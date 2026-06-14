import { useState, useEffect, useRef, useCallback, useMemo, Fragment, memo } from 'react';
import './App.css';
import * as XLSX from 'xlsx';
import type { GasPatient, VisitRecord, ModalState, SavedScheduleMeta, EmployeeList, ServiceTimes } from './types';
import {
  getGasUrl, saveGasUrl,
  fetchPatients, fetchSchedule, saveSchedule,
  fetchSaveList, loadNamedSave, saveNamedSnapshot, deleteNamedSave,
  fetchEmployees, sendExportEmail, saveToDrive, updatePatientTimes,
  fetchServiceTimes, fetchRegisteredNames, registerKirokuMonth,
} from './api';
import { getWeekdayDates, normalizeYoubi, newId, getKanaGroup, KANA_GROUPS, formatDateChip, toIso, getWeekNumInMonth, generateKirokuWeeks, parseTimeToMin } from './utils';

// サービス種別定義（全て開始/終了ペア）
const SVC_GROUPS = [
  { label: '訪問時間',              short: '訪問',     startKey: 'svcVisitTime' as const,         endKey: 'svcVisitTimeEnd' as const },
  { label: '歯科医師居宅療養管理指導Ⅰ', short: 'Dr居宅Ⅰ', startKey: 'svcDrType1Start' as const,      endKey: 'svcDrType1End' as const },
  { label: '歯科医師居宅療養管理指導Ⅱ', short: 'Dr居宅Ⅱ', startKey: 'svcDrType2Start' as const,      endKey: 'svcDrType2End' as const },
  { label: '訪問衛生指導',           short: '訪問衛生',  startKey: 'svcHygieneVisitStart' as const,  endKey: 'svcHygieneVisitEnd' as const },
  { label: '歯科衛生士等居宅療養Ⅰ',  short: '衛士居宅Ⅰ', startKey: 'svcHygieneType1Start' as const, endKey: 'svcHygieneType1End' as const },
] as const;

// 開始・終了時刻のバリデーション（"HH:MM" 形式）
function timeOrderError(start: string, end: string): string {
  if (!start || !end) return '';
  return end <= start ? `終了(${end})が開始(${start})以前です` : '';
}

function defaultSvcFromPatient(p?: GasPatient): ServiceTimes {
  return {
    svcVisitTime:         p?.svcVisitTime         ?? '',
    svcVisitTimeEnd:      p?.svcVisitTimeEnd      ?? '',
    svcDrType1Start:      p?.svcDrType1Start      ?? '',
    svcDrType1End:        p?.svcDrType1End        ?? '',
    svcDrType2Start:      p?.svcDrType2Start      ?? '',
    svcDrType2End:        p?.svcDrType2End        ?? '',
    svcHygieneVisitStart: p?.svcHygieneVisitStart ?? '',
    svcHygieneVisitEnd:   p?.svcHygieneVisitEnd   ?? '',
    svcHygieneType1Start: p?.svcHygieneType1Start ?? '',
    svcHygieneType1End:   p?.svcHygieneType1End   ?? '',
  };
}

// ── Excel ビルド ───────────────────────────────────────────────
function normalizeToIso(dateStr: string): string {
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return dateStr;
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}


function buildExcelBase64(visits: VisitRecord[], year: number, month: number): string {
  const title = `${year}年${month}月 厚生局訪問スケジュール`;

  const headers = ['患者名', '曜日', '時間', '訪問日一覧', '担当医師', '担当衛生士', 'メモ'];

  const patientMap = new Map<string, VisitRecord[]>();
  for (const v of visits) {
    const arr = patientMap.get(v.patientName) ?? [];
    arr.push({ ...v, date: normalizeToIso(v.date) });
    patientMap.set(v.patientName, arr);
  }

  const rows: (string | number)[][] = [
    [title],
    [],
    headers,
  ];

  for (const [name, pvs] of patientMap) {
    const sorted = [...pvs].sort((a, b) => a.date.localeCompare(b.date));
    const dates = sorted.map(v => {
      const d = new Date(v.date + 'T00:00:00');
      return `${d.getMonth()+1}/${d.getDate()}`;
    }).join('　');
    const youbi = (() => {
      if (sorted.length === 0) return '';
      const d = new Date(sorted[0].date + 'T00:00:00');
      return ['日','月','火','水','木','金','土'][d.getDay()];
    })();
    rows.push([
      name,
      youbi,
      sorted[0]?.time ?? '',
      dates,
      sorted[0]?.doctor ?? '',
      sorted[0]?.hygienist ?? '',
      sorted[0]?.note ?? '',
    ]);
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  // タイトル行を結合
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }];
  ws['!cols'] = [
    { wch: 16 }, { wch: 6 }, { wch: 8 }, { wch: 40 },
    { wch: 12 }, { wch: 12 }, { wch: 20 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `${month}月スケジュール`);
  return XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
}

const DEBUT_KEY = 'patientDebuts_v1';
const EMPLOYEE_CACHE_KEY = 'employeeCache_v1';

// ── 編集済み患者の管理（月別） ────────────────────────────────
function editedKey(y: number, m: number) {
  return `editedPatients_v1_${y}_${String(m).padStart(2, '0')}`;
}
function loadEdited(y: number, m: number): Set<string> {
  try { return new Set(JSON.parse(localStorage.getItem(editedKey(y, m)) ?? '[]')); }
  catch { return new Set(); }
}
function saveEdited(y: number, m: number, s: Set<string>) {
  localStorage.setItem(editedKey(y, m), JSON.stringify([...s]));
}

// ── 書類作成状態 ──────────────────────────────────────────────
type DocType = 'shijisho' | 'record';
type DocStatus = { date: string; type: 'auto' | 'edited' };
type DocMonthRecord = { shijisho?: DocStatus; record?: DocStatus };
type DocMonthMap = Record<string, DocMonthRecord>; // key = patientName

function docStatusKey(year: number, month: number) {
  return `docStatus_v1_${year}_${String(month).padStart(2,'0')}`;
}
function loadDocStatus(year: number, month: number): DocMonthMap {
  try { return JSON.parse(localStorage.getItem(docStatusKey(year, month)) ?? '{}'); }
  catch { return {}; }
}
function saveDocStatus(year: number, month: number, map: DocMonthMap) {
  localStorage.setItem(docStatusKey(year, month), JSON.stringify(map));
}
function recordDocOpen(
  year: number, month: number,
  patientName: string, docType: DocType,
  map: DocMonthMap,
): DocMonthMap {
  const existing = map[patientName]?.[docType];
  const type: 'auto' | 'edited' = existing ? 'edited' : 'auto';
  const date = new Date().toISOString().slice(0, 10);
  const next: DocMonthMap = {
    ...map,
    [patientName]: { ...map[patientName], [docType]: { date, type } },
  };
  saveDocStatus(year, month, next);
  return next;
}

function loadDebuts(): Record<string, string> {
  try { return JSON.parse(localStorage.getItem(DEBUT_KEY) ?? '{}'); }
  catch { return {}; }
}


function sortVisits(vs: VisitRecord[]): VisitRecord[] {
  return [...vs].sort((a, b) =>
    a.date !== b.date ? a.date.localeCompare(b.date) : a.patientName.localeCompare(b.patientName, 'ja'),
  );
}

function autoGenerate(
  patients: GasPatient[], year: number, month: number,
  patientDebuts: Record<string, string>,
): VisitRecord[] {
  const result: VisitRecord[] = [];
  const seen = new Set<string>();
  for (const p of patients) {
    const debut = patientDebuts[p.name];
    for (const date of getWeekdayDates(year, month, p.youbi)) {
      if (debut && date < debut) continue;
      // weekFlags が設定されている場合はその週のみ生成
      if (p.weekFlags?.length === 5) {
        const wn = getWeekNumInMonth(date);
        if (wn >= 1 && wn <= 5 && !p.weekFlags[wn - 1]) continue;
      }
      const key = `${p.name}|${date}`;
      if (seen.has(key)) continue;
      seen.add(key);
      result.push({
        id: `auto-${p.row}-${date}`,
        date,
        patientName: p.name,
        time: p.svcVisitTime || p.time,
        doctor: p.doctor,
        hygienist: p.hygienist,
        isNew: debut ? date === debut : false,
        note: '',
        svcVisitTime:         p.svcVisitTime,
        svcVisitTimeEnd:      p.svcVisitTimeEnd,
        svcDrType1Start:      p.svcDrType1Start,
        svcDrType1End:        p.svcDrType1End,
        svcDrType2Start:      p.svcDrType2Start,
        svcDrType2End:        p.svcDrType2End,
        svcHygieneVisitStart: p.svcHygieneVisitStart,
        svcHygieneVisitEnd:   p.svcHygieneVisitEnd,
        svcHygieneType1Start: p.svcHygieneType1Start,
        svcHygieneType1End:   p.svcHygieneType1End,
      });
    }
  }
  return result;
}

function mergeWithSaved(auto: VisitRecord[], saved: VisitRecord[]): VisitRecord[] {
  if (!saved.length) return auto;
  const savedNames = new Set(saved.map(v => v.patientName));
  const unscheduled = auto.filter(v => !savedNames.has(v.patientName));
  return [...saved, ...unscheduled];
}

// ── 書類GAS URL ───────────────────────────────────────────────
const GAS_03D = 'https://script.google.com/macros/s/AKfycbw9pfzzA9qtZ0zlToPV8oEKciwIeFn3fgpT_PG0zME3qo_0p_Z2YS2uWeeEP7PfXKEsLQ/exec';

function buildDocUrl(form: string, patientName: string, doctor: string, hygienist: string, svc: ServiceTimes, careLevel: 'あり' | 'なし' = 'なし', yearMonth?: string): string {
  const params: Record<string, string> = {
    patient:   patientName,
    doctor,
    hygienist,
    care: careLevel,
    vt:  svc.svcVisitTime         ?? '',
    vt2: svc.svcVisitTimeEnd      ?? '',
    d1s: svc.svcDrType1Start      ?? '',
    d1e: svc.svcDrType1End        ?? '',
    d2s: svc.svcDrType2Start      ?? '',
    d2e: svc.svcDrType2End        ?? '',
    hs:  svc.svcHygieneVisitStart ?? '',
    he:  svc.svcHygieneVisitEnd   ?? '',
    ds:  svc.svcHygieneType1Start ?? '',
    de:  svc.svcHygieneType1End   ?? '',
  };
  if (yearMonth) params.yearMonth = yearMonth;
  return `${GAS_03D}?form=${form}&${new URLSearchParams(params).toString()}`;
}

// ── カレンダーアイコン ────────────────────────────────────────
function CalIcon() {
  return (
    <svg width="12" height="12" viewBox="0 0 12 12" fill="none" aria-hidden="true"
      style={{ marginRight: 4, verticalAlign: 'middle', display: 'inline-block' }}>
      <rect x="1" y="2" width="10" height="9" rx="1.5" stroke="currentColor" strokeWidth="1.3"/>
      <line x1="1" y1="5" x2="11" y2="5" stroke="currentColor" strokeWidth="1.2"/>
      <line x1="4" y1="0.5" x2="4" y2="3.5" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round"/>
      <line x1="8" y1="0.5" x2="8" y2="3.5" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round"/>
    </svg>
  );
}

// ── スタッフバッジ（担当者ごとに固有色） ─────────────────────
const DOCTOR_PALETTE = [
  { bg: '#ddeeff', fg: '#1e5a9a' }, // ソフトブルー
  { bg: '#ddf2e5', fg: '#1a6e46' }, // セージグリーン
  { bg: '#fff3cc', fg: '#8a5e10' }, // ウォームハニー
  { bg: '#fde0d4', fg: '#8a3020' }, // テラコッタライト
  { bg: '#e0eeff', fg: '#3a4ea8' }, // ラベンダーブルー
  { bg: '#e4f6ec', fg: '#2a6040' }, // ミントグリーン
];
const HYG_PALETTE = [
  { bg: '#eedeff', fg: '#7038aa' }, // ソフトラベンダー
  { bg: '#ffdcee', fg: '#aa2a6a' }, // ソフトピンク
  { bg: '#ddf4f8', fg: '#1e6888' }, // ダスティティール
  { bg: '#fff0de', fg: '#aa5a18' }, // ウォームピーチ
  { bg: '#f8e0ff', fg: '#8a30a0' }, // モーブ
  { bg: '#ffe8e0', fg: '#aa4030' }, // コーラル
];

function nameHash(name: string): number {
  let h = 0;
  for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) >>> 0;
  return h;
}

function familyName(name: string): string {
  const space = name.indexOf(' ');
  return space > 0 ? name.slice(0, space) : name;
}

const StaffBadge = memo(function StaffBadge({ name, role }: { name: string; role: 'doctor' | 'hyg' }) {
  if (!name) return <span className="staff-empty">—</span>;
  const palette = role === 'doctor' ? DOCTOR_PALETTE : HYG_PALETTE;
  const c = palette[nameHash(name) % palette.length];
  return <span className="staff-badge" style={{ background: c.bg, color: c.fg }}>{familyName(name)}</span>;
});

// ── 書類ステータスドット（ボタン内） ─────────────────────────
function DocStatusDot({ s }: { s: DocStatus }) {
  const color = s.type === 'auto' ? '#5bc8a8' : '#f4a440';
  return (
    <span style={{ display: 'inline-block', width: 7, height: 7, borderRadius: '50%',
      background: color, marginLeft: 5, verticalAlign: 'middle', flexShrink: 0 }} />
  );
}

// ── 書類ステータスバッジ（テーブル行用） ─────────────────────
const DOC_LABELS: Record<DocType, string> = { shijisho: '指', record: '録' };
function DocBadge({ docType, s, onClick }: { docType: DocType; s?: DocStatus; onClick?: () => void }) {
  if (!s) {
    return (
      <span className="doc-badge" style={{ background: '#f0f2f5', color: '#aab0bc', border: '1px solid #dde1e8' }}
        title="未作成">
        {DOC_LABELS[docType]}
      </span>
    );
  }
  const isAuto = s.type === 'auto';
  const bg     = isAuto ? '#e6f9f3' : '#fff5e6';
  const color  = isAuto ? '#1a8a6a' : '#b86000';
  const border = isAuto ? '#9adfc8' : '#f4c47a';
  const mm = s.date.slice(5, 7).replace(/^0/, '');
  const dd = s.date.slice(8, 10).replace(/^0/, '');
  return (
    <span
      className={`doc-badge${onClick ? ' doc-badge-clickable' : ''}`}
      style={{ background: bg, color, border: `1px solid ${border}` }}
      title={onClick
        ? `${isAuto ? '自動作成' : '編集済み'} ${s.date}　▶ クリックで開く`
        : `${isAuto ? '自動作成' : '編集済み'} ${s.date}`}
      onClick={onClick ? e => { e.stopPropagation(); onClick(); } : undefined}
    >
      {DOC_LABELS[docType]}
      <span className="doc-badge-date">{mm}/{dd}</span>
    </span>
  );
}

// ── 曜日バッジ ────────────────────────────────────────────────
const YOUBI_COLOR: Record<string, { bg: string; fg: string }> = {
  '月': { bg: '#deeeff', fg: '#3a6fb5' },
  '火': { bg: '#ffeae8', fg: '#c04444' },
  '水': { bg: '#defff0', fg: '#2e9966' },
  '木': { bg: '#fff8de', fg: '#9a7010' },
  '金': { bg: '#f5eeff', fg: '#7744cc' },
  '土': { bg: '#eef0ff', fg: '#4455cc' },
  '日': { bg: '#ffeef0', fg: '#cc3344' },
};

const YoubiBadge = memo(function YoubiBadge({ youbi }: { youbi: string }) {
  const c = YOUBI_COLOR[youbi];
  if (!youbi) return <span className="youbi-empty">—</span>;
  return (
    <span className="youbi-badge" style={c ? { background: c.bg, color: c.fg } : {}}>
      {youbi}
    </span>
  );
});

// ── カレンダー ────────────────────────────────────────────────
function MonthCalendar({ year, month, selected, debutDate, onToggle }: {
  year: number; month: number;
  selected: string[];
  debutDate: string | null;
  onToggle: (date: string) => void;
}) {
  const pad = (n: number) => String(n).padStart(2, '0');
  const firstDow = new Date(year, month - 1, 1).getDay();
  const daysInMonth = new Date(year, month, 0).getDate();

  const cells: (string | null)[] = Array(firstDow).fill(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(`${year}-${pad(month)}-${pad(d)}`);

  return (
    <div className="cal">
      <div className="cal-head">
        {['日', '月', '火', '水', '木', '金', '土'].map((d, i) => (
          <div key={d} className={`cal-dow${i === 0 ? ' sun' : i === 6 ? ' sat' : ''}`}>{d}</div>
        ))}
      </div>
      <div className="cal-body">
        {cells.map((date, i) => {
          if (!date) return <div key={`e${i}`} className="cal-cell" />;
          const dow = new Date(date).getDay();
          const isSel = selected.includes(date);
          const isDisabled = !!debutDate && date < debutDate;
          const isDebut = date === debutDate;
          return (
            <div
              key={date}
              className={[
                'cal-cell cal-day',
                isSel ? 'sel' : '',
                isDisabled ? 'dis' : '',
                isDebut ? 'debut' : '',
                dow === 0 ? 'sun' : dow === 6 ? 'sat' : '',
              ].filter(Boolean).join(' ')}
              onClick={() => !isDisabled && onToggle(date)}
            >
              {new Date(date + 'T00:00:00').getDate()}
              {isDebut && <span className="debut-dot" />}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ── 患者編集モーダル ──────────────────────────────────────────
function PatientEditModal({ patientName, kana, year, month, currentVisits, patientInfo, employees, debutDate, docStatus, onSave, onDocOpen, onClose }: {
  patientName: string; kana: string;
  year: number; month: number;
  currentVisits: VisitRecord[];
  patientInfo?: GasPatient;
  employees: EmployeeList;
  debutDate: string | null;
  docStatus: DocMonthRecord;
  onSave: (visits: VisitRecord[], debut: string | null) => void;
  onDocOpen: (docType: DocType) => void;
  onClose: () => void;
}) {
  const first = currentVisits[0];
  const defSvc = defaultSvcFromPatient(patientInfo);

  const [selectedDates, setSelectedDates] = useState<string[]>(
    currentVisits.map(v => {
      if (/^\d{4}-\d{2}-\d{2}$/.test(v.date)) return v.date;
      const d = new Date(v.date);
      return isNaN(d.getTime()) ? v.date : toIso(d);
    })
  );

  // 日付→サービス時間のマップ
  const [svcTimes, setSvcTimes] = useState<Record<string, ServiceTimes>>(() => {
    const init: Record<string, ServiceTimes> = {};
    currentVisits.forEach(v => {
      init[v.date] = {
        svcVisitTime:         v.svcVisitTime         ?? defSvc.svcVisitTime,
        svcVisitTimeEnd:      v.svcVisitTimeEnd      ?? defSvc.svcVisitTimeEnd,
        svcDrType1Start:      v.svcDrType1Start      ?? defSvc.svcDrType1Start,
        svcDrType1End:        v.svcDrType1End        ?? defSvc.svcDrType1End,
        svcDrType2Start:      v.svcDrType2Start      ?? defSvc.svcDrType2Start,
        svcDrType2End:        v.svcDrType2End        ?? defSvc.svcDrType2End,
        svcHygieneVisitStart: v.svcHygieneVisitStart ?? defSvc.svcHygieneVisitStart,
        svcHygieneVisitEnd:   v.svcHygieneVisitEnd   ?? defSvc.svcHygieneVisitEnd,
        svcHygieneType1Start: v.svcHygieneType1Start ?? defSvc.svcHygieneType1Start,
        svcHygieneType1End:   v.svcHygieneType1End   ?? defSvc.svcHygieneType1End,
      };
    });
    return init;
  });
  const [kirokuLoaded, setKirokuLoaded] = useState(false);
  // 登録データ取得中は操作をブロックし、DB時間の一瞬の表示（チラつき）を防ぐ
  const [kirokuLoading, setKirokuLoading] = useState(true);
  const kirokuTimesRef = useRef<ServiceTimes[]>([]);

  // kiroku（厚生局_訪問時間記録）から週別時間を取得して適用
  useEffect(() => {
    setKirokuLoading(true);
    fetchServiceTimes(patientName, year, month).then(weekTimes => {
      kirokuTimesRef.current = weekTimes;
      if (weekTimes.length > 0) {
        setSvcTimes(prev => {
          const sorted = Object.keys(prev).sort();
          if (sorted.length === 0) return prev;
          const next = { ...prev };
          sorted.forEach((date, i) => {
            if (weekTimes[i]) next[date] = weekTimes[i];
          });
          return next;
        });
        setKirokuLoaded(true);
      }
    }).finally(() => {
      setKirokuLoading(false);
    });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [patientName, year, month]);

  const [doctor, setDoctor] = useState(first?.doctor ?? patientInfo?.doctor ?? '');
  const [hygienist, setHygienist] = useState(first?.hygienist ?? patientInfo?.hygienist ?? '');
  const [note, setNote] = useState(first?.note ?? '');
  const [localDebut, setLocalDebut] = useState<string | null>(debutDate);
  // 介護区分（DBの値をデフォルト、変更可能）
  const dbCareLevel = patientInfo?.careLevel ?? 'なし';
  const [careLevel, setCareLevel] = useState<'あり' | 'なし'>(dbCareLevel);
  const [showCareLevelEdit, setShowCareLevelEdit] = useState(false);
  const careLevelChanged = careLevel !== dbCareLevel;

  function toggleDate(date: string) {
    setSelectedDates(prev => {
      if (prev.includes(date)) {
        return prev.filter(d => d !== date);
      } else {
        const newDates = [...prev, date].sort();
        const weekIdx = newDates.indexOf(date);
        const kirokuSvc = kirokuTimesRef.current[weekIdx];
        setSvcTimes(t => ({ ...t, [date]: kirokuSvc ?? defSvc }));
        return newDates;
      }
    });
  }

  function updateSvc(date: string, field: keyof ServiceTimes, val: string) {
    setSvcTimes(prev => ({ ...prev, [date]: { ...(prev[date] ?? defSvc), [field]: val } }));
  }

  async function handleSave() {
    const newVisits: VisitRecord[] = selectedDates.map(date => {
      const svc = svcTimes[date] ?? defSvc;
      return {
        id: currentVisits.find(v => v.date === date)?.id ?? `${patientName}-${date}`,
        date, patientName,
        time: svc.svcVisitTime || '',
        doctor, hygienist,
        isNew: date === localDebut,
        note,
        ...svc,
      };
    });
    // 時間または介護区分が変更されていればマスターに自動書き戻し
    if (patientInfo?.row) {
      const firstDate = [...selectedDates].sort()[0];
      const repSvc = firstDate ? (svcTimes[firstDate] ?? defSvc) : defSvc;
      const timesChanged = Object.entries(repSvc).some(
        ([k, v]) => v !== (defSvc[k as keyof ServiceTimes] ?? '')
      );
      if (timesChanged || careLevelChanged) {
        updatePatientTimes(patientInfo.row, { ...repSvc, careLevel }).catch(() => {/* silent */});
      }
    }
    onSave(newVisits, localDebut);
  }

  const sorted = [...selectedDates].sort();
  const hasUnknownDoc = doctor && !employees.doctors.includes(doctor);
  const hasUnknownHyg = hygienist && !employees.hygienists.includes(hygienist);

  // いずれかの日付・種別で開始>終了エラーがあれば保存不可
  const hasSvcError = sorted.some(date => {
    const svc = svcTimes[date] ?? defSvc;
    return SVC_GROUPS.some(g => !!timeOrderError(svc[g.startKey], svc[g.endKey]));
  });

  return (
    <div className="overlay" onClick={onClose} role="presentation">
      <div className="modal modal-edit" onClick={e => e.stopPropagation()} style={{ position: 'relative' }}
        role="dialog" aria-modal="true" aria-label={`${patientName} の編集`}>
        {kirokuLoading && (
          <div className="kiroku-loading-overlay">
            <div className="kiroku-loading-box">
              <div className="kiroku-spinner" />
              <div>登録データを反映中…</div>
            </div>
          </div>
        )}
        <div className="edit-name-row">
          <div>
            <div className="edit-name">{patientName}</div>
            {kana && <div className="edit-kana">{kana}</div>}
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 6 }}>
            {/* 介護区分 */}
            <div className="care-level-row">
              <span className="care-level-label">介護区分</span>
              {!showCareLevelEdit ? (
                <>
                  <span className={`care-level-badge${careLevel === 'あり' ? ' care-ari' : ' care-nashi'}`}>
                    {careLevel}
                  </span>
                  <button className="care-level-change" onClick={() => setShowCareLevelEdit(true)}>
                    変更
                  </button>
                </>
              ) : (
                <>
                  <button
                    className={`care-level-btn${careLevel === 'なし' ? ' selected' : ''}`}
                    onClick={() => setCareLevel('なし')}
                  >なし</button>
                  <button
                    className={`care-level-btn${careLevel === 'あり' ? ' selected' : ''}`}
                    onClick={() => setCareLevel('あり')}
                  >あり</button>
                  <button className="care-level-change" onClick={() => setShowCareLevelEdit(false)}>完了</button>
                </>
              )}
              {careLevelChanged && (
                <span className="care-level-changed">※変更あり（保存時にDB反映）</span>
              )}
            </div>
            {localDebut && (
              <div className="debut-chip">
                新規: {localDebut.slice(5).replace('-', '/')}〜
                <button className="debut-clear" onClick={() => setLocalDebut(null)}>✕</button>
              </div>
            )}
          </div>
        </div>

        <div className="cal-label">{year}年{month}月 — クリックで訪問日を切り替え</div>
        <MonthCalendar
          year={year} month={month}
          selected={selectedDates}
          debutDate={localDebut}
          onToggle={toggleDate}
        />

        {sorted.length > 0 && (
          <div className="selected-dates">
            {sorted.map(d => formatDateChip(d)).join('　')}
          </div>
        )}

        {/* kirokuデータ読み込みバナー */}
        {kirokuLoaded && !kirokuLoading && (
          <div className="kiroku-reflected-banner">
            ✅ 先ほど登録した時間（{year}年{month}月）を反映しています
          </div>
        )}

        {/* 週別サービス時間カード */}
        {sorted.length > 0 && !kirokuLoading && (
          <div className="svc-section">
            <div className="svc-section-header">
              <div className="svc-section-label">厚生局サービス時間（週別）</div>
            </div>
            <div className="svc-cards">
              {sorted.map(date => {
                const wn = getWeekNumInMonth(date);
                const svc = svcTimes[date] ?? defSvc;
                return (
                  <div key={date} className="svc-card">
                    <div className="svc-card-header">
                      <span className="svc-card-date">{formatDateChip(date)}</span>
                      <span className="svc-card-week">第{wn}週</span>
                    </div>
                    {SVC_GROUPS.map(g => {
                      const err = timeOrderError(svc[g.startKey], svc[g.endKey]);
                      return (
                        <div key={g.short} className="svc-row">
                          <span className="svc-row-label" title={g.label}>{g.short}</span>
                          <span className="svc-pair">
                            <input
                              type="time" className={`svc-time-input${err ? ' svc-input-err' : ''}`} step="900"
                              value={svc[g.startKey]}
                              onChange={e => updateSvc(date, g.startKey, e.target.value)}
                            />
                            <span className="svc-sep">〜</span>
                            <input
                              type="time" className={`svc-time-input${err ? ' svc-input-err' : ''}`} step="900"
                              value={svc[g.endKey]}
                              onChange={e => updateSvc(date, g.endKey, e.target.value)}
                            />
                          </span>
                          {err && <span className="svc-time-err">{err}</span>}
                        </div>
                      );
                    })}
                  </div>
                );
              })}
            </div>
          </div>
        )}

        <div className="edit-fields">
          <label className="field">
            <span>診察医師</span>
            <select value={doctor} onChange={e => setDoctor(e.target.value)}>
              <option value="">— 未設定 —</option>
              {employees.doctors.map(d => <option key={d} value={d}>{d}</option>)}
              {hasUnknownDoc && <option value={doctor}>{doctor}</option>}
            </select>
          </label>

          <label className="field">
            <span>衛生士</span>
            <select value={hygienist} onChange={e => setHygienist(e.target.value)}>
              <option value="">— 未設定 —</option>
              {employees.hygienists.map(h => <option key={h} value={h}>{h}</option>)}
              {hasUnknownHyg && <option value={hygienist}>{hygienist}</option>}
            </select>
          </label>

          <label className="field">
            <span>メモ</span>
            <input type="text" value={note} onChange={e => setNote(e.target.value)} placeholder="備考など" />
          </label>
        </div>

        {/* 書類作成リンク */}
        {(() => {
          const firstDate = [...selectedDates].sort()[0];
          const repSvc = firstDate ? (svcTimes[firstDate] ?? defSvc) : defSvc;
          function openDoc(form: string, docType: DocType) {
            const url = buildDocUrl(form, patientName, doctor, hygienist, repSvc, careLevel);
            window.open(url, `doc_${patientName}_${form}`, 'noopener,noreferrer');
            onDocOpen(docType);
          }
          const isManualPat = !patientInfo;
          return (
            <div className="doc-link-row">
              <span className="doc-link-label">書類作成（訪問歯科衛生管理）</span>
              {isManualPat && (
                <span className="doc-link-note">※手動追加患者はDBに未登録のため患者名が自動選択されません</span>
              )}
              <div className="doc-link-btns">
                <button className="btn-doc-link" onClick={() => openDoc('shijisho', 'shijisho')}>
                  指示書{docStatus.shijisho && <DocStatusDot s={docStatus.shijisho} />}
                </button>
                <button className="btn-doc-link" onClick={() => openDoc('record', 'record')}>
                  指導記録簿{docStatus.record && <DocStatusDot s={docStatus.record} />}
                </button>
              </div>
            </div>
          );
        })()}

        <div className="modal-row">
          <button className="btn-delete" onClick={() => onSave([], null)}>今月から削除</button>
          <div style={{ flex: 1 }} />
          <button className="btn-cancel" onClick={onClose}>キャンセル</button>
          <button className="btn-primary" onClick={handleSave} disabled={hasSvcError}
            title={hasSvcError ? '時刻エラーを修正してから保存してください' : ''}>保存</button>
        </div>
      </div>
    </div>
  );
}

// ── 他の月に展開モーダル ─────────────────────────────────────
function ExpandMonthModal({ patientName, kana, patientInfo, year, month, onClose }: {
  patientName: string;
  kana: string;
  patientInfo?: GasPatient;
  year: number;
  month: number;
  onClose: () => void;
}) {
  type MonthEntry = { y: number; m: number; key: string; label: string };

  const pastMonths: MonthEntry[] = (() => {
    const arr: MonthEntry[] = [];
    let y = year, m = month;
    for (let i = 0; i < 6; i++) {
      m--; if (m < 1) { m = 12; y--; }
      arr.push({ y, m, key: `${y}-${String(m).padStart(2,'0')}`, label: `${y}年${m}月` });
    }
    return arr;
  })();

  const futureMonths: MonthEntry[] = (() => {
    const arr: MonthEntry[] = [];
    let y = year, m = month;
    for (let i = 0; i < 6; i++) {
      m++; if (m > 12) { m = 1; y++; }
      arr.push({ y, m, key: `${y}-${String(m).padStart(2,'0')}`, label: `${y}年${m}月` });
    }
    return arr;
  })();

  const allMonths = [...[...pastMonths].reverse(), ...futureMonths];

  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [expanding, setExpanding] = useState(false);
  const [doneMsg, setDoneMsg] = useState('');

  function toggle(key: string) {
    setSelected(prev => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key); else next.add(key);
      return next;
    });
  }

  function selectAll() { setSelected(new Set(allMonths.map(c => c.key))); }
  function clearAll()  { setSelected(new Set()); }

  async function handleExpand() {
    if (!patientInfo || selected.size === 0) return;
    setExpanding(true);
    setDoneMsg('');
    const kanaVal = patientInfo.firstName || kana || '';
    const hasKaigo = patientInfo.careLevel === 'あり';
    const baseTime = parseTimeToMin(patientInfo.time || '13:00');
    try {
      for (const c of allMonths) {
        if (!selected.has(c.key)) continue;
        const data = generateKirokuWeeks(hasKaigo, baseTime);
        await registerKirokuMonth({
          name: patientName, kana: kanaVal, kaigo: hasKaigo,
          year: c.y, month: c.m, data,
        });
      }
      setDoneMsg(`${selected.size}ヶ月分を展開しました`);
      setSelected(new Set());
    } finally {
      setExpanding(false);
    }
  }

  return (
    <div className="overlay" onClick={onClose} role="presentation">
      <div className="modal modal-expand" onClick={e => e.stopPropagation()}
        role="dialog" aria-modal="true" aria-label={`${patientName} — 他の月への展開`}>
        {/* ヘッダー */}
        <div className="expand-modal-header">
          <span className="name-avatar expand-avatar">{patientName.charAt(0)}</span>
          <div>
            <div className="expand-modal-name">{patientName}</div>
            <div className="expand-modal-sub">📅 他の月への展開</div>
          </div>
        </div>

        <p className="expand-modal-desc">
          展開したい月にチェックを入れて「展開する」を押してください。<br />
          各月ごとに<strong>新しいランダム時間</strong>を生成します（コピーではありません）。
        </p>

        <div className="expand-ctrl-row">
          <button className="btn-sel-all" onClick={selectAll}>全選択</button>
          <button className="btn-sel-all" onClick={clearAll}>全解除</button>
          {selected.size > 0 && (
            <span className="expand-count-badge">{selected.size}ヶ月選択中</span>
          )}
        </div>

        <div className="expand-two-col">
          {/* 過去 */}
          <div className="expand-col">
            <div className="expand-col-label past">◀ 過去6ヶ月</div>
            {pastMonths.map(c => (
              <label key={c.key} className={`expand-check-row${selected.has(c.key) ? ' checked' : ''}`}>
                <input type="checkbox" checked={selected.has(c.key)} onChange={() => toggle(c.key)} />
                <span className="expand-check-label">{c.label}</span>
              </label>
            ))}
          </div>

          {/* 未来 */}
          <div className="expand-col">
            <div className="expand-col-label future">▶ 未来6ヶ月</div>
            {futureMonths.map(c => (
              <label key={c.key} className={`expand-check-row${selected.has(c.key) ? ' checked' : ''}`}>
                <input type="checkbox" checked={selected.has(c.key)} onChange={() => toggle(c.key)} />
                <span className="expand-check-label">{c.label}</span>
              </label>
            ))}
          </div>
        </div>

        {doneMsg && <div className="expand-done">✅ {doneMsg}</div>}

        <div className="modal-row" style={{ marginTop: 20 }}>
          <div style={{ flex: 1 }} />
          <button className="btn-cancel" onClick={onClose}>閉じる</button>
          <button
            className="btn-primary"
            onClick={handleExpand}
            disabled={expanding || selected.size === 0 || !patientInfo}
            title={!patientInfo ? 'DBに登録されていない患者は展開できません' : ''}
          >
            {expanding ? '展開中…' : selected.size > 0 ? `${selected.size}ヶ月に展開する` : '月を選んでください'}
          </button>
        </div>
      </div>
    </div>
  );
}

// ── 印刷バッチモーダル ────────────────────────────────────────
function PrintBatchModal({ patientName, patientInfo, year, month, onClose, onDocOpen }: {
  patientName: string;
  patientInfo?: GasPatient;
  year: number;
  month: number;
  onClose: () => void;
  onDocOpen: (docType: DocType, y: number, m: number) => void;
}) {
  type MonthEntry = { y: number; m: number; key: string; label: string; isCurrent: boolean; ds: DocMonthRecord };

  const months: MonthEntry[] = (() => {
    const arr: MonthEntry[] = [];
    // 過去6ヶ月（古い順）
    const past: { y: number; m: number }[] = [];
    let y = year, m = month;
    for (let i = 0; i < 6; i++) {
      m--; if (m < 1) { m = 12; y--; }
      past.unshift({ y, m });
    }
    past.forEach(({ y, m }) => {
      arr.push({ y, m, key: `${y}-${String(m).padStart(2,'0')}`, label: `${y}年${m}月`, isCurrent: false, ds: loadDocStatus(y, m)[patientName] ?? {} });
    });
    // 今月
    arr.push({ y: year, m: month, key: `${year}-${String(month).padStart(2,'0')}`, label: `${year}年${month}月`, isCurrent: true, ds: loadDocStatus(year, month)[patientName] ?? {} });
    // 未来6ヶ月
    y = year; m = month;
    for (let i = 0; i < 6; i++) {
      m++; if (m > 12) { m = 1; y++; }
      arr.push({ y, m, key: `${y}-${String(m).padStart(2,'0')}`, label: `${y}年${m}月`, isCurrent: false, ds: loadDocStatus(y, m)[patientName] ?? {} });
    }
    return arr;
  })();

  const [selectedKeys, setSelectedKeys] = useState<Set<string>>(new Set());
  const [docType, setDocType] = useState<DocType>('shijisho');
  const [opening, setOpening] = useState(false);
  const [doneMsg, setDoneMsg] = useState('');

  function toggle(key: string) {
    setSelected(prev => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key); else next.add(key);
      return next;
    });
  }
  function setSelected(fn: (prev: Set<string>) => Set<string>) { setSelectedKeys(fn); }

  function selectHasDocs() {
    const docKey = docType === 'shijisho' ? 'shijisho' : 'record';
    setSelectedKeys(new Set(months.filter(c => c.ds[docKey]).map(c => c.key)));
  }
  function selectAll()  { setSelectedKeys(new Set(months.map(c => c.key))); }
  function clearAll()   { setSelectedKeys(new Set()); }

  function getSvc(): ServiceTimes {
    return {
      svcVisitTime:         patientInfo?.svcVisitTime         ?? '',
      svcVisitTimeEnd:      patientInfo?.svcVisitTimeEnd      ?? '',
      svcDrType1Start:      patientInfo?.svcDrType1Start      ?? '',
      svcDrType1End:        patientInfo?.svcDrType1End        ?? '',
      svcDrType2Start:      patientInfo?.svcDrType2Start      ?? '',
      svcDrType2End:        patientInfo?.svcDrType2End        ?? '',
      svcHygieneVisitStart: patientInfo?.svcHygieneVisitStart ?? '',
      svcHygieneVisitEnd:   patientInfo?.svcHygieneVisitEnd   ?? '',
      svcHygieneType1Start: patientInfo?.svcHygieneType1Start ?? '',
      svcHygieneType1End:   patientInfo?.svcHygieneType1End   ?? '',
    };
  }

  async function handlePrint() {
    if (selectedKeys.size === 0) return;
    setOpening(true);
    setDoneMsg('');
    const form   = docType === 'shijisho' ? 'shijisho' : 'record';
    const doctor    = patientInfo?.doctor    ?? '';
    const hygienist = patientInfo?.hygienist ?? '';
    const careLevel = patientInfo?.careLevel ?? 'なし';
    const svc = getSvc();
    const targets = months.filter(c => selectedKeys.has(c.key));
    try {
      for (const c of targets) {
        const url = buildDocUrl(form, patientName, doctor, hygienist, svc, careLevel, c.key);
        window.open(url, `doc_${patientName}_${form}_${c.key}`, 'noopener,noreferrer');
        onDocOpen(docType, c.y, c.m);
        // タブが連続で開くのを少し間隔を空ける（ブラウザのポップアップ対策）
        await new Promise(r => setTimeout(r, 120));
      }
      setDoneMsg(`${targets.length}ヶ月分を開きました。各タブから印刷してください。`);
    } finally {
      setOpening(false);
    }
  }

  const form_label = docType === 'shijisho' ? '居宅療養管理指導指示書' : '訪問歯科衛生指導記録簿';

  return (
    <div className="overlay" onClick={onClose} role="presentation">
      <div className="modal modal-print" onClick={e => e.stopPropagation()}
        role="dialog" aria-modal="true" aria-label={`${patientName} — 書類の一括出力`}>
        {/* ヘッダー */}
        <div className="print-modal-header">
          <span className="name-avatar print-avatar">{patientName.charAt(0)}</span>
          <div>
            <div className="print-modal-name">{patientName}</div>
            <div className="print-modal-sub">🖨 書類の一括出力</div>
          </div>
        </div>

        {/* 書類種別タブ */}
        <div className="print-doc-tabs">
          <button
            className={`print-doc-tab${docType === 'shijisho' ? ' active' : ''}`}
            onClick={() => { setDocType('shijisho'); clearAll(); }}
          >指示書</button>
          <button
            className={`print-doc-tab${docType === 'record' ? ' active' : ''}`}
            onClick={() => { setDocType('record'); clearAll(); }}
          >指導記録簿</button>
          <span className="print-doc-name">{form_label}</span>
        </div>

        {/* 選択コントロール */}
        <div className="print-ctrl-row">
          <button className="btn-sel-all" onClick={selectAll}>全選択</button>
          <button className="btn-sel-all" onClick={clearAll}>全解除</button>
          <button className="btn-sel-all btn-sel-has" onClick={selectHasDocs} title="作成済みの月だけ選択">
            作成済みのみ
          </button>
          {selectedKeys.size > 0 && (
            <span className="print-count-badge">{selectedKeys.size}ヶ月選択中</span>
          )}
        </div>

        {/* 月一覧 */}
        <div className="print-month-list">
          {months.map(c => {
            const docKey = docType === 'shijisho' ? 'shijisho' : 'record';
            const docS = c.ds[docKey];
            const isChecked = selectedKeys.has(c.key);
            return (
              <label
                key={c.key}
                className={[
                  'print-month-row',
                  isChecked ? 'checked' : '',
                  c.isCurrent ? 'current' : '',
                  docS ? 'has-doc' : '',
                ].filter(Boolean).join(' ')}
              >
                <input type="checkbox" checked={isChecked} onChange={() => toggle(c.key)} />
                <span className="print-month-label">
                  {c.label}
                  {c.isCurrent && <span className="print-current-tag">今月</span>}
                </span>
                <span className="print-month-status">
                  {docS ? (
                    <span className="print-doc-badge created">
                      {docS.type === 'edited' ? '編集済み' : '作成済み'}
                      <span className="print-doc-date">{docS.date.slice(5).replace('-','/')}</span>
                    </span>
                  ) : (
                    <span className="print-doc-badge none">未作成</span>
                  )}
                </span>
              </label>
            );
          })}
        </div>

        {doneMsg && (
          <div className="print-done">
            ✅ {doneMsg}
            <div className="print-done-note">ブラウザのポップアップがブロックされた場合は許可してください。</div>
          </div>
        )}

        <div className="modal-row" style={{ marginTop: 16 }}>
          <div style={{ flex: 1 }} />
          <button className="btn-cancel" onClick={onClose}>閉じる</button>
          <button
            className="btn-primary"
            onClick={handlePrint}
            disabled={opening || selectedKeys.size === 0}
          >
            {opening ? '開いています…' : selectedKeys.size > 0 ? `${selectedKeys.size}ヶ月分を開く` : '月を選んでください'}
          </button>
        </div>
      </div>
    </div>
  );
}

// ── 新規患者モーダル ──────────────────────────────────────────
function NewPatientModal({ year, month, employees, onAdd, onClose }: {
  year: number; month: number;
  employees: EmployeeList;
  onAdd: (v: VisitRecord) => void;
  onClose: () => void;
}) {
  const defaultDate = `${year}-${String(month).padStart(2, '0')}-01`;
  const [name, setName] = useState('');
  const [date, setDate] = useState(defaultDate);
  const [time, setTime] = useState('14:00');
  const [doctor, setDoctor] = useState('');
  const [hygienist, setHygienist] = useState('');
  const [note, setNote] = useState('');

  return (
    <div className="overlay" onClick={onClose} role="presentation">
      <div className="modal" onClick={e => e.stopPropagation()}
        role="dialog" aria-modal="true" aria-labelledby="new-patient-title">
        <h2 id="new-patient-title">新規患者を追加</h2>
        <p className="modal-note">マスター外の患者を1件追加します。</p>
        <label className="field">
          <span>患者名 <em>*</em></span>
          <input type="text" value={name} onChange={e => setName(e.target.value)}
            placeholder="例: 田中 太郎" autoFocus />
        </label>
        <label className="field">
          <span>訪問日 <em>*</em></span>
          <input type="date" value={date} onChange={e => setDate(e.target.value)} />
        </label>
        <label className="field">
          <span>時間</span>
          <input type="time" value={time} step="900" onChange={e => setTime(e.target.value)} />
        </label>
        <label className="field">
          <span>担当医師</span>
          <select value={doctor} onChange={e => setDoctor(e.target.value)}>
            <option value="">— 未設定 —</option>
            {employees.doctors.map(d => <option key={d} value={d}>{d}</option>)}
          </select>
        </label>
        <label className="field">
          <span>担当衛生士</span>
          <select value={hygienist} onChange={e => setHygienist(e.target.value)}>
            <option value="">— 未設定 —</option>
            {employees.hygienists.map(h => <option key={h} value={h}>{h}</option>)}
          </select>
        </label>
        <label className="field">
          <span>メモ</span>
          <input type="text" value={note} onChange={e => setNote(e.target.value)} />
        </label>
        <div className="modal-row">
          <button className="btn-cancel" onClick={onClose}>キャンセル</button>
          <button className="btn-primary"
            disabled={!name.trim()}
            onClick={() => name.trim() && onAdd({ id: newId(), date, patientName: name.trim(), time, doctor, hygienist, isNew: true, note })}>
            追加
          </button>
        </div>
      </div>
    </div>
  );
}

// ── スナップショットモーダル（一覧＋新規保存） ───────────────
function SaveListModal({ list, activeSave, defaultSaveName, onLoad, onDelete, onSaveNew, onClose }: {
  list: SavedScheduleMeta[];
  activeSave: string | null;
  defaultSaveName: string;
  onLoad: (name: string) => void;
  onDelete: (name: string) => void;
  onSaveNew: (name: string) => void;
  onClose: () => void;
}) {
  const [newName, setNewName] = useState(defaultSaveName);
  return (
    <div className="overlay" onClick={onClose} role="presentation">
      <div className="modal modal-wide" onClick={e => e.stopPropagation()}
        role="dialog" aria-modal="true" aria-labelledby="savelist-title">
        <h2 id="savelist-title">スナップショット管理</h2>
        <p className="modal-note">
          現在の編集内容を名前を付けて保存できます。過去に保存したバージョンをいつでも読み込めます。
        </p>

        {/* 新規保存エリア */}
        <div className="snapshot-save-row">
          <input
            className="snapshot-name-input"
            type="text"
            value={newName}
            onChange={e => setNewName(e.target.value)}
            placeholder="例: 2026年6月確定版"
          />
          <button
            className="btn-primary"
            disabled={!newName.trim()}
            onClick={() => newName.trim() && onSaveNew(newName.trim())}
          >
            この名前で保存
          </button>
        </div>

        {/* 保存済み一覧 */}
        {list.length === 0 ? (
          <p className="modal-note" style={{ marginTop: 12 }}>まだ保存済みのスナップショットはありません。</p>
        ) : (
          <ul className="save-list">
            {list.map(s => (
              <li key={s.name} className={`save-item${s.name === activeSave ? ' active' : ''}`}>
                <div className="save-info">
                  <span className="save-name">{s.name}</span>
                  <span className="save-meta">{s.count}件 · {s.createdAt}</span>
                </div>
                <div className="save-actions">
                  <button className="btn-load" onClick={() => onLoad(s.name)}>読み込み</button>
                  <button className="btn-del-sm" onClick={() => onDelete(s.name)}>削除</button>
                </div>
              </li>
            ))}
          </ul>
        )}
        <div className="modal-row" style={{ marginTop: 16 }}>
          <div style={{ flex: 1 }} />
          <button className="btn-cancel" onClick={onClose}>閉じる</button>
        </div>
      </div>
    </div>
  );
}

function ExportModal({ year, month, visits, onClose }: {
  year: number; month: number;
  visits: VisitRecord[];
  onClose: () => void;
}) {
  const [email, setEmail]       = useState('');
  const [sending, setSending]   = useState(false);
  const [saving, setSaving]     = useState(false);
  const [doneMsg, setDoneMsg]   = useState('');

  const filename = `${year}年${month}月_厚生局訪問スケジュール.xlsx`;

  function handleDownload() {
    const b64 = buildExcelBase64(visits, year, month);
    const bin = atob(b64);
    const buf = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
    setDoneMsg('ダウンロードしました');
  }

  const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  async function handleEmail() {
    if (!email.trim() || !EMAIL_RE.test(email.trim())) {
      setDoneMsg('メールアドレスの形式が正しくありません');
      return;
    }
    setSending(true);
    setDoneMsg('');
    try {
      const b64 = buildExcelBase64(visits, year, month);
      await sendExportEmail(email.trim(), filename, b64);
      setDoneMsg(`${email} に送信しました`);
    } catch (e) {
      setDoneMsg('送信エラー: ' + (e as Error).message);
    } finally {
      setSending(false);
    }
  }

  async function handleDrive() {
    setSaving(true);
    setDoneMsg('');
    try {
      const b64 = buildExcelBase64(visits, year, month);
      const { url, folderName } = await saveToDrive(filename, b64);
      setDoneMsg(`「${folderName}」に保存しました`);
      if (url) window.open(url, '_blank', 'noopener,noreferrer');
    } catch (e) {
      setDoneMsg('Drive保存エラー: ' + (e as Error).message);
    } finally {
      setSaving(false);
    }
  }

  // 患者数・訪問件数のサマリー
  const patientCount = new Set(visits.map(v => v.patientName)).size;

  return (
    <div className="overlay" onClick={onClose} role="presentation">
      <div className="modal modal-export" onClick={e => e.stopPropagation()}
        role="dialog" aria-modal="true" aria-labelledby="export-title">
        <h2 id="export-title">出力確認</h2>

        {/* サマリー */}
        <div className="export-summary">
          <div className="export-summary-title">{year}年{month}月 スケジュール</div>
          <div className="export-summary-stats">
            <span className="export-stat"><strong>{patientCount}</strong> 名</span>
            <span className="export-stat-sep">/</span>
            <span className="export-stat"><strong>{visits.length}</strong> 件の訪問</span>
          </div>
          <div className="export-filename">📄 {filename}</div>
        </div>

        {/* 1. ローカル保存 */}
        <div className="export-section">
          <div className="export-section-label">① ローカルに保存</div>
          <button className="btn-export-action" onClick={handleDownload}>
            Excel ダウンロード
          </button>
        </div>

        {/* 2. メール送信 */}
        <div className="export-section">
          <div className="export-section-label">② メールで送付</div>
          <div className="export-email-row">
            <input
              type="email"
              className="export-email-input"
              value={email}
              onChange={e => setEmail(e.target.value)}
              placeholder="送付先メールアドレス"
            />
            <button
              className="btn-export-action"
              onClick={handleEmail}
              disabled={sending || !email.trim()}
            >
              {sending ? '送信中…' : '送信'}
            </button>
          </div>
        </div>

        {/* 3. Drive保存 */}
        <div className="export-section">
          <div className="export-section-label">③ Google Drive に保存</div>
          <div className="export-drive-path">保存先: 指定フォルダ</div>
          <button
            className="btn-export-action"
            onClick={handleDrive}
            disabled={saving}
          >
            {saving ? '保存中…' : 'Drive に保存'}
          </button>
        </div>

        {/* 完了メッセージ */}
        {doneMsg && <div className="export-done">{doneMsg}</div>}

        <div className="modal-row" style={{ marginTop: 20 }}>
          <div style={{ flex: 1 }} />
          <button className="btn-cancel" onClick={onClose}>閉じる</button>
        </div>
      </div>
    </div>
  );
}

// ── メインApp ─────────────────────────────────────────────────
export default function App() {
  const now = new Date();
  const [year, setYear]   = useState(now.getFullYear());
  const [month, setMonth] = useState(now.getMonth() + 1);

  const [patients, setPatients]     = useState<GasPatient[]>([]);
  const [visits, setVisits]         = useState<VisitRecord[]>([]);
  const [saveList, setSaveList]     = useState<SavedScheduleMeta[]>([]);
  const [activeSave, setActiveSave] = useState<string | null>(null);
  const [employees, setEmployees]   = useState<EmployeeList>({ doctors: [], hygienists: [] });
  const [patientDebuts, setPatientDebuts] = useState<Record<string, string>>(loadDebuts);
  const [docStatusMap, setDocStatusMap]   = useState<DocMonthMap>(() => loadDocStatus(now.getFullYear(), now.getMonth() + 1));
  const [editedPatients, setEditedPatients] = useState<Set<string>>(() => loadEdited(now.getFullYear(), now.getMonth() + 1));
  // kirokuで登録済みの患者名（その月）。編集可否の判定に使う
  const [kirokuRegistered, setKirokuRegistered] = useState<Set<string>>(new Set());
  const [appMode] = useState<'input' | 'list'>('list');
  const [kirokuTimePatchDone, setKirokuTimePatchDone] = useState(true); // kiroku時間バックグラウンド取得完了フラグ

  const [modal, setModal]           = useState<ModalState>({ type: 'closed' });
  const [loading, setLoading]       = useState(false);
  const [saving, setSaving]         = useState(false);
  const [toast, setToast]           = useState('');
  const [showSettings, setShowSettings] = useState(!getGasUrl());
  const [gasUrlInput, setGasUrlInput]   = useState(getGasUrl());
  const [activeKana, setActiveKana] = useState<string>('全');

  function showToast(msg: string) {
    setToast(msg);
    setTimeout(() => setToast(''), 3200);
  }

  function saveDebuts(next: Record<string, string>) {
    setPatientDebuts(next);
    localStorage.setItem(DEBUT_KEY, JSON.stringify(next));
  }

  const loadData = useCallback(async (y: number, m: number) => {
    if (!getGasUrl()) return;
    setLoading(true);
    setActiveSave(null);
    setDocStatusMap(loadDocStatus(y, m));
    setEditedPatients(loadEdited(y, m));
    try {
      const [pts, saved, list, emps, registered] = await Promise.all([
        fetchPatients(),
        fetchSchedule(y, m),
        fetchSaveList(),
        fetchEmployees(),
        fetchRegisteredNames(y, m),
      ]);
      setPatients(pts);
      setSaveList(list);
      setKirokuRegistered(new Set(registered));
      localStorage.removeItem(EMPLOYEE_CACHE_KEY); // 古いキャッシュを強制クリア
      // GASがDate objectを文字列化して返す場合を除去（スプレッドシートのセル書式問題）
      const isDateStr = (s: string) => /^(Mon|Tue|Wed|Thu|Fri|Sat|Sun) /.test(s);
      const cleanList = (arr: string[]) => arr.filter(s => s && !isDateStr(s));
      const cleanedEmps: EmployeeList = {
        doctors:    cleanList(emps.doctors),
        hygienists: cleanList(emps.hygienists),
      };
      const hasEmps = cleanedEmps.doctors.length > 0 || cleanedEmps.hygienists.length > 0;
      const finalEmps: EmployeeList = hasEmps ? cleanedEmps : (() => {
        // GASリスト取得不可の場合のみ患者マスターから補完（Date文字列も除去）
        const docsFromPts = cleanList([...new Set(pts.map(p => p.doctor).filter(Boolean))]).sort((a,b) => a.localeCompare(b,'ja'));
        const hygsFromPts = cleanList([...new Set(pts.map(p => p.hygienist).filter(Boolean))]).sort((a,b) => a.localeCompare(b,'ja'));
        return { doctors: docsFromPts, hygienists: hygsFromPts };
      })();
      setEmployees(finalEmps);
      const debuts = loadDebuts();
      setPatientDebuts(debuts);
      const auto = autoGenerate(pts, y, m, debuts);
      const merged = sortVisits(mergeWithSaved(auto, saved));
      setVisits(merged);

      // kiroku登録済みだがまだ保存していない患者の時間列を非同期でパッチ
      const savedNames = new Set(saved.map(v => v.patientName));
      const needKiroku = registered.filter(n => !savedNames.has(n));
      if (needKiroku.length > 0) {
        setKirokuTimePatchDone(false);
        Promise.all(
          needKiroku.map(name =>
            fetchServiceTimes(name, y, m)
              .then(times => ({ name, visitTime: times[0]?.svcVisitTime ?? '' }))
              .catch(() => ({ name, visitTime: '' }))
          )
        ).then(results => {
          const timeMap = new Map(
            results.filter(r => r.visitTime).map(r => [r.name, r.visitTime])
          );
          if (timeMap.size > 0) {
            setVisits(prev => prev.map(v => {
              const t = timeMap.get(v.patientName);
              return t ? { ...v, time: t, svcVisitTime: t } : v;
            }));
          }
        }).catch(() => {}).finally(() => setKirokuTimePatchDone(true));
      }
    } catch (e) {
      showToast('読込エラー: ' + (e as Error).message);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { loadData(year, month); }, [year, month, loadData]);

  function prevMonth() {
    setModal({ type: 'closed' });
    if (month === 1) { setYear(y => y - 1); setMonth(12); }
    else setMonth(m => m - 1);
  }
  function nextMonth() {
    setModal({ type: 'closed' });
    if (month === 12) { setYear(y => y + 1); setMonth(1); }
    else setMonth(m => m + 1);
  }

  async function handleSave() {
    setSaving(true);
    try {
      await saveSchedule(visits);
      showToast('保存しました');
    } catch (e) {
      showToast('保存エラー: ' + (e as Error).message);
    } finally {
      setSaving(false);
    }
  }

  async function handleSaveAs(name: string) {
    setSaving(true);
    try {
      await saveNamedSnapshot(name, visits);
      const list = await fetchSaveList();
      setSaveList(list);
      setActiveSave(name);
      showToast(`「${name}」として保存しました`);
    } catch (e) {
      showToast('保存エラー: ' + (e as Error).message);
    } finally {
      setSaving(false);
    }
    setModal({ type: 'closed' });
  }

  async function handleLoadSave(name: string) {
    setLoading(true);
    try {
      const data = await loadNamedSave(name);
      setVisits(sortVisits(data));
      setActiveSave(name);
      showToast(`「${name}」を読み込みました`);
    } catch (e) {
      showToast('読み込みエラー: ' + (e as Error).message);
    } finally {
      setLoading(false);
    }
    setModal({ type: 'closed' });
  }

  async function handleDeleteSave(name: string) {
    if (!window.confirm(`「${name}」を削除しますか？`)) return;
    try {
      await deleteNamedSave(name);
      const list = await fetchSaveList();
      setSaveList(list);
      if (activeSave === name) setActiveSave(null);
      showToast(`「${name}」を削除しました`);
    } catch (e) {
      showToast('削除エラー: ' + (e as Error).message);
    }
  }

  function handleSavePatient(patientName: string, newVisits: VisitRecord[], debut: string | null) {
    setVisits(prev => sortVisits([
      ...prev.filter(v => v.patientName !== patientName),
      ...newVisits,
    ]));
    const next = { ...patientDebuts };
    if (debut) next[patientName] = debut;
    else delete next[patientName];
    saveDebuts(next);
    // 編集済みとしてマーク
    setEditedPatients(prev => {
      const s = new Set(prev);
      s.add(patientName);
      saveEdited(year, month, s);
      return s;
    });
    setModal({ type: 'closed' });
  }

  function handleResetPatient(patientName: string) {
    // 自動生成に戻す
    const pts = patients.filter(p => p.name === patientName);
    const autoVisits = autoGenerate(pts, year, month, patientDebuts);
    setVisits(prev => sortVisits([
      ...prev.filter(v => v.patientName !== patientName),
      ...autoVisits,
    ]));
    // 編集済みフラグを解除
    setEditedPatients(prev => {
      const s = new Set(prev);
      s.delete(patientName);
      saveEdited(year, month, s);
      return s;
    });
    showToast(`${patientName} をリセットしました`);
  }

  function handleDocOpen(patientName: string, docType: DocType) {
    setDocStatusMap(prev => {
      const next = recordDocOpen(year, month, patientName, docType, prev);
      return next;
    });
  }

  // テーブル行から直接書類を開く（編集モーダルを経由しない）
  function openDocFromList(pt: typeof patientList[number], docType: DocType) {
    const doctor    = pt.pvs[0]?.doctor    || pt.info?.doctor    || '';
    const hygienist = pt.pvs[0]?.hygienist || pt.info?.hygienist || '';
    const svc: ServiceTimes = {
      svcVisitTime:         pt.pvs[0]?.svcVisitTime         ?? pt.info?.svcVisitTime         ?? '',
      svcVisitTimeEnd:      pt.pvs[0]?.svcVisitTimeEnd      ?? pt.info?.svcVisitTimeEnd      ?? '',
      svcDrType1Start:      pt.pvs[0]?.svcDrType1Start      ?? pt.info?.svcDrType1Start      ?? '',
      svcDrType1End:        pt.pvs[0]?.svcDrType1End        ?? pt.info?.svcDrType1End        ?? '',
      svcDrType2Start:      pt.pvs[0]?.svcDrType2Start      ?? pt.info?.svcDrType2Start      ?? '',
      svcDrType2End:        pt.pvs[0]?.svcDrType2End        ?? pt.info?.svcDrType2End        ?? '',
      svcHygieneVisitStart: pt.pvs[0]?.svcHygieneVisitStart ?? pt.info?.svcHygieneVisitStart ?? '',
      svcHygieneVisitEnd:   pt.pvs[0]?.svcHygieneVisitEnd   ?? pt.info?.svcHygieneVisitEnd   ?? '',
      svcHygieneType1Start: pt.pvs[0]?.svcHygieneType1Start ?? pt.info?.svcHygieneType1Start ?? '',
      svcHygieneType1End:   pt.pvs[0]?.svcHygieneType1End   ?? pt.info?.svcHygieneType1End   ?? '',
    };
    const form = docType === 'shijisho' ? 'shijisho' : 'record';
    const cl = pt.info?.careLevel ?? 'なし';
    const url = buildDocUrl(form, pt.name, doctor, hygienist, svc, cl);
    window.open(url, `doc_${pt.name}_${form}`, 'noopener,noreferrer');
    handleDocOpen(pt.name, docType);
  }



  function handleAddNewVisit(v: VisitRecord) {
    setVisits(prev => sortVisits([...prev, v]));
    setModal({ type: 'closed' });
    showToast('新規患者を追加しました');
  }

  function handleSaveSettings() {
    saveGasUrl(gasUrlInput);
    setShowSettings(false);
    loadData(year, month);
  }

  function openSettings() {
    setGasUrlInput(getGasUrl());
    setShowSettings(true);
  }


  function scrollToKana(label: string) {
    setActiveKana(label);
    if (label === '全') {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    } else {
      document.getElementById(`kg-${label}`)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }

  // 患者一覧（50音ソート）- 同名患者が複数行ある場合も youbis に全曜日を収集
  const patientList = useMemo(() => {
    const map = new Map<string, { kana: string; youbis: string[]; info?: GasPatient }>();
    for (const p of patients) {
      const ex = map.get(p.name);
      if (ex) {
        const ny = normalizeYoubi(p.youbi);
        if (ny && !ex.youbis.includes(ny)) ex.youbis.push(ny);
      } else {
        const ny = normalizeYoubi(p.youbi);
        map.set(p.name, { kana: p.firstName, youbis: ny ? [ny] : [], info: p });
      }
    }
    for (const v of visits) {
      if (!map.has(v.patientName)) {
        // 手動追加患者: 訪問日から曜日を導出
        const d = new Date(v.date + 'T00:00:00');
        const youbiFromDate = isNaN(d.getTime()) ? '' : ['日','月','火','水','木','金','土'][d.getDay()];
        map.set(v.patientName, { kana: v.patientName, youbis: youbiFromDate ? [youbiFromDate] : [], info: undefined });
      }
    }
    return [...map.entries()]
      .map(([name, { kana, youbis, info }]) => ({
        name, kana, youbis, info,
        group: getKanaGroup(kana),
        pvs: sortVisits(visits.filter(v => v.patientName === name)),
        isManual: !info,
        debut: patientDebuts[name] ?? null,
      }))
      .sort((a, b) => a.kana.localeCompare(b.kana, 'ja'));
  }, [patients, visits, patientDebuts]);

  // 50音グループ別にまとめる
  const kanaGrouped = useMemo(() => {
    const groups: { label: string; rows: (typeof patientList)[number][] }[] = KANA_GROUPS.map(g => ({
      label: g,
      rows: patientList.filter(p => p.group === g),
    })).filter(g => g.rows.length > 0);
    const otherRows = patientList.filter(p => !(KANA_GROUPS as readonly string[]).includes(p.group));
    if (otherRows.length) groups.push({ label: '?', rows: otherRows });
    return groups;
  }, [patientList]);

  // 編集モーダル用の患者データ
  const editTarget = modal.type === 'editPatient'
    ? patientList.find(p => p.name === modal.patientName)
    : null;

  return (
    <div className="app">

      {/* ヘッダー */}
      <header className="hdr">
        <h1 className="hdr-title">厚生局提出スケジューラー</h1>
        <div className="hdr-nav">
          <button className="btn-nav" onClick={prevMonth} aria-label="前の月">◀</button>
          <span key={`${year}-${month}`} className="hdr-month" aria-live="polite">{year}年{month}月</span>
          <button className="btn-nav" onClick={nextMonth} aria-label="次の月">▶</button>
        </div>
        <div className="hdr-actions">
          <button className="btn-hdr-outline" onClick={() => setModal({ type: 'newPatient' })} title="マスターシートにない患者を今月だけ手動で追加します">＋ 新規患者</button>
          <button className="btn-hdr-export" onClick={() => setModal({ type: 'export' })}
            disabled={!kirokuTimePatchDone}
            title={kirokuTimePatchDone
              ? `当月（${year}年${month}月）のスケジュールをExcel出力・メール送信・Google Driveに保存します`
              : '時間情報を読み込み中…しばらくお待ちください'}>
            {kirokuTimePatchDone ? '出力' : '読込中…'}
          </button>
          <button className="btn-hdr-solid" onClick={handleSave} disabled={saving || loading}
            title={`当月（${year}年${month}月）の訪問スケジュールをクラウドに保存します。次回アクセス時にも反映されます。`}>
            {saving ? '保存中…' : '保存'}
          </button>
          <button className="btn-gear" onClick={openSettings}>⚙</button>
        </div>
      </header>

      {/* 登録バー */}
      {(() => {
        const total = patientList.length;
        const done  = patientList.filter(p => editedPatients.has(p.name)).length;
        const allDone = total > 0 && done === total;
        return (
          <div className="reg-bar">
            {total > 0 && (
              <span className={`reg-progress${allDone ? ' done' : ''}`}>
                {allDone ? '✅ 全員登録済み' : `${done} / ${total}名 登録済み`}
              </span>
            )}
            <a
              className="btn-reg-cta"
              href="https://kiroku.katanolab.dev"
              target="_blank"
              rel="noopener noreferrer"
            >
              時間を登録する →
            </a>
          </div>
        );
      })()}

      {/* 50音タブ */}
      <div className="kana-bar">
        {(['全', ...KANA_GROUPS] as string[]).map(k => (
          <button
            key={k}
            className={`kana-tab${activeKana === k ? ' active' : ''}`}
            onClick={() => scrollToKana(k)}
          >{k}</button>
        ))}
      </div>

      {/* トースト */}
      {toast && <div className="toast">{toast}</div>}

      {/* 設定モーダル */}
      {showSettings && (
        <div className="overlay" onClick={() => setShowSettings(false)} role="presentation">
          <div className="modal modal-settings" onClick={e => e.stopPropagation()}
            role="dialog" aria-modal="true" aria-labelledby="settings-title">
            <h2 id="settings-title">設定</h2>

            {/* GAS URL */}
            <div className="settings-section">
              <div className="settings-section-title">GAS URL</div>
              <p className="modal-note">デプロイしたウェブアプリのURLを入力してください。</p>
              <input className="modal-input full" type="url" value={gasUrlInput}
                onChange={e => setGasUrlInput(e.target.value)}
                placeholder="https://script.google.com/macros/s/…/exec" />
            </div>

            {/* 従業員リスト */}
            <div className="settings-section">
              <div className="settings-section-title">従業員リスト</div>
              <p className="modal-note">GASの「従業員リスト」シートから自動取得されます。</p>
              <div className="emp-status">
                <span>歯科医師 <strong>{employees.doctors.length}</strong>名
                  {employees.doctors.length > 0 && `　${employees.doctors.join('、')}`}
                </span>
                <br />
                <span>歯科衛生士 <strong>{employees.hygienists.length}</strong>名
                  {employees.hygienists.length > 0 && `　${employees.hygienists.join('、')}`}
                </span>
              </div>
            </div>

            <div className="modal-row">
              <button className="btn-cancel" onClick={() => setShowSettings(false)}>キャンセル</button>
              <button className="btn-primary" onClick={handleSaveSettings}>GAS再接続・再読込</button>
            </div>
          </div>
        </div>
      )}

      {/* テーブル */}
      <div className="table-wrap">
        {loading ? (
          <div className="loading">読み込み中…</div>
        ) : patientList.length === 0 ? (
          <div className="empty">
            {getGasUrl()
              ? 'データがありません（マスターシートのO列にTRUEの患者がいません）'
              : 'GAS URLを設定してください（右上 ⚙ ボタン）'}
          </div>
        ) : appMode === 'input' ? (
          /* ── タブ①：時間入力モード ──────────────────────────── */
          <table className="tbl">
            <thead>
              <tr>
                <th className="th-name">患者名</th>
                <th className="th-youbi">曜日</th>
                <th className="th-dates">訪問日</th>
                <th className="th-doctor">担当医師</th>
                <th className="th-hyg">担当衛生士</th>
                <th className="th-edit"></th>
              </tr>
            </thead>
            <tbody>
              {kanaGrouped.map(group => (
                <Fragment key={group.label}>
                  <tr className="group-row">
                    <td colSpan={6}>
                      <span id={`kg-${group.label}`} className="group-label">
                        {group.label !== '?' ? `${group.label}行` : 'その他'}
                      </span>
                    </td>
                  </tr>
                  {group.rows.map(pt => {
                    const isEdited = editedPatients.has(pt.name);
                    // kiroku登録済み or 既に編集済みなら編集可能（白）。未登録はグレー・閲覧のみ
                    const isRegistered = kirokuRegistered.has(pt.name) || isEdited;
                    return (
                      <tr
                        key={pt.name}
                        className={`tr-pt${pt.isManual ? ' tr-manual' : ''}${isEdited ? ' tr-step1-done' : ''}${isRegistered ? '' : ' tr-unregistered'}`}
                      >
                        <td className="td-name">
                          <span className="name-avatar">{pt.name.charAt(0)}</span>
                          <span className="name-text">{pt.name}</span>
                          {pt.isManual && <span className="badge-new">新規</span>}
                          {pt.debut && <span className="badge-debut">新</span>}
                          {isEdited && <span className="badge-step1-done">✅ 確定</span>}
                          {!isRegistered && <span className="badge-unregistered">未登録</span>}
                          {isRegistered && !pt.isManual && (
                            <span className="action-ovals">
                              <button
                                type="button"
                                className="action-oval oval-expand"
                                onClick={e => { e.stopPropagation(); setModal({ type: 'expandPatient', patientName: pt.name }); }}
                              >他の月に展開</button>
                              <button
                                type="button"
                                className="action-oval oval-print"
                                onClick={e => { e.stopPropagation(); setModal({ type: 'printBatch', patientName: pt.name }); }}
                              >他の月も一括印刷</button>
                            </span>
                          )}
                        </td>
                        <td className="td-youbi">
                          {pt.youbis.length > 0
                            ? pt.youbis.map(y => <YoubiBadge key={y} youbi={y} />)
                            : <YoubiBadge youbi="" />}
                        </td>
                        <td className="td-dates">
                          {pt.pvs.length === 0 ? (
                            <span className="no-visit">訪問なし</span>
                          ) : pt.pvs.map((v, i) => (
                            <span
                              key={v.id}
                              className={`chip${v.isNew ? ' chip-debut' : (i % 2 === 0 ? ' chip-alt' : ' chip-active')}`}
                            >
                              {formatDateChip(v.date)}
                            </span>
                          ))}
                        </td>
                        <td className="td-doctor"><StaffBadge name={pt.pvs[0]?.doctor || pt.info?.doctor || ''} role="doctor" /></td>
                        <td className="td-hyg"><StaffBadge name={pt.pvs[0]?.hygienist || pt.info?.hygienist || ''} role="hyg" /></td>
                        <td className="td-edit">
                          {isRegistered ? (
                            <button
                              className="btn-edit btn-time-input"
                              onClick={e => {
                                e.stopPropagation();
                                setModal({ type: 'editPatient', patientName: pt.name });
                              }}
                            ><CalIcon />時間入力</button>
                          ) : (
                            <button
                              className="btn-edit btn-time-input"
                              disabled
                              title="先に登録ページ（kiroku）で時間を登録してください"
                            ><CalIcon />時間入力</button>
                          )}
                          {isEdited && (
                            <button
                              className="btn-reset"
                              onClick={e => {
                                e.stopPropagation();
                                if (window.confirm(`${pt.name} をDB元データにリセットしますか？`)) {
                                  handleResetPatient(pt.name);
                                }
                              }}
                              title="自動生成（DB元データ）に戻す"
                            >↺</button>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </Fragment>
              ))}
            </tbody>
          </table>
        ) : (
          /* ── タブ②：一覧確認モード ──────────────────────────── */
          <>
            {(() => {
              const total = patientList.length;
              const done  = patientList.filter(p => editedPatients.has(p.name)).length;
              const remaining = total - done;
              return remaining > 0 ? (
                <div className="step2-warn">
                  ⚠ <strong>{remaining}名</strong>が未登録です。
                  <a className="step2-warn-link" href="https://kiroku.katanolab.dev" target="_blank" rel="noopener noreferrer">② 登録ページ</a>
                  で時間を登録してください。
                </div>
              ) : null;
            })()}
            <table className="tbl">
              <thead>
                <tr>
                  <th className="th-name">患者名</th>
                  <th className="th-youbi">曜日</th>
                  <th className="th-time">時間</th>
                  <th className="th-dates">訪問日</th>
                  <th className="th-doctor">担当医師</th>
                  <th className="th-hyg">担当衛生士</th>
                  <th className="th-edit"></th>
                </tr>
              </thead>
              <tbody>
                {kanaGrouped.map(group => (
                  <Fragment key={group.label}>
                    <tr className="group-row">
                      <td colSpan={7}>
                        <span id={`kg-${group.label}`} className="group-label">
                          {group.label !== '?' ? `${group.label}行` : 'その他'}
                        </span>
                      </td>
                    </tr>
                    {group.rows.map(pt => {
                      const isEdited = editedPatients.has(pt.name);
                      // kiroku登録済み or 編集済みのみ編集可能（白）。未登録はグレー・閲覧のみ
                      const isRegistered = kirokuRegistered.has(pt.name) || isEdited;
                      return (
                        <tr
                          key={pt.name}
                          className={`tr-pt${pt.isManual ? ' tr-manual' : ''}${isRegistered ? ' tr-edited' : ' tr-locked tr-unregistered'}`}
                          onClick={isRegistered ? () => setModal({ type: 'editPatient', patientName: pt.name }) : undefined}
                        >
                          <td className="td-name">
                            <span className="name-avatar">{pt.name.charAt(0)}</span>
                            <span className="name-text">{pt.name}</span>
                            {pt.isManual && <span className="badge-new">新規</span>}
                            {pt.debut && <span className="badge-debut">新</span>}
                            {!isRegistered && <span className="badge-unregistered">未登録</span>}
                            {(() => {
                              const ds = docStatusMap[pt.name] ?? {};
                              return (
                                <>
                                  <span className="doc-badges">
                                    <DocBadge
                                      docType="shijisho"
                                      s={ds.shijisho}
                                      onClick={ds.shijisho ? () => openDocFromList(pt, 'shijisho') : undefined}
                                    />
                                    <DocBadge
                                      docType="record"
                                      s={ds.record}
                                      onClick={ds.record ? () => openDocFromList(pt, 'record') : undefined}
                                    />
                                  </span>
                                  {isRegistered && !pt.isManual && (
                                    <span className="action-ovals">
                                      <button
                                        type="button"
                                        className="action-oval oval-expand"
                                        onClick={e => { e.stopPropagation(); setModal({ type: 'expandPatient', patientName: pt.name }); }}
                                      >他の月に展開</button>
                                      <button
                                        type="button"
                                        className="action-oval oval-print"
                                        onClick={e => { e.stopPropagation(); setModal({ type: 'printBatch', patientName: pt.name }); }}
                                      >他の月も一括印刷</button>
                                    </span>
                                  )}
                                </>
                              );
                            })()}
                          </td>
                          <td className="td-youbi">
                            {pt.youbis.length > 0
                              ? pt.youbis.map(y => <YoubiBadge key={y} youbi={y} />)
                              : <YoubiBadge youbi="" />}
                          </td>
                          <td className="td-time">{pt.pvs[0]?.time || pt.info?.time || '—'}</td>
                          <td className="td-dates">
                            {pt.pvs.length === 0 ? (
                              <span className="no-visit">訪問なし</span>
                            ) : pt.pvs.map((v, i) => (
                              <span
                                key={v.id}
                                className={`chip${v.isNew ? ' chip-debut' : (i % 2 === 0 ? ' chip-alt' : ' chip-active')}`}
                              >
                                {formatDateChip(v.date)}
                              </span>
                            ))}
                          </td>
                          <td className="td-doctor"><StaffBadge name={pt.pvs[0]?.doctor || pt.info?.doctor || ''} role="doctor" /></td>
                          <td className="td-hyg"><StaffBadge name={pt.pvs[0]?.hygienist || pt.info?.hygienist || ''} role="hyg" /></td>
                          <td className="td-edit">
                            {isRegistered ? (
                              <button
                                className="btn-edit"
                                onClick={e => {
                                  e.stopPropagation();
                                  setModal({ type: 'editPatient', patientName: pt.name });
                                }}
                              ><CalIcon />編集</button>
                            ) : (
                              <button
                                className="btn-edit btn-time-input"
                                disabled
                                title="先に登録ページ（kiroku）で時間を登録してください"
                              ><CalIcon />編集</button>
                            )}
                            {isEdited && (
                              <button
                                className="btn-reset"
                                onClick={e => {
                                  e.stopPropagation();
                                  if (window.confirm(`${pt.name} をDB元データにリセットしますか？`)) {
                                    handleResetPatient(pt.name);
                                  }
                                }}
                                title="自動生成（DB元データ）に戻す"
                              >↺</button>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </Fragment>
                ))}
              </tbody>
            </table>
          </>
        )}
      </div>

      {/* 患者編集モーダル */}
      {modal.type === 'editPatient' && editTarget && (
        <PatientEditModal
          patientName={editTarget.name}
          kana={editTarget.kana}
          year={year} month={month}
          currentVisits={editTarget.pvs}
          patientInfo={editTarget.info}
          employees={employees}
          debutDate={editTarget.debut}
          docStatus={docStatusMap[editTarget.name] ?? {}}
          onSave={(vs, debut) => handleSavePatient(editTarget.name, vs, debut)}
          onDocOpen={(docType) => handleDocOpen(editTarget.name, docType)}
          onClose={() => setModal({ type: 'closed' })}
        />
      )}

      {/* 印刷バッチモーダル */}
      {modal.type === 'printBatch' && (() => {
        const pt = patientList.find(p => p.name === modal.patientName);
        if (!pt) return null;
        return (
          <PrintBatchModal
            patientName={pt.name}
            patientInfo={pt.info}
            year={year}
            month={month}
            onClose={() => setModal({ type: 'closed' })}
            onDocOpen={(docType, y, m) => {
              // 他月のdocStatusを更新（今月のstateだけ即時反映、他月はlocalStorageに書き込み済み）
              if (y === year && m === month) {
                handleDocOpen(pt.name, docType);
              } else {
                const otherMap = loadDocStatus(y, m);
                const next = recordDocOpen(y, m, pt.name, docType, otherMap);
                saveDocStatus(y, m, next);
              }
            }}
          />
        );
      })()}

      {/* 他の月に展開モーダル */}
      {modal.type === 'expandPatient' && (() => {
        const pt = patientList.find(p => p.name === modal.patientName);
        if (!pt) return null;
        return (
          <ExpandMonthModal
            patientName={pt.name}
            kana={pt.kana}
            patientInfo={pt.info}
            year={year}
            month={month}
            onClose={() => setModal({ type: 'closed' })}
          />
        );
      })()}

      {/* 新規患者モーダル */}
      {modal.type === 'newPatient' && (
        <NewPatientModal
          year={year} month={month}
          employees={employees}
          onAdd={handleAddNewVisit}
          onClose={() => setModal({ type: 'closed' })}
        />
      )}

      {/* 出力確認 */}
      {modal.type === 'export' && (
        <ExportModal
          year={year} month={month}
          visits={visits}
          onClose={() => setModal({ type: 'closed' })}
        />
      )}

      {/* スナップショット */}
      {modal.type === 'saveList' && (
        <SaveListModal
          list={saveList}
          activeSave={activeSave}
          defaultSaveName={`${year}年${month}月`}
          onLoad={handleLoadSave}
          onDelete={handleDeleteSave}
          onSaveNew={handleSaveAs}
          onClose={() => setModal({ type: 'closed' })}
        />
      )}
    </div>
  );
}
