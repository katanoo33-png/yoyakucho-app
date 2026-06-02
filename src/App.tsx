import { useState, useEffect, useCallback, Fragment, memo } from 'react';
import './App.css';
import * as XLSX from 'xlsx';
import type { GasPatient, VisitRecord, ModalState, SavedScheduleMeta, EmployeeList } from './types';
import {
  getGasUrl, saveGasUrl,
  fetchPatients, fetchSchedule, saveSchedule,
  fetchSaveList, loadNamedSave, saveNamedSnapshot, deleteNamedSave,
  fetchEmployees, sendExportEmail, saveToDrive,
} from './api';
import { getWeekdayDates, normalizeYoubi, newId, getKanaGroup, KANA_GROUPS } from './utils';

// ── Excel ビルド ───────────────────────────────────────────────
function buildExcelBase64(visits: VisitRecord[], year: number, month: number): string {
  const title = `${year}年${month}月 厚生局訪問スケジュール`;

  // ヘッダー行
  const headers = ['患者名', '曜日', '時間', '訪問日一覧', '担当医師', '担当衛生士', 'メモ'];

  // 患者ごとにまとめる
  const patientMap = new Map<string, VisitRecord[]>();
  for (const v of visits) {
    const arr = patientMap.get(v.patientName) ?? [];
    arr.push(v);
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

function loadDebuts(): Record<string, string> {
  try { return JSON.parse(localStorage.getItem(DEBUT_KEY) ?? '{}'); }
  catch { return {}; }
}

function loadEmployeeCache(): EmployeeList {
  try {
    const s = localStorage.getItem(EMPLOYEE_CACHE_KEY);
    return s ? (JSON.parse(s) as EmployeeList) : { doctors: [], hygienists: [] };
  } catch { return { doctors: [], hygienists: [] }; }
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
      const key = `${p.name}|${date}`;
      if (seen.has(key)) continue;
      seen.add(key);
      result.push({
        id: `auto-${p.row}-${date}`,
        date,
        patientName: p.name,
        time: p.time,
        doctor: p.doctor,
        hygienist: p.hygienist,
        isNew: debut ? date === debut : false,
        note: '',
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

const StaffBadge = memo(function StaffBadge({ name, idx, role }: { name: string; idx: number; role: 'doctor' | 'hyg' }) {
  if (!name) return <span className="staff-empty">—</span>;
  const palette = role === 'doctor' ? DOCTOR_PALETTE : HYG_PALETTE;
  const c = palette[Math.max(0, idx) % palette.length];
  return <span className="staff-badge" style={{ background: c.bg, color: c.fg }}>{name}</span>;
});

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
function PatientEditModal({ patientName, kana, year, month, currentVisits, patientInfo, employees, debutDate, onSave, onClose }: {
  patientName: string; kana: string;
  year: number; month: number;
  currentVisits: VisitRecord[];
  patientInfo?: GasPatient;
  employees: EmployeeList;
  debutDate: string | null;
  onSave: (visits: VisitRecord[], debut: string | null) => void;
  onClose: () => void;
}) {
  const first = currentVisits[0];
  const [selectedDates, setSelectedDates] = useState<string[]>(currentVisits.map(v => v.date));
  const [time, setTime] = useState(first?.time ?? patientInfo?.time ?? '');
  const [doctor, setDoctor] = useState(first?.doctor ?? patientInfo?.doctor ?? '');
  const [hygienist, setHygienist] = useState(first?.hygienist ?? patientInfo?.hygienist ?? '');
  const [note, setNote] = useState(first?.note ?? '');
  const [localDebut, setLocalDebut] = useState<string | null>(debutDate);

  function toggleDate(date: string) {
    setSelectedDates(prev =>
      prev.includes(date) ? prev.filter(d => d !== date) : [...prev, date].sort()
    );
  }

  function handleSetDebut() {
    const earliest = [...selectedDates].sort()[0];
    if (earliest) setLocalDebut(earliest);
  }

  function handleSave() {
    const newVisits: VisitRecord[] = selectedDates.map(date => ({
      id: currentVisits.find(v => v.date === date)?.id ?? `${patientName}-${date}`,
      date, patientName, time, doctor, hygienist,
      isNew: date === localDebut,
      note,
    }));
    onSave(newVisits, localDebut);
  }

  const sorted = [...selectedDates].sort();
  const hasUnknownDoc = doctor && !employees.doctors.includes(doctor);
  const hasUnknownHyg = hygienist && !employees.hygienists.includes(hygienist);

  return (
    <div className="overlay" onClick={onClose}>
      <div className="modal modal-edit" onClick={e => e.stopPropagation()}>
        <div className="edit-name-row">
          <div>
            <div className="edit-name">{patientName}</div>
            {kana && <div className="edit-kana">{kana}</div>}
          </div>
          {localDebut && (
            <div className="debut-chip">
              新規: {localDebut.slice(5).replace('-', '/')}〜
              <button className="debut-clear" onClick={() => setLocalDebut(null)}>✕</button>
            </div>
          )}
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
            {sorted.map(d => d.slice(5).replace('-', '/')).join('  ')}
          </div>
        )}

        <div className="edit-fields">
          <div className="field-row2">
            <label className="field">
              <span>時間</span>
              <input type="time" value={time} step="900" onChange={e => setTime(e.target.value)} />
            </label>
            <button
              className="btn-new-flag"
              onClick={handleSetDebut}
              disabled={!selectedDates.length || !!localDebut}
              title="最初の訪問日を「新規」として設定（以前の日はグレーアウト）"
            >
              新規
            </button>
          </div>

          <label className="field">
            <span>担当医師</span>
            <select value={doctor} onChange={e => setDoctor(e.target.value)}>
              <option value="">— 未設定 —</option>
              {employees.doctors.map(d => <option key={d} value={d}>{d}</option>)}
              {hasUnknownDoc && <option value={doctor}>{doctor}</option>}
            </select>
          </label>

          <label className="field">
            <span>担当衛生士</span>
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

        <div className="modal-row">
          <button className="btn-delete" onClick={() => onSave([], null)}>今月から削除</button>
          <div style={{ flex: 1 }} />
          <button className="btn-cancel" onClick={onClose}>キャンセル</button>
          <button className="btn-primary" onClick={handleSave}>保存</button>
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
    <div className="overlay" onClick={onClose}>
      <div className="modal" onClick={e => e.stopPropagation()}>
        <h2>新規患者を追加</h2>
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

// ── 名前付き保存モーダル ──────────────────────────────────────
function SaveAsModal({ defaultName, onSave, onClose }: {
  defaultName: string;
  onSave: (name: string) => void;
  onClose: () => void;
}) {
  const [name, setName] = useState(defaultName);
  return (
    <div className="overlay" onClick={onClose}>
      <div className="modal" onClick={e => e.stopPropagation()}>
        <h2>名前を付けて保存</h2>
        <p className="modal-note">同じ名前で保存すると上書きされます。</p>
        <label className="field">
          <span>保存名 <em>*</em></span>
          <input type="text" value={name} onChange={e => setName(e.target.value)}
            placeholder="例: 2026年5月確定版" autoFocus />
        </label>
        <div className="modal-row">
          <button className="btn-cancel" onClick={onClose}>キャンセル</button>
          <button className="btn-primary" disabled={!name.trim()}
            onClick={() => name.trim() && onSave(name.trim())}>保存</button>
        </div>
      </div>
    </div>
  );
}

// ── 保存一覧モーダル ──────────────────────────────────────────
function SaveListModal({ list, activeSave, onLoad, onDelete, onClose }: {
  list: SavedScheduleMeta[];
  activeSave: string | null;
  onLoad: (name: string) => void;
  onDelete: (name: string) => void;
  onClose: () => void;
}) {
  return (
    <div className="overlay" onClick={onClose}>
      <div className="modal modal-wide" onClick={e => e.stopPropagation()}>
        <h2>保存済みスナップショット</h2>
        {list.length === 0 ? (
          <p className="modal-note">保存済みのデータはありません。</p>
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

// ── 出力確認モーダル ──────────────────────────────────────────
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
    a.click();
    URL.revokeObjectURL(a.href);
    setDoneMsg('ダウンロードしました');
  }

  async function handleEmail() {
    if (!email.trim()) return;
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
      const url = await saveToDrive(filename, b64);
      setDoneMsg('Driveに保存しました');
      if (url) window.open(url, '_blank');
    } catch (e) {
      setDoneMsg('Drive保存エラー: ' + (e as Error).message);
    } finally {
      setSaving(false);
    }
  }

  // 患者数・訪問件数のサマリー
  const patientCount = new Set(visits.map(v => v.patientName)).size;

  return (
    <div className="overlay" onClick={onClose}>
      <div className="modal modal-export" onClick={e => e.stopPropagation()}>
        <h2>出力確認</h2>

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

  const [modal, setModal]           = useState<ModalState>({ type: 'closed' });
  const [loading, setLoading]       = useState(false);
  const [saving, setSaving]         = useState(false);
  const [toast, setToast]           = useState('');
  const [showSettings, setShowSettings] = useState(!getGasUrl());
  const [gasUrlInput, setGasUrlInput]   = useState(getGasUrl());
  const [empDoctorInput, setEmpDoctorInput] = useState('');
  const [empHygInput, setEmpHygInput]       = useState('');
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
    try {
      const [pts, saved, list, emps] = await Promise.all([
        fetchPatients(),
        fetchSchedule(y, m),
        fetchSaveList(),
        fetchEmployees(),
      ]);
      setPatients(pts);
      setSaveList(list);
      // GAS が従業員を返せた場合はキャッシュ更新、空の場合はキャッシュから復元
      const hasEmps = emps.doctors.length > 0 || emps.hygienists.length > 0;
      if (hasEmps) localStorage.setItem(EMPLOYEE_CACHE_KEY, JSON.stringify(emps));
      setEmployees(hasEmps ? emps : loadEmployeeCache());
      const debuts = loadDebuts();
      setPatientDebuts(debuts);
      const auto = autoGenerate(pts, y, m, debuts);
      setVisits(sortVisits(mergeWithSaved(auto, saved)));
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
    setModal({ type: 'closed' });
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
    setEmpDoctorInput(employees.doctors.join('、'));
    setEmpHygInput(employees.hygienists.join('、'));
    setShowSettings(true);
  }

  function handleApplyManualEmployees() {
    const parse = (s: string) =>
      s.split(/[,、，\s]+/).map(x => x.trim()).filter(Boolean);
    const manual: EmployeeList = { doctors: parse(empDoctorInput), hygienists: parse(empHygInput) };
    setEmployees(manual);
    localStorage.setItem(EMPLOYEE_CACHE_KEY, JSON.stringify(manual));
    showToast('従業員リストを手動設定しました');
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
  const patientList = (() => {
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
      if (!map.has(v.patientName)) map.set(v.patientName, { kana: v.patientName, youbis: [], info: undefined });
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
  })();

  // 50音グループ別にまとめる
  const kanaGrouped: { label: string; rows: typeof patientList }[] = KANA_GROUPS.map(g => ({
    label: g,
    rows: patientList.filter(p => p.group === g),
  })).filter(g => g.rows.length > 0);
  const otherRows = patientList.filter(p => !(KANA_GROUPS as readonly string[]).includes(p.group));
  if (otherRows.length) kanaGrouped.push({ label: '?', rows: otherRows });

  // 編集モーダル用の患者データ
  const editTarget = modal.type === 'editPatient'
    ? patientList.find(p => p.name === modal.patientName)
    : null;

  return (
    <div className="app">

      {/* ヘッダー */}
      <header className="hdr">
        <span className="hdr-title">厚生局提出スケジューラー</span>
        <div className="hdr-nav">
          <button className="btn-nav" onClick={prevMonth}>◀</button>
          <span key={`${year}-${month}`} className="hdr-month">{year}年{month}月</span>
          <button className="btn-nav" onClick={nextMonth}>▶</button>
        </div>
        {activeSave && <span className="badge-save">{activeSave}</span>}
        <div className="hdr-actions">
          <button className="btn-hdr-outline" onClick={() => setModal({ type: 'newPatient' })}>＋ 新規患者</button>
          <button className="btn-hdr-outline" onClick={() => setModal({ type: 'saveList' })}>一覧</button>
          <button className="btn-hdr-outline" onClick={() => setModal({ type: 'saveAs' })}>名前付き保存</button>
          <button className="btn-hdr-export" onClick={() => setModal({ type: 'export' })}>出力</button>
          <button className="btn-hdr-solid" onClick={handleSave} disabled={saving || loading}>
            {saving ? '保存中…' : '保存'}
          </button>
          <button className="btn-gear" onClick={openSettings}>⚙</button>
        </div>
      </header>

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
        <div className="overlay" onClick={() => setShowSettings(false)}>
          <div className="modal modal-settings" onClick={e => e.stopPropagation()}>
            <h2>設定</h2>

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
              <p className="modal-note">
                GASの「従業員リスト」シートから自動取得されます。<br />
                取得できない場合・空の場合は手動入力してください（読点「、」またはカンマ「,」区切り）。
              </p>
              <div className="emp-status">
                <span>歯科医師 <strong>{employees.doctors.length}</strong>名
                  {employees.doctors.length > 0 && `　${employees.doctors.join('、')}`}
                </span>
                <br />
                <span>歯科衛生士 <strong>{employees.hygienists.length}</strong>名
                  {employees.hygienists.length > 0 && `　${employees.hygienists.join('、')}`}
                </span>
              </div>
              <label className="field">
                <span>歯科医師（手動入力）</span>
                <input type="text" value={empDoctorInput}
                  onChange={e => setEmpDoctorInput(e.target.value)}
                  placeholder="例: 村上、伊藤" />
              </label>
              <label className="field">
                <span>歯科衛生士（手動入力）</span>
                <input type="text" value={empHygInput}
                  onChange={e => setEmpHygInput(e.target.value)}
                  placeholder="例: 藤谷、清水" />
              </label>
              <button className="btn-apply-emp" onClick={handleApplyManualEmployees}>
                手動設定を適用
              </button>
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
        ) : (
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
                        {group.label !== '?' ? `${group.label}行` : 'その他'  /* eslint-disable-line */}
                      </span>
                    </td>
                  </tr>
                  {group.rows.map(pt => (
                    <tr
                      key={pt.name}
                      className={`tr-pt${pt.isManual ? ' tr-manual' : ''}`}
                      onClick={() => setModal({ type: 'editPatient', patientName: pt.name })}
                    >
                      <td className="td-name">
                        <span className="name-avatar">{pt.name.charAt(0)}</span>
                        <span className="name-text">{pt.name}</span>
                        {pt.isManual && <span className="badge-new">新規</span>}
                        {pt.debut && <span className="badge-debut">新</span>}
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
                            {`${parseInt(v.date.slice(5, 7))}/${parseInt(v.date.slice(8))}`}
                          </span>
                        ))}
                      </td>
                      <td className="td-doctor"><StaffBadge name={pt.pvs[0]?.doctor || pt.info?.doctor || ''} role="doctor" idx={employees.doctors.indexOf(pt.pvs[0]?.doctor || pt.info?.doctor || '')} /></td>
                      <td className="td-hyg"><StaffBadge name={pt.pvs[0]?.hygienist || pt.info?.hygienist || ''} role="hyg" idx={employees.hygienists.indexOf(pt.pvs[0]?.hygienist || pt.info?.hygienist || '')} /></td>
                      <td className="td-edit">
                        <button
                          className="btn-edit"
                          onClick={e => {
                            e.stopPropagation();
                            setModal({ type: 'editPatient', patientName: pt.name });
                          }}
                        ><CalIcon />編集</button>
                      </td>
                    </tr>
                  ))}
                </Fragment>
              ))}
            </tbody>
          </table>
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
          onSave={(vs, debut) => handleSavePatient(editTarget.name, vs, debut)}
          onClose={() => setModal({ type: 'closed' })}
        />
      )}

      {/* 新規患者モーダル */}
      {modal.type === 'newPatient' && (
        <NewPatientModal
          year={year} month={month}
          employees={employees}
          onAdd={handleAddNewVisit}
          onClose={() => setModal({ type: 'closed' })}
        />
      )}

      {/* 名前付き保存 */}
      {modal.type === 'saveAs' && (
        <SaveAsModal
          defaultName={`${year}年${month}月`}
          onSave={handleSaveAs}
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

      {/* 保存一覧 */}
      {modal.type === 'saveList' && (
        <SaveListModal
          list={saveList}
          activeSave={activeSave}
          onLoad={handleLoadSave}
          onDelete={handleDeleteSave}
          onClose={() => setModal({ type: 'closed' })}
        />
      )}
    </div>
  );
}
