import type { GasPatient, VisitRecord, SavedScheduleMeta, EmployeeList } from './types';

const GAS_URL_KEY = 'gasUrl_hokkyoku_v1';

export function getGasUrl(): string {
  // 環境変数(ビルド時) → localStorage の順で解決
  return import.meta.env.VITE_GAS_URL || localStorage.getItem(GAS_URL_KEY) || '';
}

export function saveGasUrl(url: string): void {
  localStorage.setItem(GAS_URL_KEY, url.trim());
}

export async function fetchPatients(): Promise<GasPatient[]> {
  const url = getGasUrl();
  if (!url) throw new Error('GAS URLが未設定です');
  const res = await fetch(`${url}?action=getPatients`, { redirect: 'follow' });
  const json = await res.json() as { ok: boolean; data?: GasPatient[]; error?: string };
  if (!json.ok) throw new Error(json.error ?? 'GAS取得エラー');
  return json.data ?? [];
}

export async function fetchSchedule(year: number, month: number): Promise<VisitRecord[]> {
  const url = getGasUrl();
  if (!url) return [];
  try {
    const res = await fetch(`${url}?action=getSchedule&year=${year}&month=${month}`, { redirect: 'follow' });
    const json = await res.json() as { ok: boolean; data?: VisitRecord[]; error?: string };
    return json.ok ? (json.data ?? []) : [];
  } catch {
    return [];
  }
}

export async function saveSchedule(visits: VisitRecord[]): Promise<void> {
  const url = getGasUrl();
  if (!url) throw new Error('GAS URLが未設定です');
  const body = new URLSearchParams();
  body.append('action', 'saveSchedule');
  body.append('data', JSON.stringify(visits));
  const res = await fetch(url, { method: 'POST', body, redirect: 'follow' });
  const json = await res.json() as { ok: boolean; error?: string };
  if (!json.ok) throw new Error(json.error ?? '保存エラー');
}

export async function fetchSaveList(): Promise<SavedScheduleMeta[]> {
  const url = getGasUrl();
  if (!url) return [];
  try {
    const res = await fetch(`${url}?action=listSaves`, { redirect: 'follow' });
    const json = await res.json() as { ok: boolean; data?: SavedScheduleMeta[]; error?: string };
    return json.ok ? (json.data ?? []) : [];
  } catch {
    return [];
  }
}

export async function loadNamedSave(name: string): Promise<VisitRecord[]> {
  const url = getGasUrl();
  if (!url) throw new Error('GAS URLが未設定です');
  const res = await fetch(`${url}?action=loadSave&name=${encodeURIComponent(name)}`, { redirect: 'follow' });
  const json = await res.json() as { ok: boolean; data?: VisitRecord[]; error?: string };
  if (!json.ok) throw new Error(json.error ?? '読み込みエラー');
  return json.data ?? [];
}

export async function saveNamedSnapshot(name: string, visits: VisitRecord[]): Promise<void> {
  const url = getGasUrl();
  if (!url) throw new Error('GAS URLが未設定です');
  const body = new URLSearchParams();
  body.append('action', 'saveNamed');
  body.append('name', name);
  body.append('data', JSON.stringify(visits));
  const res = await fetch(url, { method: 'POST', body, redirect: 'follow' });
  const json = await res.json() as { ok: boolean; error?: string };
  if (!json.ok) throw new Error(json.error ?? '保存エラー');
}

export async function deleteNamedSave(name: string): Promise<void> {
  const url = getGasUrl();
  if (!url) throw new Error('GAS URLが未設定です');
  const body = new URLSearchParams();
  body.append('action', 'deleteSave');
  body.append('name', name);
  const res = await fetch(url, { method: 'POST', body, redirect: 'follow' });
  const json = await res.json() as { ok: boolean; error?: string };
  if (!json.ok) throw new Error(json.error ?? '削除エラー');
}

export async function sendExportEmail(to: string, filename: string, base64: string): Promise<void> {
  const url = getGasUrl();
  if (!url) throw new Error('GAS URLが未設定です');
  const body = new URLSearchParams();
  body.append('action', 'sendEmail');
  body.append('to', to);
  body.append('filename', filename);
  body.append('base64', base64);
  const res = await fetch(url, { method: 'POST', body, redirect: 'follow' });
  const json = await res.json() as { ok: boolean; error?: string };
  if (!json.ok) throw new Error(json.error ?? 'メール送信エラー');
}

export async function saveToDrive(filename: string, base64: string): Promise<{ url: string; folderName: string }> {
  const url = getGasUrl();
  if (!url) throw new Error('GAS URLが未設定です');
  const body = new URLSearchParams();
  body.append('action', 'saveToDrive');
  body.append('filename', filename);
  body.append('base64', base64);
  const res = await fetch(url, { method: 'POST', body, redirect: 'follow' });
  const json = await res.json() as { ok: boolean; url?: string; folderName?: string; error?: string };
  if (!json.ok) throw new Error(json.error ?? 'Drive保存エラー');
  return { url: json.url ?? '', folderName: json.folderName ?? 'マイドライブ' };
}

export async function fetchServiceTimes(name: string, year: number, month: number): Promise<import('./types').ServiceTimes[]> {
  const url = getGasUrl();
  if (!url) return [];
  try {
    const res = await fetch(
      `${url}?action=getServiceTimes&name=${encodeURIComponent(name)}&year=${year}&month=${month}`,
      { redirect: 'follow' }
    );
    const json = await res.json() as { ok: boolean; data?: import('./types').ServiceTimes[]; error?: string };
    return json.ok ? (json.data ?? []) : [];
  } catch {
    return [];
  }
}

export async function updatePatientTimes(
  row: number,
  times: import('./types').ServiceTimes & { careLevel?: import('./types').CareLevel }
): Promise<void> {
  const url = getGasUrl();
  if (!url) throw new Error('GAS URLが未設定です');
  const body = new URLSearchParams();
  body.append('action', 'updatePatientTimes');
  body.append('row', String(row));
  body.append('times', JSON.stringify(times));
  const res = await fetch(url, { method: 'POST', body, redirect: 'follow' });
  const json = await res.json() as { ok: boolean; error?: string };
  if (!json.ok) throw new Error(json.error ?? '患者時間の更新エラー');
}

export async function fetchRegisteredNames(year: number, month: number): Promise<string[]> {
  const url = getGasUrl();
  if (!url) return [];
  try {
    const res = await fetch(`${url}?action=getRegisteredNames&year=${year}&month=${month}`, { redirect: 'follow' });
    const json = await res.json() as { ok: boolean; data?: string[] };
    return json.ok ? (json.data ?? []) : [];
  } catch {
    return [];
  }
}

// kiroku（登録ページ）GAS：厚生局_訪問時間記録へ書き込む登録エンドポイント
const KIROKU_GAS_URL = 'https://script.google.com/macros/s/AKfycbzrN2tl-V1F17UIJE-4HkL1zmZH8Mz3EDSIVRxtI5ZzfT-CLLb-GhtKkO4Nf64zgHzgjQ/exec';

export async function registerKirokuMonth(payload: {
  name: string; kana: string; kaigo: boolean;
  year: number; month: number; data: unknown[];
}): Promise<void> {
  // no-corsのためレスポンスは読めない（kirokuアプリと同じ方式）
  await fetch(KIROKU_GAS_URL, {
    method: 'POST',
    mode: 'no-cors',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify(payload),
  });
}

export async function fetchEmployees(): Promise<EmployeeList> {
  const url = getGasUrl();
  if (!url) return { doctors: [], hygienists: [] };
  try {
    const res = await fetch(`${url}?action=getEmployees`, { redirect: 'follow' });
    const json = await res.json() as { ok: boolean; data?: EmployeeList; error?: string };
    return json.ok ? (json.data ?? { doctors: [], hygienists: [] }) : { doctors: [], hygienists: [] };
  } catch {
    return { doctors: [], hygienists: [] };
  }
}
