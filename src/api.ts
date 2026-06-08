import type { GasPatient, VisitRecord, SavedScheduleMeta, EmployeeList } from './types';

const GAS_URL_KEY = 'gasUrl_hokkyoku_v1';
// 最新GASデプロイURL（clasp管理）
const DEFAULT_GAS_URL = 'https://script.google.com/macros/s/AKfycbwlM2VnVzM3wZZJx_7ayCBeHgd3EKNpOrayUKyl34yu_2fw7838JaH3A2_z0xNpyH7M/exec';

export function getGasUrl(): string {
  // 環境変数 → localStorage → デフォルト の優先順
  return import.meta.env.VITE_GAS_URL || localStorage.getItem(GAS_URL_KEY) || DEFAULT_GAS_URL;
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
