const YOUBI_DAY: Record<string, number> = {
  '日': 0, '月': 1, '火': 2, '水': 3, '木': 4, '金': 5, '土': 6,
};

export function normalizeYoubi(youbi: string): string {
  return youbi.trim().replace(/曜日$/, '').replace(/曜$/, '');
}

export function getWeekdayDates(year: number, month: number, youbi: string): string[] {
  const dayNum = YOUBI_DAY[normalizeYoubi(youbi)];
  if (dayNum === undefined) return [];
  const dates: string[] = [];
  const d = new Date(year, month - 1, 1);
  while (d.getMonth() === month - 1) {
    if (d.getDay() === dayNum) dates.push(toIso(d));
    d.setDate(d.getDate() + 1);
  }
  return dates; // 第5週まで含める
}

export function toIso(d: Date): string {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

export function toDisplayDate(iso: string): string {
  const d = new Date(iso + 'T00:00:00');
  const youbi = ['日', '月', '火', '水', '木', '金', '土'][d.getDay()];
  return `${d.getMonth() + 1}/${d.getDate()}(${youbi})`;
}

export function newId(): string {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function pad(n: number): string {
  return String(n).padStart(2, '0');
}

// ── 50音グループ ──────────────────────────────────────────────
export const KANA_GROUPS = ['あ', 'か', 'さ', 'た', 'な', 'は', 'ま', 'や', 'ら', 'わ'] as const;

export function getKanaGroup(kana: string): string {
  if (!kana) return '?';
  let code = kana.charCodeAt(0);
  if (code >= 0x3041 && code <= 0x3096) code += 0x60; // ひらがな→カタカナ
  if (code >= 0x30A1 && code <= 0x30AA) return 'あ';
  if (code >= 0x30AB && code <= 0x30B4) return 'か';
  if (code >= 0x30B5 && code <= 0x30BE) return 'さ';
  if (code >= 0x30BF && code <= 0x30C9) return 'た';
  if (code >= 0x30CA && code <= 0x30CE) return 'な';
  if (code >= 0x30CF && code <= 0x30DD) return 'は';
  if (code >= 0x30DE && code <= 0x30E2) return 'ま';
  if (code >= 0x30E3 && code <= 0x30E8) return 'や';
  if (code >= 0x30E9 && code <= 0x30ED) return 'ら';
  if (code >= 0x30EE && code <= 0x30F4) return 'わ';
  return '?';
}
