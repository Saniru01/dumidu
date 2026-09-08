import { AirFreightRecord } from '../types';

export const chartColors = [
  '#4f46e5', '#06b6d4', '#10b981', '#f59e0b', '#ef4444',
  '#8b5cf6', '#ec4899', '#3b82f6', '#14b8a6', '#f97316',
  '#6366f1', '#84cc16', '#a855f7', '#e11d48', '#0ea5e9',
  '#64748b', '#d97706', '#059669', '#2563eb', '#7c3aed'
];

export function findValue(row: Record<string, any>, possibleNames: string[], defaultVal: any = ''): any {
  const keys = Object.keys(row);
  for (const name of possibleNames) {
    const target = name.toLowerCase().replace(/[^a-z0-9]/g, '');
    for (const k of keys) {
      const cleanKey = k.toLowerCase().replace(/[^a-z0-9]/g, '');
      if (cleanKey === target && row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== '') {
        return row[k];
      }
    }
  }
  return defaultVal;
}

export function parseCleanRow(r: Record<string, any>): AirFreightRecord {
  const rawYear = findValue(r, ['No', 'Year', 'Shipment Year'], '2026');
  const year = String(parseInt(rawYear, 10) || rawYear).trim();

  const rawWeek = findValue(r, ['Shipment week', 'Week', 'Shipment Wk', 'Wk'], 0);
  const week = parseInt(rawWeek, 10) || 0;

  const plant = String(findValue(r, ['Column1', 'Plant', 'Plant Name', 'Unit'], 'Unspecified Plant')).trim();
  const dept = String(findValue(r, ['Responsible Department', 'Department', 'Resp Dept', 'Dept'], 'Unassigned')).trim();
  const reason = String(findValue(r, ['Reason for delay', 'Reason', 'Delay Reason'], 'Unspecified')).trim();

  const rawCost = findValue(r, ['Cost (USD)', 'Cost', 'Amount', 'Cost USD'], 0);
  const cost = typeof rawCost === 'number' ? rawCost : parseFloat(String(rawCost).replace(/[^0-9.-]+/g, '')) || 0;

  const month = String(findValue(r, ['Month', 'Shipment Month'], 'Unspecified')).trim();

  return { year, week, plant, dept, reason, cost, month };
}

export function aggregateBy(records: AirFreightRecord[], key: keyof AirFreightRecord): Record<string, number> {
  return records.reduce((acc: Record<string, number>, row: AirFreightRecord) => {
    const val = String(row[key] || 'Other');
    acc[val] = (acc[val] || 0) + (row.cost || 0);
    return acc;
  }, {});
}

export const monthsOrder = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
];
