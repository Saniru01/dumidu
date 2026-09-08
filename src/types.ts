export interface AirFreightRecord {
  year: string;
  week: number;
  plant: string;
  dept: string;
  reason: string;
  cost: number;
  month: string;
}

export interface FilterState {
  year: string;
  month: string;
  weekFrom: string;
  weekTo: string;
}

export interface StorageMeta {
  fileName: string;
  uploadDate: string;
}
