export type GasPatient = {
  row: number;
  lastName: string;
  firstName: string;
  name: string;
  youbi: string;
  time: string;
  doctor: string;
  hygienist: string;
};

export type VisitRecord = {
  id: string;
  date: string;
  patientName: string;
  time: string;
  doctor: string;
  hygienist: string;
  isNew: boolean;
  note: string;
};

export type SavedScheduleMeta = {
  name: string;
  count: number;
  createdAt: string;
};

export type EmployeeList = {
  doctors: string[];
  hygienists: string[];
};

export type ModalState =
  | { type: 'closed' }
  | { type: 'editPatient'; patientName: string }
  | { type: 'newPatient' }
  | { type: 'saveAs' }
  | { type: 'saveList' }
  | { type: 'export' };

export type StaffRole = 'dentist' | 'hygienist';
export type Staff = { id: string; name: string; role: StaffRole; active: boolean };
export type ViewMode = '1day' | '3day' | '7day';
