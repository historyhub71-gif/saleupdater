export interface ParsedSalesMessage {
  date: string;
  dsf: string;
  todayValue: number;
}

export interface ExcelUpdateResult {
  fileName: string;
  sheetName: string;
  rowNumber: number;
  columnNumber: number;
  updatedValue: number;
  outputUri: string;
  downloadUrl?: string;
  wasOriginalUpdated: boolean;
  backupUri?: string;
}

export interface SelectedExcelFile {
  uri: string;
  name: string;
  mimeType?: string;
  size?: number;
}

export interface AppState {
  selectedFile: SelectedExcelFile | null;
  pastedMessage: string;
  parsedMessage: ParsedSalesMessage | null;
  statusMessage: string;
  statusStep: string | null;
  errorMessage: string | null;
  successMessage: string | null;
  isLoading: boolean;
  updateResult: ExcelUpdateResult | null;
}
