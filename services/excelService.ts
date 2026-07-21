import * as FileSystem from "expo-file-system/legacy";

import { SERVER_URL } from "@/constants/app";
import { ExcelUpdateResult, ParsedSalesMessage } from "@/types/sales";

export async function previewWorkbookChange() {
  throw new Error("Preview is now handled by the server.");
}

export async function updateExcelWorkbook(params: {
  uri: string;
  fileName: string;
  parsed: ParsedSalesMessage;
  onStatus?: (step: string) => void;
}): Promise<ExcelUpdateResult> {
  const { onStatus } = params;

  let localUri = params.uri;
  if (localUri.startsWith("http://") || localUri.startsWith("https://")) {
    onStatus?.("Fetching latest workbook...");
    console.log("[excel] Fetching remote file before upload:", localUri);
    const cacheUri = FileSystem.cacheDirectory + params.fileName;
    const downloadResult = await FileSystem.downloadAsync(localUri, cacheUri);
    localUri = downloadResult.uri;
  }

  // Step 1 — Upload
  onStatus?.("Uploading...");
  console.log("[excel] Uploading Excel...", localUri);

  const response = await FileSystem.uploadAsync(
    `${SERVER_URL}/api/excel/update`,
    localUri,
    {
      fieldName: "file",
      httpMethod: "POST",
      uploadType: FileSystem.FileSystemUploadType.MULTIPART,
      parameters: {
        date: params.parsed.date,
        dsf: params.parsed.dsf,
        todayValue: String(params.parsed.todayValue),
      },
    }
  );

  console.log("[excel] Upload status:", response.status);
  console.log("[excel] Upload body:", response.body);

  // Step 2 — Server has processed the file
  onStatus?.("Updating Excel...");

  if (response.status !== 200) {
    throw new Error(response.body || `Server error: ${response.status}`);
  }

  const json = JSON.parse(response.body);

  if (!json.success) {
    throw new Error(json.message || "Server returned an error.");
  }

  // Resolve the download URL — prefer downloadUrl, fall back to output path
  const result = json.result as Record<string, unknown>;

  const downloadUrl: string =
    typeof result.downloadUrl === "string"
      ? result.downloadUrl
      : `${SERVER_URL}/api/excel/download/${encodeURIComponent(
          String(result.output ?? "updated.xlsx")
            .split(/[/\\]/)
            .pop() ?? "updated.xlsx"
        )}`;

  const filename: string =
    downloadUrl.split("/").pop() ?? "updated.xlsx";

  onStatus?.("Completed");

  // Return the remote download URL as outputUri.
  // fileService will cache it on demand when the user taps Open, Share, or Download.
  return {
    ...(result as Partial<ExcelUpdateResult>),
    outputUri: downloadUrl,
    fileName: filename,
    downloadUrl,
  } as ExcelUpdateResult;
}