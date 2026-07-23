import * as FileSystem from "expo-file-system/legacy";
import * as XLSX from "xlsx";
import { Buffer } from "buffer";

if (typeof global.Buffer === "undefined") {
  (global as any).Buffer = Buffer;
}

import { SERVER_URL } from "@/constants/app";
import { ExcelUpdateResult, ParsedSalesMessage } from "@/types/sales";

export async function logWorkbookDetails(uri: string, label: string) {
  try {
    console.log(`=== [Workbook Log: ${label}] ===`);
    console.log(`URI: ${uri}`);

    if (!uri) {
      console.log("URI is empty!");
      console.log(`====================================`);
      return;
    }

    if (uri.startsWith("http://") || uri.startsWith("https://")) {
      console.log("File is remote, skipping local details check");
      console.log(`====================================`);
      return;
    }

    const fileInfo = await FileSystem.getInfoAsync(uri, { md5: true });
    console.log(`File exists: ${fileInfo.exists}`);
    if (!fileInfo.exists) {
      console.log(`====================================`);
      return;
    }
    console.log(`File size: ${fileInfo.size} bytes`);
    console.log(`MD5: ${fileInfo.md5}`);

    const base64 = await FileSystem.readAsStringAsync(uri, {
      encoding: FileSystem.EncodingType.Base64,
    });
    
    // Parse using SheetJS (xlsx)
    const workbook = XLSX.read(base64, { type: "base64" });
    console.log(`Sheet names: ${JSON.stringify(workbook.SheetNames)}`);
    const firstSheetName = workbook.SheetNames[0];
    console.log(`First/Active sheet: ${firstSheetName}`);
    if (firstSheetName) {
      const worksheet = workbook.Sheets[firstSheetName];
      const sheetData = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1 });
      if (sheetData && sheetData.length > 0) {
        console.log(`Header row values: ${JSON.stringify(sheetData[0])}`);
      } else {
        console.log(`Header row values: (empty)`);
      }
    }
    console.log(`====================================`);
  } catch (error) {
    console.error(`[Workbook Log Error] Failed to log details for ${label}:`, error);
  }
}

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

  // Step 1 — Log details before uploading
  await logWorkbookDetails(localUri, "Pre-upload");

  // Step 2 — Upload
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

  let downloadUrl: string =
    typeof result.downloadUrl === "string"
      ? result.downloadUrl
      : `${SERVER_URL}/api/excel/download/${encodeURIComponent(
          String(result.output ?? "updated.xlsx")
            .split(/[/\\]/)
            .pop() ?? "updated.xlsx"
        )}`;

  // Sanitize the URL: Force HTTPS for production server URLs to satisfy Android's network security policy.
  // We keep http:// for local development (localhost, 10.0.2.2, etc.) to support local testing.
  if (downloadUrl.startsWith("http://") && !downloadUrl.includes("localhost") && !downloadUrl.includes("10.0.2.2")) {
    downloadUrl = downloadUrl.replace(/^http:\/\//i, "https://");
  }

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