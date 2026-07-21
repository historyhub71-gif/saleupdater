import { SelectedExcelFile } from '@/types/sales';
import * as DocumentPicker from 'expo-document-picker';
import { File as ExpoFile } from 'expo-file-system';
import * as FileSystem from 'expo-file-system/legacy';
import { StorageAccessFramework } from 'expo-file-system/legacy';
import * as IntentLauncher from 'expo-intent-launcher';
import * as Sharing from 'expo-sharing';
import { Alert, Platform } from 'react-native';

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/**
 * If the uri is a remote https:// URL, download it to the app cache and
 * return the local file:// path. Otherwise return the uri unchanged.
 */
async function ensureLocalUri(uri: string, fileName: string): Promise<string> {
  if (!uri.startsWith('http://') && !uri.startsWith('https://')) {
    return uri;
  }
  console.log('[file] Caching remote file', { uri });
  const cacheUri = FileSystem.cacheDirectory + fileName;
  const { uri: localUri } = await FileSystem.downloadAsync(uri, cacheUri);
  console.log('[file] Cached to', localUri);
  return localUri;
}

export async function pickExcelFile(): Promise<SelectedExcelFile | null> {
  try {
    const result = await DocumentPicker.getDocumentAsync({
      type: [XLSX_MIME],
      copyToCacheDirectory: true,
      multiple: false,
    });

    if (result.canceled || !result.assets?.[0]?.uri) {
      return null;
    }

    const asset = result.assets[0];
    console.log('[file] Selected URI', asset.uri);
    const fileUri = asset.uri.startsWith('file://')
      ? asset.uri
      : FileSystem.cacheDirectory + (asset.name ?? 'temp.xlsx');

    return {
      uri: fileUri,
      name: asset.name ?? 'selected-workbook.xlsx',
      mimeType: asset.mimeType ?? XLSX_MIME,
      size: asset.size,
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unable to pick the Excel file.';
    console.error('[file] Picker failed', message);
    throw new Error(message);
  }
}

export async function saveBase64ToUri(base64: string, uri: string): Promise<void> {
  try {
    new ExpoFile(uri).write(base64, { encoding: 'base64' });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unable to write the workbook to disk.';
    console.error('[file] Write failed', message);
    throw new Error(message);
  }
}

export async function openWorkbook(uri: string): Promise<void> {
  try {
    console.log('[file] Opening workbook', { uri });
    // Resolve remote URL to a local file first
    const fileName = uri.split('/').pop() ?? 'workbook.xlsx';
    let targetUri = await ensureLocalUri(uri, fileName);

    if (Platform.OS === 'android' && targetUri.startsWith('file://')) {
      targetUri = await FileSystem.getContentUriAsync(targetUri);
      console.log('[file] Converted to content URI', { targetUri });
    }

    if (targetUri.startsWith('file://') || targetUri.startsWith('content://')) {
      await IntentLauncher.startActivityAsync('android.intent.action.VIEW', {
        data: targetUri,
        flags: 1, // FLAG_GRANT_READ_URI_PERMISSION
        type: XLSX_MIME,
      });
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unable to open the workbook.';
    console.error('[file] Open failed', message);
    throw new Error(message);
  }
}

export async function shareWorkbook(uri: string, fileName: string): Promise<void> {
  try {
    console.log('[file] Sharing workbook', { uri, fileName });
    if (!(await Sharing.isAvailableAsync())) {
      throw new Error('Sharing is not available on this device.');
    }

    // Resolve remote URL to a local file first
    const localUri = await ensureLocalUri(uri, fileName);

    await Sharing.shareAsync(localUri, {
      mimeType: XLSX_MIME,
      UTI: 'com.microsoft.excel.xlsx',
      dialogTitle: `Share ${fileName}`,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unable to share the workbook.';
    console.error('[file] Share failed', message);
    throw new Error(message);
  }
}

export async function downloadWorkbook(
  sourceUri: string,
  fileName: string
): Promise<void> {
  try {
    const permission =
      await StorageAccessFramework.requestDirectoryPermissionsAsync();

    if (!permission.granted) {
      // User cancelled the folder picker — do nothing silently
      return;
    }

    const dirUri = permission.directoryUri;

    // Cloud storage providers (Google Drive, Dropbox, OneDrive, etc.)
    // are NOT writable via SAF — detect them and guide the user.
    const CLOUD_PROVIDERS = [
      'com.google.android.apps.docs',
      'com.dropbox.android',
      'com.microsoft.skydrive',
      'com.box.android',
      'com.google.android.apps.photos',
    ];
    const isCloud = CLOUD_PROVIDERS.some((pkg) => dirUri.includes(pkg));

    if (isCloud) {
      Alert.alert(
        'Choose Phone Storage',
        'Cloud folders like Google Drive cannot be used here.\n\nPlease select a folder on your phone, such as:\n• Downloads\n• Documents\n• Internal Storage',
        [{ text: 'OK' }]
      );
      return;
    }

    // Resolve remote URL to a local cached file first
    const localUri = await ensureLocalUri(sourceUri, fileName);

    const base64 = await FileSystem.readAsStringAsync(localUri, {
      encoding: FileSystem.EncodingType.Base64,
    });

    const destinationUri = await StorageAccessFramework.createFileAsync(
      dirUri,
      fileName,
      XLSX_MIME
    );

    await FileSystem.writeAsStringAsync(destinationUri, base64, {
      encoding: FileSystem.EncodingType.Base64,
    });

    Alert.alert('Download Complete', `${fileName} saved successfully.`);
  } catch (error) {
    const message =
      error instanceof Error ? error.message : 'Download failed.';

    console.error('[file] Download failed', message);
    throw new Error(message);
  }
}
