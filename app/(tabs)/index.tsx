import React, { useEffect, useMemo, useState } from "react";
import {
  Alert,
  StyleSheet,
  Text,
  TextInput,
  View,
} from "react-native";
import AsyncStorage from "@react-native-async-storage/async-storage";
import * as FileSystem from "expo-file-system/legacy";

import * as IntentLauncherModule from "expo-intent-launcher";
import * as SharingModule from "expo-sharing";

import {
  ActionCard,
  AppShell,
  PrimaryButton,
  SecondaryButton,
} from "@/components/app-shell";

import {
  APP_SUBTITLE,
  APP_TITLE,
  SERVER_URL,
} from "@/constants/app";

import {
  updateExcelWorkbook,
  logWorkbookDetails,
} from "@/services/excelService";

import {
  parseSalesMessage,
} from "@/services/parserService";

import {
  AppState,
} from "@/types/sales";

import {
  downloadWorkbook,
  openWorkbook,
  pickExcelFile,
  shareWorkbook,
} from "@/services/fileService";


if (typeof global !== "undefined") {
  (global as any).IntentLauncher = IntentLauncherModule;
  (global as any).Sharing = SharingModule;
}

const STORAGE_KEY_URI = "@SalesUpdater:workbook_uri";
const STORAGE_KEY_NAME = "@SalesUpdater:workbook_name";
const STORAGE_KEY_CONFIRMED_TS = "@SalesUpdater:workbook_confirmed_ts";
const STORAGE_KEY_UPDATED_TS = "@SalesUpdater:workbook_updated_ts";

async function savePersistedWorkbook(params: {
  uri: string;
  name: string;
  confirmedTimestamp?: number;
  updatedTimestamp?: number;
}) {
  try {
    await AsyncStorage.setItem(STORAGE_KEY_URI, params.uri);
    await AsyncStorage.setItem(STORAGE_KEY_NAME, params.name);
    if (params.confirmedTimestamp !== undefined) {
      await AsyncStorage.setItem(STORAGE_KEY_CONFIRMED_TS, String(params.confirmedTimestamp));
    }
    if (params.updatedTimestamp !== undefined) {
      await AsyncStorage.setItem(STORAGE_KEY_UPDATED_TS, String(params.updatedTimestamp));
    }
  } catch (error) {
    console.error("[Storage] Failed to save workbook metadata", error);
  }
}

async function getPersistedWorkbook() {
  try {
    const uri = await AsyncStorage.getItem(STORAGE_KEY_URI);
    const name = await AsyncStorage.getItem(STORAGE_KEY_NAME);
    const confStr = await AsyncStorage.getItem(STORAGE_KEY_CONFIRMED_TS);
    const updStr = await AsyncStorage.getItem(STORAGE_KEY_UPDATED_TS);
    
    return uri && name ? {
      uri,
      name,
      confirmedTimestamp: confStr ? parseInt(confStr, 10) : null,
      updatedTimestamp: updStr ? parseInt(updStr, 10) : null,
    } : null;
  } catch (error) {
    console.error("[Storage] Failed to load workbook metadata", error);
    return null;
  }
}

async function clearPersistedWorkbook() {
  try {
    const uri = await AsyncStorage.getItem(STORAGE_KEY_URI);
    if (uri) {
      try {
        const info = await FileSystem.getInfoAsync(uri);
        if (info.exists) {
          await FileSystem.deleteAsync(uri, { idempotent: true });
        }
      } catch (err) {
        console.warn("[Storage] Failed to delete file during clear", err);
      }
    }
    await AsyncStorage.removeItem(STORAGE_KEY_URI);
    await AsyncStorage.removeItem(STORAGE_KEY_NAME);
    await AsyncStorage.removeItem(STORAGE_KEY_CONFIRMED_TS);
    await AsyncStorage.removeItem(STORAGE_KEY_UPDATED_TS);
  } catch (error) {
    console.error("[Storage] Failed to clear workbook metadata", error);
  }
}

/**
 * Persists the workbook file in FileSystem.documentDirectory and cleans up any old persistent file.
 */
async function persistAndCleanupWorkbook(
  newUri: string,
  fileName: string,
  oldPersistentUri: string | null
): Promise<string> {
  const fileExt = fileName.split('.').pop() ?? 'xlsx';
  const targetLocalUri = `${FileSystem.documentDirectory}workbook_${Date.now()}.${fileExt}`;

  console.log(`[Storage] Persisting workbook to: ${targetLocalUri}`);

  if (newUri.startsWith("http://") || newUri.startsWith("https://")) {
    await FileSystem.downloadAsync(newUri, targetLocalUri);
  } else {
    await FileSystem.copyAsync({
      from: newUri,
      to: targetLocalUri,
    });
  }

  if (oldPersistentUri && oldPersistentUri !== targetLocalUri) {
    try {
      const oldInfo = await FileSystem.getInfoAsync(oldPersistentUri);
      if (oldInfo.exists) {
        console.log(`[Storage] Cleaning up old workbook: ${oldPersistentUri}`);
        await FileSystem.deleteAsync(oldPersistentUri, { idempotent: true });
      }
    } catch (err) {
      console.warn(`[Storage] Failed to delete old workbook: ${oldPersistentUri}`, err);
    }
  }

  return targetLocalUri;
}

function formatTimestamp(ts: number | null): string {
  if (!ts) return "Never updated";
  const date = new Date(ts);
  const pad = (n: number) => String(n).padStart(2, "0");
  const yyyy = date.getFullYear();
  const mm = pad(date.getMonth() + 1);
  const dd = pad(date.getDate());
  const hh = pad(date.getHours());
  const min = pad(date.getMinutes());
  return `${yyyy}-${mm}-${dd} ${hh}:${min}`;
}

const initialState: AppState = {
  selectedFile: null,
  pastedMessage: "",
  parsedMessage: null,
  statusMessage: "Ready to process a new sales update.",
  statusStep: null,
  errorMessage: null,
  successMessage: null,
  isLoading: false,
  updateResult: null,
};

export default function HomeScreen() {
  const [state, setState] = useState<AppState>(initialState);
  const [pendingReminder, setPendingReminder] = useState(false);
  const [isStorageLoading, setIsStorageLoading] = useState(true);
  const [workbookNotFound, setWorkbookNotFound] = useState(false);
  const [lastUpdatedTime, setLastUpdatedTime] = useState<number | null>(null);
  const [isUpdating, setIsUpdating] = useState(false);

  const isAnyLoading = state.isLoading || isUpdating;

  // Helper to handle missing or corrupted files
  const handleFileError = async (error: unknown) => {
    console.error("[Workbook Error]", error);
    await clearPersistedWorkbook();
    setWorkbookNotFound(true);
    setLastUpdatedTime(null);
    setPendingReminder(false);
    setIsUpdating(false);

    setState((current) => ({
      ...current,
      selectedFile: null,
      updateResult: null,
      isLoading: false,
      statusStep: null,
      statusMessage: "Workbook not found. Please select your workbook.",
      errorMessage: "Workbook not found",
    }));

    Alert.alert(
      "Workbook Not Found",
      "Workbook not found. Please select your workbook."
    );
  };

  // Load saved workbook on mount
  useEffect(() => {
    const loadSavedWorkbook = async () => {
      try {
        setIsStorageLoading(true);
        const saved = await getPersistedWorkbook();
        if (saved) {
          // Check if the file actually exists
          let exists = true;
          try {
            const fileInfo = await FileSystem.getInfoAsync(saved.uri);
            if (!fileInfo.exists) {
              exists = false;
            }
          } catch (err) {
            console.error("[Storage] Error checking file existence:", err);
            exists = false;
          }

          if (!exists) {
            await clearPersistedWorkbook();
            setWorkbookNotFound(true);
            setState((current) => ({
              ...current,
              selectedFile: null,
              statusMessage: "Workbook not found. Please select your workbook.",
            }));
            return;
          }

          // Log loaded workbook from storage on mount
          await logWorkbookDetails(saved.uri, "Loaded from Storage (Mount)");

          // Check if calendar month changed or if 25+ days passed using confirmedTimestamp
          const refDate = saved.confirmedTimestamp ? new Date(saved.confirmedTimestamp) : new Date();
          const currentDate = new Date();

          const isDifferentMonth =
            refDate.getFullYear() !== currentDate.getFullYear() ||
            refDate.getMonth() !== currentDate.getMonth();

          const daysPassed = saved.confirmedTimestamp
            ? (currentDate.getTime() - saved.confirmedTimestamp) / (1000 * 60 * 60 * 24)
            : 0;
          const isOverdue = daysPassed >= 25;

          setState((current) => ({
            ...current,
            selectedFile: {
              uri: saved.uri,
              name: saved.name,
              mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              size: 0,
            },
            statusMessage: "Workbook loaded successfully.",
          }));

          setLastUpdatedTime(saved.updatedTimestamp);
          setWorkbookNotFound(false);

          if (isDifferentMonth || isOverdue) {
            setPendingReminder(true);
          }
        }
      } catch (err) {
        console.error("[Storage] Error loading saved workbook:", err);
      } finally {
        setIsStorageLoading(false);
      }
    };

    loadSavedWorkbook();
  }, []);

  const handleContinueWorkbook = async () => {
    if (state.selectedFile) {
      const currentTimestamp = Date.now();
      // Confirm and reset timestamp to current time (resets overdue & month checks)
      await savePersistedWorkbook({
        uri: state.selectedFile.uri,
        name: state.selectedFile.name,
        confirmedTimestamp: currentTimestamp,
      });
      setPendingReminder(false);
      setState((current) => ({
        ...current,
        statusMessage: `Confirmed workbook: ${state.selectedFile!.name}`,
      }));
    }
  };

  const handleChangeWorkbook = async () => {
    await clearPersistedWorkbook();
    setWorkbookNotFound(false);
    setLastUpdatedTime(null);
    setPendingReminder(false);
    setState((current) => ({
      ...current,
      selectedFile: null,
      updateResult: null,
      statusMessage: "Please select a new workbook.",
    }));
    await handleSelectExcel();
  };

  const testServerConnection = async () => {
    try {
      const response = await fetch(
        `${SERVER_URL}/`
      );

      const text = await response.text();

      Alert.alert("Server", text);
    } catch (error) {
      Alert.alert(
        "Connection Failed",
        error instanceof Error
          ? error.message
          : String(error)
      );
    }
  };

  const statusItems = useMemo(() => {
    return [
      {
        label: "Selected Excel File",
        value:
          state.selectedFile?.name ??
          "None Selected",
      },
      {
        label: "Parsed Date",
        value:
          state.parsedMessage?.date ??
          "—",
      },
      {
        label: "Parsed DSF",
        value:
          state.parsedMessage?.dsf ??
          "—",
      },
      {
        label: "Today's Value",
        value: state.parsedMessage
          ? `${state.parsedMessage.todayValue}`
          : "—",
      },
    ];
  }, [
    state.selectedFile,
    state.parsedMessage,
  ]);
  const handleSelectExcel = async () => {
    try {
      setState((current) => ({
        ...current,
        isLoading: true,
        errorMessage: null,
        successMessage: null,
      }));

      const file = await pickExcelFile();

      if (file) {
        // Log original selected workbook
        await logWorkbookDetails(file.uri, "Selected by Document Picker");

        // Copy selected file to the persistent directory & cleanup old persistent copy
        const oldSaved = await getPersistedWorkbook();
        const persistentUri = await persistAndCleanupWorkbook(
          file.uri,
          file.name,
          oldSaved?.uri ?? null
        );

        // Log workbook after copy to persistent storage
        await logWorkbookDetails(persistentUri, "Copied to Persistent Storage");

        const currentTimestamp = Date.now();

        // Save metadata
        await savePersistedWorkbook({
          uri: persistentUri,
          name: file.name,
          confirmedTimestamp: currentTimestamp,
          updatedTimestamp: currentTimestamp,
        });

        setState((current) => ({
          ...current,
          selectedFile: {
            ...file,
            uri: persistentUri,
          },
          statusMessage: `Selected ${file.name}`,
        }));

        setLastUpdatedTime(currentTimestamp);
        setPendingReminder(false);
        setWorkbookNotFound(false);
      } else {
        setState((current) => ({
          ...current,
          statusMessage: "No Excel file selected.",
          errorMessage: "Excel not selected",
        }));
      }
    } catch (error) {
      setState((current) => ({
        ...current,
        statusMessage: "Unable to select Excel file.",
        errorMessage:
          error instanceof Error
            ? error.message
            : "Excel selection failed",
      }));
    } finally {
      setState((current) => ({
        ...current,
        isLoading: false,
      }));
    }
  };

  const handleParseMessage = () => {
    try {
      const parsed = parseSalesMessage(state.pastedMessage);

      setState((current) => ({
        ...current,
        parsedMessage: parsed,
        statusMessage: "Message parsed successfully.",
        errorMessage: null,
        successMessage: null,
      }));
    } catch (error) {
      setState((current) => ({
        ...current,
        parsedMessage: null,
        statusMessage: "Unable to parse message.",
        errorMessage:
          error instanceof Error
            ? error.message
            : "Invalid WhatsApp message",
      }));
    }
  };

  const handleOpenWorkbook = async () => {
    const uri =
      state.updateResult?.outputUri ??
      state.selectedFile?.uri;

    if (!uri) {
      Alert.alert(
        "No Workbook",
        "Please select a workbook first."
      );
      return;
    }

    try {
      await openWorkbook(uri);
    } catch (error) {
      await handleFileError(error);
    }
  };
  const handleShareWorkbook = async () => {
    const uri =
      state.updateResult?.outputUri ??
      state.selectedFile?.uri;

    const fileName =
      state.updateResult?.fileName ??
      state.selectedFile?.name ??
      "Workbook.xlsx";

    if (!uri) {
      Alert.alert(
        "No Workbook",
        "Please select a workbook first."
      );
      return;
    }

    try {
      await shareWorkbook(uri, fileName);
    } catch (error) {
      await handleFileError(error);
    }
  };

  const handleDownloadWorkbook = async () => {
    const uri =
      state.updateResult?.outputUri ??
      state.selectedFile?.uri;

    const fileName =
      state.updateResult?.fileName ??
      state.selectedFile?.name ??
      "Workbook.xlsx";

    if (!uri) {
      Alert.alert(
        "No Workbook",
        "Please update workbook first."
      );
      return;
    }

    try {
      await downloadWorkbook(uri, fileName);
    } catch (error) {
      await handleFileError(error);
    }
  };

  const handleUpdateExcel = async () => {
    if (!state.selectedFile) {
      Alert.alert(
        "Excel Missing",
        "Please select an Excel file."
      );
      return;
    }

    if (!state.parsedMessage) {
      Alert.alert(
        "Message Missing",
        "Please parse the WhatsApp message."
      );
      return;
    }

    try {
      setIsUpdating(true);
      setState((current) => ({
        ...current,
        isLoading: true,
        statusStep: null,
        errorMessage: null,
        successMessage: null,
      }));

      // 1. Workbook existence check on the CURRENT persistent workbook
      const savedWorkbook = await getPersistedWorkbook();
      if (!savedWorkbook) {
        throw new Error("Persistent workbook missing from storage");
      }

      // Check if the file actually exists on disk before we proceed to update it
      const fileInfo = await FileSystem.getInfoAsync(savedWorkbook.uri);
      if (!fileInfo.exists) {
        throw new Error("Local workbook file not found on disk");
      }

      // Log loaded workbook right before starting update processing
      await logWorkbookDetails(savedWorkbook.uri, "Loaded for Update");

      const result = await updateExcelWorkbook({
        uri: savedWorkbook.uri,
        fileName: savedWorkbook.name,
        parsed: state.parsedMessage,
        onStatus: (step) =>
          setState((current) => ({ ...current, statusStep: step })),
      });

      // Move the returned updated file into the persistent directory, deleting the old local file ONLY after the new file is successfully written
      const persistentUri = await persistAndCleanupWorkbook(
        result.outputUri,
        result.fileName,
        savedWorkbook.uri
      );

      const updateTimestamp = Date.now();

      // Update state with newly returned persistent outputUri
      setState((current) => ({
        ...current,
        selectedFile: {
          uri: persistentUri,
          name: result.fileName,
          mimeType:
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          size: 0,
        },
        updateResult: {
          ...result,
          outputUri: persistentUri,
        },
        statusMessage: "Workbook updated successfully.",
        successMessage: "Excel updated successfully.",
      }));

      // Replace remembered workbook in storage, keeping confirmedTimestamp intact
      await savePersistedWorkbook({
        uri: persistentUri,
        name: result.fileName,
        updatedTimestamp: updateTimestamp,
      });

      setLastUpdatedTime(updateTimestamp);

    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      const lowerMsg = errorMsg.toLowerCase();

      // Classify if the error is a genuine local filesystem error
      const isLocalFileError =
        lowerMsg.includes("enoent") ||
        lowerMsg.includes("file not found") ||
        lowerMsg.includes("no such file") ||
        lowerMsg.includes("unable to read local workbook") ||
        lowerMsg.includes("persistent workbook missing");

      if (isLocalFileError) {
        await handleFileError(error);
        return;
      }

      setState((current) => ({
        ...current,
        errorMessage: errorMsg,
        successMessage: null,
      }));

      Alert.alert(
        "Update Failed",
        errorMsg
      );
    } finally {
      setIsUpdating(false);
      setState((current) => ({
        ...current,
        isLoading: false,
        statusStep: null,
      }));
    }
  };
    return (
    <AppShell title={APP_TITLE} subtitle={APP_SUBTITLE}>
      <ActionCard>
        {isStorageLoading ? (
          <Text style={styles.loadingText}>Loading saved workbook...</Text>
        ) : workbookNotFound ? (
          <View style={styles.errorSection}>
            <Text style={styles.errorHeader}>Workbook not found</Text>
            <Text style={styles.errorSubheader}>Please select your workbook.</Text>
            <PrimaryButton
              title="Select Workbook"
              onPress={handleSelectExcel}
              disabled={isAnyLoading}
            />
          </View>
        ) : state.selectedFile ? (
          <View style={styles.workbookSection}>
            <Text style={styles.workbookHeader}>Current Workbook</Text>
            <Text style={styles.workbookName}>{state.selectedFile.name}</Text>

            <View style={styles.workbookInfoRow}>
              <Text style={styles.infoLabel}>Last Updated</Text>
              <Text style={styles.infoValue}>{formatTimestamp(lastUpdatedTime)}</Text>
            </View>

            <View style={styles.workbookStatusRow}>
              <Text style={styles.statusLabel}>Status</Text>
              <Text style={styles.statusValueReady}>Ready</Text>
            </View>

            {pendingReminder && (
              <View style={styles.reminderContainer}>
                <Text style={styles.reminderMsgText}>
                  A new monthly workbook may be available.
                </Text>
                <View style={styles.workbookButtons}>
                  <PrimaryButton
                    title="Continue Current Workbook"
                    onPress={handleContinueWorkbook}
                    disabled={isAnyLoading}
                  />
                  <SecondaryButton
                    title="Change Workbook"
                    onPress={handleChangeWorkbook}
                    disabled={isAnyLoading}
                  />
                </View>
              </View>
            )}

            {!pendingReminder && (
              <View style={styles.workbookButtons}>
                <SecondaryButton
                  title="Change Workbook"
                  onPress={handleChangeWorkbook}
                  disabled={isAnyLoading}
                />
              </View>
            )}
          </View>
        ) : (
          <PrimaryButton
            title="Select Excel File"
            onPress={handleSelectExcel}
            disabled={isAnyLoading}
          />
        )}

        <SecondaryButton
          title="Test Server"
          onPress={testServerConnection}
          disabled={isAnyLoading}
        />

        {/* Form fields are fully accessible when workbook is loaded (even during reminder) */}
        {!isStorageLoading && !workbookNotFound && state.selectedFile && (
          <>
            <TextInput
              style={styles.textArea}
              multiline
              numberOfLines={10}
              value={state.pastedMessage}
              onChangeText={(value) =>
                setState((current) => ({
                  ...current,
                  pastedMessage: value,
                }))
              }
              placeholder="Paste WhatsApp Message Here"
              textAlignVertical="top"
              editable={!isAnyLoading}
            />

            <SecondaryButton
              title="Parse Message"
              onPress={handleParseMessage}
              disabled={isAnyLoading}
            />

            <PrimaryButton
              title="Update Excel"
              onPress={handleUpdateExcel}
              disabled={isAnyLoading}
            />
          </>
        )}
      </ActionCard>

      <ActionCard>
        <Text style={styles.sectionTitle}>Status</Text>

        {statusItems.map((item) => (
          <View key={item.label} style={styles.statusRow}>
            <Text style={styles.statusLabel}>
              {item.label}
            </Text>

            <Text style={styles.statusValue}>
              {item.value}
            </Text>
          </View>
        ))}

        {state.statusStep ? (
          <Text style={styles.stepText}>
            {state.statusStep}
          </Text>
        ) : (
          <Text style={styles.statusText}>
            {state.statusMessage}
          </Text>
        )}

        {state.errorMessage ? (
          <Text style={styles.errorText}>
            {state.errorMessage}
          </Text>
        ) : null}

        {state.successMessage ? (
          <Text style={styles.successText}>
            {state.successMessage}
          </Text>
        ) : null}

        {state.updateResult && (
          <View style={styles.updateContainer}>
            <View style={styles.actionButtons}>
              <PrimaryButton
                title="Open Workbook"
                onPress={handleOpenWorkbook}
                disabled={isAnyLoading}
              />

              <SecondaryButton
                title="Download Workbook"
                onPress={handleDownloadWorkbook}
                disabled={isAnyLoading}
              />

              <SecondaryButton
                title="Share Workbook"
                onPress={handleShareWorkbook}
                disabled={isAnyLoading}
              />
            </View>
          </View>
        )}
      </ActionCard>
    </AppShell>
  );
}
const styles = StyleSheet.create({
  textArea: {
    minHeight: 180,
    borderWidth: 1,
    borderColor: "#d7e0eb",
    borderRadius: 14,
    padding: 12,
    backgroundColor: "#f9fbff",
    marginBottom: 12,
    color: "#162033",
  },

  sectionTitle: {
    fontSize: 17,
    fontWeight: "800",
    color: "#162033",
    marginBottom: 10,
  },

  loadingText: {
    fontSize: 15,
    color: "#5b6472",
    textAlign: "center",
    marginVertical: 12,
    fontStyle: "italic",
  },

  errorSection: {
    padding: 16,
    backgroundColor: "#fff5f5",
    borderRadius: 16,
    borderWidth: 1,
    borderColor: "#feb2b2",
    marginBottom: 14,
    alignItems: "center",
  },

  errorHeader: {
    fontSize: 16,
    fontWeight: "700",
    color: "#c53030",
    marginBottom: 4,
  },

  errorSubheader: {
    fontSize: 14,
    color: "#742a2a",
    marginBottom: 14,
    textAlign: "center",
  },

  workbookSection: {
    marginBottom: 14,
    padding: 14,
    backgroundColor: "#f4f8ff",
    borderRadius: 16,
    borderWidth: 1,
    borderColor: "#cbdffb",
  },

  workbookHeader: {
    fontSize: 12,
    fontWeight: "700",
    color: "#475569",
    textTransform: "uppercase",
    letterSpacing: 0.5,
    marginBottom: 4,
  },

  workbookName: {
    fontSize: 18,
    fontWeight: "800",
    color: "#0f172a",
    marginBottom: 8,
  },

  workbookInfoRow: {
    flexDirection: "row",
    justifyContent: "space-between",
    paddingVertical: 4,
    marginBottom: 4,
  },

  infoLabel: {
    fontSize: 13,
    color: "#475569",
  },

  infoValue: {
    fontSize: 13,
    fontWeight: "600",
    color: "#0f172a",
  },

  workbookStatusRow: {
    flexDirection: "row",
    alignItems: "center",
    marginBottom: 12,
  },

  statusValueReady: {
    fontSize: 13,
    fontWeight: "700",
    color: "#15803d",
    marginLeft: 8,
    backgroundColor: "#dcfce7",
    paddingHorizontal: 8,
    paddingVertical: 2,
    borderRadius: 6,
    overflow: "hidden",
  },

  reminderContainer: {
    marginTop: 8,
    backgroundColor: "#fffbeb",
    borderColor: "#fef3c7",
    borderWidth: 1,
    borderRadius: 12,
    padding: 10,
    marginBottom: 10,
  },

  reminderMsgText: {
    fontSize: 14,
    color: "#b45309",
    backgroundColor: "#fef3c7",
    padding: 10,
    borderRadius: 10,
    marginBottom: 12,
    fontWeight: "600",
  },

  workbookButtons: {
    marginTop: 4,
    gap: 8,
  },

  statusRow: {
    flexDirection: "row",
    justifyContent: "space-between",
    paddingVertical: 6,
    gap: 10,
  },

  statusLabel: {
    flex: 1,
    fontSize: 14,
    color: "#5b6472",
  },

  statusValue: {
    flex: 1,
    fontSize: 14,
    fontWeight: "600",
    color: "#162033",
    textAlign: "right",
  },

  statusText: {
    marginTop: 8,
    color: "#4a5568",
  },

  stepText: {
    marginTop: 8,
    color: "#3b82f6",
    fontWeight: "600",
    fontStyle: "italic",
  },

  errorText: {
    marginTop: 8,
    color: "#e53e3e",
    fontWeight: "700",
  },

  successText: {
    marginTop: 8,
    color: "#2f855a",
    fontWeight: "700",
  },

  updateContainer: {
    marginTop: 12,
  },

  actionButtons: {
    marginTop: 12,
    gap: 10,
  },
});