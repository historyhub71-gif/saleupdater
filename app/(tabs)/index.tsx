import React, { useMemo, useState } from "react";
import {
  Alert,
  StyleSheet,
  Text,
  TextInput,
  View,
} from "react-native";

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
        setState((current) => ({
          ...current,
          selectedFile: file,
          statusMessage: `Selected ${file.name}`,
        }));
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
      Alert.alert(
        "Open Failed",
        error instanceof Error
          ? error.message
          : "Unable to open workbook."
      );
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
      Alert.alert(
        "Share Failed",
        error instanceof Error
          ? error.message
          : "Unable to share workbook."
      );
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
      Alert.alert(
        "Download Failed",
        error instanceof Error
          ? error.message
          : "Unknown error"
      );
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
      setState((current) => ({
        ...current,
        isLoading: true,
        statusStep: null,
        errorMessage: null,
        successMessage: null,
      }));

      const result = await updateExcelWorkbook({
        uri:
          state.updateResult?.outputUri ??
          state.selectedFile!.uri,
        fileName:
          state.updateResult?.fileName ??
          state.selectedFile!.name,
        parsed: state.parsedMessage,
        onStatus: (step) =>
          setState((current) => ({ ...current, statusStep: step })),
      });

      setState((current) => ({
        ...current,
        selectedFile: {
          uri: result.outputUri,
          name: result.fileName,
          mimeType:
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          size: 0,
        },
        updateResult: result,
        statusMessage: "Workbook updated successfully.",
        statusStep: null,
        successMessage: "Excel updated successfully.",
      }));
    } catch (error) {
      setState((current) => ({
        ...current,
        statusStep: null,
        errorMessage:
          error instanceof Error
            ? error.message
            : "Update failed",
        successMessage: null,
      }));

      Alert.alert(
        "Update Failed",
        error instanceof Error
          ? error.message
          : "Unknown error"
      );
    } finally {
      setState((current) => ({
        ...current,
        isLoading: false,
      }));
    }
  };
    return (
    <AppShell title={APP_TITLE} subtitle={APP_SUBTITLE}>
      <ActionCard>
        <PrimaryButton
          title="Select Excel File"
          onPress={handleSelectExcel}
          disabled={state.isLoading}
        />

        <SecondaryButton
          title="Test Server"
          onPress={testServerConnection}
          disabled={state.isLoading}
        />

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
        />

        <SecondaryButton
          title="Parse Message"
          onPress={handleParseMessage}
          disabled={state.isLoading}
        />

        <PrimaryButton
          title="Update Excel"
          onPress={handleUpdateExcel}
          disabled={state.isLoading}
        />
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
                disabled={state.isLoading}
              />

              <SecondaryButton
                title="Download Workbook"
                onPress={handleDownloadWorkbook}
                disabled={state.isLoading}
              />

              <SecondaryButton
                title="Share Workbook"
                onPress={handleShareWorkbook}
                disabled={state.isLoading}
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