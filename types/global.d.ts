import * as IntentLauncher from 'expo-intent-launcher';
import * as Sharing from 'expo-sharing';

declare global {
  const IntentLauncher: typeof IntentLauncher;
  const Sharing: typeof Sharing;
}
