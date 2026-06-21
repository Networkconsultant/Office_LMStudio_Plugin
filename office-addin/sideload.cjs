/**
 * sideload.cjs — Properly registers and sideloads the add-in into Word.
 * Run with: node sideload.cjs
 */
const path = require("path");
const fs = require("fs");

const devSettingsWindows = require("./node_modules/office-addin-dev-settings/lib/dev-settings-windows");
const sideloadModule     = require("./node_modules/office-addin-dev-settings/lib/sideload");
const { OfficeAddinManifest } = require("./node_modules/office-addin-manifest/lib/manifestOperations");

const MANIFEST_PATH = path.resolve(__dirname, "manifest.xml");

async function main() {
  console.log("Reading manifest:", MANIFEST_PATH);
  const manifest = await OfficeAddinManifest.readManifestFile(MANIFEST_PATH);
  console.log("Add-in ID:", manifest.id);
  console.log("Default URL:", manifest.defaultSettings?.sourceLocation?.defaultValue);

  // 1. Register in registry (sets HKCU\...\WEF\Developer\<id> = manifest path)
  console.log("\nRegistering add-in...");
  await devSettingsWindows.registerAddIn(MANIFEST_PATH);
  console.log("Registered OK");

  // 2. Signal Office to refresh its add-in list
  const { RegistryKey, addBooleanValue } = require("./node_modules/office-addin-dev-settings/lib/registry");
  const refreshKey = new RegistryKey("HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer");
  await addBooleanValue(refreshKey, "RefreshAddins", true);
  console.log("RefreshAddins set");

  // 3. Generate sideload DOCX from the proper template and open Word
  console.log("\nGenerating sideload document...");
  const sideloadFile = await sideloadModule.generateSideloadFile("word", manifest, null);
  console.log("Sideload file:", sideloadFile);

  // 4. Open Word with the sideload DOCX
  const wordExe = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE";
  const { spawn } = require("child_process");
  const proc = spawn(wordExe, [sideloadFile], {
    detached: true,
    stdio: "ignore",
  });
  proc.unref();
  console.log("Word launched, PID:", proc.pid);
  console.log("\nDone! The LMStudio AI task pane should appear on the right side of Word.");
  process.exit(0);
}

main().catch((err) => {
  console.error("Error:", err.message || err);
  process.exit(1);
});
