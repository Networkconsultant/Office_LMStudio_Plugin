#!/usr/bin/env node
/**
 * LMStudio Office Plugin — Project Setup Script
 * Run once with: node setup.js
 * This creates the full project structure for both the Office Add-in and VSTO add-in.
 */
const fs = require("fs");
const path = require("path");

const ROOT = __dirname;

function write(relPath, content) {
  const fullPath = path.join(ROOT, relPath);
  fs.mkdirSync(path.dirname(fullPath), { recursive: true });
  fs.writeFileSync(fullPath, content, "utf8");
  console.log("  ✅ " + relPath);
}

console.log("\n🚀 Generating Office + LMStudio Plugin project...\n");

// ─────────────────────────────────────────────────────────────────────────────
// OFFICE ADD-IN
// ─────────────────────────────────────────────────────────────────────────────

write("office-addin/package.json", JSON.stringify({
  name: "office-lmstudio-addin",
  version: "1.0.0",
  description: "Office Add-in for LMStudio local AI models (Word, Excel, Outlook, PowerPoint)",
  scripts: {
    build: "webpack --mode production",
    "build:dev": "webpack --mode development",
    "dev-server": "webpack serve --mode development --open",
    start: "office-addin-debugging start manifest.xml",
    stop: "office-addin-debugging stop manifest.xml",
    validate: "office-addin-manifest validate manifest.xml",
    certs: "office-addin-dev-certs install"
  },
  dependencies: {
    react: "^18.2.0",
    "react-dom": "^18.2.0"
  },
  devDependencies: {
    "@types/office-js": "^1.0.396",
    "@types/react": "^18.2.0",
    "@types/react-dom": "^18.2.0",
    "css-loader": "^6.8.1",
    "html-webpack-plugin": "^5.5.3",
    "office-addin-debugging": "^5.0.20",
    "office-addin-dev-certs": "^1.9.1",
    "office-addin-manifest": "^1.8.14",
    "style-loader": "^3.3.3",
    "ts-loader": "^9.5.0",
    typescript: "^5.2.2",
    webpack: "^5.88.2",
    "webpack-cli": "^5.1.4",
    "webpack-dev-server": "^4.15.1"
  }
}, null, 2));

write("office-addin/tsconfig.json", JSON.stringify({
  compilerOptions: {
    target: "ES2017",
    lib: ["ES2017", "DOM"],
    module: "commonjs",
    jsx: "react",
    strict: false,
    moduleResolution: "node",
    esModuleInterop: true,
    skipLibCheck: true,
    outDir: "./dist"
  },
  include: ["src/**/*"],
  exclude: ["node_modules"]
}, null, 2));

write("office-addin/webpack.config.js", `const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

module.exports = async () => {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return {
    entry: {
      taskpane: "./src/taskpane/taskpane.tsx",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js"],
    },
    module: {
      rules: [
        {
          test: /\\.tsx?$/,
          use: "ts-loader",
          exclude: /node_modules/,
        },
        {
          test: /\\.css$/,
          use: ["style-loader", "css-loader"],
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: "./src/taskpane/index.html",
        filename: "taskpane.html",
        chunks: ["taskpane"],
      }),
    ],
    devServer: {
      port: 3000,
      server: {
        type: "https",
        options: httpsOptions,
      },
      headers: { "Access-Control-Allow-Origin": "*" },
    },
  };
};
`);

// ─── manifest.xml (Word + Excel + PowerPoint) ────────────────────────────────
write("office-addin/manifest.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">

  <Id>7b6ea76a-5c3d-4e8f-a1b2-3c4d5e6f7a8b</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>LMStudio Local AI</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="LMStudio AI Assistant"/>
  <Description DefaultValue="Chat with local LMStudio AI models directly in Office"/>
  <IconUrl DefaultValue="https://localhost:3000/assets/icon-80.png"/>
  <HighResolutionIconUrl DefaultValue="https://localhost:3000/assets/icon-80.png"/>
  <SupportUrl DefaultValue="https://localhost:3000"/>

  <AppDomains>
    <AppDomain>https://localhost:3000</AppDomain>
  </AppDomains>

  <Hosts>
    <Host Name="Document"/>      <!-- Word -->
    <Host Name="Workbook"/>      <!-- Excel -->
    <Host Name="Presentation"/>  <!-- PowerPoint -->
  </Hosts>

  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>
  </DefaultSettings>

  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>

      <!-- ── Word ─────────────────────────────────────────────────── -->
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Taskpane.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Word.CommandsGroup">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="Word.TaskpaneButton">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>LMStudioTaskpane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>

      <!-- ── Excel ────────────────────────────────────────────────── -->
      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Taskpane.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Excel.CommandsGroup">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="Excel.TaskpaneButton">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>LMStudioTaskpane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>

      <!-- ── PowerPoint ────────────────────────────────────────────── -->
      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Taskpane.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="PPT.CommandsGroup">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="PPT.TaskpaneButton">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>LMStudioTaskpane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>

    </Hosts>

    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://localhost:3000"/>
        <bt:Url id="LearnMoreUrl" DefaultValue="https://localhost:3000"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="LMStudio AI Assistant"/>
        <bt:String id="CommandsGroup.Label" DefaultValue="LMStudio AI"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Open AI Assistant"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Chat with local LMStudio models, transform text, and more."/>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Open the LMStudio AI Assistant panel."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`);

// ─── manifest-outlook.xml ─────────────────────────────────────────────────────
write("office-addin/manifest-outlook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <Id>8c7fb87b-6d4e-5f9a-b2c3-4d5e6f7a8b9c</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>LMStudio Local AI</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="LMStudio AI Assistant"/>
  <Description DefaultValue="Chat with local LMStudio AI models from Outlook"/>
  <IconUrl DefaultValue="https://localhost:3000/assets/icon-80.png"/>
  <HighResolutionIconUrl DefaultValue="https://localhost:3000/assets/icon-80.png"/>
  <SupportUrl DefaultValue="https://localhost:3000"/>

  <AppDomains>
    <AppDomain>https://localhost:3000</AppDomain>
  </AppDomains>

  <Hosts>
    <Host Name="Mailbox"/>
  </Hosts>

  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.1"/>
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>
        <RequestedHeight>250</RequestedHeight>
      </DesktopSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>
      </DesktopSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>

  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  </Rule>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="Taskpane.Url"/>
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="Outlook.CommandsGroup">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="Outlook.TaskpaneButton">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
          <ExtensionPoint xsi:type="MessageComposeCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="Outlook.ComposeGroup">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="Outlook.ComposeButton">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="CommandsGroup.Label" DefaultValue="LMStudio AI"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="AI Assistant"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Open the LMStudio AI Assistant panel."/>
      </bt:LongStrings>
    </Resources>
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Requirements>
        <bt:Sets DefaultMinVersion="1.5">
          <bt:Set Name="Mailbox"/>
        </bt:Sets>
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost"/>
      </Hosts>
      <Resources>
        <bt:Images/>
        <bt:Urls/>
        <bt:ShortStrings/>
        <bt:LongStrings/>
      </Resources>
    </VersionOverrides>
  </VersionOverrides>
</OfficeApp>
`);

// ─── index.html ───────────────────────────────────────────────────────────────
write("office-addin/src/taskpane/index.html", `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>LMStudio AI Assistant</title>
  <!-- Office.js MUST load before the app bundle -->
  <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
</head>
<body>
  <div id="root">
    <p style="padding:1rem;font-family:sans-serif;color:#666;">Loading LMStudio AI...</p>
  </div>
</body>
</html>
`);

// ─── src/api/lmstudio.ts ──────────────────────────────────────────────────────
write("office-addin/src/api/lmstudio.ts", `export interface LMStudioModel {
  id: string;
  object: string;
  created: number;
  owned_by: string;
}

export interface ChatMessage {
  role: "system" | "user" | "assistant";
  content: string;
}

export interface ChatCompletionRequest {
  model: string;
  messages: ChatMessage[];
  temperature?: number;
  max_tokens?: number;
  stream?: boolean;
}

export class LMStudioClient {
  private baseUrl: string;

  constructor(baseUrl = "http://localhost:1234/v1") {
    this.baseUrl = baseUrl.replace(/\\/$/, "");
  }

  setBaseUrl(url: string): void {
    this.baseUrl = url.replace(/\\/$/, "");
  }

  /** Fetch available models from LMStudio. */
  async getModels(): Promise<LMStudioModel[]> {
    const res = await fetch(\`\${this.baseUrl}/models\`);
    if (!res.ok) throw new Error(\`Failed to fetch models: \${res.statusText}\`);
    const data = await res.json();
    return data.data ?? [];
  }

  /** Single-shot chat completion (non-streaming). */
  async chat(req: ChatCompletionRequest): Promise<string> {
    const res = await fetch(\`\${this.baseUrl}/chat/completions\`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ...req, stream: false }),
    });
    if (!res.ok) throw new Error(\`Chat request failed: \${res.statusText}\`);
    const data = await res.json();
    return data.choices?.[0]?.message?.content ?? "";
  }

  /** Streaming chat completion — yields text chunks as they arrive (SSE). */
  async *chatStream(req: ChatCompletionRequest): AsyncGenerator<string> {
    const res = await fetch(\`\${this.baseUrl}/chat/completions\`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ...req, stream: true }),
    });
    if (!res.ok) throw new Error(\`Stream request failed: \${res.statusText}\`);

    const reader = res.body!.getReader();
    const decoder = new TextDecoder();
    let buffer = "";

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split("\\n");
      buffer = lines.pop()!;
      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed || trimmed === "data: [DONE]") continue;
        if (trimmed.startsWith("data: ")) {
          try {
            const json = JSON.parse(trimmed.slice(6));
            const content: string | undefined = json.choices?.[0]?.delta?.content;
            if (content) yield content;
          } catch {
            // skip malformed SSE chunks
          }
        }
      }
    }
  }
}

/** Singleton client — components may create their own or use this default. */
export const defaultClient = new LMStudioClient();
`);

// ─── taskpane.tsx (entry point) ───────────────────────────────────────────────
write("office-addin/src/taskpane/taskpane.tsx", `import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import "./styles/taskpane.css";

/* Wait for Office.js to be ready before mounting React. */
Office.onReady(() => {
  const container = document.getElementById("root");
  if (!container) return;
  const root = createRoot(container);
  root.render(<App />);
});
`);

// ─── components/App.tsx ───────────────────────────────────────────────────────
write("office-addin/src/taskpane/components/App.tsx", `import * as React from "react";
import ChatPanel from "./ChatPanel";
import TextTransform from "./TextTransform";
import Summarizer from "./Summarizer";
import FormulaHelper from "./FormulaHelper";

type Tab = "chat" | "transform" | "summarize" | "formula";

function getHostType(): string {
  if (typeof Office === "undefined" || !Office.context?.host) return "unknown";
  return Office.HostType[Office.context.host] ?? "unknown";
}

export default function App() {
  const [tab, setTab] = React.useState<Tab>("chat");
  const [baseUrl, setBaseUrl] = React.useState("http://localhost:1234/v1");
  const [showSettings, setShowSettings] = React.useState(false);
  const hostType = React.useMemo(getHostType, []);
  const isExcel = hostType === "Excel";

  const tabs: Array<{ id: Tab; label: string }> = [
    { id: "chat", label: "💬 Chat" },
    { id: "transform", label: "✨ Transform" },
    { id: "summarize", label: "📄 Summarize" },
    ...(isExcel ? [{ id: "formula" as Tab, label: "⚡ Formula" }] : []),
  ];

  return (
    <div className="app-container">
      <header className="app-header">
        <h1>🤖 LMStudio AI</h1>
        <button
          className="settings-btn"
          onClick={() => setShowSettings(!showSettings)}
          title="Settings"
        >
          ⚙️
        </button>
      </header>

      {showSettings && (
        <div className="settings-panel">
          <label>LMStudio Server URL</label>
          <input
            type="text"
            value={baseUrl}
            onChange={(e) => setBaseUrl(e.target.value)}
            placeholder="http://localhost:1234/v1"
          />
          <small>Default: http://localhost:1234/v1</small>
        </div>
      )}

      <nav className="tab-nav">
        {tabs.map((t) => (
          <button
            key={t.id}
            className={\`tab-btn\${tab === t.id ? " active" : ""}\`}
            onClick={() => setTab(t.id)}
          >
            {t.label}
          </button>
        ))}
      </nav>

      <main className="tab-content">
        {tab === "chat" && <ChatPanel baseUrl={baseUrl} />}
        {tab === "transform" && <TextTransform baseUrl={baseUrl} />}
        {tab === "summarize" && <Summarizer baseUrl={baseUrl} />}
        {tab === "formula" && isExcel && <FormulaHelper baseUrl={baseUrl} />}
      </main>
    </div>
  );
}
`);

// ─── components/ChatPanel.tsx ─────────────────────────────────────────────────
write("office-addin/src/taskpane/components/ChatPanel.tsx", `import * as React from "react";
import { LMStudioClient, ChatMessage, LMStudioModel } from "../../api/lmstudio";

interface Props {
  baseUrl: string;
}

export default function ChatPanel({ baseUrl }: Props) {
  const [messages, setMessages] = React.useState<ChatMessage[]>([]);
  const [input, setInput] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [models, setModels] = React.useState<LMStudioModel[]>([]);
  const [selectedModel, setSelectedModel] = React.useState("");
  const [streamingText, setStreamingText] = React.useState("");
  const [error, setError] = React.useState("");
  const bottomRef = React.useRef<HTMLDivElement>(null);

  const client = React.useMemo(() => new LMStudioClient(baseUrl), [baseUrl]);

  /* Load models whenever baseUrl changes */
  React.useEffect(() => {
    setError("");
    client
      .getModels()
      .then((m) => {
        setModels(m);
        if (m.length > 0 && !selectedModel) setSelectedModel(m[0].id);
      })
      .catch(() => setError("⚠️ Cannot connect to LMStudio. Is the local server running?"));
  }, [client]);

  /* Auto-scroll to latest message */
  React.useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, streamingText]);

  const send = async () => {
    if (!input.trim() || loading) return;
    const userMsg: ChatMessage = { role: "user", content: input.trim() };
    const history = [...messages, userMsg];
    setMessages(history);
    setInput("");
    setLoading(true);
    setStreamingText("");
    setError("");

    try {
      let full = "";
      for await (const chunk of client.chatStream({
        model: selectedModel,
        messages: history,
        stream: true,
      })) {
        full += chunk;
        setStreamingText(full);
      }
      setMessages([...history, { role: "assistant", content: full }]);
    } catch (e: any) {
      setError(\`Error: \${e.message}\`);
    } finally {
      setLoading(false);
      setStreamingText("");
    }
  };

  return (
    <div className="panel chat-panel">
      {/* Model selector */}
      <div className="model-row">
        <select
          value={selectedModel}
          onChange={(e) => setSelectedModel(e.target.value)}
          disabled={models.length === 0}
        >
          {models.length === 0 && <option>— no models found —</option>}
          {models.map((m) => (
            <option key={m.id} value={m.id}>
              {m.id}
            </option>
          ))}
        </select>
        <button
          className="icon-btn"
          onClick={() => setMessages([])}
          title="Clear chat"
          disabled={loading}
        >
          🗑️
        </button>
      </div>

      {error && <p className="error-msg">{error}</p>}

      {/* Message history */}
      <div className="chat-log">
        {messages.length === 0 && (
          <p className="empty-hint">Send a message to start chatting with your local AI model.</p>
        )}
        {messages.map((msg, i) => (
          <div key={i} className={\`chat-bubble \${msg.role}\`}>
            <span className="bubble-label">{msg.role === "user" ? "You" : "AI"}</span>
            <p>{msg.content}</p>
          </div>
        ))}
        {loading && (
          <div className="chat-bubble assistant streaming">
            <span className="bubble-label">AI</span>
            <p>
              {streamingText || "Thinking…"}
              {streamingText && <span className="cursor">▋</span>}
            </p>
          </div>
        )}
        <div ref={bottomRef} />
      </div>

      {/* Input area */}
      <div className="chat-input-row">
        <textarea
          value={input}
          onChange={(e) => setInput(e.target.value)}
          placeholder="Type a message… (Enter to send, Shift+Enter for newline)"
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              send();
            }
          }}
          rows={2}
          disabled={loading}
        />
        <button
          className="send-btn"
          onClick={send}
          disabled={loading || !input.trim() || !selectedModel}
          title="Send"
        >
          ➤
        </button>
      </div>
    </div>
  );
}
`);

// ─── components/TextTransform.tsx ────────────────────────────────────────────
write("office-addin/src/taskpane/components/TextTransform.tsx", `import * as React from "react";
import { LMStudioClient, LMStudioModel } from "../../api/lmstudio";

interface Props {
  baseUrl: string;
}

type TransformType = "improve" | "rewrite" | "summarize" | "translate" | "formal" | "casual" | "bulletpoints" | "expand";

const TRANSFORMS: Record<TransformType, string> = {
  improve: "Improve the grammar, clarity, and style of the following text (preserve meaning):",
  rewrite: "Rewrite the following text more clearly and concisely:",
  summarize: "Summarize the following text in 2-3 sentences:",
  translate: "Translate the following text to English (detect source language automatically):",
  formal: "Rewrite the following text in a formal, professional tone:",
  casual: "Rewrite the following text in a casual, friendly, conversational tone:",
  bulletpoints: "Convert the following text into a concise bullet-point list:",
  expand: "Expand the following text with more detail and context while staying on topic:",
};

async function readSelectionText(): Promise<string> {
  /* Try Word */
  try {
    return await new Promise<string>((resolve, reject) =>
      Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load("text");
        await ctx.sync();
        sel.text ? resolve(sel.text) : reject(new Error("empty"));
      }).catch(reject)
    );
  } catch {}
  /* Try Excel */
  try {
    return await new Promise<string>((resolve, reject) =>
      Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load("text");
        await ctx.sync();
        const flat = range.text.flat().join(" ").trim();
        flat ? resolve(flat) : reject(new Error("empty"));
      }).catch(reject)
    );
  } catch {}
  throw new Error("Could not read selection — paste text manually.");
}

async function replaceSelectionText(text: string): Promise<void> {
  try {
    await Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.insertText(text, Word.InsertLocation.replace);
      await ctx.sync();
    });
  } catch {
    throw new Error("Insert not supported in this app — copy the result manually.");
  }
}

export default function TextTransform({ baseUrl }: Props) {
  const [models, setModels] = React.useState<LMStudioModel[]>([]);
  const [selectedModel, setSelectedModel] = React.useState("");
  const [transform, setTransform] = React.useState<TransformType>("improve");
  const [sourceText, setSourceText] = React.useState("");
  const [result, setResult] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [status, setStatus] = React.useState("");

  const client = React.useMemo(() => new LMStudioClient(baseUrl), [baseUrl]);

  React.useEffect(() => {
    client.getModels().then((m) => {
      setModels(m);
      if (m.length > 0) setSelectedModel(m[0].id);
    }).catch(() => {});
  }, [client]);

  const handleGetSelection = async () => {
    try {
      const text = await readSelectionText();
      setSourceText(text);
      setStatus("");
    } catch (e: any) {
      setStatus(e.message);
    }
  };

  const handleRun = async () => {
    if (!sourceText.trim() || loading) return;
    setLoading(true);
    setResult("");
    setStatus("");
    try {
      const prompt = \`\${TRANSFORMS[transform]}\\n\\n\${sourceText}\`;
      const response = await client.chat({
        model: selectedModel,
        messages: [{ role: "user", content: prompt }],
      });
      setResult(response);
    } catch (e: any) {
      setStatus(\`Error: \${e.message}\`);
    } finally {
      setLoading(false);
    }
  };

  const handleInsert = async () => {
    if (!result) return;
    try {
      await replaceSelectionText(result);
      setStatus("✅ Inserted into document");
    } catch (e: any) {
      setStatus(e.message);
    }
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(result).then(() => setStatus("📋 Copied to clipboard"));
  };

  return (
    <div className="panel transform-panel">
      <div className="model-row">
        <select value={selectedModel} onChange={(e) => setSelectedModel(e.target.value)}>
          {models.map((m) => <option key={m.id} value={m.id}>{m.id}</option>)}
        </select>
      </div>

      <div className="field-group">
        <label>Transformation</label>
        <select value={transform} onChange={(e) => setTransform(e.target.value as TransformType)}>
          {(Object.keys(TRANSFORMS) as TransformType[]).map((k) => (
            <option key={k} value={k}>{k.charAt(0).toUpperCase() + k.slice(1)}</option>
          ))}
        </select>
      </div>

      <div className="field-group">
        <div className="field-header">
          <label>Source Text</label>
          <button className="small-btn" onClick={handleGetSelection}>📋 Get Selection</button>
        </div>
        <textarea
          value={sourceText}
          onChange={(e) => setSourceText(e.target.value)}
          placeholder="Select text in your document or paste here…"
          rows={5}
        />
      </div>

      <button className="action-btn" onClick={handleRun} disabled={loading || !sourceText.trim()}>
        {loading ? "⏳ Processing…" : "✨ Transform"}
      </button>

      {result && (
        <div className="field-group result-group">
          <div className="field-header">
            <label>Result</label>
            <span>
              <button className="small-btn" onClick={handleInsert}>📝 Insert</button>
              <button className="small-btn" onClick={handleCopy}>📋 Copy</button>
            </span>
          </div>
          <div className="result-box">{result}</div>
        </div>
      )}

      {status && <p className="status-msg">{status}</p>}
    </div>
  );
}
`);

// ─── components/Summarizer.tsx ────────────────────────────────────────────────
write("office-addin/src/taskpane/components/Summarizer.tsx", `import * as React from "react";
import { LMStudioClient, LMStudioModel } from "../../api/lmstudio";

interface Props {
  baseUrl: string;
}

type SummaryLength = "short" | "medium" | "detailed";
const LENGTH_INSTRUCTIONS: Record<SummaryLength, string> = {
  short: "in 2-3 sentences",
  medium: "in one concise paragraph",
  detailed: "in detail using a bullet-point list of key points",
};

async function getDocumentText(): Promise<string> {
  /* Word */
  try {
    return await new Promise<string>((resolve, reject) =>
      Word.run(async (ctx) => {
        const body = ctx.document.body;
        body.load("text");
        await ctx.sync();
        resolve(body.text);
      }).catch(reject)
    );
  } catch {}

  /* Outlook */
  try {
    const item = Office.context?.mailbox?.item as any;
    if (item?.body) {
      return await new Promise<string>((resolve, reject) => {
        item.body.getAsync(Office.CoercionType.Text, (result: any) => {
          result.status === Office.AsyncResultStatus.Succeeded
            ? resolve(result.value)
            : reject(new Error(result.error?.message));
        });
      });
    }
  } catch {}

  /* PowerPoint — collect text from all shapes on all slides */
  try {
    return await new Promise<string>((resolve, reject) =>
      PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();
        const texts: string[] = [];
        for (const slide of slides.items) {
          slide.shapes.load("items");
          await ctx.sync();
          for (const shape of slide.shapes.items) {
            if ((shape as any).textFrame) {
              (shape as any).textFrame.load("text");
              await ctx.sync();
              texts.push((shape as any).textFrame.text);
            }
          }
        }
        resolve(texts.filter(Boolean).join("\\n\\n"));
      }).catch(reject)
    );
  } catch {}

  throw new Error("Unable to read document content from this Office application.");
}

const MAX_CHARS = 12000; // stay within typical context windows

export default function Summarizer({ baseUrl }: Props) {
  const [models, setModels] = React.useState<LMStudioModel[]>([]);
  const [selectedModel, setSelectedModel] = React.useState("");
  const [summary, setSummary] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [status, setStatus] = React.useState("");
  const [length, setLength] = React.useState<SummaryLength>("medium");

  const client = React.useMemo(() => new LMStudioClient(baseUrl), [baseUrl]);

  React.useEffect(() => {
    client.getModels().then((m) => {
      setModels(m);
      if (m.length > 0) setSelectedModel(m[0].id);
    }).catch(() => {});
  }, [client]);

  const handleSummarize = async () => {
    setLoading(true);
    setSummary("");
    setStatus("📄 Reading document…");
    try {
      const raw = await getDocumentText();
      if (!raw.trim()) {
        setStatus("The document appears to be empty.");
        return;
      }
      const truncated = raw.length > MAX_CHARS ? raw.slice(0, MAX_CHARS) + "…" : raw;
      setStatus("🤖 Summarizing…");
      const prompt = \`Summarize the following document \${LENGTH_INSTRUCTIONS[length]}:\\n\\n\${truncated}\`;
      const response = await client.chat({
        model: selectedModel,
        messages: [{ role: "user", content: prompt }],
      });
      setSummary(response);
      setStatus("");
    } catch (e: any) {
      setStatus(\`Error: \${e.message}\`);
    } finally {
      setLoading(false);
    }
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(summary).then(() => setStatus("📋 Copied to clipboard"));
  };

  return (
    <div className="panel summarizer-panel">
      <div className="model-row">
        <select value={selectedModel} onChange={(e) => setSelectedModel(e.target.value)}>
          {models.map((m) => <option key={m.id} value={m.id}>{m.id}</option>)}
        </select>
      </div>

      <div className="field-group">
        <label>Summary Length</label>
        <select value={length} onChange={(e) => setLength(e.target.value as SummaryLength)}>
          <option value="short">Short (2-3 sentences)</option>
          <option value="medium">Medium (1 paragraph)</option>
          <option value="detailed">Detailed (bullet points)</option>
        </select>
      </div>

      <button className="action-btn" onClick={handleSummarize} disabled={loading}>
        {loading ? "⏳ Working…" : "📄 Summarize Document"}
      </button>

      {status && <p className="status-msg">{status}</p>}

      {summary && (
        <div className="field-group result-group">
          <div className="field-header">
            <label>Summary</label>
            <button className="small-btn" onClick={handleCopy}>📋 Copy</button>
          </div>
          <div className="result-box">{summary}</div>
        </div>
      )}
    </div>
  );
}
`);

// ─── components/FormulaHelper.tsx ─────────────────────────────────────────────
write("office-addin/src/taskpane/components/FormulaHelper.tsx", `import * as React from "react";
import { LMStudioClient, LMStudioModel } from "../../api/lmstudio";

interface Props {
  baseUrl: string;
}

const SYSTEM_PROMPT =
  "You are an Excel formula expert. When given a natural language description, " +
  "respond with ONLY the Excel formula on the first line (starting with =), " +
  "followed by a blank line, then a concise plain-English explanation of how it works.";

export default function FormulaHelper({ baseUrl }: Props) {
  const [models, setModels] = React.useState<LMStudioModel[]>([]);
  const [selectedModel, setSelectedModel] = React.useState("");
  const [nlInput, setNlInput] = React.useState("");
  const [formula, setFormula] = React.useState("");
  const [explanation, setExplanation] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [status, setStatus] = React.useState("");

  const client = React.useMemo(() => new LMStudioClient(baseUrl), [baseUrl]);

  React.useEffect(() => {
    client.getModels().then((m) => {
      setModels(m);
      if (m.length > 0) setSelectedModel(m[0].id);
    }).catch(() => {});
  }, [client]);

  const handleGenerate = async () => {
    if (!nlInput.trim() || loading) return;
    setLoading(true);
    setFormula("");
    setExplanation("");
    setStatus("");
    try {
      const response = await client.chat({
        model: selectedModel,
        messages: [
          { role: "system", content: SYSTEM_PROMPT },
          { role: "user", content: nlInput },
        ],
      });
      const lines = response.trim().split("\\n");
      const formulaLine = lines.find((l) => l.trim().startsWith("=")) ?? lines[0];
      const rest = lines.filter((l) => l.trim() !== formulaLine.trim()).join("\\n").trim();
      setFormula(formulaLine.trim());
      setExplanation(rest);
    } catch (e: any) {
      setStatus(\`Error: \${e.message}\`);
    } finally {
      setLoading(false);
    }
  };

  const handleInsert = async () => {
    if (!formula) return;
    try {
      await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load("address");
        await ctx.sync();
        range.formulas = [[formula]];
        await ctx.sync();
        setStatus(\`✅ Formula inserted into \${range.address}\`);
      });
    } catch (e: any) {
      setStatus(\`Insert error: \${e.message}\`);
    }
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(formula).then(() => setStatus("📋 Copied to clipboard"));
  };

  const examples = [
    'Sum column B where column A equals "Sales"',
    "Average of C2:C100 ignoring blank cells",
    "Find the last non-empty row in column D",
    "VLOOKUP: find product price from a table on Sheet2",
    "Count unique values in the range A2:A500",
  ];

  return (
    <div className="panel formula-panel">
      <div className="model-row">
        <select value={selectedModel} onChange={(e) => setSelectedModel(e.target.value)}>
          {models.map((m) => <option key={m.id} value={m.id}>{m.id}</option>)}
        </select>
      </div>

      <div className="field-group">
        <label>Describe what you want the formula to do</label>
        <textarea
          value={nlInput}
          onChange={(e) => setNlInput(e.target.value)}
          placeholder='e.g. "Sum values in B2:B100 only if column A equals Sales"'
          rows={3}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleGenerate(); }
          }}
        />
        <div className="examples">
          {examples.map((ex) => (
            <button key={ex} className="example-chip" onClick={() => setNlInput(ex)}>
              {ex}
            </button>
          ))}
        </div>
      </div>

      <button className="action-btn" onClick={handleGenerate} disabled={loading || !nlInput.trim()}>
        {loading ? "⏳ Generating…" : "⚡ Generate Formula"}
      </button>

      {formula && (
        <div className="formula-result">
          <div className="formula-box">
            <code>{formula}</code>
            <span>
              <button className="small-btn" onClick={handleInsert} title="Insert into selected cell">📥 Insert</button>
              <button className="small-btn" onClick={handleCopy} title="Copy to clipboard">📋 Copy</button>
            </span>
          </div>
          {explanation && <p className="explanation">{explanation}</p>}
        </div>
      )}

      {status && <p className="status-msg">{status}</p>}
    </div>
  );
}
`);

// ─── styles/taskpane.css ──────────────────────────────────────────────────────
write("office-addin/src/taskpane/styles/taskpane.css", `/* ── Base ── */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

body {
  font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
  font-size: 13px;
  color: #1f2937;
  background: #f9fafb;
  height: 100vh;
  overflow: hidden;
}

/* ── App container ── */
.app-container {
  display: flex;
  flex-direction: column;
  height: 100vh;
  overflow: hidden;
}

.app-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 10px 12px 8px;
  background: #1e40af;
  color: white;
  flex-shrink: 0;
}

.app-header h1 {
  font-size: 15px;
  font-weight: 600;
  letter-spacing: 0.3px;
}

.settings-btn {
  background: transparent;
  border: none;
  color: white;
  font-size: 16px;
  cursor: pointer;
  padding: 2px 4px;
  border-radius: 4px;
}
.settings-btn:hover { background: rgba(255,255,255,0.2); }

/* ── Settings panel ── */
.settings-panel {
  padding: 8px 12px;
  background: #eff6ff;
  border-bottom: 1px solid #bfdbfe;
  flex-shrink: 0;
}
.settings-panel label { font-size: 11px; color: #1e40af; font-weight: 600; display: block; margin-bottom: 4px; }
.settings-panel input {
  width: 100%;
  padding: 5px 8px;
  border: 1px solid #93c5fd;
  border-radius: 4px;
  font-size: 12px;
}
.settings-panel small { font-size: 10px; color: #6b7280; display: block; margin-top: 3px; }

/* ── Tab navigation ── */
.tab-nav {
  display: flex;
  border-bottom: 1px solid #e5e7eb;
  background: white;
  flex-shrink: 0;
  overflow-x: auto;
}

.tab-btn {
  flex: 1;
  min-width: 60px;
  padding: 7px 4px;
  font-size: 11px;
  background: none;
  border: none;
  border-bottom: 2px solid transparent;
  cursor: pointer;
  color: #6b7280;
  transition: color 0.15s, border-color 0.15s;
  white-space: nowrap;
}
.tab-btn:hover { color: #1e40af; }
.tab-btn.active { color: #1e40af; border-bottom-color: #1e40af; font-weight: 600; }

/* ── Tab content ── */
.tab-content {
  flex: 1;
  overflow: hidden;
  display: flex;
  flex-direction: column;
}

.panel {
  flex: 1;
  overflow-y: auto;
  padding: 10px 12px;
  display: flex;
  flex-direction: column;
  gap: 10px;
}

/* ── Shared controls ── */
.model-row {
  display: flex;
  gap: 6px;
  align-items: center;
}

.model-row select, .field-group select {
  flex: 1;
  min-width: 0;
  padding: 5px 8px;
  border: 1px solid #d1d5db;
  border-radius: 4px;
  font-size: 12px;
  background: white;
}

.field-group { display: flex; flex-direction: column; gap: 4px; }
.field-group label { font-size: 11px; font-weight: 600; color: #374151; }

.field-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
}

textarea {
  width: 100%;
  padding: 7px 8px;
  border: 1px solid #d1d5db;
  border-radius: 4px;
  font-size: 12px;
  resize: vertical;
  font-family: inherit;
  line-height: 1.4;
}
textarea:focus { outline: none; border-color: #3b82f6; }

.action-btn {
  padding: 8px 12px;
  background: #1e40af;
  color: white;
  border: none;
  border-radius: 5px;
  font-size: 13px;
  font-weight: 500;
  cursor: pointer;
  transition: background 0.15s;
}
.action-btn:hover:not(:disabled) { background: #1d4ed8; }
.action-btn:disabled { background: #9ca3af; cursor: not-allowed; }

.small-btn {
  padding: 3px 8px;
  background: white;
  border: 1px solid #d1d5db;
  border-radius: 4px;
  font-size: 11px;
  cursor: pointer;
  margin-left: 4px;
}
.small-btn:hover { background: #f3f4f6; border-color: #9ca3af; }

.icon-btn {
  background: none;
  border: none;
  cursor: pointer;
  font-size: 15px;
  padding: 3px 5px;
  border-radius: 4px;
  flex-shrink: 0;
}
.icon-btn:hover { background: #f3f4f6; }
.icon-btn:disabled { opacity: 0.4; cursor: not-allowed; }

.result-group { margin-top: 4px; }

.result-box {
  padding: 8px 10px;
  background: white;
  border: 1px solid #e5e7eb;
  border-radius: 5px;
  font-size: 12px;
  line-height: 1.5;
  white-space: pre-wrap;
  word-break: break-word;
  max-height: 220px;
  overflow-y: auto;
}

.status-msg {
  font-size: 11px;
  padding: 5px 8px;
  border-radius: 4px;
  background: #fef3c7;
  border: 1px solid #fde68a;
  color: #92400e;
}

.error-msg {
  font-size: 11px;
  padding: 5px 8px;
  border-radius: 4px;
  background: #fef2f2;
  border: 1px solid #fecaca;
  color: #991b1b;
}

/* ── Chat panel ── */
.chat-panel {
  padding: 8px;
}

.chat-log {
  flex: 1;
  overflow-y: auto;
  display: flex;
  flex-direction: column;
  gap: 8px;
  padding: 4px 0 8px;
  min-height: 0;
  max-height: calc(100vh - 280px);
}

.empty-hint {
  color: #9ca3af;
  font-size: 12px;
  text-align: center;
  padding: 20px 10px;
}

.chat-bubble {
  padding: 8px 10px;
  border-radius: 8px;
  max-width: 92%;
  word-break: break-word;
  line-height: 1.45;
}

.chat-bubble.user {
  background: #1e40af;
  color: white;
  align-self: flex-end;
  border-bottom-right-radius: 2px;
}

.chat-bubble.assistant {
  background: white;
  border: 1px solid #e5e7eb;
  align-self: flex-start;
  border-bottom-left-radius: 2px;
}

.bubble-label {
  display: block;
  font-size: 10px;
  font-weight: 700;
  opacity: 0.7;
  margin-bottom: 3px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
}

.chat-bubble.user .bubble-label { color: #bfdbfe; }
.chat-bubble.assistant .bubble-label { color: #6b7280; }

.chat-bubble p {
  font-size: 13px;
  margin: 0;
  white-space: pre-wrap;
}

.cursor {
  animation: blink 0.8s step-end infinite;
  font-weight: 300;
}
@keyframes blink { 50% { opacity: 0; } }

.chat-input-row {
  display: flex;
  gap: 6px;
  align-items: flex-end;
  flex-shrink: 0;
}

.chat-input-row textarea {
  flex: 1;
  min-width: 0;
  resize: none;
  padding: 7px 10px;
}

.send-btn {
  width: 36px;
  height: 36px;
  background: #1e40af;
  color: white;
  border: none;
  border-radius: 6px;
  font-size: 15px;
  cursor: pointer;
  flex-shrink: 0;
  transition: background 0.15s;
}
.send-btn:hover:not(:disabled) { background: #1d4ed8; }
.send-btn:disabled { background: #9ca3af; cursor: not-allowed; }

/* ── Formula helper ── */
.examples {
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
  margin-top: 5px;
}

.example-chip {
  padding: 3px 8px;
  border: 1px dashed #93c5fd;
  border-radius: 12px;
  background: #eff6ff;
  color: #1e40af;
  font-size: 10px;
  cursor: pointer;
  transition: background 0.1s;
}
.example-chip:hover { background: #dbeafe; }

.formula-result {
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.formula-box {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 8px 10px;
  background: #f8faff;
  border: 1px solid #bfdbfe;
  border-radius: 6px;
  gap: 8px;
}

.formula-box code {
  font-family: "Cascadia Code", "Consolas", monospace;
  font-size: 13px;
  color: #1e40af;
  word-break: break-all;
  flex: 1;
}

.explanation {
  font-size: 12px;
  color: #4b5563;
  line-height: 1.5;
  padding: 6px 8px;
  background: #f9fafb;
  border-radius: 4px;
  border-left: 3px solid #93c5fd;
}
`);

// ─────────────────────────────────────────────────────────────────────────────
// VSTO ADD-IN (C# / .NET Framework)
// ─────────────────────────────────────────────────────────────────────────────

const vstoGuid = "A1B2C3D4-E5F6-7890-ABCD-EF1234567890";
const vstoAssemblyGuid = "B2C3D4E5-F6A7-8901-BCDE-F01234567891";

write("vsto/LMStudioAddin.sln", `
Microsoft Visual Studio Solution File, Format Version 12.00
# Visual Studio Version 17
VisualStudioVersion = 17.0.31903.59
MinimumVisualStudioVersion = 10.0.40219.1
Project("{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}") = "LMStudioAddin", "LMStudioAddin\\\\LMStudioAddin.csproj", "{${vstoGuid}}"
EndProject
Global
  GlobalSection(SolutionConfigurationPlatforms) = preSolution
    Debug|AnyCPU = Debug|AnyCPU
    Release|AnyCPU = Release|AnyCPU
  EndGlobalSection
  GlobalSection(ProjectConfigurationPlatforms) = postSolution
    {${vstoGuid}}.Debug|AnyCPU.ActiveCfg = Debug|AnyCPU
    {${vstoGuid}}.Debug|AnyCPU.Build.0 = Debug|AnyCPU
    {${vstoGuid}}.Release|AnyCPU.ActiveCfg = Release|AnyCPU
    {${vstoGuid}}.Release|AnyCPU.Build.0 = Release|AnyCPU
  EndGlobalSection
  GlobalSection(SolutionProperties) = preSolution
    HideSolutionNode = FALSE
  EndGlobalSection
EndGlobal
`);

write("vsto/LMStudioAddin/LMStudioAddin.csproj", `<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <Import Project="$(MSBuildExtensionsPath)\\$(MSBuildToolsVersion)\\Microsoft.Common.props"
          Condition="Exists('$(MSBuildExtensionsPath)\\$(MSBuildToolsVersion)\\Microsoft.Common.props')" />
  <PropertyGroup>
    <Configuration Condition=" '$(Configuration)' == '' ">Debug</Configuration>
    <Platform Condition=" '$(Platform)' == '' ">AnyCPU</Platform>
    <ProjectGuid>{${vstoGuid}}</ProjectGuid>
    <OutputType>Library</OutputType>
    <RootNamespace>LMStudioAddin</RootNamespace>
    <AssemblyName>LMStudioAddin</AssemblyName>
    <TargetFrameworkVersion>v4.7.2</TargetFrameworkVersion>
    <!-- VSTO project type GUIDs: VSTO Add-In + C# -->
    <ProjectTypeGuids>{BAA0C2D2-18E2-41B9-852F-F413020CAA33};{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}</ProjectTypeGuids>
    <OfficeApplication>Word</OfficeApplication>
    <UseApplicationTrust>false</UseApplicationTrust>
    <Deterministic>true</Deterministic>
  </PropertyGroup>
  <PropertyGroup Condition=" '$(Configuration)|$(Platform)' == 'Debug|AnyCPU' ">
    <DebugSymbols>true</DebugSymbols>
    <DebugType>full</DebugType>
    <Optimize>false</Optimize>
    <OutputPath>bin\\Debug\\</OutputPath>
    <DefineConstants>DEBUG;TRACE</DefineConstants>
  </PropertyGroup>
  <PropertyGroup Condition=" '$(Configuration)|$(Platform)' == 'Release|AnyCPU' ">
    <DebugType>pdbonly</DebugType>
    <Optimize>true</Optimize>
    <OutputPath>bin\\Release\\</OutputPath>
    <DefineConstants>TRACE</DefineConstants>
  </PropertyGroup>
  <ItemGroup>
    <Reference Include="Microsoft.Office.Tools.Common, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a" />
    <Reference Include="Microsoft.Office.Tools.Word, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a" />
    <Reference Include="Microsoft.Office.Tools, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a" />
    <Reference Include="Microsoft.VisualStudio.Tools.Applications.Runtime, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a" />
    <Reference Include="Microsoft.Office.Interop.Word, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
    <Reference Include="System" />
    <Reference Include="System.Core" />
    <Reference Include="System.Net.Http" />
    <Reference Include="System.Windows.Forms" />
    <Reference Include="System.Drawing" />
  </ItemGroup>
  <ItemGroup>
    <!-- NuGet: Microsoft.Web.WebView2 — add via NuGet Package Manager in VS -->
    <Compile Include="ThisAddIn.cs" />
    <Compile Include="ThisAddIn.Ribbon.cs" />
    <Compile Include="Ribbon.cs" />
    <Compile Include="LMStudioClient.cs" />
    <Compile Include="TaskPaneControl.cs" />
    <Compile Include="Properties\\AssemblyInfo.cs" />
  </ItemGroup>
  <Import Project="$(MSBuildToolsPath)\\Microsoft.CSharp.targets" />
</Project>
`);

write("vsto/LMStudioAddin/Properties/AssemblyInfo.cs", `using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("LMStudioAddin")]
[assembly: AssemblyDescription("Office VSTO Add-in for LMStudio local AI models")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("LMStudio AI Assistant")]
[assembly: AssemblyCopyright("")]
[assembly: ComVisible(false)]
[assembly: Guid("${vstoAssemblyGuid}")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]
`);

write("vsto/LMStudioAddin/LMStudioClient.cs", `using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace LMStudioAddin
{
    public class LMStudioModel
    {
        public string Id { get; set; }
        public string Object { get; set; }
    }

    public class ChatMessage
    {
        public string Role { get; set; }
        public string Content { get; set; }
    }

    public class ChatRequest
    {
        public string Model { get; set; }
        public List<ChatMessage> Messages { get; set; }
        public double Temperature { get; set; } = 0.7;
        public bool Stream { get; set; } = false;
    }

    /// <summary>
    /// Thin wrapper around LMStudio's OpenAI-compatible local REST API.
    /// Base URL defaults to http://localhost:1234/v1.
    /// </summary>
    public class LMStudioClient : IDisposable
    {
        private readonly HttpClient _http;
        private string _baseUrl;

        public string BaseUrl
        {
            get => _baseUrl;
            set => _baseUrl = value.TrimEnd('/');
        }

        public LMStudioClient(string baseUrl = "http://localhost:1234/v1")
        {
            _baseUrl = baseUrl.TrimEnd('/');
            _http = new HttpClient { Timeout = TimeSpan.FromSeconds(120) };
            _http.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));
        }

        /// <summary>Returns all models currently loaded in LMStudio.</summary>
        public async Task<List<LMStudioModel>> GetModelsAsync()
        {
            var response = await _http.GetAsync($"{_baseUrl}/models");
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            var models = new List<LMStudioModel>();
            if (doc.RootElement.TryGetProperty("data", out var data))
            {
                foreach (var item in data.EnumerateArray())
                {
                    models.Add(new LMStudioModel
                    {
                        Id = item.GetProperty("id").GetString(),
                        Object = item.TryGetProperty("object", out var o) ? o.GetString() : ""
                    });
                }
            }
            return models;
        }

        /// <summary>Single-shot chat completion (non-streaming).</summary>
        public async Task<string> ChatAsync(ChatRequest request,
            CancellationToken cancellationToken = default)
        {
            request.Stream = false;
            var body = SerializeRequest(request);
            var content = new StringContent(body, Encoding.UTF8, "application/json");
            var response = await _http.PostAsync($"{_baseUrl}/chat/completions", content, cancellationToken);
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            return doc.RootElement
                       .GetProperty("choices")[0]
                       .GetProperty("message")
                       .GetProperty("content")
                       .GetString() ?? string.Empty;
        }

        /// <summary>
        /// Streaming chat completion — invokes <paramref name="onChunk"/> for each
        /// text token as it arrives, then returns the full assembled response.
        /// </summary>
        public async Task<string> ChatStreamAsync(
            ChatRequest request,
            Action<string> onChunk,
            CancellationToken cancellationToken = default)
        {
            request.Stream = true;
            var body = SerializeRequest(request);
            var requestMessage = new HttpRequestMessage(HttpMethod.Post,
                $"{_baseUrl}/chat/completions")
            {
                Content = new StringContent(body, Encoding.UTF8, "application/json")
            };
            requestMessage.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/event-stream"));

            using var response = await _http.SendAsync(requestMessage,
                HttpCompletionOption.ResponseHeadersRead, cancellationToken);
            response.EnsureSuccessStatusCode();

            using var stream = await response.Content.ReadAsStreamAsync();
            using var reader = new System.IO.StreamReader(stream);

            var fullText = new StringBuilder();
            string line;
            while ((line = await reader.ReadLineAsync()) != null)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (string.IsNullOrWhiteSpace(line) || line == "data: [DONE]") continue;
                if (line.StartsWith("data: "))
                {
                    try
                    {
                        using var doc = JsonDocument.Parse(line.Substring(6));
                        if (doc.RootElement
                               .GetProperty("choices")[0]
                               .GetProperty("delta")
                               .TryGetProperty("content", out var c))
                        {
                            var chunk = c.GetString();
                            if (!string.IsNullOrEmpty(chunk))
                            {
                                fullText.Append(chunk);
                                onChunk?.Invoke(chunk);
                            }
                        }
                    }
                    catch { /* skip malformed SSE chunks */ }
                }
            }
            return fullText.ToString();
        }

        private static string SerializeRequest(ChatRequest req)
        {
            var messages = new System.Text.StringBuilder();
            foreach (var m in req.Messages)
            {
                if (messages.Length > 0) messages.Append(",");
                messages.Append($"{{\"role\":\"{m.Role}\",\"content\":{JsonEncodedText.Encode(m.Content)}}}");
            }
            return $"{{\"model\":\"{req.Model}\",\"messages\":[{messages}]," +
                   $"\"temperature\":{req.Temperature},\"stream\":{req.Stream.ToString().ToLower()}}}";
        }

        public void Dispose() => _http?.Dispose();
    }
}
`);

write("vsto/LMStudioAddin/ThisAddIn.cs", `using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Word;
using Word = Microsoft.Office.Interop.Word;

namespace LMStudioAddin
{
    public partial class ThisAddIn
    {
        private CustomTaskPane _taskPane;
        private TaskPaneControl _taskPaneControl;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Task pane is created lazily when the ribbon button is clicked.
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _taskPane = null;
            _taskPaneControl?.Dispose();
        }

        /// <summary>
        /// Called by the ribbon button to toggle the AI task pane.
        /// </summary>
        public void ToggleTaskPane()
        {
            if (_taskPane == null)
            {
                _taskPaneControl = new TaskPaneControl();
                _taskPane = this.CustomTaskPanes.Add(_taskPaneControl, "LMStudio AI Assistant");
                _taskPane.Width = 340;
                _taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Visible = true;
            }
            else
            {
                _taskPane.Visible = !_taskPane.Visible;
            }
        }

        /// <summary>
        /// Returns the currently selected text in the active Word document.
        /// </summary>
        public string GetSelectedText()
        {
            try
            {
                return Application.Selection?.Text ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Replaces the current selection with <paramref name="text"/>.
        /// </summary>
        public void InsertText(string text)
        {
            try
            {
                if (Application.Selection != null)
                    Application.Selection.TypeText(text);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not insert text: {ex.Message}", "LMStudio Add-in",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
`);

write("vsto/LMStudioAddin/TaskPaneControl.cs", `using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LMStudioAddin
{
    /// <summary>
    /// WinForms UserControl that acts as the VSTO Task Pane.
    /// Provides chat, text-transform, and document-summarize panels.
    /// Optionally hosts WebView2 to reuse the React UI built from office-addin/.
    /// </summary>
    public class TaskPaneControl : UserControl
    {
        // ── Controls ──────────────────────────────────────────────────────────
        private TabControl _tabs;
        private TabPage _chatTab, _transformTab, _summarizeTab;

        // Chat tab
        private ComboBox _chatModelCombo;
        private RichTextBox _chatLog;
        private TextBox _chatInput;
        private Button _sendBtn, _clearBtn;

        // Transform tab
        private ComboBox _transformModelCombo, _transformTypeCombo;
        private RichTextBox _sourceText, _resultText;
        private Button _getSelectionBtn, _transformBtn, _insertBtn;

        // Summarize tab
        private ComboBox _summarizeModelCombo, _summaryLengthCombo;
        private RichTextBox _summaryText;
        private Button _summarizeBtn;

        // Settings
        private Panel _settingsBar;
        private TextBox _urlBox;

        // State
        private LMStudioClient _client;
        private CancellationTokenSource _cts;
        private readonly List<ChatMessage> _history = new List<ChatMessage>();

        public TaskPaneControl()
        {
            _client = new LMStudioClient();
            InitializeComponent();
            LoadModels();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.Dock = DockStyle.Fill;
            this.BackColor = Color.FromArgb(249, 250, 251);
            this.Font = new Font("Segoe UI", 9f);

            // ── Settings bar ──────────────────────────────────────────────
            _settingsBar = new Panel { Dock = DockStyle.Top, Height = 30, BackColor = Color.FromArgb(239, 246, 255) };
            var urlLabel = new Label { Text = "URL:", Left = 4, Top = 7, AutoSize = true };
            _urlBox = new TextBox
            {
                Text = "http://localhost:1234/v1",
                Left = 30, Top = 4, Width = 200, Height = 22
            };
            _urlBox.Leave += (s, e) =>
            {
                _client.BaseUrl = _urlBox.Text.Trim();
                LoadModels();
            };
            var applyBtn = new Button { Text = "✓", Left = 234, Top = 3, Width = 28, Height = 24, FlatStyle = FlatStyle.Flat };
            applyBtn.Click += (s, e) => { _client.BaseUrl = _urlBox.Text.Trim(); LoadModels(); };
            _settingsBar.Controls.AddRange(new Control[] { urlLabel, _urlBox, applyBtn });

            // ── Tabs ──────────────────────────────────────────────────────
            _tabs = new TabControl { Dock = DockStyle.Fill };

            _chatTab = new TabPage("💬 Chat");
            _transformTab = new TabPage("✨ Transform");
            _summarizeTab = new TabPage("📄 Summarize");

            _tabs.TabPages.AddRange(new[] { _chatTab, _transformTab, _summarizeTab });

            BuildChatTab();
            BuildTransformTab();
            BuildSummarizeTab();

            this.Controls.Add(_tabs);
            this.Controls.Add(_settingsBar);
            this.ResumeLayout();
        }

        // ── Chat tab ──────────────────────────────────────────────────────────
        private void BuildChatTab()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, RowCount = 4, ColumnCount = 1,
                Padding = new Padding(4)
            };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));

            // Model + clear row
            var topRow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight };
            _chatModelCombo = new ComboBox { Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };
            _clearBtn = new Button { Text = "🗑", Width = 28, Height = 23, FlatStyle = FlatStyle.Flat };
            _clearBtn.Click += (s, e) => { _history.Clear(); _chatLog.Clear(); };
            topRow.Controls.AddRange(new Control[] { _chatModelCombo, _clearBtn });

            _chatLog = new RichTextBox
            {
                Dock = DockStyle.Fill, ReadOnly = true,
                BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9f), ScrollBars = RichTextBoxScrollBars.Vertical
            };

            _chatInput = new TextBox
            {
                Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical,
                Font = new Font("Segoe UI", 9f)
            };
            _chatInput.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter && !e.Shift) { e.SuppressKeyPress = true; SendChat(); }
            };

            _sendBtn = new Button
            {
                Text = "Send ➤", Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(30, 64, 175), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            _sendBtn.Click += (s, e) => SendChat();

            panel.Controls.Add(topRow, 0, 0);
            panel.Controls.Add(_chatLog, 0, 1);
            panel.Controls.Add(_chatInput, 0, 2);
            panel.Controls.Add(_sendBtn, 0, 3);
            _chatTab.Controls.Add(panel);
        }

        private void SendChat()
        {
            var text = _chatInput.Text.Trim();
            if (string.IsNullOrEmpty(text) || _chatModelCombo.SelectedItem == null) return;
            _history.Add(new ChatMessage { Role = "user", Content = text });
            AppendChatLine("You", text, Color.FromArgb(30, 64, 175));
            _chatInput.Clear();
            _sendBtn.Enabled = false;

            _cts?.Cancel();
            _cts = new CancellationTokenSource();
            var model = _chatModelCombo.SelectedItem.ToString();
            var history = new List<ChatMessage>(_history);
            var token = _cts.Token;

            Task.Run(async () =>
            {
                try
                {
                    var response = await _client.ChatStreamAsync(
                        new ChatRequest { Model = model, Messages = history, Stream = true },
                        chunk => BeginInvoke((Action)(() => AppendStreamChunk(chunk))),
                        token);
                    _history.Add(new ChatMessage { Role = "assistant", Content = response });
                    BeginInvoke((Action)(() => FinalizeStream()));
                }
                catch (OperationCanceledException) { }
                catch (Exception ex)
                {
                    BeginInvoke((Action)(() => AppendChatLine("Error", ex.Message, Color.Red)));
                }
                finally
                {
                    BeginInvoke((Action)(() => _sendBtn.Enabled = true));
                }
            }, token);
        }

        private bool _streamingActive = false;

        private void AppendChatLine(string sender, string message, Color color)
        {
            _streamingActive = false;
            _chatLog.SelectionStart = _chatLog.TextLength;
            _chatLog.SelectionColor = color;
            _chatLog.AppendText($"[{sender}] ");
            _chatLog.SelectionColor = _chatLog.ForeColor;
            _chatLog.AppendText(message + "\\n\\n");
            _chatLog.ScrollToCaret();
        }

        private void AppendStreamChunk(string chunk)
        {
            if (!_streamingActive)
            {
                _chatLog.SelectionStart = _chatLog.TextLength;
                _chatLog.SelectionColor = Color.FromArgb(30, 64, 175);
                _chatLog.AppendText("[AI] ");
                _chatLog.SelectionColor = _chatLog.ForeColor;
                _streamingActive = true;
            }
            _chatLog.AppendText(chunk);
            _chatLog.ScrollToCaret();
        }

        private void FinalizeStream()
        {
            _streamingActive = false;
            _chatLog.AppendText("\\n\\n");
            _chatLog.ScrollToCaret();
        }

        // ── Transform tab ─────────────────────────────────────────────────────
        private void BuildTransformTab()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, RowCount = 7, ColumnCount = 1,
                Padding = new Padding(4)
            };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 45));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 45));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));

            _transformModelCombo = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };

            _transformTypeCombo = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _transformTypeCombo.Items.AddRange(new object[] {
                "Improve grammar & style", "Rewrite clearly", "Summarize",
                "Translate to English", "Make formal", "Make casual",
                "Convert to bullet points", "Expand with detail"
            });
            _transformTypeCombo.SelectedIndex = 0;

            var srcLabel = new Label { Text = "Source text:", AutoSize = true };
            _sourceText = new RichTextBox { Dock = DockStyle.Fill, ScrollBars = RichTextBoxScrollBars.Vertical };

            _getSelectionBtn = new Button { Text = "📋 Get Selection", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat };
            _getSelectionBtn.Click += (s, e) =>
            {
                var selected = Globals.ThisAddIn.GetSelectedText();
                if (!string.IsNullOrEmpty(selected)) _sourceText.Text = selected;
            };

            _transformBtn = new Button
            {
                Text = "✨ Transform", Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(30, 64, 175), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            _transformBtn.Click += (s, e) => RunTransform();

            _resultText = new RichTextBox
            {
                Dock = DockStyle.Fill, ReadOnly = true,
                ScrollBars = RichTextBoxScrollBars.Vertical, BackColor = Color.FromArgb(248, 250, 252)
            };

            _insertBtn = new Button { Text = "📝 Insert into Document", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat };
            _insertBtn.Click += (s, e) =>
            {
                if (!string.IsNullOrEmpty(_resultText.Text))
                    Globals.ThisAddIn.InsertText(_resultText.Text);
            };

            panel.Controls.Add(_transformModelCombo, 0, 0);
            panel.Controls.Add(_transformTypeCombo, 0, 1);
            panel.Controls.Add(srcLabel, 0, 2);
            panel.Controls.Add(_sourceText, 0, 3);
            panel.Controls.Add(_getSelectionBtn, 0, 4);
            panel.Controls.Add(_resultText, 0, 5);
            panel.Controls.Add(_insertBtn, 0, 6);
            _transformTab.Controls.Add(panel);
        }

        private static readonly string[] TransformPrompts = {
            "Improve the grammar, clarity, and style of the following text (preserve meaning):",
            "Rewrite the following text more clearly and concisely:",
            "Summarize the following text in 2-3 sentences:",
            "Translate the following text to English (detect source language):",
            "Rewrite the following text in a formal, professional tone:",
            "Rewrite the following text in a casual, friendly tone:",
            "Convert the following text into a concise bullet-point list:",
            "Expand the following text with more detail and context:",
        };

        private void RunTransform()
        {
            var source = _sourceText.Text.Trim();
            if (string.IsNullOrEmpty(source) || _transformModelCombo.SelectedItem == null) return;
            _transformBtn.Enabled = false;
            _resultText.Clear();
            var model = _transformModelCombo.SelectedItem.ToString();
            var prompt = TransformPrompts[_transformTypeCombo.SelectedIndex];

            Task.Run(async () =>
            {
                try
                {
                    var result = await _client.ChatAsync(new ChatRequest
                    {
                        Model = model,
                        Messages = new List<ChatMessage>
                        {
                            new ChatMessage { Role = "user", Content = $"{prompt}\\n\\n{source}" }
                        }
                    });
                    BeginInvoke((Action)(() => _resultText.Text = result));
                }
                catch (Exception ex)
                {
                    BeginInvoke((Action)(() => _resultText.Text = $"Error: {ex.Message}"));
                }
                finally
                {
                    BeginInvoke((Action)(() => _transformBtn.Enabled = true));
                }
            });
        }

        // ── Summarize tab ─────────────────────────────────────────────────────
        private void BuildSummarizeTab()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, RowCount = 4, ColumnCount = 1,
                Padding = new Padding(4)
            };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            _summarizeModelCombo = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };

            _summaryLengthCombo = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _summaryLengthCombo.Items.AddRange(new object[] {
                "Short (2-3 sentences)", "Medium (1 paragraph)", "Detailed (bullet points)"
            });
            _summaryLengthCombo.SelectedIndex = 1;

            _summarizeBtn = new Button
            {
                Text = "📄 Summarize Document", Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(30, 64, 175), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            _summarizeBtn.Click += (s, e) => RunSummarize();

            _summaryText = new RichTextBox
            {
                Dock = DockStyle.Fill, ReadOnly = true,
                BackColor = Color.FromArgb(248, 250, 252), ScrollBars = RichTextBoxScrollBars.Vertical
            };

            panel.Controls.Add(_summarizeModelCombo, 0, 0);
            panel.Controls.Add(_summaryLengthCombo, 0, 1);
            panel.Controls.Add(_summarizeBtn, 0, 2);
            panel.Controls.Add(_summaryText, 0, 3);
            _summarizeTab.Controls.Add(panel);
        }

        private static readonly string[] LengthInstructions = {
            "in 2-3 sentences",
            "in one concise paragraph",
            "in detail using a bullet-point list of key points"
        };

        private void RunSummarize()
        {
            if (_summarizeModelCombo.SelectedItem == null) return;
            _summarizeBtn.Enabled = false;
            _summaryText.Text = "Reading document…";
            var model = _summarizeModelCombo.SelectedItem.ToString();
            var instruction = LengthInstructions[_summaryLengthCombo.SelectedIndex];
            var docText = GetWordDocumentText();

            Task.Run(async () =>
            {
                try
                {
                    var prompt = $"Summarize the following document {instruction}:\\n\\n" +
                                 (docText.Length > 12000 ? docText.Substring(0, 12000) + "…" : docText);
                    var result = await _client.ChatAsync(new ChatRequest
                    {
                        Model = model,
                        Messages = new List<ChatMessage>
                        {
                            new ChatMessage { Role = "user", Content = prompt }
                        }
                    });
                    BeginInvoke((Action)(() => _summaryText.Text = result));
                }
                catch (Exception ex)
                {
                    BeginInvoke((Action)(() => _summaryText.Text = $"Error: {ex.Message}"));
                }
                finally
                {
                    BeginInvoke((Action)(() => _summarizeBtn.Enabled = true));
                }
            });
        }

        private string GetWordDocumentText()
        {
            try
            {
                return Globals.ThisAddIn.Application.ActiveDocument?.Content?.Text ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        // ── Model loading ─────────────────────────────────────────────────────
        private void LoadModels()
        {
            Task.Run(async () =>
            {
                try
                {
                    var models = await _client.GetModelsAsync();
                    BeginInvoke((Action)(() =>
                    {
                        foreach (ComboBox combo in new[] { _chatModelCombo, _transformModelCombo, _summarizeModelCombo })
                        {
                            combo.Items.Clear();
                            foreach (var m in models) combo.Items.Add(m.Id);
                            if (combo.Items.Count > 0) combo.SelectedIndex = 0;
                        }
                    }));
                }
                catch
                {
                    BeginInvoke((Action)(() =>
                    {
                        foreach (ComboBox combo in new[] { _chatModelCombo, _transformModelCombo, _summarizeModelCombo })
                        {
                            combo.Items.Clear();
                            combo.Items.Add("— check LMStudio server —");
                            combo.SelectedIndex = 0;
                        }
                    }));
                }
            });
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _cts?.Cancel();
                _cts?.Dispose();
                _client?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
`);

// ─── VSTO Ribbon (triggers task pane from Word Home tab) ─────────────────────
write("vsto/LMStudioAddin/Ribbon.cs", `using System;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace LMStudioAddin
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public Ribbon() { }

        public string GetCustomUI(string ribbonID) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'
          onLoad='Ribbon_Load'>
  <ribbon>
    <tabs>
      <tab idMso='TabHome'>
        <group id='LMStudioGroup' label='LMStudio AI'>
          <button id='ToggleTaskPane'
                  label='Open AI Assistant'
                  imageMso='HappyFace'
                  size='large'
                  onAction='ToggleTaskPane_Click'
                  screentip='Open the LMStudio AI task pane'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void ToggleTaskPane_Click(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.ToggleTaskPane();
        }
    }
}
`);

// Update ThisAddIn.cs to wire up the Ribbon
// (Patch: add CreateRibbonExtensibilityObject override)
write("vsto/LMStudioAddin/ThisAddIn.Ribbon.cs", `using System;
using Microsoft.Office.Core;

namespace LMStudioAddin
{
    public partial class ThisAddIn
    {
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }
    }
}
`);

// ─────────────────────────────────────────────────────────────────────────────
// SETUP INSTRUCTIONS FILE
// ─────────────────────────────────────────────────────────────────────────────
write("SETUP.txt", `
╔══════════════════════════════════════════════════════════════════════╗
║          LMStudio Office Plugin — Setup Instructions                ║
╚══════════════════════════════════════════════════════════════════════╝

Prerequisites
─────────────
• LMStudio installed and running with "Local Server" enabled (port 1234)
• Node.js 18+ and npm (for the Office Add-in)
• Visual Studio 2022 with "Office/SharePoint Development" workload (for VSTO)

═══════════════════════════════════════════════════════════════════════
OPTION A: Office Add-in  (works on Desktop + Office Online)
═══════════════════════════════════════════════════════════════════════

1. Install dependencies
   cd office-addin
   npm install

2. Install dev HTTPS certificates (first time only)
   npm run certs

3. Start the dev server
   npm run dev-server
   → Opens at https://localhost:3000

4. Sideload the add-in into Office
   • Word/Excel/PowerPoint:
       File → Options → Trust Center → Trust Center Settings
       → Trusted Add-in Catalogs → Add "https://localhost:3000"
     OR use the npm script:
       npm start          (auto-sideloads using office-addin-debugging)

5. Using the add-in
   • A ribbon button "Open AI Assistant" appears in the Home tab
   • Click it to open the task pane
   • Use ⚙️ to change the LMStudio server URL if needed

For production deployment:
   npm run build          → outputs to office-addin/dist/
   Host dist/ on any HTTPS server and update manifest.xml URLs.

Outlook (separate manifest):
   Sideload office-addin/manifest-outlook.xml through
   Outlook → File → Manage Add-ins → Add from file.

═══════════════════════════════════════════════════════════════════════
OPTION B: VSTO Add-in  (Windows Desktop only, deeper integration)
═══════════════════════════════════════════════════════════════════════

1. Open vsto/LMStudioAddin.sln in Visual Studio 2022

2. Install NuGet package (optional, for WebView2):
   Tools → NuGet Package Manager → Console:
   Install-Package Microsoft.Web.WebView2

3. Build and Run (F5)
   • VS will deploy and start Word with the add-in loaded
   • A "LMStudio AI" group appears in the Word Home ribbon
   • Click "Open AI Assistant" to open the task pane

4. The task pane has three tabs:
   💬 Chat      — converse with the selected LMStudio model
   ✨ Transform — rewrite/summarize/translate selected text
   📄 Summarize — summarize the entire active document

═══════════════════════════════════════════════════════════════════════
LMStudio Server Configuration
═══════════════════════════════════════════════════════════════════════
Default URL: http://localhost:1234/v1

To use a different port or host:
  • Office Add-in: click ⚙️ in the task pane header
  • VSTO: change the URL in the settings bar at the top of the pane

Ensure LMStudio has at least one model loaded before opening the add-in.
`);

console.log("\n✅ Project scaffold complete!\n");
console.log("Files created:");
console.log("  office-addin/   — React + TypeScript Office Add-in");
console.log("  vsto/           — C# VSTO Add-in for Word");
console.log("  SETUP.txt       — Full setup instructions\n");
console.log("Next steps:");
console.log("  1. cd office-addin && npm install");
console.log("  2. npm run certs");
console.log("  3. npm run dev-server");
console.log("  4. Sideload manifest.xml into Office\n");
