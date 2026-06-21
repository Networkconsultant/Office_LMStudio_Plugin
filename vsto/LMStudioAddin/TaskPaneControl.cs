using System;
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
            _chatLog.AppendText(message + "\n\n");
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
            _chatLog.AppendText("\n\n");
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
                            new ChatMessage { Role = "user", Content = $"{prompt}\n\n{source}" }
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
                    var prompt = $"Summarize the following document {instruction}:\n\n" +
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
