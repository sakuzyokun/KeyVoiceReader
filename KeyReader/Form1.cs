// KeyVoiceReader.cs
// 単一ファイルで動く WinForms 常駐アプリ
// - グローバル低レベルキーボードフック
// - 説明モード（わかりやすい説明） / 通常モード（キー名）
// - JP/US キーボード配列に配慮（数字列の Shift 記号など）
// - Shift/Control/Alt/Win の組み合わせ読み上げ
// - 読み上げエンジン: System.Speech.Synthesis
// - タスクトレイ常駐、設定ウィンドウ（音声選択/速度/説明モード切替）
// - 設定ウィンドウ上のコントロールにマウスをかざすと読み上げ
// - プロセス優先度: High（必要に応じて変更可）

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Speech.Synthesis;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace KeyReader
{
    /// <summary>
    /// タスクトレイ常駐用 ApplicationContext
    /// </summary>
    public class AppContext : ApplicationContext
    {
        private NotifyIcon tray;
        private SettingsForm settingsForm;

        public AppContext()
        {
            // 音声初期化
            SpeechService.Instance.Initialize();

            // キーフック開始
            KeyboardHook.Instance.Start();

            // トレイメニュー
            var menu = new ContextMenuStrip();
            var miToggleExplain = new ToolStripMenuItem("説明モード", null, (s, e) => ToggleExplain())
            {
                CheckOnClick = true,
                Checked = KeyNameMapper.UseExplainMode
            };
            var miVoice = new ToolStripMenuItem("音声設定", null, (s, e) => ShowSettings());
            var miExit = new ToolStripMenuItem("終了", null, (s, e) => Exit());
            menu.Items.Add(miToggleExplain);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(miVoice);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(miExit);

            tray = new NotifyIcon
            {
                Text = "KeyVoiceReader",
                Visible = true,
                ContextMenuStrip = menu,
                Icon = System.Drawing.SystemIcons.Information
            };

            tray.DoubleClick += (s, e) => ShowSettings();
        }

        private void ToggleExplain()
        {
            KeyNameMapper.UseExplainMode = !KeyNameMapper.UseExplainMode;
            var now = KeyNameMapper.UseExplainMode ? "ON" : "OFF";
            SpeechService.Instance.SpeakQuick($"説明モード {now}");
        }

        private void ShowSettings()
        {
            if (settingsForm == null || settingsForm.IsDisposed)
                settingsForm = new SettingsForm();
            settingsForm.Show();
            settingsForm.Activate();
        }

        private void Exit()
        {
            tray.Visible = false;
            KeyboardHook.Instance.Stop();
            SpeechService.Instance.Dispose();
            Application.Exit();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                tray?.Dispose();
                settingsForm?.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 設定ウィンドウ（音声選択、速度、説明モード切替）
    /// すべてのコントロールにマウスをかざすと、そのテキスト/説明を読み上げる
    /// </summary>
    public class SettingsForm : Form
    {
        ComboBox cmbVoices;
        TrackBar trkRate;
        Label lblRateValue;
        CheckBox chkExplain;
        Button btnClose;
        ToolTip tip;

        public SettingsForm()
        {
            Text = "KeyVoiceReader 設定";
            Width = 460;
            Height = 240;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;

            tip = new ToolTip();

            var lblVoice = new Label { Text = "音声:", Left = 16, Top = 20, Width = 80 };
            cmbVoices = new ComboBox { Left = 100, Top = 16, Width = 320, DropDownStyle = ComboBoxStyle.DropDownList };
            var lblRate = new Label { Text = "速度:", Left = 16, Top = 65, Width = 80 };
            trkRate = new TrackBar { Left = 100, Top = 60, Width = 260, Minimum = -10, Maximum = 10, TickFrequency = 1 };
            lblRateValue = new Label { Left = 370, Top = 65, Width = 50, TextAlign = System.Drawing.ContentAlignment.MiddleLeft };
            chkExplain = new CheckBox { Left = 100, Top = 110, Width = 320, Text = "説明モード（Backspace=1文字前を消す など）" };
            btnClose = new Button { Text = "閉じる", Left = 340, Top = 150, Width = 80 };

            Controls.AddRange(new Control[] { lblVoice, cmbVoices, lblRate, trkRate, lblRateValue, chkExplain, btnClose });

            // 初期値
            cmbVoices.Items.AddRange(SpeechService.Instance.GetInstalledVoiceNames());
            var current = SpeechService.Instance.CurrentVoiceName;
            if (!string.IsNullOrEmpty(current))
            {
                var idx = cmbVoices.Items.IndexOf(current);
                if (idx >= 0) cmbVoices.SelectedIndex = idx; else if (cmbVoices.Items.Count > 0) cmbVoices.SelectedIndex = 0;
            }
            else if (cmbVoices.Items.Count > 0) cmbVoices.SelectedIndex = 0;

            trkRate.Value = SpeechService.Instance.Rate;
            UpdateRateLabel();
            chkExplain.Checked = KeyNameMapper.UseExplainMode;

            // イベント
            cmbVoices.SelectedIndexChanged += (s, e) =>
            {
                var name = cmbVoices.SelectedItem?.ToString();
                if (!string.IsNullOrEmpty(name))
                {
                    SpeechService.Instance.SelectVoice(name);
                    SpeechService.Instance.SpeakQuick($"音声を {name} に変更");
                }
            };
            trkRate.ValueChanged += (s, e) =>
            {
                SpeechService.Instance.Rate = trkRate.Value;
                UpdateRateLabel();
            };
            chkExplain.CheckedChanged += (s, e) =>
            {
                KeyNameMapper.UseExplainMode = chkExplain.Checked;
                var now = chkExplain.Checked ? "ON" : "OFF";
                SpeechService.Instance.SpeakQuick($"説明モード {now}");
                // 逆にするヒントもツールチップで提示
                tip.SetToolTip(chkExplain, "このチェックは説明モード。チェックを外すと通常モード（キー名）\n'説明モードを今のモードの逆にする' → チェック切替");
            };
            btnClose.Click += (s, e) => Close();

            // すべてのコントロールにホバー読み上げを付与
            AttachHoverSpeak(this);

            // 初回ヘルプ
            tip.SetToolTip(cmbVoices, "利用可能な音声を選択");
            tip.SetToolTip(trkRate, "読み上げ速度");
            tip.SetToolTip(chkExplain, "説明モードをON/OFF（ONだと 'Backspace=1文字前を消す' など）");
            tip.SetToolTip(btnClose, "ウィンドウを閉じる（アプリはタスクトレイに常駐）");
        }

        private void UpdateRateLabel()
        {
            lblRateValue.Text = trkRate.Value.ToString();
            SpeechService.Instance.SpeakQuick($"速度 {trkRate.Value}");
        }

        // 子孫コントロールすべてに MouseEnter を再帰的に付与
        private void AttachHoverSpeak(Control root)
        {
            foreach (Control c in root.Controls)
            {
                c.MouseEnter += OnHoverSpeak;
                // アクセシビリティ名や説明があれば採用
                if (string.IsNullOrEmpty(c.AccessibleName)) c.AccessibleName = c.Text;
                if (c is TrackBar && string.IsNullOrEmpty(c.AccessibleDescription)) c.AccessibleDescription = "読み上げ速度のスライダー";
                if (c is ComboBox && string.IsNullOrEmpty(c.AccessibleDescription)) c.AccessibleDescription = "音声選択のコンボボックス";
                AttachHoverSpeak(c);
            }
            root.MouseEnter += OnHoverSpeak;
        }

        private void OnHoverSpeak(object sender, EventArgs e)
        {
            if (sender is Control c)
            {
                // 説明優先 → アクセシブル名 → テキスト → ツールチップ
                string text = c.AccessibleDescription;
                if (string.IsNullOrWhiteSpace(text)) text = c.AccessibleName;
                if (string.IsNullOrWhiteSpace(text)) text = c.Text;
                if (string.IsNullOrWhiteSpace(text)) text = tip.GetToolTip(c);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    SpeechService.Instance.SpeakQuick(text);
                }
            }
        }
    }

    /// <summary>
    /// 音声サービス（単一インスタンス）
    /// </summary>
    public sealed class SpeechService : IDisposable
    {
        private static readonly Lazy<SpeechService> _lazy = new Lazy<SpeechService>(() => new SpeechService());
        public static SpeechService Instance => _lazy.Value;

        private SpeechSynthesizer synth;
        private readonly object _lock = new object();
        private Thread speakThread;
        private AutoResetEvent speakEvent = new AutoResetEvent(false);
        private volatile string pendingText = null; // 最新だけ話す（古い要求は上書き）

        private SpeechService() { }

        public int Rate
        {
            get => synth?.Rate ?? 0;
            set { lock (_lock) { if (synth != null) synth.Rate = Math.Max(-10, Math.Min(10, value)); } }
        }

        public string CurrentVoiceName { get; private set; }

        public void Initialize()
        {
            lock (_lock)
            {
                if (synth != null) return;
                synth = new SpeechSynthesizer();
                synth.Rate = 0;
                try { synth.SetOutputToDefaultAudioDevice(); } catch { }

                // 別スレッドで逐次 Speak（割り込み対応）
                speakThread = new Thread(SpeakWorker)
                {
                    IsBackground = true,
                    Priority = ThreadPriority.AboveNormal,
                    Name = "SpeechWorker"
                };
                speakThread.Start();
            }
        }

        public void SelectVoice(string name)
        {
            lock (_lock)
            {
                try { synth.SelectVoice(name); CurrentVoiceName = name; }
                catch { }
            }
        }

        public string[] GetInstalledVoiceNames()
        {
            lock (_lock)
            {
                return synth?.GetInstalledVoices()?.Select(v => v.VoiceInfo?.Name).Where(n => !string.IsNullOrEmpty(n)).ToArray() ?? Array.Empty<string>();
            }
        }

        // 直近のテキストだけを喋る（連打されるキー向けにスルー）
        public void SpeakQuick(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return;
            pendingText = text;
            speakEvent.Set();
        }

        private void SpeakWorker()
        {
            while (true)
            {
                speakEvent.WaitOne();
                var text = pendingText;
                if (text == null) continue;
                lock (_lock)
                {
                    try
                    {
                        synth.SpeakAsyncCancelAll();
                        synth.SpeakAsync(text);
                    }
                    catch { }
                }
                // 少し待って連打抑止
                Thread.Sleep(20);
            }
        }

        public void Dispose()
        {
            lock (_lock)
            {
                try { synth?.SpeakAsyncCancelAll(); } catch { }
                try { synth?.Dispose(); } catch { }
                synth = null;
            }
        }
    }

    /// <summary>
    /// グローバル低レベルキーボードフック
    /// </summary>
    public sealed class KeyboardHook
    {
        public static KeyboardHook Instance { get; } = new KeyboardHook();

        private KeyboardHook() { }

        private IntPtr hookID = IntPtr.Zero;
        private LowLevelKeyboardProc proc;
        private DateTime lastSpokeAt = DateTime.MinValue;
        private Keys lastKey;

        public void Start()
        {
            if (hookID != IntPtr.Zero) return;
            proc = HookCallback;
            using (var curProcess = Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                hookID = SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        public void Stop()
        {
            if (hookID != IntPtr.Zero)
            {
                UnhookWindowsHookEx(hookID);
                hookID = IntPtr.Zero;
            }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int msg = wParam.ToInt32();
                if (msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN)
                {
                    KBDLLHOOKSTRUCT data = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                    Keys key = (Keys)data.vkCode;

                    // 自分のホットキー抑止などのため必要ならここに条件追加

                    string text = KeyNameMapper.GetReadableName(data.vkCode, data.scanCode);

                    // 連打の抑制（同じキーを超高速に繰り返し喋らない）
                    var now = DateTime.UtcNow;
                    if (key == lastKey && (now - lastSpokeAt).TotalMilliseconds < 50)
                    {
                        // スキップ
                    }
                    else
                    {
                        SpeechService.Instance.SpeakQuick(text);
                        lastKey = key;
                        lastSpokeAt = now;
                    }
                }
            }
            return CallNextHookEx(hookID, nCode, wParam, lParam);
        }

        // WinAPI
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);
    }

    /// <summary>
    /// キー名のわかりやすい表示・説明
    /// </summary>
    public static class KeyNameMapper
    {
        public static bool UseExplainMode { get; set; } = false;

        // JP/USの数字列 Shift 記号
        private static readonly Dictionary<char, string> ShiftSymbolsUS = new Dictionary<char, string>
        {
            ['1'] = "!",
            ['2'] = "@",
            ['3'] = "#",
            ['4'] = "$",
            ['5'] = "%",
            ['6'] = "^",
            ['7'] = "&",
            ['8'] = "*",
            ['9'] = "(",
            ['0'] = ")"
        };
        private static readonly Dictionary<char, string> ShiftSymbolsJP = new Dictionary<char, string>
        {
            ['1'] = "!",
            ['2'] = '"'.ToString(),
            ['3'] = "#",
            ['4'] = "$",
            ['5'] = "%",
            ['6'] = "&",
            ['7'] = "'",
            ['8'] = "(",
            ['9'] = ")",
            ['0'] = ""
        };

        // OEM系・特殊キーのフレンドリーネーム（両配列共通の概念 + JP固有メモ）
        private static readonly Dictionary<Keys, string> Friendly = new Dictionary<Keys, string>
        {
            { Keys.Back, "Backspace" },
            { Keys.Delete, "Delete" },
            { Keys.Enter, "Enter" },
            { Keys.Space, "Space" },
            { Keys.Tab, "Tab" },
            { Keys.Escape, "Escape" },
            { Keys.CapsLock, "Caps Lock" },
            { Keys.NumLock, "Num Lock" },
            { Keys.Scroll, "Scroll Lock" },
            { Keys.PrintScreen, "Print Screen" },
            { Keys.Pause, "Pause" },
            { Keys.Insert, "Insert" },
            { Keys.Home, "Home" },
            { Keys.End, "End" },
            { Keys.PageUp, "Page Up" },
            { Keys.PageDown, "Page Down" },
            { Keys.Left, "Left" },
            { Keys.Right, "Right" },
            { Keys.Up, "Up" },
            { Keys.Down, "Down" },
            { Keys.Apps, "Menu" },
            { Keys.LWin, "Left Windows" },
            { Keys.RWin, "Right Windows" },
            { Keys.Oemtilde, "バッククォート" }, // US:`~ / JP: 半角/全角キーは別（Keys.Kana など）
            { Keys.OemMinus, "ハイフン" }, // US: - _
            { Keys.Oemplus, "イコール" }, // US: = +
            { Keys.OemOpenBrackets, "左角かっこ" }, // [ {
            { Keys.Oem6, "右角かっこ" }, // ] }
            { Keys.Oem5, "バックスラッシュ" }, // \
            { Keys.Oem1, "セミコロン" }, // ; :
            { Keys.Oem7, "クォート" }, // ' "
            { Keys.Oemcomma, "カンマ" }, // , <
            { Keys.OemPeriod, "ピリオド" }, // . >
            { Keys.OemQuestion, "スラッシュ" }, // / ?
            { Keys.Oem102, "不等号キー" }, // JP/ISO: \|/<> の追加キー
            { Keys.KanaMode, "かな" }, // JP: 半角/全角 漢字 切替 (環境で Keys.Kana/KanaMode になる)
            { Keys.HanjaMode, "変換" }, // 参考: 一部環境
            { Keys.JunjaMode, "無変換" },
            { Keys.ProcessKey, "IME処理キー" }
        };

        // 説明モード
        private static readonly Dictionary<Keys, string> Explain = new Dictionary<Keys, string>
        {
            { Keys.Back, "1文字前を消す" },
            { Keys.Delete, "1文字後を消す" },
            { Keys.Enter, "改行" },
            { Keys.Space, "スペース" },
            { Keys.Tab, "次の入力欄へ移動" },
            { Keys.Escape, "キャンセル" },
            { Keys.CapsLock, "英字の大文字固定を切り替え" },
            { Keys.NumLock, "テンキーの数字入力を切り替え" },
            { Keys.Scroll, "スクロールロック（スクロールのオンオフ）" },
            { Keys.PrintScreen, "画面をキャプチャ" },
            { Keys.Pause, "一時停止" },
            { Keys.Insert, "挿入と上書きを切り替え" },
            { Keys.Home, "行頭へ移動" },
            { Keys.End, "行末へ移動" },
            { Keys.PageUp, "1ページ上へ" },
            { Keys.PageDown, "1ページ下へ" },
            { Keys.Left, "左へ移動" },
            { Keys.Right, "右へ移動" },
            { Keys.Up, "上へ移動" },
            { Keys.Down, "下へ移動" },
        };

        // 現在のレイアウトが日本語かどうか
        private static bool IsJapaneseLayout()
        {
            IntPtr hkl = GetKeyboardLayout(0);
            ushort langId = (ushort)((long)hkl & 0xFFFF);
            ushort primary = (ushort)(langId & 0xFF);
            return primary == 0x11; // LANG_JAPANESE
        }

        public static string GetReadableName(uint vkCode, uint scanCode)
        {
            Keys key = (Keys)vkCode;

            bool shift = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
            bool ctrl = (Control.ModifierKeys & Keys.Control) == Keys.Control;
            bool alt = (Control.ModifierKeys & Keys.Alt) == Keys.Alt;
            bool win = (GetAsyncKeyState((int)Keys.LWin) < 0) || (GetAsyncKeyState((int)Keys.RWin) < 0);

            // コンビネーションの前置き（Ctrl/Alt/Shift/Win）
            var mods = new List<string>();
            if (ctrl) mods.Add("Control");
            if (alt) mods.Add("Alt");
            if (win) mods.Add("Windows");
            if (shift) mods.Add("Shift");

            string core = GetKeyCoreName(key, vkCode, shift);
            string baseText = mods.Count > 0 ? string.Join(" + ", mods) + " + " + core : core;

            if (!UseExplainMode) return baseText;

            // 説明モード：修飾キーだけのときはそのまま、文字キーなら補足
            if (Explain.TryGetValue(key, out var desc))
            {
                return desc;
            }
            // 代表的なショートカットの説明（例）
            if (ctrl && key == Keys.C) return "コピー";
            if (ctrl && key == Keys.V) return "貼り付け";
            if (ctrl && key == Keys.X) return "切り取り";
            if (ctrl && key == Keys.Z) return "元に戻す";
            if (alt && key == Keys.F4) return "アプリを終了";

            return baseText; // 既定はそのまま
        }

        private static string GetKeyCoreName(Keys key, uint vkCode, bool shift)
        {
            // 1) まずフレンドリーネーム辞書
            if (Friendly.TryGetValue(key, out var friendly))
                return friendly;

            // 2) 英字（A-Z）：ShiftやCapsLockによる大文字/小文字
            if (key >= Keys.A && key <= Keys.Z)
            {
                bool caps = IsToggled(Keys.CapsLock);
                bool upper = shift ^ caps;
                var ch = key.ToString()[0]; // 'A'..
                return upper ? ch.ToString() : char.ToLower(ch).ToString();
            }

            // 3) 数字（0-9）：レイアウト別のShift記号
            if (key >= Keys.D0 && key <= Keys.D9)
            {
                char digit = (char)('0' + (vkCode - (uint)Keys.D0));
                if (shift)
                {
                    var dict = IsJapaneseLayout() ? ShiftSymbolsJP : ShiftSymbolsUS;
                    if (dict.TryGetValue(digit, out var sym) && !string.IsNullOrEmpty(sym))
                        return sym;
                }
                return digit.ToString();
            }

            // 4) テンキー
            if (key >= Keys.NumPad0 && key <= Keys.NumPad9)
            {
                return ((char)('0' + (vkCode - (uint)Keys.NumPad0))).ToString();
            }
            if (key == Keys.Add) return "+";
            if (key == Keys.Subtract) return "-";
            if (key == Keys.Multiply) return "*";
            if (key == Keys.Divide) return "/";
            if (key == Keys.Decimal) return ".";

            // 5) ファンクションキー
            if (key >= Keys.F1 && key <= Keys.F24)
                return key.ToString();

            // 6) その他/OEM：できるだけ ToUnicode で文字化
            string uni = TryToUnicode(vkCode, shift);
            if (!string.IsNullOrEmpty(uni)) return uni;

            // 7) それでもダメなら enum の名前
            return key.ToString();
        }

        private static string TryToUnicode(uint vk, bool shift)
        {
            byte[] state = new byte[256];
            GetKeyboardState(state);
            // 現在の修飾の状態を反映（Shift押下）
            if (shift) state[0x10] = 0x80; // VK_SHIFT
            // CapsLock 状態も反映
            if (IsToggled(Keys.CapsLock)) state[0x14] = 0x01;

            uint sc = MapVirtualKey(vk, 0);
            StringBuilder sb = new StringBuilder(8);
            int rc = ToUnicode(vk, sc, state, sb, sb.Capacity, 0);
            if (rc > 0)
            {
                var s = sb.ToString();
                if (!string.IsNullOrWhiteSpace(s)) return s;
            }
            return null;
        }

        private static bool IsToggled(Keys key)
        {
            short state = GetKeyState((int)key);
            return (state & 0x0001) != 0;
        }

        // WinAPI for keyboard state
        [DllImport("user32.dll")]
        private static extern int ToUnicode(uint wVirtKey, uint wScanCode, byte[] lpKeyState, [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pwszBuff, int cchBuff, uint wFlags);

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        [DllImport("user32.dll")]
        private static extern bool GetKeyboardState(byte[] lpKeyState);

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int nVirtKey);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        private static extern IntPtr GetKeyboardLayout(uint idThread);
    }
}
