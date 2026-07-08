using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
using OfficeMasking.Core;




namespace ExcelChatAddin
{
    public partial class ThisAddIn
    {
        private readonly Dictionary<int, Microsoft.Office.Tools.CustomTaskPane> _panesByHwnd
            = new Dictionary<int, Microsoft.Office.Tools.CustomTaskPane>();

        private readonly Dictionary<int, TaskPaneHost> _hostsByHwnd
            = new Dictionary<int, TaskPaneHost>();

        private Office.CommandBarButton _sendBtn;
        private Office.CommandBarButton _manageBtn;
        private Office.CommandBarButton _previewBtn;
        private const string MENU_TAG = "OfficeChat_SendSelectionToChat";
        private bool _maskRegisterDialogOpen = false;

        private bool _menusInitialized = false;
        private DateTime _lastRegisterClick = DateTime.MinValue;
        private DateTime _lastManageClick = DateTime.MinValue;
        private bool _registerDialogOpen = false;
        private bool _manageDialogOpen = false;
        private int _inManageClick = 0;
        private int _inRegisterClick = 0;
        private int _inHotKeyRegister = 0;
        private DateTime _lastHotKey = DateTime.MinValue;
        private long _lastHotKeyTicks = 0; // DateTimeより安定


        private HotKeyWindow _hotKeyWindow;
        private const int HOTKEY_ID_REGISTER = 0x1234;

        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;

        [DllImport("user32.dll", SetLastError = true, EntryPoint = "RegisterHotKey")]
        private static extern bool RegisterHotKeyNative(
            IntPtr hWnd,
            int id,
            uint fsModifiers,
            uint vk
        );

        [DllImport("user32.dll", SetLastError = true, EntryPoint = "UnregisterHotKey")]
        private static extern bool UnregisterHotKeyNative(
            IntPtr hWnd,
            int id
        );




        // 追加：マスキング関連（右クリックに追加するメニュー）
        //private const string MaskMenuTagRegister = "ExcelChatAddin.Mask.Register";
        private const string MaskMenuTagManage = "ExcelChatAddin.Mask.Manage";
        private const string MaskMenuTagPreview = "ExcelChatAddin.Mask.Preview";


        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // OfficeMasking.Core 初期化
            Paths.InitLegacyDllDirectory();
            MaskingEngine.SetLogger(DebugMaskingLogger.Instance);

            // H-1: 外部送信直前ガードの警告ダイアログを差し込む（未マスクの登録語を検出した場合）。
            MaskingSendGuard.ConfirmSendDespiteLeaks = ShowSendLeakConfirmDialog;

            PurgeMaskMenus();     // 全掃除

            AddMaskManageMenu();  // ★通常モード専用
            AddMaskPreviewMenu(); // ★マスキング確認メニュー
            //AddMaskRegisterMenus(); // ★編集モード専用
            RegisterHotKey_CtrlShiftM();

            if (MaskingEngine.Instance.IsAvailable)
            {
                AddCellContextMenu(); // 既存：チャット転送
            }
            else
            {
                RemoveCellContextMenu();
                try { UnregisterHotKeys(); } catch { }
                ShowMaskingUnavailableMessage("Secure Chat");
            }

            this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
        }

        private string BuildMaskingUnavailableMessage()
        {
            var details = MaskingEngine.Instance.AvailabilityErrorMessage;
            if (string.IsNullOrWhiteSpace(details))
            {
                details = "マスキング辞書を読み込めないため、マスキング機能を利用できません。";
            }

            return details + "\n\nこの状態では Secure Chat を使用できません。";
        }

        private void ShowMaskingUnavailableMessage(string title)
        {
            MessageBox.Show(BuildMaskingUnavailableMessage(), title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// H-1: 送信直前ガードが平文残存を検出したときの確認ダイアログ。
        /// 「はい」で送信続行、「いいえ」で中止（MaskingSendGuard が送信を止める）。
        /// </summary>
        private bool ShowSendLeakConfirmDialog(System.Collections.Generic.IReadOnlyList<string> leakedWords)
        {
            string words = string.Join(", ", leakedWords);
            var result = MessageBox.Show(
                "外部LLM（Gemini）へ送信する内容に、マスキング辞書の登録単語が未マスクのまま含まれています。\n\n"
                + "検出された単語: " + words + "\n\n"
                + "このまま送信すると機密情報が外部へ送られる可能性があります。送信を続行しますか？",
                "マスキング警告（送信前チェック）",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);
            return result == DialogResult.Yes;
        }

        private bool EnsureMaskingAvailable(string title)
        {
            if (MaskingEngine.Instance.IsAvailable) return true;
            ShowMaskingUnavailableMessage(title);
            return false;
        }

        private bool EnsureMaskingDataDirConfigured()
        {
            if (Paths.IsMaskingDataDirConfigured) return true;

            MessageBox.Show(
                "環境変数 OFFICE_MASKING_DATA_DIR が未設定です。\n"
                + "Secure Chat を開く前に設定してください。",
                "Secure Chat",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            return false;
        }

        private void AddMaskManageMenu()
        {
            try
            {
                var cb = this.Application.CommandBars["Cell"];
                if (cb == null) return;

                RemoveCommandBarControl(cb, MaskMenuTagManage);

                _manageBtn = (Office.CommandBarButton)cb.Controls.Add(
                    Office.MsoControlType.msoControlButton, Temporary: true);

                _manageBtn.Caption = "辞書管理…";
                _manageBtn.Tag = MaskMenuTagManage;
                _manageBtn.Click += BtnMng_Click;
            }
            catch { }
        }

        private void AddMaskPreviewMenu()
        {
            try
            {
                var cb = this.Application.CommandBars["Cell"];
                if (cb == null) return;

                RemoveCommandBarControl(cb, MaskMenuTagPreview);

                _previewBtn = (Office.CommandBarButton)cb.Controls.Add(
                    Office.MsoControlType.msoControlButton, Temporary: true);

                _previewBtn.Caption = "マスキング確認";
                _previewBtn.Tag = MaskMenuTagPreview;
                _previewBtn.Click += BtnPreview_Click;
            }
            catch { }
        }

        private void RegisterHotKey_CtrlShiftM()
        {
            if (!MaskingEngine.Instance.IsAvailable) return;

            try { UnregisterHotKeys(); } catch { } // ★先に掃除

            try
            {
                System.Diagnostics.Debug.WriteLine("[HotKey] Initializing message-only HotKeyWindow");

                _hotKeyWindow = new HotKeyWindow();
                _hotKeyWindow.HotKeyPressed += () =>
                {
                    System.Diagnostics.Debug.WriteLine("[HotKey] Ctrl+Alt+M pressed!");
                    RunMaskRegisterFromShortcut();
                };

                // Windows API で Ctrl+Alt+M を登録（Ctrl+Shift+M は既に他のアプリで使用中）
                IntPtr hwnd = _hotKeyWindow?.WindowHandle ?? IntPtr.Zero;
                bool ok = false;
                if (hwnd != IntPtr.Zero)
                    ok = RegisterHotKeyNative(hwnd, HOTKEY_ID_REGISTER, MOD_CONTROL | MOD_ALT, (uint)Keys.M);
                System.Diagnostics.Debug.WriteLine($"[HotKey] RegisterHotKey result: {ok}");

                if (!ok)
                {
                    int err = Marshal.GetLastWin32Error();
                    System.Diagnostics.Debug.WriteLine($"[HotKey] RegisterHotKey error code: {err}");
                    MessageBox.Show($"Ctrl+Alt+M の登録に失敗しました。\nWin32Error={err}\n\nこのショートキーは既に別のアプリケーションで使用されている可能性があります。", "ホットキー登録");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HotKey] Exception: {ex}");
                MessageBox.Show($"ホットキー登録中に例外が発生しました。\nエラー: {ex.Message}", "ホットキー登録");
            }
        }



        private void RunMaskRegisterFromShortcut()
        {
            if (!EnsureMaskingAvailable("マスキング登録")) return;

            // ★時間ガードはロック前（ここで return してもロック不要）
            long nowTicks = DateTime.UtcNow.Ticks;
            long last = System.Threading.Interlocked.Read(ref _lastHotKeyTicks);

            // 600ms = 600 * 10,000 ticks
            if (nowTicks - last < 600L * 10_000L)
                return;

            System.Threading.Interlocked.Exchange(ref _lastHotKeyTicks, nowTicks);

            // ★排他（ここから先は必ず finally で解除）
            if (System.Threading.Interlocked.Exchange(ref _inHotKeyRegister, 1) == 1)
                return;

            try
            {
                string selected = TryGetSelectedTextInEditMode();
                if (string.IsNullOrWhiteSpace(selected))
                {
                    MessageBox.Show(
                        "セル編集モードで、登録したい文字列を選択してから Ctrl+Alt+M を押してください。",
                        "マスキング登録");
                    return;
                }

                var rules = MaskingEngine.Instance.GetAllRules();
                if (rules != null && rules.TryGetValue(selected, out var ph))
                {
                    MessageBox.Show($"すでに登録済みです。\n\n対象: {selected}\n置換: {ph}", "マスキング登録");
                    return;
                }

                var owner = new Win32Window(new IntPtr(this.Application.Hwnd));
                using (var dlg = new RegisterDialog(selected))
                {
                    var r = dlg.ShowDialog(owner);
                    if (r != DialogResult.OK) return;

                    if (dlg.IsNewCategory)
                        MaskingEngine.Instance.AddRule(selected, dlg.SelectedCategory, dlg.Meaning, dlg.AliasList, dlg.CaseInsensitive);
                    else
                        MaskingEngine.Instance.AddRuleWithPlaceholder(selected, dlg.SelectedPlaceholder, dlg.Meaning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "マスキング登録");
            }
            finally
            {
                System.Threading.Interlocked.Exchange(ref _inHotKeyRegister, 0);
            }
        }


        private void UnregisterHotKeys()
        {
            try
            {
                IntPtr hwnd = _hotKeyWindow?.WindowHandle ?? IntPtr.Zero;
                if (hwnd != IntPtr.Zero)
                {
                    bool ok = UnregisterHotKeyNative(hwnd, HOTKEY_ID_REGISTER);
                    System.Diagnostics.Debug.WriteLine($"[HotKey] UnregisterHotKey result: {ok}");
                }
            }
            catch { }

            try
            {
                if (_hotKeyWindow != null)
                {
                    _hotKeyWindow.Dispose();
                    _hotKeyWindow = null;
                }
            }
            catch { }
        }

        private void PurgeMaskMenus()
        {
            string[] bars = { "Cell" };

            foreach (var name in bars)
            {
                try
                {
                    var cb = this.Application.CommandBars[name];
                    if (cb == null) continue;

                    //RemoveCommandBarControl(cb, MaskMenuTagRegister);
                    RemoveCommandBarControl(cb, MaskMenuTagManage);
                    RemoveCommandBarControl(cb, MaskMenuTagPreview);
                }
                catch { }
            }
        }



        //private void EnsureMaskMenus()
        //{
        //    if (_menusInitialized) return;
        //    _menusInitialized = true;

        //    // ★これだけにする（最重要）
        //    TryAddMaskMenusToBar("Cell");
        //}
        //private void TryAddMaskMenusToBar(string commandBarName)
        //{
        //    try
        //    {
        //        var cb = this.Application.CommandBars[commandBarName];
        //        if (cb == null) return;

        //        // 同Tagを全削除（複数残っていても全消し）
        //        RemoveCommandBarControl(cb, MaskMenuTagRegister);
        //        RemoveCommandBarControl(cb, MaskMenuTagManage);

        //        // Register
        //        var btnReg = (Office.CommandBarButton)cb.Controls.Add(
        //            Office.MsoControlType.msoControlButton, Temporary: true);
        //        btnReg.Caption = "選択文字列をマスキング登録…";
        //        btnReg.Tag = MaskMenuTagRegister;
        //        btnReg.Click += BtnReg_Click;

        //        // Manage
        //        var btnMng = (Office.CommandBarButton)cb.Controls.Add(
        //            Office.MsoControlType.msoControlButton, Temporary: true);
        //        btnMng.Caption = "辞書管理…";
        //        btnMng.Tag = MaskMenuTagManage;
        //        btnMng.Click += BtnMng_Click;
        //    }
        //    catch { }
        //}
        private void CleanupMaskMenus()
        {
            string[] bars = { "Cell", "Text", "Edit", "Formula Bar" };
            foreach (var name in bars)
            {
                try
                {
                    var cb = this.Application.CommandBars[name];
                    if (cb == null) continue;
                    //RemoveCommandBarControl(cb, MaskMenuTagRegister);
                    RemoveCommandBarControl(cb, MaskMenuTagManage);
                    RemoveCommandBarControl(cb, MaskMenuTagPreview);
                }
                catch { }
            }
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook wb, ref bool cancel)
        {
            try { RemoveCellContextMenu(); } catch { }

            try { CleanupMaskMenus(); } catch { }
            try { UnregisterHotKeys(); } catch { }
        }



        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                if (_sendBtn != null) _sendBtn.Click -= Btn_Click;
            }
            catch { }
            try
            {
                if (_manageBtn != null) _manageBtn.Click -= BtnMng_Click;
            }
            catch { }
            try
            {
                if (_previewBtn != null) _previewBtn.Click -= BtnPreview_Click;
            }
            catch { }

            RemoveCellContextMenu();

            try
            {
                //this.Application.SheetBeforeRightClick -= Application_SheetBeforeRightClick;
                this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
                try { UnregisterHotKeys(); } catch { }
            }
            catch { }
        }

        

        // =========================================================
        // 右クリック時：メニューを差し込む（毎回）
        // =========================================================
        private void Application_SheetBeforeRightClick(object Sh, Excel.Range Target, ref bool Cancel)
        {
            //try
            //{
            //    // セルの右クリックメニュー
            //    Office.CommandBar cb = this.Application.CommandBars["Cell"];

            //    // 既存の同Tagボタンを消す（重複防止）
            //    RemoveCommandBarControl(cb, MaskMenuTagRegister);
            //    RemoveCommandBarControl(cb, MaskMenuTagManage);

            //    // ① 選択文字列をマスキング登録
            //    var btnReg = (Office.CommandBarButton)cb.Controls.Add(
            //        Office.MsoControlType.msoControlButton, Temporary: true);
            //    btnReg.Caption = "選択文字列をマスキング登録…";
            //    btnReg.Tag = MaskMenuTagRegister;
            //    btnReg.Click += BtnReg_Click;

            //    // ② 辞書管理
            //    var btnMng = (Office.CommandBarButton)cb.Controls.Add(
            //        Office.MsoControlType.msoControlButton, Temporary: true);
            //    btnMng.Caption = "辞書管理…";
            //    btnMng.Tag = MaskMenuTagManage;
            //    btnMng.Click += BtnMng_Click;
            //}
            //catch
            //{
            //    // 右クリックメニューは環境差があるので握りつぶしでOK
            //}
        }

        private static void RemoveCommandBarControl(Office.CommandBar cb, string tag)
        {
            try
            {
                for (int i = cb.Controls.Count; i >= 1; i--)
                {
                    var c = cb.Controls[i];
                    if (c != null && string.Equals(c.Tag, tag, StringComparison.OrdinalIgnoreCase))
                        c.Delete();
                }
            }
            catch { }
        }
        

        private DateTime _lastRegClick = DateTime.MinValue;
        // =========================================================
        // ① マスキング登録…
        // =========================================================
        //private void BtnReg_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        //{
           
        //    // ★同時発火（複数バー）を完全に止める
        //    if (System.Threading.Interlocked.Exchange(ref _inRegisterClick, 1) == 1)
        //        return;

        //    // 1) 同一クリック多重発火を時間で弾く（ExcelのCOMイベント対策）
        //    var now = DateTime.UtcNow;
        //    if ((now - _lastRegisterClick).TotalMilliseconds < 800) return;
        //    _lastRegisterClick = now;

        //    // 2) ダイアログの多重表示を弾く
        //    if (_registerDialogOpen) return;
        //    _registerDialogOpen = true;

        //    try
        //    {
        //        string selected = TryGetSelectedTextInEditMode();
        //        if (string.IsNullOrWhiteSpace(selected))
        //        {
        //            MessageBox.Show("セル編集モードで、登録したい文字列を選択してから実行してください。", "マスキング登録");
        //            return;
        //        }

        //        // ★既に登録済みチェック（メッセージはここで1回だけ）
        //        var rules = MaskingEngine.Instance.GetAllRules();
        //        if (rules != null && rules.TryGetValue(selected, out var ph))
        //        {
        //            MessageBox.Show($"すでに登録済みです。\n\n対象: {selected}\n置換: {ph}", "マスキング登録");
        //            return;
        //        }

        //        var owner = new Win32Window(new IntPtr(this.Application.Hwnd));

        //        using (var dlg = new RegisterDialog(selected))
        //        {
        //            var r = dlg.ShowDialog(owner);
        //            if (r != DialogResult.OK) return;

        //            if (dlg.IsNewCategory)
        //                MaskingEngine.Instance.AddRule(selected, dlg.SelectedCategory);
        //            else
        //                MaskingEngine.Instance.AddRuleWithPlaceholder(selected, dlg.SelectedPlaceholder);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString(), "マスキング登録");
        //    }
        //    finally
        //    {
        //        _registerDialogOpen = false;
        //    }
        //}

        private void BtnPreview_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (!EnsureMaskingAvailable("マスキング確認")) return;

            try
            {
                string text = GetSelectedRangeText();
                if (string.IsNullOrWhiteSpace(text))
                {
                    MessageBox.Show("セルを選択してから実行してください。", "マスキング確認");
                    return;
                }

                string masked = MaskingEngine.Instance.Mask(text);

                var owner = new System.Windows.Interop.WindowInteropHelper(new System.Windows.Window());
                var win = new MaskPreviewWindow(masked);
                var helper = new System.Windows.Interop.WindowInteropHelper(win);
                helper.Owner = new IntPtr(this.Application.Hwnd);
                win.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "マスキング確認");
            }
        }

        private string GetSelectedRangeText()
        {
            try
            {
                var sel = this.Application.Selection as Excel.Range;
                if (sel == null) return "";

                object v = sel.Value2;
                if (v == null) return "";

                if (!(v is object[,]))
                    return Convert.ToString(v) ?? "";

                var arr = (object[,])v;
                int r1 = arr.GetLowerBound(0), r2 = arr.GetUpperBound(0);
                int c1 = arr.GetLowerBound(1), c2 = arr.GetUpperBound(1);

                var sb = new System.Text.StringBuilder();
                for (int r = r1; r <= r2; r++)
                {
                    for (int c = c1; c <= c2; c++)
                    {
                        if (c > c1) sb.Append('\t');
                        sb.Append(arr[r, c]?.ToString() ?? "");
                    }
                    if (r < r2) sb.AppendLine();
                }
                return sb.ToString();
            }
            catch { return ""; }
        }

        private void BtnMng_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ShowDictionaryManager();
        }

        public void ShowDictionaryManager()
        {
            if (!EnsureMaskingAvailable("辞書管理")) return;

            // ★完全排他（同時発火を物理的に止める）
            if (System.Threading.Interlocked.Exchange(ref _inManageClick, 1) == 1)
                return;

            try
            {
                var owner = new Win32Window(new IntPtr(this.Application.Hwnd));
                using (var dlg = new DictionaryManager())
                {
                    dlg.ShowDialog(owner);
                }
            }
            finally
            {
                System.Threading.Interlocked.Exchange(ref _inManageClick, 0);
            }
        }

        /// <summary>
        /// デバッグ用: マスク→送信→アンマスクの往復を目視確認するフォームを開く。
        /// </summary>
        public void ShowMaskingDebug()
        {
            try
            {
                var owner = new Win32Window(new IntPtr(this.Application.Hwnd));
                using (var dlg = new MaskingDebugForm())
                {
                    dlg.ShowDialog(owner);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "マスキング診断（デバッグ）");
            }
        }

        public void ShowTableSchemaSettings()
        {
            try
            {
                var owner = new Win32Window(new IntPtr(this.Application.Hwnd));
                using (var dlg = new IssueSchemaSettingsDialog(this.Application))
                {
                    dlg.ShowDialog(owner);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "表設定");
            }
        }

        public void ShowIssueSchemaSettings()
        {
            // 既存呼び出し互換
            ShowTableSchemaSettings();
        }

        public void ShowTableRelationSettings()
        {
            try
            {
                var owner = new Win32Window(new IntPtr(this.Application.Hwnd));
                using (var dlg = new TableRelationSettingsDialog(this.Application))
                {
                    dlg.ShowDialog(owner);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "関係設定");
            }
        }

        public void ShowSchemaTemplateInsert()
        {
            try
            {
                var items = SchemaTemplateManager.LoadAll();
                if (items.Count == 0)
                {
                    MessageBox.Show("保存済みのテンプレートがありません。", "テンプレートから挿入");
                    return;
                }

                var owner = new Win32Window(new IntPtr(this.Application.Hwnd));

                SchemaTemplateEntry tmpl;
                using (var dlg = new SchemaTemplateListDialog())
                {
                    if (dlg.ShowDialog(owner) != DialogResult.OK || dlg.SelectedTemplate == null) return;
                    tmpl = dlg.SelectedTemplate;
                }

                string newTableName;
                using (var nameDlg = new SchemaTemplateTableNameDialog(tmpl.Name))
                {
                    if (nameDlg.ShowDialog(owner) != DialogResult.OK) return;
                    newTableName = nameDlg.TableName;
                }

                if (string.IsNullOrWhiteSpace(newTableName))
                {
                    MessageBox.Show("テーブル名を入力してください。", "入力エラー");
                    return;
                }

                if (ExcelTableExists(newTableName))
                {
                    MessageBox.Show($"テーブル「{newTableName}」は既にExcelブック内に存在します。\n別のテーブル名を指定してください。",
                        "テーブル名重複", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var cols = (tmpl.Columns ?? new List<IssueSchemaColumn>())
                    .Select(c => new IssueSchemaColumn
                    {
                        ColumnLetter = c.ColumnLetter,
                        ColumnName = c.ColumnName,
                        IsKey = c.IsKey,
                        IsRequired = c.IsRequired,
                        ValueType = c.ValueType,
                        AllowedValues = c.AllowedValues != null ? new List<string>(c.AllowedValues) : new List<string>(),
                        ExampleValue = c.ExampleValue,
                        Meaning = c.Meaning,
                        UpdateMode = c.UpdateMode
                    })
                    .ToList();

                if (cols.Count == 0)
                {
                    MessageBox.Show("テンプレートに列定義がありません。", "テンプレートから挿入");
                    return;
                }

                var keyCols = cols.Where(x => x.IsKey).ToList();
                var cfg = new IssueSchemaConfig
                {
                    TableName = newTableName,
                    SheetName = newTableName,
                    HeaderRow = Math.Max(1, tmpl.HeaderRow),
                    DataStartRow = Math.Max(2, tmpl.DataStartRow),
                    ValuePolicy = "strict",
                    KeyColumnLetter = keyCols.Count > 0 ? keyCols[0].ColumnLetter : cols[0].ColumnLetter,
                    Columns = cols
                };

                var store = IssueSchemaManager.LoadStore();
                IssueSchemaManager.Upsert(store, cfg);
                IssueSchemaManager.SaveStore(store);

                CreateNewSheetWithTable(cfg);

                MessageBox.Show($"テンプレート「{tmpl.Name}」から新しいテーブル「{newTableName}」を作成しました。",
                    "テンプレートから挿入");
            }
            catch (Exception ex)
            {
                MessageBox.Show("テンプレートからの挿入に失敗しました: " + ex.Message, "テンプレートから挿入");
            }
        }

        public void ShowSchemaTemplateManager()
        {
            try
            {
                var owner = new Win32Window(new IntPtr(this.Application.Hwnd));
                using (var dlg = new SchemaTemplateListDialog(manageOnly: true))
                {
                    dlg.ShowDialog(owner);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "テンプレート管理");
            }
        }

        private bool ExcelTableExists(string tableName)
        {
            try
            {
                var wb = this.Application?.ActiveWorkbook;
                if (wb == null) return false;

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    if (ws.ListObjects == null) continue;
                    foreach (Excel.ListObject lo in ws.ListObjects)
                    {
                        if (string.Equals(lo.Name, tableName, StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                }
            }
            catch { }
            return false;
        }

        private void CreateNewSheetWithTable(IssueSchemaConfig cfg)
        {
            if (cfg == null || cfg.Columns == null || cfg.Columns.Count == 0) return;

            var wb = this.Application?.ActiveWorkbook;
            if (wb == null) return;

            var ws = wb.Worksheets.Add() as Excel.Worksheet;
            if (ws == null) return;
            try { ws.Name = cfg.TableName; } catch { }

            foreach (var c in cfg.Columns)
            {
                int col = ColumnLetterToIndex(c.ColumnLetter);
                if (col <= 0) continue;
                var headerCell = ws.Cells[cfg.HeaderRow, col] as Excel.Range;
                if (headerCell != null) headerCell.Value2 = c.ColumnName;
            }

            int minCol = cfg.Columns.Min(c => ColumnLetterToIndex(c.ColumnLetter));
            int maxCol = cfg.Columns.Max(c => ColumnLetterToIndex(c.ColumnLetter));
            if (minCol <= 0 || maxCol <= 0) return;

            int endRow = Math.Max(cfg.DataStartRow, cfg.HeaderRow + 1);
            var topLeft = ws.Cells[cfg.HeaderRow, minCol] as Excel.Range;
            var bottomRight = ws.Cells[endRow, maxCol] as Excel.Range;
            var tableRange = ws.Range[topLeft, bottomRight];
            if (tableRange == null) return;

            try
            {
                var lo = ws.ListObjects.Add(
                    Excel.XlListObjectSourceType.xlSrcRange,
                    tableRange,
                    Type.Missing,
                    Excel.XlYesNoGuess.xlYes,
                    Type.Missing);

                if (lo != null)
                {
                    try { lo.Name = cfg.TableName; } catch { }
                }
            }
            catch { }
        }

        private static int ColumnLetterToIndex(string letter)
        {
            if (string.IsNullOrWhiteSpace(letter)) return 0;
            int index = 0;
            foreach (var ch in letter.Trim().ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z') return 0;
                index = index * 26 + (ch - 'A' + 1);
            }
            return index;
        }



        private string TryGetSelectedTextInEditMode()
        {
            try
            {
                // クリップボード退避
                string before = "";
                try { before = Clipboard.ContainsText() ? Clipboard.GetText() : ""; } catch { }

                // 変化検知しやすいように一旦クリア（空にできない環境もあるので try）
                try { Clipboard.Clear(); } catch { }

                // Excelに対して「コピー」を実行（SendKeysより安定）
                try
                {
                    this.Application.CommandBars.ExecuteMso("Copy");
                }
                catch
                {
                    // ExecuteMso が効かない環境は最後の手として SendKeys
                    SendKeys.SendWait("^c");
                }

                System.Threading.Thread.Sleep(80);

                // コピー結果取得
                string copied = "";
                try { copied = Clipboard.ContainsText() ? Clipboard.GetText() : ""; } catch { }

                // クリップボード復元（親切）
                try { Clipboard.SetText(before); } catch { }

                copied = (copied ?? "").Trim();

                // 何も取れてない or 変化してないなら「選択が取れてない」扱い
                if (string.IsNullOrWhiteSpace(copied)) return "";
                if (string.Equals(copied, before, StringComparison.Ordinal)) return "";

                // 「セル全体コピー」を弾きたい場合はここで比較（必要なら）
                // var full = Convert.ToString((this.Application.ActiveCell as Excel.Range)?.Text) ?? "";
                // if (!string.IsNullOrEmpty(full) && string.Equals(copied, full.Trim(), StringComparison.Ordinal)) return "";

                return copied;
            }
            catch
            {
                return "";
            }
        }




        // =========================================================
        // ② 辞書管理…
        // =========================================================
        

        // 選択セルのテキストを取得（複数なら左上セル）
        private string GetSelectedCellText()
        {
            try
            {
                var sel = this.Application.Selection as Excel.Range;
                if (sel == null) return "";

                var cell = sel.Cells[1, 1] as Excel.Range;

                // 表示文字（Text）優先。空なら Value2
                string t = Convert.ToString(cell.Text);
                if (!string.IsNullOrWhiteSpace(t)) return t.Trim();

                var v = cell.Value2;
                return v != null ? Convert.ToString(v).Trim() : "";
            }
            catch { return ""; }
        }

        // =========================================================
        // 既存：セルメニューに「選択範囲をチャットへ転送」
        // =========================================================
        private void AddCellContextMenu()
        {
            if (!MaskingEngine.Instance.IsAvailable) return;

            try
            {
                var cellBar = this.Application.CommandBars["Cell"];
                RemoveCellContextMenu();

                _sendBtn = (Office.CommandBarButton)cellBar.Controls.Add(
                    Type: Office.MsoControlType.msoControlButton,
                    Temporary: true);

                _sendBtn.Caption = "選択範囲をチャットへ転送";
                _sendBtn.Tag = MENU_TAG;
                _sendBtn.Visible = true;

                _sendBtn.Click -= Btn_Click;
                _sendBtn.Click += Btn_Click;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "ExcelChatAddin(AddCellContextMenu)");
            }
        }

        private void RemoveCellContextMenu()
        {
            try
            {
                var cellBar = this.Application.CommandBars["Cell"];
                foreach (Office.CommandBarControl c in cellBar.Controls)
                {
                    if (c.Tag == MENU_TAG)
                    {
                        c.Delete();
                        break;
                    }
                }
            }
            catch { }
        }

        // 既存：@range トークン追加
        private void Btn_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (!EnsureMaskingAvailable("Secure Chat")) return;

            try
            {
                var sel = this.Application.Selection as Excel.Range;
                if (sel == null) return;

                var ws = sel.Worksheet as Excel.Worksheet;
                string sheetName = ws?.Name ?? "";
                string addressA1 = sel.Address[false, false, Excel.XlReferenceStyle.xlA1];

                // トークン生成
                string token = $"@range({sheetName},{addressA1}) ";

                // ペイン表示
                ShowChat();

                // 入力欄に追記
                AppendRangeTokenToInput(token);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "ExcelChatAddin(Btn_Click)");
            }
        }

        public void ShowChat()
        {
            if (!EnsureMaskingDataDirConfigured()) return;
            if (!EnsureMaskingAvailable("Secure Chat")) return;

            try
            {
                var win = this.Application.ActiveWindow;
                if (win == null)
                {
                    MessageBox.Show("ActiveWindow is null", "ExcelChatAddin");
                    return;
                }

                int hwnd = win.Hwnd;

                if (!_panesByHwnd.TryGetValue(hwnd, out var pane) || pane == null)
                {
                    var host = new TaskPaneHost();
                    host.SetApplication(this.Application);

                    pane = this.CustomTaskPanes.Add(host, "Secure Chat", win);
                    pane.Width = 400;

                    _panesByHwnd[hwnd] = pane;
                    _hostsByHwnd[hwnd] = host;
                }

                pane.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "ExcelChatAddin(ShowChat)");
            }
        }

        public void ClearHighlights()
        {
            try
            {
                var win = this.Application.ActiveWindow;
                if (win == null) return;
                int hwnd = win.Hwnd;
                if (_hostsByHwnd.TryGetValue(hwnd, out var host) && host != null)
                {
                    host.ClearHighlights();
                }
                else
                {
                    MessageBox.Show("チャットパネルが開かれていません。", "ハイライト解除");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ハイライト解除エラー");
            }
        }

        private void AppendRangeTokenToInput(string token)
        {
            var win = this.Application.ActiveWindow;
            if (win == null) return;

            int hwnd = win.Hwnd;

            if (_hostsByHwnd.TryGetValue(hwnd, out var host) && host != null)
            {
                host.AppendToInput(token);
            }
        }



        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new ChatRibbon();
        }
    }
}
