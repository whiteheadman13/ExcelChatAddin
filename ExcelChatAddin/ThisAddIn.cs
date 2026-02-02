using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;




namespace ExcelChatAddin
{
    public partial class ThisAddIn
    {
        private readonly Dictionary<int, Microsoft.Office.Tools.CustomTaskPane> _panesByHwnd
            = new Dictionary<int, Microsoft.Office.Tools.CustomTaskPane>();

        private readonly Dictionary<int, TaskPaneHost> _hostsByHwnd
            = new Dictionary<int, TaskPaneHost>();

        // 既存：チャットへ転送
        private Office.CommandBarButton _sendBtn;
        private const string MENU_TAG = "OfficeChat_SendSelectionToChat";
        private bool _maskRegisterDialogOpen = false;

        private bool _menusInitialized = false;
        private DateTime _lastRegisterClick = DateTime.MinValue;
        private DateTime _lastManageClick = DateTime.MinValue;
        private bool _registerDialogOpen = false;
        private bool _manageDialogOpen = false;
        private int _inManageClick = 0;
        private int _inRegisterClick = 0;

        private HotKeyWindow _hotKeyWindow;
        private const int HOTKEY_ID_REGISTER = 0x1234;

        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;

        [DllImport("user32.dll", SetLastError = true, EntryPoint = "RegisterHotKey")]
        private static extern bool RegisterHotKeyNative(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true, EntryPoint = "UnregisterHotKey")]
        private static extern bool UnregisterHotKeyNative(IntPtr hWnd, int id);






        // 追加：マスキング関連（右クリックに追加するメニュー）
        //private const string MaskMenuTagRegister = "ExcelChatAddin.Mask.Register";
        private const string MaskMenuTagManage = "ExcelChatAddin.Mask.Manage";


        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            AddCellContextMenu(); // 既存：チャット転送

            PurgeMaskMenus();     // 全掃除

            AddMaskManageMenu();  // ★通常モード専用
            //AddMaskRegisterMenus(); // ★編集モード専用
            RegisterHotKey_CtrlShiftM();
            this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
        }
        private void AddMaskManageMenu()
        {
            try
            {
                var cb = this.Application.CommandBars["Cell"];
                if (cb == null) return;

                RemoveCommandBarControl(cb, MaskMenuTagManage);

                var btnMng = (Office.CommandBarButton)cb.Controls.Add(
                    Office.MsoControlType.msoControlButton, Temporary: true);

                btnMng.Caption = "辞書管理…";
                btnMng.Tag = MaskMenuTagManage;
                btnMng.Click += BtnMng_Click;
            }
            catch { }
        }

        private void RegisterHotKey_CtrlShiftM()
        {
            IntPtr hwnd = new IntPtr(this.Application.Hwnd);

            _hotKeyWindow = new HotKeyWindow(hwnd);
            _hotKeyWindow.HotKeyPressed += () =>
            {
                try
                {
                    System.Windows.Forms.Control.FromHandle(hwnd)?.BeginInvoke((Action)(() =>
                    {
                        RunMaskRegisterFromShortcut();
                    }));
                }
                catch
                {
                    RunMaskRegisterFromShortcut();
                }
            };

            bool ok = RegisterHotKeyNative(hwnd, HOTKEY_ID_REGISTER, MOD_CONTROL | MOD_SHIFT, (uint)Keys.M);


            if (!ok)
            {
                int err = Marshal.GetLastWin32Error();
                MessageBox.Show(
                    $"Ctrl+Shift+M のショートカット登録に失敗しました。\n\nWin32Error: {err}\n他アプリ/Excel設定と競合している可能性があります。",
                    "マスキング登録");
            }
        }

        private void RunMaskRegisterFromShortcut()
        {
            try
            {
                string selected = TryGetSelectedTextInEditMode();
                if (string.IsNullOrWhiteSpace(selected))
                {
                    MessageBox.Show(
                        "セル編集モードで、登録したい文字列を選択してから Ctrl+Shift+M を押してください。",
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
                        MaskingEngine.Instance.AddRule(selected, dlg.SelectedCategory);
                    else
                        MaskingEngine.Instance.AddRuleWithPlaceholder(selected, dlg.SelectedPlaceholder);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "マスキング登録");
            }
        }

        private void UnregisterHotKeys()
        {
            try
            {
                IntPtr hwnd = new IntPtr(this.Application.Hwnd);
                UnregisterHotKeyNative(hwnd, HOTKEY_ID_REGISTER);
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

        //private void TryAddRegisterToBar(string barName)
        //{
        //    try
        //    {
        //        var cb = this.Application.CommandBars[barName];
        //        if (cb == null) return;

        //        RemoveCommandBarControl(cb, MaskMenuTagRegister);

        //        var btnReg = (Office.CommandBarButton)cb.Controls.Add(
        //            Office.MsoControlType.msoControlButton, Temporary: true);
        //        btnReg.Caption = "選択文字列をマスキング登録…";
        //        btnReg.Tag = MaskMenuTagRegister;
        //        btnReg.Click += BtnReg_Click;
        //    }
        //    catch { }
        //}

        //private void TryAddRegisterToFirstExistingBar(string[] bars)
        //{
        //    foreach (var name in bars)
        //    {
        //        try
        //        {
        //            var cb = this.Application.CommandBars[name];
        //            if (cb == null) continue;

        //            RemoveCommandBarControl(cb, MaskMenuTagRegister);

        //            var btnReg = (Office.CommandBarButton)cb.Controls.Add(
        //                Office.MsoControlType.msoControlButton, Temporary: true);
        //            btnReg.Caption = "選択文字列をマスキング登録…";
        //            btnReg.Tag = MaskMenuTagRegister;
        //            btnReg.Click += BtnReg_Click;

        //            break; // 1つだけ
        //        }
        //        catch { }
        //    }
        //}


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

        private void BtnMng_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
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
                _inManageClick = 0;
            }
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


    }
}
