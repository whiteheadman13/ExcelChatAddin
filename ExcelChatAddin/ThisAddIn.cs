using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelChatAddin
{
    public partial class ThisAddIn
    {
        private readonly Dictionary<int, Microsoft.Office.Tools.CustomTaskPane> _panesByHwnd
            = new Dictionary<int, Microsoft.Office.Tools.CustomTaskPane>();

        private readonly Dictionary<int, TaskPaneHost> _hostsByHwnd
            = new Dictionary<int, TaskPaneHost>();

        private Office.CommandBarButton _sendBtn;
        private const string MENU_TAG = "OfficeChat_SendSelectionToChat";

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            AddCellContextMenu();
            this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try { if (_sendBtn != null) _sendBtn.Click -= Btn_Click; } catch { }

            RemoveCellContextMenu();
            this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook wb, ref bool cancel)
        {
            RemoveCellContextMenu();
        }

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

        private void Btn_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                var sel = this.Application.Selection as Excel.Range;
                if (sel == null) return;

                var ws = sel.Worksheet as Excel.Worksheet;
                string sheetName = ws?.Name ?? "";
                string addressA1 = sel.Address[false, false, Excel.XlReferenceStyle.xlA1];

                // トークン生成（PowerPointの @slide と同じノリ）
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
