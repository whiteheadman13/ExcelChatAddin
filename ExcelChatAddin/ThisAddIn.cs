using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;


namespace ExcelChatAddin
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane _myTaskPane;
        private TaskPaneHost _taskPaneHost;

        private const string MENU_TAG = "OfficeChat_SendSelectionToChat";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            AddCellContextMenu();
            this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;

            // ▼ PowerPointと同じ：右側タスクペイン
            _taskPaneHost = new TaskPaneHost();
            _taskPaneHost.SetApplication(this.Application);

            _myTaskPane = this.CustomTaskPanes.Add(_taskPaneHost, "Secure Chat");
            _myTaskPane.Width = 400;
            _myTaskPane.Visible = false;
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            RemoveCellContextMenu();
            this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook wb, ref bool cancel)
        {
            // 事故防止：残留しやすいのでここでも削除
            RemoveCellContextMenu();
        }

        private void AddCellContextMenu()
        {
            try
            {
                var cellBar = this.Application.CommandBars["Cell"];
                RemoveCellContextMenu(); // 二重追加防止

                var btn = (Office.CommandBarButton)cellBar.Controls.Add(
                    Type: Office.MsoControlType.msoControlButton,
                    Temporary: true);

                btn.Caption = "選択範囲をチャットへ転送";
                btn.Tag = MENU_TAG;
                btn.Visible = true;
                btn.Click += Btn_Click;
            }
            catch
            {
                // 必要ならログ
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
            catch
            {
                // 無視でOK
            }
        }

        private void Btn_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                var sel = this.Application.Selection as Excel.Range;
                if (sel == null) return;

                string tsv = RangeToTsv(sel);
                System.Windows.Forms.Clipboard.SetText(tsv);
                ShowChat(tsv);

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "ExcelChatAddin");
            }
        }
        public void ShowChat(string text = "")
        {
            if (_myTaskPane == null) return;

            _myTaskPane.Visible = true;

            if (!string.IsNullOrEmpty(text) && _taskPaneHost != null)
            {
                _taskPaneHost.PassTextToChat(text);
            }
        }


        private string RangeToTsv(Excel.Range range)
        {
            object v = range.Value2;

            // 単一セル
            if (!(v is object[,]))
            {
                return SanitizeCell(v);
            }

            // 複数セル
            var arr = (object[,])v;
            int rowCount = arr.GetLength(0);
            int colCount = arr.GetLength(1);

            var sb = new StringBuilder();

            for (int r = 1; r <= rowCount; r++)
            {
                for (int c = 1; c <= colCount; c++)
                {
                    if (c > 1) sb.Append('\t');
                    sb.Append(SanitizeCell(arr[r, c]));
                }
                if (r < rowCount) sb.AppendLine();
            }

            return sb.ToString();
        }

        private string SanitizeCell(object value)
        {
            string s = value?.ToString() ?? "";
            return s.Replace("\r\n", " ")
                    .Replace("\n", " ")
                    .Replace("\r", " ")
                    .Replace("\t", " ");
        }


        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
