using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelChatAddin
{
    public partial class TaskPaneHost : UserControl
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);
        private ElementHost _elementHost;
        private ChatView _chatView;
        private Excel.Application _excelApp;

        public TaskPaneHost()
        {
            InitializeComponent();
            Build();
        }

        private void Build()
        {
            _elementHost = new ElementHost();
            _elementHost.Dock = DockStyle.Fill;
            _elementHost.TabStop = true; // ★フォーカスが入りやすくなる

            _chatView = new ChatView();
            _chatView.SetHost(this);

            _elementHost.Child = _chatView;
            this.Controls.Add(_elementHost);
        }

        public void SetApplication(object application)
        {
            _excelApp = application as Excel.Application;
        }

        public void AppendToInput(string text)
        {
            _chatView?.AppendText(text);
        }

        public void FocusInput()
        {
            // TaskPane表示直後に呼ぶ用
            if (this.IsHandleCreated)
            {
                this.BeginInvoke((Action)(() =>
                {
                    try
                    {
                        _elementHost.Focus();
                        _chatView?.FocusInput();
                    }
                    catch { }
                }));
            }
            else
            {
                try
                {
                    _elementHost.Focus();
                    _chatView?.FocusInput();
                }
                catch { }
            }
        }

        public void SelectExcelRange(string sheetName, string addressA1)
        {
            // ★ Excel COM は WinForms UI スレッドに寄せる（安定化）
            if (this.IsHandleCreated)
            {
                this.BeginInvoke((Action)(() => SelectExcelRangeCore(sheetName, addressA1)));
            }
            else
            {
                SelectExcelRangeCore(sheetName, addressA1);
            }
        }

        private void SelectExcelRangeCore(string sheetName, string addressA1)
        {
            try
            {
                if (_excelApp == null) return;

                var wb = _excelApp.ActiveWorkbook;
                if (wb == null) return;

                var ws = wb.Worksheets.Item[sheetName] as Excel.Worksheet;
                if (ws == null) return;

                ws.Activate();
                var r = ws.Range[addressA1];
                r.Select();

                // ★ここから追加：Excelへフォーカスを戻す
                _excelApp.ActiveWindow?.Activate();
                var hwnd = new IntPtr(_excelApp.Hwnd);
                SetForegroundWindow(hwnd);
                SetFocus(hwnd);
            }
            catch
            {
            }
        }

    }
}
