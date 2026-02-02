using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelChatAddin
{
    internal class HotKeyWindow : NativeWindow, IDisposable
    {
        public event Action HotKeyPressed;

        public HotKeyWindow(IntPtr hwnd)
        {
            AssignHandle(hwnd);
        }

        protected override void WndProc(ref Message m)
        {
            const int WM_HOTKEY = 0x0312;
            if (m.Msg == WM_HOTKEY)
            {
                HotKeyPressed?.Invoke();
            }
            base.WndProc(ref m);
        }

        public void Dispose()
        {
            try { ReleaseHandle(); } catch { }
        }
    }
}
