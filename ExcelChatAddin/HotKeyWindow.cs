using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelChatAddin
{
    internal class HotKeyWindow : NativeWindow, IDisposable
    {
        public event Action HotKeyPressed;

        // Create a dedicated message-only window to receive WM_HOTKEY safely.
        public HotKeyWindow()
        {
            var cp = new CreateParams();
            // Message-only window: set parent to HWND_MESSAGE (-3)
            cp.Parent = new IntPtr(-3);
            CreateHandle(cp);
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
            try { DestroyHandle(); } catch { }
        }

        // Expose the created window handle for external APIs
        public IntPtr WindowHandle => this.Handle;
    }
}
