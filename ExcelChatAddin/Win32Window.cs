using System;
using System.Windows.Forms;

namespace ExcelChatAddin
{
    internal sealed class Win32Window : IWin32Window
    {
        public IntPtr Handle { get; }
        public Win32Window(IntPtr handle) => Handle = handle;
    }
}
