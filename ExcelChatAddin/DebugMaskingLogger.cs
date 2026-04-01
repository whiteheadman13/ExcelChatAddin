using System;
using OfficeMasking.Core;

namespace ExcelChatAddin
{
    /// <summary>
    /// OfficeMasking.Core の IMaskingLogger を ExcelChatAddin の DebugLogger に委譲するアダプタ。
    /// </summary>
    internal sealed class DebugMaskingLogger : IMaskingLogger
    {
        public static readonly DebugMaskingLogger Instance = new DebugMaskingLogger();

        private DebugMaskingLogger() { }

        public void LogInfo(string message) => DebugLogger.LogInfo(message);
        public void LogError(string message) => DebugLogger.LogError(message);
        public void LogException(Exception ex, string context) => DebugLogger.LogException(ex, context);
    }
}
