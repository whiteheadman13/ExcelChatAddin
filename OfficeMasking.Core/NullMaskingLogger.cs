using System;

namespace OfficeMasking.Core
{
    /// <summary>
    /// 何も出力しないロガー。ロガー未設定時の既定値。
    /// </summary>
    public sealed class NullMaskingLogger : IMaskingLogger
    {
        public static readonly NullMaskingLogger Instance = new NullMaskingLogger();

        private NullMaskingLogger() { }

        public void LogInfo(string message) { }
        public void LogError(string message) { }
        public void LogException(Exception ex, string context) { }
    }
}
