using System;

namespace OfficeMasking.Core
{
    /// <summary>
    /// MaskingEngine が利用するロガーのインターフェース。
    /// 各アドイン側で具象クラスを提供する。
    /// </summary>
    public interface IMaskingLogger
    {
        void LogInfo(string message);
        void LogError(string message);
        void LogException(Exception ex, string context);
    }
}
