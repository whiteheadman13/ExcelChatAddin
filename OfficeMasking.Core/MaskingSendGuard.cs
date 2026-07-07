using System;
using System.Collections.Generic;

namespace OfficeMasking.Core
{
    /// <summary>
    /// 外部LLM（Gemini 等）への送信直前の最終セーフティネット（H-1）。
    /// マスク呼び忘れ・新規送信経路の追加漏れがあっても、辞書の登録単語が平文のまま
    /// 外部へ送信される前にユーザーへ警告し、送信を中止できるようにする。
    ///
    /// UI（警告ダイアログ）は Core から切り離し、アプリ側が <see cref="ConfirmSendDespiteLeaks"/> を
    /// 起動時に差し込む。未設定のまま平文残存を検出した場合は「安全側＝送信中止」で例外を投げる。
    /// ローカルLLM（Ollama 等）は生データ送信が仕様のため、このガードは呼ばないこと。
    /// </summary>
    public static class MaskingSendGuard
    {
        /// <summary>
        /// 平文残存を検出したときに呼ばれる確認関数（true=送信続行、false=中止）。
        /// アプリ側で MessageBox 等を差し込む。null の場合は安全側（中止）で例外を投げる。
        /// </summary>
        public static Func<IReadOnlyList<string>, bool> ConfirmSendDespiteLeaks;

        /// <summary>
        /// 外部送信ペイロード（システムプロンプト・ユーザープロンプト等）に登録単語が
        /// 平文のまま残っていないか検査する。残存があればユーザーへ警告し、
        /// 中止が選ばれた（または確認関数が未設定の）場合は OperationCanceledException で送信を停止する。
        /// </summary>
        public static void EnsureSafeForExternalSend(params string[] payloadParts)
        {
            EnsureSafeForExternalSend(MaskingEngine.Instance, ConfirmSendDespiteLeaks, payloadParts);
        }

        internal static void EnsureSafeForExternalSend(
            MaskingEngine engine,
            Func<IReadOnlyList<string>, bool> confirm,
            params string[] payloadParts)
        {
            if (engine == null || payloadParts == null) return;

            var leaked = new List<string>();
            foreach (var part in payloadParts)
            {
                if (string.IsNullOrEmpty(part)) continue;
                foreach (var word in engine.FindRegisteredWordsIn(part))
                {
                    if (!leaked.Contains(word)) leaked.Add(word);
                }
            }
            if (leaked.Count == 0) return;

            // 平文残存を検出。確認関数が無ければ安全側＝中止。
            bool proceed = confirm != null && confirm(leaked);
            if (!proceed)
            {
                throw new OperationCanceledException(
                    "マスキング辞書の登録単語が未マスクのまま外部LLMへ送信されようとしたため、送信を中止しました。");
            }
        }
    }
}
