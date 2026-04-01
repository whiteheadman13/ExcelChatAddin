using Microsoft.VisualStudio.TestTools.UnitTesting;

// グローバル状態（環境変数・シングルトン）を使うテストが多いため、並列実行を禁止する
[assembly: DoNotParallelize]
