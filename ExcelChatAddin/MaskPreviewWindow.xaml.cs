using System;
using System.Windows;

namespace ExcelChatAddin
{
    public partial class MaskPreviewWindow : Window
    {
        public MaskPreviewWindow(string maskedText)
        {
            InitializeComponent();
            BodyBox.Text = maskedText ?? "";
        }

        private void BtnRegister_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selected = BodyBox.SelectedText?.Trim();
                if (string.IsNullOrWhiteSpace(selected))
                {
                    MessageBox.Show("プレビュー内で登録したい文字列を選択してから実行してください。", "マスキング登録");
                    return;
                }

                using (var dlg = new RegisterDialog(selected))
                {
                    var result = dlg.ShowDialog();
                    if (result != System.Windows.Forms.DialogResult.OK) return;

                    string placeholder;
                    if (dlg.IsNewCategory)
                    {
                        MaskingEngine.Instance.AddRule(selected, dlg.SelectedCategory);
                        var rules = MaskingEngine.Instance.GetAllRules();
                        if (!rules.TryGetValue(selected, out placeholder) || string.IsNullOrWhiteSpace(placeholder))
                        {
                            MessageBox.Show("登録に失敗しました（プレースホルダ取得不可）。", "マスキング登録");
                            return;
                        }
                    }
                    else
                    {
                        placeholder = dlg.SelectedPlaceholder;
                        if (string.IsNullOrWhiteSpace(placeholder))
                        {
                            MessageBox.Show("既存タグが選択されていません。", "マスキング登録");
                            return;
                        }

                        MaskingEngine.Instance.AddRuleWithPlaceholder(selected, placeholder);
                    }

                    ReplaceCurrentSelection(placeholder);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "マスキング登録");
            }
        }

        private void ReplaceCurrentSelection(string placeholder)
        {
            if (string.IsNullOrWhiteSpace(placeholder)) return;

            int start = BodyBox.SelectionStart;
            int length = BodyBox.SelectionLength;
            if (length <= 0) return;

            string current = BodyBox.Text ?? string.Empty;
            BodyBox.Text = current.Remove(start, length).Insert(start, placeholder);
            BodyBox.SelectionStart = start;
            BodyBox.SelectionLength = placeholder.Length;
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
