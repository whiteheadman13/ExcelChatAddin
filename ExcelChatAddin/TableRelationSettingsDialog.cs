using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelChatAddin
{
    public class TableRelationSettingsDialog : Form
    {
        private readonly Excel.Application _excelApp;
        private TableSchemaStore _schemaStore;
        private TableRelationStore _relationStore;

        private DataGridView _gridTypeMaster;
        private DataGridView _gridTableRules;
        private DataGridView _gridRelations;

        public TableRelationSettingsDialog(Excel.Application app)
        {
            _excelApp = app;
            _schemaStore = IssueSchemaManager.LoadStore();
            _relationStore = TableRelationManager.LoadStore();

            InitializeLayout();
            LoadAllGrids();
        }

        private void InitializeLayout()
        {
            Text = "関係設定（MDM一覧編集）";
            Size = new Size(1300, 800);
            StartPosition = FormStartPosition.CenterParent;

            var top = new Panel { Dock = DockStyle.Top, Height = 78 };
            top.Controls.Add(new Label
            {
                Text = "メタ情報(JSON): " + Paths.TableRelationsPath,
                AutoSize = false,
                Width = 1260,
                Height = 24,
                Location = new Point(12, 10)
            });
            top.Controls.Add(new Label
            {
                Text = "実データ(Excel): シート『" + TableRelationSheetStore.SheetName + "』",
                AutoSize = false,
                Width = 1260,
                Height = 24,
                Location = new Point(12, 38)
            });

            var tabs = new TabControl { Dock = DockStyle.Fill };

            _gridTypeMaster = CreateTypeMasterGrid();
            _gridTableRules = CreateTableRuleGrid();
            _gridRelations = CreateRelationsGrid();

            var tabType = new TabPage("関係種別マスタ");
            tabType.Controls.Add(_gridTypeMaster);
            tabs.TabPages.Add(tabType);

            var tabRule = new TabPage("テーブル間関係ルール");
            tabRule.Controls.Add(_gridTableRules);
            tabs.TabPages.Add(tabRule);

            var tabRelations = new TabPage("関係一覧");
            tabRelations.Controls.Add(_gridRelations);
            tabs.TabPages.Add(tabRelations);

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 52 };

            var btnPasteTsv = new Button { Text = "TSV貼付", Width = 110, Height = 30, Location = new Point(12, 10) };
            var btnDeleteRows = new Button { Text = "選択行削除", Width = 110, Height = 30, Location = new Point(132, 10) };
            var btnValidate = new Button { Text = "検証", Width = 110, Height = 30, Location = new Point(252, 10) };
            var btnMatrix = new Button { Text = "マトリクス表示", Width = 120, Height = 30, Location = new Point(372, 10) };
            var btnSave = new Button { Text = "保存", Width = 110, Height = 30, Location = new Point(900, 10), Font = new Font(DefaultFont, FontStyle.Bold) };
            var btnClose = new Button { Text = "閉じる", Width = 110, Height = 30, Location = new Point(1020, 10) };

            btnPasteTsv.Click += (s, e) => PasteTsvToRelations();
            btnDeleteRows.Click += (s, e) => DeleteSelectedRows(tabs.SelectedTab);
            btnValidate.Click += (s, e) => ValidateOnly();
            btnMatrix.Click += (s, e) =>
            {
                try
                {
                    using (var f = new TableRelationMatrixForm(_excelApp))
                    {
                        f.ShowDialog(this);
                        // マトリクスで変更があった場合は一覧を再読込
                        _relationStore = TableRelationManager.LoadStore();
                        LoadAllGrids();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("マトリクス表示に失敗しました: " + ex.Message, "関係設定");
                }
            };
            btnSave.Click += (s, e) => SaveAll();
            btnClose.Click += (s, e) => Close();

            bottom.Controls.Add(btnPasteTsv);
            bottom.Controls.Add(btnDeleteRows);
            bottom.Controls.Add(btnValidate);
            bottom.Controls.Add(btnMatrix);
            bottom.Controls.Add(btnSave);
            bottom.Controls.Add(btnClose);

            Controls.Add(tabs);
            Controls.Add(top);
            Controls.Add(bottom);
        }

        private DataGridView CreateTypeMasterGrid()
        {
            var grid = CreateBaseGrid();
            grid.Columns.Add("RelationTypeCode", "関係種別コード");
            grid.Columns.Add("RelationTypeName", "関係種別名");
            grid.Columns.Add("Description", "説明");
            grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "IsEnabled", HeaderText = "使用可否" });
            grid.Columns.Add("SortOrder", "表示順");
            return grid;
        }

        private DataGridView CreateTableRuleGrid()
        {
            var grid = CreateBaseGrid();
            grid.Columns.Add(new DataGridViewComboBoxColumn { Name = "FromTableName", HeaderText = "元テーブル" });
            grid.Columns.Add(new DataGridViewComboBoxColumn { Name = "ToTableName", HeaderText = "先テーブル" });
            grid.Columns.Add(new DataGridViewComboBoxColumn { Name = "RelationTypeCode", HeaderText = "関係種別コード" });
            grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "IsAllowed", HeaderText = "関係許可" });
            grid.Columns.Add("Notes", "備考");
            return grid;
        }

        private DataGridView CreateRelationsGrid()
        {
            var grid = CreateBaseGrid();
            grid.Columns.Add(new DataGridViewComboBoxColumn { Name = "FromTableName", HeaderText = "元テーブル" });
            grid.Columns.Add("FromKey", "元キー");
            grid.Columns.Add(new DataGridViewComboBoxColumn { Name = "ToTableName", HeaderText = "先テーブル" });
            grid.Columns.Add("ToKey", "先キー");
            grid.Columns.Add(new DataGridViewComboBoxColumn { Name = "RelationTypeCode", HeaderText = "関係種別コード" });
            grid.Columns.Add("Meaning", "意味");
            grid.Columns.Add("Notes", "備考");
            grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "IsDecoupling", HeaderText = "疎結合化" });
            grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "IsEnabled", HeaderText = "有効" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "UpdatedAtUtc", HeaderText = "更新日時(UTC)", ReadOnly = true });
            return grid;
        }

        private static DataGridView CreateBaseGrid()
        {
            return new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true
            };
        }

        private void LoadAllGrids()
        {
            TableRelationManager.Normalize(_relationStore);
            ApplyComboSources();

            _gridTypeMaster.Rows.Clear();
            foreach (var t in _relationStore.RelationTypes)
            {
                _gridTypeMaster.Rows.Add(t.RelationTypeCode, t.RelationTypeName, t.Description, t.IsEnabled, t.SortOrder);
            }

            _gridTableRules.Rows.Clear();
            foreach (var r in _relationStore.TableRules)
            {
                _gridTableRules.Rows.Add(r.FromTableName, r.ToTableName, r.RelationTypeCode, r.IsAllowed, r.Notes);
            }

            var relations = TableRelationSheetStore.LoadRelations(_excelApp);
            _gridRelations.Rows.Clear();
            foreach (var r in relations.OrderByDescending(x => x.UpdatedAtUtc))
            {
                _gridRelations.Rows.Add(
                    r.FromTableName,
                    r.FromKey,
                    r.ToTableName,
                    r.ToKey,
                    r.RelationTypeCode,
                    r.Meaning,
                    r.Notes,
                    r.IsDecoupling,
                    r.IsEnabled,
                    r.UpdatedAtUtc.ToString("yyyy-MM-dd HH:mm:ss"));
            }
        }

        private void ApplyComboSources()
        {
            var tables = (_schemaStore?.Tables ?? new List<IssueSchemaConfig>())
                .Select(x => (x?.TableName ?? "").Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToArray();

            var relationCodes = (_relationStore?.RelationTypes ?? new List<RelationTypeMasterItem>())
                .Where(x => x.IsEnabled)
                .OrderBy(x => x.SortOrder)
                .ThenBy(x => x.RelationTypeCode)
                .Select(x => x.RelationTypeCode)
                .ToArray();

            ApplyComboDataSource(_gridTableRules, "FromTableName", tables);
            ApplyComboDataSource(_gridTableRules, "ToTableName", tables);
            ApplyComboDataSource(_gridTableRules, "RelationTypeCode", relationCodes);

            ApplyComboDataSource(_gridRelations, "FromTableName", tables);
            ApplyComboDataSource(_gridRelations, "ToTableName", tables);
            ApplyComboDataSource(_gridRelations, "RelationTypeCode", relationCodes);
        }

        private static void ApplyComboDataSource(DataGridView grid, string columnName, string[] data)
        {
            if (!(grid.Columns[columnName] is DataGridViewComboBoxColumn col)) return;
            col.DataSource = null;
            col.DataSource = data ?? Array.Empty<string>();
        }

        private void PasteTsvToRelations()
        {
            try
            {
                var text = Clipboard.GetText();
                var rows = TableRelationManager.ParseTsvForRelations(text);
                if (rows.Count == 0)
                {
                    MessageBox.Show("TSV形式のデータが見つかりません。\n列順: 元テーブル\t元キー\t先テーブル\t先キー\t関係種別コード\t意味\t備考\t疎結合化\t有効", "TSV貼付");
                    return;
                }

                foreach (var r in rows)
                {
                    _gridRelations.Rows.Add(
                        r.FromTableName,
                        r.FromKey,
                        r.ToTableName,
                        r.ToKey,
                        r.RelationTypeCode,
                        r.Meaning,
                        r.Notes,
                        r.IsDecoupling,
                        r.IsEnabled,
                        DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss"));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("TSV貼付に失敗しました: " + ex.Message, "TSV貼付");
            }
        }

        private void DeleteSelectedRows(TabPage selectedTab)
        {
            if (selectedTab == null) return;

            DataGridView target = null;
            if (selectedTab.Text == "関係種別マスタ") target = _gridTypeMaster;
            else if (selectedTab.Text == "テーブル間関係ルール") target = _gridTableRules;
            else if (selectedTab.Text == "関係一覧") target = _gridRelations;
            if (target == null) return;

            foreach (DataGridViewRow row in target.SelectedRows.Cast<DataGridViewRow>().ToList())
            {
                if (row.IsNewRow) continue;
                target.Rows.Remove(row);
            }
        }

        private void ValidateOnly()
        {
            var errors = ValidateAll();
            if (errors.Count == 0)
            {
                MessageBox.Show("検証OKです。", "関係設定");
                return;
            }

            MessageBox.Show("検証エラー:\n- " + string.Join("\n- ", errors.Take(20))
                + (errors.Count > 20 ? "\n..." : ""), "関係設定", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void SaveAll()
        {
            try
            {
                _gridTypeMaster.EndEdit();
                _gridTableRules.EndEdit();
                _gridRelations.EndEdit();

                var store = BuildStoreFromGrid();
                var relations = CollectRelationsFromGrid();
                var errors = ValidateAll(store, relations);
                if (errors.Count > 0)
                {
                    MessageBox.Show("保存できません。\n- " + string.Join("\n- ", errors.Take(20))
                        + (errors.Count > 20 ? "\n..." : ""), "関係設定", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                TableRelationManager.SaveStore(store);
                TableRelationSheetStore.SaveRelations(_excelApp, relations);

                _relationStore = TableRelationManager.LoadStore();
                LoadAllGrids();

                MessageBox.Show("関係定義を保存しました。", "関係設定");
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存に失敗しました: " + ex.Message, "関係設定");
            }
        }

        private List<string> ValidateAll(TableRelationStore targetStore = null, List<TableRecordRelation> targetRelations = null)
        {
            var store = targetStore ?? BuildStoreFromGrid();
            var relations = targetRelations ?? CollectRelationsFromGrid();
            var tableNames = (_schemaStore?.Tables ?? new List<IssueSchemaConfig>())
                .Select(x => x.TableName)
                .Where(x => !string.IsNullOrWhiteSpace(x));
            var tableKeys = BuildTableKeySet();

            return TableRelationManager.Validate(store, relations, tableNames, tableKeys);
        }

        private TableRelationStore BuildStoreFromGrid()
        {
            var store = new TableRelationStore();

            foreach (DataGridViewRow row in _gridTypeMaster.Rows)
            {
                if (row.IsNewRow) continue;
                var code = ReadString(row, "RelationTypeCode");
                var name = ReadString(row, "RelationTypeName");
                if (string.IsNullOrWhiteSpace(code) && string.IsNullOrWhiteSpace(name)) continue;

                store.RelationTypes.Add(new RelationTypeMasterItem
                {
                    RelationTypeCode = TableRelationManager.NormalizeCode(code),
                    RelationTypeName = name,
                    Description = ReadString(row, "Description"),
                    IsEnabled = ReadBool(row, "IsEnabled", true),
                    SortOrder = ReadInt(row, "SortOrder", 0)
                });
            }

            foreach (DataGridViewRow row in _gridTableRules.Rows)
            {
                if (row.IsNewRow) continue;
                if (IsAllEmpty(row, "FromTableName", "ToTableName", "RelationTypeCode", "Notes")) continue;

                store.TableRules.Add(new TableRelationRule
                {
                    FromTableName = ReadString(row, "FromTableName"),
                    ToTableName = ReadString(row, "ToTableName"),
                    RelationTypeCode = TableRelationManager.NormalizeCode(ReadString(row, "RelationTypeCode")),
                    IsAllowed = ReadBool(row, "IsAllowed", true),
                    Notes = ReadString(row, "Notes")
                });
            }

            TableRelationManager.Normalize(store);
            store.Relations = new List<TableRecordRelation>();
            return store;
        }

        private List<TableRecordRelation> CollectRelationsFromGrid()
        {
            var relations = new List<TableRecordRelation>();
            foreach (DataGridViewRow row in _gridRelations.Rows)
            {
                if (row.IsNewRow) continue;
                if (IsAllEmpty(row, "FromTableName", "FromKey", "ToTableName", "ToKey", "RelationTypeCode", "Meaning", "Notes")) continue;

                relations.Add(new TableRecordRelation
                {
                    FromTableName = ReadString(row, "FromTableName"),
                    FromKey = ReadString(row, "FromKey"),
                    ToTableName = ReadString(row, "ToTableName"),
                    ToKey = ReadString(row, "ToKey"),
                    RelationTypeCode = TableRelationManager.NormalizeCode(ReadString(row, "RelationTypeCode")),
                    Meaning = ReadString(row, "Meaning"),
                    Notes = ReadString(row, "Notes"),
                    IsDecoupling = ReadBool(row, "IsDecoupling", false),
                    IsEnabled = ReadBool(row, "IsEnabled", true),
                    UpdatedAtUtc = DateTime.UtcNow
                });
            }

            return relations;
        }

        private Dictionary<string, HashSet<string>> BuildTableKeySet()
        {
            var result = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var wb = _excelApp?.ActiveWorkbook;
                if (wb == null) return result;

                var schemaByTable = (_schemaStore?.Tables ?? new List<IssueSchemaConfig>())
                    .Where(x => x != null && !string.IsNullOrWhiteSpace(x.TableName))
                    .ToDictionary(x => x.TableName, StringComparer.OrdinalIgnoreCase);

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    if (ws?.ListObjects == null) continue;
                    foreach (Excel.ListObject lo in ws.ListObjects)
                    {
                        var tableName = lo?.Name;
                        if (string.IsNullOrWhiteSpace(tableName)) continue;
                        if (!schemaByTable.TryGetValue(tableName, out var schema)) continue;

                        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        result[tableName] = keys;

                        if (lo.DataBodyRange == null) continue;

                        var absoluteKeyCol = ColumnLetterToIndex(schema.KeyColumnLetter);
                        if (absoluteKeyCol <= 0) continue;

                        var relativeKeyCol = absoluteKeyCol - lo.Range.Column + 1;
                        if (relativeKeyCol <= 0 || relativeKeyCol > lo.Range.Columns.Count) continue;

                        for (int r = 1; r <= lo.DataBodyRange.Rows.Count; r++)
                        {
                            var cell = lo.DataBodyRange.Cells[r, relativeKeyCol] as Excel.Range;
                            var key = (Convert.ToString(cell?.Value2) ?? "").Trim();
                            if (!string.IsNullOrWhiteSpace(key)) keys.Add(key);
                        }
                    }
                }
            }
            catch
            {
            }

            return result;
        }

        private static int ColumnLetterToIndex(string letter)
        {
            if (string.IsNullOrWhiteSpace(letter)) return 0;
            int index = 0;
            foreach (var ch in letter.Trim().ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z') return 0;
                index = index * 26 + (ch - 'A' + 1);
            }
            return index;
        }

        private static bool IsAllEmpty(DataGridViewRow row, params string[] columnNames)
        {
            foreach (var name in columnNames)
            {
                if (!string.IsNullOrWhiteSpace(ReadString(row, name))) return false;
            }
            return true;
        }

        private static string ReadString(DataGridViewRow row, string columnName)
        {
            return (row.Cells[columnName].Value?.ToString() ?? "").Trim();
        }

        private static bool ReadBool(DataGridViewRow row, string columnName, bool defaultValue)
        {
            try
            {
                var value = row.Cells[columnName].Value;
                if (value == null) return defaultValue;
                return Convert.ToBoolean(value);
            }
            catch
            {
                return defaultValue;
            }
        }

        private static int ReadInt(DataGridViewRow row, string columnName, int defaultValue)
        {
            var text = ReadString(row, columnName);
            if (int.TryParse(text, out var value)) return value;
            return defaultValue;
        }
    }
}
