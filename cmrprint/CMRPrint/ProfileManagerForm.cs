using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace CMRPrint
{
    public sealed class ProfileManagerForm : Form
    {
        private const int SplitLeftMinWidth = 220;
        private const int SplitRightMinWidth = 650;
        private readonly List<ProfileRecord> _profiles;
        private readonly List<CmrField> _places;
        private readonly ListBox _listProfiles;
        private readonly TextBox _txtName;
        private readonly TextBox _txtCountry;
        private readonly TextBox _txtPlace;
        private readonly DataGridView _gridAssignments;
        private readonly TextBox _txtAssignmentValue;
        private readonly SplitContainer _splitMain;
        private bool _isSyncingAssignmentValue;

        public ProfileManagerForm(string title, IEnumerable<ProfileRecord> profiles, IEnumerable<CmrField> places)
        {
            Text = title;
            StartPosition = FormStartPosition.Manual;
            AutoScaleMode = AutoScaleMode.None;
            var workingArea = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1400, 900);
            Bounds = new Rectangle(
                workingArea.Left + 12,
                workingArea.Top + 12,
                Math.Max(900, workingArea.Width - 24),
                Math.Max(700, workingArea.Height - 24));
            MinimumSize = new Size(900, 700);

            _profiles = profiles
                .Select(CloneProfile)
                .OrderBy(profile => profile.Name)
                .ToList();
            _places = places.OrderBy(place => place.PlaceNumber).ToList();

            _splitMain = new SplitContainer
            {
                Dock = DockStyle.Fill,
                FixedPanel = FixedPanel.Panel1,
            };

            _listProfiles = new ListBox
            {
                Dock = DockStyle.Fill,
            };
            _listProfiles.SelectedIndexChanged += (_, _) => LoadSelectedProfile();

            var leftPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
            };

            var leftButtons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 64,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 10, 0, 12),
            };

            var btnNew = new Button
            {
                Text = "New",
                Size = new Size(70, 30),
            };
            btnNew.Click += (_, _) => CreateNewProfile();

            var btnDelete = new Button
            {
                Text = "Delete",
                Size = new Size(70, 30),
            };
            btnDelete.Click += (_, _) => DeleteSelectedProfile();

            leftButtons.Controls.Add(btnNew);
            leftButtons.Controls.Add(btnDelete);
            leftPanel.Controls.Add(_listProfiles);
            leftPanel.Controls.Add(leftButtons);
            _splitMain.Panel1.Controls.Add(leftPanel);

            var rightPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12, 12, 12, 18),
            };

            var editorTop = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 88,
                ColumnCount = 4,
                RowCount = 2,
            };
            editorTop.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90F));
            editorTop.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            editorTop.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90F));
            editorTop.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            editorTop.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));
            editorTop.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));

            var lblName = new Label { Text = "Name:", AutoSize = true, Anchor = AnchorStyles.Left };
            _txtName = new TextBox { Dock = DockStyle.Fill };

            var lblCountry = new Label { Text = "Country:", AutoSize = true, Anchor = AnchorStyles.Left };
            _txtCountry = new TextBox { Dock = DockStyle.Fill };

            var lblPlace = new Label { Text = "Place:", AutoSize = true, Anchor = AnchorStyles.Left };
            _txtPlace = new TextBox { Dock = DockStyle.Fill };

            editorTop.Controls.Add(lblName, 0, 0);
            editorTop.Controls.Add(_txtName, 1, 0);
            editorTop.SetColumnSpan(_txtName, 3);
            editorTop.Controls.Add(lblCountry, 0, 1);
            editorTop.Controls.Add(_txtCountry, 1, 1);
            editorTop.Controls.Add(lblPlace, 2, 1);
            editorTop.Controls.Add(_txtPlace, 3, 1);

            _gridAssignments = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                AutoGenerateColumns = false,
                RowHeadersVisible = false,
                RowTemplate = { Height = 54 },
            };
            _gridAssignments.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            _gridAssignments.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            var fieldColumn = new DataGridViewComboBoxColumn
            {
                HeaderText = "CMR Field",
                Width = 220,
                DataSource = _places,
                DisplayMember = "Description",
                ValueMember = "FieldName",
            };
            var valueColumn = new DataGridViewTextBoxColumn
            {
                HeaderText = "Value",
                Width = 430,
            };
            _gridAssignments.Columns.Add(fieldColumn);
            _gridAssignments.Columns.Add(valueColumn);

            var lblAssignmentValue = new Label
            {
                Text = "Selected value (multiline):",
                AutoSize = true,
                Dock = DockStyle.Top,
            };

            _txtAssignmentValue = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                AcceptsReturn = true,
                ScrollBars = ScrollBars.Vertical,
                WordWrap = true,
            };
            _txtAssignmentValue.TextChanged += (_, _) => SaveSelectedAssignmentValue();
            _gridAssignments.SelectionChanged += (_, _) => LoadSelectedAssignmentValue();
            _gridAssignments.CellValueChanged += (_, _) => LoadSelectedAssignmentValue();
            _gridAssignments.RowsAdded += (_, _) => LoadSelectedAssignmentValue();

            var btnSave = new Button
            {
                Text = "Save",
                Size = new Size(70, 30),
            };
            btnSave.Click += (_, _) => SaveSelectedProfile();

            var btnClose = new Button
            {
                Text = "Done",
                Size = new Size(70, 30),
                DialogResult = DialogResult.OK,
            };

            var bottomButtons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 68,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 12, 0, 14),
            };
            bottomButtons.Controls.Add(btnClose);
            bottomButtons.Controls.Add(btnSave);

            var assignmentEditorPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 170,
                Padding = new Padding(0, 8, 0, 0),
            };
            assignmentEditorPanel.Controls.Add(_txtAssignmentValue);
            assignmentEditorPanel.Controls.Add(lblAssignmentValue);

            var gridPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 8, 0, 8),
            };
            gridPanel.Controls.Add(_gridAssignments);

            rightPanel.Controls.Add(gridPanel);
            rightPanel.Controls.Add(assignmentEditorPanel);
            rightPanel.Controls.Add(bottomButtons);
            rightPanel.Controls.Add(editorTop);
            _splitMain.Panel2.Controls.Add(rightPanel);

            Controls.Add(_splitMain);
            Shown += (_, _) => ApplySplitLayout();
            Resize += (_, _) => ApplySplitLayout();

            PerformLayout();
            RefreshList();
        }

        private void ApplySplitLayout()
        {
            _splitMain.Panel1MinSize = SplitLeftMinWidth;
            _splitMain.Panel2MinSize = SplitRightMinWidth;

            var maxLeftWidth = Math.Max(SplitLeftMinWidth, _splitMain.Width - SplitRightMinWidth);
            if (maxLeftWidth <= SplitLeftMinWidth)
            {
                _splitMain.SplitterDistance = SplitLeftMinWidth;
                return;
            }

            _splitMain.SplitterDistance = Math.Min(260, maxLeftWidth);
        }

        public List<ProfileRecord> GetProfiles()
        {
            return _profiles
                .Where(profile => !string.IsNullOrWhiteSpace(profile.Name))
                .OrderBy(profile => profile.Name)
                .ToList();
        }

        private void RefreshList()
        {
            var selected = _listProfiles.SelectedItem as ProfileRecord;
            _listProfiles.Items.Clear();
            _listProfiles.Items.AddRange(_profiles.Cast<object>().ToArray());
            if (selected != null)
            {
                var match = _profiles.FirstOrDefault(profile => profile.Name == selected.Name);
                if (match != null)
                    _listProfiles.SelectedItem = match;
            }

            if (_listProfiles.SelectedIndex < 0 && _listProfiles.Items.Count > 0)
                _listProfiles.SelectedIndex = 0;
        }

        private void CreateNewProfile()
        {
            var profile = new ProfileRecord { Name = "New profile" };
            _profiles.Add(profile);
            RefreshList();
            _listProfiles.SelectedItem = profile;
        }

        private void DeleteSelectedProfile()
        {
            if (_listProfiles.SelectedItem is not ProfileRecord profile)
                return;

            _profiles.Remove(profile);
            RefreshList();
            if (_listProfiles.Items.Count == 0)
                ClearEditor();
        }

        private void LoadSelectedProfile()
        {
            if (_listProfiles.SelectedItem is not ProfileRecord profile)
            {
                ClearEditor();
                return;
            }

            _txtName.Text = profile.Name;
            _txtCountry.Text = profile.Country;
            _txtPlace.Text = profile.Place;

            _gridAssignments.Rows.Clear();
            foreach (var assignment in profile.FieldAssignments)
            {
                _gridAssignments.Rows.Add(assignment.FieldName, assignment.Value);
            }
            LoadSelectedAssignmentValue();
        }

        private void SaveSelectedProfile()
        {
            if (_listProfiles.SelectedItem is not ProfileRecord profile)
                return;

            profile.Name = _txtName.Text.Trim();
            profile.Country = _txtCountry.Text.Trim();
            profile.Place = _txtPlace.Text.Trim();
            profile.FieldAssignments = ReadAssignments();

            RefreshList();
            _listProfiles.SelectedItem = profile;
        }

        private List<FieldAssignment> ReadAssignments()
        {
            var result = new List<FieldAssignment>();
            foreach (DataGridViewRow row in _gridAssignments.Rows)
            {
                if (row.IsNewRow)
                    continue;

                var fieldName = row.Cells[0].Value?.ToString() ?? string.Empty;
                var value = row.Cells[1].Value?.ToString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(fieldName) || string.IsNullOrWhiteSpace(value))
                    continue;

                result.Add(new FieldAssignment { FieldName = fieldName, Value = value });
            }

            return result;
        }

        private void ClearEditor()
        {
            _txtName.Clear();
            _txtCountry.Clear();
            _txtPlace.Clear();
            _gridAssignments.Rows.Clear();
            _txtAssignmentValue.Clear();
        }

        private void LoadSelectedAssignmentValue()
        {
            if (_isSyncingAssignmentValue)
                return;

            if (_txtAssignmentValue == null)
                return;

            _isSyncingAssignmentValue = true;
            try
            {
                if (_gridAssignments.CurrentRow == null || _gridAssignments.CurrentRow.IsNewRow)
                {
                    _txtAssignmentValue.Clear();
                    return;
                }

                _txtAssignmentValue.Text = NormalizeMultilineText(_gridAssignments.CurrentRow.Cells[1].Value?.ToString() ?? string.Empty);
            }
            finally
            {
                _isSyncingAssignmentValue = false;
            }
        }

        private void SaveSelectedAssignmentValue()
        {
            if (_isSyncingAssignmentValue)
                return;

            if (_txtAssignmentValue == null)
                return;

            if (_gridAssignments.CurrentRow == null || _gridAssignments.CurrentRow.IsNewRow)
                return;

            _gridAssignments.CurrentRow.Cells[1].Value = NormalizeMultilineText(_txtAssignmentValue.Text);
        }

        private static string NormalizeMultilineText(string value)
        {
            return string.Join(
                Environment.NewLine,
                value.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'));
        }

        private static ProfileRecord CloneProfile(ProfileRecord profile)
        {
            return new ProfileRecord
            {
                Name = profile.Name,
                Country = profile.Country,
                Place = profile.Place,
                FieldAssignments = profile.FieldAssignments
                    .Select(assignment => new FieldAssignment { FieldName = assignment.FieldName, Value = assignment.Value })
                    .ToList()
            };
        }
    }
}
