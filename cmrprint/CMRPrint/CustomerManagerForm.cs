using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace CMRPrint
{
    public sealed class CustomerManagerForm : Form
    {
        private const int SplitLeftMinWidth = 220;
        private const int SplitRightMinWidth = 650;
        private readonly List<Customer> _customers;
        private readonly List<ProfileRecord> _exporters;
        private readonly List<ProfileRecord> _transportInfos;
        private readonly List<ProfileRecord> _loadingPlaces;
        private readonly List<CmrField> _places;
        private readonly ListBox _listCustomers;
        private readonly TextBox _txtName;
        private readonly TextBox _txtAddress;
        private readonly TextBox _txtCity;
        private readonly TextBox _txtCountry;
        private readonly TextBox _txtVat;
        private readonly TextBox _txtPlaceOfIssue;
        private readonly ComboBox _cmbExporter;
        private readonly ComboBox _cmbTransport;
        private readonly ComboBox _cmbLoadingPlace;
        private readonly DataGridView _gridAssignments;
        private readonly TextBox _txtAssignmentValue;
        private readonly SplitContainer _splitMain;
        private bool _isSyncingAssignmentValue;

        public CustomerManagerForm(
            IEnumerable<Customer> customers,
            IEnumerable<ProfileRecord> exporters,
            IEnumerable<ProfileRecord> transportInfos,
            IEnumerable<ProfileRecord> loadingPlaces,
            IEnumerable<CmrField> places)
        {
            Text = "Customer Info";
            StartPosition = FormStartPosition.Manual;
            AutoScaleMode = AutoScaleMode.None;
            var workingArea = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1400, 900);
            Bounds = new Rectangle(
                workingArea.Left + 12,
                workingArea.Top + 12,
                Math.Max(900, workingArea.Width - 24),
                Math.Max(700, workingArea.Height - 24));
            MinimumSize = new Size(900, 700);

            _customers = customers.Select(CloneCustomer).OrderBy(customer => customer.Name).ToList();
            _exporters = exporters.OrderBy(profile => profile.Name).ToList();
            _transportInfos = transportInfos.OrderBy(profile => profile.Name).ToList();
            _loadingPlaces = loadingPlaces.OrderBy(profile => profile.Name).ToList();
            _places = places.OrderBy(place => place.PlaceNumber).ToList();

            _splitMain = new SplitContainer
            {
                Dock = DockStyle.Fill,
                FixedPanel = FixedPanel.Panel1,
            };

            _listCustomers = new ListBox
            {
                Dock = DockStyle.Fill,
            };
            _listCustomers.SelectedIndexChanged += (_, _) => LoadSelectedCustomer();

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
            btnNew.Click += (_, _) => CreateNewCustomer();

            var btnDelete = new Button
            {
                Text = "Delete",
                Size = new Size(70, 30),
            };
            btnDelete.Click += (_, _) => DeleteSelectedCustomer();

            leftButtons.Controls.Add(btnNew);
            leftButtons.Controls.Add(btnDelete);
            leftPanel.Controls.Add(_listCustomers);
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
                Height = 200,
                ColumnCount = 4,
                RowCount = 5,
                Padding = new Padding(0),
            };
            editorTop.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90F));
            editorTop.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            editorTop.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 95F));
            editorTop.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            editorTop.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            editorTop.RowStyles.Add(new RowStyle(SizeType.Absolute, 64F));
            editorTop.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            editorTop.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            editorTop.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));

            var lblName = new Label { Text = "Name:", AutoSize = true, Anchor = AnchorStyles.Left };
            _txtName = new TextBox { Dock = DockStyle.Fill };
            var lblAddress = new Label { Text = "Address:", AutoSize = true, Anchor = AnchorStyles.Left };
            _txtAddress = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical };
            var lblCity = new Label { Text = "City:", AutoSize = true, Anchor = AnchorStyles.Left };
            _txtCity = new TextBox { Dock = DockStyle.Fill };
            var lblCountry = new Label { Text = "Country:", AutoSize = true, Anchor = AnchorStyles.Left };
            _txtCountry = new TextBox { Dock = DockStyle.Fill };
            var lblVat = new Label { Text = "VAT:", AutoSize = true, Anchor = AnchorStyles.Left };
            _txtVat = new TextBox { Dock = DockStyle.Fill };
            var lblPlace = new Label { Text = "Place/date:", AutoSize = true, Anchor = AnchorStyles.Left };
            _txtPlaceOfIssue = new TextBox { Dock = DockStyle.Fill };

            var lblExporter = new Label { Text = "Exporter:", AutoSize = true, Anchor = AnchorStyles.Left };
            _cmbExporter = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbExporter.Items.Add(string.Empty);
            _cmbExporter.Items.AddRange(_exporters.Select(profile => profile.Name).Cast<object>().ToArray());

            var lblTransport = new Label { Text = "Transport:", AutoSize = true, Anchor = AnchorStyles.Left };
            _cmbTransport = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbTransport.Items.Add(string.Empty);
            _cmbTransport.Items.AddRange(_transportInfos.Select(profile => profile.Name).Cast<object>().ToArray());

            var lblLoading = new Label { Text = "Loading place:", AutoSize = true, Anchor = AnchorStyles.Left };
            _cmbLoadingPlace = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbLoadingPlace.Items.Add(string.Empty);
            _cmbLoadingPlace.Items.AddRange(_loadingPlaces.Select(profile => profile.Name).Cast<object>().ToArray());

            var lblHint = new Label
            {
                Text = "Assign customer-specific text to CMR fields here. Manual fields stay free: 7, 9, 17.",
                AutoSize = true,
                Dock = DockStyle.Top,
                Padding = new Padding(0, 8, 0, 8),
            };

            editorTop.Controls.Add(lblName, 0, 0);
            editorTop.Controls.Add(_txtName, 1, 0);
            editorTop.SetColumnSpan(_txtName, 3);
            editorTop.Controls.Add(lblAddress, 0, 1);
            editorTop.Controls.Add(_txtAddress, 1, 1);
            editorTop.SetColumnSpan(_txtAddress, 3);
            editorTop.Controls.Add(lblCity, 0, 2);
            editorTop.Controls.Add(_txtCity, 1, 2);

            var countryVatPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
            };
            countryVatPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
            countryVatPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            countryVatPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 45F));
            countryVatPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            countryVatPanel.Controls.Add(lblCountry, 0, 0);
            countryVatPanel.Controls.Add(_txtCountry, 1, 0);
            countryVatPanel.Controls.Add(lblVat, 2, 0);
            countryVatPanel.Controls.Add(_txtVat, 3, 0);
            editorTop.Controls.Add(countryVatPanel, 2, 2);
            editorTop.SetColumnSpan(countryVatPanel, 2);

            editorTop.Controls.Add(lblPlace, 0, 3);
            editorTop.Controls.Add(_txtPlaceOfIssue, 1, 3);
            editorTop.Controls.Add(lblExporter, 0, 4);
            editorTop.Controls.Add(_cmbExporter, 1, 4);
            editorTop.Controls.Add(lblTransport, 2, 4);
            editorTop.Controls.Add(_cmbTransport, 3, 4);

            var loadingRow = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 40,
                ColumnCount = 2,
                Padding = new Padding(0, 0, 0, 6),
            };
            loadingRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90F));
            loadingRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            loadingRow.Controls.Add(lblLoading, 0, 0);
            loadingRow.Controls.Add(_cmbLoadingPlace, 1, 0);

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
                Width = 250,
                DataSource = _places,
                DisplayMember = "Description",
                ValueMember = "FieldName",
            };
            var valueColumn = new DataGridViewTextBoxColumn
            {
                HeaderText = "Value",
                Width = 590,
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
            btnSave.Click += (_, _) => SaveSelectedCustomer();

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
                Height = 180,
                Padding = new Padding(0, 8, 0, 0),
            };
            assignmentEditorPanel.Controls.Add(_txtAssignmentValue);
            assignmentEditorPanel.Controls.Add(lblAssignmentValue);

            var gridPanel = new Panel
            {
                Dock = DockStyle.Fill,
            };
            gridPanel.Controls.Add(_gridAssignments);

            rightPanel.Controls.Add(gridPanel);
            rightPanel.Controls.Add(assignmentEditorPanel);
            rightPanel.Controls.Add(bottomButtons);
            rightPanel.Controls.Add(lblHint);
            rightPanel.Controls.Add(loadingRow);
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

            _splitMain.SplitterDistance = Math.Min(250, maxLeftWidth);
        }

        public List<Customer> GetCustomers()
        {
            return _customers
                .Where(customer => !string.IsNullOrWhiteSpace(customer.Name))
                .OrderBy(customer => customer.Name)
                .ToList();
        }

        private void RefreshList()
        {
            var selected = _listCustomers.SelectedItem as Customer;
            _listCustomers.Items.Clear();
            _listCustomers.Items.AddRange(_customers.Cast<object>().ToArray());

            if (selected != null)
            {
                var match = _customers.FirstOrDefault(customer => customer.Name == selected.Name && customer.Address == selected.Address);
                if (match != null)
                    _listCustomers.SelectedItem = match;
            }

            if (_listCustomers.SelectedIndex < 0 && _listCustomers.Items.Count > 0)
                _listCustomers.SelectedIndex = 0;
        }

        private void CreateNewCustomer()
        {
            var customer = new Customer { Name = "New customer" };
            _customers.Add(customer);
            RefreshList();
            _listCustomers.SelectedItem = customer;
        }

        private void DeleteSelectedCustomer()
        {
            if (_listCustomers.SelectedItem is not Customer customer)
                return;

            _customers.Remove(customer);
            RefreshList();
            if (_listCustomers.Items.Count == 0)
                ClearEditor();
        }

        private void LoadSelectedCustomer()
        {
            if (_listCustomers.SelectedItem is not Customer customer)
            {
                ClearEditor();
                return;
            }

            _txtName.Text = customer.Name;
            _txtAddress.Text = customer.Address;
            _txtCity.Text = customer.City;
            _txtCountry.Text = customer.Country;
            _txtVat.Text = customer.VatNumber;
            _txtPlaceOfIssue.Text = customer.PlaceOfIssue;
            _cmbExporter.SelectedItem = string.IsNullOrWhiteSpace(customer.ExporterProfileName) ? string.Empty : customer.ExporterProfileName;
            _cmbTransport.SelectedItem = string.IsNullOrWhiteSpace(customer.TransportProfileName) ? string.Empty : customer.TransportProfileName;
            _cmbLoadingPlace.SelectedItem = string.IsNullOrWhiteSpace(customer.LoadingPlaceProfileName) ? string.Empty : customer.LoadingPlaceProfileName;

            _gridAssignments.Rows.Clear();
            foreach (var assignment in customer.FieldAssignments)
            {
                _gridAssignments.Rows.Add(assignment.FieldName, assignment.Value);
            }
            LoadSelectedAssignmentValue();
        }

        private void SaveSelectedCustomer()
        {
            if (_listCustomers.SelectedItem is not Customer customer)
                return;

            customer.Name = _txtName.Text.Trim();
            customer.Address = _txtAddress.Text.Trim();
            customer.City = _txtCity.Text.Trim();
            customer.Country = _txtCountry.Text.Trim();
            customer.VatNumber = _txtVat.Text.Trim();
            customer.PlaceOfIssue = _txtPlaceOfIssue.Text.Trim();
            customer.ExporterProfileName = _cmbExporter.SelectedItem?.ToString() ?? string.Empty;
            customer.TransportProfileName = _cmbTransport.SelectedItem?.ToString() ?? string.Empty;
            customer.LoadingPlaceProfileName = _cmbLoadingPlace.SelectedItem?.ToString() ?? string.Empty;
            customer.FieldAssignments = ReadAssignments();

            RefreshList();
            _listCustomers.SelectedItem = customer;
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
            _txtAddress.Clear();
            _txtCity.Clear();
            _txtCountry.Clear();
            _txtVat.Clear();
            _txtPlaceOfIssue.Clear();
            _cmbExporter.SelectedIndex = 0;
            _cmbTransport.SelectedIndex = 0;
            _cmbLoadingPlace.SelectedIndex = 0;
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

        private static Customer CloneCustomer(Customer customer)
        {
            return new Customer
            {
                Name = customer.Name,
                Address = customer.Address,
                City = customer.City,
                Country = customer.Country,
                VatNumber = customer.VatNumber,
                ExporterProfileName = customer.ExporterProfileName,
                TransportProfileName = customer.TransportProfileName,
                LoadingPlaceProfileName = customer.LoadingPlaceProfileName,
                PlaceOfIssue = customer.PlaceOfIssue,
                FieldAssignments = customer.FieldAssignments
                    .Select(assignment => new FieldAssignment { FieldName = assignment.FieldName, Value = assignment.Value })
                    .ToList()
            };
        }
    }
}
