using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace CMRPrint
{
    public sealed class BatchPrintSelection
    {
        public Customer Customer { get; init; } = new();
        public string Field9Value { get; init; } = string.Empty;
    }

    public sealed class BatchPrintForm : Form
    {
        private const string Field9StarterText = "xx Pal\r\nxx DC\r\nxx DCO\r\nxx DCS";
        private readonly List<Customer> _customers;
        private readonly TextBox _txtSearch;
        private readonly FlowLayoutPanel _countriesPanel;
        private readonly Dictionary<Customer, TextBox> _field9Inputs = new();
        private readonly Dictionary<Customer, CheckBox> _selectedChecks = new();

        public BatchPrintForm(IEnumerable<Customer> customers)
        {
            _customers = customers.OrderBy(c => c.Country).ThenBy(c => c.Name).ToList();

            Text = "Batch Print CMRs";
            StartPosition = FormStartPosition.CenterParent;
            Width = 980;
            Height = 720;
            MinimumSize = new Size(840, 620);

            var lblSearch = new Label
            {
                Text = "Search customer:",
                AutoSize = true,
                Location = new Point(16, 18),
            };

            _txtSearch = new TextBox
            {
                Location = new Point(140, 14),
                Width = 320,
            };
            _txtSearch.TextChanged += (_, _) => RenderCountries();

            var lblHint = new Label
            {
                Text = "Field 9 = nature of goods. Tick customers to add them to this print run.",
                AutoSize = true,
                Location = new Point(480, 18),
            };

            _countriesPanel = new FlowLayoutPanel
            {
                Location = new Point(16, 50),
                Size = new Size(930, 560),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
            };

            var btnPrint = new Button
            {
                Text = "Print Selected",
                DialogResult = DialogResult.OK,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Location = new Point(786, 625),
                Size = new Size(160, 34),
            };

            var btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Location = new Point(676, 625),
                Size = new Size(100, 34),
            };

            Controls.Add(lblSearch);
            Controls.Add(_txtSearch);
            Controls.Add(lblHint);
            Controls.Add(_countriesPanel);
            Controls.Add(btnPrint);
            Controls.Add(btnCancel);

            AcceptButton = btnPrint;
            CancelButton = btnCancel;

            RenderCountries();
        }

        public List<BatchPrintSelection> GetSelectedJobs()
        {
            return _selectedChecks
                .Where(entry => entry.Value.Checked)
                .Select(entry => new BatchPrintSelection
                {
                    Customer = entry.Key,
                    Field9Value = _field9Inputs.TryGetValue(entry.Key, out var input) ? input.Text.Trim() : string.Empty,
                })
                .ToList();
        }

        private void RenderCountries()
        {
            _countriesPanel.SuspendLayout();
            _countriesPanel.Controls.Clear();

            var search = _txtSearch.Text.Trim();
            var filtered = _customers.Where(customer =>
                string.IsNullOrWhiteSpace(search) ||
                customer.Name.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                customer.Address.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                customer.City.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                customer.Country.Contains(search, StringComparison.OrdinalIgnoreCase));

            foreach (var countryGroup in filtered.GroupBy(customer => string.IsNullOrWhiteSpace(customer.Country) ? "(No country)" : customer.Country))
            {
                var group = new GroupBox
                {
                    Text = countryGroup.Key,
                    Width = _countriesPanel.ClientSize.Width - 30,
                    Height = Math.Max(100, 44 + (countryGroup.Count() * 76)),
                    Padding = new Padding(10, 24, 10, 10),
                };

                var top = 24;
                foreach (var customer in countryGroup)
                {
                    var row = new Panel
                    {
                        Location = new Point(10, top),
                        Size = new Size(group.Width - 30, 72),
                        Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
                    };

                    var lblCustomer = new Label
                    {
                        Text = $"{customer.Name} | {customer.Address} | {customer.City}",
                        Location = new Point(0, 5),
                        Size = new Size(490, 20),
                        AutoEllipsis = true,
                    };

                    var lblField9 = new Label
                    {
                        Text = "Field 9:",
                        Location = new Point(500, 5),
                        Size = new Size(55, 20),
                    };

                    if (!_field9Inputs.TryGetValue(customer, out var txtField9))
                    {
                        txtField9 = new TextBox
                        {
                            Width = 120,
                            Height = 64,
                            Multiline = true,
                            ScrollBars = ScrollBars.Vertical,
                            AcceptsReturn = true,
                            Text = Field9StarterText
                        };
                        _field9Inputs[customer] = txtField9;
                    }

                    txtField9.Location = new Point(558, 2);
                    txtField9.Size = new Size(120, 64);

                    if (!_selectedChecks.TryGetValue(customer, out var chkSelected))
                    {
                        chkSelected = new CheckBox
                        {
                            Text = "Add",
                            AutoSize = true,
                        };
                        _selectedChecks[customer] = chkSelected;
                    }

                    chkSelected.Location = new Point(690, 24);

                    row.Controls.Add(lblCustomer);
                    row.Controls.Add(lblField9);
                    row.Controls.Add(txtField9);
                    row.Controls.Add(chkSelected);
                    group.Controls.Add(row);
                    top += 76;
                }

                _countriesPanel.Controls.Add(group);
            }

            _countriesPanel.ResumeLayout();
        }
    }
}
