using System;
using System.Drawing;
using System.Windows.Forms;

namespace CMRPrint
{
    public class CmrFieldControl : UserControl
    {
        private Label lblDescription = null!;
        private TextBox txtValue = null!;
        private NumericUpDown nudFontSize = null!;
        private NumericUpDown nudOffset = null!;
        private NumericUpDown nudFieldWidth = null!;
        private NumericUpDown nudFieldHeight = null!;
        private Label lblFontSize = null!;
        private Label lblOffset = null!;
        private Label lblFieldWidth = null!;
        private Label lblFieldHeight = null!;

        public CmrField Place { get; private set; }
        private readonly CmrTemplate _template;
        public event EventHandler? LayoutSettingsChanged;

        public CmrFieldControl(CmrField place, CmrTemplate template)
        {
            Place = place;
            _template = template;
            InitializeComponent();
            LoadTemplateValues();
        }

        private void InitializeComponent()
        {
            lblDescription = new Label
            {
                Text = Place.Description,
                AutoSize = false,
                Location = new System.Drawing.Point(8, 6),
                Size = new System.Drawing.Size(250, 20),
                Font = new System.Drawing.Font("Arial", 8, System.Drawing.FontStyle.Bold),
            };

            txtValue = new TextBox
            {
                Location = new System.Drawing.Point(8, 30),
                Size = new System.Drawing.Size(250, 98),
                Multiline = true,
                AcceptsReturn = true,
                ScrollBars = ScrollBars.Vertical,
                WordWrap = true,
            };

            lblFontSize = new Label
            {
                Text = "Font:",
                AutoSize = true,
                Location = new System.Drawing.Point(280, 10),
            };

            nudFontSize = new NumericUpDown
            {
                Location = new System.Drawing.Point(324, 8),
                Size = new System.Drawing.Size(64, 23),
                Minimum = 6,
                Maximum = 20,
                Value = Place.DefaultFontSize,
            };

            lblOffset = new Label
            {
                Text = "Offset:",
                AutoSize = true,
                Location = new System.Drawing.Point(402, 10),
            };

            nudOffset = new NumericUpDown
            {
                Location = new System.Drawing.Point(454, 8),
                Size = new System.Drawing.Size(64, 23),
                Minimum = -100,
                Maximum = 100,
                Value = 0,
            };

            lblFieldWidth = new Label
            {
                Text = "W:",
                AutoSize = true,
                Location = new System.Drawing.Point(280, 40),
            };

            nudFieldWidth = new NumericUpDown
            {
                Location = new System.Drawing.Point(302, 38),
                Size = new System.Drawing.Size(86, 23),
                Minimum = 30,
                Maximum = 400,
                Value = 140,
            };

            lblFieldHeight = new Label
            {
                Text = "H:",
                AutoSize = true,
                Location = new System.Drawing.Point(402, 40),
            };

            nudFieldHeight = new NumericUpDown
            {
                Location = new System.Drawing.Point(424, 38),
                Size = new System.Drawing.Size(94, 23),
                Minimum = 18,
                Maximum = 220,
                Value = 52,
            };

            Controls.Add(lblDescription);
            Controls.Add(txtValue);
            Controls.Add(lblFontSize);
            Controls.Add(nudFontSize);
            Controls.Add(lblOffset);
            Controls.Add(nudOffset);
            Controls.Add(lblFieldWidth);
            Controls.Add(nudFieldWidth);
            Controls.Add(lblFieldHeight);
            Controls.Add(nudFieldHeight);

            txtValue.BringToFront();
            nudFontSize.ValueChanged += (_, _) => LayoutSettingsChanged?.Invoke(this, EventArgs.Empty);
            nudOffset.ValueChanged += (_, _) => LayoutSettingsChanged?.Invoke(this, EventArgs.Empty);
            nudFieldWidth.ValueChanged += (_, _) => LayoutSettingsChanged?.Invoke(this, EventArgs.Empty);
            nudFieldHeight.ValueChanged += (_, _) => LayoutSettingsChanged?.Invoke(this, EventArgs.Empty);

            BorderStyle = BorderStyle.FixedSingle;
            BackColor = SystemColors.Control;
            Height = 138;
            Width = 530;
        }

        private void LoadTemplateValues()
        {
            if (_template.FontSizes.TryGetValue(Place.FieldName, out var fontSize))
            {
                nudFontSize.Value = fontSize;
            }

            if (_template.VerticalOffsets.TryGetValue(Place.FieldName, out var offset))
            {
                nudOffset.Value = offset;
            }

            if (_template.FieldWidths.TryGetValue(Place.FieldName, out var fieldWidth))
            {
                nudFieldWidth.Value = fieldWidth;
            }

            if (_template.FieldHeights.TryGetValue(Place.FieldName, out var fieldHeight))
            {
                nudFieldHeight.Value = fieldHeight;
            }
        }

        public string GetValue()
        {
            return txtValue.Text;
        }

        public void SetValue(string value)
        {
            txtValue.Text = value;
        }

        public int GetFontSize()
        {
            return (int)nudFontSize.Value;
        }

        public int GetVerticalOffset()
        {
            return (int)nudOffset.Value;
        }

        public int GetFieldWidth()
        {
            return (int)nudFieldWidth.Value;
        }

        public int GetFieldHeight()
        {
            return (int)nudFieldHeight.Value;
        }
    }
}
