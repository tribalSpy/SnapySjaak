using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Windows.Forms;

namespace CMRPrint
{
    public partial class MainForm : Form
    {
        private sealed class BatchPrintJob
        {
            public Customer Customer { get; init; } = new();
            public string Field9Value { get; init; } = string.Empty;
        }

        private const int MenuFieldWidth = 530;
        private const int MenuFieldHeight = 138;
        private const int MenuFieldSpacing = 10;
        private const float PreviewScale = 0.90f;
        private const int PreviewPadding = 18;
        private const int DefaultDocumentFieldWidth = 140;
        private const int DefaultDocumentFieldHeight = 52;
        private const short PrintCopies = 4;
        private const int A4WidthHundredthsInch = 827;
        private const int A4HeightHundredthsInch = 1169;
        private const string LegacyField9FieldName = "GrossWeight";
        private const string RenamedField9FieldName = "NatureofGoods";
        private const string Field9StarterText = "xx Pal\r\nxx DC\r\nxx DCO\r\nxx DCS";

        private AppDataStore _appData = new();
        private readonly List<Customer> _customers = new();
        private readonly List<ProfileRecord> _exporters = new();
        private readonly List<ProfileRecord> _transportInfos = new();
        private readonly List<ProfileRecord> _loadingPlaces = new();
        private readonly Dictionary<string, CmrFieldControl> _fieldControls = new();
        private readonly Dictionary<string, Panel> _previewFieldControls = new();
        private readonly Dictionary<string, PointF> _fieldPositions = new();
        private readonly Dictionary<string, Size> _fieldSizes = new();
        private readonly Dictionary<string, string> _documentFieldValues = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, CmrField> _cmrPlaces;
        private CmrTemplate _currentTemplate = new();
        private Bitmap? _cmrPreviewBitmap;
        private Control? _draggedPreviewField;
        private Panel? _selectedPreviewField;
        private Point _dragOffset;
        private readonly Panel _guideVertical = new() { BackColor = Color.Red, Width = 1, Visible = false };
        private readonly Panel _guideHorizontal = new() { BackColor = Color.Red, Height = 1, Visible = false };
        private readonly Panel _guideVerticalCenter = new() { BackColor = Color.OrangeRed, Width = 1, Visible = false };
        private readonly Panel _guideHorizontalCenter = new() { BackColor = Color.OrangeRed, Height = 1, Visible = false };
        private Queue<BatchPrintJob> _batchPrintQueue = new();
        private BatchPrintJob? _activeBatchPrintJob;
        private Panel _panelDailyFields = null!;
        private TextBox _txtField7 = null!;
        private TextBox _txtField9 = null!;
        private TextBox _txtField17 = null!;
        private TextBox _txtAutofillSummary = null!;

        public MainForm()
        {
            InitializeComponent();
            _cmrPlaces = CmrPlaces.GetStandardPlaces().ToDictionary(p => p.FieldName);
            printDocument1.DefaultPageSettings.PaperSize = new PaperSize("A4", A4WidthHundredthsInch, A4HeightHundredthsInch);
            printDocument1.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);
            printDocument1.DefaultPageSettings.Landscape = false;
            printDocument1.OriginAtMargins = false;
            panelCmrPreview.Controls.Add(_guideVertical);
            panelCmrPreview.Controls.Add(_guideHorizontal);
            panelCmrPreview.Controls.Add(_guideVerticalCenter);
            panelCmrPreview.Controls.Add(_guideHorizontalCenter);
            InitializeDailyWorkspace();
            EnsureManualFieldDefaults();
            _txtField9.Text = GetFieldValue(GetField9FieldName());
            LoadAppData();
            LoadTemplates();
            LoadCustomersIntoComboBox();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            GenerateCmrPreview();
            RegenerateCmrFields();
            ConfigureDailyMode();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);

            if (_panelDailyFields != null && !panelCmrPreview.Visible)
            {
                ApplyDailyLayout();
            }
        }

        private void InitializeDailyWorkspace()
        {
            var field9LabelText = GetField9LabelText();

            _panelDailyFields = new Panel
            {
                BorderStyle = BorderStyle.FixedSingle,
                Location = flowLayoutPanelFields.Location,
                Size = flowLayoutPanelFields.Size,
                Anchor = flowLayoutPanelFields.Anchor,
                AutoScroll = true,
            };

            var lblHint = new Label
            {
                Text = "Daily CMR input. Auto-filled data comes from the selected customer and linked profiles.",
                Location = new Point(12, 12),
                AutoSize = true,
            };

            var lblField7 = new Label { Text = "Field 7 - Packaging / marks", Location = new Point(12, 50), AutoSize = true };
            _txtField7 = new TextBox { Location = new Point(12, 74), Size = new Size(420, 88), Multiline = true, ScrollBars = ScrollBars.Vertical };
            _txtField7.TextChanged += (_, _) => SetFieldValue("PackagingType", _txtField7.Text);

            var lblField9 = new Label { Text = field9LabelText, Location = new Point(460, 50), AutoSize = true };
            _txtField9 = new TextBox
            {
                Location = new Point(460, 74),
                Size = new Size(300, 116),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                AcceptsReturn = true,
            };
            _txtField9.TextChanged += (_, _) => SetFieldValue(GetField9FieldName(), _txtField9.Text);

            var lblField17 = new Label { Text = "Field 17 - Transport authorizations", Location = new Point(12, 208), AutoSize = true };
            _txtField17 = new TextBox
            {
                Location = new Point(12, 232),
                Size = new Size(748, 120),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                AcceptsReturn = true,
            };
            _txtField17.TextChanged += (_, _) => SetFieldValue("TransportAuthorizations", _txtField17.Text);

            var lblSummary = new Label { Text = "Auto-filled summary", Location = new Point(12, 372), AutoSize = true };
            _txtAutofillSummary = new TextBox
            {
                Location = new Point(12, 396),
                Size = new Size(780, 170),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
            };

            _panelDailyFields.Controls.AddRange(new Control[]
            {
                lblHint, lblField7, _txtField7, lblField9, _txtField9, lblField17, _txtField17, lblSummary, _txtAutofillSummary
            });

            Controls.Add(_panelDailyFields);
            _panelDailyFields.BringToFront();
        }

        private void ConfigureDailyMode()
        {
            panelCmrPreview.Visible = false;
            panelAdjust.Visible = false;
            flowLayoutPanelFields.Visible = false;
            btnSaveTemplate.Visible = false;
            btnDeleteTemplate.Visible = false;
            btnNewCustomer.Visible = false;
            btnSaveCustomer.Visible = false;
            txtConsignorName.ReadOnly = true;
            txtConsignorAddress.ReadOnly = true;
            txtConsignorCity.ReadOnly = true;
            txtConsignorCountry.ReadOnly = true;
            txtVatNumber.ReadOnly = true;
            btnLoadTemplate.Text = "Use";
            ApplyDailyLayout();
        }

        private void ApplyDailyLayout()
        {
            var sideMargin = 12;
            var topY = menuStrip1.Bottom + 8;
            var availableWidth = ClientSize.Width - (sideMargin * 2);

            panelCustomer.Location = new Point(sideMargin, topY);
            panelCustomer.Size = new Size(availableWidth, 170);

            _panelDailyFields.Location = new Point(sideMargin, panelCustomer.Bottom + 10);
            _panelDailyFields.Size = new Size(availableWidth, ClientSize.Height - _panelDailyFields.Location.Y - panelTemplates.Height - 20);

            panelTemplates.Location = new Point(sideMargin, ClientSize.Height - panelTemplates.Height - 12);
            panelTemplates.Size = new Size(availableWidth, panelTemplates.Height);
        }

        private void LoadTemplates()
        {
            cmbTemplates.Items.Clear();
            var templates = TemplateManager.GetAvailableTemplates();
            foreach (var template in templates)
            {
                cmbTemplates.Items.Add(template);
            }
        }

        private void GenerateCmrPreview()
        {
            var maxX = _cmrPlaces.Values.Max(place => place.DefaultX) + 180f;
            var maxY = _cmrPlaces.Values.Max(place => place.DefaultY) + 120f;
            var previewWidth = (int)Math.Ceiling(maxX * PreviewScale) + (PreviewPadding * 2);
            var previewHeight = (int)Math.Ceiling(maxY * PreviewScale) + (PreviewPadding * 2);

            _cmrPreviewBitmap = new Bitmap(previewWidth, previewHeight);
            panelCmrPreview.AutoScrollMinSize = new Size(previewWidth + PreviewPadding, previewHeight + PreviewPadding);
            using var g = Graphics.FromImage(_cmrPreviewBitmap);
            g.Clear(Color.White);
            g.DrawRectangle(Pens.Black, PreviewPadding, PreviewPadding, previewWidth - (PreviewPadding * 2), previewHeight - (PreviewPadding * 2));
            g.DrawString("CMR Layout", new Font("Arial", 10, FontStyle.Bold), Brushes.Black, PreviewPadding + 6, PreviewPadding + 6);

            var smallFont = new Font("Arial", 7);
            foreach (var place in _cmrPlaces.Values.OrderBy(p => p.PlaceNumber))
            {
                var previewPoint = DocumentToPreview(place.DefaultX, place.DefaultY);
                var fieldSize = DocumentSizeToPreview(GetFieldSize(place));
                g.DrawRectangle(Pens.Gainsboro, previewPoint.X, previewPoint.Y, fieldSize.Width, fieldSize.Height);
                g.DrawString($"{place.PlaceNumber}", smallFont, Brushes.DarkGreen, previewPoint.X + 3, previewPoint.Y + 2);
            }
        }

        private void LoadAppData()
        {
            _appData = AppDataManager.Load();

            _customers.Clear();
            _customers.AddRange(_appData.Customers.OrderBy(customer => customer.Name));

            _exporters.Clear();
            _exporters.AddRange(_appData.Exporters.OrderBy(profile => profile.Name));

            _transportInfos.Clear();
            _transportInfos.AddRange(_appData.TransportInfos.OrderBy(profile => profile.Name));

            _loadingPlaces.Clear();
            _loadingPlaces.AddRange(_appData.LoadingPlaces.OrderBy(profile => profile.Name));
        }

        private void SaveAppData()
        {
            _appData.Customers = _customers.OrderBy(customer => customer.Name).ToList();
            _appData.Exporters = _exporters.OrderBy(profile => profile.Name).ToList();
            _appData.TransportInfos = _transportInfos.OrderBy(profile => profile.Name).ToList();
            _appData.LoadingPlaces = _loadingPlaces.OrderBy(profile => profile.Name).ToList();
            AppDataManager.Save(_appData);
        }

        private void SetFieldValue(string fieldName, string value)
        {
            var normalizedFieldName = NormalizeFieldName(fieldName);
            _documentFieldValues[normalizedFieldName] = value;
            if (_fieldControls.TryGetValue(normalizedFieldName, out var control))
            {
                control.SetValue(value);
            }
        }

        private string GetFieldValue(string fieldName)
        {
            var normalizedFieldName = NormalizeFieldName(fieldName);
            return _documentFieldValues.TryGetValue(normalizedFieldName, out var value) ? value : string.Empty;
        }

        private string NormalizeFieldName(string fieldName)
        {
            return string.Equals(fieldName, LegacyField9FieldName, StringComparison.OrdinalIgnoreCase)
                ? GetField9FieldName()
                : fieldName;
        }

        private string GetField9FieldName()
        {
            return _cmrPlaces.ContainsKey(RenamedField9FieldName)
                ? RenamedField9FieldName
                : LegacyField9FieldName;
        }

        private string GetField9LabelText()
        {
            if (_cmrPlaces.TryGetValue(GetField9FieldName(), out var place))
            {
                return $"Field 9 - {place.Description[(place.Description.IndexOf('.') + 1)..].Trim()}";
            }

            return "Field 9";
        }

        private bool IsManualField(string fieldName)
        {
            var normalizedFieldName = NormalizeFieldName(fieldName);
            return string.Equals(normalizedFieldName, "PackagingType", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedFieldName, GetField9FieldName(), StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedFieldName, "TransportAuthorizations", StringComparison.OrdinalIgnoreCase);
        }

        private void EnsureManualFieldDefaults()
        {
            if (string.IsNullOrWhiteSpace(GetFieldValue(GetField9FieldName())))
            {
                SetFieldValue(GetField9FieldName(), Field9StarterText);
            }
        }

        private void RegenerateCmrFields()
        {
            _fieldControls.Clear();
            _previewFieldControls.Clear();
            flowLayoutPanelFields.Controls.Clear();
            panelCmrPreview.Controls.Clear();
            panelCmrPreview.Controls.Add(_guideVertical);
            panelCmrPreview.Controls.Add(_guideHorizontal);
            panelCmrPreview.Controls.Add(_guideVerticalCenter);
            panelCmrPreview.Controls.Add(_guideHorizontalCenter);
            _selectedPreviewField = null;
            HideGuides();

            var places = _cmrPlaces.Values.OrderBy(p => p.PlaceNumber).ToList();
            var top = 10;

            foreach (var place in places)
            {
                if (!_fieldPositions.ContainsKey(place.FieldName))
                {
                    _fieldPositions[place.FieldName] = new PointF(place.DefaultX, place.DefaultY);
                }

                if (!_fieldSizes.ContainsKey(place.FieldName))
                {
                    _fieldSizes[place.FieldName] = new Size(DefaultDocumentFieldWidth, DefaultDocumentFieldHeight);
                }

                var control = new CmrFieldControl(place, _currentTemplate);
                control.Tag = place.FieldName;
                control.Size = new Size(MenuFieldWidth, MenuFieldHeight);
                control.Location = new Point(10, top);
                control.LayoutSettingsChanged += FieldControl_LayoutSettingsChanged;
                control.SetValue(GetFieldValue(place.FieldName));
                _fieldControls[place.FieldName] = control;
                flowLayoutPanelFields.Controls.Add(control);

                var previewControl = CreatePreviewFieldControl(place);
                _previewFieldControls[place.FieldName] = previewControl;
                panelCmrPreview.Controls.Add(previewControl);
                previewControl.BringToFront();

                top += MenuFieldHeight + MenuFieldSpacing;
            }

            flowLayoutPanelFields.AutoScrollMinSize = new Size(0, top + 10);

            if (customerComboBox.SelectedItem is Customer customer)
            {
                ApplyCustomerToFields(customer);
            }

            panelCmrPreview.Invalidate();
        }

        private Panel CreatePreviewFieldControl(CmrField place)
        {
            var previewControl = new Panel
            {
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.WhiteSmoke,
                Tag = place.FieldName,
                Cursor = Cursors.SizeAll,
            };

            var label = new Label
            {
                Dock = DockStyle.Fill,
                Text = GetPreviewFieldLabel(place),
                Font = new Font("Arial", 7.5f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true,
                Padding = new Padding(4, 0, 4, 0),
                BackColor = Color.Transparent,
                Cursor = Cursors.SizeAll,
            };

            previewControl.Controls.Add(label);
            AttachPreviewDragHandlers(previewControl);
            AttachPreviewDragHandlers(label);
            ApplyPreviewFieldSize(previewControl, GetFieldSize(place));
            PositionPreviewField(previewControl, GetFieldPosition(place));

            return previewControl;
        }

        private void FieldControl_LayoutSettingsChanged(object? sender, EventArgs e)
        {
            if (sender is not CmrFieldControl fieldControl)
                return;

            var fieldName = fieldControl.Place.FieldName;
            _fieldSizes[fieldName] = new Size(fieldControl.GetFieldWidth(), fieldControl.GetFieldHeight());

            if (_previewFieldControls.TryGetValue(fieldName, out var previewControl))
            {
                ApplyPreviewFieldSize(previewControl, _fieldSizes[fieldName]);
                PositionPreviewField(previewControl, GetFieldPosition(fieldControl.Place));
            }

            GenerateCmrPreview();
            panelCmrPreview.Invalidate();
        }

        private void AttachPreviewDragHandlers(Control control)
        {
            control.MouseDown += PreviewField_MouseDown;
            control.MouseMove += PreviewField_MouseMove;
            control.MouseUp += PreviewField_MouseUp;
            control.Click += PreviewField_Click;
        }

        private void PreviewField_Click(object? sender, EventArgs e)
        {
            var previewField = sender as Control;
            if (previewField == null)
                return;

            SelectPreviewField(previewField.Parent == panelCmrPreview ? previewField as Panel : previewField.Parent as Panel);
        }

        private void PreviewField_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
                return;

            _draggedPreviewField = sender is Control child && child.Parent == panelCmrPreview ? child : (sender as Control)?.Parent;
            if (_draggedPreviewField == null)
                return;

            SelectPreviewField(_draggedPreviewField as Panel);
            _dragOffset = e.Location;
            if (!ReferenceEquals(sender, _draggedPreviewField))
            {
                _dragOffset = _draggedPreviewField.PointToClient(Cursor.Position);
            }
        }

        private void PreviewField_MouseMove(object? sender, MouseEventArgs e)
        {
            if (_draggedPreviewField == null || e.Button != MouseButtons.Left)
                return;

            var cursorPoint = panelCmrPreview.PointToClient(Cursor.Position);
            var canvasSize = panelCmrPreview.AutoScrollMinSize;
            var maxX = Math.Max(PreviewPadding, canvasSize.Width - _draggedPreviewField.Width - PreviewPadding);
            var maxY = Math.Max(PreviewPadding, canvasSize.Height - _draggedPreviewField.Height - PreviewPadding);

            var newX = Math.Max(PreviewPadding, Math.Min(cursorPoint.X - _dragOffset.X, maxX));
            var newY = Math.Max(PreviewPadding, Math.Min(cursorPoint.Y - _dragOffset.Y, maxY));

            _draggedPreviewField.Location = new Point(newX, newY);
            SyncDocumentPositionFromPreview(_draggedPreviewField);
            UpdateGuides(_draggedPreviewField as Panel);
        }

        private void PreviewField_MouseUp(object? sender, MouseEventArgs e)
        {
            _draggedPreviewField = null;
        }

        private void PositionPreviewField(Control previewControl, PointF documentPoint)
        {
            var previewPoint = DocumentToPreview(documentPoint.X, documentPoint.Y);
            var canvasSize = panelCmrPreview.AutoScrollMinSize;
            var maxX = Math.Max(PreviewPadding, canvasSize.Width - previewControl.Width - PreviewPadding);
            var maxY = Math.Max(PreviewPadding, canvasSize.Height - previewControl.Height - PreviewPadding);
            previewControl.Location = new Point(
                Math.Max(PreviewPadding, Math.Min(previewPoint.X, maxX)),
                Math.Max(PreviewPadding, Math.Min(previewPoint.Y, maxY)));
        }

        private void SyncDocumentPositionFromPreview(Control previewControl)
        {
            if (previewControl.Tag is not string fieldName)
                return;

            _fieldPositions[fieldName] = PreviewToDocument(previewControl.Left, previewControl.Top);
        }

        private void ApplyPreviewFieldSize(Panel previewControl, Size documentSize)
        {
            var previewSize = DocumentSizeToPreview(documentSize);
            previewControl.Size = previewSize;
        }

        private void SelectPreviewField(Panel? previewField)
        {
            if (_selectedPreviewField == previewField)
            {
                UpdateGuides(previewField);
                return;
            }

            if (_selectedPreviewField != null)
            {
                _selectedPreviewField.BackColor = Color.WhiteSmoke;
            }

            _selectedPreviewField = previewField;

            if (_selectedPreviewField != null)
            {
                _selectedPreviewField.BackColor = Color.MistyRose;
                _selectedPreviewField.Focus();
                UpdateGuides(_selectedPreviewField);
            }
            else
            {
                HideGuides();
            }
        }

        private void UpdateGuides(Panel? previewField)
        {
            if (previewField == null)
            {
                HideGuides();
                return;
            }

            var canvas = panelCmrPreview.AutoScrollMinSize;
            _guideVertical.Location = new Point(previewField.Left, 0);
            _guideVertical.Height = canvas.Height;
            _guideVertical.Visible = true;

            _guideHorizontal.Location = new Point(0, previewField.Top);
            _guideHorizontal.Width = canvas.Width;
            _guideHorizontal.Visible = true;

            _guideVerticalCenter.Location = new Point(previewField.Left + (previewField.Width / 2), 0);
            _guideVerticalCenter.Height = canvas.Height;
            _guideVerticalCenter.Visible = true;

            _guideHorizontalCenter.Location = new Point(0, previewField.Top + (previewField.Height / 2));
            _guideHorizontalCenter.Width = canvas.Width;
            _guideHorizontalCenter.Visible = true;

            _guideVertical.BringToFront();
            _guideHorizontal.BringToFront();
            _guideVerticalCenter.BringToFront();
            _guideHorizontalCenter.BringToFront();
            previewField.BringToFront();
        }

        private void HideGuides()
        {
            _guideVertical.Visible = false;
            _guideHorizontal.Visible = false;
            _guideVerticalCenter.Visible = false;
            _guideHorizontalCenter.Visible = false;
        }

        private void MoveSelectedPreviewField(int dx, int dy)
        {
            if (_selectedPreviewField == null)
                return;

            var canvasSize = panelCmrPreview.AutoScrollMinSize;
            var maxX = Math.Max(PreviewPadding, canvasSize.Width - _selectedPreviewField.Width - PreviewPadding);
            var maxY = Math.Max(PreviewPadding, canvasSize.Height - _selectedPreviewField.Height - PreviewPadding);
            var newX = Math.Max(PreviewPadding, Math.Min(_selectedPreviewField.Left + dx, maxX));
            var newY = Math.Max(PreviewPadding, Math.Min(_selectedPreviewField.Top + dy, maxY));

            _selectedPreviewField.Location = new Point(newX, newY);
            SyncDocumentPositionFromPreview(_selectedPreviewField);
            UpdateGuides(_selectedPreviewField);
        }

        private Point DocumentToPreview(float x, float y)
        {
            return new Point(
                PreviewPadding + (int)Math.Round(x * PreviewScale),
                PreviewPadding + (int)Math.Round(y * PreviewScale));
        }

        private Size DocumentSizeToPreview(Size size)
        {
            return new Size(
                Math.Max(28, (int)Math.Round(size.Width * PreviewScale)),
                Math.Max(18, (int)Math.Round(size.Height * PreviewScale)));
        }

        private PointF PreviewToDocument(int x, int y)
        {
            return new PointF(
                Math.Max(0, (x - PreviewPadding) / PreviewScale),
                Math.Max(0, (y - PreviewPadding) / PreviewScale));
        }

        private PointF GetFieldPosition(CmrField place)
        {
            return _fieldPositions.TryGetValue(place.FieldName, out var position)
                ? position
                : new PointF(place.DefaultX, place.DefaultY);
        }

        private Size GetFieldSize(CmrField place)
        {
            return _fieldSizes.TryGetValue(place.FieldName, out var size)
                ? size
                : new Size(DefaultDocumentFieldWidth, DefaultDocumentFieldHeight);
        }

        private static string GetPreviewFieldLabel(CmrField place)
        {
            var shortDescription = place.Description;
            var separatorIndex = shortDescription.IndexOf(' ');
            if (separatorIndex >= 0)
            {
                shortDescription = shortDescription[(separatorIndex + 1)..];
            }

            if (shortDescription.Length > 14)
            {
                shortDescription = shortDescription[..14];
            }

            return $"{place.PlaceNumber}. {shortDescription}";
        }

        private void btnSaveTemplate_Click(object sender, EventArgs e)
        {
            var templateName = PromptForTemplateName();
            if (string.IsNullOrEmpty(templateName))
                return;

            _currentTemplate.Name = templateName;
            _currentTemplate.FontSizes.Clear();
            _currentTemplate.VerticalOffsets.Clear();
            _currentTemplate.FieldPositions.Clear();
            _currentTemplate.FieldWidths.Clear();
            _currentTemplate.FieldHeights.Clear();

            foreach (var field in _fieldControls.Values)
            {
                _currentTemplate.FontSizes[field.Place.FieldName] = field.GetFontSize();
                _currentTemplate.VerticalOffsets[field.Place.FieldName] = field.GetVerticalOffset();
                _currentTemplate.FieldWidths[field.Place.FieldName] = field.GetFieldWidth();
                _currentTemplate.FieldHeights[field.Place.FieldName] = field.GetFieldHeight();
            }

            foreach (var entry in _fieldPositions)
            {
                _currentTemplate.FieldPositions[entry.Key] = entry.Value;
            }

            TemplateManager.SaveTemplate(_currentTemplate);
            LoadTemplates();
            cmbTemplates.SelectedItem = templateName;
            MessageBox.Show(this, $"Template '{templateName}' saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnLoadTemplate_Click(object sender, EventArgs e)
        {
            if (cmbTemplates.SelectedItem is not string templateName)
            {
                MessageBox.Show(this, "Please select a template to load.", "Selection required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            LoadTemplateIntoMain(templateName);
            MessageBox.Show(this, $"Template '{templateName}' loaded.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void LoadTemplateIntoMain(string templateName)
        {
            _currentTemplate = TemplateManager.LoadTemplate(templateName);
            LoadFieldPositionsFromTemplate();
            LoadFieldSizesFromTemplate();
            GenerateCmrPreview();
            RegenerateCmrFields();
            cmbTemplates.SelectedItem = templateName;
        }

        private void btnDeleteTemplate_Click(object sender, EventArgs e)
        {
            if (cmbTemplates.SelectedItem is not string templateName)
            {
                MessageBox.Show(this, "Please select a template to delete.", "Selection required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show(this, $"Delete template '{templateName}'?", "Confirm deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                TemplateManager.DeleteTemplate(templateName);
                LoadTemplates();
                MessageBox.Show(this, "Template deleted.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private string PromptForTemplateName()
        {
            using var form = new Form
            {
                Text = "Save Template",
                Width = 300,
                Height = 150,
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
            };

            var label = new Label { Text = "Template name:", Left = 10, Top = 20, Width = 260 };
            var textBox = new TextBox { Left = 10, Top = 50, Width = 260, Text = DateTime.Now.ToString("CMR_yyyy-MM-dd_HHmm") };
            var btnOK = new Button { Text = "OK", Left = 110, Top = 85, Width = 80, DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "Cancel", Left = 200, Top = 85, Width = 70, DialogResult = DialogResult.Cancel };

            form.Controls.Add(label);
            form.Controls.Add(textBox);
            form.Controls.Add(btnOK);
            form.Controls.Add(btnCancel);
            form.AcceptButton = btnOK;
            form.CancelButton = btnCancel;

            return form.ShowDialog(this) == DialogResult.OK ? textBox.Text : string.Empty;
        }

        private void LoadCustomersIntoComboBox()
        {
            customerComboBox.Items.Clear();
            customerComboBox.Items.AddRange(_customers.Cast<object>().ToArray());
            if (customerComboBox.Items.Count > 0)
            {
                customerComboBox.SelectedIndex = 0;
            }
        }

        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog(this) != DialogResult.OK)
                return;

            var imported = ExcelImporter.ImportCustomers(openFileDialog.FileName);
            if (imported.Count == 0)
            {
                MessageBox.Show(this, "No valid customers found.", "Import result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _customers.AddRange(imported);
            SaveAppData();
            LoadCustomersIntoComboBox();
            MessageBox.Show(this, $"Imported {imported.Count} customers.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnNewCustomer_Click(object sender, EventArgs e)
        {
            customerComboBox.SelectedIndex = -1;
        }

        private void btnSaveCustomer_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtConsignorName.Text))
            {
                MessageBox.Show(this, "Consignor name is required.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var customer = new Customer
            {
                Name = txtConsignorName.Text.Trim(),
                Address = txtConsignorAddress.Text.Trim(),
                City = txtConsignorCity.Text.Trim(),
                Country = txtConsignorCountry.Text.Trim(),
                VatNumber = txtVatNumber.Text.Trim(),
            };

            if (customerComboBox.SelectedItem is Customer existing)
            {
                existing.Name = customer.Name;
                existing.Address = customer.Address;
                existing.City = customer.City;
                existing.Country = customer.Country;
                existing.VatNumber = customer.VatNumber;
            }
            else
            {
                _customers.Add(customer);
            }

            SaveAppData();
            LoadCustomersIntoComboBox();
            customerComboBox.SelectedItem = _customers.FirstOrDefault(c => c.Name == customer.Name);
            MessageBox.Show(this, "Customer saved.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void customerComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (customerComboBox.SelectedItem is Customer customer)
            {
                txtConsignorName.Text = customer.Name;
                txtConsignorAddress.Text = customer.Address;
                txtConsignorCity.Text = customer.City;
                txtConsignorCountry.Text = customer.Country;
                txtVatNumber.Text = customer.VatNumber;
                ApplyCustomerToFields(customer);
            }
            else
            {
                ClearFields();
            }
        }

        private void ClearFields()
        {
            txtConsignorName.Clear();
            txtConsignorAddress.Clear();
            txtConsignorCity.Clear();
            txtConsignorCountry.Clear();
            txtVatNumber.Clear();
            SetFieldValue("PackagingType", string.Empty);
            SetFieldValue("TransportAuthorizations", string.Empty);
            EnsureManualFieldDefaults();
            _txtField7.Text = GetFieldValue("PackagingType");
            _txtField9.Text = GetFieldValue(GetField9FieldName());
            _txtField17.Text = GetFieldValue("TransportAuthorizations");
            _txtAutofillSummary.Clear();
        }

        private void ApplyCustomerToFields(Customer customer)
        {
            var preservedManualValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["PackagingType"] = GetFieldValue("PackagingType"),
                [GetField9FieldName()] = GetFieldValue(GetField9FieldName()),
                ["TransportAuthorizations"] = GetFieldValue("TransportAuthorizations")
            };

            foreach (var fieldControl in _fieldControls.Values.OrderBy(control => control.Place.PlaceNumber))
            {
                if (!IsManualField(fieldControl.Place.FieldName))
                {
                    SetFieldValue(fieldControl.Place.FieldName, string.Empty);
                }
            }

            ApplyProfileAssignments(GetProfileByName(_exporters, customer.ExporterProfileName));
            ApplyProfileAssignments(GetProfileByName(_transportInfos, customer.TransportProfileName));
            ApplyProfileAssignments(GetProfileByName(_loadingPlaces, customer.LoadingPlaceProfileName));
            ApplyAssignments(customer.FieldAssignments);

            if (_fieldControls.TryGetValue("ConsignorName", out var consignorControl) && string.IsNullOrWhiteSpace(consignorControl.GetValue()))
            {
                SetFieldValue("ConsignorName", BuildCustomerBlock(customer));
            }

            if (_fieldControls.TryGetValue("ExportDate", out _))
            {
                var dateValue = string.IsNullOrWhiteSpace(customer.PlaceOfIssue)
                    ? DateTime.Today.ToString("dd-MM-yyyy")
                    : $"{customer.PlaceOfIssue} {DateTime.Today:dd-MM-yyyy}";
                var combinedPlaceDate = MergeFieldLines(GetFieldValue("ConsignorRemarks"), GetFieldValue("ExportDate"), dateValue);
                SetFieldValue("ConsignorRemarks", combinedPlaceDate);
                SetFieldValue("ExportDate", combinedPlaceDate);
            }

            foreach (var manualField in preservedManualValues)
            {
                SetFieldValue(manualField.Key, manualField.Value);
            }

            EnsureManualFieldDefaults();
            _txtField7.Text = GetFieldValue("PackagingType");
            _txtField9.Text = GetFieldValue(GetField9FieldName());
            _txtField17.Text = GetFieldValue("TransportAuthorizations");
            RefreshAutofillSummary(customer);
        }

        private static string MergeFieldLines(params string[] values)
        {
            var lines = new List<string>();
            foreach (var value in values)
            {
                if (string.IsNullOrWhiteSpace(value))
                    continue;

                foreach (var line in value
                    .Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)
                    .Select(part => part.Trim())
                    .Where(part => !string.IsNullOrWhiteSpace(part)))
                {
                    if (!lines.Contains(line, StringComparer.OrdinalIgnoreCase))
                    {
                        lines.Add(line);
                    }
                }
            }

            return string.Join(Environment.NewLine, lines);
        }

        private void ApplyProfileAssignments(ProfileRecord? profile)
        {
            if (profile == null)
                return;

            ApplyAssignments(profile.FieldAssignments);
        }

        private void ApplyAssignments(IEnumerable<FieldAssignment> assignments)
        {
            foreach (var assignment in assignments)
            {
                if (IsManualField(assignment.FieldName))
                    continue;

                if (_fieldControls.TryGetValue(assignment.FieldName, out var control))
                {
                    SetFieldValue(assignment.FieldName, assignment.Value);
                }
            }
        }

        private static ProfileRecord? GetProfileByName(IEnumerable<ProfileRecord> profiles, string name)
        {
            return string.IsNullOrWhiteSpace(name)
                ? null
                : profiles.FirstOrDefault(profile => string.Equals(profile.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        private void RefreshAutofillSummary(Customer customer)
        {
            var lines = new List<string>
            {
                $"Customer: {customer.Name}",
                $"Exporter profile: {customer.ExporterProfileName}",
                $"Transport profile: {customer.TransportProfileName}",
                $"Loading place: {customer.LoadingPlaceProfileName}",
                $"Field 21: {GetFieldValue("ExportDate")}"
            };

            foreach (var place in _cmrPlaces.Values.OrderBy(place => place.PlaceNumber))
            {
                if (IsManualField(place.FieldName))
                    continue;

                var value = GetFieldValue(place.FieldName);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    lines.Add($"{place.Description}: {value.Replace(Environment.NewLine, " | ")}");
                }
            }

            _txtAutofillSummary.Text = string.Join(Environment.NewLine, lines);
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            if (!EnsurePrinterAvailable(requireUserSelection: true))
                return;

            try
            {
                printPreviewDialog.Document = printDocument1;
                printPreviewDialog.ShowDialog(this);
            }
            catch (InvalidPrinterException)
            {
                MessageBox.Show(this, "No valid printer is available for preview. Please choose an installed printer first.", "Printer required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (!EnsurePrinterAvailable(requireUserSelection: true))
                return;

            _batchPrintQueue.Clear();
            _activeBatchPrintJob = null;
            printDocument1.PrinterSettings.Copies = PrintCopies;
            printDialog.PrinterSettings = printDocument1.PrinterSettings;

            try
            {
                if (printDialog.ShowDialog(this) != DialogResult.OK)
                    return;
            }
            catch (InvalidPrinterException)
            {
                MessageBox.Show(this, "No valid printer is available. Please choose an installed printer first.", "Printer required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _batchPrintQueue.Clear();
            _activeBatchPrintJob = null;
            printDocument1.PrinterSettings.Copies = PrintCopies;
            printDocument1.Print();
        }

        private void LoadFieldPositionsFromTemplate()
        {
            _fieldPositions.Clear();

            foreach (var place in _cmrPlaces.Values)
            {
                if (_currentTemplate.FieldPositions.TryGetValue(place.FieldName, out var position))
                {
                    _fieldPositions[place.FieldName] = position;
                }
                else
                {
                    _fieldPositions[place.FieldName] = new PointF(place.DefaultX, place.DefaultY);
                }
            }
        }

        private void LoadFieldSizesFromTemplate()
        {
            _fieldSizes.Clear();

            foreach (var place in _cmrPlaces.Values)
            {
                var width = _currentTemplate.FieldWidths.TryGetValue(place.FieldName, out var savedWidth)
                    ? savedWidth
                    : DefaultDocumentFieldWidth;
                var height = _currentTemplate.FieldHeights.TryGetValue(place.FieldName, out var savedHeight)
                    ? savedHeight
                    : DefaultDocumentFieldHeight;

                _fieldSizes[place.FieldName] = new Size(width, height);
            }
        }

        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            var g = e.Graphics;
            var places = CmrPlaces.GetStandardPlaces();
            var hardMarginX = e.PageSettings.HardMarginX;
            var hardMarginY = e.PageSettings.HardMarginY;
            g.TranslateTransform(-hardMarginX, -hardMarginY);

            var documentBounds = GetDocumentBounds();
            var pageBounds = e.PageBounds;
            var scaleX = pageBounds.Width / documentBounds.Width;
            var scaleY = pageBounds.Height / documentBounds.Height;
            var fontScale = Math.Min(scaleX, scaleY);

            foreach (var place in places)
            {
                if (_fieldControls.TryGetValue(place.FieldName, out var control))
                {
                    var fontSize = Math.Max(6f, control.GetFontSize() * fontScale);
                    var offset = control.GetVerticalOffset();
                    using var font = new Font("Arial", fontSize);
                    var position = GetFieldPosition(place);
                    var fieldSize = GetFieldSize(place);
                    var bounds = new RectangleF(
                        position.X * scaleX,
                        (position.Y + offset) * scaleY,
                        fieldSize.Width * scaleX,
                        fieldSize.Height * scaleY);
                    var text = ResolveFieldText(place, control);

                    if (!string.IsNullOrEmpty(text))
                    {
                        using var format = new StringFormat
                        {
                            Alignment = StringAlignment.Near,
                            LineAlignment = StringAlignment.Near,
                            Trimming = StringTrimming.Word,
                        };
                        g.DrawString(text, font, Brushes.Black, bounds, format);
                    }
                }
            }

            if (_batchPrintQueue.Count > 0)
            {
                _activeBatchPrintJob = _batchPrintQueue.Dequeue();
                e.HasMorePages = true;
            }
            else
            {
                _activeBatchPrintJob = null;
                e.HasMorePages = false;
            }
        }

        private string ResolveFieldText(CmrField place, CmrFieldControl control)
        {
            if (_activeBatchPrintJob == null)
            {
                return GetFieldValue(place.FieldName);
            }

            return place.FieldName switch
            {
                "ConsignorName" => BuildCustomerBlock(_activeBatchPrintJob.Customer),
                var fieldName when string.Equals(fieldName, GetField9FieldName(), StringComparison.OrdinalIgnoreCase) => string.IsNullOrWhiteSpace(_activeBatchPrintJob.Field9Value) ? GetFieldValue(fieldName) : _activeBatchPrintJob.Field9Value,
                _ => GetFieldValue(place.FieldName),
            };
        }

        private static string BuildCustomerBlock(Customer customer)
        {
            var lines = new[]
            {
                customer.Name,
                customer.Address,
                customer.City,
                customer.Country
            };

            return string.Join(Environment.NewLine, lines.Where(line => !string.IsNullOrWhiteSpace(line)));
        }

        private SizeF GetDocumentBounds()
        {
            var maxX = 0f;
            var maxY = 0f;

            foreach (var place in _cmrPlaces.Values)
            {
                var position = GetFieldPosition(place);
                var size = GetFieldSize(place);
                maxX = Math.Max(maxX, position.X + size.Width);
                maxY = Math.Max(maxY, position.Y + size.Height);
            }

            return new SizeF(Math.Max(1f, maxX + 20f), Math.Max(1f, maxY + 20f));
        }

        private void panelCmrPreview_Paint(object sender, PaintEventArgs e)
        {
            if (_cmrPreviewBitmap != null)
            {
                e.Graphics.TranslateTransform(panelCmrPreview.AutoScrollPosition.X, panelCmrPreview.AutoScrollPosition.Y);
                e.Graphics.DrawImageUnscaled(_cmrPreviewBitmap, 0, 0);
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnMoveUp_Click(object sender, EventArgs e)
        {
            MoveSelectedPreviewField(0, -1);
        }

        private void btnMoveDown_Click(object sender, EventArgs e)
        {
            MoveSelectedPreviewField(0, 1);
        }

        private void btnMoveLeft_Click(object sender, EventArgs e)
        {
            MoveSelectedPreviewField(-1, 0);
        }

        private void btnMoveRight_Click(object sender, EventArgs e)
        {
            MoveSelectedPreviewField(1, 0);
        }

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            var step = e.Shift ? 5 : 1;
            switch (e.KeyCode)
            {
                case Keys.Up:
                    MoveSelectedPreviewField(0, -step);
                    e.Handled = true;
                    break;
                case Keys.Down:
                    MoveSelectedPreviewField(0, step);
                    e.Handled = true;
                    break;
                case Keys.Left:
                    MoveSelectedPreviewField(-step, 0);
                    e.Handled = true;
                    break;
                case Keys.Right:
                    MoveSelectedPreviewField(step, 0);
                    e.Handled = true;
                    break;
            }
        }

        private void batchPrintToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_currentTemplate.Name))
            {
                MessageBox.Show(this, "Load or save a template first, then open batch print.", "Template required", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var form = new BatchPrintForm(_customers);
            if (form.ShowDialog(this) != DialogResult.OK)
                return;

            var jobs = form.GetSelectedJobs();
            if (jobs.Count == 0)
            {
                MessageBox.Show(this, "No customers were selected for batch printing.", "Nothing selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!EnsurePrinterAvailable(requireUserSelection: true))
                return;

            printDocument1.PrinterSettings.Copies = PrintCopies;
            printDialog.PrinterSettings = printDocument1.PrinterSettings;
            try
            {
                if (printDialog.ShowDialog(this) != DialogResult.OK)
                    return;
            }
            catch (InvalidPrinterException)
            {
                MessageBox.Show(this, "No valid printer is available. Please choose an installed printer first.", "Printer required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _batchPrintQueue = new Queue<BatchPrintJob>(jobs.Select(job => new BatchPrintJob
            {
                Customer = job.Customer,
                Field9Value = job.Field9Value
            }));

            _activeBatchPrintJob = _batchPrintQueue.Dequeue();
            printDocument1.PrinterSettings.Copies = PrintCopies;
            printDocument1.Print();
        }

        private bool EnsurePrinterAvailable(bool requireUserSelection)
        {
            var installedPrinters = PrinterSettings.InstalledPrinters.Cast<string>().ToList();
            if (installedPrinters.Count == 0)
            {
                MessageBox.Show(this, "No printers are installed in Windows. Please install a printer or PDF printer first.", "No printers found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            var currentPrinter = printDocument1.PrinterSettings.PrinterName;
            var hasValidCurrentPrinter = !string.IsNullOrWhiteSpace(currentPrinter) &&
                                         installedPrinters.Any(printer => string.Equals(printer, currentPrinter, StringComparison.OrdinalIgnoreCase));

            if (hasValidCurrentPrinter && !requireUserSelection)
                return true;

            if (hasValidCurrentPrinter && requireUserSelection)
            {
                return ShowPrinterSelection(installedPrinters, currentPrinter);
            }

            return ShowPrinterSelection(installedPrinters, installedPrinters.First());
        }

        private bool ShowPrinterSelection(List<string> installedPrinters, string selectedPrinter)
        {
            using var form = new Form
            {
                Text = "Choose Printer",
                Width = 460,
                Height = 170,
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
            };

            var label = new Label
            {
                Text = "Available printers:",
                Left = 12,
                Top = 20,
                Width = 420,
            };

            var combo = new ComboBox
            {
                Left = 12,
                Top = 48,
                Width = 420,
                DropDownStyle = ComboBoxStyle.DropDownList,
            };
            combo.Items.AddRange(installedPrinters.Cast<object>().ToArray());
            combo.SelectedItem = installedPrinters.FirstOrDefault(printer => string.Equals(printer, selectedPrinter, StringComparison.OrdinalIgnoreCase)) ?? installedPrinters.First();

            var btnOk = new Button
            {
                Text = "OK",
                Left = 262,
                Top = 86,
                Width = 80,
                DialogResult = DialogResult.OK,
            };

            var btnCancel = new Button
            {
                Text = "Cancel",
                Left = 352,
                Top = 86,
                Width = 80,
                DialogResult = DialogResult.Cancel,
            };

            form.Controls.Add(label);
            form.Controls.Add(combo);
            form.Controls.Add(btnOk);
            form.Controls.Add(btnCancel);
            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;

            if (form.ShowDialog(this) != DialogResult.OK || combo.SelectedItem is not string printerName)
                return false;

            printDocument1.PrinterSettings.PrinterName = printerName;
            printDialog.PrinterSettings.PrinterName = printerName;
            return true;
        }

        private void templateEditorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var form = new TemplateEditorForm(_currentTemplate.Name);
            if (form.ShowDialog(this) != DialogResult.OK)
                return;

            LoadTemplates();
            if (!string.IsNullOrWhiteSpace(form.SelectedTemplateName))
            {
                LoadTemplateIntoMain(form.SelectedTemplateName);
            }
        }

        private void exporterInfoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var form = new ProfileManagerForm("Exporter Info", _exporters, _cmrPlaces.Values);
            if (form.ShowDialog(this) != DialogResult.OK)
                return;

            _exporters.Clear();
            _exporters.AddRange(form.GetProfiles());
            SaveAppData();
        }

        private void transportInfoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var form = new ProfileManagerForm("Transport Info", _transportInfos, _cmrPlaces.Values);
            if (form.ShowDialog(this) != DialogResult.OK)
                return;

            _transportInfos.Clear();
            _transportInfos.AddRange(form.GetProfiles());
            SaveAppData();
        }

        private void loadingPlacesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var form = new ProfileManagerForm("Loading Places", _loadingPlaces, _cmrPlaces.Values);
            if (form.ShowDialog(this) != DialogResult.OK)
                return;

            _loadingPlaces.Clear();
            _loadingPlaces.AddRange(form.GetProfiles());
            SaveAppData();
        }

        private void customerInfoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var form = new CustomerManagerForm(_customers, _exporters, _transportInfos, _loadingPlaces, _cmrPlaces.Values);
            if (form.ShowDialog(this) != DialogResult.OK)
                return;

            _customers.Clear();
            _customers.AddRange(form.GetCustomers());
            SaveAppData();
            LoadCustomersIntoComboBox();
        }
    }
}
