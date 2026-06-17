using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace CMRPrint
{
    public sealed class TemplateEditorForm : Form
    {
        private const int MenuFieldWidth = 530;
        private const int MenuFieldHeight = 138;
        private const int MenuFieldSpacing = 10;
        private const float PreviewScale = 0.90f;
        private const int PreviewPadding = 18;
        private const int DefaultDocumentFieldWidth = 140;
        private const int DefaultDocumentFieldHeight = 52;

        private readonly Panel _panelPreview;
        private readonly Panel _panelAdjust;
        private readonly Panel _panelFields;
        private readonly ComboBox _cmbTemplates;
        private readonly Dictionary<string, CmrField> _places;
        private readonly Dictionary<string, CmrFieldControl> _fieldControls = new();
        private readonly Dictionary<string, Panel> _previewFieldControls = new();
        private readonly Dictionary<string, PointF> _fieldPositions = new();
        private readonly Dictionary<string, Size> _fieldSizes = new();
        private readonly Panel _guideVertical = new() { BackColor = Color.Red, Width = 1, Visible = false };
        private readonly Panel _guideHorizontal = new() { BackColor = Color.Red, Height = 1, Visible = false };
        private readonly Panel _guideVerticalCenter = new() { BackColor = Color.OrangeRed, Width = 1, Visible = false };
        private readonly Panel _guideHorizontalCenter = new() { BackColor = Color.OrangeRed, Height = 1, Visible = false };
        private Bitmap? _previewBitmap;
        private CmrTemplate _currentTemplate = new();
        private Panel? _selectedPreviewField;
        private Control? _draggedPreviewField;
        private Point _dragOffset;

        public string SelectedTemplateName => _currentTemplate.Name;

        public TemplateEditorForm(string currentTemplateName)
        {
            _places = CmrPlaces.GetStandardPlaces().ToDictionary(place => place.FieldName);
            _currentTemplate = string.IsNullOrWhiteSpace(currentTemplateName)
                ? new CmrTemplate()
                : TemplateManager.LoadTemplate(currentTemplateName);

            Text = "Template Editor";
            StartPosition = FormStartPosition.CenterParent;
            Width = 1480;
            Height = 860;
            MinimumSize = new Size(1320, 820);
            KeyPreview = true;
            KeyDown += TemplateEditorForm_KeyDown;

            _panelPreview = new Panel
            {
                Location = new Point(12, 12),
                Size = new Size(580, 710),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left,
                AutoScroll = true,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
            };
            _panelPreview.Paint += PanelPreview_Paint;

            _panelAdjust = new Panel
            {
                Location = new Point(12, 730),
                Size = new Size(580, 75),
                Anchor = AnchorStyles.Left | AnchorStyles.Bottom,
                BorderStyle = BorderStyle.FixedSingle,
            };

            var lblAdjust = new Label
            {
                Text = "Select a box, then use buttons or keyboard arrows. Shift = bigger step.",
                Location = new Point(10, 10),
                AutoSize = true,
            };
            var btnLeft = new Button { Text = "Left", Location = new Point(292, 38), Size = new Size(56, 24) };
            var btnRight = new Button { Text = "Right", Location = new Point(414, 38), Size = new Size(60, 24) };
            var btnUp = new Button { Text = "Up", Location = new Point(354, 10), Size = new Size(48, 24) };
            var btnDown = new Button { Text = "Down", Location = new Point(354, 38), Size = new Size(54, 24) };
            btnLeft.Click += (_, _) => MoveSelectedPreviewField(-1, 0);
            btnRight.Click += (_, _) => MoveSelectedPreviewField(1, 0);
            btnUp.Click += (_, _) => MoveSelectedPreviewField(0, -1);
            btnDown.Click += (_, _) => MoveSelectedPreviewField(0, 1);
            _panelAdjust.Controls.AddRange(new Control[] { lblAdjust, btnLeft, btnRight, btnUp, btnDown });

            _panelFields = new Panel
            {
                Location = new Point(606, 12),
                Size = new Size(850, 720),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle,
            };

            var bottomPanel = new Panel
            {
                Location = new Point(606, 742),
                Size = new Size(850, 63),
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                BorderStyle = BorderStyle.FixedSingle,
            };

            _cmbTemplates = new ComboBox
            {
                Location = new Point(10, 16),
                Size = new Size(420, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
            };

            var btnLoad = new Button { Text = "Load", Location = new Point(440, 15), Size = new Size(82, 30) };
            var btnSave = new Button { Text = "Save", Location = new Point(530, 15), Size = new Size(82, 30) };
            var btnDelete = new Button { Text = "Delete", Location = new Point(620, 15), Size = new Size(82, 30) };
            var btnClose = new Button { Text = "Done", Location = new Point(748, 15), Size = new Size(82, 30), DialogResult = DialogResult.OK };
            btnLoad.Click += BtnLoad_Click;
            btnSave.Click += BtnSave_Click;
            btnDelete.Click += BtnDelete_Click;
            bottomPanel.Controls.AddRange(new Control[] { _cmbTemplates, btnLoad, btnSave, btnDelete, btnClose });

            Controls.AddRange(new Control[] { _panelPreview, _panelAdjust, _panelFields, bottomPanel });

            _panelPreview.Controls.Add(_guideVertical);
            _panelPreview.Controls.Add(_guideHorizontal);
            _panelPreview.Controls.Add(_guideVerticalCenter);
            _panelPreview.Controls.Add(_guideHorizontalCenter);

            LoadTemplates();
            LoadCurrentTemplateIntoWorkingState();
            GeneratePreview();
            RegenerateEditorControls();
        }

        private void LoadTemplates()
        {
            _cmbTemplates.Items.Clear();
            var templates = TemplateManager.GetAvailableTemplates();
            foreach (var template in templates)
            {
                _cmbTemplates.Items.Add(template);
            }

            if (!string.IsNullOrWhiteSpace(_currentTemplate.Name))
            {
                _cmbTemplates.SelectedItem = _currentTemplate.Name;
            }
        }

        private void LoadCurrentTemplateIntoWorkingState()
        {
            _fieldPositions.Clear();
            _fieldSizes.Clear();

            foreach (var place in _places.Values)
            {
                _fieldPositions[place.FieldName] = _currentTemplate.FieldPositions.TryGetValue(place.FieldName, out var position)
                    ? position
                    : new PointF(place.DefaultX, place.DefaultY);
                _fieldSizes[place.FieldName] = new Size(
                    _currentTemplate.FieldWidths.TryGetValue(place.FieldName, out var width) ? width : DefaultDocumentFieldWidth,
                    _currentTemplate.FieldHeights.TryGetValue(place.FieldName, out var height) ? height : DefaultDocumentFieldHeight);
            }
        }

        private void GeneratePreview()
        {
            var maxX = _places.Values.Max(place => GetFieldPosition(place).X + GetFieldSize(place).Width) + 80f;
            var maxY = _places.Values.Max(place => GetFieldPosition(place).Y + GetFieldSize(place).Height) + 80f;
            var previewWidth = (int)Math.Ceiling(maxX * PreviewScale) + (PreviewPadding * 2);
            var previewHeight = (int)Math.Ceiling(maxY * PreviewScale) + (PreviewPadding * 2);

            _previewBitmap = new Bitmap(previewWidth, previewHeight);
            _panelPreview.AutoScrollMinSize = new Size(previewWidth + PreviewPadding, previewHeight + PreviewPadding);
            using var g = Graphics.FromImage(_previewBitmap);
            g.Clear(Color.White);
            g.DrawRectangle(Pens.Black, PreviewPadding, PreviewPadding, previewWidth - (PreviewPadding * 2), previewHeight - (PreviewPadding * 2));
            g.DrawString("CMR Layout", new Font("Arial", 10, FontStyle.Bold), Brushes.Black, PreviewPadding + 6, PreviewPadding + 6);

            var smallFont = new Font("Arial", 7);
            foreach (var place in _places.Values.OrderBy(place => place.PlaceNumber))
            {
                var point = DocumentToPreview(GetFieldPosition(place).X, GetFieldPosition(place).Y);
                var size = DocumentSizeToPreview(GetFieldSize(place));
                g.DrawRectangle(Pens.Gainsboro, point.X, point.Y, size.Width, size.Height);
                g.DrawString($"{place.PlaceNumber}", smallFont, Brushes.DarkGreen, point.X + 3, point.Y + 2);
            }
        }

        private void RegenerateEditorControls()
        {
            _fieldControls.Clear();
            _previewFieldControls.Clear();
            _panelFields.Controls.Clear();
            _panelPreview.Controls.Clear();
            _panelPreview.Controls.Add(_guideVertical);
            _panelPreview.Controls.Add(_guideHorizontal);
            _panelPreview.Controls.Add(_guideVerticalCenter);
            _panelPreview.Controls.Add(_guideHorizontalCenter);
            _selectedPreviewField = null;
            HideGuides();

            var top = 10;
            foreach (var place in _places.Values.OrderBy(place => place.PlaceNumber))
            {
                var control = new CmrFieldControl(place, _currentTemplate)
                {
                    Size = new Size(MenuFieldWidth, MenuFieldHeight),
                    Location = new Point(10, top),
                    Tag = place.FieldName
                };
                control.LayoutSettingsChanged += (_, _) => OnFieldLayoutChanged(control);
                _fieldControls[place.FieldName] = control;
                _panelFields.Controls.Add(control);

                var previewControl = CreatePreviewFieldControl(place);
                _previewFieldControls[place.FieldName] = previewControl;
                _panelPreview.Controls.Add(previewControl);
                previewControl.BringToFront();

                top += MenuFieldHeight + MenuFieldSpacing;
            }

            _panelFields.AutoScrollMinSize = new Size(0, top + 10);
            _panelPreview.Invalidate();
        }

        private Panel CreatePreviewFieldControl(CmrField place)
        {
            var panel = new Panel
            {
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.WhiteSmoke,
                Tag = place.FieldName,
                Cursor = Cursors.SizeAll,
            };
            var label = new Label
            {
                Dock = DockStyle.Fill,
                Text = place.Description,
                Font = new Font("Arial", 7.5f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true,
                Padding = new Padding(4, 0, 4, 0),
                BackColor = Color.Transparent,
                Cursor = Cursors.SizeAll,
            };
            panel.Controls.Add(label);
            AttachPreviewDragHandlers(panel);
            AttachPreviewDragHandlers(label);
            ApplyPreviewFieldSize(panel, GetFieldSize(place));
            PositionPreviewField(panel, GetFieldPosition(place));
            return panel;
        }

        private void OnFieldLayoutChanged(CmrFieldControl control)
        {
            _fieldSizes[control.Place.FieldName] = new Size(control.GetFieldWidth(), control.GetFieldHeight());
            _currentTemplate.FontSizes[control.Place.FieldName] = control.GetFontSize();
            _currentTemplate.VerticalOffsets[control.Place.FieldName] = control.GetVerticalOffset();
            if (_previewFieldControls.TryGetValue(control.Place.FieldName, out var preview))
            {
                ApplyPreviewFieldSize(preview, _fieldSizes[control.Place.FieldName]);
                PositionPreviewField(preview, GetFieldPosition(control.Place));
                UpdateGuides(_selectedPreviewField);
            }
            GeneratePreview();
            _panelPreview.Invalidate();
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
            var control = sender as Control;
            if (control == null)
                return;

            SelectPreviewField(control.Parent == _panelPreview ? control as Panel : control.Parent as Panel);
        }

        private void PreviewField_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
                return;

            _draggedPreviewField = sender is Control child && child.Parent == _panelPreview ? child : (sender as Control)?.Parent;
            if (_draggedPreviewField == null)
                return;

            SelectPreviewField(_draggedPreviewField as Panel);
            _dragOffset = ReferenceEquals(sender, _draggedPreviewField) ? e.Location : _draggedPreviewField.PointToClient(Cursor.Position);
        }

        private void PreviewField_MouseMove(object? sender, MouseEventArgs e)
        {
            if (_draggedPreviewField == null || e.Button != MouseButtons.Left)
                return;

            var cursorPoint = _panelPreview.PointToClient(Cursor.Position);
            var canvas = _panelPreview.AutoScrollMinSize;
            var maxX = Math.Max(PreviewPadding, canvas.Width - _draggedPreviewField.Width - PreviewPadding);
            var maxY = Math.Max(PreviewPadding, canvas.Height - _draggedPreviewField.Height - PreviewPadding);
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

        private void SelectPreviewField(Panel? panel)
        {
            if (_selectedPreviewField != null)
                _selectedPreviewField.BackColor = Color.WhiteSmoke;

            _selectedPreviewField = panel;
            if (_selectedPreviewField == null)
            {
                HideGuides();
                return;
            }

            _selectedPreviewField.BackColor = Color.MistyRose;
            UpdateGuides(_selectedPreviewField);
        }

        private void UpdateGuides(Panel? previewField)
        {
            if (previewField == null)
            {
                HideGuides();
                return;
            }

            var canvas = _panelPreview.AutoScrollMinSize;
            _guideVertical.Location = new Point(previewField.Left, 0);
            _guideVertical.Height = canvas.Height;
            _guideHorizontal.Location = new Point(0, previewField.Top);
            _guideHorizontal.Width = canvas.Width;
            _guideVerticalCenter.Location = new Point(previewField.Left + (previewField.Width / 2), 0);
            _guideVerticalCenter.Height = canvas.Height;
            _guideHorizontalCenter.Location = new Point(0, previewField.Top + (previewField.Height / 2));
            _guideHorizontalCenter.Width = canvas.Width;

            _guideVertical.Visible = _guideHorizontal.Visible = _guideVerticalCenter.Visible = _guideHorizontalCenter.Visible = true;
            _guideVertical.BringToFront();
            _guideHorizontal.BringToFront();
            _guideVerticalCenter.BringToFront();
            _guideHorizontalCenter.BringToFront();
            previewField.BringToFront();
        }

        private void HideGuides()
        {
            _guideVertical.Visible = _guideHorizontal.Visible = _guideVerticalCenter.Visible = _guideHorizontalCenter.Visible = false;
        }

        private void MoveSelectedPreviewField(int dx, int dy)
        {
            if (_selectedPreviewField == null)
                return;

            var canvas = _panelPreview.AutoScrollMinSize;
            var maxX = Math.Max(PreviewPadding, canvas.Width - _selectedPreviewField.Width - PreviewPadding);
            var maxY = Math.Max(PreviewPadding, canvas.Height - _selectedPreviewField.Height - PreviewPadding);
            var newX = Math.Max(PreviewPadding, Math.Min(_selectedPreviewField.Left + dx, maxX));
            var newY = Math.Max(PreviewPadding, Math.Min(_selectedPreviewField.Top + dy, maxY));
            _selectedPreviewField.Location = new Point(newX, newY);
            SyncDocumentPositionFromPreview(_selectedPreviewField);
            UpdateGuides(_selectedPreviewField);
        }

        private void PositionPreviewField(Control previewControl, PointF documentPoint)
        {
            var previewPoint = DocumentToPreview(documentPoint.X, documentPoint.Y);
            var canvas = _panelPreview.AutoScrollMinSize;
            var maxX = Math.Max(PreviewPadding, canvas.Width - previewControl.Width - PreviewPadding);
            var maxY = Math.Max(PreviewPadding, canvas.Height - previewControl.Height - PreviewPadding);
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
            previewControl.Size = DocumentSizeToPreview(documentSize);
        }

        private Point DocumentToPreview(float x, float y)
        {
            return new Point(
                PreviewPadding + (int)Math.Round(x * PreviewScale),
                PreviewPadding + (int)Math.Round(y * PreviewScale));
        }

        private PointF PreviewToDocument(int x, int y)
        {
            return new PointF(
                Math.Max(0, (x - PreviewPadding) / PreviewScale),
                Math.Max(0, (y - PreviewPadding) / PreviewScale));
        }

        private Size DocumentSizeToPreview(Size size)
        {
            return new Size(
                Math.Max(28, (int)Math.Round(size.Width * PreviewScale)),
                Math.Max(18, (int)Math.Round(size.Height * PreviewScale)));
        }

        private PointF GetFieldPosition(CmrField place) => _fieldPositions[place.FieldName];
        private Size GetFieldSize(CmrField place) => _fieldSizes[place.FieldName];

        private void PanelPreview_Paint(object? sender, PaintEventArgs e)
        {
            if (_previewBitmap == null)
                return;

            e.Graphics.TranslateTransform(_panelPreview.AutoScrollPosition.X, _panelPreview.AutoScrollPosition.Y);
            e.Graphics.DrawImageUnscaled(_previewBitmap, 0, 0);
        }

        private void TemplateEditorForm_KeyDown(object? sender, KeyEventArgs e)
        {
            var step = e.Shift ? 5 : 1;
            switch (e.KeyCode)
            {
                case Keys.Left: MoveSelectedPreviewField(-step, 0); e.Handled = true; break;
                case Keys.Right: MoveSelectedPreviewField(step, 0); e.Handled = true; break;
                case Keys.Up: MoveSelectedPreviewField(0, -step); e.Handled = true; break;
                case Keys.Down: MoveSelectedPreviewField(0, step); e.Handled = true; break;
            }
        }

        private void BtnLoad_Click(object? sender, EventArgs e)
        {
            if (_cmbTemplates.SelectedItem is not string templateName)
            {
                MessageBox.Show(this, "Select a template to load.", "Template", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _currentTemplate = TemplateManager.LoadTemplate(templateName);
            LoadCurrentTemplateIntoWorkingState();
            GeneratePreview();
            RegenerateEditorControls();
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            var templateName = PromptForTemplateName();
            if (string.IsNullOrWhiteSpace(templateName))
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
                _currentTemplate.FieldPositions[entry.Key] = entry.Value;

            TemplateManager.SaveTemplate(_currentTemplate);
            LoadTemplates();
            _cmbTemplates.SelectedItem = templateName;
            MessageBox.Show(this, $"Template '{templateName}' saved.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnDelete_Click(object? sender, EventArgs e)
        {
            if (_cmbTemplates.SelectedItem is not string templateName)
                return;

            if (MessageBox.Show(this, $"Delete template '{templateName}'?", "Delete template", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            TemplateManager.DeleteTemplate(templateName);
            if (string.Equals(_currentTemplate.Name, templateName, StringComparison.OrdinalIgnoreCase))
                _currentTemplate = new CmrTemplate();
            LoadTemplates();
        }

        private string PromptForTemplateName()
        {
            using var form = new Form
            {
                Text = "Save Template",
                Width = 320,
                Height = 160,
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
            };
            var label = new Label { Text = "Template name:", Left = 12, Top = 20, Width = 280 };
            var textBox = new TextBox
            {
                Left = 12,
                Top = 50,
                Width = 280,
                Text = string.IsNullOrWhiteSpace(_currentTemplate.Name) ? DateTime.Now.ToString("CMR_yyyy-MM-dd_HHmm") : _currentTemplate.Name
            };
            var btnOk = new Button { Text = "OK", Left = 132, Top = 88, Width = 70, DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "Cancel", Left = 212, Top = 88, Width = 80, DialogResult = DialogResult.Cancel };
            form.Controls.AddRange(new Control[] { label, textBox, btnOk, btnCancel });
            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;
            return form.ShowDialog(this) == DialogResult.OK ? textBox.Text.Trim() : string.Empty;
        }
    }
}
