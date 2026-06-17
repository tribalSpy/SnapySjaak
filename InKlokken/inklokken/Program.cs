using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.Json;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new ClockingForm());
    }
}

internal sealed class ClockingForm : Form
{
    private static readonly string AppDirectory = AppContext.BaseDirectory;
    private static readonly string DataDirectory = Path.Combine(AppDirectory, "data");
    private static readonly string RecordsDirectory = Path.Combine(DataDirectory, "records");
    private static readonly string ExportsDirectory = Path.Combine(DataDirectory, "exports");
    private static readonly string ConfigPath = Path.Combine(DataDirectory, "config.json");
    private static readonly string DefaultEmployeeFile = Path.Combine(AppDirectory, "employees.csv");

    private readonly Dictionary<string, Employee> employees = new(StringComparer.OrdinalIgnoreCase);

    private readonly Label homeTitleLabel;
    private readonly Label statusLabel;
    private readonly Label detailLabel;
    private readonly Label fileLabel;
    private readonly TextBox scanBox;
    private readonly DataGridView historyGrid;
    private readonly Button editSelectedButton;
    private readonly Button deleteSelectedButton;
    private readonly TabControl tabs;
    private readonly DateTimePicker exportDatePicker;
    private readonly ComboBox manualEmployeeComboBox;
    private readonly DateTimePicker manualDatePicker;
    private readonly CheckBox manualStartCheckBox;
    private readonly DateTimePicker manualStartPicker;
    private readonly CheckBox manualFinishCheckBox;
    private readonly DateTimePicker manualFinishPicker;
    private readonly Dictionary<string, Employee> manualLookup = new(StringComparer.OrdinalIgnoreCase);
    private readonly System.Windows.Forms.Timer scannerFocusTimer;
    private char employeeDelimiter = ',';

    private string? employeeFilePath;

    public ClockingForm()
    {
        Directory.CreateDirectory(DataDirectory);
        Directory.CreateDirectory(RecordsDirectory);
        Directory.CreateDirectory(ExportsDirectory);

        Text = "InKlokken";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(940, 640);
        Size = new Size(1100, 760);
        BackColor = Color.FromArgb(246, 248, 251);

        var mainLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(18),
            ColumnCount = 1,
            RowCount = 3,
        };
        mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        Controls.Add(mainLayout);

        var headerPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            AutoSize = true,
        };
        headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        mainLayout.Controls.Add(headerPanel, 0, 0);

        var titleLabel = new Label
        {
            Text = "Clocking App",
            Font = new Font("Segoe UI", 24, FontStyle.Bold),
            Dock = DockStyle.Fill,
            AutoSize = true,
            TextAlign = ContentAlignment.MiddleLeft,
        };
        headerPanel.Controls.Add(titleLabel, 0, 0);

        var loadButton = new Button
        {
            Text = "Load Employee CSV",
            AutoSize = true,
            Padding = new Padding(14, 8, 14, 8),
        };
        loadButton.Click += (_, _) => ChooseEmployeeFile();
        headerPanel.Controls.Add(loadButton, 1, 0);

        fileLabel = new Label
        {
            Text = "Employee CSV: not loaded",
            ForeColor = Color.FromArgb(70, 70, 70),
            AutoSize = true,
            Padding = new Padding(0, 0, 0, 10),
        };
        mainLayout.Controls.Add(fileLabel, 0, 1);

        tabs = new TabControl
        {
            Dock = DockStyle.Fill,
        };
        mainLayout.Controls.Add(tabs, 0, 2);

        var homePage = new TabPage("Home");
        tabs.TabPages.Add(homePage);

        var homeLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(24),
            ColumnCount = 1,
            RowCount = 5,
        };
        homeLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        homeLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 76));
        homeLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        homeLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
        homeLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 84));
        homeLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));
        homePage.Controls.Add(homeLayout);

        homeTitleLabel = new Label
        {
            Text = "InKlokken",
            Font = new Font("Segoe UI", 28, FontStyle.Bold),
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
        };
        homeLayout.Controls.Add(homeTitleLabel, 0, 0);

        var statusGroup = CreateGroup("Scan Result");
        homeLayout.Controls.Add(statusGroup, 0, 1);
        var statusLayout = CreateVerticalLayout();
        statusGroup.Controls.Add(statusLayout);

        statusLabel = new Label
        {
            Text = "Ready",
            Font = new Font("Segoe UI", 34, FontStyle.Bold),
            ForeColor = Color.FromArgb(22, 80, 146),
            AutoSize = false,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
        };
        statusLayout.Controls.Add(statusLabel, 0, 0);

        detailLabel = new Label
        {
            Text = "Load your employee CSV and scan a QR code.",
            Font = new Font("Segoe UI", 20, FontStyle.Regular),
            AutoSize = false,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
        };
        statusLayout.Controls.Add(detailLabel, 0, 1);

        var scanHintLabel = new Label
        {
            Text = "Scanner stays active in the field below.",
            Font = new Font("Segoe UI", 11, FontStyle.Regular),
            AutoSize = false,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
        };
        homeLayout.Controls.Add(scanHintLabel, 0, 2);
        statusLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 78));
        statusLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        var scannerGroup = CreateGroup("Scan");
        homeLayout.Controls.Add(scannerGroup, 0, 3);
        var scannerGroupLayout = CreateVerticalLayout();
        scannerGroup.Controls.Add(scannerGroupLayout);

        scanBox = new TextBox
        {
            Font = new Font("Consolas", 26, FontStyle.Regular),
            Dock = DockStyle.Fill,
            TextAlign = HorizontalAlignment.Center,
            TabIndex = 0,
        };
        scanBox.KeyDown += ScanBoxOnKeyDown;
        scannerGroupLayout.Controls.Add(scanBox, 0, 0);

        var manualButton = new Button
        {
            Text = "Process Manual Entry",
            Dock = DockStyle.Fill,
            Height = 42,
        };
        manualButton.Click += (_, _) => ProcessScan();
        scannerGroupLayout.Controls.Add(manualButton, 0, 1);
        scannerGroupLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
        scannerGroupLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
        scannerGroupLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        var managePage = new TabPage("Records");
        tabs.TabPages.Add(managePage);

        var manageLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(6),
            ColumnCount = 2,
            RowCount = 1,
        };
        manageLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 64));
        manageLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 36));
        managePage.Controls.Add(manageLayout);

        var historyGroup = CreateGroup("Today's Scans");
        manageLayout.Controls.Add(historyGroup, 0, 0);

        historyGrid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            RowHeadersVisible = false,
        };
        historyGrid.Columns.Add("Time", "Time");
        historyGrid.Columns.Add("TBNR", "TBNR");
        historyGrid.Columns.Add("Name", "Name");
        historyGrid.Columns.Add("Type", "Type");
        historyGrid.Columns.Add("Direction", "IN/OUT");
        historyGrid.Columns.Add("Source", "Source");
        historyGroup.Controls.Add(historyGrid);

        var historyButtons = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 44,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(0, 8, 0, 0),
        };
        historyGroup.Controls.Add(historyButtons);

        editSelectedButton = new Button
        {
            Text = "Edit Selected",
            Width = 130,
            Enabled = false,
        };
        editSelectedButton.Click += (_, _) => EditSelectedRecord();
        historyButtons.Controls.Add(editSelectedButton);

        deleteSelectedButton = new Button
        {
            Text = "Delete Selected",
            Width = 130,
            Enabled = false,
        };
        deleteSelectedButton.Click += (_, _) => DeleteSelectedRecord();
        historyButtons.Controls.Add(deleteSelectedButton);

        historyGrid.SelectionChanged += (_, _) => UpdateHistoryButtons();

        var exportGroup = CreateGroup("Export Daily File");
        manageLayout.Controls.Add(exportGroup, 1, 0);
        var exportLayout = CreateVerticalLayout();
        exportGroup.Controls.Add(exportLayout);

        var exportLabel = new Label
        {
            Text = "Choose a date and export the CSV file.",
            Dock = DockStyle.Fill,
        };
        exportLayout.Controls.Add(exportLabel, 0, 3);

        exportDatePicker = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "yyyy-MM-dd",
            Dock = DockStyle.Fill,
            Font = new Font("Consolas", 14, FontStyle.Regular),
        };
        exportLayout.Controls.Add(exportDatePicker, 0, 2);

        var exportButton = new Button
        {
            Text = "Export CSV",
            Dock = DockStyle.Fill,
            Height = 42,
        };
        exportButton.Click += (_, _) => ExportDay();
        exportLayout.Controls.Add(exportButton, 0, 1);

        var openFolderButton = new Button
        {
            Text = "Open Export Folder",
            Dock = DockStyle.Fill,
            Height = 42,
        };
        openFolderButton.Click += (_, _) => OpenExportFolder();
        exportLayout.Controls.Add(openFolderButton, 0, 0);
        exportLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
        exportLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
        exportLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
        exportLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
        exportLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        var manualPage = new TabPage("Manual Corrections");
        tabs.TabPages.Add(manualPage);

        var manualLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Padding = new Padding(18),
            ColumnCount = 2,
            RowCount = 8,
            AutoSize = true,
        };
        manualLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        manualLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65));
        manualPage.Controls.Add(manualLayout);

        manualLayout.Controls.Add(new Label { Text = "Employee", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
        manualEmployeeComboBox = new ComboBox
        {
            Dock = DockStyle.Top,
            DropDownStyle = ComboBoxStyle.DropDown,
            AutoCompleteMode = AutoCompleteMode.SuggestAppend,
            AutoCompleteSource = AutoCompleteSource.ListItems,
        };
        manualLayout.Controls.Add(manualEmployeeComboBox, 1, 0);

        manualLayout.Controls.Add(new Label { Text = "Date", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 1);
        manualDatePicker = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "yyyy-MM-dd",
            Dock = DockStyle.Top,
            Value = DateTime.Today,
        };
        manualLayout.Controls.Add(manualDatePicker, 1, 1);

        manualStartCheckBox = new CheckBox
        {
            Text = "Add start time (IN)",
            Dock = DockStyle.Fill,
            Checked = true,
        };
        manualLayout.Controls.Add(manualStartCheckBox, 0, 2);
        manualStartPicker = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "HH:mm",
            ShowUpDown = true,
            Dock = DockStyle.Top,
            Value = DateTime.Today.AddHours(8),
        };
        manualLayout.Controls.Add(manualStartPicker, 1, 2);
        manualStartCheckBox.CheckedChanged += (_, _) => manualStartPicker.Enabled = manualStartCheckBox.Checked;

        manualFinishCheckBox = new CheckBox
        {
            Text = "Add finish time (OUT)",
            Dock = DockStyle.Fill,
            Checked = false,
        };
        manualLayout.Controls.Add(manualFinishCheckBox, 0, 3);
        manualFinishPicker = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "HH:mm",
            ShowUpDown = true,
            Dock = DockStyle.Top,
            Value = DateTime.Today.AddHours(17),
            Enabled = false,
        };
        manualLayout.Controls.Add(manualFinishPicker, 1, 3);
        manualFinishCheckBox.CheckedChanged += (_, _) => manualFinishPicker.Enabled = manualFinishCheckBox.Checked;

        var manualInfo = new Label
        {
            Text = "Use this page when someone forgot to scan in or out. You can add just one time or both.",
            Dock = DockStyle.Fill,
            Height = 50,
        };
        manualLayout.Controls.Add(manualInfo, 0, 4);
        manualLayout.SetColumnSpan(manualInfo, 2);

        var saveManualButton = new Button
        {
            Text = "Save Manual Correction",
            Height = 42,
            Dock = DockStyle.Top,
        };
        saveManualButton.Click += (_, _) => SaveManualCorrection();
        manualLayout.Controls.Add(saveManualButton, 0, 5);
        manualLayout.SetColumnSpan(saveManualButton, 2);

        employeeFilePath = LoadConfiguredEmployeeFile();
        LoadEmployeesIfPossible();
        RefreshTodayHistory();

        scannerFocusTimer = new System.Windows.Forms.Timer { Interval = 300 };
        scannerFocusTimer.Tick += (_, _) =>
        {
            if (tabs.SelectedTab?.Text == "Home" && Visible && Enabled && !ContainsComboBoxDropDownFocus())
            {
                if (!scanBox.Focused)
                {
                    scanBox.Focus();
                    scanBox.SelectionStart = scanBox.TextLength;
                }
            }
        };
        scannerFocusTimer.Start();

        tabs.SelectedIndexChanged += (_, _) =>
        {
            if (tabs.SelectedTab?.Text == "Home")
            {
                BeginInvoke(FocusScanner);
            }
        };
        Shown += (_, _) => FocusScanner();
    }

    private GroupBox CreateGroup(string title)
    {
        return new GroupBox
        {
            Text = title,
            Dock = DockStyle.Fill,
            Padding = new Padding(14),
            Margin = new Padding(0, 0, 12, 0),
            BackColor = Color.White,
        };
    }

    private static TableLayoutPanel CreateVerticalLayout()
    {
        return new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Margin = new Padding(0),
            Padding = new Padding(0),
        };
    }

    private void FocusScanner()
    {
        if (tabs.SelectedTab?.Text != "Home")
        {
            return;
        }

        scanBox.Focus();
        scanBox.SelectAll();
    }

    private bool ContainsComboBoxDropDownFocus()
    {
        return manualEmployeeComboBox.DroppedDown;
    }

    private string? LoadConfiguredEmployeeFile()
    {
        if (File.Exists(DefaultEmployeeFile))
        {
            return DefaultEmployeeFile;
        }

        if (!File.Exists(ConfigPath))
        {
            return null;
        }

        try
        {
            var config = JsonSerializer.Deserialize<AppConfig>(File.ReadAllText(ConfigPath));
            if (config?.EmployeeFile is { Length: > 0 } && File.Exists(config.EmployeeFile))
            {
                return config.EmployeeFile;
            }
        }
        catch
        {
            return null;
        }

        return null;
    }

    private void SaveConfig()
    {
        var json = JsonSerializer.Serialize(new AppConfig(employeeFilePath), new JsonSerializerOptions
        {
            WriteIndented = true,
        });
        File.WriteAllText(ConfigPath, json);
    }

    private void ChooseEmployeeFile()
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Choose employee CSV",
            InitialDirectory = AppDirectory,
            Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        employeeFilePath = dialog.FileName;
        SaveConfig();
        LoadEmployeesIfPossible();
    }

    private void LoadEmployeesIfPossible()
    {
        if (string.IsNullOrWhiteSpace(employeeFilePath))
        {
            fileLabel.Text = "Employee CSV: not loaded";
            SetStatus("Load employee file", "Click 'Load Employee CSV' and select the file with TBNR, type, and name.", Color.DarkGoldenrod);
            return;
        }

        try
        {
            employees.Clear();
            foreach (var employee in ReadEmployees(employeeFilePath))
            {
                employees[employee.Tbnr] = employee;
            }

            if (employees.Count == 0)
            {
                throw new InvalidOperationException("No employees found in the CSV.");
            }

            fileLabel.Text = $"Employee CSV: {employeeFilePath}";
            RefreshManualEmployeeList();
            SetStatus("Ready", "Scan a QR code.", Color.FromArgb(22, 80, 146));
        }
        catch (Exception ex)
        {
            fileLabel.Text = $"Employee CSV: {employeeFilePath}";
            SetStatus("CSV error", ex.Message, Color.Firebrick);
            MessageBox.Show(this, ex.Message, "CSV Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private IEnumerable<Employee> ReadEmployees(string path)
    {
        using var reader = new StreamReader(path, Encoding.UTF8, true);
        var headerLine = reader.ReadLine();
        if (string.IsNullOrWhiteSpace(headerLine))
        {
            throw new InvalidOperationException("The employee CSV has no header row.");
        }

        employeeDelimiter = DetectDelimiter(headerLine);
        var headers = ParseCsvLine(headerLine, employeeDelimiter);
        var indexByName = headers
            .Select((name, index) => new { Name = name.Trim().ToLowerInvariant(), Index = index })
            .ToDictionary(item => item.Name, item => item.Index);

        foreach (var required in new[] { "tbnr", "type", "name" })
        {
            if (!indexByName.ContainsKey(required))
            {
                throw new InvalidOperationException("Missing required columns. Expected TBNR, type, and name.");
            }
        }

        while (!reader.EndOfStream)
        {
            var line = reader.ReadLine();
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            if (line.Trim() == "\"")
            {
                continue;
            }

            var values = ParseCsvLine(line, employeeDelimiter);
            var tbnr = GetValue(values, indexByName["tbnr"]).Trim();
            if (string.IsNullOrWhiteSpace(tbnr))
            {
                continue;
            }

            yield return new Employee(
                tbnr.ToUpperInvariant(),
                GetValue(values, indexByName["type"]).Trim(),
                GetValue(values, indexByName["name"]).Trim());
        }
    }

    private static string GetValue(IReadOnlyList<string> values, int index) =>
        index >= 0 && index < values.Count ? values[index] : string.Empty;

    private static char DetectDelimiter(string line)
    {
        var semicolons = line.Count(ch => ch == ';');
        var commas = line.Count(ch => ch == ',');
        return semicolons > commas ? ';' : ',';
    }

    private static List<string> ParseCsvLine(string line, char delimiter = ',')
    {
        var values = new List<string>();
        var current = new StringBuilder();
        var insideQuotes = false;

        for (var i = 0; i < line.Length; i++)
        {
            var ch = line[i];
            if (ch == '"')
            {
                if (insideQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    current.Append('"');
                    i++;
                }
                else
                {
                    insideQuotes = !insideQuotes;
                }
            }
            else if (ch == delimiter && !insideQuotes)
            {
                values.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(ch);
            }
        }

        values.Add(current.ToString());
        return values;
    }

    private void ScanBoxOnKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode != Keys.Enter)
        {
            return;
        }

        e.SuppressKeyPress = true;
        ProcessScan();
    }

    private void ProcessScan()
    {
        var code = scanBox.Text.Trim().ToUpperInvariant();
        scanBox.Clear();
        FocusScanner();

        if (string.IsNullOrWhiteSpace(code))
        {
            return;
        }

        if (employees.Count == 0)
        {
            SetStatus("No employee file", "Load the employee CSV before scanning.", Color.DarkGoldenrod);
            return;
        }

        if (!employees.TryGetValue(code, out var employee))
        {
            SetStatus("Not found", $"{code} is not in the employee CSV.", Color.Firebrick);
            return;
        }

        var now = DateTime.Now;
        var direction = NextDirection(code, now.Date);
        SaveDailyRecords(now.Date, ReadDailyRecords(now.Date).Append(CreateRecord(employee, direction, now, "scanner")));
        SetStatus(employee.Name, $"{now:HH:mm:ss}  {direction}", direction == "IN" ? Color.SeaGreen : Color.DarkOrange);
        RefreshTodayHistory();
    }

    private static string DailyFile(DateTime date) => Path.Combine(RecordsDirectory, $"{date:yyyy-MM-dd}.csv");

    private static string DailyExportFile(DateTime date) => Path.Combine(ExportsDirectory, $"clocking-{date:yyyy-MM-dd}.csv");

    private string NextDirection(string tbnr, DateTime date)
    {
        var scanCount = ReadDailyRecords(date).Count(record => record.Tbnr.Equals(tbnr, StringComparison.OrdinalIgnoreCase));
        return scanCount % 2 == 0 ? "IN" : "OUT";
    }

    private DailyRecord CreateRecord(Employee employee, string direction, DateTime timestamp, string source)
    {
        return new DailyRecord(
            timestamp.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            timestamp.ToString("HH:mm:ss", CultureInfo.InvariantCulture),
            employee.Tbnr,
            employee.Name,
            employee.EmployeeType,
            direction,
            source,
            Guid.NewGuid().ToString("N"));
    }

    private static string EscapeCsv(string value)
    {
        if (!value.Contains('"') && !value.Contains(',') && !value.Contains('\n') && !value.Contains('\r'))
        {
            return value;
        }

        return $"\"{value.Replace("\"", "\"\"")}\"";
    }

    private void RefreshTodayHistory()
    {
        historyGrid.Rows.Clear();

        exportDatePicker.Value = DateTime.Today;
        var rows = ReadDailyRecords(DateTime.Today)
            .OrderByDescending(record => record.Timestamp)
            .ToList();

        foreach (var row in rows)
        {
            var index = historyGrid.Rows.Add(
                row.Time,
                row.Tbnr,
                row.Name,
                row.Type,
                row.Direction,
                row.Source);
            historyGrid.Rows[index].Tag = row;
        }

        UpdateHistoryButtons();
    }

    private void RefreshManualEmployeeList()
    {
        var selected = manualEmployeeComboBox.Text;
        manualLookup.Clear();
        var items = employees.Values
            .OrderBy(employee => employee.Name)
            .ThenBy(employee => employee.Tbnr)
            .Select(employee =>
            {
                var display = $"{employee.Name} ({employee.Tbnr})";
                manualLookup[display] = employee;
                manualLookup[employee.Name] = employee;
                manualLookup[employee.Tbnr] = employee;
                return display;
            })
            .ToArray();

        manualEmployeeComboBox.Items.Clear();
        manualEmployeeComboBox.Items.AddRange(items);
        manualEmployeeComboBox.Text = selected;
    }

    private void SaveManualCorrection()
    {
        if (employees.Count == 0)
        {
            MessageBox.Show(this, "Load the employee CSV before adding manual corrections.", "No employee file", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (!TryResolveManualEmployee(out var employee))
        {
            MessageBox.Show(this, "Choose a valid employee from the list or type a valid TBNR.", "Employee not found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!manualStartCheckBox.Checked && !manualFinishCheckBox.Checked)
        {
            MessageBox.Show(this, "Choose at least a start time or a finish time.", "No time selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var date = manualDatePicker.Value.Date;
        var newRecords = new List<DailyRecord>(ReadDailyRecords(date));

        if (manualStartCheckBox.Checked)
        {
            var start = date.Add(manualStartPicker.Value.TimeOfDay);
            newRecords.Add(CreateRecord(employee, "IN", start, "manual"));
        }

        if (manualFinishCheckBox.Checked)
        {
            var finish = date.Add(manualFinishPicker.Value.TimeOfDay);
            newRecords.Add(CreateRecord(employee, "OUT", finish, "manual"));
        }

        SaveDailyRecords(date, newRecords);
        SetStatus("Manual saved", $"{employee.Name}{Environment.NewLine}{date:yyyy-MM-dd}", Color.MediumSlateBlue);
        RefreshTodayHistory();
        MessageBox.Show(this, "Manual correction saved.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void UpdateHistoryButtons()
    {
        var hasSelection = historyGrid.SelectedRows.Count > 0 && historyGrid.SelectedRows[0].Tag is DailyRecord;
        editSelectedButton.Enabled = hasSelection;
        deleteSelectedButton.Enabled = hasSelection;
    }

    private void EditSelectedRecord()
    {
        if (historyGrid.SelectedRows.Count == 0 || historyGrid.SelectedRows[0].Tag is not DailyRecord record)
        {
            return;
        }

        if (!ShowRecordEditor(record, out var updatedRecord))
        {
            return;
        }

        var date = record.Timestamp.Date;
        var records = ReadDailyRecords(date);
        var matchIndex = records.FindIndex(existing => existing.RecordIdentity == record.RecordIdentity);
        if (matchIndex < 0)
        {
            matchIndex = records.FindIndex(existing => existing.LooseIdentity == record.LooseIdentity);
        }

        if (matchIndex < 0)
        {
            MessageBox.Show(this, "Could not find the selected record to update.", "Edit failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        records[matchIndex] = updatedRecord;
        SaveDailyRecords(date, records);
        SetStatus("Record updated", $"{updatedRecord.Name}{Environment.NewLine}{updatedRecord.Time}", Color.SteelBlue);
        RefreshTodayHistory();
    }

    private void DeleteSelectedRecord()
    {
        if (historyGrid.SelectedRows.Count == 0 || historyGrid.SelectedRows[0].Tag is not DailyRecord record)
        {
            return;
        }

        var confirm = MessageBox.Show(
            this,
            $"Delete this record?{Environment.NewLine}{record.Name} - {record.Time} - {record.Direction}",
            "Delete record",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        if (confirm != DialogResult.Yes)
        {
            return;
        }

        var date = record.Timestamp.Date;
        var records = ReadDailyRecords(date);
        var matchIndex = records.FindIndex(existing => existing.RecordIdentity == record.RecordIdentity);
        if (matchIndex < 0)
        {
            matchIndex = records.FindIndex(existing => existing.LooseIdentity == record.LooseIdentity);
        }

        if (matchIndex < 0)
        {
            MessageBox.Show(this, "Could not find the selected record to delete.", "Delete failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        records.RemoveAt(matchIndex);
        SaveDailyRecords(date, records);
        SetStatus("Record deleted", $"{record.Name}{Environment.NewLine}{record.Time}", Color.Firebrick);
        RefreshTodayHistory();
    }

    private bool ShowRecordEditor(DailyRecord record, out DailyRecord updatedRecord)
    {
        updatedRecord = record;

        using var dialog = new Form
        {
            Text = "Edit Clock Record",
            StartPosition = FormStartPosition.CenterParent,
            ClientSize = new Size(460, 240),
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(16),
            ColumnCount = 2,
            RowCount = 5,
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65));
        dialog.Controls.Add(layout);

        layout.Controls.Add(new Label { Text = "Employee", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
        var employeeBox = new ComboBox
        {
            Dock = DockStyle.Fill,
            DropDownStyle = ComboBoxStyle.DropDown,
            AutoCompleteMode = AutoCompleteMode.SuggestAppend,
            AutoCompleteSource = AutoCompleteSource.ListItems,
        };
        var employeeItems = employees.Values
            .OrderBy(employee => employee.Name)
            .ThenBy(employee => employee.Tbnr)
            .Select(employee => $"{employee.Name} ({employee.Tbnr})")
            .ToArray();
        employeeBox.Items.AddRange(employeeItems);
        employeeBox.Text = $"{record.Name} ({record.Tbnr})";
        layout.Controls.Add(employeeBox, 1, 0);

        layout.Controls.Add(new Label { Text = "Time", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 1);
        var timePicker = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "HH:mm:ss",
            ShowUpDown = true,
            Dock = DockStyle.Fill,
            Value = record.Timestamp,
        };
        layout.Controls.Add(timePicker, 1, 1);

        layout.Controls.Add(new Label { Text = "Direction", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 2);
        var directionBox = new ComboBox
        {
            Dock = DockStyle.Fill,
            DropDownStyle = ComboBoxStyle.DropDownList,
        };
        directionBox.Items.AddRange(new object[] { "IN", "OUT" });
        directionBox.SelectedItem = record.Direction.ToUpperInvariant();
        layout.Controls.Add(directionBox, 1, 2);

        layout.Controls.Add(new Label { Text = "Source", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 3);
        var sourceLabel = new Label
        {
            Dock = DockStyle.Fill,
            Text = record.Source,
            TextAlign = ContentAlignment.MiddleLeft,
        };
        layout.Controls.Add(sourceLabel, 1, 3);

        var buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false,
        };
        var saveButton = new Button { Text = "Save", DialogResult = DialogResult.OK, Width = 100 };
        var cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 100 };
        buttonPanel.Controls.Add(saveButton);
        buttonPanel.Controls.Add(cancelButton);
        layout.Controls.Add(buttonPanel, 0, 4);
        layout.SetColumnSpan(buttonPanel, 2);

        dialog.AcceptButton = saveButton;
        dialog.CancelButton = cancelButton;

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return false;
        }

        if (!TryResolveEmployeeText(employeeBox.Text.Trim(), out var employee))
        {
            MessageBox.Show(this, "Choose a valid employee name.", "Employee not found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        updatedRecord = new DailyRecord(
            record.Date,
            timePicker.Value.ToString("HH:mm:ss", CultureInfo.InvariantCulture),
            employee.Tbnr,
            employee.Name,
            employee.EmployeeType,
            directionBox.SelectedItem?.ToString() ?? record.Direction,
            record.Source,
            record.RecordId);
        return true;
    }

    private bool TryResolveManualEmployee(out Employee employee)
    {
        var raw = manualEmployeeComboBox.Text.Trim();
        return TryResolveEmployeeText(raw, out employee);
    }

    private bool TryResolveEmployeeText(string raw, out Employee employee)
    {
        if (manualLookup.TryGetValue(raw, out employee!))
        {
            return true;
        }

        var parenIndex = raw.LastIndexOf(" (", StringComparison.Ordinal);
        if (parenIndex > 0 && raw.EndsWith(")", StringComparison.Ordinal))
        {
            var namePart = raw[..parenIndex].Trim();
            if (manualLookup.TryGetValue(namePart, out employee!))
            {
                return true;
            }

            var tbnrPart = raw[(parenIndex + 2)..^1].Trim();
            if (employees.TryGetValue(tbnrPart, out employee!))
            {
                return true;
            }
        }

        return employees.TryGetValue(raw, out employee!);
    }

    private List<DailyRecord> ReadDailyRecords(DateTime date)
    {
        var file = DailyFile(date);
        var records = new List<DailyRecord>();
        if (!File.Exists(file))
        {
            return records;
        }

        foreach (var line in File.ReadLines(file, Encoding.UTF8).Skip(1))
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            var values = ParseCsvLine(line);
            var dateText = GetValue(values, 0);
            var timeText = GetValue(values, 1);
            var source = GetValue(values, 6);
            var recordId = GetValue(values, 7);
            records.Add(new DailyRecord(
                dateText,
                timeText,
                GetValue(values, 2),
                GetValue(values, 3),
                GetValue(values, 4),
                GetValue(values, 5),
                string.IsNullOrWhiteSpace(source) ? "scanner" : source,
                string.IsNullOrWhiteSpace(recordId) ? Guid.NewGuid().ToString("N") : recordId));
        }

        return records;
    }

    private void SaveDailyRecords(DateTime date, IEnumerable<DailyRecord> records)
    {
        var file = DailyFile(date);
        var sorted = records
            .OrderBy(record => record.Timestamp)
            .ThenBy(record => record.Tbnr, StringComparer.OrdinalIgnoreCase)
            .ThenBy(record => record.Direction, StringComparer.OrdinalIgnoreCase)
            .ToList();

        using var writer = new StreamWriter(file, false, Encoding.UTF8);
        writer.WriteLine("date,time,tbnr,name,type,direction,source,record_id");
        foreach (var record in sorted)
        {
            writer.WriteLine(string.Join(",",
                EscapeCsv(record.Date),
                EscapeCsv(record.Time),
                EscapeCsv(record.Tbnr),
                EscapeCsv(record.Name),
                EscapeCsv(record.Type),
                EscapeCsv(record.Direction),
                EscapeCsv(record.Source),
                EscapeCsv(record.RecordId)));
        }

        SaveDailySummaryExport(date, sorted);
    }

    private void ExportDay()
    {
        var date = exportDatePicker.Value.Date;
        var records = ReadDailyRecords(date);
        if (records.Count == 0)
        {
            MessageBox.Show(this, $"No record file found for {date:yyyy-MM-dd}.", "No file", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        SaveDailySummaryExport(date, records);
        var source = DailyExportFile(date);

        using var dialog = new SaveFileDialog
        {
            Title = "Export daily CSV",
            Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            FileName = $"clocking-{date:yyyy-MM-dd}.csv",
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        File.Copy(source, dialog.FileName, overwrite: true);
        MessageBox.Show(this, $"Saved export to:{Environment.NewLine}{dialog.FileName}", "Exported", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void OpenExportFolder()
    {
        Directory.CreateDirectory(ExportsDirectory);
        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = $"\"{ExportsDirectory}\"",
            UseShellExecute = true,
        });
    }

    private void SaveDailySummaryExport(DateTime date, IEnumerable<DailyRecord> records)
    {
        Directory.CreateDirectory(ExportsDirectory);
        var exportFile = DailyExportFile(date);
        var summaries = BuildDailySummaries(records);

        using var writer = new StreamWriter(exportFile, false, new UTF8Encoding(true));
        writer.WriteLine($"Date - {date:dd-MM-yyyy};;;");
        writer.WriteLine(";;;");
        writer.WriteLine("name;Start(IN);finish(out);total");

        foreach (var summary in summaries)
        {
            writer.WriteLine(string.Join(";",
                EscapeSemicolon(summary.Name),
                EscapeSemicolon(summary.StartText),
                EscapeSemicolon(summary.FinishText),
                EscapeSemicolon(summary.TotalText)));
        }
    }

    private static List<DailySummaryRow> BuildDailySummaries(IEnumerable<DailyRecord> records)
    {
        return records
            .GroupBy(record => record.Tbnr, StringComparer.OrdinalIgnoreCase)
            .Select(group =>
            {
                var ordered = group.OrderBy(record => record.Timestamp).ToList();
                var displayName = ordered.First().Name;
                var firstIn = ordered.FirstOrDefault(record => record.Direction.Equals("IN", StringComparison.OrdinalIgnoreCase));
                var lastOut = ordered.LastOrDefault(record => record.Direction.Equals("OUT", StringComparison.OrdinalIgnoreCase));

                DateTime? openIn = null;
                TimeSpan total = TimeSpan.Zero;
                foreach (var record in ordered)
                {
                    if (record.Direction.Equals("IN", StringComparison.OrdinalIgnoreCase))
                    {
                        openIn = record.Timestamp;
                    }
                    else if (record.Direction.Equals("OUT", StringComparison.OrdinalIgnoreCase) && openIn is DateTime start && record.Timestamp >= start)
                    {
                        total += record.Timestamp - start;
                        openIn = null;
                    }
                }

                return new DailySummaryRow(
                    displayName,
                    firstIn?.Timestamp.ToString("HH:mm", CultureInfo.InvariantCulture) ?? string.Empty,
                    lastOut?.Timestamp.ToString("HH:mm", CultureInfo.InvariantCulture) ?? string.Empty,
                    total > TimeSpan.Zero ? total.ToString(@"hh\:mm", CultureInfo.InvariantCulture) : string.Empty,
                    firstIn?.Timestamp ?? DateTime.MaxValue);
            })
            .OrderBy(summary => summary.SortTime)
            .ThenBy(summary => summary.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string EscapeSemicolon(string value)
    {
        if (!value.Contains('"') && !value.Contains(';') && !value.Contains('\n') && !value.Contains('\r'))
        {
            return value;
        }

        return $"\"{value.Replace("\"", "\"\"")}\"";
    }

    private void SetStatus(string title, string detail, Color color)
    {
        statusLabel.Text = title;
        statusLabel.ForeColor = color;
        detailLabel.Text = detail;
    }

    private sealed record Employee(string Tbnr, string EmployeeType, string Name);

    private sealed record AppConfig(string? EmployeeFile);

    private sealed record DailySummaryRow(
        string Name,
        string StartText,
        string FinishText,
        string TotalText,
        DateTime SortTime);

    private sealed record DailyRecord(
        string Date,
        string Time,
        string Tbnr,
        string Name,
        string Type,
        string Direction,
        string Source,
        string RecordId)
    {
        public DateTime Timestamp =>
            DateTime.ParseExact($"{Date} {Time}", "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

        public string RecordIdentity => RecordId;

        public string LooseIdentity => $"{Date}|{Time}|{Tbnr}|{Name}|{Type}|{Direction}|{Source}";
    }
}
