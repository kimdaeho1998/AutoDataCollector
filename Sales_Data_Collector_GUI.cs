using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

public sealed class CollectorForm : Form
{
    private const string MenuOutputDirectory = @"C:\Users\USER\Desktop\WorkSpace\DATA\Menu";

    private const string ServiceConfigFileName = "collector_service.conf";

    private static string BrandPrefix
    {
        get
        {
            return Environment.GetEnvironmentVariable("COLLECTOR_BRAND_PREFIX") ?? String.Empty;
        }
    }

    private readonly TextBox salesTemplatePath = new TextBox();
    private readonly TextBox salesOutputPath = new TextBox();
    private readonly DateTimePicker salesStartDate = new DateTimePicker();
    private readonly DateTimePicker salesEndDate = new DateTimePicker();
    private readonly RadioButton salesAllStores = new RadioButton();
    private readonly RadioButton salesSpecificStore = new RadioButton();
    private readonly TextBox salesStoreName = new TextBox();
    private readonly Label salesStatus = new Label();

    private readonly TextBox menuTemplatePath = new TextBox();
    private readonly TextBox menuOutputPath = new TextBox();
    private readonly NumericUpDown menuYear = new NumericUpDown();
    private readonly TextBox menuMonth = new TextBox();
    private readonly RadioButton menuAllStores = new RadioButton();
    private readonly RadioButton menuSpecificStore = new RadioButton();
    private readonly TextBox menuStoreName = new TextBox();
    private readonly Label menuStatus = new Label();
    private readonly ServiceConfig serviceConfig;

    public CollectorForm()
    {
        serviceConfig = ServiceConfig.Load(Path.Combine(Application.StartupPath, ServiceConfigFileName));

        Text = "MagicERP 매출 자동수집 W7";
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new System.Drawing.Size(760, 500);
        Font = new System.Drawing.Font("Malgun Gothic", 9F);
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;

        var title = new Label
        {
            Text = "MagicERP 매출 자동수집 W7",
            Left = 24,
            Top = 22,
            Width = 500,
            Height = 32,
            Font = new System.Drawing.Font("Malgun Gothic", 16F, System.Drawing.FontStyle.Bold)
        };
        Controls.Add(title);

        var tabs = new TabControl { Left = 20, Top = 66, Width = 720, Height = 410 };
        var salesTab = new TabPage("매출 수집");
        var menuTab = new TabPage("메뉴 분석");
        tabs.TabPages.Add(salesTab);
        tabs.TabPages.Add(menuTab);
        Controls.Add(tabs);

        BuildSalesTab(salesTab);
        BuildMenuTab(menuTab);
    }

    private void BuildSalesTab(Control parent)
    {
        AddLabel(parent, "원본 파일", 24, 28);
        salesTemplatePath.SetBounds(130, 24, 455, 28);
        parent.Controls.Add(salesTemplatePath);
        AddButton(parent, "찾아보기", 595, 23, delegate { ChooseExcel(salesTemplatePath); });

        AddLabel(parent, "조회 시작일", 24, 73);
        ConfigureDatePicker(salesStartDate, 130, 69);
        parent.Controls.Add(salesStartDate);
        AddLabel(parent, "조회 종료일", 24, 113);
        ConfigureDatePicker(salesEndDate, 130, 109);
        parent.Controls.Add(salesEndDate);
        var yesterday = DateTime.Today.AddDays(-1);
        salesStartDate.Value = yesterday;
        salesEndDate.Value = yesterday;

        AddLabel(parent, "대상 가맹점", 24, 158);
        salesAllStores.Text = "전체 가맹점";
        salesAllStores.SetBounds(130, 154, 110, 28);
        salesAllStores.Checked = true;
        salesAllStores.CheckedChanged += delegate { UpdateSalesScope(); };
        parent.Controls.Add(salesAllStores);
        salesSpecificStore.Text = "특정 가맹점";
        salesSpecificStore.SetBounds(252, 154, 120, 28);
        salesSpecificStore.CheckedChanged += delegate { UpdateSalesScope(); };
        parent.Controls.Add(salesSpecificStore);

        AddLabel(parent, "매장명", 24, 198);
        salesStoreName.SetBounds(130, 194, 455, 28);
        parent.Controls.Add(salesStoreName);

        AddLabel(parent, "저장 위치", 24, 243);
        salesOutputPath.SetBounds(130, 239, 455, 28);
        salesOutputPath.Text = Path.Combine(Application.StartupPath, "output", "sales_collection.xlsx");
        parent.Controls.Add(salesOutputPath);
        AddButton(parent, "찾아보기", 595, 238, delegate { ChooseOutput(salesOutputPath, "sales_collection.xlsx"); });

        salesStatus.SetBounds(24, 316, 490, 25);
        salesStatus.Text = "CMD 창에서 로그인 후 매출 수집과 기록을 진행합니다.";
        salesStatus.ForeColor = System.Drawing.Color.FromArgb(43, 108, 176);
        parent.Controls.Add(salesStatus);
        AddButton(parent, "실행", 505, 306, delegate { LaunchSalesCollector(); }, 180, 36);
        UpdateSalesScope();
    }

    private void BuildMenuTab(Control parent)
    {
        AddLabel(parent, "원본 파일", 24, 28);
        menuTemplatePath.SetBounds(130, 24, 455, 28);
        parent.Controls.Add(menuTemplatePath);
        AddButton(parent, "찾아보기", 595, 23, delegate
        {
            ChooseExcel(menuTemplatePath);
            TryDetectYearMonthFromPath(menuTemplatePath.Text);
        });

        AddLabel(parent, "연도", 24, 73);
        menuYear.SetBounds(130, 69, 100, 28);
        menuYear.Minimum = 2000;
        menuYear.Maximum = 2099;
        menuYear.DecimalPlaces = 0;
        menuYear.ThousandsSeparator = false;
        menuYear.TextAlign = HorizontalAlignment.Left;
        menuYear.UpDownAlign = LeftRightAlignment.Right;
        menuYear.Value = DateTime.Today.Year;
        parent.Controls.Add(menuYear);

        AddLabel(parent, "월", 252, 73);
        menuMonth.SetBounds(300, 69, 80, 28);
        menuMonth.TextAlign = HorizontalAlignment.Left;
        menuMonth.Text = DateTime.Today.Month.ToString(CultureInfo.InvariantCulture);
        parent.Controls.Add(menuMonth);

        AddLabel(parent, "대상 가맹점", 24, 118);
        menuAllStores.Text = "전체 가맹점";
        menuAllStores.SetBounds(130, 114, 110, 28);
        menuAllStores.Checked = true;
        menuAllStores.CheckedChanged += delegate { UpdateMenuScope(); };
        parent.Controls.Add(menuAllStores);
        menuSpecificStore.Text = "특정 가맹점";
        menuSpecificStore.SetBounds(252, 114, 120, 28);
        menuSpecificStore.CheckedChanged += delegate { UpdateMenuScope(); };
        parent.Controls.Add(menuSpecificStore);

        AddLabel(parent, "매장명", 24, 158);
        menuStoreName.SetBounds(130, 154, 455, 28);
        parent.Controls.Add(menuStoreName);

        AddLabel(parent, "저장 파일", 24, 203);
        menuOutputPath.SetBounds(130, 199, 420, 28);
        menuOutputPath.Text = Path.Combine(MenuOutputDirectory, "26년_상품별데이터_통합_07월.xlsx");
        parent.Controls.Add(menuOutputPath);
        AddButton(parent, "찾아보기", 595, 198, delegate { ChooseOutput(menuOutputPath, "menu_monthly_collection.xlsx"); });

        menuStatus.SetBounds(24, 316, 430, 25);
        menuStatus.Text = "CMD 창에서 로그인 후 월별 메뉴 분석과 기록을 진행합니다.";
        menuStatus.ForeColor = System.Drawing.Color.FromArgb(43, 108, 176);
        parent.Controls.Add(menuStatus);
        AddButton(parent, "실행", 505, 306, delegate { LaunchMenuCollector(); }, 180, 36);
        UpdateMenuScope();
    }

    private static void AddLabel(Control parent, string text, int left, int top)
    {
        var width = text.Trim().Length <= 2 ? 35 : 95;
        parent.Controls.Add(new Label { Text = text, Left = left, Top = top, Width = width, Height = 26 });
    }

    private static void AddButton(Control parent, string text, int left, int top, EventHandler click, int width = 90, int height = 30)
    {
        var button = new Button { Text = text, Left = left, Top = top, Width = width, Height = height };
        button.Click += click;
        parent.Controls.Add(button);
    }

    private static void ConfigureDatePicker(DateTimePicker picker, int left, int top)
    {
        picker.SetBounds(left, top, 150, 28);
        picker.Format = DateTimePickerFormat.Custom;
        picker.CustomFormat = "yyyy-MM-dd";
    }

    private void UpdateSalesScope()
    {
        salesStoreName.Enabled = salesSpecificStore.Checked;
    }

    private void UpdateMenuScope()
    {
        menuStoreName.Enabled = menuSpecificStore.Checked;
    }

    private void ChooseExcel(TextBox target)
    {
        using (var dialog = new OpenFileDialog { Filter = "Excel workbook (*.xlsx)|*.xlsx", FileName = target.Text })
        {
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                target.Text = dialog.FileName;
            }
        }
    }

    private void ChooseOutput(TextBox target, string defaultName)
    {
        var initialDirectory = Path.GetDirectoryName(target.Text);
        if (String.IsNullOrWhiteSpace(initialDirectory) || !Directory.Exists(initialDirectory))
        {
            initialDirectory = Application.StartupPath;
        }

        using (var dialog = new SaveFileDialog
        {
            Filter = "Excel workbook (*.xlsx)|*.xlsx",
            FileName = String.IsNullOrWhiteSpace(Path.GetFileName(target.Text)) ? defaultName : Path.GetFileName(target.Text),
            InitialDirectory = initialDirectory,
            OverwritePrompt = true
        })
        {
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                target.Text = dialog.FileName;
            }
        }
    }

    private void TryDetectYearMonthFromPath(string path)
    {
        var fileName = Path.GetFileName(path) ?? String.Empty;
        var match = Regex.Match(fileName, @"(?<yy>\d{2})년\s*(?<mm>\d{1,2})월");
        if (!match.Success)
        {
            return;
        }

        int yy;
        int mm;
        if (!Int32.TryParse(match.Groups["yy"].Value, out yy) || !Int32.TryParse(match.Groups["mm"].Value, out mm))
        {
            return;
        }
        if (mm < 1 || mm > 12)
        {
            return;
        }

        menuYear.Value = 2000 + yy;
        menuMonth.Text = mm.ToString(CultureInfo.InvariantCulture);
    }

    private void LaunchSalesCollector()
    {
        try
        {
            if (!ValidateExcelInput(salesTemplatePath.Text, salesOutputPath.Text))
            {
                return;
            }
            if (salesEndDate.Value.Date < salesStartDate.Value.Date)
            {
                MessageBox.Show(this, "조회 종료일은 조회 시작일보다 빠를 수 없습니다.", "날짜 범위", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (salesSpecificStore.Checked && String.IsNullOrWhiteSpace(salesStoreName.Text))
            {
                MessageBox.Show(this, "특정 가맹점을 선택한 경우 매장명을 입력하세요.", "매장 선택", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var command = new StringBuilder();
            command.Append(Quote(CollectorPath()));
            AppendServiceArguments(command);
            command.Append(" --production-write --template ").Append(Quote(salesTemplatePath.Text));
            command.Append(" --production-output ").Append(Quote(salesOutputPath.Text));
            for (var day = salesStartDate.Value.Date; day <= salesEndDate.Value.Date; day = day.AddDays(1))
            {
                command.Append(" --date ").Append(day.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
            }
            if (salesAllStores.Checked)
            {
                command.Append(" --all-stores");
            }
            else
            {
                command.Append(" --store-name ").Append(Quote(BrandPrefix + salesStoreName.Text.Trim()));
            }

            LaunchCommand(command.ToString(), salesStatus);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "실행 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void LaunchMenuCollector()
    {
        try
        {
            var selectedMonth = ParseMenuMonth();

            Directory.CreateDirectory(
                MenuOutputDirectory
            );

            string outputFileName = Path.GetFileName(
                menuOutputPath.Text.Trim()
            );

            if (
                String.IsNullOrWhiteSpace(outputFileName)
                || !String.Equals(
                    Path.GetExtension(outputFileName),
                    ".xlsx",
                    StringComparison.OrdinalIgnoreCase
                )
            )
            {
                string sourceBaseName = Path.GetFileNameWithoutExtension(
                    menuTemplatePath.Text.Trim()
                );

                if (String.IsNullOrWhiteSpace(sourceBaseName))
                {
                    sourceBaseName = "menu_monthly_collection";
                }

                outputFileName = String.Format(
                    CultureInfo.InvariantCulture,
                    "{0}_{1:00}\uC6D4.xlsx",
                    sourceBaseName,
                    selectedMonth
                );
            }

            menuOutputPath.Text = Path.Combine(
                MenuOutputDirectory,
                outputFileName
            );

            if (
                !ValidateExcelInput(
                    menuTemplatePath.Text,
                    menuOutputPath.Text
                )
            )
            {
                return;
            }

            if (
                menuSpecificStore.Checked
                && String.IsNullOrWhiteSpace(
                    menuStoreName.Text
                )
            )
            {
                MessageBox.Show(
                    this,
                    "특정 가맹점을 선택한 경우 매장명을 입력하세요.",
                    "매장 선택",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );

                return;
            }

            var command = new StringBuilder();

            command.Append(
                Quote(
                    CollectorPath()
                )
            );

            AppendServiceArguments(
                command
            );

            command.Append(
                " --menu-monthly-batch"
            );

            command.Append(
                " --template "
            ).Append(
                Quote(
                    menuTemplatePath.Text
                )
            );

            command.Append(
                " --output "
            ).Append(
                Quote(
                    menuOutputPath.Text
                )
            );

            command.Append(
                " --year "
            ).Append(
                ((int)menuYear.Value).ToString(
                    CultureInfo.InvariantCulture
                )
            );

            command.Append(
                " --month "
            ).Append(
                selectedMonth.ToString(
                    CultureInfo.InvariantCulture
                )
            );

            if (
                menuAllStores.Checked
            )
            {
                command.Append(
                    " --all-stores"
                );
            }
            else
            {
                command.Append(
                    " --store-name "
                ).Append(
                    Quote(
                        BrandPrefix
                        + menuStoreName.Text.Trim()
                    )
                );
            }

            LaunchCommand(
                command.ToString(),
                menuStatus
            );
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                this,
                ex.Message,
                "실행 실패",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }

    private bool ValidateExcelInput(string template, string output)
    {
        if (!File.Exists(template) || !String.Equals(Path.GetExtension(template), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show(this, "유효한 원본 Excel 파일(.xlsx)을 선택하세요.", "원본 파일 확인", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
        if (String.IsNullOrWhiteSpace(output) || !String.Equals(Path.GetExtension(output), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show(this, "저장 파일은 .xlsx 확장자로 지정하세요.", "저장 파일 확인", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
        if (String.Equals(Path.GetFullPath(template), Path.GetFullPath(output), StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show(this, "원본 파일과 저장 파일은 같을 수 없습니다.", "저장 파일 확인", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
        return true;
    }

    private int ParseMenuMonth()
    {
        int value;
        if (!Int32.TryParse(menuMonth.Text.Trim(), out value) || value < 1 || value > 12)
        {
            throw new InvalidOperationException("메뉴 분석 월은 1부터 12 사이 숫자로 입력하세요.");
        }
        return value;
    }

    private static string CollectorPath()
    {
        var collector = Path.Combine(Application.StartupPath, "MagicERP_AutoCollector.exe");
        if (!File.Exists(collector))
        {
            throw new FileNotFoundException("MagicERP_AutoCollector.exe를 GUI EXE와 같은 폴더에 두세요.", collector);
        }
        return collector;
    }

    private void LaunchCommand(string command, Label statusLabel)
    {
        try
        {
            var launcher = Path.Combine(Application.StartupPath, "run_magic_erp_collector.cmd");
            var script = "@echo off\r\n"
                + "chcp 949 >nul\r\n"
                + BuildServiceEnvironmentScript()
                + command + "\r\n"
                + "if errorlevel 1 (\r\n"
                + "  echo.\r\n"
                + "  echo [ERROR] Collection failed. Press any key to close this window.\r\n"
                + "  pause >nul\r\n"
                + "  exit\r\n"
                + ")\r\n"
                + "echo.\r\n"
                + "echo [OK] Collection completed. Press any key to close this window.\r\n"
                + "pause >nul\r\n"
                + "exit\r\n";
            File.WriteAllText(launcher, script, Encoding.Default);
            Process.Start(new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = "/k run_magic_erp_collector.cmd",
                WorkingDirectory = Application.StartupPath,
                UseShellExecute = true
            });
            statusLabel.Text = "CMD 창에서 로그인 후 수집과 기록을 진행합니다.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "실행 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static string Quote(string value)
    {
        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }

    private void AppendServiceArguments(StringBuilder command)
    {
        if (!serviceConfig.HasRequiredCore)
        {
            throw new InvalidOperationException(
                ServiceConfigFileName + "에 COLLECTOR_BASE_URL, COLLECTOR_BRAND_IDX, COLLECTOR_BRAND_NAME 값을 설정하세요.");
        }

        command.Append(" --base-url ").Append(Quote(serviceConfig.BaseUrl));
        command.Append(" --brand-idx ").Append(Quote(serviceConfig.BrandIdx));
        command.Append(" --brand-name ").Append(Quote(serviceConfig.BrandName));
    }

    private string BuildServiceEnvironmentScript()
    {
        var script = new StringBuilder();
        foreach (var item in serviceConfig.Values)
        {
            if (String.IsNullOrWhiteSpace(item.Value))
            {
                continue;
            }
            script.Append("set \"").Append(item.Key).Append("=").Append(EscapeSetValue(item.Value)).Append("\"\r\n");
        }
        return script.ToString();
    }

    private static string EscapeSetValue(string value)
    {
        return value.Replace("\"", "");
    }

    [STAThread]
    public static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new CollectorForm());
    }
}

public sealed class ServiceConfig
{
    public readonly System.Collections.Generic.Dictionary<string, string> Values =
        new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    public string BaseUrl
    {
        get { return Get("COLLECTOR_BASE_URL"); }
    }

    public string BrandIdx
    {
        get { return Get("COLLECTOR_BRAND_IDX"); }
    }

    public string BrandName
    {
        get { return Get("COLLECTOR_BRAND_NAME"); }
    }

    public bool HasRequiredCore
    {
        get
        {
            return !String.IsNullOrWhiteSpace(BaseUrl)
                && !String.IsNullOrWhiteSpace(BrandIdx)
                && !String.IsNullOrWhiteSpace(BrandName);
        }
    }

    public static ServiceConfig Load(string path)
    {
        var config = new ServiceConfig();
        if (!File.Exists(path))
        {
            return config;
        }

        foreach (var rawLine in File.ReadAllLines(path, Encoding.UTF8))
        {
            var line = rawLine.Trim();
            if (line.Length == 0 || line.StartsWith("#"))
            {
                continue;
            }

            var index = line.IndexOf('=');
            if (index <= 0)
            {
                continue;
            }

            var key = line.Substring(0, index).Trim();
            var value = line.Substring(index + 1).Trim();
            if (key.StartsWith("COLLECTOR_", StringComparison.OrdinalIgnoreCase))
            {
                config.Values[key] = value;
            }
        }

        return config;
    }

    private string Get(string key)
    {
        string value;
        return Values.TryGetValue(key, out value) ? value : String.Empty;
    }
}
