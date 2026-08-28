using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

public sealed class CollectorForm : Form
{
    private static string BrandPrefix => Environment.GetEnvironmentVariable("COLLECTOR_BRAND_PREFIX") ?? String.Empty;
    private readonly TextBox templatePath = new TextBox();
    private readonly TextBox outputPath = new TextBox();
    private readonly DateTimePicker startDate = new DateTimePicker();
    private readonly DateTimePicker endDate = new DateTimePicker();
    private readonly RadioButton allStores = new RadioButton();
    private readonly RadioButton specificStore = new RadioButton();
    private readonly TextBox storeName = new TextBox();
    private readonly Label status = new Label();

    public CollectorForm()
    {
        Text = "매출 데이터 수집 W6";
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new System.Drawing.Size(710, 415);
        Font = new System.Drawing.Font("Malgun Gothic", 9F);
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;

        var title = new Label { Text = "매출 데이터 수집 W6", Left = 24, Top = 22, Width = 500, Height = 30, Font = new System.Drawing.Font("Malgun Gothic", 16F, System.Drawing.FontStyle.Bold) };
        Controls.Add(title);
        AddLabel("원본 파일", 24, 77);
        templatePath.SetBounds(130, 73, 455, 28);
        templatePath.Text = String.Empty;
        Controls.Add(templatePath);
        var browse = new Button { Text = "찾아보기", Left = 595, Top = 72, Width = 90, Height = 30 };
        browse.Click += delegate { ChooseTemplate(); };
        Controls.Add(browse);

        AddLabel("조회 시작일", 24, 122);
        ConfigureDatePicker(startDate, 130, 118);
        Controls.Add(startDate);
        AddLabel("조회 종료일", 24, 162);
        ConfigureDatePicker(endDate, 130, 158);
        Controls.Add(endDate);
        var yesterday = DateTime.Today.AddDays(-1);
        startDate.Value = yesterday;
        endDate.Value = yesterday;

        AddLabel("대상 가맹점", 24, 207);
        allStores.Text = "전체 가맹점";
        allStores.SetBounds(130, 203, 110, 28);
        allStores.Checked = true;
        allStores.CheckedChanged += delegate { UpdateScope(); };
        Controls.Add(allStores);
        specificStore.Text = "특정 가맹점";
        specificStore.SetBounds(252, 203, 110, 28);
        specificStore.CheckedChanged += delegate { UpdateScope(); };
        Controls.Add(specificStore);

        AddLabel("매장명", 24, 247);
        storeName.SetBounds(130, 243, 455, 28);
        Controls.Add(storeName);

        AddLabel("저장 위치", 24, 292);
        outputPath.SetBounds(130, 288, 455, 28);
        outputPath.Text = Path.Combine(Application.StartupPath, "output", "sales_collection.xlsx");
        Controls.Add(outputPath);
        var chooseOutput = new Button { Text = "찾아보기", Left = 595, Top = 287, Width = 90, Height = 30 };
        chooseOutput.Click += delegate { ChooseOutput(); };
        Controls.Add(chooseOutput);

        status.SetBounds(24, 356, 460, 25);
        status.Text = "기간과 대상을 선택한 뒤 실행하세요.";
        status.ForeColor = System.Drawing.Color.FromArgb(43, 108, 176);
        Controls.Add(status);
        var execute = new Button { Text = "실행", Left = 505, Top = 346, Width = 180, Height = 36 };
        execute.Click += delegate { LaunchCollector(); };
        Controls.Add(execute);
        UpdateScope();
    }

    private void AddLabel(string text, int left, int top)
    {
        Controls.Add(new Label { Text = text, Left = left, Top = top, Width = 95, Height = 26 });
    }

    private static void ConfigureDatePicker(DateTimePicker picker, int left, int top)
    {
        picker.SetBounds(left, top, 150, 28);
        picker.Format = DateTimePickerFormat.Custom;
        picker.CustomFormat = "yyyy-MM-dd";
    }

    private void UpdateScope()
    {
        storeName.Enabled = specificStore.Checked;
    }

    private void ChooseTemplate()
    {
        using (var dialog = new OpenFileDialog { Filter = "Excel workbook (*.xlsx)|*.xlsx", FileName = templatePath.Text })
        {
            if (dialog.ShowDialog(this) == DialogResult.OK) templatePath.Text = dialog.FileName;
        }
    }

    private void ChooseOutput()
    {
        using (var dialog = new SaveFileDialog { Filter = "Excel workbook (*.xlsx)|*.xlsx", FileName = Path.GetFileName(outputPath.Text), InitialDirectory = Path.GetDirectoryName(outputPath.Text) })
        {
            if (dialog.ShowDialog(this) == DialogResult.OK) outputPath.Text = dialog.FileName;
        }
    }

    private void LaunchCollector()
    {
        if (!File.Exists(templatePath.Text))
        {
            MessageBox.Show(this, "유효한 원본 Excel 템플릿을 선택하세요.", "템플릿 확인", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        if (String.IsNullOrWhiteSpace(outputPath.Text) || !String.Equals(Path.GetExtension(outputPath.Text), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show(this, "출력 파일은 .xlsx 확장자로 지정하세요.", "출력 경로", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        if (endDate.Value.Date < startDate.Value.Date)
        {
            MessageBox.Show(this, "종료일은 시작일보다 빠를 수 없습니다.", "날짜 범위", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        if (specificStore.Checked && String.IsNullOrWhiteSpace(storeName.Text))
        {
            MessageBox.Show(this, "특정 가맹점의 정확한 매장명을 입력하세요.", "매장 선택", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var collector = Path.Combine(Application.StartupPath, "Sales_Data_Collector.exe");
        if (!File.Exists(collector))
        {
            MessageBox.Show(this, "Sales_Data_Collector.exe를 GUI EXE와 같은 폴더에 두세요.", "수집기 확인", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        var command = new StringBuilder();
        command.Append(Quote(collector)).Append(" --production-write --template ").Append(Quote(templatePath.Text)).Append(" --production-output ").Append(Quote(outputPath.Text));
        for (var day = startDate.Value.Date; day <= endDate.Value.Date; day = day.AddDays(1)) command.Append(" --date ").Append(day.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
        if (allStores.Checked) command.Append(" --all-stores");
        else command.Append(" --store-name ").Append(Quote(BrandPrefix + storeName.Text.Trim()));

        try
        {
            var launcher = Path.Combine(Application.StartupPath, "run_sales_data_collector.cmd");
            var script = "@echo off\r\n" + command + "\r\n"
                + "if errorlevel 1 (\r\n"
                + "  echo.\r\n"
                + "  echo [ERROR] Collection failed. Closing this window in 10 seconds.\r\n"
                + "  timeout /t 10 /nobreak >nul\r\n"
                + "  exit\r\n"
                + ")\r\n"
                + "echo.\r\n"
                + "echo [OK] Collection completed. Press any key to close this window.\r\n"
                + "pause >nul\r\n"
                + "exit\r\n";
            File.WriteAllText(launcher, script, Encoding.Default);
            Process.Start(new ProcessStartInfo { FileName = "cmd.exe", Arguments = "/k run_sales_data_collector.cmd", WorkingDirectory = Application.StartupPath, UseShellExecute = true });
            status.Text = "CMD 창에서 로그인 후 드라이런과 기록을 진행 중입니다.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "실행 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static string Quote(string value) { return "\"" + value.Replace("\"", "\\\"") + "\""; }

    [STAThread]
    public static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new CollectorForm());
    }
}
