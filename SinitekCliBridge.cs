using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

public static class SinitekCliBridge
{
    private const string AddinDir = @"C:\Sinitek\SinitekExcelAddin";
    private const string AddinDll = @"C:\Sinitek\SinitekExcelAddin\SinitekExcel.dll";
    private const int StockSearchTimeoutMs = 15000;
    private const string DirectorySheetName = "\u76ee\u5f55";

    private static readonly CurrencyUnitSpec[] CurrencyUnits = new[]
    {
        new CurrencyUnitSpec("1", 1D, "\u5143", "\u80a1"),
        new CurrencyUnitSpec("0.001", 0.001D, "\u5343\u5143", "\u5343\u80a1"),
        new CurrencyUnitSpec("0.0001", 0.0001D, "\u4e07\u5143", "\u4e07\u80a1"),
        new CurrencyUnitSpec("0.000001", 0.000001D, "\u767e\u4e07\u5143", "\u767e\u4e07\u80a1"),
        new CurrencyUnitSpec("0.00000001", 0.00000001D, "\u4ebf\u5143", "\u4ebf\u80a1")
    };

    private static readonly PredictionIndicatorDefinition[] PredictionIndicators = new[]
    {
        new PredictionIndicatorDefinition("sales", 22, "\u7a0e\u91d1\u53ca\u9644\u52a0/\u8425\u4e1a\u6536\u5165", "\u7a0e\u91d1\u53ca\u9644\u52a0\u7387", "\u7a0e\u91d1\u53ca\u9644\u52a0\u5360\u6536\u5165\u6bd4"),
        new PredictionIndicatorDefinition("sales", 25, "\u9500\u552e\u8d39\u7528/\u8425\u4e1a\u6536\u5165", "\u9500\u552e\u8d39\u7528\u7387", "\u9500\u552e\u8d39\u7528\u5360\u6536\u5165\u6bd4"),
        new PredictionIndicatorDefinition("sales", 28, "\u7ba1\u7406\u8d39\u7528/\u8425\u4e1a\u6536\u5165", "\u7ba1\u7406\u8d39\u7528\u7387", "\u7ba1\u7406\u8d39\u7528\u5360\u6536\u5165\u6bd4"),
        new PredictionIndicatorDefinition("sales", 31, "\u7814\u53d1\u8d39\u7528/\u8425\u4e1a\u6536\u5165", "\u7814\u53d1\u8d39\u7528\u7387", "\u7814\u53d1\u8d39\u7528\u5360\u6536\u5165\u6bd4"),
        new PredictionIndicatorDefinition("sales", 34, "\u5176\u4ed6\u6536\u76ca/\u8425\u4e1a\u6536\u5165", "\u5176\u4ed6\u6536\u76ca\u7387", "\u5176\u4ed6\u6536\u76ca\u5360\u6536\u5165\u6bd4"),
        new PredictionIndicatorDefinition("sales", 37, "\u6295\u8d44\u6536\u76ca/\u8425\u4e1a\u6536\u5165", "\u6295\u8d44\u6536\u76ca\u7387", "\u6295\u8d44\u6536\u76ca\u5360\u6536\u5165\u6bd4"),
        new PredictionIndicatorDefinition("sales", 40, "\u8d44\u4ea7\u5904\u7f6e\u6536\u76ca/\u8425\u4e1a\u6536\u5165", "\u8d44\u4ea7\u5904\u7f6e\u6536\u76ca\u7387", "\u8d44\u4ea7\u5904\u7f6e\u6536\u76ca\u5360\u6536\u5165\u6bd4"),
        new PredictionIndicatorDefinition("sales", 42, "\u8425\u4e1a\u5916\u6536\u5165"),
        new PredictionIndicatorDefinition("sales", 43, "\u8425\u4e1a\u5916\u652f\u51fa"),
        new PredictionIndicatorDefinition("sales", 48, "\u6240\u5f97\u7a0e\u7a0e\u7387", "\u6240\u5f97\u7a0e\u7387", "\u6709\u6548\u6240\u5f97\u7a0e\u7387"),
        new PredictionIndicatorDefinition("sales", 51, "\u5c11\u6570\u80a1\u4e1c\u635f\u76ca/\u51c0\u5229\u6da6", "\u5c11\u6570\u80a1\u4e1c\u635f\u76ca\u7387", "\u5c11\u6570\u80a1\u4e1c\u635f\u76ca\u5360\u51c0\u5229\u6da6\u6bd4"),
        new PredictionIndicatorDefinition("capital", 11, "\u5e94\u6536\u7968\u636e\u5468\u8f6c\u5929\u6570", "\u5e94\u6536\u7968\u636e\u5468\u8f6c\u5929\u6570\uff08\u76f8\u5bf9\u4e8e\u8425\u4e1a\u6536\u5165\uff09", "\u5e94\u6536\u7968\u636e/\u8425\u4e1a\u6536\u5165"),
        new PredictionIndicatorDefinition("capital", 15, "\u5e94\u6536\u8d26\u6b3e\u5468\u8f6c\u5929\u6570", "\u5e94\u6536\u8d26\u6b3e\u5468\u8f6c\u5929\u6570\uff08\u76f8\u5bf9\u4e8e\u8425\u4e1a\u6536\u5165\uff09", "\u5e94\u6536\u8d26\u6b3e/\u8425\u4e1a\u6536\u5165"),
        new PredictionIndicatorDefinition("capital", 19, "\u5e94\u6536\u6b3e\u9879\u878d\u8d44/\u8425\u4e1a\u6536\u5165", "\u5e94\u6536\u6b3e\u9879\u878d\u8d44\u5360\u6536\u5165\u6bd4", "\u5e94\u6536\u6b3e\u9879\u878d\u8d44\u6bd4\u8425\u4e1a\u6536\u5165"),
        new PredictionIndicatorDefinition("capital", 22, "\u9884\u4ed8\u8d26\u6b3e/\u8425\u4e1a\u6210\u672c", "\u9884\u4ed8\u8d26\u6b3e\u5360\u6210\u672c\u6bd4", "\u9884\u4ed8\u8d26\u6b3e\u6bd4\u8425\u4e1a\u6210\u672c"),
        new PredictionIndicatorDefinition("capital", 25, "\u5176\u4ed6\u5e94\u6536\u6b3e\u5468\u8f6c\u5929\u6570", "\u5176\u4ed6\u5e94\u6536\u6b3e\u5468\u8f6c\u5929\u6570\uff08\u76f8\u5bf9\u4e8e\u8425\u4e1a\u6536\u5165\uff09", "\u5176\u4ed6\u5e94\u6536\u6b3e/\u8425\u4e1a\u6536\u5165"),
        new PredictionIndicatorDefinition("capital", 29, "\u5b58\u8d27\u5468\u8f6c\u5929\u6570", "\u5b58\u8d27\u5468\u8f6c\u5929\u6570\uff08\u76f8\u5bf9\u4e8e\u8425\u4e1a\u6210\u672c\uff09", "\u5b58\u8d27/\u8425\u4e1a\u6210\u672c"),
        new PredictionIndicatorDefinition("capital", 33, "\u5408\u540c\u8d44\u4ea7/\u8425\u4e1a\u6536\u5165", "\u5408\u540c\u8d44\u4ea7\u5360\u6536\u5165\u6bd4", "\u5408\u540c\u8d44\u4ea7\u6bd4\u8425\u4e1a\u6536\u5165"),
        new PredictionIndicatorDefinition("capital", 39, "\u5e94\u4ed8\u7968\u636e\u5468\u8f6c\u5929\u6570", "\u5e94\u4ed8\u7968\u636e\u5468\u8f6c\u5929\u6570\uff08\u76f8\u5bf9\u4e8e\u8425\u4e1a\u6210\u672c\uff09", "\u5e94\u4ed8\u7968\u636e/\u8425\u4e1a\u6210\u672c"),
        new PredictionIndicatorDefinition("capital", 43, "\u5e94\u4ed8\u8d26\u6b3e\u5468\u8f6c\u5929\u6570", "\u5e94\u4ed8\u8d26\u6b3e\u5468\u8f6c\u5929\u6570\uff08\u76f8\u5bf9\u4e8e\u8425\u4e1a\u6210\u672c\uff09", "\u5e94\u4ed8\u8d26\u6b3e/\u8425\u4e1a\u6210\u672c"),
        new PredictionIndicatorDefinition("capital", 47, "\u9884\u6536\u8d26\u6b3e/\u8425\u4e1a\u6536\u5165", "\u9884\u6536\u8d26\u6b3e\u5360\u6536\u5165\u6bd4", "\u9884\u6536\u8d26\u6b3e\u6bd4\u8425\u4e1a\u6536\u5165"),
        new PredictionIndicatorDefinition("capital", 50, "\u5408\u540c\u8d1f\u503a/\u8425\u4e1a\u6536\u5165", "\u5408\u540c\u8d1f\u503a\u5360\u6536\u5165\u6bd4", "\u5408\u540c\u8d1f\u503a\u6bd4\u8425\u4e1a\u6536\u5165"),
        new PredictionIndicatorDefinition("capital", 53, "\u5e94\u4ed8\u804c\u5de5\u85aa\u916c/\u8425\u4e1a\u6210\u672c", "\u5e94\u4ed8\u804c\u5de5\u85aa\u916c\u5360\u6210\u672c\u6bd4", "\u5e94\u4ed8\u804c\u5de5\u85aa\u916c\u6bd4\u8425\u4e1a\u6210\u672c"),
        new PredictionIndicatorDefinition("capital", 56, "\u5e94\u4ea4\u7a0e\u8d39/\u8425\u4e1a\u6536\u5165", "\u5e94\u4ea4\u7a0e\u8d39\u5360\u6536\u5165\u6bd4", "\u5e94\u4ea4\u7a0e\u8d39\u6bd4\u8425\u4e1a\u6536\u5165"),
        new PredictionIndicatorDefinition("capital", 59, "\u5176\u4ed6\u5e94\u4ed8\u6b3e/\u8425\u4e1a\u6210\u672c", "\u5176\u4ed6\u5e94\u4ed8\u6b3e\u5360\u6210\u672c\u6bd4", "\u5176\u4ed6\u5e94\u4ed8\u6b3e\u6bd4\u8425\u4e1a\u6210\u672c"),
    };

    [DllImport("user32.dll")]
    private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool VirtualProtect(IntPtr lpAddress, UIntPtr dwSize, uint flNewProtect, out uint lpflOldProtect);

    private static bool documentPropertyPatchInstalled;

    private sealed class PerfTracer
    {
        private readonly bool enabled;
        private readonly string scope;
        private readonly Stopwatch stopwatch;
        private long lastMilliseconds;

        public PerfTracer(bool enabled, string scope)
        {
            this.enabled = enabled;
            this.scope = string.IsNullOrWhiteSpace(scope) ? "cli" : scope;
            if (enabled)
            {
                stopwatch = Stopwatch.StartNew();
                Mark("start");
            }
        }

        public void Mark(string phase)
        {
            if (!enabled)
            {
                return;
            }

            long totalMilliseconds = stopwatch.ElapsedMilliseconds;
            long elapsedMilliseconds = totalMilliseconds - lastMilliseconds;
            lastMilliseconds = totalMilliseconds;
            Console.Error.WriteLine("PERF Phase=" + scope + "." + phase
                + " ElapsedMs=" + elapsedMilliseconds.ToString(CultureInfo.InvariantCulture)
                + " TotalMs=" + totalMilliseconds.ToString(CultureInfo.InvariantCulture));
        }
    }

    public static string Inspect(string workbookPath, bool visible, bool perfTrace)
    {
        var tracer = new PerfTracer(perfTrace, "inspect");
        using (var session = new ExcelSession(workbookPath, visible, true, tracer))
        {
            var lines = new List<string>();
            lines.Add("Workbook=" + session.Workbook.FullName);
            lines.Add("ExcelVersion=" + session.App.Version);
            lines.Add("PluginVersion=" + Safe(() => SinitekExcel.WriterUtil.ModelUtil.PluginVer));
            lines.Add("ModelVersion=" + Safe(() => SinitekExcel.WriterUtil.LoadUtil.getModelVersion()));
            lines.Add("ModelType=" + Safe(() => SinitekExcel.WriterUtil.LoadUtil.getModelType().ToString(CultureInfo.InvariantCulture)));
            lines.Add("WorkbookModelVersion=" + GetOfficeDocProperty("ModelVersion"));
            lines.Add("WorkbookModelType=" + GetOfficeDocProperty("ModelType"));
            lines.Add("ModelUrl=" + Safe(() => SinitekExcel.WriterUtil.ModelUtil.ModelUrl));
            lines.Add("LoginState=" + Safe(() => SinitekExcel.WriterUtil.WriteUtil.LoginState));
            lines.Add("UserName=" + Safe(() => SinitekExcel.WriterUtil.WriteUtil.UserName));
            lines.Add("StkCode=" + GetDocProperty("StkCode"));
            lines.Add("GSCode=" + GetDocProperty("GSCode"));
            lines.Add("HistoryYear=" + GetDocProperty("HistoryYear"));
            lines.Add("ForecastYear=" + GetDocProperty("ForecastYear"));
            lines.Add("CurrencyUnit=" + GetDocProperty("CurrencyUnit"));
            lines.Add("Sheets=" + session.Workbook.Worksheets.Count);
            tracer.Mark("complete");
            return string.Join(Environment.NewLine, lines);
        }
    }

    public static string Login(string username, string password, bool perfTrace)
    {
        var tracer = new PerfTracer(perfTrace, "login");
        LoadIpConfig();
        tracer.Mark("ip-config");
        if (string.IsNullOrWhiteSpace(username) || string.IsNullOrEmpty(password))
        {
            throw new ArgumentException("Username and password are required for login.");
        }

        string message = string.Empty;
        string result = SinitekExcel.WriterBLL.WriterLoginBLL.Instance.Login(username, password, ref message);
        if (string.Equals(result, "200", StringComparison.Ordinal))
        {
            SinitekExcel.WriterUtil.WriteUtil.UserName = username.Trim();
        }

        var lines = new List<string>();
        lines.Add("LoginResult=" + result);
        lines.Add("LoginState=" + Safe(() => SinitekExcel.WriterUtil.WriteUtil.LoginState));
        lines.Add("UserName=" + Safe(() => SinitekExcel.WriterUtil.WriteUtil.UserName));
        if (!string.IsNullOrWhiteSpace(message))
        {
            lines.Add("Message=" + message);
        }

        tracer.Mark("complete");
        return string.Join(Environment.NewLine, lines);
    }

    public static string OutputDirect(
        string workbookPath,
        string outWorkbook,
        bool saveOriginal,
        bool visible,
        string stockCode,
        string currencyUnit,
        string username,
        string password,
        bool perfTrace)
    {
        var tracer = new PerfTracer(perfTrace, "output");
        using (var session = new ExcelSession(workbookPath, visible, false, tracer))
        {
            LoginIfProvided(username, password);
            tracer.Mark("login");
            EnsureLoggedIn();
            tracer.Mark("ensure-logged-in");
            CurrencyUnitSpec currency = ApplyCurrencyUnit(session.Workbook, currencyUnit);
            SinitekExcel.WriterUtil.ModelUtil.CurrencyUnit = currency.Scale;
            SetDocProperty("CurrencyUnit", currency.ScaleText);
            tracer.Mark("currency");

            string modelStock = string.IsNullOrWhiteSpace(stockCode) ? GetDocProperty("StkCode") : stockCode.Trim();
            if (string.IsNullOrWhiteSpace(modelStock))
            {
                throw new InvalidOperationException("Stock code is required. Pass -Stock or run update first.");
            }

            session.App.CalculateFull();
            tracer.Mark("calculate-full");
            string suffix = ResolveMailSuffix(username);
            string handlerType = ResolveHandlerType("btnOutput");
            object handler = Activator.CreateInstance(ResolveType(handlerType));
            bool ok = (bool)InvokeMethod(handler, "GetOutputSheet", true, suffix, modelStock);
            tracer.Mark("output-sheet");
            SaveIfRequested(session.Workbook, outWorkbook, saveOriginal);
            tracer.Mark("save");
            tracer.Mark("complete");
            return "OutputDirect=" + ok
                + Environment.NewLine + "Handler=" + handlerType
                + Environment.NewLine + "MailSuffix=" + suffix
                + Environment.NewLine + "Stock=" + modelStock
                + Environment.NewLine + "CurrencyUnit=" + currency.CurrencyUnitName
                + Environment.NewLine + "CurrencyUnitScale=" + currency.ScaleText;
        }
    }

    public static string UpdateDirect(
        string workbookPath,
        string outWorkbook,
        bool saveOriginal,
        bool visible,
        string stockCode,
        string gsdm,
        string stockName,
        int historyYear,
        int forecastYear,
        string currencyUnit,
        string segmentDimension,
        string peerStock,
        bool updateDirectory,
        bool updateSrcData,
        bool migrate,
        bool addOutput,
        string username,
        string password,
        bool perfTrace)
    {
        var tracer = new PerfTracer(perfTrace, "update");
        using (var session = new ExcelSession(workbookPath, visible, false, tracer))
        {
            LoginIfProvided(username, password);
            tracer.Mark("login");
            EnsureLoggedIn();
            tracer.Mark("ensure-logged-in");
            UpdateExecutionResult result = RunUpdateDirectInSession(session, new UpdateOptions
            {
                WorkbookPath = workbookPath,
                OutWorkbook = outWorkbook,
                SaveOriginal = saveOriginal,
                OutputActionName = "Update",
                StockCode = NormalizeStockInputForSearch(stockCode),
                Gsdm = gsdm,
                StockName = stockName,
                HistoryYear = historyYear,
                ForecastYear = forecastYear,
                CurrencyUnit = currencyUnit,
                SegmentDimension = segmentDimension,
                PeerStock = peerStock,
                UpdateDirectory = updateDirectory,
                UpdateSrcData = updateSrcData,
                Migrate = migrate,
                AddOutput = addOutput,
                Username = username
            }, tracer);
            tracer.Mark("complete");
            return result.OutputText;
        }
    }

    public static string ProduceBatchDirect(
        string workbookPath,
        string outputDir,
        bool visible,
        string stocks,
        int historyYear,
        int forecastYear,
        string currencyUnit,
        string segmentDimension,
        string peerStock,
        bool updateDirectory,
        bool updateSrcData,
        bool migrate,
        string username,
        string password,
        bool perfTrace)
    {
        var tracer = new PerfTracer(perfTrace, "batch-produce");
        List<string> stockList = ParseStockList(stocks);
        if (stockList.Count == 0)
        {
            throw new ArgumentException("Batch produce requires at least one stock in -Stocks.");
        }
        if (string.IsNullOrWhiteSpace(outputDir))
        {
            throw new ArgumentException("Batch produce requires outputDir.");
        }

        var lines = new List<string>();
        lines.Add("BatchProduce=true");
        lines.Add("BatchCount=" + stockList.Count.ToString(CultureInfo.InvariantCulture));

        using (var applicationSession = new ExcelApplicationSession(visible, tracer))
        {
            LoginIfProvided(username, password);
            tracer.Mark("batch-login");

            for (int i = 0; i < stockList.Count; i++)
            {
                string inputStock = stockList[i];
                string normalizedStock = NormalizeStockInputForSearch(inputStock);
                tracer.Mark("item-" + (i + 1).ToString(CultureInfo.InvariantCulture) + "-start");
                try
                {
                    using (var session = new ExcelSession(applicationSession, workbookPath, false, tracer))
                    {
                        EnsureLoggedIn();
                        tracer.Mark("item-" + (i + 1).ToString(CultureInfo.InvariantCulture) + "-ensure-logged-in");
                        UpdateExecutionResult result = RunUpdateDirectInSession(session, new UpdateOptions
                        {
                            WorkbookPath = workbookPath,
                            OutputDir = outputDir,
                            OutputActionName = "Produce",
                            StockCode = normalizedStock,
                            HistoryYear = historyYear,
                            ForecastYear = forecastYear,
                            CurrencyUnit = currencyUnit,
                            SegmentDimension = segmentDimension,
                            PeerStock = peerStock,
                            UpdateDirectory = updateDirectory,
                            UpdateSrcData = updateSrcData,
                            Migrate = migrate,
                            AddOutput = true,
                            Username = username
                        }, tracer);

                        lines.Add("Item=" + (i + 1).ToString(CultureInfo.InvariantCulture));
                        lines.Add("InputStock=" + inputStock);
                        lines.Add(result.OutputText);
                        if (!string.IsNullOrWhiteSpace(result.Artifact))
                        {
                            lines.Add("Artifact=" + result.Artifact);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Batch produce failed for stock '" + inputStock + "': " + ex.Message, ex);
                }
                tracer.Mark("item-" + (i + 1).ToString(CultureInfo.InvariantCulture) + "-complete");
            }

            tracer.Mark("complete");
            return string.Join(Environment.NewLine, lines);
        }
    }

    private static UpdateExecutionResult RunUpdateDirectInSession(
        ExcelSession session,
        UpdateOptions options,
        PerfTracer tracer)
    {
        var stock = ResolveStock(options.StockCode, options.Gsdm, options.StockName);
        tracer.Mark("resolve-stock");
        string artifact = options.OutWorkbook;
        if (string.IsNullOrWhiteSpace(artifact) && !string.IsNullOrWhiteSpace(options.OutputDir))
        {
            artifact = BuildOutputWorkbookPath(options.OutputDir, options.WorkbookPath, options.OutputActionName, stock.StockCode);
        }
        string previousGsdm = GetDocProperty("GSCode");
        PeerStockSelection peerSelection = ResolvePeerStockSelection(stock, options.PeerStock, previousGsdm);
        tracer.Mark("resolve-peer-stock");
        string handlerType = ResolveHandlerType("btnUpdate");
        object handler = Activator.CreateInstance(ResolveType(handlerType));
        tracer.Mark("handler");

        SinitekExcel.WriterUtil.ModelUtil.StockCode = stock.StockCode;
        SinitekExcel.WriterUtil.ModelUtil.StockName = stock.StockName;
        SinitekExcel.WriterUtil.ModelUtil.Gsdm = stock.Gsdm;
        SinitekExcel.WriterUtil.ModelUtil.HistoryYear = options.HistoryYear;
        ApplyYearSettings(session.Workbook, options.HistoryYear, options.ForecastYear);
        CurrencyUnitSpec currency = ApplyCurrencyUnit(session.Workbook, options.CurrencyUnit);
        SinitekExcel.WriterUtil.ModelUtil.CurrencyUnit = currency.Scale;
        SinitekExcel.WriterUtil.ModelUtil.UpdateCheck = true;

        SetDocProperty("StkCode", stock.StockCode);
        SetDocProperty("GSCode", stock.Gsdm);
        SetDocProperty("HistoryYear", options.HistoryYear.ToString(CultureInfo.InvariantCulture));
        SetDocProperty("ForecastYear", options.ForecastYear.ToString(CultureInfo.InvariantCulture));
        SetDocProperty("CurrencyUnit", currency.ScaleText);
        var dimensionToPlugin = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
        dimensionToPlugin.Add("industry", new[] { "1", "\u6309\u884c\u4e1a" });
        dimensionToPlugin.Add("product",  new[] { "2", "\u6309\u4ea7\u54c1" });
        dimensionToPlugin.Add("region",   new[] { "3", "\u6309\u5730\u533a" });
        string segmentDimension = options.SegmentDimension;
        if (string.IsNullOrWhiteSpace(segmentDimension))
        {
            segmentDimension = "product";
        }
        string[] mapped;
        if (!dimensionToPlugin.TryGetValue(segmentDimension, out mapped))
        {
            mapped = dimensionToPlugin["product"];
        }
        SetDocProperty("CompanyManagementType", mapped[0]);
        SetDocProperty("CompanyManagementName", mapped[1]);
        SetDocProperty("PeerStock", peerSelection.Gsdms);
        SetDocProperty("UpdateDirectory", options.UpdateDirectory.ToString(CultureInfo.InvariantCulture));
        SetDocProperty("UpdateSrcData", options.UpdateSrcData.ToString(CultureInfo.InvariantCulture));
        tracer.Mark("set-properties");

        var failures = new List<string>();
        int lastYear = 0;
        string lastReportDate = string.Empty;
        Excel.XlCalculation oldCalculation = session.App.Calculation;
        try
        {
            session.App.Calculation = Excel.XlCalculation.xlCalculationManual;
            tracer.Mark("calculation-manual");

            if (IsModernHongKongHistoryHandler(handler))
            {
                RunModernHongKongUpdate(session, handler, stock, options.UpdateDirectory, options.UpdateSrcData, out lastYear, out lastReportDate, failures, tracer);
            }
            else
            {
                RunStandardUpdate(session, handler, stock, options.UpdateDirectory, options.UpdateSrcData, out lastYear, out lastReportDate, failures, tracer);
            }

            SinitekExcel.WriterUtil.ModelUtil.LastYear = lastYear;

            int migrateYears = 0;
            string previousYear = GetDocProperty("year1");
            if (!string.IsNullOrWhiteSpace(previousYear))
            {
                SetDocProperty("lastUpdateYear", previousYear);
                migrateYears = lastYear - SafeInt(previousYear, lastYear);
            }
            SetDocProperty("year1", lastYear.ToString(CultureInfo.InvariantCulture));
            tracer.Mark("year-state");

            if (options.Migrate && migrateYears > 0)
            {
                if (!(bool)InvokeMethod(handler, "MigrateData", false, migrateYears))
                {
                    failures.Add("migrate");
                }
                tracer.Mark("migrate");
            }

            if (options.AddOutput)
            {
                string outputType = ResolveHandlerType("btnOutput");
                object outputHandler = Activator.CreateInstance(ResolveType(outputType));
                if (!(bool)InvokeMethod(outputHandler, "GetOutputSheet", true, ResolveMailSuffix(options.Username), stock.StockCode))
                {
                    failures.Add("output");
                }
                tracer.Mark("output-sheet");
            }

            session.App.CalculateFull();
            tracer.Mark("calculate-full");
        }
        finally
        {
            session.App.Calculation = oldCalculation;
        }

        SaveIfRequested(session.Workbook, artifact, options.SaveOriginal);
        tracer.Mark("save");

        var lines = new List<string>();
        lines.Add("Update=" + (failures.Count == 0));
        lines.Add("Handler=" + handlerType);
        lines.Add("Stock=" + stock.StockCode);
        lines.Add("GSCode=" + stock.Gsdm);
        lines.Add("LastReportDate=" + lastReportDate);
        lines.Add("LastYear=" + lastYear.ToString(CultureInfo.InvariantCulture));
        lines.Add("HistoryYear=" + options.HistoryYear.ToString(CultureInfo.InvariantCulture));
        lines.Add("ForecastYear=" + options.ForecastYear.ToString(CultureInfo.InvariantCulture));
        lines.Add("CurrencyUnit=" + currency.CurrencyUnitName);
        lines.Add("CurrencyUnitScale=" + currency.ScaleText);
        lines.Add("PeerStock=" + peerSelection.Gsdms);
        lines.Add("PeerStockSource=" + peerSelection.Source);
        lines.Add("PeerStockCount=" + peerSelection.Count.ToString(CultureInfo.InvariantCulture));
        if (failures.Count > 0)
        {
            lines.Add("Failures=" + string.Join(",", failures));
        }

        return new UpdateExecutionResult(string.Join(Environment.NewLine, lines), artifact);
    }

    public static string PredictionSettingsDirect(
        string workbookPath,
        string outWorkbook,
        bool saveOriginal,
        bool visible,
        string predictionScope,
        string predictionRows,
        string predictionIndicators,
        string predictionMethod,
        string predictionSettings,
        bool perfTrace)
    {
        var tracer = new PerfTracer(perfTrace, "predict");
        using (var session = new ExcelSession(workbookPath, visible, false, tracer))
        {
            string handlerType = ResolveHandlerType("btnSet");
            Type formType = ResolvePredictionFormType(handlerType);
            tracer.Mark("form-type");
            var applied = new List<string>();
            var appliedIndicators = new List<string>();

            object form = Activator.CreateInstance(formType);
            try
            {
                InvokePredictionFormLoad(form);
                tracer.Mark("form-load");

                var checkBoxes = new List<System.Windows.Forms.CheckBox>();
                var comboBoxes = new List<System.Windows.Forms.ComboBox>();
                CollectControls((System.Windows.Forms.Control)form, checkBoxes);
                CollectControls((System.Windows.Forms.Control)form, comboBoxes);
                tracer.Mark("collect-controls");

                var targets = BuildPredictionTargets(checkBoxes, comboBoxes);
                var selected = ResolvePredictionSelections(
                    targets,
                    predictionScope,
                    predictionRows,
                    predictionIndicators,
                    predictionMethod,
                    predictionSettings);
                tracer.Mark("resolve-selections");

                foreach (KeyValuePair<PredictionTarget, int> item in selected)
                {
                    PredictionTarget target = item.Key;
                    int methodIndex = item.Value;

                    SetDocProperty(target.ComboName, methodIndex.ToString(CultureInfo.InvariantCulture));
                    InvokePredictionFormula(form, target, methodIndex);
                    applied.Add(target.ComboName + "=" + PredictionMethodName(methodIndex));
                    if (!string.IsNullOrWhiteSpace(target.IndicatorName))
                    {
                        appliedIndicators.Add(target.IndicatorName + "=" + PredictionMethodName(methodIndex));
                    }
                }
                tracer.Mark("apply-settings");
            }
            finally
            {
                IDisposable disposable = form as IDisposable;
                if (disposable != null)
                {
                    disposable.Dispose();
                }
            }

            session.App.CalculateFull();
            tracer.Mark("calculate-full");
            SaveIfRequested(session.Workbook, outWorkbook, saveOriginal);
            tracer.Mark("save");

            var lines = new List<string>();
            lines.Add("PredictionSettings=true");
            lines.Add("Handler=" + handlerType);
            lines.Add("Form=" + formType.FullName);
            lines.Add("AppliedCount=" + applied.Count.ToString(CultureInfo.InvariantCulture));
            lines.Add("Applied=" + string.Join(";", applied.ToArray()));
            if (appliedIndicators.Count > 0)
            {
                lines.Add("AppliedIndicators=" + string.Join(";", appliedIndicators.ToArray()));
            }
            tracer.Mark("complete");
            return string.Join(Environment.NewLine, lines);
        }
    }

    private static void RunStandardUpdate(
        ExcelSession session,
        object handler,
        Sinitek.WriterModel.StockData stock,
        bool updateDirectory,
        bool updateSrcData,
        out int lastYear,
        out string lastReportDate,
        List<string> failures,
        PerfTracer tracer)
    {
        var reportYear = DateTime.Now.Year - 1;
        int lastYearFlag = (int)InvokeMethod(handler, "IsLastYear", false, stock.StockCode, reportYear);
        tracer.Mark("standard-is-last-year");
        if (lastYearFlag == -1)
        {
            throw new InvalidOperationException("Failed to validate latest annual report period.");
        }

        lastYear = lastYearFlag > 0 ? reportYear : reportYear - 1;
        lastReportDate = (string)InvokeMethod(handler, "GetLastReportDate", false, stock.StockCode);
        tracer.Mark("standard-last-report-date");
        if (string.IsNullOrWhiteSpace(lastReportDate))
        {
            throw new InvalidOperationException("Failed to get latest report date for " + stock.StockCode + ".");
        }

        if (updateDirectory)
        {
            if (!(bool)InvokeMethod(handler, "UpdateDirectory", false, lastReportDate))
            {
                failures.Add("directory");
            }
            tracer.Mark("standard-update-directory");
        }

        if (updateSrcData)
        {
            SetDocProperty("base_stkcode", stock.StockCode);
            if (!(bool)InvokeMethod(handler, "UpdateSrcData", false, lastReportDate))
            {
                failures.Add("src-data");
            }
            tracer.Mark("standard-update-src-data");
            if (!(bool)InvokeMethod(handler, "UpdateFinancialNotes", false))
            {
                failures.Add("financial-notes");
            }
            tracer.Mark("standard-update-financial-notes");
            if (SheetExists(session.Workbook, "Company operating data") || SheetExists(session.Workbook, "公司经营数据"))
            {
                if (!(bool)InvokeMethod(handler, "UpdateCompanyManagement", false))
                {
                    failures.Add("company-management");
                }
                tracer.Mark("standard-update-company-management");
            }
            if (!(bool)InvokeMethod(handler, "UpdatePeerAnalysis", false, lastReportDate))
            {
                failures.Add("peer-analysis");
            }
            tracer.Mark("standard-update-peer-analysis");
        }
    }

    private static void RunModernHongKongUpdate(
        ExcelSession session,
        object handler,
        Sinitek.WriterModel.StockData stock,
        bool updateDirectory,
        bool updateSrcData,
        out int lastYear,
        out string lastReportDate,
        List<string> failures,
        PerfTracer tracer)
    {
        object periods = InvokeMethod(handler, "GetReportPeriodList", false, stock.StockCode);
        ArrayList reportPeriodList = periods as ArrayList;
        tracer.Mark("hk-report-period-list");
        if (reportPeriodList == null || reportPeriodList.Count == 0)
        {
            throw new InvalidOperationException("Failed to get Hong Kong report period list for " + stock.StockCode + ".");
        }

        Hashtable latestPeriod = InvokeMethod(handler, "GetlastReportPeriod", false, stock.StockCode) as Hashtable;
        Hashtable lastYearPeriod = InvokeMethod(handler, "GetLastYearReportPeriod", false, stock.StockCode) as Hashtable;
        tracer.Mark("hk-last-periods");
        lastReportDate = GetHashtableText(latestPeriod, "REPORTDATE", "reportDate", "ReportDate");
        string lastYearReportDate = GetHashtableText(lastYearPeriod, "REPORTDATE", "reportDate", "ReportDate");
        string lastYearText = GetHashtableText(lastYearPeriod, "YEAR", "year", "Year");
        if (string.IsNullOrWhiteSpace(lastYearText) && !string.IsNullOrWhiteSpace(lastYearReportDate) && lastYearReportDate.Length >= 4)
        {
            lastYearText = lastYearReportDate.Substring(0, 4);
        }
        if (string.IsNullOrWhiteSpace(lastReportDate)
            || string.IsNullOrWhiteSpace(lastYearReportDate)
            || !int.TryParse(lastYearText, NumberStyles.Integer, CultureInfo.InvariantCulture, out lastYear))
        {
            throw new InvalidOperationException("Failed to resolve Hong Kong report periods for " + stock.StockCode + ".");
        }

        if (updateDirectory)
        {
            if (!(bool)InvokeMethod(handler, "UpdateDirectory", false, lastYearReportDate, lastReportDate))
            {
                failures.Add("directory");
            }
            tracer.Mark("hk-update-directory");
        }

        if (updateSrcData)
        {
            SetDocProperty("base_stkcode", stock.StockCode);
            if (!(bool)InvokeMethod(handler, "UpdateSrcData", false, reportPeriodList, latestPeriod))
            {
                failures.Add("src-data");
            }
            tracer.Mark("hk-update-src-data");
        }
    }

    private static bool IsModernHongKongHistoryHandler(object handler)
    {
        if (handler == null)
        {
            return false;
        }

        Type type = handler.GetType();
        return HasInstanceMethod(type, "GetReportPeriodList", new[] { typeof(string) })
            && HasInstanceMethod(type, "GetlastReportPeriod", new[] { typeof(string) })
            && HasInstanceMethod(type, "GetLastYearReportPeriod", new[] { typeof(string) })
            && HasInstanceMethod(type, "UpdateDirectory", new[] { typeof(string), typeof(string) })
            && HasInstanceMethod(type, "UpdateSrcData", new[] { typeof(ArrayList), typeof(Hashtable) });
    }

    private static bool HasInstanceMethod(Type type, string name, Type[] parameterTypes)
    {
        return type.GetMethod(
            name,
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
            null,
            parameterTypes,
            null) != null;
    }

    private static string GetHashtableText(Hashtable table, params string[] keys)
    {
        if (table == null)
        {
            return string.Empty;
        }

        foreach (string key in keys)
        {
            if (table.ContainsKey(key) && table[key] != null)
            {
                return Convert.ToString(table[key], CultureInfo.InvariantCulture) ?? string.Empty;
            }
        }

        foreach (object key in table.Keys)
        {
            string keyText = Convert.ToString(key, CultureInfo.InvariantCulture);
            foreach (string expected in keys)
            {
                if (string.Equals(keyText, expected, StringComparison.OrdinalIgnoreCase) && table[key] != null)
                {
                    return Convert.ToString(table[key], CultureInfo.InvariantCulture) ?? string.Empty;
                }
            }
        }

        return string.Empty;
    }

    private static string ResolveHandlerType(string buttonId)
    {
        string modelVersion = ResolveCurrentModelVersion();
        int modelType = SafeInt(ResolveCurrentModelType(), -1);
        string typeKey = ModelTypeKey(modelType);

        Assembly asm = Assembly.LoadFrom(AddinDll);
        using (var stream = asm.GetManifestResourceStream("SinitekExcel.Resources.activesetting.xml"))
        {
            if (stream == null)
            {
                throw new InvalidOperationException("Cannot find embedded activesetting.xml.");
            }

            XDocument doc = XDocument.Load(stream);
            XElement setting = null;
            foreach (XElement candidate in doc.Root.Elements("setting"))
            {
                if ((string)candidate.Attribute("type") == typeKey)
                {
                    setting = candidate;
                    break;
                }
            }
            if (setting == null)
            {
                throw new InvalidOperationException("No mapping setting for model type " + typeKey + ".");
            }

            XElement mapping = null;
            foreach (XElement candidate in setting.Elements("mapping"))
            {
                if ((string)candidate.Attribute("version") == modelVersion)
                {
                    mapping = candidate;
                    break;
                }
            }
            if (mapping == null)
            {
                throw new InvalidOperationException("No handler mapping for model version " + modelVersion + ".");
            }

            XElement item = null;
            foreach (XElement candidate in mapping.Elements("item"))
            {
                if ((string)candidate.Attribute("name") == buttonId)
                {
                    item = candidate;
                    break;
                }
            }
            if (item == null)
            {
                throw new InvalidOperationException("No handler mapping for button " + buttonId + " in " + modelVersion + ".");
            }

            return item.Value.Trim();
        }
    }

    private static string ModelTypeKey(int modelType)
    {
        switch (modelType)
        {
            case 0: return "model";
            case 1: return "hkModel";
            case 2: return "usaModel";
            case 3: return "bkModel";
            case 4: return "orgiHkModel";
            default: throw new InvalidOperationException("Unsupported model type: " + modelType);
        }
    }

    private static Type ResolvePredictionFormType(string handlerType)
    {
        string candidate;
        if (handlerType.IndexOf(".USABLL.", StringComparison.Ordinal) >= 0)
        {
            string ns = handlerType.Substring(0, handlerType.LastIndexOf('.'));
            candidate = ns + ".Form.USPredictedDataForm";
        }
        else if (handlerType.IndexOf(".HKBLL.", StringComparison.Ordinal) >= 0
            || handlerType.IndexOf(".OrgiHKBLL.", StringComparison.Ordinal) >= 0)
        {
            string ns = handlerType.Substring(0, handlerType.LastIndexOf('.'));
            candidate = ns + ".Form.HKPredictedDataForm";
        }
        else if (handlerType.IndexOf(".BKBLL.", StringComparison.Ordinal) >= 0)
        {
            string ns = handlerType.Substring(0, handlerType.LastIndexOf('.'));
            candidate = ns + ".Form.APredictedDataForm";
        }
        else if (handlerType.IndexOf(".ASBLL.", StringComparison.Ordinal) >= 0)
        {
            string ns = handlerType.Substring(0, handlerType.LastIndexOf('.'));
            candidate = ns + ".Form.APredictedDataForm";
        }
        else
        {
            throw new InvalidOperationException("Unsupported prediction handler mapping: " + handlerType + ".");
        }

        Type type = Type.GetType(candidate + ", SinitekExcel")
            ?? Assembly.LoadFrom(AddinDll).GetType(candidate);
        if (type != null)
        {
            return type;
        }

        throw new InvalidOperationException("Cannot resolve prediction form type for handler " + handlerType + ".");
    }

    private static void InvokePredictionFormLoad(object form)
    {
        MethodInfo method = FindInstanceMethod(form.GetType(), new[] { "LoginForm_Load", "Form_Load" }, 2);
        if (method == null)
        {
            return;
        }

        InvokeMethodInfo(form, method, new object[] { form, EventArgs.Empty });
    }

    private static void InvokePredictionFormula(object form, PredictionTarget target, int methodIndex)
    {
        MethodInfo method = form.GetType().GetMethod(
            "HandleIndicatorFormula",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
            null,
            new[] { typeof(int), typeof(int), typeof(string) },
            null);
        if (method == null)
        {
            method = form.GetType().GetMethod(
                "Combobox_SelectedIndexChanged",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                null,
                new[] { typeof(int), typeof(int), typeof(string) },
                null);
        }
        if (method == null)
        {
            throw new MissingMethodException(form.GetType().FullName, "HandleIndicatorFormula");
        }

        InvokeMethodInfo(form, method, new object[] { methodIndex, target.Row, ResolvePredictionSheetName(target) });
    }

    private static MethodInfo FindInstanceMethod(Type type, string[] names, int parameterCount)
    {
        BindingFlags flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
        foreach (string name in names)
        {
            foreach (MethodInfo method in type.GetMethods(flags))
            {
                if (string.Equals(method.Name, name, StringComparison.Ordinal)
                    && method.GetParameters().Length == parameterCount)
                {
                    return method;
                }
            }
        }

        return null;
    }

    private static void InvokeMethodInfo(object target, MethodInfo method, object[] args)
    {
        try
        {
            method.Invoke(target, args);
        }
        catch (TargetInvocationException ex)
        {
            Exception inner = ex.InnerException ?? ex;
            throw new InvalidOperationException("Invocation failed: "
                + target.GetType().FullName + "." + method.Name + ": " + inner.Message, inner);
        }
    }

    private static void CollectControls<T>(System.Windows.Forms.Control root, List<T> result)
        where T : System.Windows.Forms.Control
    {
        foreach (System.Windows.Forms.Control child in root.Controls)
        {
            T typed = child as T;
            if (typed != null)
            {
                result.Add(typed);
            }

            CollectControls(child, result);
        }
    }

    private static List<PredictionTarget> BuildPredictionTargets(
        List<System.Windows.Forms.CheckBox> checkBoxes,
        List<System.Windows.Forms.ComboBox> comboBoxes)
    {
        var comboNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (System.Windows.Forms.ComboBox comboBox in comboBoxes)
        {
            comboNames.Add(comboBox.Name);
        }

        var targets = new List<PredictionTarget>();
        foreach (System.Windows.Forms.CheckBox checkBox in checkBoxes)
        {
            string prefix;
            int row;
            if (!TryParsePredictionCheckBoxName(checkBox.Name, out prefix, out row))
            {
                continue;
            }

            string comboName = prefix + "cmb_" + row.ToString(CultureInfo.InvariantCulture);
            if (!comboNames.Contains(comboName))
            {
                continue;
            }

            string scope = PredictionScopeFromPrefix(prefix);
            targets.Add(new PredictionTarget(prefix, scope, row, comboName, ResolvePredictionIndicatorName(scope, row)));
        }

        if (targets.Count == 0)
        {
            throw new InvalidOperationException("No prediction checkboxes and comboboxes were found in the plugin form.");
        }

        return targets;
    }

    private static Dictionary<PredictionTarget, int> ResolvePredictionSelections(
        List<PredictionTarget> targets,
        string predictionScope,
        string predictionRows,
        string predictionIndicators,
        string predictionMethod,
        string predictionSettings)
    {
        var selected = new Dictionary<PredictionTarget, int>();
        bool hasBulkMethod = !string.IsNullOrWhiteSpace(predictionMethod);
        bool hasSettings = !string.IsNullOrWhiteSpace(predictionSettings);
        if (!hasBulkMethod && !hasSettings)
        {
            throw new ArgumentException("Pass -PredictionMethod for bulk setting, or -PredictionSettings for explicit setting.");
        }

        if (hasBulkMethod)
        {
            int methodIndex = ResolvePredictionMethodIndex(predictionMethod);
            HashSet<string> scopes = ParsePredictionScopes(predictionScope);
            HashSet<int> rows = ParsePredictionRows(predictionRows);
            HashSet<PredictionTarget> indicatorTargets = ResolvePredictionIndicatorTargets(targets, predictionIndicators);
            bool hasRowFilter = rows.Count > 0;
            bool hasIndicatorFilter = indicatorTargets.Count > 0;
            foreach (PredictionTarget target in targets)
            {
                bool selectedByFilter = (!hasRowFilter && !hasIndicatorFilter)
                    || (hasRowFilter && rows.Contains(target.Row))
                    || (hasIndicatorFilter && indicatorTargets.Contains(target));
                if (MatchesPredictionScope(target, scopes) && selectedByFilter)
                {
                    selected[target] = methodIndex;
                }
            }

            if (selected.Count == 0)
            {
                throw new ArgumentException("No prediction indicators matched -PredictionScope/-PredictionRows/-PredictionIndicators.");
            }
        }

        if (hasSettings)
        {
            foreach (string entry in SplitList(predictionSettings))
            {
                int equalIndex = entry.IndexOf('=');
                if (equalIndex <= 0 || equalIndex == entry.Length - 1)
                {
                    throw new ArgumentException("Prediction setting must be key=method: " + entry);
                }

                string key = entry.Substring(0, equalIndex).Trim();
                string method = entry.Substring(equalIndex + 1).Trim();
                PredictionTarget target = ResolvePredictionTarget(targets, key);
                selected[target] = ResolvePredictionMethodIndex(method);
            }
        }

        return selected;
    }

    private static HashSet<string> ParsePredictionScopes(string predictionScope)
    {
        var scopes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string scope in SplitList(string.IsNullOrWhiteSpace(predictionScope) ? "all" : predictionScope))
        {
            scopes.Add(NormalizePredictionToken(scope));
        }

        if (scopes.Count == 0)
        {
            scopes.Add("all");
        }

        return scopes;
    }

    private static HashSet<int> ParsePredictionRows(string predictionRows)
    {
        var rows = new HashSet<int>();
        foreach (string rowText in SplitList(predictionRows))
        {
            int row;
            if (!int.TryParse(rowText, NumberStyles.Integer, CultureInfo.InvariantCulture, out row))
            {
                throw new ArgumentException("Prediction row must be an integer: " + rowText);
            }

            rows.Add(row);
        }

        return rows;
    }

    private static HashSet<PredictionTarget> ResolvePredictionIndicatorTargets(
        List<PredictionTarget> targets,
        string predictionIndicators)
    {
        var selected = new HashSet<PredictionTarget>();
        foreach (string indicator in SplitList(predictionIndicators))
        {
            selected.Add(ResolvePredictionTarget(targets, indicator));
        }

        return selected;
    }

    private static IEnumerable<string> SplitList(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            yield break;
        }

        string[] parts = value.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (string part in parts)
        {
            string trimmed = part.Trim();
            if (trimmed.Length > 0)
            {
                yield return trimmed;
            }
        }
    }

    private static bool MatchesPredictionScope(PredictionTarget target, HashSet<string> scopes)
    {
        if (scopes.Contains("all"))
        {
            return true;
        }

        return scopes.Contains(target.Scope)
            || scopes.Contains(target.NormalizedPrefix);
    }

    private static PredictionTarget ResolvePredictionTarget(List<PredictionTarget> targets, string key)
    {
        string normalizedKey = NormalizePredictionToken(key);
        var matches = new List<PredictionTarget>();
        foreach (PredictionTarget target in targets)
        {
            if (string.Equals(normalizedKey, NormalizePredictionToken(target.ComboName), StringComparison.OrdinalIgnoreCase))
            {
                return target;
            }

            if (string.Equals(normalizedKey, target.Scope + target.Row.ToString(CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedKey, target.Scope + ":" + target.Row.ToString(CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedKey, target.NormalizedPrefix + target.Row.ToString(CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedKey, target.NormalizedPrefix + ":" + target.Row.ToString(CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase))
            {
                matches.Add(target);
            }
        }

        if (matches.Count == 1)
        {
            return matches[0];
        }

        if (matches.Count > 1)
        {
            throw new ArgumentException("Prediction setting key is ambiguous: " + key);
        }

        matches = ResolvePredictionTargetsByIndicator(targets, normalizedKey);
        if (matches.Count == 1)
        {
            return matches[0];
        }

        if (matches.Count > 1)
        {
            throw new ArgumentException("Prediction indicator name is ambiguous: " + key);
        }

        throw new ArgumentException("Prediction setting key not found: " + key);
    }

    private static List<PredictionTarget> ResolvePredictionTargetsByIndicator(
        List<PredictionTarget> targets,
        string normalizedKey)
    {
        var matches = new List<PredictionTarget>();
        foreach (PredictionIndicatorDefinition indicator in PredictionIndicators)
        {
            if (!indicator.Matches(normalizedKey))
            {
                continue;
            }

            PredictionTarget target = FindPredictionTarget(targets, indicator.Scope, indicator.Row);
            if (target != null && !matches.Contains(target))
            {
                matches.Add(target);
            }
        }

        return matches;
    }

    private static PredictionTarget FindPredictionTarget(
        List<PredictionTarget> targets,
        string scope,
        int row)
    {
        foreach (PredictionTarget target in targets)
        {
            if (string.Equals(target.Scope, scope, StringComparison.OrdinalIgnoreCase)
                && target.Row == row)
            {
                return target;
            }
        }

        return null;
    }

    private static string ResolvePredictionIndicatorName(string scope, int row)
    {
        foreach (PredictionIndicatorDefinition indicator in PredictionIndicators)
        {
            if (string.Equals(indicator.Scope, scope, StringComparison.OrdinalIgnoreCase)
                && indicator.Row == row)
            {
                return indicator.Name;
            }
        }

        return string.Empty;
    }

    private static bool TryParsePredictionCheckBoxName(string name, out string prefix, out int row)
    {
        prefix = string.Empty;
        row = 0;

        if (string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        int underscore = name.LastIndexOf('_');
        if (underscore <= 3 || underscore == name.Length - 1)
        {
            return false;
        }

        if (!int.TryParse(name.Substring(underscore + 1), NumberStyles.Integer, CultureInfo.InvariantCulture, out row))
        {
            return false;
        }

        string head = name.Substring(0, underscore);
        if (head.EndsWith("ckb", StringComparison.OrdinalIgnoreCase))
        {
            prefix = head.Substring(0, head.Length - 3);
            return prefix.Length > 0;
        }

        return false;
    }

    private static int ResolvePredictionMethodIndex(string method)
    {
        string token = NormalizePredictionToken(method);
        switch (token)
        {
            case "0":
            case "latest":
            case "latest1":
            case "last":
            case "last1":
            case "recent1":
                return 0;
            case "1":
            case "avg2":
            case "average2":
            case "arithmetic2":
                return 1;
            case "2":
            case "avg3":
            case "average3":
            case "arithmetic3":
                return 2;
            case "3":
            case "weighted2":
            case "wavg2":
            case "weightedavg2":
                return 3;
            case "4":
            case "weighted3":
            case "wavg3":
            case "weightedavg3":
                return 4;
            case "5":
            case "custom":
            case "manual":
            case "user":
                return 5;
            case "6":
            case "zero":
            case "tozero":
                return 6;
            default:
                throw new ArgumentException("Unsupported prediction method '" + method
                    + "'. Supported: latest1, avg2, avg3, weighted2, weighted3, custom, zero, or indexes 0-6.");
        }
    }

    private static string PredictionMethodName(int methodIndex)
    {
        switch (methodIndex)
        {
            case 0: return "latest1";
            case 1: return "avg2";
            case 2: return "avg3";
            case 3: return "weighted2";
            case 4: return "weighted3";
            case 5: return "custom";
            case 6: return "zero";
            default: return methodIndex.ToString(CultureInfo.InvariantCulture);
        }
    }

    private static string NormalizePredictionToken(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var builder = new StringBuilder(value.Length);
        foreach (char ch in value.Trim())
        {
            if (char.IsLetterOrDigit(ch))
            {
                builder.Append(char.ToLowerInvariant(ch));
            }
        }

        return builder.ToString();
    }

    private static string PredictionScopeFromPrefix(string prefix)
    {
        string normalized = NormalizePredictionToken(prefix);
        if (normalized.EndsWith("sales", StringComparison.Ordinal))
        {
            return "sales";
        }
        if (normalized.EndsWith("capital", StringComparison.Ordinal))
        {
            return "capital";
        }
        if (normalized.EndsWith("investment", StringComparison.Ordinal))
        {
            return "investment";
        }
        if (normalized.EndsWith("assume", StringComparison.Ordinal))
        {
            return "assume";
        }

        return normalized;
    }

    private static string ResolvePredictionSheetName(PredictionTarget target)
    {
        switch (target.Scope)
        {
            case "sales":
                return "\u9500\u552e\u9884\u6d4b";
            case "capital":
                return "\u8d44\u4ea7\u9884\u6d4b";
            case "investment":
                return "\u6295\u8d44\u4e0e\u51cf\u503c";
            case "assume":
                return "\u5047\u8bbe";
            default:
                throw new ArgumentException("Cannot map prediction scope to sheet: " + target.Scope);
        }
    }

    private static Sinitek.WriterModel.StockData ResolveStock(string stockCode, string gsdm, string stockName)
    {
        string code = FirstNonEmpty(stockCode, GetDocProperty("StkCode"));
        string resolvedGsdm = FirstNonEmpty(gsdm, GetDocProperty("GSCode"));
        if (!string.IsNullOrWhiteSpace(code) && !string.IsNullOrWhiteSpace(resolvedGsdm))
        {
            return new Sinitek.WriterModel.StockData
            {
                StockCode = code.Trim(),
                Gsdm = resolvedGsdm.Trim(),
                StockName = FirstNonEmpty(stockName, code.Trim())
            };
        }

        if (string.IsNullOrWhiteSpace(code))
        {
            throw new ArgumentException("Stock code is required. Pass -Stock or use a workbook that already has StkCode.");
        }

        var stocks = SearchStocks(code, 10, ResolveCurrentModelVersion(), ResolveCurrentModelType());
        Sinitek.WriterModel.StockData match = null;
        foreach (var candidate in stocks)
        {
            if (string.Equals(candidate.StockCode, code, StringComparison.OrdinalIgnoreCase))
            {
                match = candidate;
                break;
            }
        }
        if (match == null && stocks.Count > 0)
        {
            match = stocks[0];
        }
        if (match == null)
        {
            throw new InvalidOperationException("Cannot resolve stock via /api/stock. Pass both -Stock and -Gsdm.");
        }

        if (!string.IsNullOrWhiteSpace(stockName))
        {
            match.StockName = stockName;
        }

        return match;
    }

    private static List<string> ParseStockList(string stocks)
    {
        var result = new List<string>();
        if (string.IsNullOrWhiteSpace(stocks))
        {
            return result;
        }

        foreach (string item in stocks.Split(new[] { ',', ';', '\r', '\n', '\t', ' ' }, StringSplitOptions.RemoveEmptyEntries))
        {
            string stock = item.Trim();
            if (!string.IsNullOrWhiteSpace(stock))
            {
                result.Add(stock);
            }
        }

        return result;
    }

    private static string NormalizeStockInputForSearch(string stockCode)
    {
        string text = stockCode == null ? string.Empty : stockCode.Trim();
        if (text.EndsWith(".HK", StringComparison.OrdinalIgnoreCase))
        {
            return text.Substring(0, text.Length - 3);
        }

        return text;
    }

    private static string BuildOutputWorkbookPath(string outputDir, string workbookPath, string actionName, string stockCode)
    {
        string directory = Path.GetFullPath(outputDir);
        Directory.CreateDirectory(directory);
        string workbookName = Path.GetFileNameWithoutExtension(workbookPath);
        if (string.IsNullOrWhiteSpace(workbookName))
        {
            workbookName = "workbook";
        }

        string normalizedStock = NormalizeStockCodeForFileName(stockCode);
        string stockPart = string.IsNullOrWhiteSpace(normalizedStock) ? string.Empty : "-" + normalizedStock;
        string timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
        string candidate = Path.Combine(directory, workbookName + "-" + actionName + stockPart + "-" + timestamp + ".xlsx");
        return EnsureUniquePath(candidate);
    }

    private static string NormalizeStockCodeForFileName(string stockCode)
    {
        if (string.IsNullOrWhiteSpace(stockCode))
        {
            return string.Empty;
        }

        StringBuilder builder = new StringBuilder();
        foreach (char ch in stockCode.Trim())
        {
            if (char.IsDigit(ch))
            {
                builder.Append(ch);
            }
        }

        return builder.ToString();
    }

    private static string EnsureUniquePath(string path)
    {
        if (!File.Exists(path))
        {
            return path;
        }

        string directory = Path.GetDirectoryName(path);
        string name = Path.GetFileNameWithoutExtension(path);
        string extension = Path.GetExtension(path);
        for (int i = 2; i < 10000; i++)
        {
            string candidate = Path.Combine(directory, name + "-" + i.ToString("00", CultureInfo.InvariantCulture) + extension);
            if (!File.Exists(candidate))
            {
                return candidate;
            }
        }

        throw new IOException("Cannot create a unique output filename for " + path);
    }

    private static PeerStockSelection ResolvePeerStockSelection(
        Sinitek.WriterModel.StockData stock,
        string explicitPeerStock,
        string previousGsdm)
    {
        if (!string.IsNullOrWhiteSpace(explicitPeerStock))
        {
            string normalized = NormalizePeerStockList(explicitPeerStock);
            return new PeerStockSelection(normalized, "explicit");
        }

        if (stock == null || string.IsNullOrWhiteSpace(stock.Gsdm) || IsNEEQStock(stock.Gsdm))
        {
            return new PeerStockSelection(string.Empty, "none");
        }

        string existingPeerStock = NormalizePeerStockList(GetDocProperty("PeerStock"));
        if (string.Equals(NormalizeGsdm(previousGsdm), NormalizeGsdm(stock.Gsdm), StringComparison.OrdinalIgnoreCase))
        {
            return new PeerStockSelection(existingPeerStock, "workbook");
        }

        var peers = GetDefaultPeerStocks(stock.Gsdm);
        var gsdms = new List<string>();
        foreach (var peer in peers)
        {
            AddUniqueGsdm(gsdms, peer.Gsdm);
        }

        return new PeerStockSelection(string.Join(",", gsdms.ToArray()), "cloud-default");
    }

    private static List<Sinitek.WriterModel.StockData> GetDefaultPeerStocks(string gsdm)
    {
        var peers = new List<Sinitek.WriterModel.StockData>();
        string url = (SinitekExcel.WriterUtil.ModelUtil.ModelUrl ?? string.Empty).TrimEnd('/') + "/api/company/analysis/gsdms";
        IDictionary parameters = new Hashtable();
        parameters.Add("gsdm", gsdm);

        object webHandler = SinitekExcel.WriterUtil.Web.ModelWebDataHandler.NewInstance();
        Hashtable response = InvokeMethod(webHandler, "GetMap", false, url, parameters) as Hashtable;
        if (response == null || !response.ContainsKey("data") || response["data"] == null)
        {
            return peers;
        }

        JArray data = ToJArray(response["data"]);
        if (data == null)
        {
            return peers;
        }

        foreach (JToken token in data)
        {
            var peer = DeserializeStock(token);
            if (peer != null && !string.IsNullOrWhiteSpace(peer.Gsdm))
            {
                peers.Add(peer);
            }
        }

        return peers;
    }

    private static JArray ToJArray(object value)
    {
        JArray array = value as JArray;
        if (array != null)
        {
            return array;
        }

        JToken token = value as JToken;
        if (token != null)
        {
            return FindStockArray(token, 0);
        }

        string text = Convert.ToString(value, CultureInfo.InvariantCulture);
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        text = text.Trim();
        if (!text.StartsWith("[", StringComparison.Ordinal) && !text.StartsWith("{", StringComparison.Ordinal))
        {
            return null;
        }

        return FindStockArray(JToken.Parse(text), 0);
    }

    private static string NormalizePeerStockList(string peerStock)
    {
        var gsdms = new List<string>();
        if (!string.IsNullOrWhiteSpace(peerStock))
        {
            foreach (string item in peerStock.Split(new[] { ',', ';', '\r', '\n', '\t', ' ' }, StringSplitOptions.RemoveEmptyEntries))
            {
                AddUniqueGsdm(gsdms, item);
            }
        }

        return string.Join(",", gsdms.ToArray());
    }

    private static void AddUniqueGsdm(List<string> gsdms, string gsdm)
    {
        string normalized = NormalizeGsdm(gsdm);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return;
        }

        foreach (string existing in gsdms)
        {
            if (string.Equals(existing, normalized, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
        }

        gsdms.Add(normalized);
    }

    private static string NormalizeGsdm(string gsdm)
    {
        return string.IsNullOrWhiteSpace(gsdm) ? string.Empty : gsdm.Trim();
    }

    private static bool IsNEEQStock(string gsdm)
    {
        if (string.IsNullOrWhiteSpace(gsdm))
        {
            return false;
        }

        Type commonType = Assembly.LoadFrom(AddinDll).GetType("Sinitek.WriterUtil.Common");
        if (commonType == null)
        {
            return false;
        }

        MethodInfo method = commonType.GetMethod("IsNEEQStock", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static);
        if (method == null)
        {
            return false;
        }

        return (bool)method.Invoke(null, new object[] { gsdm });
    }

    private static List<Sinitek.WriterModel.StockData> SearchStocks(string query, int count, string modelVersion, string modelType)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            throw new ArgumentException("Stock search query is required.");
        }

        string url = (SinitekExcel.WriterUtil.ModelUtil.ModelUrl ?? string.Empty).TrimEnd('/') + "/api/stock";
        string postBody = BuildStockSearchPostBody(url, query.Trim(), count <= 0 ? 10 : count, modelVersion, modelType);

        string json = PostForm(url, postBody, "Stock search");
        bool foundStockArray;
        string serverMessage;
        var stocks = ParseStockSearchJson(json, out foundStockArray, out serverMessage);
        if (!foundStockArray)
        {
            string message = "Stock search response did not contain a stock list.";
            if (!string.IsNullOrWhiteSpace(serverMessage))
            {
                message += " Server message: " + serverMessage;
            }
            message += " Raw response: " + Abbreviate(json, 500);
            throw new InvalidOperationException(message);
        }

        return stocks;
    }

    private static string BuildStockSearchPostBody(string url, string query, int count, string modelVersion, string modelType)
    {
        var parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        parameters["searchText"] = query;
        parameters["count"] = count.ToString(CultureInfo.InvariantCulture);

        AddPluginWebParameters(parameters, url, modelVersion, modelType);

        return FormatFormBody(parameters);
    }

    private static void AddPluginWebParameters(IDictionary<string, string> parameters, string url, string modelVersion, string modelType)
    {
        object webHandler = SinitekExcel.WriterUtil.Web.ModelWebDataHandler.NewInstance();
        string timestamp = (string)InvokeMethod(webHandler, "GetTimeStamp", false);
        string relativeUrl = url.Replace(SinitekExcel.WriterUtil.ModelUtil.ModelUrl ?? string.Empty, string.Empty);
        int queryIndex = relativeUrl.IndexOf("?", StringComparison.Ordinal);
        if (queryIndex >= 0)
        {
            relativeUrl = relativeUrl.Substring(0, queryIndex);
        }

        parameters["tokenid"] = Safe(() => SinitekExcel.WriterUtil.WriteUtil.TokenId);
        parameters["timestamp"] = timestamp;
        parameters["sign"] = (string)InvokeMethod(webHandler, "GetSign", false, timestamp, relativeUrl);
        parameters["modelversion"] = modelVersion ?? string.Empty;
        parameters["pluginversion"] = Safe(() => SinitekExcel.WriterUtil.ModelUtil.PluginVer);
        parameters["modeltype"] = modelType ?? string.Empty;
        parameters["ov"] = Safe(() => SinitekExcel.WriterUtil.LoadUtil.getOfficeVer());
        parameters["uuid"] = Safe(() => SinitekExcel.WriterBLL.WriterUuidBLL.Instance.GetUuid());
    }

    private static string FormatFormBody(IDictionary<string, string> parameters)
    {
        var builder = new StringBuilder();
        foreach (KeyValuePair<string, string> pair in parameters)
        {
            if (builder.Length > 0)
            {
                builder.Append('&');
            }
            builder.Append(Uri.EscapeDataString(pair.Key));
            builder.Append('=');
            builder.Append(Uri.EscapeDataString(pair.Value ?? string.Empty));
        }

        return builder.ToString();
    }

    private static string PostForm(string url, string postBody, string operation)
    {
        try
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.Accept = "application/json,text/plain,*/*";
            request.Timeout = StockSearchTimeoutMs;
            request.ReadWriteTimeout = StockSearchTimeoutMs;

            byte[] bodyBytes = Encoding.UTF8.GetBytes(postBody);
            request.ContentLength = bodyBytes.Length;
            using (Stream stream = request.GetRequestStream())
            {
                stream.Write(bodyBytes, 0, bodyBytes.Length);
            }

            using (var response = (HttpWebResponse)request.GetResponse())
            using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
            {
                return reader.ReadToEnd();
            }
        }
        catch (WebException ex)
        {
            if (ex.Status == WebExceptionStatus.Timeout)
            {
                throw new TimeoutException(operation + " timed out after "
                    + (StockSearchTimeoutMs / 1000).ToString(CultureInfo.InvariantCulture)
                    + " seconds: " + url, ex);
            }

            string detail = ReadWebExceptionBody(ex);
            if (!string.IsNullOrWhiteSpace(detail))
            {
                detail = " Response: " + detail;
            }
            throw new InvalidOperationException(operation + " request failed: " + ex.Message + detail, ex);
        }
    }

    private static List<Sinitek.WriterModel.StockData> ParseStockSearchJson(string json, out bool foundStockArray, out string serverMessage)
    {
        foundStockArray = false;
        serverMessage = string.Empty;
        var result = new List<Sinitek.WriterModel.StockData>();
        if (string.IsNullOrWhiteSpace(json))
        {
            return result;
        }

        JToken root = JToken.Parse(json);
        serverMessage = ExtractServerMessage(root);

        JArray array = FindStockArray(root, 0);
        if (array == null)
        {
            return result;
        }
        foundStockArray = true;

        foreach (JToken token in array)
        {
            var stock = DeserializeStock(token);
            if (stock != null && HasStockFields(stock))
            {
                result.Add(stock);
            }
        }

        return result;
    }

    private static Sinitek.WriterModel.StockData DeserializeStock(JToken token)
    {
        var stock = JsonConvert.DeserializeObject<Sinitek.WriterModel.StockData>(token.ToString());
        if (stock == null)
        {
            stock = new Sinitek.WriterModel.StockData();
        }

        JObject obj = token as JObject;
        if (obj == null)
        {
            return stock;
        }

        stock.StockCode = FirstNonEmpty(stock.StockCode, JsonValue(obj, "stkcode", "stockCode", "code", "scode"));
        stock.StockName = FirstNonEmpty(stock.StockName, JsonValue(obj, "sname", "stockName", "name", "stkname"));
        stock.Gsdm = FirstNonEmpty(stock.Gsdm, JsonValue(obj, "gsdm", "gscode", "gsCode", "companyCode", "orgcode"));
        stock.StockType = FirstNonEmpty(stock.StockType, JsonValue(obj, "stype", "stockType", "type"));
        stock.MarketCode = FirstNonEmpty(stock.MarketCode, JsonValue(obj, "marketCode", "mktcode", "market"));
        stock.MarketName = FirstNonEmpty(stock.MarketName, JsonValue(obj, "marketName", "mktname"));
        stock.SecurityCode = FirstNonEmpty(stock.SecurityCode, JsonValue(obj, "securityCode", "secuCode", "seccode"));
        stock.IndustryCode = FirstNonEmpty(stock.IndustryCode, JsonValue(obj, "industryCode", "industry"));
        if (string.IsNullOrWhiteSpace(stock.Gsdm) && !string.IsNullOrWhiteSpace(stock.SecurityCode))
        {
            stock.Gsdm = stock.SecurityCode;
        }
        if (string.IsNullOrWhiteSpace(stock.SecurityCode) && !string.IsNullOrWhiteSpace(stock.Gsdm) && stock.Gsdm.IndexOf('.') >= 0)
        {
            stock.SecurityCode = stock.Gsdm;
        }

        return stock;
    }

    private static string JsonValue(JObject obj, params string[] keys)
    {
        if (obj == null)
        {
            return string.Empty;
        }

        foreach (string key in keys)
        {
            foreach (JProperty property in obj.Properties())
            {
                if (string.Equals(property.Name, key, StringComparison.OrdinalIgnoreCase)
                    && property.Value != null
                    && property.Value.Type != JTokenType.Null)
                {
                    return property.Value.Type == JTokenType.String
                        ? property.Value.ToString()
                        : property.Value.ToString(Formatting.None);
                }
            }
        }

        return string.Empty;
    }

    private static bool HasStockFields(Sinitek.WriterModel.StockData stock)
    {
        return !string.IsNullOrWhiteSpace(stock.StockCode)
            || !string.IsNullOrWhiteSpace(stock.StockName)
            || !string.IsNullOrWhiteSpace(stock.Gsdm)
            || !string.IsNullOrWhiteSpace(stock.SecurityCode);
    }

    private static JArray FindStockArray(JToken token, int depth)
    {
        if (token == null || depth > 4)
        {
            return null;
        }

        JArray array = token as JArray;
        if (array != null)
        {
            return array;
        }

        JValue value = token as JValue;
        if (value != null && value.Type == JTokenType.String)
        {
            string text = Convert.ToString(value.Value, CultureInfo.InvariantCulture);
            if (!string.IsNullOrWhiteSpace(text))
            {
                text = text.Trim();
                if (text.StartsWith("{", StringComparison.Ordinal) || text.StartsWith("[", StringComparison.Ordinal))
                {
                    return FindStockArray(JToken.Parse(text), depth + 1);
                }
            }
            return null;
        }

        JObject obj = token as JObject;
        if (obj == null)
        {
            return null;
        }

        string[] preferredKeys = new[] { "stock", "data", "rows", "list", "result", "results" };
        foreach (string key in preferredKeys)
        {
            JToken child = obj[key];
            JArray found = FindStockArray(child, depth + 1);
            if (found != null)
            {
                return found;
            }
        }

        return null;
    }

    private static string ExtractServerMessage(JToken token)
    {
        JObject obj = token as JObject;
        if (obj == null)
        {
            return string.Empty;
        }

        foreach (string key in new[] { "message", "Message", "msg", "Msg", "error", "Error" })
        {
            JToken value = obj[key];
            if (value != null && value.Type != JTokenType.Null)
            {
                return value.ToString();
            }
        }

        return string.Empty;
    }

    private static string Abbreviate(string value, int maxLength)
    {
        if (string.IsNullOrEmpty(value) || value.Length <= maxLength)
        {
            return value ?? string.Empty;
        }

        return value.Substring(0, maxLength) + "...";
    }

    private static string ReadWebExceptionBody(WebException ex)
    {
        if (ex.Response == null)
        {
            return string.Empty;
        }

        try
        {
            using (Stream stream = ex.Response.GetResponseStream())
            {
                if (stream == null)
                {
                    return string.Empty;
                }

                using (var reader = new StreamReader(stream, Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
        }
        catch
        {
            return string.Empty;
        }
    }

    private static void LoginIfProvided(string username, string password)
    {
        LoadIpConfig();
        if (!string.IsNullOrWhiteSpace(username) || !string.IsNullOrEmpty(password))
        {
            Login(username, password, false);
        }
    }

    private static string ResolveCurrentModelVersion()
    {
        string modelVersion = Safe(() => SinitekExcel.WriterUtil.LoadUtil.getModelVersion());
        if (string.IsNullOrWhiteSpace(modelVersion) || modelVersion.StartsWith("ERR:", StringComparison.Ordinal))
        {
            modelVersion = GetOfficeDocProperty("ModelVersion");
        }
        if (string.IsNullOrWhiteSpace(modelVersion))
        {
            throw new InvalidOperationException("Cannot resolve ModelVersion from workbook.");
        }

        return modelVersion;
    }

    private static string ResolveCurrentModelType()
    {
        int modelType = -1;
        try
        {
            modelType = SinitekExcel.WriterUtil.LoadUtil.getModelType();
        }
        catch
        {
            modelType = -1;
        }

        if (modelType >= 0)
        {
            return modelType.ToString(CultureInfo.InvariantCulture);
        }

        string modelTypeText = GetOfficeDocProperty("ModelType");
        if (string.IsNullOrWhiteSpace(modelTypeText))
        {
            throw new InvalidOperationException("Cannot resolve ModelType from workbook.");
        }

        return modelTypeText;
    }

    private static void EnsureLoggedIn()
    {
        if (!string.Equals(SinitekExcel.WriterUtil.WriteUtil.LoginState, "200", StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Not logged in. Pass -Username/-Password or run the action in a process that logs in first.");
        }

        string tokenId = NullToEmpty(SinitekExcel.WriterUtil.WriteUtil.TokenId);
        if (string.IsNullOrWhiteSpace(tokenId))
        {
            throw new InvalidOperationException("Login token is empty after login.");
        }

        string modelVersion = ResolveCurrentModelVersion();
        string modelType = ResolveCurrentModelType();
        string rightResponse;
        if (!CheckModelRightDirect(modelVersion, modelType, out rightResponse))
        {
            throw new InvalidOperationException("Current account has no model permission for this workbook. "
                + "UserName=" + NullToEmpty(SinitekExcel.WriterUtil.WriteUtil.UserName)
                + ", ModelVersion=" + modelVersion
                + ", ModelType=" + modelType
                + ", Response=" + Abbreviate(rightResponse, 500));
        }
    }

    private static bool CheckModelRightDirect(string modelVersion, string modelType, out string responseJson)
    {
        string userName = SinitekExcel.WriterUtil.WriteUtil.UserName;
        if (string.IsNullOrWhiteSpace(userName))
        {
            responseJson = string.Empty;
            return false;
        }

        string url = (SinitekExcel.WriterUtil.ModelUtil.ModelUrl ?? string.Empty).TrimEnd('/') + "/api/checkmodelauth";
        var parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        parameters["email"] = userName;
        AddPluginWebParameters(parameters, url, modelVersion, modelType);

        responseJson = PostForm(url, FormatFormBody(parameters), "Model right check");
        if (string.IsNullOrWhiteSpace(responseJson))
        {
            return false;
        }

        JToken root = JToken.Parse(responseJson);
        string authFlag = FindJsonString(root, "authflag");
        return string.Equals(authFlag, "1", StringComparison.Ordinal);
    }

    private static string FindJsonString(JToken token, string key)
    {
        if (token == null)
        {
            return string.Empty;
        }

        JObject obj = token as JObject;
        if (obj != null)
        {
            foreach (JProperty property in obj.Properties())
            {
                if (string.Equals(property.Name, key, StringComparison.OrdinalIgnoreCase)
                    && property.Value != null
                    && property.Value.Type != JTokenType.Null)
                {
                    return property.Value.Type == JTokenType.String
                        ? property.Value.ToString()
                        : property.Value.ToString(Formatting.None);
                }
            }

            foreach (JProperty property in obj.Properties())
            {
                string found = FindJsonString(property.Value, key);
                if (!string.IsNullOrWhiteSpace(found))
                {
                    return found;
                }
            }
        }

        JArray array = token as JArray;
        if (array != null)
        {
            foreach (JToken item in array)
            {
                string found = FindJsonString(item, key);
                if (!string.IsNullOrWhiteSpace(found))
                {
                    return found;
                }
            }
        }

        return string.Empty;
    }

    private static void LoadIpConfig()
    {
        var list = new ArrayList();
        SinitekExcel.WriterBLL.WriterIPConfigBLL.Instance.getIPInfomation(ref list);
    }

    private static Type ResolveType(string typeName)
    {
        Type type = Type.GetType(typeName + ", SinitekExcel")
            ?? Assembly.LoadFrom(AddinDll).GetType(typeName);
        if (type == null)
        {
            throw new InvalidOperationException("Cannot resolve type: " + typeName);
        }

        return type;
    }

    private static object InvokeMethod(object target, string name, bool publicOnly, params object[] args)
    {
        BindingFlags flags = BindingFlags.Instance | (publicOnly ? BindingFlags.Public : BindingFlags.Public | BindingFlags.NonPublic);
        MethodInfo method = target.GetType().GetMethod(name, flags);
        if (method == null)
        {
            throw new MissingMethodException(target.GetType().FullName, name);
        }

        try
        {
            return method.Invoke(target, args);
        }
        catch (TargetInvocationException ex)
        {
            Exception inner = ex.InnerException ?? ex;
            throw new InvalidOperationException("Invocation failed: "
                + target.GetType().FullName + "." + name + ": " + inner.Message, inner);
        }
    }

    private static string GetDocProperty(string name)
    {
        string value = string.Empty;
        try
        {
            SinitekExcel.WriterUtil.ModelUtil.getDocProperty(name, ref value);
        }
        catch
        {
            value = string.Empty;
        }

        return value ?? string.Empty;
    }

    private static string GetOfficeDocProperty(string name)
    {
        string value;
        return TryGetOfficeDocProperty(null, name, out value) ? value : string.Empty;
    }

    private static bool TryGetOfficeDocProperty(object workbook, string name, out string value)
    {
        object props = null;
        object prop = null;
        value = string.Empty;
        try
        {
            if (workbook == null)
            {
                workbook = SinitekExcel.WriterUtil.WriteUtil.Application.ActiveWorkbook;
            }
            props = workbook.GetType().InvokeMember("CustomDocumentProperties", BindingFlags.GetProperty, null, workbook, null);
            prop = props.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, props, new object[] { name });
            object rawValue = prop.GetType().InvokeMember("Value", BindingFlags.GetProperty, null, prop, null);
            value = Convert.ToString(rawValue, CultureInfo.InvariantCulture) ?? string.Empty;
            return true;
        }
        catch
        {
            value = string.Empty;
            return false;
        }
        finally
        {
            if (prop != null && Marshal.IsComObject(prop))
            {
                Marshal.ReleaseComObject(prop);
            }
            if (props != null && Marshal.IsComObject(props))
            {
                Marshal.ReleaseComObject(props);
            }
        }
    }

    private static void SetDocProperty(string name, string value)
    {
        try
        {
            SetOfficeDocProperty(name, value ?? string.Empty);
        }
        catch (Exception ex)
        {
            Exception detail = UnwrapInvocationException(ex);
            throw new InvalidOperationException("Failed to set workbook document property "
                + name + ": " + detail.Message, ex);
        }
    }

    private static void SetOfficeDocProperty(string name, string value)
    {
        SetOfficeDocProperty(null, name, value);
    }

    private static void SetOfficeDocProperty(object workbook, string name, string value)
    {
        object props = null;
        object prop = null;
        try
        {
            if (workbook == null)
            {
                workbook = SinitekExcel.WriterUtil.WriteUtil.Application.ActiveWorkbook;
            }
            props = workbook.GetType().InvokeMember("CustomDocumentProperties", BindingFlags.GetProperty, null, workbook, null);
            try
            {
                prop = props.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, props, new object[] { name });
            }
            catch (Exception ex)
            {
                if (!IsMissingComMember(ex))
                {
                    throw;
                }

                AddOfficeDocProperty(props, name, value);
                return;
            }

            try
            {
                prop.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, prop, new object[] { value ?? string.Empty });
            }
            catch (Exception ex)
            {
                if (!IsMissingComMember(ex))
                {
                    throw;
                }

                prop.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, prop, null);
                Marshal.ReleaseComObject(prop);
                prop = null;
                AddOfficeDocProperty(props, name, value);
            }
        }
        finally
        {
            if (prop != null && Marshal.IsComObject(prop))
            {
                Marshal.ReleaseComObject(prop);
            }
            if (props != null && Marshal.IsComObject(props))
            {
                Marshal.ReleaseComObject(props);
            }
        }
    }

    private static void AddOfficeDocProperty(object props, string name, string value)
    {
        props.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, props, new object[]
        {
            name,
            false,
            4,
            value ?? string.Empty
        });
    }

    private static bool IsMissingComMember(Exception ex)
    {
        Exception detail = UnwrapInvocationException(ex);

        COMException com = detail as COMException;
        if (com != null)
        {
            return unchecked((uint)com.ErrorCode) == 0x8002000B
                || unchecked((uint)com.ErrorCode) == 0x800A03EC
                || unchecked((uint)com.ErrorCode) == 0x80070057;
        }

        return detail is ArgumentException
            && unchecked((uint)detail.HResult) == 0x80070057;
    }

    private static Exception UnwrapInvocationException(Exception ex)
    {
        while (ex is TargetInvocationException && ex.InnerException != null)
        {
            ex = ex.InnerException;
        }

        return ex;
    }

    private static void InstallDocumentPropertyPatch()
    {
        if (documentPropertyPatchInstalled)
        {
            return;
        }

        Type modelUtil = ResolveType("SinitekExcel.WriterUtil.ModelUtil");
        BindingFlags flags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static;

        MethodInfo originalGet = modelUtil.GetMethod("getDocProperty", flags, null, new[] { typeof(string), typeof(string).MakeByRefType() }, null);
        MethodInfo originalSet = modelUtil.GetMethod("setDocProperty", flags, null, new[] { typeof(string), typeof(string) }, null);
        MethodInfo originalSetWorkbook = FindStaticMethod(modelUtil, "setDocProperty", 3);

        MethodInfo patchedGet = typeof(SinitekCliBridge).GetMethod("PatchedGetDocProperty", BindingFlags.NonPublic | BindingFlags.Static);
        MethodInfo patchedSet = typeof(SinitekCliBridge).GetMethod("PatchedSetDocProperty", BindingFlags.NonPublic | BindingFlags.Static);
        MethodInfo patchedSetWorkbook = typeof(SinitekCliBridge).GetMethod("PatchedSetDocPropertyForWorkbook", BindingFlags.NonPublic | BindingFlags.Static);

        PatchMethod(originalGet, patchedGet);
        PatchMethod(originalSet, patchedSet);
        PatchMethod(originalSetWorkbook, patchedSetWorkbook);

        documentPropertyPatchInstalled = true;
    }

    private static MethodInfo FindStaticMethod(Type type, string name, int parameterCount)
    {
        foreach (MethodInfo method in type.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static))
        {
            if (method.Name == name && method.GetParameters().Length == parameterCount)
            {
                return method;
            }
        }

        return null;
    }

    private static void PatchMethod(MethodInfo original, MethodInfo replacement)
    {
        if (original == null || replacement == null)
        {
            throw new InvalidOperationException("Cannot install document property patch because a method was not found.");
        }

        RuntimeHelpers.PrepareMethod(original.MethodHandle);
        RuntimeHelpers.PrepareMethod(replacement.MethodHandle);

        IntPtr originalPointer = original.MethodHandle.GetFunctionPointer();
        IntPtr replacementPointer = replacement.MethodHandle.GetFunctionPointer();

        byte[] patch;
        if (IntPtr.Size == 8)
        {
            patch = new byte[12];
            patch[0] = 0x48;
            patch[1] = 0xB8;
            BitConverter.GetBytes(replacementPointer.ToInt64()).CopyTo(patch, 2);
            patch[10] = 0xFF;
            patch[11] = 0xE0;
        }
        else
        {
            patch = new byte[7];
            patch[0] = 0xB8;
            BitConverter.GetBytes(replacementPointer.ToInt32()).CopyTo(patch, 1);
            patch[5] = 0xFF;
            patch[6] = 0xE0;
        }

        uint oldProtect;
        if (!VirtualProtect(originalPointer, (UIntPtr)patch.Length, 0x40, out oldProtect))
        {
            throw new InvalidOperationException("VirtualProtect failed while installing document property patch. Win32Error=" + Marshal.GetLastWin32Error());
        }

        Marshal.Copy(patch, 0, originalPointer, patch.Length);

        uint ignored;
        VirtualProtect(originalPointer, (UIntPtr)patch.Length, oldProtect, out ignored);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static bool PatchedGetDocProperty(string name, ref string value)
    {
        return TryGetOfficeDocProperty(null, name, out value);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void PatchedSetDocProperty(string name, string value)
    {
        SetOfficeDocProperty(null, name, value);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void PatchedSetDocPropertyForWorkbook(object workbook, string name, string value)
    {
        SetOfficeDocProperty(workbook, name, value);
    }

    private static void TrySetDocProperty(string name, string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        try
        {
            SetDocProperty(name, value);
        }
        catch
        {
        }
    }

    private static void InitializeModelContextFromWorkbook()
    {
        TrySetDocProperty("ModelVersion", GetOfficeDocProperty("ModelVersion"));
        TrySetDocProperty("ModelType", GetOfficeDocProperty("ModelType"));
    }

    private static void ApplyYearSettings(Excel.Workbook workbook, int historyYear, int forecastYear)
    {
        Excel.Worksheet directorySheet = null;
        try
        {
            directorySheet = GetWorksheet(workbook, DirectorySheetName);
            SetCellValue(directorySheet, 3, "D", historyYear);
            SetCellValue(directorySheet, 4, "D", forecastYear);
        }
        finally
        {
            if (directorySheet != null && Marshal.IsComObject(directorySheet))
            {
                Marshal.ReleaseComObject(directorySheet);
            }
        }
    }

    private static CurrencyUnitSpec ApplyCurrencyUnit(Excel.Workbook workbook, string currencyUnit)
    {
        CurrencyUnitSpec currency = ResolveCurrencyUnit(currencyUnit);

        Excel.Worksheet directorySheet = null;
        try
        {
            directorySheet = GetWorksheet(workbook, DirectorySheetName);
            SetCellValue(directorySheet, 2, "D", currency.ShareUnitName);
            SetCellValue(directorySheet, 5, "D", currency.CurrencyUnitName);
            SetCellValue(directorySheet, 6, "D", currency.Scale);
        }
        finally
        {
            if (directorySheet != null && Marshal.IsComObject(directorySheet))
            {
                Marshal.ReleaseComObject(directorySheet);
            }
        }

        return currency;
    }

    private static CurrencyUnitSpec ResolveCurrencyUnit(string currencyUnit)
    {
        string text = (currencyUnit ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            return CurrencyUnits[3];
        }

        foreach (CurrencyUnitSpec currency in CurrencyUnits)
        {
            if (string.Equals(text, currency.ScaleText, StringComparison.OrdinalIgnoreCase)
                || string.Equals(text, currency.CurrencyUnitName, StringComparison.OrdinalIgnoreCase)
                || string.Equals(text, currency.ShareUnitName, StringComparison.OrdinalIgnoreCase))
            {
                return currency;
            }
        }

        double parsed;
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed)
            || double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out parsed))
        {
            foreach (CurrencyUnitSpec currency in CurrencyUnits)
            {
                if (Math.Abs(parsed - currency.Scale) <= 0.000000000001D)
                {
                    return currency;
                }
            }
        }

        throw new ArgumentException("Unsupported currency unit '" + text
            + "'. Supported values are 1, 0.001, 0.0001, 0.000001, 0.00000001.");
    }

    private static Excel.Worksheet GetWorksheet(Excel.Workbook workbook, string name)
    {
        Excel.Sheets sheets = null;
        try
        {
            sheets = workbook.Worksheets;
            int count = sheets.Count;
            for (int i = 1; i <= count; i++)
            {
                Excel.Worksheet sheet = null;
                bool returnSheet = false;
                try
                {
                    sheet = (Excel.Worksheet)sheets[i];
                    if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                    {
                        returnSheet = true;
                        return sheet;
                    }
                }
                finally
                {
                    if (!returnSheet && sheet != null && Marshal.IsComObject(sheet))
                    {
                        Marshal.ReleaseComObject(sheet);
                    }
                }
            }
        }
        finally
        {
            if (sheets != null && Marshal.IsComObject(sheets))
            {
                Marshal.ReleaseComObject(sheets);
            }
        }

        throw new InvalidOperationException("Worksheet not found: " + name);
    }

    private static void SetCellValue(Excel.Worksheet sheet, int row, string column, object value)
    {
        Excel.Range cell = null;
        try
        {
            cell = (Excel.Range)sheet.Cells[row, column];
            cell.Value2 = value;
        }
        finally
        {
            if (cell != null && Marshal.IsComObject(cell))
            {
                Marshal.ReleaseComObject(cell);
            }
        }
    }

    private static bool SheetExists(Excel.Workbook workbook, string name)
    {
        foreach (Excel.Worksheet sheet in workbook.Worksheets)
        {
            try
            {
                if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            finally
            {
                Marshal.ReleaseComObject(sheet);
            }
        }

        return false;
    }

    private static void SaveIfRequested(Excel.Workbook workbook, string outWorkbook, bool saveOriginal)
    {
        if (!string.IsNullOrWhiteSpace(outWorkbook))
        {
            string fullPath = Path.GetFullPath(outWorkbook);
            Directory.CreateDirectory(Path.GetDirectoryName(fullPath));
            workbook.SaveAs(fullPath);
            return;
        }

        if (saveOriginal)
        {
            workbook.Save();
        }
    }

    private static string ResolveMailSuffix(string username)
    {
        string resolvedUsername = FirstNonEmpty(
            username,
            Safe(() => SinitekExcel.WriterUtil.WriteUtil.UserName),
            Environment.GetEnvironmentVariable("SINITEK_USERNAME"));
        int at = string.IsNullOrWhiteSpace(resolvedUsername) ? -1 : resolvedUsername.IndexOf('@');
        if (at >= 0 && at < resolvedUsername.Length - 1)
        {
            return "@" + resolvedUsername.Substring(at + 1).Trim();
        }

        throw new InvalidOperationException("Cannot resolve output mail suffix from username. Set SINITEK_USERNAME to an email address such as your.name@domainname.com, or pass -Username with an email address.");
    }

    private static string FirstNonEmpty(params string[] values)
    {
        foreach (string value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value.Trim();
            }
        }

        return string.Empty;
    }

    private static string NullToEmpty(string value)
    {
        return value ?? string.Empty;
    }

    private static int SafeInt(string value, int fallback)
    {
        int parsed;
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) ? parsed : fallback;
    }

    private static string Safe(Func<string> getter)
    {
        try
        {
            return getter() ?? string.Empty;
        }
        catch (Exception ex)
        {
            return "ERR: " + ex.Message;
        }
    }

    private sealed class UpdateExecutionResult
    {
        public UpdateExecutionResult(string outputText, string artifact)
        {
            OutputText = outputText ?? string.Empty;
            Artifact = artifact ?? string.Empty;
        }

        public string OutputText { get; private set; }
        public string Artifact { get; private set; }
    }

    private sealed class UpdateOptions
    {
        public string WorkbookPath { get; set; }
        public string OutWorkbook { get; set; }
        public bool SaveOriginal { get; set; }
        public string OutputDir { get; set; }
        public string OutputActionName { get; set; }
        public string StockCode { get; set; }
        public string Gsdm { get; set; }
        public string StockName { get; set; }
        public int HistoryYear { get; set; }
        public int ForecastYear { get; set; }
        public string CurrencyUnit { get; set; }
        public string SegmentDimension { get; set; }
        public string PeerStock { get; set; }
        public bool UpdateDirectory { get; set; }
        public bool UpdateSrcData { get; set; }
        public bool Migrate { get; set; }
        public bool AddOutput { get; set; }
        public string Username { get; set; }
    }

    private sealed class CurrencyUnitSpec
    {
        public CurrencyUnitSpec(string scaleText, double scale, string currencyUnitName, string shareUnitName)
        {
            ScaleText = scaleText;
            Scale = scale;
            CurrencyUnitName = currencyUnitName;
            ShareUnitName = shareUnitName;
        }

        public string ScaleText { get; private set; }
        public double Scale { get; private set; }
        public string CurrencyUnitName { get; private set; }
        public string ShareUnitName { get; private set; }
    }

    private sealed class PeerStockSelection
    {
        public PeerStockSelection(string gsdms, string source)
        {
            Gsdms = gsdms ?? string.Empty;
            Source = source ?? string.Empty;
            Count = string.IsNullOrWhiteSpace(Gsdms) ? 0 : Gsdms.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Length;
        }

        public string Gsdms { get; private set; }
        public string Source { get; private set; }
        public int Count { get; private set; }
    }

    private sealed class PredictionTarget
    {
        public PredictionTarget(
            string prefix,
            string scope,
            int row,
            string comboName,
            string indicatorName)
        {
            NormalizedPrefix = NormalizePredictionToken(prefix);
            Scope = scope;
            Row = row;
            ComboName = comboName;
            IndicatorName = indicatorName ?? string.Empty;
        }

        public string NormalizedPrefix { get; private set; }
        public string Scope { get; private set; }
        public int Row { get; private set; }
        public string ComboName { get; private set; }
        public string IndicatorName { get; private set; }
    }

    private sealed class PredictionIndicatorDefinition
    {
        private readonly string[] normalizedAliases;

        public PredictionIndicatorDefinition(string scope, int row, string name, params string[] aliases)
        {
            Scope = scope;
            Row = row;
            Name = name;

            var values = new List<string>();
            values.Add(name);
            if (aliases != null)
            {
                values.AddRange(aliases);
            }

            normalizedAliases = new string[values.Count];
            for (int i = 0; i < values.Count; i++)
            {
                normalizedAliases[i] = NormalizePredictionToken(values[i]);
            }
        }

        public string Scope { get; private set; }
        public int Row { get; private set; }
        public string Name { get; private set; }

        public bool Matches(string normalizedKey)
        {
            foreach (string alias in normalizedAliases)
            {
                if (string.Equals(normalizedKey, alias, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }
    }

    private sealed class ExcelSession : IDisposable
    {
        public Excel.Application App { get; private set; }
        public Excel.Workbook Workbook { get; private set; }
        private bool disposed;
        private ExcelApplicationSession applicationSession;
        private bool ownsApplicationSession;

        public ExcelSession(string workbookPath, bool visible, bool readOnly, PerfTracer tracer)
        {
            applicationSession = new ExcelApplicationSession(visible, tracer);
            ownsApplicationSession = true;
            App = applicationSession.App;
            OpenWorkbook(workbookPath, readOnly, tracer);
        }

        public ExcelSession(ExcelApplicationSession applicationSession, string workbookPath, bool readOnly, PerfTracer tracer)
        {
            if (applicationSession == null || applicationSession.App == null)
            {
                throw new ArgumentNullException("applicationSession");
            }

            this.applicationSession = applicationSession;
            ownsApplicationSession = false;
            App = applicationSession.App;
            OpenWorkbook(workbookPath, readOnly, tracer);
        }

        private void OpenWorkbook(string workbookPath, bool readOnly, PerfTracer tracer)
        {
            Workbook = App.Workbooks.Open(Path.GetFullPath(workbookPath), 0, readOnly);
            ((Excel._Workbook)Workbook).Activate();
            InstallDocumentPropertyPatch();
            InitializeModelContextFromWorkbook();
            tracer.Mark("workbook-open");
        }

        public void Dispose()
        {
            if (disposed)
            {
                return;
            }

            disposed = true;
            try
            {
                if (Workbook != null)
                {
                    Workbook.Close(false);
                    Marshal.ReleaseComObject(Workbook);
                }
            }
            finally
            {
                if (ownsApplicationSession && applicationSession != null)
                {
                    applicationSession.Dispose();
                }
            }
        }
    }

    private sealed class ExcelApplicationSession : IDisposable
    {
        public Excel.Application App { get; private set; }
        private bool disposed;
        private int processId;
        private HashSet<int> existingExcelProcessIds;

        public ExcelApplicationSession(bool visible, PerfTracer tracer)
        {
            Directory.SetCurrentDirectory(AddinDir);
            LoadIpConfig();
            tracer.Mark("ip-config");

            existingExcelProcessIds = GetExcelProcessIds();
            App = (Excel.Application)Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            tracer.Mark("excel-created");
            try
            {
                GetWindowThreadProcessId(new IntPtr(App.Hwnd), out processId);
            }
            catch
            {
                processId = 0;
            }
            if (processId <= 0)
            {
                processId = FindNewExcelProcessId();
            }
            WriteExcelPidFile(processId);
            App.Visible = visible;
            App.DisplayAlerts = false;
            try
            {
                App.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            }
            catch
            {
            }

            SinitekExcel.WriterUtil.WriteUtil.Application = App;

            var connect = new SinitekExcel.Connect();
            Array custom = Array.CreateInstance(typeof(object), 0);
            connect.OnConnection(App, Extensibility.ext_ConnectMode.ext_cm_Startup, null, ref custom);
            SinitekExcel.WriterUtil.WriteUtil.Application = App;
            tracer.Mark("addin-connected");
        }

        public void Dispose()
        {
            if (disposed)
            {
                return;
            }

            disposed = true;
            try
            {
                if (App != null)
                {
                    App.Quit();
                    Marshal.ReleaseComObject(App);
                }
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                KillExcelProcessIfStillAlive();
            }
        }

        private void KillExcelProcessIfStillAlive()
        {
            if (processId <= 0)
            {
                return;
            }

            try
            {
                Process process = Process.GetProcessById(processId);
                if (!process.HasExited)
                {
                    process.Kill();
                    process.WaitForExit(5000);
                }
            }
            catch
            {
            }
        }

        private static void WriteExcelPidFile(int excelProcessId)
        {
            if (excelProcessId <= 0)
            {
                return;
            }

            string path = Environment.GetEnvironmentVariable("SINITEK_EXCEL_PID_FILE");
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            try
            {
                string directory = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                File.WriteAllText(path, excelProcessId.ToString(CultureInfo.InvariantCulture), Encoding.UTF8);
            }
            catch
            {
            }
        }

        private static HashSet<int> GetExcelProcessIds()
        {
            var ids = new HashSet<int>();
            foreach (Process process in Process.GetProcessesByName("EXCEL"))
            {
                try
                {
                    ids.Add(process.Id);
                }
                finally
                {
                    process.Dispose();
                }
            }

            return ids;
        }

        private int FindNewExcelProcessId()
        {
            foreach (Process process in Process.GetProcessesByName("EXCEL"))
            {
                try
                {
                    if (existingExcelProcessIds == null || !existingExcelProcessIds.Contains(process.Id))
                    {
                        return process.Id;
                    }
                }
                finally
                {
                    process.Dispose();
                }
            }

            return 0;
        }
    }
}
