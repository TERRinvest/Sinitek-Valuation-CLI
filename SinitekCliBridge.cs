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

public sealed class FakeRibbonControl : IRibbonControl
{
    private readonly string id;

    public FakeRibbonControl(string id)
    {
        this.id = id;
    }

    public string Id { get { return id; } }
    public object Context { get { return null; } }
    public string Tag { get { return null; } }
}

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

    [DllImport("user32.dll")]
    private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool VirtualProtect(IntPtr lpAddress, UIntPtr dwSize, uint flNewProtect, out uint lpflOldProtect);

    private static bool documentPropertyPatchInstalled;

    public static string Inspect(string workbookPath, bool visible)
    {
        using (var session = new ExcelSession(workbookPath, visible, true))
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
            return string.Join(Environment.NewLine, lines);
        }
    }

    public static string Login(string username, string password)
    {
        LoadIpConfig();
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

        return string.Join(Environment.NewLine, lines);
    }

    public static string StockSearch(string workbookPath, string query, int count, string username, string password)
    {
        using (var session = new ExcelSession(workbookPath, false, true))
        {
            LoginIfProvided(username, password);
            EnsureTokenForStockSearch();

            var stocks = SearchStocks(query, count, ResolveCurrentModelVersion(), ResolveCurrentModelType());
            return FormatStockSearchResult(stocks);
        }
    }

    private static string FormatStockSearchResult(List<Sinitek.WriterModel.StockData> stocks)
    {
        if (stocks.Count == 0)
        {
            return "No stocks returned.";
        }

        var lines = new List<string>();
        foreach (var stock in stocks)
        {
            lines.Add(string.Join("\t", new[]
            {
                NullToEmpty(stock.StockCode),
                NullToEmpty(stock.StockName),
                NullToEmpty(stock.Gsdm),
                NullToEmpty(stock.MarketCode),
                NullToEmpty(stock.MarketName),
                NullToEmpty(stock.SecurityCode),
                NullToEmpty(stock.IndustryCode)
            }));
        }

        return string.Join(Environment.NewLine, lines);
    }

    public static string InvokeButton(
        string workbookPath,
        string outWorkbook,
        bool saveOriginal,
        bool visible,
        string buttonId,
        string username,
        string password)
    {
        if (string.IsNullOrWhiteSpace(buttonId))
        {
            throw new ArgumentException("Button id is required.");
        }

        using (var session = new ExcelSession(workbookPath, visible, false))
        {
            LoginIfProvided(username, password);
            var connect = new SinitekExcel.Connect();
            connect.btnOnAction(new FakeRibbonControl(buttonId));
            SaveIfRequested(session.Workbook, outWorkbook, saveOriginal);
            return "ButtonInvoked=" + buttonId;
        }
    }

    public static string InvokeHandler(
        string workbookPath,
        string outWorkbook,
        bool saveOriginal,
        bool visible,
        string handlerType,
        string username,
        string password)
    {
        if (string.IsNullOrWhiteSpace(handlerType))
        {
            throw new ArgumentException("Handler type is required.");
        }

        using (var session = new ExcelSession(workbookPath, visible, false))
        {
            LoginIfProvided(username, password);
            var handler = CreateHandler(handlerType);
            handler.handle();
            SaveIfRequested(session.Workbook, outWorkbook, saveOriginal);
            return "HandlerInvoked=" + handlerType;
        }
    }

    public static string OutputDirect(
        string workbookPath,
        string outWorkbook,
        bool saveOriginal,
        bool visible,
        string stockCode,
        string currencyUnit,
        string username,
        string password)
    {
        using (var session = new ExcelSession(workbookPath, visible, false))
        {
            LoginIfProvided(username, password);
            EnsureLoggedIn();
            CurrencyUnitSpec currency = ApplyCurrencyUnit(session.Workbook, currencyUnit);
            SinitekExcel.WriterUtil.ModelUtil.CurrencyUnit = currency.Scale;
            SetDocProperty("CurrencyUnit", currency.ScaleText);

            string modelStock = string.IsNullOrWhiteSpace(stockCode) ? GetDocProperty("StkCode") : stockCode.Trim();
            if (string.IsNullOrWhiteSpace(modelStock))
            {
                throw new InvalidOperationException("Stock code is required. Pass -Stock or run update first.");
            }

            session.App.CalculateFull();
            string suffix = ResolveMailSuffix(username);
            string handlerType = ResolveHandlerType("btnOutput");
            object handler = Activator.CreateInstance(ResolveType(handlerType));
            bool ok = (bool)InvokeMethod(handler, "GetOutputSheet", true, suffix, modelStock);
            SaveIfRequested(session.Workbook, outWorkbook, saveOriginal);
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
        string companyManagementType,
        string companyManagementName,
        string peerStock,
        bool updateDirectory,
        bool updateSrcData,
        bool migrate,
        bool addOutput,
        string username,
        string password)
    {
        using (var session = new ExcelSession(workbookPath, visible, false))
        {
            LoginIfProvided(username, password);
            EnsureLoggedIn();

            var stock = ResolveStock(stockCode, gsdm, stockName);
            string handlerType = ResolveHandlerType("btnUpdate");
            object handler = Activator.CreateInstance(ResolveType(handlerType));

            SinitekExcel.WriterUtil.ModelUtil.StockCode = stock.StockCode;
            SinitekExcel.WriterUtil.ModelUtil.StockName = stock.StockName;
            SinitekExcel.WriterUtil.ModelUtil.Gsdm = stock.Gsdm;
            SinitekExcel.WriterUtil.ModelUtil.HistoryYear = historyYear;
            ApplyYearSettings(session.Workbook, historyYear, forecastYear);
            CurrencyUnitSpec currency = ApplyCurrencyUnit(session.Workbook, currencyUnit);
            SinitekExcel.WriterUtil.ModelUtil.CurrencyUnit = currency.Scale;
            SinitekExcel.WriterUtil.ModelUtil.UpdateCheck = true;

            SetDocProperty("StkCode", stock.StockCode);
            SetDocProperty("GSCode", stock.Gsdm);
            SetDocProperty("HistoryYear", historyYear.ToString(CultureInfo.InvariantCulture));
            SetDocProperty("ForecastYear", forecastYear.ToString(CultureInfo.InvariantCulture));
            SetDocProperty("CurrencyUnit", currency.ScaleText);
            SetDocProperty("CompanyManagementName", string.IsNullOrWhiteSpace(companyManagementName) ? "\u6309\u4ea7\u54c1" : companyManagementName);
            SetDocProperty("CompanyManagementType", companyManagementType);
            SetDocProperty("PeerStock", peerStock ?? string.Empty);
            SetDocProperty("UpdateDirectory", updateDirectory.ToString(CultureInfo.InvariantCulture));
            SetDocProperty("UpdateSrcData", updateSrcData.ToString(CultureInfo.InvariantCulture));

            var reportYear = DateTime.Now.Year - 1;
            int lastYearFlag = (int)InvokeMethod(handler, "IsLastYear", false, stock.StockCode, reportYear);
            if (lastYearFlag == -1)
            {
                throw new InvalidOperationException("Failed to validate latest annual report period.");
            }

            int lastYear = lastYearFlag > 0 ? reportYear : reportYear - 1;
            SinitekExcel.WriterUtil.ModelUtil.LastYear = lastYear;

            int migrateYears = 0;
            string previousYear = GetDocProperty("year1");
            if (!string.IsNullOrWhiteSpace(previousYear))
            {
                SetDocProperty("lastUpdateYear", previousYear);
                migrateYears = lastYear - SafeInt(previousYear, lastYear);
            }
            SetDocProperty("year1", lastYear.ToString(CultureInfo.InvariantCulture));

            string lastReportDate = (string)InvokeMethod(handler, "GetLastReportDate", false, stock.StockCode);
            if (string.IsNullOrWhiteSpace(lastReportDate))
            {
                throw new InvalidOperationException("Failed to get latest report date for " + stock.StockCode + ".");
            }

            var failures = new List<string>();
            Excel.XlCalculation oldCalculation = session.App.Calculation;
            try
            {
                session.App.Calculation = Excel.XlCalculation.xlCalculationManual;

                if (updateDirectory)
                {
                    if (!(bool)InvokeMethod(handler, "UpdateDirectory", false, lastReportDate))
                    {
                        failures.Add("directory");
                    }
                }

                if (updateSrcData)
                {
                    SetDocProperty("base_stkcode", stock.StockCode);
                    if (!(bool)InvokeMethod(handler, "UpdateSrcData", false, lastReportDate))
                    {
                        failures.Add("src-data");
                    }
                    if (!(bool)InvokeMethod(handler, "UpdateFinancialNotes", false))
                    {
                        failures.Add("financial-notes");
                    }
                    if (SheetExists(session.Workbook, "Company operating data") || SheetExists(session.Workbook, "公司经营数据"))
                    {
                        if (!(bool)InvokeMethod(handler, "UpdateCompanyManagement", false))
                        {
                            failures.Add("company-management");
                        }
                    }
                    if (!(bool)InvokeMethod(handler, "UpdatePeerAnalysis", false, lastReportDate))
                    {
                        failures.Add("peer-analysis");
                    }
                }

                if (migrate && migrateYears > 0)
                {
                    if (!(bool)InvokeMethod(handler, "MigrateData", false, migrateYears))
                    {
                        failures.Add("migrate");
                    }
                }

                if (addOutput)
                {
                    string outputType = ResolveHandlerType("btnOutput");
                    object outputHandler = Activator.CreateInstance(ResolveType(outputType));
                    if (!(bool)InvokeMethod(outputHandler, "GetOutputSheet", true, ResolveMailSuffix(username), stock.StockCode))
                    {
                        failures.Add("output");
                    }
                }

                session.App.CalculateFull();
            }
            finally
            {
                session.App.Calculation = oldCalculation;
            }

            SaveIfRequested(session.Workbook, outWorkbook, saveOriginal);

            var lines = new List<string>();
            lines.Add("UpdateDirect=" + (failures.Count == 0));
            lines.Add("Handler=" + handlerType);
            lines.Add("Stock=" + stock.StockCode);
            lines.Add("GSCode=" + stock.Gsdm);
            lines.Add("LastReportDate=" + lastReportDate);
            lines.Add("LastYear=" + lastYear.ToString(CultureInfo.InvariantCulture));
            lines.Add("HistoryYear=" + historyYear.ToString(CultureInfo.InvariantCulture));
            lines.Add("ForecastYear=" + forecastYear.ToString(CultureInfo.InvariantCulture));
            lines.Add("CurrencyUnit=" + currency.CurrencyUnitName);
            lines.Add("CurrencyUnitScale=" + currency.ScaleText);
            if (failures.Count > 0)
            {
                lines.Add("Failures=" + string.Join(",", failures));
            }

            return string.Join(Environment.NewLine, lines);
        }
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

    private static List<Sinitek.WriterModel.StockData> SearchStocks(string query, int count, string modelVersion, string modelType)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            throw new ArgumentException("Stock search query is required.");
        }

        string url = (SinitekExcel.WriterUtil.ModelUtil.ModelUrl ?? string.Empty).TrimEnd('/') + "/api/stock";
        string postBody = BuildStockSearchPostBody(url, query.Trim(), count <= 0 ? 10 : count, modelVersion, modelType);

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
                string json = reader.ReadToEnd();
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
        }
        catch (WebException ex)
        {
            if (ex.Status == WebExceptionStatus.Timeout)
            {
                throw new TimeoutException("Stock search timed out after "
                    + (StockSearchTimeoutMs / 1000).ToString(CultureInfo.InvariantCulture)
                    + " seconds: " + url, ex);
            }

            string detail = ReadWebExceptionBody(ex);
            if (!string.IsNullOrWhiteSpace(detail))
            {
                detail = " Response: " + detail;
            }
            throw new InvalidOperationException("Stock search request failed: " + ex.Message + detail, ex);
        }
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
            Login(username, password);
        }
    }

    private static void EnsureTokenForStockSearch()
    {
        string tokenId = Safe(() => SinitekExcel.WriterUtil.WriteUtil.TokenId);
        if (string.IsNullOrWhiteSpace(tokenId) || tokenId.StartsWith("ERR:", StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Stock search requires login token. Set SINITEK_USERNAME/SINITEK_PASSWORD, or pass -Username and -Password.");
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

    private static SinitekExcel.IWriterInterface.IHandler CreateHandler(string handlerType)
    {
        object instance = Activator.CreateInstance(ResolveType(handlerType));
        var handler = instance as SinitekExcel.IWriterInterface.IHandler;
        if (handler == null)
        {
            throw new InvalidOperationException(handlerType + " is not an IHandler.");
        }

        return handler;
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

    private static object GetStaticProperty(Type type, string name)
    {
        PropertyInfo property = type.GetProperty(name, BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
        if (property == null)
        {
            throw new MissingMemberException(type.FullName, name);
        }

        return property.GetValue(null, null);
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

    private static double SafeDouble(string value, double fallback)
    {
        double parsed;
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed) ? parsed : fallback;
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

    private sealed class ExcelSession : IDisposable
    {
        public Excel.Application App { get; private set; }
        public Excel.Workbook Workbook { get; private set; }
        private bool disposed;
        private int processId;
        private HashSet<int> existingExcelProcessIds;

        public ExcelSession(string workbookPath, bool visible, bool readOnly)
        {
            Directory.SetCurrentDirectory(AddinDir);
            LoadIpConfig();

            existingExcelProcessIds = GetExcelProcessIds();
            App = (Excel.Application)Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
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

            Workbook = App.Workbooks.Open(Path.GetFullPath(workbookPath), 0, readOnly);
            ((Excel._Workbook)Workbook).Activate();
            InstallDocumentPropertyPatch();
            InitializeModelContextFromWorkbook();
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
                if (App != null)
                {
                    App.Quit();
                    Marshal.ReleaseComObject(App);
                }
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
