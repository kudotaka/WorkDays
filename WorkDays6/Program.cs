﻿using System.Text;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Utf8StringInterpolation;
using ZLogger;
using ZLogger.Providers;

//==
var builder = ConsoleApp.CreateBuilder(args);
builder.ConfigureServices((ctx,services) =>
{
    // Register appconfig.json to IOption<MyConfig>
    services.Configure<MyConfig>(ctx.Configuration);

    // Using Cysharp/ZLogger for logging to file
    services.AddLogging(logging =>
    {
        logging.ClearProviders();
        logging.SetMinimumLevel(LogLevel.Trace);
        var jstTimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time");
        var utcTimeZoneInfo = TimeZoneInfo.Utc;
        logging.AddZLoggerConsole(options =>
        {
            options.UsePlainTextFormatter(formatter => 
            {
                formatter.SetPrefixFormatter($"{0:yyyy-MM-dd'T'HH:mm:sszzz}|{1:short}|", (in MessageTemplate template, in LogInfo info) => template.Format(TimeZoneInfo.ConvertTime(info.Timestamp.Utc, jstTimeZoneInfo), info.LogLevel));
                formatter.SetExceptionFormatter((writer, ex) => Utf8String.Format(writer, $"{ex.Message}"));
            });
        });
        logging.AddZLoggerRollingFile(options =>
        {
            options.UsePlainTextFormatter(formatter => 
            {
                formatter.SetPrefixFormatter($"{0:yyyy-MM-dd'T'HH:mm:sszzz}|{1:short}|", (in MessageTemplate template, in LogInfo info) => template.Format(TimeZoneInfo.ConvertTime(info.Timestamp.Utc, jstTimeZoneInfo), info.LogLevel));
                formatter.SetExceptionFormatter((writer, ex) => Utf8String.Format(writer, $"{ex.Message}"));
            });

            // File name determined by parameters to be rotated
            options.FilePathSelector = (timestamp, sequenceNumber) => $"logs/{timestamp.ToLocalTime():yyyy-MM-dd}_{sequenceNumber:00}.log";
            
            // The period of time for which you want to rotate files at time intervals.
            options.RollingInterval = RollingInterval.Day;
            
            // Limit of size if you want to rotate by file size. (KB)
            options.RollingSizeKB = 1024;        
        });
    });
});

var app = builder.Build();
app.AddCommands<WorkDaysApp>();
app.Run();


public class WorkDaysApp : ConsoleAppBase
{
    bool isAllPass = true;

    readonly ILogger<WorkDaysApp> logger;
    readonly IOptions<MyConfig> config;

    List<MyWorkDay> listMyWorkDay = new List<MyWorkDay>();

    public WorkDaysApp(ILogger<WorkDaysApp> logger,IOptions<MyConfig> config)
    {
        this.logger = logger;
        this.config = config;
    }

//    [Command("")]
    public void Days(string firstexcel, string secondexcel, string printday)
    {
//== start
        logger.ZLogInformation($"==== tool {getMyFileVersion()} ====");
        if (!File.Exists(firstexcel))
        {
            logger.ZLogError($"[NG] first excel file is missing.");
            return;
        }
        if (!File.Exists(secondexcel))
        {
            logger.ZLogError($"[NG] second excel file is missing.");
            return;
        }

        string firstExcelSheetName = config.Value.FirstExcelSheetName;
        string secondExcelSheetName = config.Value.SecondExcelSheetName;
        int firstDataRow = config.Value.FirstDataRow;
        int siteKeyColumn = config.Value.SiteKeyColumn;
        int siteNumberColumn = config.Value.SiteNumberColumn;
        int siteNameColumn = config.Value.SiteNameColumn;
        int statusColumn = config.Value.StatusColumn;
        int workDayCountColumn = config.Value.WorkDayCountColumn;
        int workDaysColumn = config.Value.WorkDaysColumn;

        FileStream fsFirstExcel = new FileStream(firstexcel, FileMode.Open, FileAccess.Read, FileShare.Read);
        using XLWorkbook xlWorkbookFristExcel = new XLWorkbook(fsFirstExcel);
        IXLWorksheets sheetsFristExcel = xlWorkbookFristExcel.Worksheets;
        foreach (IXLWorksheet? sheet in sheetsFristExcel)
        {
            if (firstExcelSheetName.Equals(sheet.Name))
            {
                int lastUsedRowNumber = sheet.LastRowUsed() == null ? 0 : sheet.LastRowUsed().RowNumber();
                logger.ZLogInformation($"シート名:{sheet.Name}, 最後の行:{lastUsedRowNumber}");

                for (int r = firstDataRow; r < lastUsedRowNumber + 1; r++)
                {
                    IXLCell cellWorkDayCount = sheet.Cell(r, workDayCountColumn);
                    int workCount = -1;
                    switch (cellWorkDayCount.DataType)
                    {
                        case XLDataType.Number:
                            workCount = cellWorkDayCount.GetValue<int>();
                            break;
                        case XLDataType.Text:
                            break;
                        default:
                            logger.ZLogError($"workCount is NOT type ( Number | Text ) at sheet:{sheet.Name} row:{r}");
                            continue;
                    }
                    IXLCell cellWorkDaysColumn = sheet.Cell(r, workDaysColumn);

                    string workDays = replaceDateTimeString(cellWorkDaysColumn.GetValue<string>());
                    logger.ZLogTrace($"工事日数:{workCount}, 工事日:{workDays}");
                    MyWorkDay wd = new MyWorkDay();
                    wd.workDayCount = workCount;
                    wd.siteNumber = sheet.Cell(r, siteNumberColumn).Value.ToString();
                    wd.siteName = sheet.Cell(r, siteNameColumn).Value.ToString();
                    wd.siteKey = sheet.Cell(r, siteKeyColumn).Value.ToString();
                    wd.status = sheet.Cell(r, statusColumn).Value.ToString();
                    List<DateTime> listDateTime = new List<DateTime>();
                    foreach (var day in workDays.Split("|"))
                    {
                        try
                        {
                            DateTime dt = DateTime.Parse(day);
                            if (listDateTime.Contains(dt))
                            {
                                isAllPass = false;
                                DateTime errDt = new DateTime(1900,1,1);
                                listDateTime.Add(errDt);
                                logger.ZLogError($"[ERROR] 重複した日付を発見しました:{day},key:{wd.siteKey},拠点番号:{wd.siteNumber},拠点名:{wd.siteName}");
                            }
                            else
                            {
                                listDateTime.Add(dt);
                            }
                        }
                        catch (FormatException)
                        {
                            isAllPass = false;
                            DateTime errDt = new DateTime(1900,1,1);
                            listDateTime.Add(errDt);
                            logger.ZLogTrace($"エラー 日付に変換できませんでした:{day},拠点名:{wd.siteName}");
                        }
                        catch (System.Exception)
                        {
                            isAllPass = false;
                            throw;
                        }
                    }
                    listDateTime.Sort();
                    wd.workDays = listDateTime;
                    listMyWorkDay.Add(wd);
                }
            }
            else
            {
                logger.ZLogTrace($"Miss {sheet.Name}");
            }
        }



//== print
        printMyWorkDays();

//== check
        checkWorkDayCount();

//== check
        checkWorkDayAtDayOfWeek();

//== print day
        printToday(printday);

//== finish
        if (isAllPass)
        {
            logger.ZLogInformation($"== [Congratulations!] すべての確認項目をパスしました ==");
        }
        logger.ZLogInformation($"==== tool finish ====");
    }

    private void printToday(string printday)
    {
        logger.ZLogInformation($"== start 工事日の拠点 ==");
        bool isDateError = false;
        StringBuilder sb = new StringBuilder();
        string checkStatusAtWork = config.Value.CheckStatusAtWork;
        try
        {
            DateTime day = DateTime.Parse(printday);
            sb.AppendLine($"");
            sb.AppendLine($"{convertDateTimeToDateAndDayofweek(day)} の拠点は以下です");
            foreach (var workDay in listMyWorkDay)
            {
                if (checkStatusAtWork.Equals(workDay.status))
                {
                    if (workDay.workDays.Contains(new DateTime(1900,1,1)))
                    {
                        isDateError = true;
                        logger.ZLogError($"日付エラー key:{workDay.siteKey},拠点番号:{workDay.siteNumber},拠点名:{workDay.siteName},工事日:{convertDateTimeToDate(workDay.workDays)}");
                        continue;
                    }
                    int index = workDay.workDays.IndexOf(day)+1;
                    if (index > 0)
                    {
                        sb.AppendLine($"{workDay.siteNumber}_{workDay.siteName} ({index})");
                    }
                }
            }
            logger.ZLogInformation($"{sb.ToString()}");
        }
        catch (FormatException)
        {
            isAllPass = false;
            logger.LogError($"[NG] エラー 日付に変換できませんでした:{printday}");
        }
        catch (System.Exception)
        {
            isAllPass = false;
            throw;
        }
        logger.ZLogInformation($"== end 工事日の拠点 ==");
    }

    private void checkWorkDayCount()
    {
        logger.ZLogInformation($"== start 工事日数と工事日の日数一致の確認 ==");
        bool isDateError = false;
        bool isError = false;
        string checkStatusAtWork = config.Value.CheckStatusAtWork;
        foreach (var workDay in listMyWorkDay)
        {
            if (checkStatusAtWork.Equals(workDay.status))
            {
                if (workDay.workDays.Contains(new DateTime(1900,1,1)))
                {
                    isDateError = true;
                    logger.ZLogError($"日付エラー key:{workDay.siteKey},拠点番号:{workDay.siteNumber},拠点名:{workDay.siteName},工事日:{convertDateTimeToDate(workDay.workDays)}");
                    continue;
                }
                if (workDay.workDayCount == workDay.workDays.Count)
                {
                    logger.ZLogTrace($"[checkWorkDayCount] 一致");
                }
                else
                {
                    isError = true;
                    logger.ZLogError($"不一致エラー 拠点番号:{workDay.siteNumber},拠点名:{workDay.siteName},工事日数:{workDay.workDayCount},工事日:{convertDateTimeToDate(workDay.workDays)}");
                }
            }
        }
        if (isDateError)
        {
            isAllPass = false;
            logger.ZLogError($"[ERROR] 日付にエラーが発見されました");
        }
        if (isError)
        {
            isAllPass = false;
            logger.ZLogInformation($"[NG] 工事日数と工事日の日数の不一致が発見されました");
        }
        if (!isDateError && !isError)
        {
            logger.ZLogInformation($"[OK] 工事日数と工事日の日数の不一致はありませんでした");
        }
        logger.ZLogInformation($"== end 工事日数と工事日の日数一致の確認 ==");
    }

    private void checkWorkDayAtDayOfWeek()
    {
        logger.ZLogInformation($"== start 工事日と曜日の確認 ==");
        bool isDateError = false;
        bool isWarning = false;
        Dictionary<string,DateTime> dicHolidays = new Dictionary<string, DateTime>();
        string checkStatusAtWork = config.Value.CheckStatusAtWork;
        string publicHolidaysInJapan = config.Value.PublicHolidaysInJapan;
        string bussinessHolidays = config.Value.BussinessHolidays;
        foreach (var holiday in publicHolidaysInJapan.Split('|'))
        {
            dicHolidays.Add(holiday, DateTime.Parse(holiday));
        }
        foreach (var holiday in bussinessHolidays.Split('|'))
        {
            dicHolidays.Add(holiday, DateTime.Parse(holiday));
        }
        logger.ZLogTrace($"[checkWorkDayAtDayOfWeek] {bussinessHolidays}");
        logger.ZLogTrace($"[checkWorkDayAtDayOfWeek] {convertDateTimeToDate(dicHolidays.Values.ToList<DateTime>())}");
        foreach (var workDay in listMyWorkDay)
        {
            if (checkStatusAtWork.Equals(workDay.status))
            {
                if (workDay.workDays.Contains(new DateTime(1900,1,1)))
                {
                    isDateError = true;
                    logger.ZLogError($"日付エラー key:{workDay.siteKey},拠点番号:{workDay.siteNumber},拠点名:{workDay.siteName},工事日:{convertDateTimeToDate(workDay.workDays)}");
                    continue;
                }
                foreach (var day in workDay.workDays)
                {
                    if (dicHolidays.ContainsKey(day.ToString("yyyy/MM/dd")))
                    {
                        isWarning = true;
                        logger.ZLogWarning($"要注意！ 休日:{day.ToString("yyyy/MM/dd")},拠点番号:{workDay.siteNumber},拠点名:{workDay.siteName}");
                    }
                    else
                    {
                        switch (day.DayOfWeek)
                        {
                            case DayOfWeek.Sunday:
                                isWarning = true;
                                logger.ZLogWarning($"要注意！ 日曜:{day.ToString("yyyy/MM/dd")},拠点番号:{workDay.siteNumber},拠点名:{workDay.siteName}");
                                break;
                            case DayOfWeek.Saturday:
                                isWarning = true;
                                logger.ZLogWarning($"要注意！ 土曜:{day.ToString("yyyy/MM/dd")},拠点番号:{workDay.siteNumber},拠点名:{workDay.siteName}");
                                break;
                            default:
//                                logger.ZLogTrace($"[checkWorkDayAtDayOfWeek] 平日:{day.ToString("yyyy/MM/dd")}");
                                break;
                        }
                    }
                }
            }
            else
            {
                logger.ZLogTrace($"除外しました key:{workDay.siteKey},status:{workDay.status}");
            }
        }

        if (isDateError)
        {
            isAllPass = false;
            logger.ZLogError($"[ERROR] 日付にエラーが発見されました");
        }
        if (isWarning)
        {
            isAllPass = false;
            logger.ZLogInformation($"[WARNING] 工事日と曜日に土日祝が発見されました");
        }
        if (!isDateError && !isWarning)
        {
            logger.ZLogInformation($"[OK] 工事日と曜日に土日祝は含まれていませんでした");
        }
        logger.ZLogInformation($"== end 工事日と曜日の確認 ==");
    }

    private void printMyWorkDays()
    {
        logger.ZLogTrace($"== start print ==");
        foreach (var workDay in listMyWorkDay)
        {
//            logger.ZLogTrace($"workDayCount:{workDay.workDayCount},workDays:{string.Join(";",workDay.workDays)}");
            logger.ZLogTrace($"siteKey:{workDay.siteKey},siteNumber:{workDay.siteNumber},siteName:{workDay.siteName},status:{workDay.status},workDayCount:{workDay.workDayCount},workDays:{convertDateTimeToDate(workDay.workDays)}");
        }
        logger.ZLogTrace($"== end print ==");
    }

    private string convertDateTimeToDate(List<DateTime> listDateTime)
    {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < listDateTime.Count; i++)
        {
            sb.Append(listDateTime[i].ToString("yyyy/MM/dd"));
            if (i < listDateTime.Count - 1)
            {
                sb.Append(" & ");
            }
        }
        return sb.ToString();
    }

    private string convertDateTimeToDateAndDayofweek(DateTime day)
    {
        StringBuilder sb = new StringBuilder();
        switch (day.DayOfWeek)
        {
        case DayOfWeek.Sunday:
            sb.Append(day.ToString("yyyy/MM/dd(日)"));
            break;
        case DayOfWeek.Monday:
            sb.Append(day.ToString("yyyy/MM/dd(月)"));
            break;
        case DayOfWeek.Tuesday:
            sb.Append(day.ToString("yyyy/MM/dd(火)"));
            break;
        case DayOfWeek.Wednesday:
            sb.Append(day.ToString("yyyy/MM/dd(水)"));
            break;
        case DayOfWeek.Thursday:
            sb.Append(day.ToString("yyyy/MM/dd(木)"));
            break;
        case DayOfWeek.Friday:
            sb.Append(day.ToString("yyyy/MM/dd(金)"));
            break;
        case DayOfWeek.Saturday:
            sb.Append(day.ToString("yyyy/MM/dd(土)"));
            break;
        }
        return sb.ToString();
    }

    private string replaceDateTimeString(string dateTimeString)
    {
        return dateTimeString.Replace(" ","").Replace("、",",").Replace(",","|");
    }

    private string getTime()
    {
        var jstTimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time");
        return TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, jstTimeZoneInfo).ToString("yyyy-MM-dd'T'HH:mm:sszzz");
    }

    private string getMyFileVersion()
    {
        System.Diagnostics.FileVersionInfo ver = System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location);
        return ver.InternalName + "(" + ver.FileVersion + ")";
    }
}

//==
public class MyConfig
{
    public int FirstDataRow {get; set;} = -1;
    public int SiteKeyColumn {get; set;} = -1;
    public int SiteNumberColumn {get; set;} = -1;
    public int SiteNameColumn {get; set;} = -1;
    public int StatusColumn {get; set;} = -1;
    public int WorkDayCountColumn {get; set;} = -1;
    public int WorkDaysColumn {get; set;} = -1;
    public string FirstExcelSheetName {get; set;} = "";
    public string SecondExcelSheetName {get; set;} = "";
    public string PublicHolidaysInJapan {get; set;} = "";
    public string BussinessHolidays {get; set;} = "";
    public string CheckStatusAtSurvey {get; set;} = "";
    public string CheckStatusAtWork {get; set;} = "";
    
}

public class MyWorkDay
{
    public string siteKey = "";
    public string siteNumber = "";
    public string siteName = "";
    public string status = "";
    public int workDayCount = -1;
    public List<DateTime> workDays = new List<DateTime>();
}