﻿using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
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

    Dictionary<string, MyWorkDay> dicFirstMyWorkDay = new Dictionary<string, MyWorkDay>();
    Dictionary<string, MyWorkDay> dicSecondMyWorkDay = new Dictionary<string, MyWorkDay>();

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
        int secondExcelFirstDataRow = config.Value.SecondExcelFirstDataRow;
        int secondExcelSiteKeyColumn = config.Value.SecondExcelSiteKeyColumn;
        int secondExcelSiteNameColumn = config.Value.SecondExcelSiteNameColumn;
        int secondExcelWorkDayCountColumn = config.Value.SecondExcelWorkDayCountColumn;
        int secondExcelWorkDaysColumn = config.Value.SecondExcelWorkDaysColumn;
        string ignoreSiteKeySuffix = config.Value.IgnoreSiteKeySuffix;

        Dictionary<string,string> dicIgnoreFirstExcelAtSiteKey = new Dictionary<string, string>();
        string ignoreFirstExcelAtSiteKey = config.Value.IgnoreFirstExcelAtSiteKey;
        foreach (var ignore in ignoreFirstExcelAtSiteKey.Split(','))
        {
            dicIgnoreFirstExcelAtSiteKey.Add(ignore, "");
        }
        FileStream fsFirstExcel = new FileStream(firstexcel, FileMode.Open, FileAccess.Read, FileShare.Read);
        using XLWorkbook xlWorkbookFristExcel = new XLWorkbook(fsFirstExcel);
        IXLWorksheets sheetsFristExcel = xlWorkbookFristExcel.Worksheets;
        foreach (IXLWorksheet? sheet in sheetsFristExcel)
        {
            if (firstExcelSheetName.Equals(sheet.Name))
            {
                int lastUsedRowNumber = sheet.LastRowUsed() == null ? 0 : sheet.LastRowUsed().RowNumber();
                logger.ZLogInformation($"firstexcel シート名:{sheet.Name}, 最後の行:{lastUsedRowNumber}");

                for (int r = firstDataRow; r < lastUsedRowNumber + 1; r++)
                {
                    IXLCell cellWorkDayCount = sheet.Cell(r, workDayCountColumn);
                    int workCount = -1;
                    switch (cellWorkDayCount.DataType)
                    {
                        case XLDataType.Number:
                            workCount = cellWorkDayCount.GetValue<int>();
                            break;
                        default:
                            logger.ZLogError($"workCount is NOT type ( Number ) at siteKey]{sheet.Cell(r, siteKeyColumn).Value.ToString()} sheet:{sheet.Name} row:{r}");
                            continue;
                    }
                    IXLCell cellWorkDaysColumn = sheet.Cell(r, workDaysColumn);
                    string workDays = "";
                    switch (cellWorkDaysColumn.DataType)
                    {
                        case XLDataType.DateTime:
                            workDays = cellWorkDaysColumn.GetValue<DateTime>().ToString("yyyy/MM/dd");
                            break;
                        case XLDataType.Text:
                            workDays = replaceDateTimeString(cellWorkDaysColumn.GetValue<string>());
                            break;
                        case XLDataType.Blank:
                            logger.ZLogTrace($"workDays is Blank type at sheet:{sheet.Name} row:{r}");
                            break;
                        default:
                            logger.ZLogError($"workDays is NOT type ( DateTime | Text ) at sheet:{sheet.Name} row:{r}");
                            continue;
                    }

                    MyWorkDay wd = new MyWorkDay();
                    wd.workDayCount = workCount;
                    wd.siteNumber = sheet.Cell(r, siteNumberColumn).Value.ToString();
                    wd.siteName = convertZero(sheet.Cell(r, siteNameColumn).Value.ToString());
                    wd.siteKey = sheet.Cell(r, siteKeyColumn).Value.ToString();
                    wd.status = sheet.Cell(r, statusColumn).Value.ToString();
                    logger.ZLogTrace($"拠点キー:{wd.siteKey}, 工事日数:{workCount}, 工事日:{workDays}");
                    if (isIgnoreSiteKey(wd.siteKey, dicIgnoreFirstExcelAtSiteKey))
                    {
                        logger.ZLogTrace($"[FitstExcel] 除外しました {wd.siteKey}");
                    }
                    else
                    {
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
                        if (wd.siteKey.EndsWith(ignoreSiteKeySuffix))
                        {
                            continue;
                        }
                        dicFirstMyWorkDay.Add(wd.siteKey, wd);
                    }
                }
            }
            else
            {
                logger.ZLogTrace($"Miss {sheet.Name}");
            }
        }

        Dictionary<string,string> dicIgnoreSecondExcelAtSiteKey = new Dictionary<string, string>();
        string ignoreSecondExcelAtSiteKey = config.Value.IgnoreSecondExcelAtSiteKey;
        foreach (var ignore in ignoreSecondExcelAtSiteKey.Split(','))
        {
            dicIgnoreSecondExcelAtSiteKey.Add(ignore, "");
        }
        FileStream fsSecondExcel = new FileStream(secondexcel, FileMode.Open, FileAccess.Read, FileShare.Read);
        using XLWorkbook xlWorkbookSecondExcel = new XLWorkbook(fsSecondExcel);
        IXLWorksheets sheetsSecondExcel = xlWorkbookSecondExcel.Worksheets;
        foreach (IXLWorksheet? sheet in sheetsSecondExcel)
        {
            if (secondExcelSheetName.Equals(sheet.Name))
            {
                int lastUsedRowNumber = sheet.LastRowUsed() == null ? 0 : sheet.LastRowUsed().RowNumber();
                logger.ZLogInformation($"secondexcel シート名:{sheet.Name}, 最後の行:{lastUsedRowNumber}");

                for (int r = secondExcelFirstDataRow; r < lastUsedRowNumber + 1; r++)
                {
                    IXLCell cellWorkDayCount = sheet.Cell(r, secondExcelWorkDayCountColumn);
                    int workCount = -1;
                    switch (cellWorkDayCount.DataType)
                    {
                        case XLDataType.Number:
                            workCount = cellWorkDayCount.GetValue<int>();
                            break;
                        default:
                            logger.ZLogError($"workCount is NOT type ( Number ) at siteKey]{sheet.Cell(r, siteKeyColumn).Value.ToString()} sheet:{sheet.Name} row:{r}");
                            continue;
                    }
                    IXLCell cellWorkDaysColumn = sheet.Cell(r, secondExcelWorkDaysColumn);
 //                   IXLCell cellWorkDaysColumn = sheet.Cell(r, workDaysColumn);
                    string workDays = "";
                    switch (cellWorkDaysColumn.DataType)
                    {
                        case XLDataType.DateTime:
                            workDays = cellWorkDaysColumn.GetValue<DateTime>().ToString("yyyy/MM/dd");
                            break;
                        case XLDataType.Text:
                            workDays = replaceDateTimeString(cellWorkDaysColumn.GetValue<string>());
                            break;
                        case XLDataType.Blank:
                            logger.ZLogTrace($"workDays is Blank type at sheet:{sheet.Name} row:{r}");
                            break;
                        default:
                            logger.ZLogError($"workDays is NOT type ( DateTime | Text ) at sheet:{sheet.Name} row:{r}");
                            continue;
                    }

                    MyWorkDay wd = new MyWorkDay();
                    wd.workDayCount = workCount;
                    wd.siteName = convertZero(sheet.Cell(r, secondExcelSiteNameColumn).Value.ToString());
                    wd.siteKey = sheet.Cell(r, secondExcelSiteKeyColumn).Value.ToString();
                    logger.ZLogTrace($"拠点キー:{wd.siteKey}, 工事日数:{workCount}, 工事日:{workDays}");
                    if (isIgnoreSiteKey(wd.siteKey, dicIgnoreFirstExcelAtSiteKey))
                    {
                        logger.ZLogTrace($"[SecondExcel] 除外しました {wd.siteKey}");
                    }
                    else
                    {
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
                        if (wd.siteKey.EndsWith(ignoreSiteKeySuffix))
                        {
                            continue;
                        }
                        dicSecondMyWorkDay.Add(wd.siteKey, wd);
                    }
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

//== diff
        diffFirstAndSecond();

//== finish
        if (isAllPass)
        {
            logger.ZLogInformation($"== [Congratulations!] すべての確認項目をパスしました ==");
        }
        logger.ZLogInformation($"==== tool finish ====");
    }

    private string convertZero(string target)
    {
        int index = target.IndexOf('-');
        switch (index)
        {
            case 1:
                return "000"+target;
            case 2:
                return "00"+target;
            case 3:
                return "0"+target;
            default:
                break;
        }
        int index2 = target.IndexOf('_');
        switch (index2)
        {
            case 1:
                return "000"+target;
            case 2:
                return "00"+target;
            case 3:
                return "0"+target;
            default:
                break;
        }
        return target;
    }

    private bool isErrorAtDiffList<T>(string name, List<T> list1, List<T> list2)
    {
        logger.ZLogTrace($"{string.Join("|", list1)}");
        logger.ZLogTrace($"{string.Join("|", list2)}");

        bool isError = false;
        var siteKey12 = list1.Except(list2);
        var siteKey21 = list2.Except(list1);
        if (siteKey12.Count() > 0)
        {
            isError = true;
            logger.ZLogInformation($"不一致が発見されました 比較パラメーター:{name} [1st-2nd] {string.Join("|",siteKey12)}");
        }
        if (siteKey21.Count() > 0)
        {
            isError = true;
            logger.ZLogInformation($"不一致が発見されました 比較パラメーター:{name} [2nd-1st] {string.Join("|",siteKey21)}");
        }

        if (isError)
        {
            return true;
        }
        return false;
    }

    private bool isErrorAtCompareString(string name, string siteKey, string s1, string s2)
    {
        if (!s1.Equals(s2))
        {
            logger.ZLogError($"不一致が発見されました ({name}) {siteKey} 1st:{s1} 2nd:{s2}");
            return true;
        }
        return false;
    }

    private bool isErrorAtCompareInt(string name, string siteKey, int i1, int i2)
    {
        if (!i1.Equals(i2))
        {
            logger.ZLogError($"不一致が発見されました ({name}) {siteKey} 1st:{i1} 2nd:{i2}");
            return true;
        }
        return false;
    }

    private bool isErrorAtCompareListDateTime(string name, string siteKey, int i1, int i2)
    {
        if (!i1.Equals(i2))
        {
            logger.ZLogError($"不一致が発見されました ({name}) {siteKey} 1st:{i1} 2nd:{i2}");
            return true;
        }
        return false;
    }

    private bool isErrorAtCompareList(string name, string siteKey, List<DateTime> d1, List<DateTime> d2)
    {
        bool isError = false;

        if (isErrorAtCompareInt(name+"の合計日数", siteKey, d1.Count, d2.Count))
        {
            isError = true;
        }
        var siteKeyIntersect = d1.Intersect(d2);
        var siteKey12 = d1.Except(d2);
        var siteKey21 = d2.Except(d1);
        if (siteKeyIntersect.Count() > 0)
        {
            logger.ZLogTrace($"一致 ({name}) {siteKey} [1st=2nd] {convertDateTimeToDate(siteKeyIntersect)}");
        }
        if (siteKey12.Count() > 0)
        {
            isError = true;
            logger.ZLogInformation($"不一致が発見されました ({name}) {siteKey} [1st-2nd] {convertDateTimeToDate(siteKey12)}");
        }
        if (siteKey21.Count() > 0)
        {
            isError = true;
            logger.ZLogInformation($"不一致が発見されました ({name}) {siteKey} [2nd-1st] {convertDateTimeToDate(siteKey21)}");
        }

        
        return isError;
    }

    private bool isIgnoreSiteKey(string siteKey, Dictionary<string,string> dicIgnore)
    {
        return dicIgnore.ContainsKey(siteKey);
    }

    private void diffFirstAndSecond()
    {
        logger.ZLogInformation($"== start 2つのExcelファイルの比較 ==");
        bool isDateError = false;
        bool isError = false;
        string checkStatusAtWork = config.Value.CheckStatusAtWork;

        if (isErrorAtDiffList("拠点キー", dicFirstMyWorkDay.Keys.ToList(), dicSecondMyWorkDay.Keys.ToList()))
        {
            isError = true;
        }
        foreach (var key in dicFirstMyWorkDay.Keys)
        {
            MyWorkDay wd1 = dicFirstMyWorkDay[key];
            if (dicSecondMyWorkDay.ContainsKey(key))
            {
                MyWorkDay wd2 = dicSecondMyWorkDay[key];
                if (checkStatusAtWork.Equals(wd1.status))
                {
                    if (isErrorAtCompareString("拠点名", key, wd1.siteName, wd2.siteName) |
                        isErrorAtCompareInt("工事日数", key, wd1.workDayCount, wd2.workDayCount) |
                        isErrorAtCompareList("工事日", key, wd1.workDays, wd2.workDays) )
                    {
                        isError = true;
                    }
                }
            }
            else
            {
                isError = true;
                logger.ZLogError($"[ERROR] 2つ目のExcelに拠点キー({key})が見つかりませんでした");
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
            logger.ZLogInformation($"[NG] 2つのExcelファイルの不一致が発見されました");
        }
        if (!isDateError && !isError)
        {
            logger.ZLogInformation($"[OK] 2つのExcelファイルの不一致はありませんでした");
        }
        logger.ZLogInformation($"== end 2つのExcelファイルの比較 ==");
    }

    private void printToday(string printday)
    {
        logger.ZLogInformation($"== start 工事日の拠点 ==");
        bool isDateError = false;
        string checkStatusAtWork = config.Value.CheckStatusAtWork;
        foreach (var targetday in printday.Split('|'))
        {
            StringBuilder sb = new StringBuilder();
            try
            {
                DateTime day = DateTime.Parse(targetday);
                sb.AppendLine($"");
                sb.AppendLine($"{convertDateTimeToDateAndDayofweek(day)} の拠点は以下です");
                sb.AppendLine($"");
                foreach (var workDay in dicFirstMyWorkDay.Values.ToList())
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
                        int max = workDay.workDays.Count;
                        if (index > 0)
                        {
                            sb.AppendLine($"{workDay.siteName} ({index}/{max})");
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
        }
        logger.ZLogInformation($"== end 工事日の拠点 ==");
    }

    private void checkWorkDayCount()
    {
        logger.ZLogInformation($"== start 工事日数と工事日の日数一致の確認 ==");
        bool isDateError = false;
        bool isError = false;
        string checkStatusAtWork = config.Value.CheckStatusAtWork;
        foreach (var workDay in dicFirstMyWorkDay.Values.ToList())
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
        foreach (var workDay in dicFirstMyWorkDay.Values.ToList())
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
        foreach (var workDay in dicFirstMyWorkDay.Values.ToList())
        {
            logger.ZLogDebug($"キー:{workDay.siteKey},拠点名:{workDay.siteName},ステータス:{workDay.status},工事日:{workDay.workDayCount},workDays:{convertDateTimeToDate(workDay.workDays)}");
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
                sb.Append("|");
            }
        }
        return sb.ToString();
    }

    private string convertDateTimeToDate(IEnumerable<DateTime> dateTimes)
    {
        List<DateTime> tmpList = dateTimes.ToList();
        tmpList.Sort();
        return convertDateTimeToDate(tmpList);
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
    
    public int SecondExcelFirstDataRow {get; set;} = -1;
    public int SecondExcelSiteKeyColumn {get; set;} = -1;
    public int SecondExcelSiteNameColumn {get; set;} = -1;
    public int SecondExcelWorkDayCountColumn {get; set;} = -1;
    public int SecondExcelWorkDaysColumn {get; set;} = -1;
    public string IgnoreSiteKeySuffix {get; set;} = "";
    public string IgnoreFirstExcelAtSiteKey {get; set;} = "";
    public string IgnoreSecondExcelAtSiteKey {get; set;} = "";
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