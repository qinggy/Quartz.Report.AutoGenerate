using Esd.Cloud.Common;
using Esd.Report.AutoGenerate.Application;
using Esd.Report.AutoGenerate.Application.Model;
using Esd.Report.AutoGenerate.Application.Services;
using Esd.Report.AutoGenerate.Service;
using EsdPec.Cloud.BusinessImpl.NewReportCenter;
using EsdPec.Cloud.DbModel;
using EsdPec.Cloud.Infrastructure.Utilities;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Esd.Report.AutoGenerate.Application.Service
{
    public class ReportAutoGenerateService : IReportAutoGenerateService
    {
        private const string DatabaseFile = "database.config";
        private readonly string REPORT_TPL_START_FLAG = "#{{start}}#";
        private readonly DatabaseExpression sqlExpression = DatabaseFile.XmlToObject<DatabaseExpression>();
        private readonly RptInformationService rptInformationService
            = new RptInformationService();

        public void AutoGenerate(JobDetail jobDetail)
        {
            Task.Factory.StartNew(() =>
            {
                return GenerateReportPdf(jobDetail.RptId, jobDetail.CompanyId);
            }).ContinueWith(r =>
            {
                var emailDetail = new Database().DbContext.Database.SqlQuery<EmailDetail>(sqlExpression.GetEmailSet).ToList();
                //CommHelper.SendMailUseGmail();
            });
        }

        public string GenerateReportPdf(Guid rptId, Guid companyId)
        {
            var currentRptInfo = rptInformationService.GetCurrentRptInformation(rptId);
            var filePath = Path.Combine(Report.AutoGenerate.Service.AppSettings.GetValue("xlsSaveDirectory"), currentRptInfo.RptPath);
            if (!System.IO.File.Exists(filePath))
                throw new NullReferenceException("The report file does not exist");
            var startTime = GetStartTime((RptTypeEnum)currentRptInfo.RptType);
            IWorkbook workbook = GetExportWorkBook(filePath, (int)currentRptInfo.RptType, "2", startTime, startTime, false, companyId);
            if (workbook == null)
                throw new NullReferenceException("the workbook is null");
            ReplaceSensitiveDataToEmpty(workbook);
            MemoryStream mStream = new MemoryStream();
            workbook.Write(mStream);
            mStream.Flush();
            workbook.Close();
            mStream.Close();
            return CommHelper.ConvertXlsToPdf(mStream, companyId.ToString(), currentRptInfo.Name);
        }

        private string GetStartTime(RptTypeEnum type)
        {
            switch (type)
            {
                case RptTypeEnum.MonthCrossDay:
                    return DateTime.Now.ToString("yyyy-MM");
                case RptTypeEnum.YearCrossMonth:
                case RptTypeEnum.YoY:
                    return DateTime.Now.ToString("yyyy");
                case RptTypeEnum.DayCheckTable:
                    return DateTime.Now.ToString("yyyy-MM-dd");
                default:
                    return DateTime.Now.ToString("yyyy-MM-dd");
            }
        }

        private void ReplaceSensitiveDataToEmpty(IWorkbook workbook)
        {
            ISheet sheet = workbook.GetSheetAt(0);
            IRow row = null;
            ICellStyle cellStyle = GetCellNoneStyle(workbook);
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                row = sheet.GetRow(i);
                if (row == null) continue;
                for (int j = 0; j < 4; j++)
                {
                    ICell cell = row.GetCell(j);
                    if (cell == null) continue;
                    cell.SetCellValue("");
                    cell.CellStyle = cellStyle;
                }
            }
        }

        private ICellStyle GetCellNoneStyle(IWorkbook workbook)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();
            return cellStyle;
        }

        private HSSFWorkbook GetExportWorkBook(string filePath, int rptType, string precision,
            string startTime, string endTime, bool isShouldResizeColumn, Guid companyId)
        {
            HSSFWorkbook workbook = new HSSFWorkbook(new FileStream(filePath, FileMode.Open, FileAccess.Read));
            ISheet sheet = workbook.GetSheetAt(0);
            HSSFRow row = (HSSFRow)sheet.GetRow(0);
            row.ZeroHeight = true;
            sheet.SetColumnHidden(0, true);
            sheet.SetColumnHidden(1, true);
            sheet.SetColumnHidden(2, true);
            sheet.SetColumnHidden(3, true);
            GetReportDataSource(sheet, workbook, rptType, precision, startTime, endTime, isShouldResizeColumn, companyId);
            sheet.ForceFormulaRecalculation = true;
            return workbook;
        }

        private void GetReportDataSource(ISheet sheet, HSSFWorkbook workbook, int rptType, string precision,
            string startTime, string endTime, bool isShouldResizeColumn, Guid companyId)
        {
            List<Guid> bmfIds = new List<Guid>();
            GetBmfIds(sheet, bmfIds);

            Trans trans = new Trans();
            DateTime stime = DateTime.Now;
            DateTime etime = DateTime.Now;
            switch (rptType)
            {
                #region 年(每月)报表
                case 1:
                    trans.stime = startTime;
                    trans.etime = trans.stime;
                    var monthRecords = DataHelper.GetMonRecordByWhere(trans, bmfIds, companyId);
                    SheetVariableReplace(sheet, workbook, monthRecords.OrderBy(a => a.HTime).ToList(),
                        new DateTime(int.Parse(trans.stime), 1, 1),
                        new DateTime(int.Parse(trans.stime), 12, 31),
                        rptType, precision, isShouldResizeColumn);
                    break;
                #endregion
                #region 月(每日)报表
                case 2:
                    var dateArr = startTime.Split('-');
                    var year = dateArr[0];
                    var month = dateArr[1];
                    var days = DateTime.DaysInMonth(int.Parse(year), int.Parse(month));
                    trans.stime = year + "-" + month + "-01";
                    trans.etime = year + "-" + month + "-" + days;
                    var dayRecords = DataHelper.GetDayRecordByWhere(trans, bmfIds, companyId);
                    SheetVariableReplace(sheet, workbook, dayRecords.OrderBy(a => a.HTime).ToList(),
                        DateTime.Parse(trans.stime), DateTime.Parse(trans.etime),
                        rptType, precision, isShouldResizeColumn);
                    break;
                #endregion
                #region 同比报表
                case 3:
                    trans.stime = startTime;
                    trans.etime = trans.stime;
                    var yearMonthRecords = DataHelper.GetMonRecordByWhere(trans, bmfIds, companyId);
                    trans.stime = (Convert.ToInt32(trans.stime) - 1).ToString();
                    trans.tempetime = trans.stime;
                    var lastYearMonthRecords = DataHelper.GetMonRecordByWhere(trans, bmfIds, companyId);
                    var yoyList = new List<ComparsionDto>();
                    DateTime yoysTime = new DateTime(int.Parse(startTime), 1, 1);
                    DateTime yoyeTime = new DateTime(int.Parse(startTime), 12, 1);
                    DateTime lastHtime = yoysTime.AddYears(-1);
                    foreach (var bmfId in bmfIds)
                    {
                        lastHtime = yoysTime.AddYears(-1);
                        for (DateTime sstime = yoysTime; sstime <= yoyeTime; sstime = sstime.AddMonths(1), lastHtime = lastHtime.AddMonths(1))
                        {
                            var mfId = bmfId.ToString();
                            double totalData = 0.0d;
                            var currentData = yearMonthRecords.FirstOrDefault(a => a.MfId == mfId && a.HTime == sstime);
                            var lastData = lastYearMonthRecords.FirstOrDefault(a => a.MfId == mfId && a.HTime == lastHtime);
                            if (lastData == null || lastData.TotalData == 0)
                            {
                                if (currentData != null && currentData.TotalData != 0)
                                {
                                    totalData = 100;
                                }
                                else
                                {
                                    totalData = 0;
                                }
                            }
                            else if (currentData == null || currentData.TotalData == 0)
                            {
                                if (lastData != null && lastData.TotalData != 0)
                                {
                                    totalData = -100;
                                }
                                else
                                {
                                    totalData = 0;
                                }
                            }
                            else
                            {
                                totalData = (currentData.TotalData - lastData.TotalData) / lastData.TotalData;
                            }

                            ComparsionDto record = new ComparsionDto
                            {
                                HTime = sstime,
                                MfId = mfId,
                                CurrentData = currentData == null ? 0 : currentData.TotalData,
                                LastData = lastData == null ? 0 : lastData.TotalData,
                                Rate = totalData,
                            };
                            yoyList.Add(record);
                        }
                    }
                    SheetVariableReplace(sheet, workbook, yoyList.OrderBy(a => a.HTime).ToList(),
                        yoysTime, yoyeTime, rptType, precision, isShouldResizeColumn);
                    break;
                #endregion
                #region 环比
                case 4:
                    bmfIds = bmfIds.Distinct().ToList();
                    var now = DateTime.Now;
                    var yesterday = now.AddDays(-1);
                    var qoqmonth = DateTime.Parse(now.ToString("yyyy-MM-01"));
                    var lastMonth = qoqmonth.AddMonths(-1);
                    List<ComparsionClass> cmpList = new List<ComparsionClass>();
                    foreach (var mfId in bmfIds)
                    {
                        BaseRecord dayRecord = DataHelper.GetDayRecord(mfId, companyId, now);
                        ComparsionClass comparsion = new ComparsionClass();
                        comparsion.CurrentVal = dayRecord == null ? 0 : dayRecord.TotalData;
                        dayRecord = DataHelper.GetDayRecord(mfId, companyId, yesterday);
                        comparsion.YesterDayVal = dayRecord == null ? 0 : dayRecord.TotalData;
                        dayRecord = DataHelper.GetMonthRecord(mfId, companyId, qoqmonth);
                        comparsion.CurrentMonthVal = dayRecord == null ? 0 : dayRecord.TotalData;
                        dayRecord = DataHelper.GetMonthRecord(mfId, companyId, lastMonth);
                        comparsion.LastMonthVal = dayRecord == null ? 0 : dayRecord.TotalData;
                        comparsion.ReferenceId = mfId;
                        cmpList.Add(comparsion);
                    }
                    SheetVariableReplace(sheet, workbook, cmpList, DateTime.Now, DateTime.Now,
                        rptType, precision, isShouldResizeColumn);
                    break;
                #endregion
                #region 日抄表能耗报表
                case 5:
                    var factory = ServiceFactory.CreateService<ICollectValuesService>();
                    var time = DateTime.Parse(startTime);
                    stime = time.AddSeconds(-1);
                    List<RealTimeData> datas = new List<RealTimeData>();
                    foreach (var bmfId in bmfIds)
                    {
                        RealTimeData data = new RealTimeData();
                        var autodata = factory.GetEarliestAutoDatasByMfIdLessThanTime(companyId, bmfId, stime);//cy
                        if (autodata != null)
                            data.LastTime = autodata.Mt.ToString("yyyy-MM-dd HH:mm:ss");

                        data.PreData = autodata == null ? 0 : autodata.Mv;
                        var nextautodata = factory.GetEarliestAutoDatasByMfIdLessThanTime(companyId, bmfId, DateTime.Now);//cy
                        data.NextData = nextautodata == null ? 0 : nextautodata.Mv;
                        data.mftId = bmfId.ToString();
                        data.Sum = data.NextData - data.PreData;
                        if (nextautodata != null)
                            data.Htime = nextautodata.Mt.ToString("yyyy-MM-dd HH:mm:ss");//当前时间
                        datas.Add(data);
                    }
                    factory.CloseService();
                    SheetVariableReplace(sheet, workbook, datas, DateTime.Now, DateTime.Now,
                       rptType, precision, isShouldResizeColumn);
                    break;
                #endregion
                #region 峰谷平能耗报表
                case 6:
                #endregion
                #region 日跨时间范围报表
                case 7:
                #endregion
                #region 班组能耗报表
                case 8:
                #endregion
                #region 时间点抄表报表
                case 9:
                #endregion
                #region 灵活式结构
                case 10:
                #endregion
                #region 不定类型报表
                case 11:
                    break;
                    #endregion
            }
        }

        private static void GetBmfIds(ISheet sheet, List<Guid> bmfIds)
        {
            IRow row = null;
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                row = sheet.GetRow(i);
                if (row != null)
                {
                    ICell cell = row.GetCell(1);
                    if (cell == null) continue;
                    if (cell.CellType == CellType.String)
                    {
                        var strBmfId = cell.StringCellValue;
                        if (!string.IsNullOrEmpty(strBmfId))
                        {
                            Guid bmfId = Guid.Empty;
                            if (Guid.TryParse(strBmfId, out bmfId))
                            {
                                bmfIds.Add(bmfId);
                            }
                        }
                    }
                }
            }
        }

        private bool IsGuid(string strSrc)
        {
            Guid g = Guid.Empty;
            return Guid.TryParse(strSrc, out g);
        }

        private int GetCellWidth(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    var numWidth = cell.NumericCellValue.ToString().Length * 10;
                    return numWidth > 60 ? numWidth > 255 ? 235 : numWidth : 60;
                case CellType.String:
                    var width = cell.StringCellValue.ToString().Length * 10;
                    return width > 60 ? width > 255 ? 235 : width : 60;
                default:
                    return 60;
            }
        }

        private DateTime GetTimeStep(int type, DateTime startTime, ref string format)
        {
            switch (type)
            {
                case 1:
                case 3:
                    format = "yyyy";
                    return startTime.AddYears(1);
                case 2:
                    format = "yyyy-MM";
                    return startTime.AddMonths(1);
                case 5:
                    format = "yyyy-MM-dd";
                    return startTime.AddDays(1);
                default:
                    format = "yyyy-MM-dd";
                    return startTime.AddDays(1);
            }
        }

        private string GetDateFormat(int type)
        {
            switch (type)
            {
                case 1:
                case 3:
                    return "yyyy";
                case 2:
                    return "yyyy-MM";
                case 5:
                    return "yyyy-MM-dd";
                default:
                    return "yyyy-MM-dd";
            }
        }

        private void DrawTableHeader(DateTime sTime, DateTime eTime, IWorkbook workbook, ISheet sheet, int reportType,
            int startColumnIndex, int rowIndex, dynamic valuePair, string precision)
        {
            int startIndex = startColumnIndex;
            IRow row = sheet.GetRow(rowIndex - 1);
            IRow dateRow = sheet.GetRow(rowIndex - 2);
            ICellStyle cellstyle = GetCellStyle(workbook, precision);
            switch (reportType)
            {
                #region 年（每月）
                case 1:
                    for (int i = 1; i <= 12; i++, startIndex++)
                    {
                        ICell dateCell = dateRow.CreateCell(startIndex);
                        dateCell.SetCellValue(i + "月");
                        dateCell.CellStyle = cellstyle;
                        dateCell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                        dateCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                        ICell eCell = row.CreateCell(startIndex);
                        eCell.SetCellValue("能耗");
                        eCell.CellStyle = cellstyle;
                        eCell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                        eCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                    }
                    break;
                #endregion
                #region 月（每日）
                case 2:
                    var year = sTime.Year;
                    var month = sTime.Month;
                    var days = DateTime.DaysInMonth(year, month);
                    for (int i = 1; i <= days; i++, startIndex++)
                    {
                        ICell dateCell = dateRow.CreateCell(startIndex);
                        dateCell.SetCellValue(i + "号");
                        dateCell.CellStyle = cellstyle;
                        dateCell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                        dateCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                        ICell eCell = row.CreateCell(startIndex);
                        eCell.SetCellValue("能耗");
                        eCell.CellStyle = cellstyle;
                        eCell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                        eCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                    }
                    break;
                #endregion
                #region 同比
                case 3:
                    for (int i = 1; i <= 12; i++, startIndex++)
                    {
                        ICell dateCell = dateRow.CreateCell(startIndex);
                        dateCell.SetCellValue(i + "月");
                        CellRangeAddress dateCellRegion = new CellRangeAddress(rowIndex - 2, rowIndex - 2, startIndex, startIndex + 2);
                        sheet.AddMergedRegion(dateCellRegion);
                        dateCell.CellStyle = cellstyle;
                        dateCell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                        dateCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                        ICell eCell = row.CreateCell(startIndex);
                        eCell.SetCellValue("今年");
                        eCell.CellStyle = cellstyle;
                        eCell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                        eCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                        startIndex++;
                        ICell yeCell = row.CreateCell(startIndex);
                        yeCell.SetCellValue("去年");
                        yeCell.CellStyle = cellstyle;
                        yeCell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                        yeCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                        startIndex++;
                        ICell qeCell = row.CreateCell(startIndex);
                        qeCell.SetCellValue("同比(%)");
                        qeCell.CellStyle = cellstyle;
                        qeCell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                        qeCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                    }
                    ICell totalDateCell = dateRow.CreateCell(startIndex);
                    totalDateCell.SetCellValue("合计");
                    CellRangeAddress totalDateCellRegion = new CellRangeAddress(rowIndex - 2, rowIndex - 2, startIndex, startIndex + 2);
                    sheet.AddMergedRegion(totalDateCellRegion);
                    totalDateCell.CellStyle = cellstyle;
                    totalDateCell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                    totalDateCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                    ICell yearCell = row.CreateCell(startIndex);
                    yearCell.SetCellValue("今年");
                    yearCell.CellStyle = cellstyle;
                    yearCell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                    yearCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                    startIndex++;
                    ICell lastYearCell = row.CreateCell(startIndex);
                    lastYearCell.SetCellValue("去年");
                    lastYearCell.CellStyle = cellstyle;
                    lastYearCell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                    lastYearCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                    startIndex++;
                    ICell totaleCell = row.CreateCell(startIndex);
                    totaleCell.SetCellValue("同比(%)");
                    totaleCell.CellStyle = cellstyle;
                    totaleCell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                    totaleCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                    break;
                #endregion
                #region 环比
                case 4:
                    List<string> hbHeaderList = new List<string> { "昨日", "今日", "环比(%)", "本月", "上月", "环比(%)" };
                    IRow hbRow = sheet.GetRow(rowIndex - 2);
                    foreach (var header in hbHeaderList)
                    {
                        ICell cell = hbRow.CreateCell(startIndex);
                        sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 2, rowIndex - 1, startIndex, startIndex));
                        cell.SetCellValue(header);
                        cell.CellStyle = cellstyle;
                        cell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                        cell.CellStyle.FillPattern = FillPattern.SolidForeground;
                        startIndex++;
                    }
                    break;
                #endregion
                #region 日抄表能耗
                case 5:
                    List<string> rcHeaderList = new List<string> { "上一表底", "上一表底时间", "当前表底时间", "当前表底", "能耗" };
                    IRow rcbRow = sheet.GetRow(rowIndex - 2);
                    foreach (var header in rcHeaderList)
                    {
                        ICell cell = rcbRow.CreateCell(startIndex);
                        sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 2, rowIndex - 1, startIndex, startIndex));
                        cell.SetCellValue(header);
                        cell.CellStyle = cellstyle;
                        cell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                        cell.CellStyle.FillPattern = FillPattern.SolidForeground;
                        startIndex++;
                    }
                    break;
                #endregion
                #region 峰谷平
                case 6:
                    var fgpSection = (List<string>)valuePair.sectionList;
                    string titleFormat = GetDateFormat(reportType);
                    int mergeStartIndex = startIndex;
                    for (DateTime issTime = sTime; issTime <= eTime; issTime = GetTimeStep(reportType, issTime, ref titleFormat), startIndex++)
                    {
                        foreach (var fgp in fgpSection)
                        {
                            ICell fgpCell = row.CreateCell(startIndex);
                            startIndex++;
                            fgpCell.SetCellValue(fgp);
                            fgpCell.CellStyle = cellstyle;
                            fgpCell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                            fgpCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                        }
                        ICell eCell = row.CreateCell(startIndex);
                        eCell.SetCellValue("总能耗");
                        eCell.CellStyle = cellstyle;
                        eCell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                        eCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                        ICell dateCell = dateRow.CreateCell(mergeStartIndex);
                        var region = new CellRangeAddress(rowIndex - 2, rowIndex - 2, mergeStartIndex, startIndex);
                        sheet.AddMergedRegion(region);
                        dateCell.SetCellValue(issTime.ToString(titleFormat));
                        dateCell.CellStyle = cellstyle;
                        dateCell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                        dateCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                        sheet.SetColumnWidth(startIndex, 60 * 256);
                        mergeStartIndex = startIndex + 1;
                        for (int i = region.FirstRow; i <= region.LastRow; i++)
                        {
                            IRow rowRegion = HSSFCellUtil.GetRow(i, (HSSFSheet)sheet);
                            for (int j = region.FirstColumn; j <= region.LastColumn; j++)
                            {
                                ICell singleCell = HSSFCellUtil.GetCell(rowRegion, (short)j);
                                singleCell.CellStyle = cellstyle;
                                dateCell.CellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                                dateCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                            }
                        }
                    }
                    break;
                #endregion
                #region 时间跨范围
                case 7:
                #endregion
                #region 班组能耗
                case 8:
                #endregion
                #region 时间点抄表
                case 9:
                #endregion
                #region 灵活式
                case 10:
                    break;
                    #endregion
            }
        }

        private double GetPrecision(string precision, double totalData)
        {
            if (string.IsNullOrEmpty(precision)) precision = "2";
            switch (precision)
            {
                case "0":
                case "2":
                    return Math.Round(totalData, 2);
                case "1":
                    return Math.Round(totalData, 1);
                case "3":
                    return Math.Round(totalData, 3);
                case "4":
                    return Math.Round(totalData, 4);
            }
            return Math.Round(totalData, 2);
        }

        private string GetFormat(string precision)
        {
            if (string.IsNullOrEmpty(precision)) precision = "2";
            switch (precision)
            {
                case "0":
                case "2":
                    return "0.00";
                case "1":
                    return "0.0";
                case "3":
                    return "0.00";
                case "4":
                    return "0.000";
            }
            return "0.00";
        }

        private ICellStyle GetCellStyle(IWorkbook workbook, string precision, bool ifShowBg = false, bool ifDrw = false)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat(GetFormat(precision));
            cellStyle.Alignment = HorizontalAlignment.Center;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BottomBorderColor = HSSFColor.Black.Index;
            cellStyle.LeftBorderColor = HSSFColor.Black.Index;
            cellStyle.RightBorderColor = HSSFColor.Black.Index;
            cellStyle.TopBorderColor = HSSFColor.Black.Index;
            if (ifShowBg)
            {
                cellStyle.FillForegroundColor = HSSFColor.Yellow.Index;
                cellStyle.FillPattern = FillPattern.SolidForeground;
            }
            if (ifDrw)
            {
                cellStyle.FillForegroundColor = HSSFColor.Aqua.Index;
                cellStyle.FillPattern = FillPattern.SolidForeground;
            }
            return cellStyle;
        }

        private void ExcelReplaceValueDiffByType(IWorkbook workbook, ISheet sheet, IRow row, dynamic valuePair, DateTime startTime, DateTime endTime,
            int reportType, string bmfId, string precision, int startIndex, ref int endIndex)
        {
            ICellStyle cellstyle = GetCellStyle(workbook, precision);
            switch (reportType)
            {
                #region 年（每月）
                case 1:
                    var monthList = (List<MonRecord>)valuePair;
                    for (DateTime sTime = startTime; sTime <= endTime; sTime = sTime.AddMonths(1), startIndex++)
                    {
                        ICell cell = row.CreateCell(startIndex);
                        var cellValue = monthList.FirstOrDefault(a => a.MfId == bmfId && a.HTime == sTime);
                        if (cellValue == null) cellValue = new MonRecord();
                        cell.SetCellValue(cellValue == null ? 0.00 : GetPrecision(precision, cellValue.TotalData));
                        cell.CellStyle = cellstyle;
                    }
                    endIndex = startIndex;
                    break;
                #endregion
                #region 月（每天）
                case 2:
                    var dayList = (List<DayRecord>)valuePair;
                    for (DateTime sTime = startTime; sTime <= endTime; sTime = sTime.AddDays(1), startIndex++)
                    {
                        ICell cell = row.CreateCell(startIndex);
                        var cellValue = dayList.FirstOrDefault(a => a.MfId == bmfId && a.HTime == sTime);
                        if (cellValue == null) cellValue = new DayRecord();
                        cell.SetCellValue(cellValue != null ? GetPrecision(precision, cellValue.TotalData) : 0.00);
                        cell.CellStyle = cellstyle;
                    }
                    endIndex = startIndex;
                    break;
                #endregion
                #region 同比
                case 3:
                    var yoyList = (List<ComparsionDto>)valuePair;
                    for (DateTime sTime = startTime; sTime <= endTime; sTime = sTime.AddMonths(1), startIndex++)
                    {
                        var yoyEntity = yoyList.FirstOrDefault(a => a.MfId == bmfId && a.HTime == sTime);
                        if (yoyEntity == null) yoyEntity = new ComparsionDto();
                        ICell currentCell = row.CreateCell(startIndex);
                        currentCell.SetCellValue(GetPrecision(precision, yoyEntity.CurrentData));
                        currentCell.CellStyle = cellstyle;
                        startIndex = startIndex + 1;
                        ICell lastCell = row.CreateCell(startIndex);
                        lastCell.SetCellValue(GetPrecision(precision, yoyEntity.LastData));
                        lastCell.CellStyle = cellstyle;
                        startIndex = startIndex + 1;
                        ICell rateCell = row.CreateCell(startIndex);
                        rateCell.SetCellValue(GetPrecision(precision, yoyEntity.Rate));
                        rateCell.CellStyle = cellstyle;
                    }
                    var totalCurrent = yoyList.Where(a => a.MfId == bmfId).Sum(a => a.CurrentData);
                    var totalLast = yoyList.Where(a => a.MfId == bmfId).Sum(a => a.LastData);
                    var totalRate = totalLast == 0 ? 100 : (totalCurrent - totalLast) / totalLast;
                    ICell totalCell = row.CreateCell(startIndex);
                    totalCell.SetCellValue(GetPrecision(precision, totalCurrent));
                    totalCell.CellStyle = cellstyle;
                    startIndex = startIndex + 1;
                    ICell totalLastCell = row.CreateCell(startIndex);
                    totalLastCell.SetCellValue(GetPrecision(precision, totalLast));
                    totalLastCell.CellStyle = cellstyle;
                    startIndex = startIndex + 1;
                    ICell totalRateCell = row.CreateCell(startIndex);
                    totalRateCell.SetCellValue(GetPrecision(precision, totalRate));
                    totalRateCell.CellStyle = cellstyle;
                    endIndex = startIndex;
                    break;
                #endregion
                #region 环比
                case 4:
                    var qoqList = (List<ComparsionClass>)valuePair;
                    var mfId = new Guid(bmfId);
                    var qoqEntity = qoqList.FirstOrDefault(a => a.ReferenceId == mfId);
                    if (qoqEntity == null) qoqEntity = new ComparsionClass();
                    ICell yesterdayValue = row.CreateCell(startIndex);
                    yesterdayValue.SetCellValue(GetPrecision(precision, qoqEntity.YesterDayVal));
                    yesterdayValue.CellStyle = cellstyle;
                    startIndex++;
                    ICell todayValue = row.CreateCell(startIndex);
                    todayValue.SetCellValue(GetPrecision(precision, qoqEntity.CurrentVal));
                    todayValue.CellStyle = cellstyle;
                    startIndex++;
                    ICell qoqRate = row.CreateCell(startIndex);
                    var rateValue = qoqEntity.YesterDayVal == 0 ? 0 : GetPrecision(precision, (qoqEntity.CurrentVal - qoqEntity.YesterDayVal) * 100 / qoqEntity.YesterDayVal);
                    qoqRate.SetCellValue(rateValue);
                    qoqRate.CellStyle = cellstyle;
                    startIndex++;
                    ICell monthValue = row.CreateCell(startIndex);
                    monthValue.SetCellValue(GetPrecision(precision, qoqEntity.CurrentMonthVal));
                    monthValue.CellStyle = cellstyle;
                    startIndex++;
                    ICell lastMonthValue = row.CreateCell(startIndex);
                    lastMonthValue.SetCellValue(GetPrecision(precision, qoqEntity.LastMonthVal));
                    lastMonthValue.CellStyle = cellstyle;
                    startIndex++;
                    ICell qoqMonthRate = row.CreateCell(startIndex);
                    var monthRateValue = qoqEntity.LastMonthVal == 0 ? 0.0 : GetPrecision(precision, (qoqEntity.CurrentMonthVal - qoqEntity.LastMonthVal) * 100 / qoqEntity.LastMonthVal);
                    qoqMonthRate.SetCellValue(monthRateValue);
                    qoqMonthRate.CellStyle = cellstyle;

                    endIndex = startIndex;
                    break;
                #endregion
                #region 日抄表能耗
                case 5:
                    var dayCheckList = (List<RealTimeData>)valuePair;
                    var dayCheckEntity = dayCheckList.FirstOrDefault(a => a.mftId == bmfId);
                    if (dayCheckEntity == null) dayCheckEntity = new RealTimeData();
                    ICell preValue = row.CreateCell(startIndex);
                    preValue.SetCellValue(GetPrecision(precision, dayCheckEntity.PreData));
                    preValue.CellStyle = cellstyle;
                    startIndex++;
                    ICell preDateValue = row.CreateCell(startIndex);
                    preDateValue.SetCellValue(dayCheckEntity.LastTime);
                    preDateValue.CellStyle = cellstyle;
                    startIndex++;
                    ICell currentDateValue = row.CreateCell(startIndex);
                    currentDateValue.SetCellValue(dayCheckEntity.Htime);
                    currentDateValue.CellStyle = cellstyle;
                    startIndex++;
                    ICell currentValue = row.CreateCell(startIndex);
                    currentValue.SetCellValue(GetPrecision(precision, dayCheckEntity.NextData));
                    currentValue.CellStyle = cellstyle;
                    startIndex++;
                    ICell energyValue = row.CreateCell(startIndex);
                    energyValue.SetCellValue(GetPrecision(precision, dayCheckEntity.Sum));
                    energyValue.CellStyle = cellstyle;
                    startIndex++;

                    endIndex = startIndex;
                    break;
                #endregion
                #region 峰谷平
                case 6:
                #endregion
                #region 日跨时间范围
                case 7:
                #endregion
                #region 班组能耗
                case 8:
                #endregion
                #region 时间点
                case 9:
                #endregion
                #region 灵活式
                case 10:
                    break;
                    #endregion
            }
        }

        private void SheetVariableReplace(ISheet sheet, IWorkbook workbook, dynamic valuePair,
            DateTime startTime, DateTime endTime, int reportType, string precision, bool isShouldResizeColumn = true)
        {
            bool hadBeenDraw = false;
            IRow row;
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                row = sheet.GetRow(i);
                if (row != null)
                {
                    var bmfId = string.Empty;
                    var startIndex = 0;
                    int endIndex = 0;
                    for (int j = 0; j <= row.LastCellNum; j++)
                    {
                        startIndex = j;
                        sheet.AutoSizeColumn(j);
                        ICell cell = row.GetCell(j);
                        if (cell == null) continue;
                        string cellValue = cell.ToString();
                        if (j == 1 && IsGuid(cellValue))
                        {
                            bmfId = cellValue;
                            continue;
                        }
                        if (cellValue == REPORT_TPL_START_FLAG)
                        {
                            if (!hadBeenDraw)
                            {
                                hadBeenDraw = true;
                                DrawTableHeader(startTime, endTime, workbook, sheet, reportType, startIndex, i, valuePair, precision);
                            }
                            ExcelReplaceValueDiffByType(workbook, sheet, row, valuePair, startTime, endTime,
                                reportType, bmfId, precision, startIndex, ref endIndex);
                            j = endIndex;
                        }
                        if (startIndex != j)
                        {
                            cell = row.GetCell(j);
                            if (cell == null) continue;
                        }
                        if (cell.CellType == CellType.Formula)
                        {
                            HSSFFormulaEvaluator eva = new HSSFFormulaEvaluator(workbook);
                            if (cell.ToString().Contains(".xls")) continue;
                            double result = eva.Evaluate(cell).NumberValue;
                            cell.SetCellValue(result);
                            sheet.ForceFormulaRecalculation = true;
                        }
                        if (isShouldResizeColumn)
                            sheet.SetColumnWidth(j, (int)(GetCellWidth(cell)) * 256);
                    }
                }
            }
        }

    }
}
