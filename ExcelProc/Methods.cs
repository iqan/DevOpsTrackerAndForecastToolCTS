using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelProc
{
    class Methods
    {
        public static int ExportToExcel(DataTable dt, string dest)
        {

            /*Set up work book, work sheets, and excel application*/
            try
            {

                var file = new FileInfo(dest);
                using (var xp = new ExcelPackage(file))
                {
                    string tableName = "Forecast_" + DateTime.Today.ToString("dd-MM-yyyy");
                    ExcelWorksheet ws = xp.Workbook.Worksheets.Add(tableName);

                    //Headers
                    ws.Cells["A1"].Value = "Month";
                    ws.Cells["B1"].Value = "Project#";
                    ws.Cells["C1"].Value = "Project Name";
                    ws.Cells["D1"].Value = "Resource Name";
                    ws.Cells["E1"].Value = "Billing Period";
                    ws.Cells["F1"].Value = "Rate";
                    ws.Cells["G1"].Value = "Leaves";
                    ws.Cells["H1"].Value = "Billing days";
                    ws.Cells["I1"].Value = "Billing (Total)";

                    using (var range = ws.Cells["A1:I1"])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightCoral);
                        range.Style.Font.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.ShrinkToFit = false;
                    }

                    //var resources = dt.AsEnumerable().Select(r => r.Field<int>("Resource Name")).ToList();
                    List<Resource> resources = new List<Resource>();

                    foreach (DataRow row in dt.Rows)
                    {
                        if (row["Resource Name"].ToString() != "")
                        {
                            Resource res = new Resource();
                            res.ProjectId = long.Parse((string) row[0]);
                            res.ProjectName = (string) row[1];
                            res.ResourceName = (string) row[2];
                            res.BillingPeriod = "na";
                            res.Rate = int.Parse((string) row[8]);
                            res.Leaves = 0;
                            res.BillingDays = 20;
                            res.TotalBilling = res.Rate*res.BillingDays;
                            res.EndDate = DateTime.Parse((string) row[7]);
                            res.StartDate = DateTime.Parse((string) row[6]);
                            resources.Add(res);
                        }
                    }

                    int i = 2;

                    DateTime fromDate = System.DateTime.Today;
                    DateTime toDate = new DateTime(fromDate.Year +1,3,31);
                    if(App.Current.Properties["FromDate"] != null)
                        fromDate = (DateTime) App.Current.Properties["FromDate"];
                    if(App.Current.Properties["FromDate"] != null)
                        toDate = (DateTime) App.Current.Properties["ToDate"];

                    //int bilDates = BillingDays(fromDate, toDate);

                    for (DateTime index = fromDate; index < toDate; index = index.AddMonths(1))
                    {
                        foreach (var res in resources)
                        {
                            var mn = new DateTimeFormatInfo();
                            int count = 0;
                            int days = 0;

                            DateRange range = new DateRange(res.StartDate, res.EndDate);

                            for (DateTime index2 = index; index2 < index.AddMonths(1); index2 = index2.AddDays(1))
                            {
                                DateTime[] bps = GetBillingPeriodGeneral(index2);

                                if (range.Includes(index))
                                {
                                    if (res.StartDate >= bps[0])
                                        days = BillingDays(res.StartDate, bps[1]);
                                    else if (res.EndDate <= bps[1])
                                        days = BillingDays(bps[0], res.EndDate);
                                    else
                                        days = BillingDays(bps[0], bps[1]);

                                    if (index2 == bps[1] && count == 0)
                                    {
                                        ws.Cells[i, 1].Value = mn.GetAbbreviatedMonthName(index.Month) + "-" +
                                                               index.ToString("yy");
                                        ws.Cells[i, 2].Value = res.ProjectId;
                                        ws.Cells[i, 3].Value = res.ProjectName;
                                        ws.Cells[i, 4].Value = res.ResourceName;
                                        ws.Cells[i, 5].Value = "From " + bps[0].ToString("MMM") + " " + bps[0].Day +
                                                               " till " + bps[1].ToString("MMM") + " " + bps[1].Day;
                                        ws.Cells[i, 6].Value = res.Rate;
                                        ws.Cells[i, 7].Value = res.Leaves;
                                        ws.Cells[i, 8].Value = days;
                                        ws.Cells[i, 9].Value = days*res.Rate;
                                        i++;
                                        count++;
                                    }
                                }
                            }
                        }
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    xp.Save();
                }
                MessageBox.Show("File Exported successfully.", "Export sucessful");
                return 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error While Exporting");
                return -1;
            }
        }

        public static DataTable ExcelSheetToDataTable(string path, string sName)
        {
            try
            {
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {
                    using (var stream = File.OpenRead(path))
                    {
                        pck.Load(stream);
                    }
                    var ws = pck.Workbook.Worksheets.First();

                    if (sName != string.Empty)
                        ws = pck.Workbook.Worksheets[sName];

                    DataTable tbl = new DataTable();
                    foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    {
                        tbl.Columns.Add(firstRowCell.Text);
                    }
                    for (int rowNum = 2; rowNum <= ws.Dimension.End.Row; rowNum++)
                    {
                        var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                        DataRow row = tbl.Rows.Add();
                        foreach (var cell in wsRow)
                        {
                            row[cell.Start.Column - 1] = cell.Text;
                        }
                    }
                    return tbl;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error while reading file!");
                return null;
            }
        }

        #region Billing Period logic
        // Get billing period
        public static DateTime[] GetBillingPeriod(DateTime index)
        {
            List<DateTime> billingPS = new List<DateTime>();
            List<DateTime> billingES = new List<DateTime>();
            billingPS.Add(DateTime.Parse("26-Dec-16"));
            billingES.Add(DateTime.Parse("20-Jan-17"));
            billingPS.Add(DateTime.Parse("23-Jan-17"));
            billingES.Add(DateTime.Parse("24-Feb-17"));
            billingPS.Add(DateTime.Parse("27-Feb-17"));
            billingES.Add(DateTime.Parse("24-Mar-17"));
            billingPS.Add(DateTime.Parse("28-Mar-16"));
            billingES.Add(DateTime.Parse("22-Apr-16"));
            billingPS.Add(DateTime.Parse("25-Apr-16"));
            billingES.Add(DateTime.Parse("27-May-16"));
            billingPS.Add(DateTime.Parse("30-May-16"));
            billingES.Add(DateTime.Parse("24-Jun-16"));
            billingPS.Add(DateTime.Parse("27-Jun-16"));
            billingES.Add(DateTime.Parse("22-Jul-16"));
            billingPS.Add(DateTime.Parse("25-Jul-16"));
            billingES.Add(DateTime.Parse("26-Aug-16"));
            billingPS.Add(DateTime.Parse("29-Aug-16"));
            billingES.Add(DateTime.Parse("23-Sep-16"));
            billingPS.Add(DateTime.Parse("26-Sep-16"));
            billingES.Add(DateTime.Parse("21-Oct-16"));
            billingPS.Add(DateTime.Parse("24-Oct-16"));
            billingES.Add(DateTime.Parse("25-Nov-16"));
            billingPS.Add(DateTime.Parse("28-Nov-16"));
            billingES.Add(DateTime.Parse("23-Dec-16"));

            //return bp[index.Month-1];
            DateTime[] temp = {billingPS[index.Month - 1], billingES[index.Month - 1]};

            return temp;
        }

        public static DateTime[] GetBillingPeriodGeneral(DateTime index)
        {
            List<DateTime> billingPS = new List<DateTime>();
            List<DateTime> billingES = new List<DateTime>();

            DateTime tempData = new DateTime();

            DateTime financialYearStartDate = new DateTime(index.Year, 4, 1);
            DateTime financialYearEndDate = new DateTime(index.Year+1, 3, 31);

            switch (financialYearStartDate.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    tempData = financialYearStartDate;
                    break;
                case DayOfWeek.Tuesday:
                    tempData = financialYearStartDate.AddDays(-1);
                    break;
                case DayOfWeek.Wednesday:
                    tempData = financialYearStartDate.AddDays(-2);
                    break;
                case DayOfWeek.Thursday:
                    tempData = financialYearStartDate.AddDays(-3);
                    break;
                case DayOfWeek.Friday:
                    tempData = financialYearStartDate.AddDays(-4);
                    break;
                case DayOfWeek.Saturday:
                    tempData = financialYearStartDate.AddDays(-5);
                    break;
                case DayOfWeek.Sunday:
                    tempData = financialYearStartDate.AddDays(-6);
                    break;
            }

            billingPS.Add(tempData);
            bool change = false;
            for (DateTime i = tempData; i <= financialYearEndDate;)
            {
                if (change)
                {
                    i = i.AddDays(1);
                    billingPS.Add(i);
                    change = false;
                }
                else
                {
                    if (i.Month == 2 || i.Month == 5 || i.Month == 8 || i.Month == 11)
                        i = i.AddDays(35);
                    else
                        i = i.AddDays(28);
                    billingES.Add(i);
                    change = true;
                }
            }

            DateTime[] temp = new DateTime[2];
            if (index.Month >= 4)
            {
                temp[0] = billingPS[index.Month - 4];
                temp[1] = billingES[index.Month - 4];
            }
            else
            {
                temp[0] = billingPS[index.Month + 8];
                temp[1] = billingES[index.Month + 8];
            }

            return temp;
        }
        #endregion

        #region Billing days count

        //getting total days

        public static int BillingDays(DateTime startDate, DateTime endDate)
        {
            int count = 0;
            for (DateTime index = startDate; index <= endDate; index = index.AddDays(1))
            {
                if (index.DayOfWeek != DayOfWeek.Sunday && index.DayOfWeek != DayOfWeek.Saturday)
                {
                    count++;
                }
            }
            return count;
        }

        public static int BillingDaysWithDateExclusion(DateTime startDate, DateTime endDate, Boolean excludeWeekends,
            List<DateTime> excludeDates)
        {
            int count = 0;
            for (DateTime index = startDate; index < endDate; index = index.AddDays(1))
            {
                if (excludeWeekends && index.DayOfWeek != DayOfWeek.Sunday && index.DayOfWeek != DayOfWeek.Saturday)
                {
                    bool excluded = false;
                    ;
                    for (int i = 0; i < excludeDates.Count; i++)
                    {
                        if (index.Date.CompareTo(excludeDates[i].Date) == 0)
                        {
                            excluded = true;
                            break;
                        }
                    }

                    if (!excluded)
                        count++;
                }
            }
            return count;
        }

        public static int GetBillingDaysInMonth(int m)
        {
            List<int> days = new List<int>();
            days.Add(5);
            days.Add(4);
            days.Add(4);
            days.Add(5);
            days.Add(4);
            days.Add(4);
            days.Add(5);
            days.Add(4);
            days.Add(4);
            days.Add(5);
            days.Add(4);
            days.Add(4);
            return days[m - 1];
        }

        #endregion
    }

    public interface IRange<T>
    {
        T Start { get; }
        T End { get; }
        bool Includes(T value);
        bool Includes(IRange<T> range);
    }
    public class DateRange : IRange<DateTime>
    {
        public DateRange(DateTime start, DateTime end)
        {
            Start = start;
            End = end;
        }

        public DateTime Start { get; private set; }
        public DateTime End { get; private set; }

        public bool Includes(DateTime value)
        {
            return (Start <= value) && (value <= End);
        }

        public bool Includes(IRange<DateTime> range)
        {
            return (Start <= range.Start) && (range.End <= End);
        }

        //usage
        //DateRange range = new DateRange(startDate, endDate);
        //range.Includes(date);
    }
}