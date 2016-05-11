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
                            res.ProjectId = long.Parse((string)row[0]);
                            res.ProjectName = (string)row[1];
                            res.ResourceName = (string)row[2];
                            res.BillingPeriod = "na";
                            res.Rate = int.Parse((string)row[8]);
                            res.Leaves = 0;
                            res.BillingDays = 20;
                            res.TotalBilling = res.Rate * res.BillingDays;
                            res.EndDate = DateTime.Parse((string)row[7]);
                            res.StartDate = DateTime.Parse((string)row[6]);
                            resources.Add(res);
                        }
                    }

                    int i = 2;

                    int fy = System.DateTime.Now.Year;
                    int fm = Convert.ToDateTime(System.DateTime.Now).Month; ;
                    int ty = 2017;
                    int tm = 3;
                    if (App.Current.Properties["FromYear"] != null)
                        fy = int.Parse((string)App.Current.Properties["FromYear"]);
                    if (App.Current.Properties["FromMon"] != null)
                        fm = (int)App.Current.Properties["FromMon"];
                    if (App.Current.Properties["ToYear"] != null)
                        ty = int.Parse((string)App.Current.Properties["ToYear"]);
                    int x = 31;
                    if (App.Current.Properties["ToMon"] != null)
                    {
                        tm = (int)App.Current.Properties["ToMon"];

                        switch ((int)App.Current.Properties["ToMon"])
                        {
                            case 1:
                            case 3:
                            case 5:
                            case 7:
                            case 8:
                            case 10:
                            case 12:
                                x = 31;
                                break;
                            case 2:
                                //x = (ty % 4 == 0 && ty %400 ==0)? 29: 28;
                                if (ty % 400 == 0)
                                    x = 29;
                                else if (ty % 100 == 0)
                                    x = 28;
                                else if (ty % 4 == 0)
                                    x = 29;
                                else
                                    x = 28;
                                break;
                            default:
                                x = 30;
                                break;
                        }
                    }
                    DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
                    var cal = dfi.Calendar;
                    var week = cal.GetWeekOfYear(DateTime.Parse("28-Mar-16"), dfi.CalendarWeekRule,
                        dfi.FirstDayOfWeek);

                    DateTime fromDate = new DateTime(fy, fm, 1);
                    DateTime toDate = new DateTime(ty, tm, x);

                    int bilDates = BillingDays(fromDate, toDate);

                    for (DateTime index = fromDate; index < toDate; index = index.AddMonths(1))
                    {
                        foreach (var res in resources)
                        {
                            var mn = new DateTimeFormatInfo();
                            int days = 0;
                            int count = 0;
                            
                            DateRange range = new DateRange(res.StartDate, res.EndDate);

                            for (DateTime index2 = index; index2 < index.AddMonths(1); index2 = index2.AddDays(1))
                            {
                                DateTime[] bps = GetBillingPeriod(index2);
                                 
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
                                        ws.Cells[i, 1].Value = mn.GetAbbreviatedMonthName(index.Month) + "-" + index.ToString("yy");
                                        ws.Cells[i, 2].Value = res.ProjectId;
                                        ws.Cells[i, 3].Value = res.ProjectName;
                                        ws.Cells[i, 4].Value = res.ResourceName;
                                        ws.Cells[i, 5].Value = "From " + bps[0].ToString("MMM") + " " + bps[0].Day + " till " + bps[1].ToString("MMM") + " " + bps[1].Day;
                                        ws.Cells[i, 6].Value = res.Rate;
                                        ws.Cells[i, 7].Value = res.Leaves;
                                        ws.Cells[i, 8].Value = days;
                                        ///ws.Cells[i, 8].Value = GetBillingDaysInMonth(index.Month) * 5;
                                        ws.Cells[i, 9].Value = days * res.Rate;
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
            DateTime[] temp = { billingPS[index.Month - 1], billingES[index.Month - 1] };

            return temp;
        }

        //getting total days

        public static int BillingDays(DateTime startDate, DateTime endDate)
        {
            int count = 0;
            for (DateTime index = startDate; index < endDate; index = index.AddDays(1))
            {
                if (index.DayOfWeek != DayOfWeek.Sunday && index.DayOfWeek != DayOfWeek.Saturday)
                {
                    bool excluded = false;
                    if (!excluded)
                        count++;
                }
            }
            return count;
        }
        public static int BillingDaysWithDateExclusion(DateTime startDate, DateTime endDate, Boolean excludeWeekends, List<DateTime> excludeDates)
        {
            int count = 0;
            for (DateTime index = startDate; index < endDate; index = index.AddDays(1))
            {
                if (excludeWeekends && index.DayOfWeek != DayOfWeek.Sunday && index.DayOfWeek != DayOfWeek.Saturday)
                {
                    bool excluded = false; ;
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
