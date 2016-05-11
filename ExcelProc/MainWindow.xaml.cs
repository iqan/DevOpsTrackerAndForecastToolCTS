using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Data;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Forms;
using MessageBox = System.Windows.MessageBox;
using OfficeOpenXml.Style;
using System.Globalization;

namespace ExcelProc
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            App.Current.Properties["isName"] = string.Empty;
            App.Current.Properties["esName"] = string.Empty;
            InitializeComponent();
            SetComboboxItems();
        }

        #region Import Methods
        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            //// Set filter for file extension and default file extension
            dlg.Filter = "All Excel Files (.xlsx, .xls)|*.xlsx;*.xls";

            // Display OpenFileDialog by calling ShowDialog method
            bool? result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox
            if (result == true)
            {
                // Open document
                string filename = dlg.FileName;
                ImportPath.FontStyle = FontStyles.Italic;
                ImportPath.Text = filename;
                ImportSheetName.Items.Clear();
                ImportSheetName.Items.Add("Select Sheet");
                GetImpSheetName(filename);
            }
        }

        private void BtnImportPreview_Click(object sender, RoutedEventArgs e)
        {
            PreviewFile.Visibility = Visibility.Visible;
            var tbl = new DataTable();
            string str = string.Empty;
            if (App.Current.Properties["isName"].ToString() != string.Empty)
                str = App.Current.Properties["isName"].ToString();
            tbl = ExcelSheetToDataTable(ImportPath.Text, str);
            if (tbl != null)
            {
                PreviewFile.DataContext = tbl.DefaultView;
            }
        }

        private void GetImpSheetName(string path)
        {
            try
            {
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {
                    using (var stream = File.OpenRead(path))
                    {
                        pck.Load(stream);
                    }
                    foreach (var x in pck.Workbook.Worksheets)
                    {
                        ImportSheetName.Items.Add(x.Name);
                    }
                }
                BtnPreviewImport.IsEnabled = true;
                ImportSheetName.IsEnabled = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error while reading file!");
            }
        }
        #endregion

        #region Export Methods
        private void BtnExportPreview_Click(object sender, RoutedEventArgs e)
        {
            PreviewFile.Visibility = Visibility.Visible;
            var tbl = new DataTable();

            //MessageBox.Show("prop esName = " + str);
            tbl = ExcelSheetToDataTable(ExportPath.Text, string.Empty);
            if (tbl != null)
            {
                PreviewFile.DataContext = tbl.DefaultView;
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            PreviewFile.DataContext = null;

            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = "Document"; // Default file name
            dlg.DefaultExt = ".xlsx"; // Default file extension
            dlg.Filter = "Excel documents (.xlsx)|*.xlsx"; // Filter files by extension

            // Show save file dialog box
            dlg.ShowDialog();

            // Process save file dialog box results
            // Save document
            App.Current.Properties["esName"] = dlg.FileName;
            ExportPath.FontStyle = FontStyles.Italic;
            ExportPath.Text = App.Current.Properties["esName"].ToString();
            if (App.Current.Properties["esName"].ToString() != string.Empty)
            {
                BtnExport_Do.IsEnabled = true;
                FromMonth.IsEnabled = true;
                ToMonth.IsEnabled = true;
                FromYear.IsEnabled = true;
                ToYear.IsEnabled = true;
            }   
        }

        private void BtnExport_Do_Click(object sender, RoutedEventArgs e)
        {
            string impSheet = string.Empty;
            string destUrl = App.Current.Properties["esName"].ToString();

            if (App.Current.Properties["isName"].ToString() != string.Empty)
                impSheet = App.Current.Properties["isName"].ToString();
            DataTable dt = ExcelSheetToDataTable(ImportPath.Text, impSheet);
            
            ExportToExcel(dt, destUrl);
        }

        private void ExportToExcel(DataTable dt, string dest)
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
                            res.ResourceName =  (string)row[2];
                            res.BillingPeriod = "na";
                            res.Rate = int.Parse((string)row[8]);
                            res.Leaves = 0;
                            res.BillingDays = 20;
                            res.TotalBilling = res.Rate * res.BillingDays;
                            resources.Add(res);
                        }
                    }
                    
                    int i = 2;

                    int fy = System.DateTime.Now.Year;
                    int fm = Convert.ToDateTime(System.DateTime.Now).Month; ;
                    int ty = 2017;
                    int tm = 3;
                    if (App.Current.Properties["FromYear"] != null)
                        fy = int.Parse((string) App.Current.Properties["FromYear"]);
                    if (App.Current.Properties["FromMon"] != null)
                        fm = (int)App.Current.Properties["FromMon"];
                    if (App.Current.Properties["ToYear"] != null)
                        ty = int.Parse((string)App.Current.Properties["ToYear"]);
                    int x = 31;
                    if (App.Current.Properties["ToMon"] != null)
                    {
                        tm = (int) App.Current.Properties["ToMon"];

                        switch ((int) App.Current.Properties["ToMon"])
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
                                if (ty%400 == 0)
                                    x = 29;
                                else if (ty%100 == 0)
                                    x = 28;
                                else if (ty%4 == 0)
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
                    DateTime toDate = new DateTime(ty,tm,x);

                    int bilDates = BillingDays(fromDate, toDate);

                    for (DateTime index = fromDate; index < toDate; index = index.AddMonths(1))
                    {
                        foreach (var res in resources)
                        {
                            var mn = new System.Globalization.DateTimeFormatInfo();
                            string m = mn.GetAbbreviatedMonthName(index.Month);

                            string billingPeriod = getBillingPeriod(index);
                            ws.Cells[i, 1].Value = mn.GetAbbreviatedMonthName(index.Month) + "-" + index.ToString("yy");
                            ws.Cells[i, 2].Value = res.ProjectId;
                            ws.Cells[i, 3].Value = res.ProjectName;
                            ws.Cells[i, 4].Value = res.ResourceName;
                            ws.Cells[i, 5].Value = getBillingPeriod(index);
                            ws.Cells[i, 6].Value = res.Rate;
                            ws.Cells[i, 7].Value = res.Leaves;
                            ws.Cells[i, 8].Value = GetBillingDaysInMonth(index.Month) * 5;
                            ws.Cells[i, 9].Value = res.TotalBilling;
                            i++;
                        }
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    xp.Save();
                }
                MessageBox.Show("File Exported successfully.", "Export sucessful");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error While Exporting");
                BtnPreviewExport.IsEnabled = false;
            }
        }
        #endregion

        //Common Methods
        private DataTable ExcelSheetToDataTable(string path,string sName)
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
                    if (App.Current.Properties["esName"].ToString() != string.Empty)
                    {
                        BtnPreviewExport.IsEnabled = true;   
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

        #region Import sheet name and export file path with filename
                private void ImportSheetName_SelectionChanged(object sender, SelectionChangedEventArgs e)
                {
                    App.Current.Properties["isName"] = ImportSheetName.SelectedItem;
                }
                private void ExportPath_SelectionChanged(object sender, RoutedEventArgs e)
                {
                    App.Current.Properties["esName"] = ExportPath.Text;
                }
        #endregion

        #region user tips methods
        // User Tips Methods
        private void BtnImport_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "Select a file to Import.";
        }

        private void BtnImport_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "";
        }

        private void BtnExport_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "Select a folder to Export Excel file into.";
        }

        private void BtnExport_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "";
        }

        private void ImportSheetName_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "Select a Sheet.";
        }

        private void ImportSheetName_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "";
        }

        private void BtnPreviewImport_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "Click to see preview of selected worksheet.";
        }

        private void BtnPreviewImport_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "";
        }

        private void BtnPreviewExport_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "Click to see preview of exporting worksheet.";
        }

        private void BtnPreviewExport_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "";
        }

        private void BtnExport_Do_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "Click to export worksheet to file in selected folder.";
        }

        private void BtnExport_Do_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "";
        }

        private void FromMonth_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "Select starting Month.";
        }

        private void FromMonth_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "";
        }

        private void FromYear_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "Select starting Year.";
        }

        private void FromYear_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "";
        }

        private void ToMonth_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "Select ending Month.";
        }

        private void ToMonth_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "";
        }

        private void ToYear_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "Select ending Year.";
        }

        private void ToYear_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            InfoLabel.Content = "";
        }

        #endregion

        #region Set some properties

        public void SetComboboxItems()
        {
            var months = System.Globalization.DateTimeFormatInfo.InvariantInfo.MonthNames;

            FromMonth.Items.Add("Select Month");
            ToMonth.Items.Add("Select Month");
            FromYear.Items.Add("Year");
            ToYear.Items.Add("Year");

            for (int i = 2016; i <= 2017; i++)
            {
                ToYear.Items.Add(i.ToString());
            }

            for (int i = 2016; i <= 2016; i++)
            {
                FromYear.Items.Add(i.ToString());
            }

            foreach (var month in months)
            {
                try
                {
                    int d = Convert.ToDateTime(month + " 01, 2000").Month;
                    if (System.DateTime.Now.Month <= d)
                    {
                        FromMonth.Items.Add(month);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
                
                ToMonth.Items.Add(month);
            }
        }

        #endregion

        #region selecting month and year
        private void FromMonth_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var month = FromMonth.SelectedItem;
            if (!month.ToString().Contains("Select"))
            {
                App.Current.Properties["FromMon"] = DateTime.ParseExact((string)month, "MMMM", null).Month;
            }
        }
        private void FromYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var year = FromYear.SelectedItem;
            if (!year.ToString().Contains("Year"))
                App.Current.Properties["FromYear"] = FromYear.SelectedItem;
        }

        private void ToMonth_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var month = ToMonth.SelectedItem;
            if (!month.ToString().Contains("Select"))
                App.Current.Properties["ToMon"] = DateTime.ParseExact((string)month, "MMMM", null).Month;
        }

        private void ToYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var year = ToYear.SelectedItem;
            if (!year.ToString().Contains("Year"))
                App.Current.Properties["ToYear"] = ToYear.SelectedItem;
        }

        #endregion


        // Get billing period
        private string getBillingPeriod(DateTime index)
        {
            List<string> bp = new List<string>();
            bp.Add("From 26th Dec till 20th Jan");
            bp.Add("From 23rd Jan till 24th Feb");
            bp.Add("From 27th Feb till 24th Mar"); 
            bp.Add("From 28th Mar till 22nd Apr");
            bp.Add("From 25th Apr till 27th May");
            bp.Add("From 30th May till 24th Jun");
            bp.Add("From 27th Jun till 22nd July");
            bp.Add("From 25th Jul till 26th Aug");
            bp.Add("From 29th Aug till 23rd Sep");
            bp.Add("From 26th Sep till 21st Oct");
            bp.Add("From 24th Oct till 25th Nov"); 
            bp.Add("From 28th Nov till 23rd Dec"); 
            
            return bp[index.Month-1];
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

        private int GetBillingDaysInMonth(int m)
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
}
