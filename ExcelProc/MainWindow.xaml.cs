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
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.YellowGreen);
                        range.Style.Font.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.ShrinkToFit = false;
                    }

                    //var resources = dt.AsEnumerable().Select(r => r.Field<int>("Resource Name")).ToList();
                    List<Resource> resources = new List<Resource>(dt.Rows.Count);
                    
                    foreach (DataRow row in dt.Rows)
                    {
                        if (row["Resource Name"].ToString() != "")
                        {
                            Resource res = new Resource();
                            res.ProjectId = (string)row[0];
                            res.ProjectName = (string)row[1];
                            res.ResourceName =  (string)row[2];
                            res.BillingPeriod = "na";
                            res.Rate = (int)row[8];
                            res.Leaves = 0;
                            res.BillingDays = 20;
                            res.TotalBilling = res.Rate * res.BillingDays;
                            resources.Add(res);
                        }
                    }

                    //DataRow[] dr = dt.AsEnumerable().Where(dr => dr.Field<string>("Resource Name") == "put name");
                    ws.InsertRow(2, resources.Count() + 1);
                    int i = 2;

                    foreach (var res in resources)
                    {
                        ws.Cells[i, 1].Value = "Month";
                        ws.Cells[i, 2].Value = res.ProjectId;
                        ws.Cells[i, 3].Value = res.ProjectName;
                        ws.Cells[i, 4].Value = res.ResourceName;
                        ws.Cells[i, 5].Value = res.BillingPeriod;
                        ws.Cells[i, 6].Value = res.Rate;
                        ws.Cells[i, 7].Value = res.Leaves;
                        ws.Cells[i, 8].Value = res.BillingDays;
                        ws.Cells[i, 9].Value = res.TotalBilling;
                        i++;
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    xp.Save();
                }

                InfoLabel.Content = "File exported successfully.";
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
        #endregion

        #region Set some properties

        public void SetComboboxItems()
        {
            var months = System.Globalization.DateTimeFormatInfo.InvariantInfo.MonthNames;
            FromMonth.Items.Add("Select Month");
            ToMonth.Items.Add("Select Month");
            FromYear.Items.Add("Year");
            ToYear.Items.Add("Year");

            for (int i = 2015; i < 2099; i++)
            {
                FromYear.Items.Add(i.ToString());
                ToYear.Items.Add(i.ToString());
            }

            foreach (var month in months)
            {
                FromMonth.Items.Add(month);
                ToMonth.Items.Add(month);
            }
        }

        #endregion
    }
}
