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
            
            int res = Methods.ExportToExcel(dt, destUrl);

            if (res == -1)
                BtnPreviewExport.IsEnabled = false;
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


        

    }
}
