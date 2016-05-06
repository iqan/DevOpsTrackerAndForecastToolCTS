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
            tbl = PreviewExcel(ImportPath.Text, str);
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
            tbl = PreviewExcel(ImportPath.Text, "Sheet1");
            if (tbl != null)
            {
                PreviewFile.DataContext = tbl.DefaultView;
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = "Document"; // Default file name
            dlg.DefaultExt = ".xls"; // Default file extension
            dlg.Filter = "Excel documents (.xls, .xlsx)|*.xls;*.xlsx"; // Filter files by extension

            // Show save file dialog box
            dlg.ShowDialog();

            // Process save file dialog box results
            // Save document
            App.Current.Properties["esName"] += dlg.FileName;
            ExportPath.FontStyle = FontStyles.Italic;
            ExportPath.Text = App.Current.Properties["esName"].ToString();
            if (App.Current.Properties["esName"].ToString() != string.Empty)
                BtnExport_Do.IsEnabled = true;
        }

        private void BtnExport_Do_Click(object sender, RoutedEventArgs e)
        {
            string impSheet = string.Empty;
            string destUrl = App.Current.Properties["esName"].ToString();

            if (App.Current.Properties["isName"].ToString() != string.Empty)
                impSheet = App.Current.Properties["isName"].ToString();
            DataTable dt = PreviewExcel(ImportPath.Text, impSheet);
            ExportToExcel(dt, destUrl);
        }
        #endregion

        //Common Methods
        private DataTable PreviewExcel(string path,string sName)
        {
            try
            {
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {
                    using (var stream = File.OpenRead(path))
                    {
                        pck.Load(stream);
                    }
                    var ws = pck.Workbook.Worksheets["Sheet1"];

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

        //Export to excel
        private void ExportToExcel(DataTable dt, string dest)
        {

            /*Set up work book, work sheets, and excel application*/
            try
            {
                using (ExcelPackage xp = new ExcelPackage())
                {
                    using (var stream = File.OpenWrite(dest))
                    {
                        xp.Load(stream);
                    }
                    string tableName = "Sheet1";
                    if (App.Current.Properties["isName"].ToString() != string.Empty)
                        tableName = App.Current.Properties["isName"].ToString();
                    ExcelWorksheet ws = xp.Workbook.Worksheets.Add(tableName);

                    int rowstart = 2;
                    int colstart = 2;
                    int rowend = rowstart;
                    int colend = colstart + dt.Columns.Count;

                    ws.Cells[rowstart, colstart, rowend, colend].Merge = true;
                    ws.Cells[rowstart, colstart, rowend, colend].Value = dt.TableName;
                    ws.Cells[rowstart, colstart, rowend, colend].Style.HorizontalAlignment =
                        OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    ws.Cells[rowstart, colstart, rowend, colend].Style.Font.Bold = true;
                    ws.Cells[rowstart, colstart, rowend, colend].Style.Fill.PatternType =
                        OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    ws.Cells[rowstart, colstart, rowend, colend].Style.Fill.BackgroundColor.SetColor(
                        System.Drawing.Color.LightGray);

                    rowstart += 2;
                    rowend = rowstart + dt.Rows.Count;
                    ws.Cells[rowstart, colstart].LoadFromDataTable(dt, true);
                    int i = 1;
                    foreach (DataColumn dc in dt.Columns)
                    {
                        i++;
                        if (dc.DataType == typeof (decimal))
                            ws.Column(i).Style.Numberformat.Format = "#0.00";
                    }
                    ws.Cells[ws.Dimension.Address].AutoFitColumns();



                    ws.Cells[rowstart, colstart, rowend, colend].Style.Border.Top.Style =
                        ws.Cells[rowstart, colstart, rowend, colend].Style.Border.Bottom.Style =
                            ws.Cells[rowstart, colstart, rowend, colend].Style.Border.Left.Style =
                                ws.Cells[rowstart, colstart, rowend, colend].Style.Border.Right.Style =
                                    OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,"Error While Exporting");
                BtnPreviewExport.IsEnabled = false;  
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
    }
}
