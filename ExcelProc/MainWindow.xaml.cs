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
            PreviewFile.DataContext = tbl.DefaultView;
        }

        private void BtnExportPreview_Click(object sender, RoutedEventArgs e)
        {
            PreviewFile.Visibility = Visibility.Visible;
            var tbl = new DataTable();
            string str = string.Empty;
            if (App.Current.Properties["esName"].ToString() != string.Empty)
                str = App.Current.Properties["esName"].ToString();
            tbl = PreviewExcel(ImportPath.Text, str);
            PreviewFile.DataContext = tbl.DefaultView;
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            fbd.ShowDialog();
            DialogResult result = fbd.ShowDialog();

            if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
                string filepath = fbd.SelectedPath;
                ImportPath.FontStyle = FontStyles.Italic;
                ExportPath.Text = filepath;
            }
        }

        private void GetImpSheetName(string path)
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
        }

        private DataTable PreviewExcel(string path,string sName)
        {

            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets["SOW-PO Tracker 2016"];

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

        //Export to excel
        private void ExportToExcel(DataTable dt)
        {

            /*Set up work book, work sheets, and excel application*/
            
            try
            {
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while exporting excel file.");
            }
        }

        private void ImportSheetName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            App.Current.Properties["isName"] = ImportSheetName.SelectedItem;
        }
    }
}
