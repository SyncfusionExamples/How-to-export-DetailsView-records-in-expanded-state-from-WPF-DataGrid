using Microsoft.Win32;
using Syncfusion.UI.Xaml.Grid.Converter;
using Syncfusion.XlsIO;
using System.IO;
using System.Windows;

namespace DataGridWPFSample
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private void ExportToExcelOnClicked(object sender, RoutedEventArgs e)
        {
            ExcelExportingOptions exportoption = new ExcelExportingOptions();
            var excelengine = dataGrid.ExportToExcel(dataGrid.View, exportoption);
            var workbook = excelengine.Excel.Workbooks[0];
            workbook.Worksheets[0].Range.ExpandGroup(ExcelGroupBy.ByRows, ExpandCollapseFlags.IncludeSubgroups);
            SaveFileDialog sfd = new SaveFileDialog
            {
                FilterIndex = 2,
                Filter = "Excel 97 to 2003 Files(*.xls)|*.xls|Excel 2007 to 2010 Files(*.xlsx)|*.xlsx|Excel 2013 File(*.xlsx)|*.xlsx"
            };
            if (sfd.ShowDialog() == true)
            {
                using (Stream stream = sfd.OpenFile())
                {
                    if (sfd.FilterIndex == 1)
                        workbook.Version = ExcelVersion.Excel97to2003;
                    else if (sfd.FilterIndex == 2)
                        workbook.Version = ExcelVersion.Excel2010;
                    else
                        workbook.Version = ExcelVersion.Excel2013;
                    workbook.SaveAs(stream);
                }
                if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                {
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
            }
        }
    }
}