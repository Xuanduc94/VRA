using Microsoft.Win32;
using System.Windows;
using Viettel_Report_Automation.Controllers;

namespace Viettel_Report_Automation
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private string fileWord = null;
        private string fileExcel = null;
        private string fileChamDiem = null;
        public MainWindow()
        {
            InitializeComponent();
        }
        private void btn_openFile_Click(object sender, RoutedEventArgs e)
        {
            string lbl = "File: ";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word File|*.docx";
            if (openFileDialog.ShowDialog().Value == true)
            {
                if (openFileDialog.FileName.Length > 20)
                {
                    string[] fileName = openFileDialog.FileName.Split("\\");
                    lbl_fileName.Content = lbl + "...../" + fileName[fileName.Length - 1];
                }
                else
                {
                    lbl_fileName.Content = lbl + openFileDialog.FileName;
                }
                this.fileWord = openFileDialog.FileName.ToString();

            }
        }



        private async void btnExtract_Click(object sender, RoutedEventArgs e)
        {


            if (this.fileExcel == null)
            {
                MessageBox.Show("Vui lòng chọn file excel", "Cảnh báo", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else if (this.fileWord == null)
            {
                MessageBox.Show("Vui lòng chọn file word", "Cảnh báo", MessageBoxButton.OK, MessageBoxImage.Warning);

            }
            else
            {
                Progress<string> progress = new Progress<string>(value =>
                {
                    lblProcess.Content = value;

                });
                await Task.Run(() =>
                {
                    ReportExtractController reportExtractController = new ReportExtractController();
                    new SettingController().SettingScore(fileChamDiem, progress);
                    reportExtractController.generateReport(this.fileChamDiem, fileExcel, fileWord, progress);
                });
              
                
            }
        }

        private void btn_openFileExcel_Click(object sender, RoutedEventArgs e)
        {
            string lbl = "File: ";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel File|*.xlsx";
            if (openFileDialog.ShowDialog().Value == true)
            {
                if (openFileDialog.FileName.Length > 20)
                {
                    string[] fileName = openFileDialog.FileName.Split("\\");
                    lbl_fileName_excel.Content = lbl + "...../" + fileName[fileName.Length - 1];
                }
                else
                {
                    lbl_fileName_excel.Content = lbl + openFileDialog.FileName;
                }
                fileExcel = openFileDialog.FileName.ToString();
                kpi_area.IsEnabled = true;
                word_are.IsEnabled = true;
            }
        }


        private async void btn_chamdiemkpi_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel File|*.xlsx";
            openFileDialog.Title = "File theo dõi chấm điểm KPI";
            if (openFileDialog.ShowDialog() == true)
            {
                this.fileChamDiem = openFileDialog.FileName;
                lbl_chamdiem.Content = openFileDialog.FileName.Length > 50 ? openFileDialog.FileName.Substring(0, 30) + "..." : openFileDialog.FileName;



            }
        }
    }
}