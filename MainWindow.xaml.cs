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


using Microsoft.Expression.Encoder;
using Microsoft.Expression.Encoder.Devices;
using Microsoft.Expression.Encoder.ScreenCapture;
using System.Windows.Forms;

using System.Drawing;
using Microsoft.Office.Interop.Excel;
namespace videoRecorder
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow :System.Windows. Window
    {
        private ScreenCaptureJob job;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)// record video
        {
            startRecording();
        }
        private void startRecording()
        {
            job = new ScreenCaptureJob();
            System.Drawing.Size workingArea = SystemInformation.WorkingArea.Size;
            System.Drawing.Rectangle captureRect = new System.Drawing.Rectangle(0, 0, workingArea.Width-(workingArea.Width%4), workingArea.Height-(workingArea.Height%4));
            job.CaptureRectangle = captureRect;

            job.ShowFlashingBoundary = true;
            job.ShowCountdown = true;
            job.CaptureMouseCursor = true;
            job.OutputPath = @"D:/RecoreFolder";
            job.Start();
            
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (job.Status == RecordStatus.Running)
                job.Stop();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            runExcel();
        }

        private void runExcel()
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            //app.Visible = true;
            //app.WindowState = XlWindowState.xlMaximized;
            Workbook wb = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = wb.Worksheets[1];
            DateTime currentDate = DateTime.Now;



            // ws.Range["H2"].Value= "this is a header";


            // Range header_range = ws.get_Range("H2:L1");


            //  header_range.Merge(4);

            //  header_range.Interior.Color = System.Drawing.Color.Yellow.ToArgb();
            //  header_range.Font.Size = 20;


            //ws.Range["R2"].Value = "از تاریخ :";
            //ws.Range["R3"].Value = "تا تاریخ :";

            Range range = ws.get_Range("Q2:M2");
            ws.Hyperlinks.Add(range, @"D:\RecoreFolder\rett.PNG", Type.Missing, "Microsoft", "click me");
            //Range taTarikhRange = ws.get_Range("Q3:M3");

            //azTarikhRange.Merge(4);
            //taTarikhRange.Merge(4);

            //// azTarikhRange.Font.Size = 13;
            //// taTarikhRange.Font.Size = 13;
            //// azTarikhRange.Interior.Color = System.Drawing.Color.Red.ToArgb();
            //// taTarikhRange.Interior.Color = System.Drawing.Color.Red.ToArgb();

            //Range azazTarikh = ws.get_Range("R2");
            //Range taTarikh = ws.get_Range("R3");


            //azazTarikh.Font.Size = 13;
            //taTarikh.Font.Size = 13;

            //Range generateDate = ws.get_Range("C2:E2");
            //ws.Range["B2"].Value = "تاریخ گزارش گیری";

            //generateDate.Merge(3);

            //ws.Range["C2"].Value = manageUI.getInstance().daystr;

            ////azazTarikh.Interior.Color = System.Drawing.Color.Yellow.ToArgb();
            //// taTarikh.Interior.Color = System.Drawing.Color.Yellow.ToArgb();


            //ws.Range["M2"].Value = fromDateStr;
            //ws.Range["M3"].Value = toDateStr;
            //ws.Range["Q5"].Value = "نوع هدف";
            //ws.Range["P5"].Value = " نام هدف";
            //ws.Range["O5"].Value = "فرکانس";
            //ws.Range["N5"].Value = "پهنای باند";
            //ws.Range["M5"].Value = "SNR";
            //ws.Range["L5"].Value = "وضعیت";
            //ws.Range["K5"].Value = "مود2";

            //ws.Range["j5"].Value = "زمان اولین مشاهده";
            //ws.Range["I5"].Value = "زمان آخرین مشاهده";

            //ws.Range["H5"].Value = "زاویه شروع";
            //ws.Range["G5"].Value = "زاویه پایان";
            //ws.Range["F5"].Value = "ارتفاع";
            //ws.Range["E5"].Value = "فاصله (km)";
            //ws.Range["D5"].Value = "سرعت";



            //ws.Range["C5"].Value = "آدرس عکس";



            //int counter = Model.tablesData.EnemyList.Count();

            //int i = 6;
            ////   foreach (Model.EnemyDb en in Model.tablesData.EnemyList)
            ////   {
            ////   ws.Range["Q"+i].Value = "نوع هدف";
            ////   ws.Range["P"+i].Value = en.Name;//" نام هدف";
            ////   ws.Range["O"+i].Value = en.Frequency;// "فرکانس";
            ////   ws.Range["N"+i].Value =en.BW;// "پهنای باند";
            ////   ws.Range["M"+i].Value =en.MaxPow;// "SNR";
            ////   ws.Range["L"+i].Value =en.State;// "وضعیت";
            ////   i++;
            ////   }
            //radDataPager1.Dispatcher.Invoke((System.Action)delegate()
            //{
            //    for (int k = 0; k < radDataPager1.PageCount; k++)
            //    {

            //        Thread.Sleep(20);
            //        for (int item = 0; item < radGridView.Items.Count; item++)
            //        {
            //            Thread.Sleep(20);
            //            Model.EnemyDb en = (Model.EnemyDb)radGridView.Items[item];
            //            ws.Range["Q" + i].Value = en.Type;
            //            ws.Range["P" + i].Value = en.Name;//" نام هدف";
            //            ws.Range["O" + i].Value = en.Frequency;// "فرکانس";
            //            ws.Range["N" + i].Value = en.BW;// "پهنای باند";
            //            ws.Range["M" + i].Value = en.MaxPow;// "SNR";
            //            ws.Range["L" + i].Value = en.IsMoving == true ? "متحرک" : "ثابت";
            //            ws.Range["K" + i].Value = en.mode2;

            //            ws.Range["j" + i].Value = en.firstObservedPer;
            //            ws.Range["I" + i].Value = en.lastObservedPer;
            //            double startAzim = DbModel.getInstance().azimStartAz(en.sqlID);
            //            ws.Range["H" + i].Value = startAzim > 0 ? startAzim.ToString() : "---";
            //            ws.Range["G" + i].Value = en.Azimuth;
            //            ws.Range["F" + i].Value = en.Elevation;
            //            ws.Range["E" + i].Value = en.Value;
            //            ws.Range["D" + i].Value = en.Velocity;
            //            string imageAddr = DbModel.getInstance().imageAddr(en.sqlID);
            //            ws.Range["C" + i].Value = imageAddr != "ندارد" ? MainWindow.signalShot + "\\" + imageAddr : imageAddr;



            //            i++;
            //        }
            //        radDataPager1.PagedSource.MoveToNextPage();
            //    }
            //    radDataPager1.PagedSource.MoveToFirstPage();

            //});


            app.Visible = true;
            app.WindowState = XlWindowState.xlMaximized;

            try
            {
               // ws.SaveAs(MainWindow.configFolderPath + "\\report" + "/tt.xlsx");
            }
            catch
            {

            }
        }
    }
}
