using System;
using System.Collections.Generic;
using System.IO;
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
using OFFICECORE = Microsoft.Office.Core;
using POWERPOINT = Microsoft.Office.Interop.PowerPoint;
using EXCEL = Microsoft.Office.Interop.Excel;
using WORD = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Windows.Threading;

namespace OfficeRunLauncher
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// 开始时间
        /// </summary>
        private static DateTime beginTime;

        private static int Forgur = 20;
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="path"></param>
        public MainWindow(string path)
        {
            this.DataContext = this;
            beginTime = DateTime.Now;

            DispatcherTimer clock = new DispatcherTimer();
            clock.Interval = new TimeSpan(0, 0, 1);
            clock.Tick += new EventHandler(Strat);
            clock.Start();

            InitializeComponent();
            GetFileType(path); 
        }

        /// <summary>
        /// 获得文件类型
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private void GetFileType(string path)
        {
            FileInfo fileinfo = new FileInfo(path);
            if (fileinfo.Exists)
            {
                boxpath.ToolTip = path;
                boxpath.Text = path.Length > 35 ? path.Substring(0, 33) + "......" : path;
                boxsize.Text = Math.Round((decimal)fileinfo.Length / 1024 / 1024, 2).ToString() + " MB";
                boxlasttime.Text = fileinfo.LastWriteTime.ToString();
                switch(fileinfo.Extension)
                {
                    case ".xls":
                    case ".xlsx": 
                        boxtype.Text = "Excel" + "（" + fileinfo.Extension + ")";
                        OpenExcelDoucument(path);
                        break;
                    case ".doc":
                    case ".docx":
                    case ".dot":
                    case ".xml":
                        boxtype.Text = "Word" + "（" + fileinfo.Extension + ")";
                        OpenWordDoucument(path);
                        break;
                    case ".ppt":
                    case ".pptx":
                        boxtype.Text = "PowerPoint" + "（" + fileinfo.Extension + ")";
                        OpenPPTDoucument(path);
                        break;
                }
            }
        }

        /// <summary>
        /// 计时器
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void Strat(object sender, EventArgs e)
        {
            this.Topmost = true;
            Forgur--;
            autoshutdown.Content = "此窗口将在（" + Forgur.ToString() + "）内关闭";
            if (Forgur == 0)
            {
                Application.Current.Shutdown();
            }
        }

        /// <summary>
        /// 求间隔时间
        /// </summary>
        private void MathSpanTime()
        {
            TimeSpan ts = DateTime.Now - beginTime;
            boxtime.Text = ts.Milliseconds.ToString() + " 毫秒";
            boxstatus.Text = "链 接 已 成 功";
        }

        /// <summary>
        /// 打开Word
        /// </summary>
        /// <param name="path"></param>
        private void OpenWordDoucument(string path)
        {
            try
            {
                Microsoft.Office.Interop.Word._Application oWord;
                Microsoft.Office.Interop.Word._Document oDoc;
                oWord = new Microsoft.Office.Interop.Word.Application();
                oWord.Visible = true;
                oDoc = oWord.Documents.Open(path);

                MathSpanTime();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 打开Excel
        /// </summary>
        /// <param name="path"></param>
        private void OpenExcelDoucument(string path)
        {
            try
            {
                Microsoft.Office.Interop.Excel._Application oExcel;
                Microsoft.Office.Interop.Excel._Workbook oXls;
                oExcel = new Microsoft.Office.Interop.Excel.Application();
                oExcel.Visible = true;
                oXls = oExcel.Workbooks.Open(path);

                MathSpanTime();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 打开PowerPoint
        /// </summary>
        /// <param name="path"></param>
        private void OpenPPTDoucument(string path)
        {
            try
            {
                POWERPOINT.Application objApp = null;
                POWERPOINT.Presentation objPresSet = null;
                POWERPOINT.SlideShowSettings objSSS;
                bool bAssistantOn;
                objApp = new POWERPOINT.Application();
                objPresSet = objApp.Presentations.Open(path);

                MathSpanTime();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 异型窗口移动
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void f_hand_MouseMove(object sender, MouseEventArgs e)
        {
            Point pp = Mouse.GetPosition(this);
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        /// <summary>
        /// 最小化
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void e_mini_MouseDown(object sender, MouseButtonEventArgs e)
        {
            string msg = "OfficeRunLauncher\r\n作者：LingMin\r\n帮朋友的朋友写的帮助软件\r\n解决Office每次通过文件打开的时候出现\r\n正在安装 . . . 的问题\r\n需要探讨技术细节可以邮件来访：\r\nKid--L@Hotmail.com";
            MessageBox.Show(msg);
        }

        /// <summary>
        /// 关闭
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void e_close_MouseDown(object sender, MouseButtonEventArgs e)
        {
            Application.Current.Shutdown();
        }
    }
}
