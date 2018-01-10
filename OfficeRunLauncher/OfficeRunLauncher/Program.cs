using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace OfficeRunLauncher
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            App app = new App();
            app.Run(new MainWindow(args.FirstOrDefault()));
        }
    }
}
