using System;
using System.Windows.Forms;

namespace CodeWalker.MloExporter
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MloExporterForm());

            GTAFolder.UpdateSettings();
        }
    }
}
