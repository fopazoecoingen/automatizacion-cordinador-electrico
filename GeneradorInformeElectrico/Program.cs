using System;
using System.Windows.Forms;

namespace GeneradorInformeElectrico;

static class Program
{
    [STAThread]
    static void Main()
    {
        LogHelper.IniciarSesion();
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}
