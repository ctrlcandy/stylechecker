using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace stylechecker
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        /// 

        [DllImport("kernel32.dll")]
        static extern bool AttachConsole(int dwProcessId);
        private const int ATTACH_PARENT_PROCESS = -1;

        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form());
            }
            else
            {
                if (args[0] == "-crun")
                {
                    AttachConsole(ATTACH_PARENT_PROCESS);
                    Console.WriteLine("\nLimited functionality mode");

                    AttachConsole(ATTACH_PARENT_PROCESS);
                    Console.WriteLine("Font:");
                    string Font = Console.ReadLine();

                    AttachConsole(ATTACH_PARENT_PROCESS);
                    Console.WriteLine("Font size:");
                    int FontSize = Convert.ToInt32(Console.ReadLine());

                    AttachConsole(ATTACH_PARENT_PROCESS);
                    Console.WriteLine("Line spacing:");
                    double LineSpacing = Convert.ToDouble(Console.ReadLine());

                    AttachConsole(ATTACH_PARENT_PROCESS);
                    Console.WriteLine("Alignment:");
                    string Alignment = Console.ReadLine();

                    Stylechecker s = new Stylechecker(Font, FontSize, LineSpacing, Alignment);
                    s.MyDocument(args[0]);

                    AttachConsole(ATTACH_PARENT_PROCESS);
                    Console.WriteLine(s.ResultErrors);

                }
                else if (args.Length == 1 && args[0] != "-help" && args[0] != "--help" && args[0] != "-man")
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new Form(args[0]));
                }
                else
                {
                    AttachConsole(ATTACH_PARENT_PROCESS);
                    Console.WriteLine("\nUse -crun for console run");
                }
            }
        }
    }
}
