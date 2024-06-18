using System;
using ImportSheetConsole.Properties;

namespace ImportSheetConsole
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            //Settings.Default.Host = "127.0.0.1";
            //Settings.Default.Port = 8000;
            //Settings.Default.Login = "Administrator";
            //Settings.Default.Password = "1234";

            var script = new DefaultScheme.ScriptImplemetation();
            script.Execute(@"d:\1\Опросный лист по умолчанию.xlsx");
            
            Console.WriteLine("Для завершения, нажмите любую клавишу...");
            Console.ReadLine();
        }
    }
}