using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace SWTableConvertor
{
    internal static class Program
    {
        [STAThread]
        public static void Main()
        {
            // Set up fallback path resolution for SOLIDWORKS Interop DLLs
            AppDomain.CurrentDomain.AssemblyResolve += ResolveSolidWorksInterop;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }

        private static Assembly ResolveSolidWorksInterop(object sender, ResolveEventArgs args)
        {
            string simpleName = new AssemblyName(args.Name).Name + ".dll";
            
            // 1. Look in the local application folder
            string local = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, simpleName);
            if (File.Exists(local))
                return Assembly.LoadFrom(local);

            // 2. Look in the standard SOLIDWORKS installation folder
            string swPath = Path.Combine(@"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS", simpleName);
            if (File.Exists(swPath))
                return Assembly.LoadFrom(swPath);

            return null;
        }
    }
}
