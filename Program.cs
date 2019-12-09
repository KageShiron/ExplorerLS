using MicroBatchFramework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExplorerUtils
{
    class Program
    {
        static void Main(string[] args)
        {
            BatchHost.CreateDefaultBuilder().RunBatchEngineAsync<ExplorerBatch>(args);
        }

        class ShellWindow
        {
            public ShellWindow(string uri, long hwnd)
            {
                this.Uri = uri;
                this.Hwnd = new IntPtr(hwnd);
            }
            public string Uri { get; set; }
            public IntPtr Hwnd { get; set; }
        }

        public class ExplorerBatch : BatchBase
        {
            static readonly Guid CLSID_ShellWindows = new Guid("{9BA05972-F6A8-11CF-A442-00A0C90A8F39}");

            public void Default()
            {
                Last();
            }

            [Command("last", "Last actived explorer")]
            public void Last(
                [Option("l", "no output if last actived window have no filesystem path.")] bool lastOnly = false)
            {
                dynamic shWin = Activator.CreateInstance(Type.GetTypeFromCLSID(CLSID_ShellWindows));
                var wins = new List<ShellWindow>();

                foreach (dynamic w in shWin)
                {
                    if ("EXPLORER.EXE".Equals(Path.GetFileName((string)w.FullName).ToUpper(), StringComparison.Ordinal))
                        wins.Add(new ShellWindow(w.LocationURL, w.HWND));
                }

                PInvoke.User32.EnumWindows((hwnd, _) =>
                {
                    foreach (var w in wins)
                    {
                        if (w.Hwnd == hwnd)
                        {
                            if (w.Uri == "")
                            {
                                if (lastOnly)
                                {
                                    Environment.Exit(1);
                                }
                                else { continue; }
                            }
                            Console.WriteLine(new Uri(w.Uri).LocalPath);
                            return false;
                        }
                    }
                    return true;
                }, IntPtr.Zero);
            }


            [Command("all", "Show all explorer path")]
            public void All()
            {
                dynamic wins = Activator.CreateInstance(Type.GetTypeFromCLSID(CLSID_ShellWindows));

                foreach (dynamic w in wins)
                {
                    if (!string.IsNullOrWhiteSpace(w.LocationURL) && "EXPLORER.EXE".Equals(Path.GetFileName((string)w.FullName).ToUpper(), StringComparison.Ordinal))
                        Console.WriteLine(new Uri(w.LocationURL).LocalPath);
                }
            }
        }
    }
}