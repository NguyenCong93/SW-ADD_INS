using System;
using System.IO;
using System.Reflection;

namespace DumpAPI
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Assembly asm = Assembly.LoadFile(@"C:\Users\VISC_PC01\.gemini\antigravity\scratch\SwAutomationAddin\SolidWorks.Interop.sldworks.dll");
                Type tInterMgr = asm.GetType("SolidWorks.Interop.sldworks.IInterferenceDetectionMgr");
                using (StreamWriter sw = new StreamWriter("dump_mgr.txt"))
                {
                    sw.WriteLine("IInterferenceDetectionMgr Properties:");
                    foreach (var p in tInterMgr.GetProperties())
                    {
                        sw.WriteLine(p.Name + " : " + p.PropertyType.Name);
                    }
                }
            }
            catch(Exception ex)
            {
                File.WriteAllText("dump_mgr.txt", ex.ToString());
            }
        }
    }
}
