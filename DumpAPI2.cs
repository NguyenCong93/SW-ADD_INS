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
                Type tBody = asm.GetType("SolidWorks.Interop.sldworks.IBody2");
                using (StreamWriter sw = new StreamWriter("dump.txt"))
                {
                    sw.WriteLine("Body2 Methods:");
                    foreach (var m in tBody.GetMethods())
                    {
                        if (m.Name.Contains("GetBodyBox") || m.Name.Contains("Select")) {
                            sw.Write(m.Name + "(");
                            var pars = m.GetParameters();
                            for(int i=0; i<pars.Length; i++) {
                                sw.Write(pars[i].ParameterType.Name + " " + pars[i].Name + (i<pars.Length-1?", ":""));
                            }
                            sw.WriteLine(")");
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                File.WriteAllText("dump.txt", ex.ToString());
            }
        }
    }
}
