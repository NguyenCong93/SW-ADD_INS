using System;
using System.Reflection;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace DumpAPI
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Type swType = Type.GetTypeFromProgID("SldWorks.Application");
                ISldWorks swApp = (ISldWorks)Activator.CreateInstance(swType);
                if (swApp == null) return;
                
                ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
                if (swModel == null || swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    Console.WriteLine("Please open an assembly.");
                    return;
                }
                
                AssemblyDoc swAssy = (AssemblyDoc)swModel;
                InterferenceDetectionMgr mgr = swAssy.InterferenceDetectionManager;
                int count = mgr.GetInterferenceCount();
                mgr.Done();
                object[] inters = (object[])mgr.GetInterferences();
                
                if (inters != null && inters.Length > 0)
                {
                    object inter = inters[0];
                    Type t = inter.GetType();
                    Console.WriteLine("Type: " + t.Name);
                    
                    Console.WriteLine("Properties:");
                    foreach (var p in t.GetProperties())
                    {
                        Console.WriteLine("- " + p.Name + " : " + p.PropertyType.Name);
                    }
                    
                    Console.WriteLine("Methods:");
                    foreach (var m in t.GetMethods())
                    {
                        Console.WriteLine("- " + m.Name);
                    }
                }
                else
                {
                    Console.WriteLine("No interferences found.");
                }
                mgr.Done();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
