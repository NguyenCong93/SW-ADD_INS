using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;

namespace SwAutomationAddin
{
    [Guid("A8F8DE2E-6D3A-4F39-A1C8-53B4B2F3D9C1"), ComVisible(true)]
    public class SwAddin : ISwAddin
    {
        #region Local Variables
        private ISldWorks iSwApp = null;
        private int addinID = 0;
        #endregion

        #region SolidWorks Registration
        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 1);

                addinkey.SetValue("Description", "SolidWorks Automation Add-in");
                addinkey.SetValue("Title", "SwAutomationAddin");

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, 1, Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error Registering COM: " + e.Message);
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error Unregistering COM: " + e.Message);
            }
        }
        #endregion

        #region ISwAddin Implementation
        private ITaskpaneView swTaskpaneView;
        private MainTaskPane swTaskpaneHost;

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            iSwApp = (ISldWorks)ThisSW;
            addinID = cookie;

            // Set up callbacks here
            iSwApp.SetAddinCallbackInfo(0, this, addinID);

            // Create Taskpane View and load UI
            LoadTaskpane();
            
            return true;
        }

        public bool DisconnectFromSW()
        {
            // Destroy Taskpane View
            UnloadTaskpane();
            
            System.Runtime.InteropServices.Marshal.ReleaseComObject(iSwApp);
            iSwApp = null;
            return true;
        }

        private void LoadTaskpane()
        {
            string bitmapPath = ""; // Path to icon, skip for now
            string tooltip = "SwAutomation Add-in";

            swTaskpaneView = iSwApp.CreateTaskpaneView2(bitmapPath, tooltip);
            
            swTaskpaneHost = (MainTaskPane)swTaskpaneView.AddControl("SwAutomationAddin.MainTaskPane", "");
            if (swTaskpaneHost != null)
            {
                swTaskpaneHost.Setup(iSwApp);
            }
        }

        private void UnloadTaskpane()
        {
            if (swTaskpaneHost != null)
            {
                swTaskpaneHost.Dispose();
                swTaskpaneHost = null;
            }
            if (swTaskpaneView != null)
            {
                swTaskpaneView.DeleteView();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(swTaskpaneView);
                swTaskpaneView = null;
            }
        }
        #endregion
    }
}
