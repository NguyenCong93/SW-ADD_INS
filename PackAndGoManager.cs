using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomationAddin
{
    public class PackAndGoManager
    {
        private ISldWorks _swApp;

        public PackAndGoManager(ISldWorks swApp)
        {
            _swApp = swApp;
        }

        /// <summary>
        /// Thực hiện Pack and Go: tạo thư mục chứa Assembly + tất cả components
        /// </summary>
        public void ExecutePackAndGo()
        {
            try
            {
                ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
                if (swModel == null || swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    _swApp.SendMsgToUser("Vui lòng mở một Assembly để thực hiện Pack and Go.");
                    return;
                }

                string assyPath = swModel.GetPathName();
                if (string.IsNullOrEmpty(assyPath))
                {
                    _swApp.SendMsgToUser("Vui lòng lưu Assembly trước khi thực hiện Pack and Go.");
                    return;
                }

                string assyDir = Path.GetDirectoryName(assyPath);
                string assyName = Path.GetFileNameWithoutExtension(assyPath);

                // Tạo thư mục Pack_and_Go_YYYYMMDD_HHMMSS
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string packDir = Path.Combine(assyDir, $"Pack_and_Go_{assyName}_{timestamp}");

                if (!Directory.Exists(packDir))
                {
                    Directory.CreateDirectory(packDir);
                }

                // Copy Assembly chính
                string newAssyPath = Path.Combine(packDir, Path.GetFileName(assyPath));
                File.Copy(assyPath, newAssyPath, true);

                // Copy tất cả components
                AssemblyDoc swAssy = (AssemblyDoc)swModel;
                object[] components = (object[])swAssy.GetComponents(false);

                int copiedCount = 0;
                if (components != null)
                {
                    foreach (object compObj in components)
                    {
                        Component2 swComp = (Component2)compObj;
                        if (swComp.IsSuppressed() || swComp.IsHidden(true)) continue;

                        ModelDoc2 compModel = (ModelDoc2)swComp.GetModelDoc2();
                        if (compModel == null) continue;

                        string compPath = compModel.GetPathName();
                        if (!string.IsNullOrEmpty(compPath))
                        {
                            string newCompPath = Path.Combine(packDir, Path.GetFileName(compPath));
                            if (!File.Exists(newCompPath))
                            {
                                File.Copy(compPath, newCompPath, true);
                                copiedCount++;
                            }
                        }
                    }
                }

                _swApp.SendMsgToUser(
                    string.Format("Pack and Go hoàn tất!\n\n" +
                    "Thư mục: {0}\n" +
                    "Assembly: 1\n" +
                    "Components: {1}",
                    packDir, copiedCount));
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi Pack and Go: " + ex.Message);
            }
        }
    }
}