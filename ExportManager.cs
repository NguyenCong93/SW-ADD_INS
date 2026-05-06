using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomationAddin
{
    public class ExportManager
    {
        private ISldWorks _swApp;

        public ExportManager(ISldWorks swApp)
        {
            _swApp = swApp;
        }

        /// <summary>
        /// Xuất file ra định dạng PDF hoặc 3D PDF (cho Part/Assy)
        /// </summary>
        public void ExportToPDF()
        {
            try
            {
                ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
                if (swModel == null)
                {
                    _swApp.SendMsgToUser("Không có tài liệu nào đang mở.");
                    return;
                }

                string originalPath = swModel.GetPathName();
                if (string.IsNullOrEmpty(originalPath))
                {
                    _swApp.SendMsgToUser("Vui lòng lưu file trước khi xuất PDF.");
                    return;
                }

                string dir = Path.GetDirectoryName(originalPath);
                string exportDir = Path.Combine(dir, "Exported_Files");
                if (!Directory.Exists(exportDir))
                {
                    Directory.CreateDirectory(exportDir);
                }

                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(originalPath);
                string newPath = Path.Combine(exportDir, fileNameWithoutExt + ".pdf");

                int errors = 0;
                int warnings = 0;

                ModelDocExtension swModExt = swModel.Extension;
                
                // Thiết lập Export Data cho PDF
                ExportPdfData swExportData = (ExportPdfData)_swApp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);

                if (swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                {
                    // Note: For 3D PDF, you usually need to configure the ExportPdfData
                }

                bool status = swModExt.SaveAs(newPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, swExportData, ref errors, ref warnings);

                if (status)
                {
                    _swApp.SendMsgToUser("Đã xuất PDF thành công: " + newPath);
                }
                else
                {
                    _swApp.SendMsgToUser(string.Format("Lỗi xuất PDF. Mã lỗi: {0}", errors));
                }
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi trong quá trình xuất PDF: " + ex.Message);
            }
        }

        /// <summary>
        /// Xuất file ra định dạng STP / X_T với tên được chỉ định, lưu cùng thư mục với file gốc
        /// </summary>
        public void ExportToStepOrParasolid(string targetName, string extension = ".stp")
        {
            try
            {
                ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
                if (swModel == null)
                {
                    _swApp.SendMsgToUser("Không có tài liệu nào đang mở.");
                    return;
                }

                int docType = swModel.GetType();
                if (docType != (int)swDocumentTypes_e.swDocPART && docType != (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    _swApp.SendMsgToUser("Chỉ có thể xuất STEP/Parasolid cho Part và Assembly.");
                    return;
                }

                string originalPath = swModel.GetPathName();
                if (string.IsNullOrEmpty(originalPath))
                {
                    _swApp.SendMsgToUser("Vui lòng lưu file gốc trước khi xuất.");
                    return;
                }

                string dir = Path.GetDirectoryName(originalPath);
                string exportDir = Path.Combine(dir, "Exported_Files");
                if (!Directory.Exists(exportDir))
                {
                    Directory.CreateDirectory(exportDir);
                }

                string newPath = Path.Combine(exportDir, targetName + extension);

                int errors = 0;
                int warnings = 0;

                // SaveAs với tuỳ chọn swSaveAsOptions_Copy: chỉ xuất file bản sao ra, file gốc đang mở vẫn không bị thay đổi
                bool status = swModel.Extension.SaveAs(
                    newPath, 
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion, 
                    (int)swSaveAsOptions_e.swSaveAsOptions_Copy | (int)swSaveAsOptions_e.swSaveAsOptions_Silent, 
                    null, 
                    ref errors, 
                    ref warnings);

                if (status)
                {
                    _swApp.SendMsgToUser("Đã xuất thành công: " + newPath);
                }
                else
                {
                    _swApp.SendMsgToUser(string.Format("Lỗi xuất file. Mã lỗi: {0}", errors));
                }
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi trong quá trình xuất file: " + ex.Message);
            }
        }
    }
}
