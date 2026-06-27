using System;
using System.IO;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomationAddin
{
    /// <summary>
    /// Enum để xác định loại file đang mở
    /// </summary>
    public enum FileType
    {
        Part,
        Assembly,
        Drawing,
        Unknown
    }

    /// <summary>
    /// Enum loại đặt tên file
    /// </summary>
    public enum FileNamingRule
    {
        DrawingNo,      // Mã BV (Chi tiết lẻ)
        UnitName,       // Tên cụm con
        MachineName     // Tên cụm tổng
    }

    /// <summary>
    /// Enum cho tuỳ chọn xuất sheet
    /// </summary>
    public enum SheetExportOption
    {
        ActiveSheet,        // Sheet đang mở
        AllSheetsInOneFile, // Gộp tất cả vào 1 file
        SeparateFiles,      // Tách riêng từng file
        RangeSheets         // Xuất theo dải
    }

    public class ExportManager
    {
        private ISldWorks _swApp;

        public ExportManager(ISldWorks swApp)
        {
            _swApp = swApp;
        }

        #region Helper Methods

        /// <summary>
        /// Lấy loại file đang mở
        /// </summary>
        public FileType GetCurrentFileType()
        {
            try
            {
                ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
                if (swModel == null) return FileType.Unknown;

                int docType = swModel.GetType();

                if (docType == (int)swDocumentTypes_e.swDocPART)
                    return FileType.Part;
                else if (docType == (int)swDocumentTypes_e.swDocASSEMBLY)
                    return FileType.Assembly;
                else if (docType == (int)swDocumentTypes_e.swDocDRAWING)
                    return FileType.Drawing;

                return FileType.Unknown;
            }
            catch
            {
                return FileType.Unknown;
            }
        }

        /// <summary>
        /// Tạo thư mục Exported_Files
        /// </summary>
        private string EnsureExportDirectory()
        {
            ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
            if (swModel == null) return null;

            string originalPath = swModel.GetPathName();
            if (string.IsNullOrEmpty(originalPath)) return null;

            string dir = Path.GetDirectoryName(originalPath);
            string exportDir = Path.Combine(dir, "Exported_Files");

            if (!Directory.Exists(exportDir))
            {
                Directory.CreateDirectory(exportDir);
            }

            return exportDir;
        }

        /// <summary>
        /// Lấy tên file dựa theo rule
        /// </summary>
        private string GetFileName(FileNamingRule rule)
        {
            ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
            if (swModel == null) return "Export";

            CustomPropertyManager pm = swModel.Extension.get_CustomPropertyManager("");
            string propName = "";

            switch (rule)
            {
                case FileNamingRule.DrawingNo:
                    propName = "Title"; // Mã BV
                    break;
                case FileNamingRule.UnitName:
                    propName = "Unit"; // Cụm con
                    break;
                case FileNamingRule.MachineName:
                    propName = "Name Project"; // Cụm tổng
                    break;
                default:
                    return Path.GetFileNameWithoutExtension(swModel.GetPathName());
            }

            string valOut = "";
            string resValOut = "";
            bool wasResolved = false;
            bool linkToProp = false;

            pm.Get6(propName, false, out valOut, out resValOut, out wasResolved, out linkToProp);

            return !string.IsNullOrEmpty(resValOut) ? resValOut : Path.GetFileNameWithoutExtension(swModel.GetPathName());
        }

        /// <summary>
        /// Lấy tên file tuỳ theo loại file
        /// </summary>
        public string GetExportFileName(FileNamingRule rule)
        {
            return GetFileName(rule);
        }

        #endregion

        #region PDF Export

        /// <summary>
        /// Export PDF 2D từ Drawing
        /// </summary>
        public void ExportDrawingToPDF2D(
            FileNamingRule namingRule,
            SheetExportOption sheetOption,
            string sheetRange = "")
        {
            try
            {
                ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
                if (swModel == null)
                {
                    _swApp.SendMsgToUser("Không có tài liệu nào đang mở.");
                    return;
                }

                if (swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                {
                    _swApp.SendMsgToUser("Chỉ dùng được cho Drawing.");
                    return;
                }

                string exportDir = EnsureExportDirectory();
                if (string.IsNullOrEmpty(exportDir)) return;

                DrawingDoc swDrawing = (DrawingDoc)swModel;
                string baseFileName = GetFileName(namingRule);
                int sheetCount = swDrawing.GetSheetCount();

                List<string> exportedFiles = new List<string>();

                // Xử lý theo lựa chọn sheet
                switch (sheetOption)
                {
                    case SheetExportOption.ActiveSheet:
                        ExportSingleSheetPDF(swDrawing, baseFileName, exportDir, exportedFiles);
                        break;

                    case SheetExportOption.AllSheetsInOneFile:
                        ExportAllSheetsPDF(swDrawing, baseFileName, exportDir, exportedFiles);
                        break;

                    case SheetExportOption.SeparateFiles:
                        ExportSheetsToSeparatePDF(swDrawing, baseFileName, exportDir, exportedFiles);
                        break;

                    case SheetExportOption.RangeSheets:
                        ExportSheetsInRangePDF(swDrawing, baseFileName, sheetRange, exportDir, exportedFiles);
                        break;
                }

                if (exportedFiles.Count > 0)
                {
                    _swApp.SendMsgToUser($"✓ Xuất PDF 2D thành công!\n" +
                        $"Số file: {exportedFiles.Count}\n" +
                        $"Thư mục: {exportDir}");
                }
                else
                {
                    _swApp.SendMsgToUser("Không có file nào được xuất.");
                }
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi: " + ex.Message);
            }
        }

        /// <summary>
        /// Export PDF 3D từ Part/Assembly
        /// </summary>
        public void ExportToPDF3D(FileNamingRule namingRule)
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
                if (docType == (int)swDocumentTypes_e.swDocDRAWING)
                {
                    _swApp.SendMsgToUser("PDF 3D không dùng được cho Drawing.");
                    return;
                }

                string exportDir = EnsureExportDirectory();
                if (string.IsNullOrEmpty(exportDir)) return;

                string fileName = GetFileName(namingRule);
                string pdfPath = Path.Combine(exportDir, fileName + "_3D.pdf");

                int errors = 0;
                int warnings = 0;

                bool status = swModel.Extension.SaveAs(
                    pdfPath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    null,
                    ref errors,
                    ref warnings);

                if (status)
                {
                    _swApp.SendMsgToUser($"✓ Xuất PDF 3D thành công!\n{pdfPath}");
                }
                else
                {
                    _swApp.SendMsgToUser($"✗ Lỗi xuất PDF 3D. Mã lỗi: {errors}");
                }
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi: " + ex.Message);
            }
        }

        private void ExportSingleSheetPDF(DrawingDoc swDrawing, string baseFileName,
            string exportDir, List<string> exportedFiles)
        {
            try
            {
                string pdfPath = Path.Combine(exportDir, baseFileName + ".pdf");
                int errors = 0;
                int warnings = 0;

                bool status = ((ModelDoc2)swDrawing).Extension.SaveAs(
                    pdfPath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    null,
                    ref errors,
                    ref warnings);

                if (status)
                {
                    exportedFiles.Add(pdfPath);
                }
            }
            catch { }
        }

        private void ExportAllSheetsPDF(DrawingDoc swDrawing, string baseFileName,
            string exportDir, List<string> exportedFiles)
        {
            try
            {
                string pdfPath = Path.Combine(exportDir, baseFileName + "_AllSheets.pdf");
                int errors = 0;
                int warnings = 0;

                bool status = ((ModelDoc2)swDrawing).Extension.SaveAs(
                    pdfPath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    null,
                    ref errors,
                    ref warnings);

                if (status)
                {
                    exportedFiles.Add(pdfPath);
                }
            }
            catch { }
        }

        private void ExportSheetsToSeparatePDF(DrawingDoc swDrawing, string baseFileName,
            string exportDir, List<string> exportedFiles)
        {
            try
            {
                int sheetCount = swDrawing.GetSheetCount();

                for (int i = 1; i <= sheetCount; i++)
                {
                    Sheet2 sheet = swDrawing.GetSheet(i - 1) as Sheet2;
                    if (sheet == null) continue;

                    string sheetName = sheet.GetName();
                    string pdfPath = Path.Combine(exportDir, $"{baseFileName}_{sheetName}.pdf");

                    // Activate sheet
                    swDrawing.ActivateSheet(sheetName);

                    int errors = 0;
                    int warnings = 0;

                    bool status = ((ModelDoc2)swDrawing).Extension.SaveAs(
                        pdfPath,
                        (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                        (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                        null,
                        ref errors,
                        ref warnings);

                    if (status)
                    {
                        exportedFiles.Add(pdfPath);
                    }
                }
            }
            catch { }
        }

        private void ExportSheetsInRangePDF(DrawingDoc swDrawing, string baseFileName,
            string sheetRange, string exportDir, List<string> exportedFiles)
        {
            try
            {
                List<int> sheetIndices = ParseSheetRange(sheetRange, swDrawing.GetSheetCount());

                foreach (int idx in sheetIndices)
                {
                    Sheet2 sheet = swDrawing.GetSheet(idx - 1) as Sheet2;
                    if (sheet == null) continue;

                    string sheetName = sheet.GetName();
                    string pdfPath = Path.Combine(exportDir, $"{baseFileName}_{sheetName}.pdf");

                    swDrawing.ActivateSheet(sheetName);

                    int errors = 0;
                    int warnings = 0;

                    bool status = ((ModelDoc2)swDrawing).Extension.SaveAs(
                        pdfPath,
                        (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                        (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                        null,
                        ref errors,
                        ref warnings);

                    if (status)
                    {
                        exportedFiles.Add(pdfPath);
                    }
                }
            }
            catch { }
        }

        #endregion

        #region DXF/DWG Export

        /// <summary>
        /// Export DXF/DWG từ Drawing
        /// </summary>
        public void ExportDrawingToDxfDwg(
            bool isDxf,
            string cadVersion,
            FileNamingRule namingRule,
            SheetExportOption sheetOption,
            string sheetRange = "")
        {
            try
            {
                ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
                if (swModel == null)
                {
                    _swApp.SendMsgToUser("Không có tài liệu nào đang mở.");
                    return;
                }

                if (swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                {
                    _swApp.SendMsgToUser("Chỉ dùng được cho Drawing.");
                    return;
                }

                string exportDir = EnsureExportDirectory();
                if (string.IsNullOrEmpty(exportDir)) return;

                DrawingDoc swDrawing = (DrawingDoc)swModel;
                string baseFileName = GetFileName(namingRule);
                string ext = isDxf ? ".dxf" : ".dwg";

                List<string> exportedFiles = new List<string>();

                switch (sheetOption)
                {
                    case SheetExportOption.ActiveSheet:
                        ExportSingleDxfDwg(swDrawing, baseFileName, ext, exportDir, exportedFiles);
                        break;

                    case SheetExportOption.SeparateFiles:
                        ExportDxfDwgSeparate(swDrawing, baseFileName, ext, exportDir, exportedFiles);
                        break;

                    case SheetExportOption.RangeSheets:
                        ExportDxfDwgRange(swDrawing, baseFileName, sheetRange, ext, exportDir, exportedFiles);
                        break;
                }

                if (exportedFiles.Count > 0)
                {
                    string formatName = isDxf ? "DXF" : "DWG";
                    _swApp.SendMsgToUser($"✓ Xuất {formatName} thành công!\n" +
                        $"Số file: {exportedFiles.Count}\n" +
                        $"Thư mục: {exportDir}");
                }
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi: " + ex.Message);
            }
        }

        private void ExportSingleDxfDwg(DrawingDoc swDrawing, string baseFileName,
            string ext, string exportDir, List<string> exportedFiles)
        {
            try
            {
                string filePath = Path.Combine(exportDir, baseFileName + ext);

                int errors = 0;
                int warnings = 0;

                bool status = ((ModelDoc2)swDrawing).Extension.SaveAs(
                    filePath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    null,
                    ref errors,
                    ref warnings);

                if (status)
                {
                    exportedFiles.Add(filePath);
                }
            }
            catch { }
        }

        private void ExportDxfDwgSeparate(DrawingDoc swDrawing, string baseFileName,
            string ext, string exportDir, List<string> exportedFiles)
        {
            try
            {
                int sheetCount = swDrawing.GetSheetCount();

                for (int i = 1; i <= sheetCount; i++)
                {
                    Sheet2 sheet = swDrawing.GetSheet(i - 1) as Sheet2;
                    if (sheet == null) continue;

                    string sheetName = sheet.GetName();
                    string filePath = Path.Combine(exportDir, $"{baseFileName}_{sheetName}{ext}");

                    swDrawing.ActivateSheet(sheetName);

                    int errors = 0;
                    int warnings = 0;

                    bool status = ((ModelDoc2)swDrawing).Extension.SaveAs(
                        filePath,
                        (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                        (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                        null,
                        ref errors,
                        ref warnings);

                    if (status)
                    {
                        exportedFiles.Add(filePath);
                    }
                }
            }
            catch { }
        }

        private void ExportDxfDwgRange(DrawingDoc swDrawing, string baseFileName,
            string sheetRange, string ext, string exportDir, List<string> exportedFiles)
        {
            try
            {
                List<int> sheetIndices = ParseSheetRange(sheetRange, swDrawing.GetSheetCount());

                foreach (int idx in sheetIndices)
                {
                    Sheet2 sheet = swDrawing.GetSheet(idx - 1) as Sheet2;
                    if (sheet == null) continue;

                    string sheetName = sheet.GetName();
                    string filePath = Path.Combine(exportDir, $"{baseFileName}_{sheetName}{ext}");

                    swDrawing.ActivateSheet(sheetName);

                    int errors = 0;
                    int warnings = 0;

                    bool status = ((ModelDoc2)swDrawing).Extension.SaveAs(
                        filePath,
                        (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                        (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                        null,
                        ref errors,
                        ref warnings);

                    if (status)
                    {
                        exportedFiles.Add(filePath);
                    }
                }
            }
            catch { }
        }

        #endregion

        #region STEP/X_T Export

        /// <summary>
        /// Export STEP hoặc X_T từ Part/Assembly
        /// </summary>
        public void ExportToStepOrXt(
            bool isStep,
            FileNamingRule namingRule)
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
                if (docType == (int)swDocumentTypes_e.swDocDRAWING)
                {
                    _swApp.SendMsgToUser("STEP/X_T không dùng được cho Drawing.");
                    return;
                }

                string exportDir = EnsureExportDirectory();
                if (string.IsNullOrEmpty(exportDir)) return;

                string fileName = GetFileName(namingRule);
                string ext = isStep ? ".stp" : ".xt";
                string filePath = Path.Combine(exportDir, fileName + ext);

                int errors = 0;
                int warnings = 0;

                bool status = swModel.Extension.SaveAs(
                    filePath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Copy | (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    null,
                    ref errors,
                    ref warnings);

                if (status)
                {
                    string formatName = isStep ? "STEP" : "X_T";
                    _swApp.SendMsgToUser($"✓ Xuất {formatName} thành công!\n{filePath}");
                }
                else
                {
                    _swApp.SendMsgToUser($"✗ Lỗi xuất file. Mã lỗi: {errors}");
                }
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi: " + ex.Message);
            }
        }

        #endregion

        #region Utility Methods

        /// <summary>
        /// Parse sheet range: "1,3-5,7" -> [1, 3, 4, 5, 7]
        /// </summary>
        private List<int> ParseSheetRange(string rangeStr, int maxSheets)
        {
            List<int> result = new List<int>();

            if (string.IsNullOrWhiteSpace(rangeStr))
                return result;

            string[] parts = rangeStr.Split(',');

            foreach (string part in parts)
            {
                string trimmed = part.Trim();

                if (trimmed.Contains("-"))
                {
                    string[] rangeParts = trimmed.Split('-');
                    if (int.TryParse(rangeParts[0], out int start) && int.TryParse(rangeParts[1], out int end))
                    {
                        for (int i = start; i <= end && i <= maxSheets; i++)
                        {
                            if (!result.Contains(i))
                                result.Add(i);
                        }
                    }
                }
                else if (int.TryParse(trimmed, out int single))
                {
                    if (single <= maxSheets && !result.Contains(single))
                        result.Add(single);
                }
            }

            return result;
        }

        #endregion
    }
}