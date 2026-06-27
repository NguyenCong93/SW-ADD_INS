using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomationAddin
{
    public class BomItem
    {
        public string PartName { get; set; }
        public string PartType { get; set; } // MF (Gia công) or PUR (Tiêu chuẩn)
        public string Material { get; set; }
        public string SurfaceTreatment { get; set; }
        public string Manufacturer { get; set; }
        public int Quantity { get; set; }
        public int QuantityPerSet { get; set; } // Số lượng/1 set

        public string UniqueKey
        {
            get
            {
                return string.Format("{0}|{1}|{2}|{3}|{4}",
                    PartName, PartType, Material, SurfaceTreatment, Manufacturer);
            }
        }
    }

    public class BomManager
    {
        private ISldWorks _swApp;

        public BomManager(ISldWorks swApp)
        {
            _swApp = swApp;
        }

        /// <summary>
        /// Xuất BOM ra Excel: tách Tiêu chuẩn (PUR) và Gia công (MF) vào 2 sheet
        /// </summary>
        public void ExportBomToExcel()
        {
            try
            {
                ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
                if (swModel == null || swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    _swApp.SendMsgToUser("Vui lòng mở Assembly để xuất BOM.");
                    return;
                }

                AssemblyDoc swAssy = (AssemblyDoc)swModel;
                object[] components = (object[])swAssy.GetComponents(false);

                if (components == null || components.Length == 0)
                {
                    _swApp.SendMsgToUser("Assembly không có chi tiết nào.");
                    return;
                }

                // Lấy Mã thiết bị (Drawing NO) từ Assembly
                string deviceCode = GetPropertyValue(swModel, "Title");

                // Thu thập dữ liệu BOM
                Dictionary<string, BomItem> bomDict = new Dictionary<string, BomItem>();

                foreach (object compObj in components)
                {
                    Component2 swComp = (Component2)compObj;

                    // Bỏ qua chi tiết bị ẩn hoặc suppressed
                    if (swComp.IsSuppressed() || swComp.IsHidden(true))
                        continue;

                    ModelDoc2 compModel = (ModelDoc2)swComp.GetModelDoc2();
                    if (compModel == null)
                        continue;

                    // Lấy thông tin chi tiết
                    string partName = GetPropertyValue(compModel, "Part Name");
                    if (string.IsNullOrEmpty(partName))
                        partName = swComp.Name2;

                    string partType = GetPropertyValue(compModel, "MakeOrBuy"); // MF or PUR
                    string material = GetPropertyValue(compModel, "Material");
                    string surface = GetPropertyValue(compModel, "Surface Treatment");
                    string manufacturer = GetPropertyValue(compModel, "Manufacturer");

                    BomItem item = new BomItem()
                    {
                        PartName = partName,
                        PartType = partType,
                        Material = material,
                        SurfaceTreatment = surface,
                        Manufacturer = manufacturer,
                        Quantity = 1,
                        QuantityPerSet = 1
                    };

                    if (bomDict.ContainsKey(item.UniqueKey))
                    {
                        bomDict[item.UniqueKey].Quantity++;
                    }
                    else
                    {
                        bomDict.Add(item.UniqueKey, item);
                    }
                }

                // Tách thành 2 nhóm
                Dictionary<string, BomItem> bomMF = new Dictionary<string, BomItem>(); // Gia công
                Dictionary<string, BomItem> bomPUR = new Dictionary<string, BomItem>(); // Tiêu chuẩn

                foreach (var kvp in bomDict)
                {
                    if (kvp.Value.PartType == "MF")
                        bomMF.Add(kvp.Key, kvp.Value);
                    else
                        bomPUR.Add(kvp.Key, kvp.Value);
                }

                // Ghi Excel
                WriteToExcelWithTemplate(swModel.GetPathName(), deviceCode, bomMF, bomPUR);
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi BOM: " + ex.Message);
            }
        }

        /// <summary>
        /// Ghi dữ liệu vào Excel (tạo file nếu không có template)
        /// </summary>
        private void WriteToExcelWithTemplate(string assyPath, string deviceCode,
            Dictionary<string, BomItem> bomMF, Dictionary<string, BomItem> bomPUR)
        {
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    _swApp.SendMsgToUser("Không tìm thấy Microsoft Excel.");
                    return;
                }

                dynamic excelApp = Activator.CreateInstance(excelType);
                excelApp.Visible = true;

                // Tạo hoặc mở template
                string dir = Path.GetDirectoryName(assyPath);
                string templatePath = Path.Combine(dir, "BOM_Template.xlsx");
                string outputPath = Path.Combine(dir, Path.GetFileNameWithoutExtension(assyPath) + "_BOM.xlsx");

                dynamic workbook;

                if (File.Exists(templatePath))
                {
                    workbook = excelApp.Workbooks.Open(templatePath);
                }
                else
                {
                    // Tạo workbook mới nếu không có template
                    workbook = excelApp.Workbooks.Add();
                }

                // Xóa sheets cũ nếu có
                int sheetCount = workbook.Sheets.Count;
                for (int i = sheetCount; i >= 1; i--)
                {
                    try { workbook.Sheets[i].Delete(); } catch { }
                }

                // Tạo sheet Gia công (MF)
                dynamic worksheetMF = workbook.Sheets.Add();
                worksheetMF.Name = "Gia công (MF)";
                WriteBomToSheet(worksheetMF, deviceCode, bomMF, true);

                // Tạo sheet Tiêu chuẩn (PUR)
                dynamic worksheetPUR = workbook.Sheets.Add();
                worksheetPUR.Name = "Tiêu chuẩn (PUR)";
                WriteBomToSheet(worksheetPUR, deviceCode, bomPUR, false);

                // Lưu file
                if (File.Exists(outputPath))
                    File.Delete(outputPath);

                workbook.SaveAs(outputPath);
                excelApp.Quit();

                _swApp.SendMsgToUser($"✓ Xuất BOM thành công!\n{outputPath}");
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi Excel: " + ex.Message);
            }
        }

        /// <summary>
        /// Ghi dữ liệu BOM vào một sheet
        /// </summary>
        private void WriteBomToSheet(dynamic worksheet, string deviceCode,
            Dictionary<string, BomItem> bomDict, bool isMF)
        {
            // Header
            worksheet.Cells[1, 1].Value = "Mã Thiết Bị";
            worksheet.Cells[1, 2].Value = "STT";
            worksheet.Cells[1, 3].Value = "Tên Chi Tiết";
            worksheet.Cells[1, 4].Value = "Loại";
            worksheet.Cells[1, 5].Value = "Vật Liệu";
            worksheet.Cells[1, 6].Value = "Bề Mặt";
            worksheet.Cells[1, 7].Value = "Hãng SX";
            worksheet.Cells[1, 8].Value = "SL/1 Set";
            worksheet.Cells[1, 9].Value = "SL Tổng";

            // Format header
            var headerRange = worksheet.Range["A1:I1"];
            headerRange.Interior.Color = 0x4472C4; // Xanh đậm
            headerRange.Font.Color = 0xFFFFFF;
            headerRange.Font.Bold = true;

            int row = 2;
            int stt = 1;

            foreach (var kvp in bomDict)
            {
                BomItem item = kvp.Value;

                worksheet.Cells[row, 1].Value = deviceCode;
                worksheet.Cells[row, 2].Value = stt++;
                worksheet.Cells[row, 3].Value = item.PartName;
                worksheet.Cells[row, 4].Value = item.PartType;
                worksheet.Cells[row, 5].Value = item.Material;
                worksheet.Cells[row, 6].Value = item.SurfaceTreatment;
                worksheet.Cells[row, 7].Value = item.Manufacturer;
                worksheet.Cells[row, 8].Value = item.QuantityPerSet;
                worksheet.Cells[row, 9].Value = item.Quantity;

                // Highlight màu theo loại
                var cellRange = worksheet.Range[$"A{row}:I{row}"];
                if (isMF)
                {
                    cellRange.Interior.Color = 0xE7E6E6; // Xám nhạt
                }
                else
                {
                    cellRange.Interior.Color = 0xF0F0F0; // Xám rất nhạt
                }

                row++;
            }

            // Auto-fit columns
            worksheet.Columns.AutoFit();
        }

        /// <summary>
        /// Lấy giá trị property từ document
        /// </summary>
        private string GetPropertyValue(ModelDoc2 swModel, string propName)
        {
            try
            {
                CustomPropertyManager propMgr = swModel.Extension.get_CustomPropertyManager("");
                string valOut = "";
                string resValOut = "";
                bool wasResolved = false;
                bool linkToProp = false;

                propMgr.Get6(propName, false, out valOut, out resValOut, out wasResolved, out linkToProp);
                return resValOut;
            }
            catch
            {
                return "";
            }
        }
    }
}