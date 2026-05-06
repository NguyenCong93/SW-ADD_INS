using System;
using System.Collections.Generic;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomationAddin
{
    public class BomItem
    {
        public string PartName { get; set; }
        public string PartType { get; set; } // MF or PUR
        public string Material { get; set; }
        public string SurfaceTreatment { get; set; }
        public int Quantity { get; set; }

        public string UniqueKey 
        { 
            get { return string.Format("{0}|{1}|{2}|{3}", PartName, PartType, Material, SurfaceTreatment); } 
        }
    }

    public class BomManager
    {
        private ISldWorks _swApp;

        public BomManager(ISldWorks swApp)
        {
            _swApp = swApp;
        }

        public void ExportBomToExcel()
        {
            try
            {
                ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
                if (swModel == null || swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    _swApp.SendMsgToUser("Vui lòng mở một file Assembly để xuất BOM.");
                    return;
                }

                AssemblyDoc swAssy = (AssemblyDoc)swModel;
                object[] components = (object[])swAssy.GetComponents(false);

                if (components == null || components.Length == 0)
                {
                    _swApp.SendMsgToUser("Assembly không có chi tiết nào.");
                    return;
                }

                Dictionary<string, BomItem> bomDict = new Dictionary<string, BomItem>();

                foreach (object compObj in components)
                {
                    Component2 swComp = (Component2)compObj;

                    // Bỏ qua các chi tiết bị ẩn hoặc suppressed
                    if (swComp.IsSuppressed() || swComp.IsHidden(true)) continue;

                    ModelDoc2 compModel = (ModelDoc2)swComp.GetModelDoc2();
                    if (compModel == null) continue;

                    CustomPropertyManager propMgr = compModel.Extension.get_CustomPropertyManager("");
                    
                    string partName = GetPropertyValue(propMgr, "Part Name");
                    if (string.IsNullOrEmpty(partName)) partName = swComp.Name2; // Fallback

                    string partType = GetPropertyValue(propMgr, "Part Type");
                    string material = GetPropertyValue(propMgr, "Material");
                    string surface = GetPropertyValue(propMgr, "Surface Treatment");

                    BomItem item = new BomItem()
                    {
                        PartName = partName,
                        PartType = partType,
                        Material = material,
                        SurfaceTreatment = surface,
                        Quantity = 1
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

                // Ghi ra Excel dùng Late Binding
                WriteToExcel(bomDict, swModel.GetPathName());
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi trong quá trình quét dữ liệu BOM: " + ex.Message);
            }
        }

        private string GetPropertyValue(CustomPropertyManager propMgr, string propName)
        {
            string valOut = "";
            string resValOut = "";
            bool wasResolved = false;
            bool linkToProp = false;
            propMgr.Get6(propName, false, out valOut, out resValOut, out wasResolved, out linkToProp);
            return resValOut;
        }

        private void WriteToExcel(Dictionary<string, BomItem> bomDict, string assyPath)
        {
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    _swApp.SendMsgToUser("Không tìm thấy Microsoft Excel trên máy tính.");
                    return;
                }

                dynamic excelApp = Activator.CreateInstance(excelType);
                excelApp.Visible = true;
                dynamic workbooks = excelApp.Workbooks;
                dynamic workbook = workbooks.Add();
                dynamic worksheet = workbook.Worksheets[1];

                worksheet.Name = "BOM List";

                // Headers
                worksheet.Cells[1, 1].Value = "STT";
                worksheet.Cells[1, 2].Value = "Tên Chi Tiết";
                worksheet.Cells[1, 3].Value = "Loại";
                worksheet.Cells[1, 4].Value = "Vật Liệu";
                worksheet.Cells[1, 5].Value = "Bề Mặt";
                worksheet.Cells[1, 6].Value = "Số Lượng";

                int row = 2;
                int stt = 1;

                foreach (var kvp in bomDict)
                {
                    BomItem item = kvp.Value;
                    worksheet.Cells[row, 1].Value = stt++;
                    worksheet.Cells[row, 2].Value = item.PartName;
                    worksheet.Cells[row, 3].Value = item.PartType;
                    worksheet.Cells[row, 4].Value = item.Material;
                    worksheet.Cells[row, 5].Value = item.SurfaceTreatment;
                    worksheet.Cells[row, 6].Value = item.Quantity;

                    // Nếu là hàng gia công (MF) thì có thể highlight màu hoặc định dạng riêng
                    if (item.PartType == "MF")
                    {
                        worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 6]].Interior.Color = 14281213; // Màu xám nhạt
                    }
                    else if (item.PartType == "PUR")
                    {
                        worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 6]].Interior.Color = 14342874; // Màu xanh nhạt
                    }

                    row++;
                }

                // Căn chỉnh tự động kích thước cột
                worksheet.Columns.AutoFit();

                // Tuỳ chọn lưu file tự động cùng thư mục gốc
                if (!string.IsNullOrEmpty(assyPath))
                {
                    string dir = Path.GetDirectoryName(assyPath);
                    string name = Path.GetFileNameWithoutExtension(assyPath);
                    string excelPath = Path.Combine(dir, name + "_BOM.xlsx");
                    
                    // Nếu file đã tồn tại thì tự động xóa hoặc bỏ qua, ở đây dùng cảnh báo
                    if (!File.Exists(excelPath))
                    {
                        workbook.SaveAs(excelPath);
                        _swApp.SendMsgToUser("Đã xuất BOM thành công: " + excelPath);
                    }
                    else
                    {
                        _swApp.SendMsgToUser("Đã tạo bảng BOM. Vui lòng lưu file Excel thủ công vì file BOM đã tồn tại.");
                    }
                }
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi khi xuất ra Excel: " + ex.Message);
            }
        }
    }
}
