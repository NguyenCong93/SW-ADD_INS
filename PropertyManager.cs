using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomationAddin
{
    public class PropertyManager
    {
        private ISldWorks _swApp;

        public PropertyManager(ISldWorks swApp)
        {
            _swApp = swApp;
        }

        /// <summary>
        /// Ghi các thuộc tính vào file hiện tại, hoặc các chi tiết (components) đang được chọn nếu ở trong Assembly
        /// </summary>
        public void WritePropertiesToCurrentDoc(
            string unitName, string partName, string material, 
            string designer, string surfaceTreatment, string heatTreatment, string partType)
        {
            try
            {
                ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;

                if (swModel == null)
                {
                    _swApp.SendMsgToUser("Không có tài liệu nào đang mở.");
                    return;
                }

                // Nếu là Assembly, kiểm tra xem có component nào được chọn không
                if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
                    int selCount = swSelMgr.GetSelectedObjectCount2(-1);
                    
                    if (selCount > 0)
                    {
                        int successCount = 0;
                        for (int i = 1; i <= selCount; i++)
                        {
                            Component2 swComp = (Component2)swSelMgr.GetSelectedObjectsComponent4(i, -1);
                            if (swComp != null)
                            {
                                ModelDoc2 compModel = (ModelDoc2)swComp.GetModelDoc2();
                                if (compModel != null)
                                {
                                    CustomPropertyManager custPropMgr = compModel.Extension.get_CustomPropertyManager("");
                                    WritePropsToMgr(custPropMgr, unitName, partName, material, designer, surfaceTreatment, heatTreatment, partType);
                                    successCount++;
                                }
                            }
                        }
                        if (successCount > 0)
                        {
                            _swApp.SendMsgToUser(string.Format("Đã cập nhật thuộc tính cho {0} chi tiết được chọn.", successCount));
                            return;
                        }
                    }
                }

                // Nếu không có gì được chọn hoặc không phải Assembly, ghi cho file hiện tại
                CustomPropertyManager activeDocPropMgr = swModel.Extension.get_CustomPropertyManager("");
                WritePropsToMgr(activeDocPropMgr, unitName, partName, material, designer, surfaceTreatment, heatTreatment, partType);
                
                _swApp.SendMsgToUser("Đã cập nhật Custom Properties cho tài liệu hiện tại.");
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi trong quá trình ghi thuộc tính: " + ex.Message);
            }
        }

        private void WritePropsToMgr(CustomPropertyManager custPropMgr, string unitName, string partName, string material, string designer, string surfaceTreatment, string heatTreatment, string partType)
        {
            AddOrUpdateProperty(custPropMgr, "Unit Name", unitName);
            AddOrUpdateProperty(custPropMgr, "Part Name", partName);
            AddOrUpdateProperty(custPropMgr, "Material", material);
            AddOrUpdateProperty(custPropMgr, "Designer", designer);
            AddOrUpdateProperty(custPropMgr, "Surface Treatment", surfaceTreatment);
            AddOrUpdateProperty(custPropMgr, "Heat Treatment", heatTreatment);
            AddOrUpdateProperty(custPropMgr, "Part Type", partType);
        }

        private void AddOrUpdateProperty(CustomPropertyManager mgr, string propName, string propValue)
        {
            if (string.IsNullOrEmpty(propValue))
                return;

            int res = mgr.Add3(propName, (int)swCustomInfoType_e.swCustomInfoText, propValue, (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
            if (res != (int)swCustomInfoAddResult_e.swCustomInfoAddResult_AddedOrChanged)
            {
                // Fallback nếu có lỗi
                mgr.Set2(propName, propValue);
            }
        }

        /// <summary>
        /// Tạo tên file mới dựa theo quy tắc: [Project ID]-[Unit]-[Type]-[ASM]-[Serial]-[Rev]
        /// </summary>
        public string GenerateFileName(string projectId, string unit, string type, string asm, string serial, string rev)
        {
            return string.Format("{0}-{1}-{2}-{3}-{4}-{5}", projectId, unit, type, asm, serial, rev);
        }

        /// <summary>
        /// Đổi tên file (Save As Copy and Open / RenameDocument) 
        /// Theo yêu cầu, ta giữ nguyên tên file gốc và chỉ dùng tên mới cho xuất STP/X-T ở module khác.
        /// Tuy nhiên, nếu muốn đổi tên part trong Feature Tree của Assembly thì có thể dùng hàm này.
        /// </summary>
        public void RenameActiveDocumentInTree(string newName)
        {
            ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
            if (swModel == null) return;
            
            // Rename ở mức Component trong Assembly (nếu đang chọn Component)
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            Component2 swComp = (Component2)swSelMgr.GetSelectedObjectsComponent4(1, -1);

            if (swComp != null)
            {
                swComp.Name2 = newName;
                swModel.EditRebuild3();
                _swApp.SendMsgToUser("Đã đổi tên component trong Feature Tree thành: " + newName);
            }
            else
            {
                _swApp.SendMsgToUser("Vui lòng chọn một Component trong Assembly để đổi tên.");
            }
        }
    }
}
