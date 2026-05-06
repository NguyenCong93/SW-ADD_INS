using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomationAddin
{
    public class DrawingManager
    {
        private ISldWorks _swApp;

        public DrawingManager(ISldWorks swApp)
        {
            _swApp = swApp;
        }

        public void CreateDrawingAndAutoDim()
        {
            try
            {
                ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
                if (swModel == null)
                {
                    _swApp.SendMsgToUser("Vui lòng mở một Part hoặc Assembly để tạo bản vẽ.");
                    return;
                }

                int docType = swModel.GetType();
                if (docType != (int)swDocumentTypes_e.swDocPART && docType != (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    _swApp.SendMsgToUser("Chỉ có thể tạo bản vẽ từ Part hoặc Assembly.");
                    return;
                }

                string modelPath = swModel.GetPathName();
                if (string.IsNullOrEmpty(modelPath))
                {
                    _swApp.SendMsgToUser("Vui lòng lưu Part/Assembly trước khi tạo bản vẽ.");
                    return;
                }

                // Lấy template bản vẽ mặc định
                string templatePath = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);
                if (string.IsNullOrEmpty(templatePath))
                {
                    _swApp.SendMsgToUser("Không tìm thấy Template bản vẽ mặc định.");
                    return;
                }

                // Tạo bản vẽ mới
                ModelDoc2 swDrawing = (ModelDoc2)_swApp.NewDocument(templatePath, (int)swDwgPaperSizes_e.swDwgPaperA3size, 0, 0);
                if (swDrawing == null)
                {
                    _swApp.SendMsgToUser("Không thể tạo bản vẽ mới.");
                    return;
                }

                DrawingDoc swDrawDoc = (DrawingDoc)swDrawing;

                // Thiết lập khổ giấy và scale tổng quát
                swDrawDoc.SetupSheet5("Sheet1", (int)swDwgPaperSizes_e.swDwgPaperA3size, (int)swDwgTemplates_e.swDwgTemplateA3size, 1, 10, true, templatePath, 0.42, 0.297, "Default", true);

                // Chèn hình chiếu chuẩn: Front, Top, Right, Iso
                // Tọa độ tạm thời, cần scale dựa theo bounding box sau này
                View viewFront = swDrawDoc.CreateDrawViewFromModelView3(modelPath, "*Front", 0.1, 0.1, 0);
                View viewTop = swDrawDoc.CreateDrawViewFromModelView3(modelPath, "*Top", 0.1, 0.2, 0);
                View viewRight = swDrawDoc.CreateDrawViewFromModelView3(modelPath, "*Right", 0.2, 0.1, 0);
                View viewIso = swDrawDoc.CreateDrawViewFromModelView3(modelPath, "*Isometric", 0.3, 0.2, 0);

                if (viewFront != null)
                {
                    // Gọi chức năng Auto-DIM (Insert Model Items) cho Front View
                    swDrawDoc.ActivateView(viewFront.Name);
                    
                    // Chèn các kích thước Driving Dimension (thuộc tính chuẩn ASME Y14.5 sẽ cấu hình trong SolidWorks settings)
                    // 0 is usually the default for swInsertAnnotations_e.swInsertDimensionsMarkedForDrawing
                    // 32768 is usually swInsertAnnotation_e.swInsertDimensions
                    swDrawDoc.InsertModelAnnotations3(0, 32768, true, true, false, false);
                }

                if (viewTop != null)
                {
                    swDrawDoc.ActivateView(viewTop.Name);
                    swDrawDoc.InsertModelAnnotations3(0, 32768, true, true, false, false);
                }

                if (viewRight != null)
                {
                    swDrawDoc.ActivateView(viewRight.Name);
                    swDrawDoc.InsertModelAnnotations3(0, 32768, true, true, false, false);
                }

                // Sắp xếp lại kích thước tự động (Auto-Arrange)
                swDrawing.Extension.AlignDimensions((int)swAlignDimensionType_e.swAlignDimensionType_AutoArrange, 0.01);

                swDrawing.ForceRebuild3(true);
                _swApp.SendMsgToUser("Đã tạo bản vẽ và Auto-DIM thành công!");
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi trong quá trình tạo bản vẽ: " + ex.Message);
            }
        }
    }
}
