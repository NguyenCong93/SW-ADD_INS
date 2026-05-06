using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomationAddin
{
    [ComVisible(true)]
    [ProgId("SwAutomationAddin.MainTaskPane")]
    public class MainTaskPane : UserControl
    {
        private ISldWorks _swApp;
        private PropertyManager _propMgr;
        private InterferenceManager _interferenceMgr;
        private ExportManager _exportMgr;
        private BomManager _bomMgr;
        private DrawingManager _drawingMgr;

        // UI Controls - TabControl
        private TabControl mainTabControl;
        private TabPage tabExport, tabProps, tabRename, tabDrawing, tabAbout;

        // Utilities
        private GroupBox gbUtilities;
        private Button btnCheckClash, btnNextClash, btnPrevClash, btnEndClash;
        private Button btnPackAndGo;

        // Export Tab Controls
        // PDF
        private CheckBox chkPdf2D, chkPdf3D;
        private RadioButton radPdfSheetActive, radPdfSheetAll, radPdfSheetSeparate, radPdfSheetRange;
        private TextBox txtPdfSheetRange;
        // DXF/DWG
        private CheckBox chkDxf, chkDwg;
        private ComboBox cbDxfVersion;
        // STP/XT
        private CheckBox chkStp, chkXt;
        // Naming options (Shared for all exports)
        private RadioButton radNameDrawingNo, radNameUnitName, radNameMachineName;
        // Export Action
        private Button btnDoExport;
        private Button btnExportBomList;

        // Properties Tab Controls
        private Button btnSetupProperties;
        
        // Rename Tab Controls
        private Button btnBatchRename;

        // Drawing Tab Controls
        private Button btnAutoDrawing;

        private System.Windows.Forms.Timer _uiTimer;

        public MainTaskPane()
        {
            InitializeUI();
            
            _uiTimer = new System.Windows.Forms.Timer();
            _uiTimer.Interval = 1000;
            _uiTimer.Tick += UiTimer_Tick;
            _uiTimer.Start();
        }

        private void UiTimer_Tick(object sender, EventArgs e)
        {
            if (_swApp == null || gbUtilities == null) return;
            try
            {
                ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
                if (swModel != null && swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    btnCheckClash.Enabled = true;
                }
                else
                {
                    btnCheckClash.Enabled = false;
                }
            }
            catch { }
        }

        public void Setup(ISldWorks swApp)
        {
            _swApp = swApp;
            _propMgr = new PropertyManager(_swApp);
            _interferenceMgr = new InterferenceManager(_swApp);
            _exportMgr = new ExportManager(_swApp);
            _bomMgr = new BomManager(_swApp);
            _drawingMgr = new DrawingManager(_swApp);
        }

        private void InitializeUI()
        {
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            this.AutoScroll = true;
            this.Padding = new Padding(5);
            this.Width = 320;
            this.Height = 1000;

            mainTabControl = new TabControl();
            mainTabControl.Width = 300;
            mainTabControl.Height = 580;
            mainTabControl.Location = new Point(5, 5);

            tabRename = new TabPage("🔄 Đổi Tên");
            tabProps = new TabPage("⚙️ Thuộc Tính");
            tabExport = new TabPage("📥 Xuất File");
            tabDrawing = new TabPage("📐 Bản Vẽ");
            tabAbout = new TabPage("ℹ️ About");

            mainTabControl.TabPages.Add(tabRename);
            mainTabControl.TabPages.Add(tabProps);
            mainTabControl.TabPages.Add(tabExport);
            mainTabControl.TabPages.Add(tabDrawing);
            mainTabControl.TabPages.Add(tabAbout);

            BuildRenameTab(tabRename);
            BuildPropsTab(tabProps);
            BuildExportTab(tabExport);
            BuildDrawingTab(tabDrawing);
            BuildAboutTab(tabAbout);

            this.Controls.Add(mainTabControl);

            // --- Group: Utilities (Always visible at bottom) ---
            gbUtilities = new GroupBox() { Text = "TIỆN ÍCH NHANH", Font = new Font("Segoe UI", 9F, FontStyle.Bold), Location = new Point(5, 590), Width = 300, Height = 145 };
            
            btnCheckClash = new Button() { Text = "Kiểm Tra Va Chạm", Location = new Point(10, 25), Width = 280, Height = 30, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            StyleButton(btnCheckClash);
            btnCheckClash.Click += BtnCheckClash_Click;
            gbUtilities.Controls.Add(btnCheckClash);

            btnPrevClash = new Button() { Text = "<< Prev", Location = new Point(10, 65), Width = 85, Height = 30 };
            StyleButton(btnPrevClash);
            btnPrevClash.Enabled = false;
            btnPrevClash.Click += BtnPrevClash_Click;
            gbUtilities.Controls.Add(btnPrevClash);

            btnNextClash = new Button() { Text = "Next >>", Location = new Point(105, 65), Width = 85, Height = 30 };
            StyleButton(btnNextClash);
            btnNextClash.Enabled = false;
            btnNextClash.Click += BtnNextClash_Click;
            gbUtilities.Controls.Add(btnNextClash);

            btnEndClash = new Button() { Text = "X Kết Thúc", Location = new Point(200, 65), Width = 90, Height = 30 };
            StyleButton(btnEndClash);
            btnEndClash.Enabled = false;
            btnEndClash.Click += BtnEndClash_Click;
            gbUtilities.Controls.Add(btnEndClash);

            btnPackAndGo = new Button() { Text = "Pack and Go", Location = new Point(10, 105), Width = 280, Height = 30, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            StyleButton(btnPackAndGo);
            // btnPackAndGo.Click += ...
            gbUtilities.Controls.Add(btnPackAndGo);

            this.Controls.Add(gbUtilities);
        }

        private void StyleButton(Button btn)
        {
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderColor = Color.FromArgb(204, 204, 204);
            btn.BackColor = Color.FromArgb(245, 245, 245);
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(229, 243, 255); 
            btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(204, 232, 255); 
            btn.Cursor = Cursors.Hand;
        }

        private void BuildExportTab(TabPage page)
        {
            page.AutoScroll = true;
            int yPos = 10;

            // Naming Rule
            GroupBox gbName = new GroupBox() { Text = "Tên file", Location = new Point(5, yPos), Width = 280, Height = 95 };
            radNameDrawingNo = new RadioButton() { Text = "Drawing NO (Chi tiết lẻ)", Location = new Point(10, 20), Width = 250, Checked = true };
            radNameUnitName = new RadioButton() { Text = "Unit Name (Cụm con)", Location = new Point(10, 45), Width = 250 };
            radNameMachineName = new RadioButton() { Text = "Machine Name (Cụm tổng)", Location = new Point(10, 70), Width = 250 };
            gbName.Controls.Add(radNameDrawingNo);
            gbName.Controls.Add(radNameUnitName);
            gbName.Controls.Add(radNameMachineName);
            page.Controls.Add(gbName);
            yPos += 105;

            // Format Selection
            GroupBox gbFormat = new GroupBox() { Text = "Định dạng file", Location = new Point(5, yPos), Width = 280, Height = 100 };
            chkPdf2D = new CheckBox() { Text = "PDF 2D", Location = new Point(10, 20), Width = 100, Checked = true };
            chkPdf3D = new CheckBox() { Text = "PDF 3D (Part/Asm)", Location = new Point(120, 20), Width = 150 };
            
            chkDxf = new CheckBox() { Text = "DXF", Location = new Point(10, 45), Width = 60 };
            chkDwg = new CheckBox() { Text = "DWG", Location = new Point(70, 45), Width = 60 };
            cbDxfVersion = new ComboBox() { Location = new Point(140, 43), Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            cbDxfVersion.Items.AddRange(new object[] { "2013", "2018" });
            cbDxfVersion.SelectedIndex = 0;
            
            chkStp = new CheckBox() { Text = "STEP", Location = new Point(10, 70), Width = 80 };
            chkXt = new CheckBox() { Text = "X_T", Location = new Point(120, 70), Width = 80 };
            
            gbFormat.Controls.Add(chkPdf2D); gbFormat.Controls.Add(chkPdf3D);
            gbFormat.Controls.Add(chkDxf); gbFormat.Controls.Add(chkDwg); gbFormat.Controls.Add(cbDxfVersion);
            gbFormat.Controls.Add(chkStp); gbFormat.Controls.Add(chkXt);
            page.Controls.Add(gbFormat);
            yPos += 110;

            // Sheet Options (For PDF/DWG Drawing)
            GroupBox gbSheet = new GroupBox() { Text = "Tuỳ Chọn Sheet (Bản Vẽ)", Location = new Point(5, yPos), Width = 280, Height = 120 };
            radPdfSheetActive = new RadioButton() { Text = "Sheet đang mở", Location = new Point(10, 20), Width = 250, Checked = true };
            radPdfSheetAll = new RadioButton() { Text = "Gộp tất cả sheet vào 1 file", Location = new Point(10, 45), Width = 250 };
            radPdfSheetSeparate = new RadioButton() { Text = "Tách mỗi sheet 1 file", Location = new Point(10, 70), Width = 250 };
            radPdfSheetRange = new RadioButton() { Text = "Dải sheet:", Location = new Point(10, 95), Width = 80 };
            txtPdfSheetRange = new TextBox() { Location = new Point(95, 93), Width = 150 };
            gbSheet.Controls.Add(radPdfSheetActive); gbSheet.Controls.Add(radPdfSheetAll);
            gbSheet.Controls.Add(radPdfSheetSeparate); gbSheet.Controls.Add(radPdfSheetRange);
            gbSheet.Controls.Add(txtPdfSheetRange);
            page.Controls.Add(gbSheet);
            yPos += 130;

            btnDoExport = new Button() { Text = "THỰC HIỆN XUẤT FILE", Location = new Point(5, yPos), Width = 280, Height = 40, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            StyleButton(btnDoExport);
            page.Controls.Add(btnDoExport);
            yPos += 50;

            // BOM List 
            GroupBox gbBom = new GroupBox() { Text = "Xuất BOM List Excel", Location = new Point(5, yPos), Width = 280, Height = 80 };
            btnExportBomList = new Button() { Text = "Xuất BOM (Tiêu chuẩn & Gia công)", Location = new Point(10, 25), Width = 260, Height = 40 };
            StyleButton(btnExportBomList);
            gbBom.Controls.Add(btnExportBomList);
            page.Controls.Add(gbBom);
        }

        private void BuildPropsTab(TabPage page)
        {
            btnSetupProperties = new Button() { Text = "Cài Đặt System Properties Mặc Định", Location = new Point(10, 20), Width = 270, Height = 50, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            StyleButton(btnSetupProperties);
            page.Controls.Add(btnSetupProperties);
            
            Label lblInfo = new Label() { Text = "Tính năng này sẽ tự động link đường dẫn System Options của SolidWorks tới thư mục:\n\nC:\\VISC_SOLIDWORKS\\1. VISC_TEMPLATE\\My Custom Property Files\n\nGiúp bạn dùng chuẩn chung ngay lập tức.", Location = new Point(10, 80), Width = 270, Height = 100 };
            page.Controls.Add(lblInfo);
        }

        private void BuildRenameTab(TabPage page)
        {
            btnBatchRename = new Button() { Text = "Đổi Tên Chi Tiết Hàng Loạt", Location = new Point(10, 20), Width = 270, Height = 50, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            StyleButton(btnBatchRename);
            page.Controls.Add(btnBatchRename);

            Label lblInfo = new Label() { Text = "(Giao diện tính năng Đổi tên hàng loạt sẽ được phát triển sau khi thống nhất luồng)", Location = new Point(10, 80), Width = 270, Height = 50 };
            page.Controls.Add(lblInfo);
        }

        private void BuildDrawingTab(TabPage page)
        {
            btnAutoDrawing = new Button() { Text = "Tạo Bản Vẽ Tự Động", Location = new Point(10, 20), Width = 270, Height = 50, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            StyleButton(btnAutoDrawing);
            page.Controls.Add(btnAutoDrawing);

            Label lblInfo = new Label() { Text = "- Part: 1 Sheet\n- Assembly: Multi-sheet (BOM + Tổng thể + Các chi tiết lẻ)", Location = new Point(10, 80), Width = 270, Height = 100 };
            page.Controls.Add(lblInfo);
        }

        private void BuildAboutTab(TabPage page)
        {
            Label lblTitle = new Label() { Text = "VISC SolidWorks Add-in", Location = new Point(10, 30), Width = 270, Height = 30, TextAlign = ContentAlignment.TopCenter, Font = new Font("Segoe UI", 12F, FontStyle.Bold) };
            Label lblDesc = new Label() { Text = "Phiên bản: 1.0.0\nPhát triển bởi đội ngũ R&D", Location = new Point(10, 60), Width = 270, Height = 50, TextAlign = ContentAlignment.TopCenter };
            page.Controls.Add(lblTitle);
            page.Controls.Add(lblDesc);
        }

        private void BtnCheckClash_Click(object sender, EventArgs e)
        {
            if (_interferenceMgr == null) return;
            try
            {
                int count = _interferenceMgr.RunInterferenceDetectionAndHighlight();
                if (count > 0)
                {
                    btnCheckClash.Text = string.Format("Kiểm Tra Va Chạm ({0} lỗi)", count);
                    btnNextClash.Enabled = true;
                    btnPrevClash.Enabled = true;
                    btnEndClash.Enabled = true;
                    _interferenceMgr.ZoomToNextClash();
                }
                else if (count == 0)
                {
                    btnCheckClash.Text = "Kiểm Tra Va Chạm (0 lỗi)";
                    btnNextClash.Enabled = false;
                    btnPrevClash.Enabled = false;
                    btnEndClash.Enabled = true;
                }
                else
                {
                    btnCheckClash.Text = "Kiểm Tra Va Chạm";
                    btnNextClash.Enabled = false;
                    btnPrevClash.Enabled = false;
                    btnEndClash.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi khi kiểm tra va chạm: " + ex.Message);
            }
        }

        private void BtnNextClash_Click(object sender, EventArgs e)
        {
            if (_interferenceMgr == null) return;
            _interferenceMgr.ZoomToNextClash();
        }

        private void BtnPrevClash_Click(object sender, EventArgs e)
        {
            if (_interferenceMgr == null) return;
            _interferenceMgr.ZoomToPrevClash();
        }

        private void BtnEndClash_Click(object sender, EventArgs e)
        {
            if (_interferenceMgr == null) return;
            _interferenceMgr.EndClashDetection();
            btnCheckClash.Text = "Kiểm Tra Va Chạm";
            btnNextClash.Enabled = false;
            btnPrevClash.Enabled = false;
            btnEndClash.Enabled = false;
        }
    }
}
