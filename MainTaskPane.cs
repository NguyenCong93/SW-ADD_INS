using System;
using System.ComponentModel;
using System.Drawing;
using System.Threading.Tasks;
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
        #region Fields
        private ISldWorks _swApp;
        private PropertyManager _propMgr;
        private InterferenceManager _interferenceMgr;
        private ExportManager _exportMgr;
        private BomManager _bomMgr;
        private DrawingManager _drawingMgr;
        private bool _disposed = false;

        // UI Style Constants
        private static readonly Font UIFont = new Font("Segoe UI", 9F);
        private static readonly Font UIFontBold = new Font("Segoe UI", 9F, FontStyle.Bold);
        private static readonly Font UIFontLarge = new Font("Segoe UI", 12F, FontStyle.Bold);
        private static readonly Color BtnBg = Color.FromArgb(245, 245, 245);
        private static readonly Color BtnBorder = Color.FromArgb(204, 204, 204);
        private static readonly Color BtnHover = Color.FromArgb(229, 243, 255);
        private static readonly Color BtnPress = Color.FromArgb(204, 232, 255);
        private const int FIELD_HEIGHT = 28;
        private const int MARGIN = 5;
        private const int CONTENT_WIDTH = 280;

        // UI Controls
        private TabControl mainTabControl;
        private TabPage tabExport, tabProps, tabRename, tabDrawing, tabAbout;
        private Panel pnlUtilities;
        private GroupBox gbUtilities;
        private Button btnCheckClash, btnNextClash, btnPrevClash, btnEndClash, btnPackAndGo;
        private CheckBox chkPdf2D, chkPdf3D, chkDxf, chkDwg, chkStp, chkXt;
        private RadioButton radPdfSheetActive, radPdfSheetAll, radPdfSheetSeparate, radPdfSheetRange;
        private TextBox txtPdfSheetRange;
        private ComboBox cbDxfVersion;
        private RadioButton radNameDrawingNo, radNameUnitName, radNameMachineName;
        private Button btnDoExport, btnExportBomList;
        private Button btnSetupProperties, btnApplyProps;
        private TextBox txtPropTitle, txtPropPartName, txtPropUnit, txtPropNameProject;
        private TextBox txtPropMaterial, txtPropSurface, txtPropHeat, txtPropManufacturer;
        private ComboBox cbPropDesign, cbPropMakeOrBuy;
        private Button btnBatchRename, btnAutoDrawing;
        private StatusStrip statusBar;
        private ToolStripStatusLabel lblStatus;
        #endregion

        #region Constructor & Dispose
        public MainTaskPane()
        {
            InitializeUI();
        }

        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Release managed resources
                    if (statusBar != null) statusBar.Dispose();
                }
                // Release COM references
                _interferenceMgr = null;
                _exportMgr = null;
                _bomMgr = null;
                _drawingMgr = null;
                _propMgr = null;
                _swApp = null;
                _disposed = true;
            }
            base.Dispose(disposing);
        }
        #endregion

        #region Setup & SW Event Hooks
        public void Setup(ISldWorks swApp)
        {
            _swApp = swApp;
            _propMgr = new PropertyManager(_swApp);
            _interferenceMgr = new InterferenceManager(_swApp);
            _exportMgr = new ExportManager(_swApp);
            _bomMgr = new BomManager(_swApp);
            _drawingMgr = new DrawingManager(_swApp);

            // Hook SW events instead of polling with Timer
            try
            {
                ((SldWorks)_swApp).ActiveDocChangeNotify += OnActiveDocChanged;
                ((SldWorks)_swApp).FileOpenPostNotify += delegate(string f) { UpdateUIState(); return 0; };
                ((SldWorks)_swApp).FileCloseNotify += delegate(string f, int r) { BeginInvoke((Action)UpdateUIState); return 0; };
            }
            catch { /* Fallback: UI state updated on tab switch */ }

            UpdateUIState();
        }

        private int OnActiveDocChanged()
        {
            if (InvokeRequired)
                BeginInvoke((Action)UpdateUIState);
            else
                UpdateUIState();
            return 0;
        }

        private void UpdateUIState()
        {
            if (_swApp == null) return;
            try
            {
                ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
                bool isAsm = swModel != null && swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY;
                bool isDrw = swModel != null && swModel.GetType() == (int)swDocumentTypes_e.swDocDRAWING;

                btnCheckClash.Enabled = isAsm;
                btnExportBomList.Enabled = isAsm;
                chkPdf3D.Enabled = !isDrw;
                chkStp.Enabled = !isDrw;
                chkXt.Enabled = !isDrw;

                SetStatus(swModel != null ? "Đang mở: " + swModel.GetTitle() : "Chưa mở file");
            }
            catch { }
        }
        #endregion

        #region UI Helpers
        private void StyleButton(Button btn)
        {
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderColor = BtnBorder;
            btn.BackColor = BtnBg;
            btn.FlatAppearance.MouseOverBackColor = BtnHover;
            btn.FlatAppearance.MouseDownBackColor = BtnPress;
            btn.Cursor = Cursors.Hand;
            btn.Font = UIFont;
        }

        private Button CreateButton(string text, bool bold = false)
        {
            var btn = new Button() { Text = text, Height = 35, Font = bold ? UIFontBold : UIFont, Dock = DockStyle.Top };
            StyleButton(btn);
            return btn;
        }

        private Label CreateLabel(string text)
        {
            return new Label() { Text = text, Width = 95, Height = 20, TextAlign = ContentAlignment.MiddleLeft, Font = UIFont };
        }

        private void SetStatus(string msg)
        {
            if (lblStatus != null) lblStatus.Text = msg;
        }

        private void RunAsync(Action work, Action onDone)
        {
            SetStatus("Đang xử lý...");
            this.Enabled = false;
            Task.Run(() =>
            {
                try { work(); }
                catch (Exception ex)
                {
                    BeginInvoke((Action)(() => {
                        if (_swApp != null) _swApp.SendMsgToUser("Lỗi: " + ex.Message);
                        SetStatus("Lỗi!");
                    }));
                    return;
                }
                BeginInvoke((Action)(() => { this.Enabled = true; if (onDone != null) onDone(); }));
            });
        }
        #endregion

        #region InitializeUI
        private void InitializeUI()
        {
            this.Font = UIFont;
            this.AutoScroll = false;
            this.Dock = DockStyle.Fill;

            // Status bar at very bottom
            statusBar = new StatusStrip();
            lblStatus = new ToolStripStatusLabel("Sẵn sàng");
            statusBar.Items.Add(lblStatus);
            statusBar.Dock = DockStyle.Bottom;
            this.Controls.Add(statusBar);

            // Utilities panel at bottom (above status bar)
            pnlUtilities = new Panel() { Dock = DockStyle.Bottom, Height = 155, Padding = new Padding(MARGIN) };
            gbUtilities = new GroupBox() { Text = "TIỆN ÍCH NHANH", Font = UIFontBold, Dock = DockStyle.Fill };
            BuildUtilitiesGroup(gbUtilities);
            pnlUtilities.Controls.Add(gbUtilities);
            this.Controls.Add(pnlUtilities);

            // Tab control fills remaining space
            mainTabControl = new TabControl() { Dock = DockStyle.Fill, Multiline = true, Font = UIFont };

            tabRename  = new TabPage("🔄 Đổi Tên");
            tabProps   = new TabPage("⚙️ Thuộc Tính");
            tabExport  = new TabPage("📥 Xuất File");
            tabDrawing = new TabPage("📐 Bản Vẽ");
            tabAbout   = new TabPage("ℹ️ About");

            mainTabControl.TabPages.AddRange(new[] { tabRename, tabProps, tabExport, tabDrawing, tabAbout });

            BuildRenameTab(tabRename);
            BuildPropsTab(tabProps);
            BuildExportTab(tabExport);
            BuildDrawingTab(tabDrawing);
            BuildAboutTab(tabAbout);

            this.Controls.Add(mainTabControl);
        }

        private void BuildUtilitiesGroup(GroupBox gb)
        {
            var tbl = new TableLayoutPanel() { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 3 };
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33));
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));

            btnCheckClash = CreateButton("Kiểm Tra Va Chạm", true);
            btnCheckClash.Dock = DockStyle.Fill;
            btnCheckClash.Click += BtnCheckClash_Click;
            tbl.Controls.Add(btnCheckClash, 0, 0);
            tbl.SetColumnSpan(btnCheckClash, 3);

            btnPrevClash = CreateButton("<< Prev"); btnPrevClash.Dock = DockStyle.Fill; btnPrevClash.Enabled = false;
            btnPrevClash.Click += BtnPrevClash_Click;
            tbl.Controls.Add(btnPrevClash, 0, 1);

            btnNextClash = CreateButton("Next >>"); btnNextClash.Dock = DockStyle.Fill; btnNextClash.Enabled = false;
            btnNextClash.Click += BtnNextClash_Click;
            tbl.Controls.Add(btnNextClash, 1, 1);

            btnEndClash = CreateButton("X Kết Thúc"); btnEndClash.Dock = DockStyle.Fill; btnEndClash.Enabled = false;
            btnEndClash.Click += BtnEndClash_Click;
            tbl.Controls.Add(btnEndClash, 2, 1);

            btnPackAndGo = CreateButton("Pack and Go", true); btnPackAndGo.Dock = DockStyle.Fill;
            tbl.Controls.Add(btnPackAndGo, 0, 2);
            tbl.SetColumnSpan(btnPackAndGo, 3);

            gb.Controls.Add(tbl);
        }
        #endregion

        #region Tab Builders
        private void BuildExportTab(TabPage page)
        {
            page.AutoScroll = true;
            var flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true, Padding = new Padding(MARGIN) };

            // Naming Rule
            GroupBox gbName = new GroupBox() { Text = "Tên file", Width = CONTENT_WIDTH, Height = 95 };
            radNameDrawingNo = new RadioButton() { Text = "Drawing NO (Chi tiết lẻ)", Location = new Point(10, 20), Width = 250, Checked = true };
            radNameUnitName = new RadioButton() { Text = "Unit Name (Cụm con)", Location = new Point(10, 45), Width = 250 };
            radNameMachineName = new RadioButton() { Text = "Machine Name (Cụm tổng)", Location = new Point(10, 70), Width = 250 };
            gbName.Controls.AddRange(new Control[] { radNameDrawingNo, radNameUnitName, radNameMachineName });
            flow.Controls.Add(gbName);

            // Format Selection
            GroupBox gbFormat = new GroupBox() { Text = "Định dạng file", Width = CONTENT_WIDTH, Height = 100 };
            chkPdf2D = new CheckBox() { Text = "PDF 2D", Location = new Point(10, 20), Width = 100, Checked = true };
            chkPdf3D = new CheckBox() { Text = "PDF 3D (Part/Asm)", Location = new Point(120, 20), Width = 150 };
            chkDxf = new CheckBox() { Text = "DXF", Location = new Point(10, 45), Width = 60 };
            chkDwg = new CheckBox() { Text = "DWG", Location = new Point(70, 45), Width = 60 };
            cbDxfVersion = new ComboBox() { Location = new Point(140, 43), Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            cbDxfVersion.Items.AddRange(new object[] { "2013", "2018" }); cbDxfVersion.SelectedIndex = 0;
            chkStp = new CheckBox() { Text = "STEP", Location = new Point(10, 70), Width = 80 };
            chkXt = new CheckBox() { Text = "X_T", Location = new Point(120, 70), Width = 80 };
            gbFormat.Controls.AddRange(new Control[] { chkPdf2D, chkPdf3D, chkDxf, chkDwg, cbDxfVersion, chkStp, chkXt });
            flow.Controls.Add(gbFormat);

            // Sheet Options
            GroupBox gbSheet = new GroupBox() { Text = "Tuỳ Chọn Sheet (Bản Vẽ)", Width = CONTENT_WIDTH, Height = 120 };
            radPdfSheetActive = new RadioButton() { Text = "Sheet đang mở", Location = new Point(10, 20), Width = 250, Checked = true };
            radPdfSheetAll = new RadioButton() { Text = "Gộp tất cả sheet vào 1 file", Location = new Point(10, 45), Width = 250 };
            radPdfSheetSeparate = new RadioButton() { Text = "Tách mỗi sheet 1 file", Location = new Point(10, 70), Width = 250 };
            radPdfSheetRange = new RadioButton() { Text = "Dải sheet:", Location = new Point(10, 95), Width = 80 };
            txtPdfSheetRange = new TextBox() { Location = new Point(95, 93), Width = 150 };
            gbSheet.Controls.AddRange(new Control[] { radPdfSheetActive, radPdfSheetAll, radPdfSheetSeparate, radPdfSheetRange, txtPdfSheetRange });
            flow.Controls.Add(gbSheet);

            btnDoExport = CreateButton("THỰC HIỆN XUẤT FILE", true); btnDoExport.Width = CONTENT_WIDTH; btnDoExport.Height = 40;
            flow.Controls.Add(btnDoExport);

            GroupBox gbBom = new GroupBox() { Text = "Xuất BOM List Excel", Width = CONTENT_WIDTH, Height = 60 };
            btnExportBomList = CreateButton("Xuất BOM (Tiêu chuẩn & Gia công)"); btnExportBomList.Dock = DockStyle.Fill;
            gbBom.Controls.Add(btnExportBomList);
            flow.Controls.Add(gbBom);

            page.Controls.Add(flow);
        }

        private void BuildPropsTab(TabPage page)
        {
            page.AutoScroll = true;
            var flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true, Padding = new Padding(MARGIN) };

            btnSetupProperties = CreateButton("Cài Đặt System Properties Mặc Định", true); btnSetupProperties.Width = CONTENT_WIDTH;
            flow.Controls.Add(btnSetupProperties);

            GroupBox gbInputs = new GroupBox() { Text = "Nhập Thuộc Tính Nhanh", Width = CONTENT_WIDTH, Height = 370 };
            int y = 20;
            txtPropTitle = AddFieldRow(gbInputs, "Mã BV (Title):", ref y);
            txtPropPartName = AddFieldRow(gbInputs, "Tên CT (Part Name):", ref y);
            txtPropUnit = AddFieldRow(gbInputs, "Cụm (Unit):", ref y);
            txtPropNameProject = AddFieldRow(gbInputs, "Máy (Project):", ref y);

            gbInputs.Controls.Add(CreateLabel("Thiết kế:")); gbInputs.Controls[gbInputs.Controls.Count - 1].Location = new Point(10, y);
            cbPropDesign = new ComboBox() { Location = new Point(105, y - 3), Width = 165 };
            cbPropDesign.Items.AddRange(new object[] { "VISC", "R&D" });
            gbInputs.Controls.Add(cbPropDesign); y += FIELD_HEIGHT;

            txtPropMaterial = AddFieldRow(gbInputs, "Vật liệu:", ref y);
            txtPropSurface = AddFieldRow(gbInputs, "Xử lý bề mặt:", ref y);
            txtPropHeat = AddFieldRow(gbInputs, "Nhiệt luyện:", ref y);
            txtPropManufacturer = AddFieldRow(gbInputs, "Hãng SX:", ref y);

            gbInputs.Controls.Add(CreateLabel("Phân loại:")); gbInputs.Controls[gbInputs.Controls.Count - 1].Location = new Point(10, y);
            cbPropMakeOrBuy = new ComboBox() { Location = new Point(105, y - 3), Width = 165, DropDownStyle = ComboBoxStyle.DropDownList };
            cbPropMakeOrBuy.Items.AddRange(new object[] { "MF (Gia công)", "PUR (Tiêu chuẩn/Mua ngoài)" });
            cbPropMakeOrBuy.SelectedIndex = 0;
            gbInputs.Controls.Add(cbPropMakeOrBuy); y += FIELD_HEIGHT + 5;

            btnApplyProps = CreateButton("LƯU THUỘC TÍNH VÀO FILE", true);
            btnApplyProps.Location = new Point(10, y); btnApplyProps.Width = 260; btnApplyProps.Height = 40; btnApplyProps.Dock = DockStyle.None;
            btnApplyProps.Click += BtnApplyProps_Click;
            gbInputs.Controls.Add(btnApplyProps);

            flow.Controls.Add(gbInputs);
            page.Controls.Add(flow);
        }

        private TextBox AddFieldRow(GroupBox parent, string label, ref int y)
        {
            parent.Controls.Add(new Label() { Text = label, Location = new Point(10, y), Width = 95, Height = 20, Font = UIFont });
            var txt = new TextBox() { Location = new Point(105, y - 3), Width = 165 };
            parent.Controls.Add(txt);
            y += FIELD_HEIGHT;
            return txt;
        }

        private void BuildRenameTab(TabPage page)
        {
            btnBatchRename = CreateButton("Đổi Tên Chi Tiết Hàng Loạt", true);
            btnBatchRename.Dock = DockStyle.Top; btnBatchRename.Height = 50;
            page.Controls.Add(btnBatchRename);

            var lbl = new Label() { Text = "(Tính năng Đổi tên hàng loạt sẽ được phát triển tiếp)", Dock = DockStyle.Top, Height = 50, Padding = new Padding(10) };
            page.Controls.Add(lbl);
        }

        private void BuildDrawingTab(TabPage page)
        {
            btnAutoDrawing = CreateButton("Tạo Bản Vẽ Tự Động", true);
            btnAutoDrawing.Dock = DockStyle.Top; btnAutoDrawing.Height = 50;
            page.Controls.Add(btnAutoDrawing);

            var lbl = new Label() { Text = "- Part: 1 Sheet\n- Assembly: Multi-sheet (BOM + Tổng thể + Các chi tiết lẻ)", Dock = DockStyle.Top, Height = 80, Padding = new Padding(10) };
            page.Controls.Add(lbl);
        }

        private void BuildAboutTab(TabPage page)
        {
            var lblTitle = new Label() { Text = "VISC SolidWorks Add-in", Dock = DockStyle.Top, Height = 40, TextAlign = ContentAlignment.MiddleCenter, Font = UIFontLarge };
            var lblDesc = new Label() { Text = "Phiên bản: 1.0.0\nPhát triển bởi đội ngũ R&D", Dock = DockStyle.Top, Height = 50, TextAlign = ContentAlignment.TopCenter };
            page.Controls.Add(lblDesc);
            page.Controls.Add(lblTitle);
        }
        #endregion

        #region Event Handlers
        private void BtnCheckClash_Click(object sender, EventArgs e)
        {
            if (_interferenceMgr == null || _swApp == null) return;
            ModelDoc2 m = (ModelDoc2)_swApp.ActiveDoc;
            if (m == null || m.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                SetStatus("Chỉ dùng được trong Assembly!");
                return;
            }
            try
            {
                SetStatus("Đang kiểm tra va chạm...");
                int count = _interferenceMgr.RunInterferenceDetectionAndHighlight();
                if (count > 0)
                {
                    btnCheckClash.Text = string.Format("Kiểm Tra Va Chạm ({0} lỗi)", count);
                    btnNextClash.Enabled = true; btnPrevClash.Enabled = true; btnEndClash.Enabled = true;
                    _interferenceMgr.ZoomToNextClash();
                    SetStatus(string.Format("Phát hiện {0} va chạm", count));
                }
                else
                {
                    btnCheckClash.Text = "Kiểm Tra Va Chạm (0 lỗi)";
                    btnNextClash.Enabled = false; btnPrevClash.Enabled = false; btnEndClash.Enabled = true;
                    SetStatus("Không phát hiện va chạm");
                }
            }
            catch (Exception ex) { SetStatus("Lỗi: " + ex.Message); }
        }

        private void BtnNextClash_Click(object sender, EventArgs e) { if (_interferenceMgr != null) _interferenceMgr.ZoomToNextClash(); }
        private void BtnPrevClash_Click(object sender, EventArgs e) { if (_interferenceMgr != null) _interferenceMgr.ZoomToPrevClash(); }

        private void BtnEndClash_Click(object sender, EventArgs e)
        {
            if (_interferenceMgr != null) _interferenceMgr.EndClashDetection();
            btnCheckClash.Text = "Kiểm Tra Va Chạm";
            btnNextClash.Enabled = false; btnPrevClash.Enabled = false; btnEndClash.Enabled = false;
            SetStatus("Kết thúc kiểm tra va chạm");
        }

        private void BtnApplyProps_Click(object sender, EventArgs e)
        {
            if (_swApp == null) return;
            ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
            if (swModel == null)
            {
                SetStatus("Vui lòng mở file trước!");
                return;
            }
            try
            {
                CustomPropertyManager pm = swModel.Extension.get_CustomPropertyManager("");
                SetProp(pm, "Title", txtPropTitle.Text);
                SetProp(pm, "Part Name", txtPropPartName.Text);
                SetProp(pm, "Unit", txtPropUnit.Text);
                SetProp(pm, "Name Project", txtPropNameProject.Text);
                SetProp(pm, "Design", cbPropDesign.Text);
                SetProp(pm, "Material", txtPropMaterial.Text);
                SetProp(pm, "Surface Treatment", txtPropSurface.Text);
                SetProp(pm, "Heat Treatment", txtPropHeat.Text);
                SetProp(pm, "Manufacturer", txtPropManufacturer.Text);
                SetProp(pm, "MakeOrBuy", cbPropMakeOrBuy.SelectedIndex == 0 ? "MF" : "PUR");
                SetStatus("Lưu thuộc tính thành công!");
            }
            catch (Exception ex) { SetStatus("Lỗi: " + ex.Message); }
        }

        private void SetProp(CustomPropertyManager pm, string name, string val)
        {
            if (!string.IsNullOrEmpty(val))
                pm.Add3(name, (int)swCustomInfoType_e.swCustomInfoText, val, (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
        }
        #endregion
    }
}
