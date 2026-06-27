using System;
using System.Windows.Forms;

namespace SwAutomationAddin
{
    public class BatchRenameForm : Form
    {
        private TextBox txtNewName;
        private Button btnOK, btnCancel;

        public string NewName { get; private set; }

        public BatchRenameForm()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Đổi Tên Chi Tiết";
            this.Width = 350;
            this.Height = 150;
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Label
            var lbl = new Label()
            {
                Text = "Nhập tên mới:",
                Location = new System.Drawing.Point(10, 15),
                Width = 100
            };
            this.Controls.Add(lbl);

            // TextBox
            txtNewName = new TextBox()
            {
                Location = new System.Drawing.Point(10, 40),
                Width = 310,
                Height = 25
            };
            this.Controls.Add(txtNewName);

            // OK Button
            btnOK = new Button()
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Location = new System.Drawing.Point(165, 80),
                Width = 70
            };
            this.Controls.Add(btnOK);

            // Cancel Button
            btnCancel = new Button()
            {
                Text = "Hủy",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(250, 80),
                Width = 70
            };
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (this.DialogResult == DialogResult.OK)
            {
                NewName = txtNewName.Text;
            }
            base.OnFormClosing(e);
        }
    }
}