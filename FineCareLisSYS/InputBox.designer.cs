namespace FineCareLisSYS
{
    partial class mInputBox
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_OK = new System.Windows.Forms.Button();
            this.tb_Input = new System.Windows.Forms.TextBox();
            this.btn_Cancle = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn_OK
            // 
            this.btn_OK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btn_OK.Location = new System.Drawing.Point(36, 72);
            this.btn_OK.Name = "btn_OK";
            this.btn_OK.Size = new System.Drawing.Size(68, 24);
            this.btn_OK.TabIndex = 0;
            this.btn_OK.Text = "Ok";
            this.btn_OK.UseVisualStyleBackColor = true;
            this.btn_OK.Click += new System.EventHandler(this.btn_OK_Click);
            // 
            // tb_Input
            // 
            this.tb_Input.Location = new System.Drawing.Point(36, 27);
            this.tb_Input.Name = "tb_Input";
            this.tb_Input.Size = new System.Drawing.Size(153, 21);
            this.tb_Input.TabIndex = 2;
            // 
            // btn_Cancle
            // 
            this.btn_Cancle.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btn_Cancle.Location = new System.Drawing.Point(121, 72);
            this.btn_Cancle.Name = "btn_Cancle";
            this.btn_Cancle.Size = new System.Drawing.Size(68, 24);
            this.btn_Cancle.TabIndex = 3;
            this.btn_Cancle.Text = "Cancel";
            this.btn_Cancle.UseVisualStyleBackColor = true;
            // 
            // mInputBox
            // 
            this.AcceptButton = this.btn_OK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btn_Cancle;
            this.ClientSize = new System.Drawing.Size(226, 119);
            this.ControlBox = false;
            this.Controls.Add(this.btn_Cancle);
            this.Controls.Add(this.tb_Input);
            this.Controls.Add(this.btn_OK);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "mInputBox";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Please enter the name of the item to be added：";
            this.Load += new System.EventHandler(this.mInputBox_Load);
            this.Activated += new System.EventHandler(this.mInputBox_Activated);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_OK;
        private System.Windows.Forms.Button btn_Cancle;
        public System.Windows.Forms.TextBox tb_Input;
    }
}