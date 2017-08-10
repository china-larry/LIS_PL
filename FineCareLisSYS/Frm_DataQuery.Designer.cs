namespace FineCareLisSYS
{
    partial class Frm_DataQuery
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Frm_DataQuery));
            this.dgv_Result = new System.Windows.Forms.DataGridView();
            this.col_xlh = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_Swatch = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_CheckTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_Name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_Sex = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_Age = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_ProName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_Result = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_Doc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_RecvTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label11 = new System.Windows.Forms.Label();
            this.DTPickerS = new System.Windows.Forms.DateTimePicker();
            this.DTPickerE = new System.Windows.Forms.DateTimePicker();
            this.cmb_Doctor = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.btn_Query = new System.Windows.Forms.Button();
            this.btn_Export = new System.Windows.Forms.Button();
            this.cmb_QyType = new System.Windows.Forms.ComboBox();
            this.sfDialog = new System.Windows.Forms.SaveFileDialog();
            this.label2 = new System.Windows.Forms.Label();
            this.tb_Name = new System.Windows.Forms.TextBox();
            this.tb_Swatch = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cmb_ItemName = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Result)).BeginInit();
            this.SuspendLayout();
            // 
            // dgv_Result
            // 
            this.dgv_Result.AllowUserToAddRows = false;
            this.dgv_Result.AllowUserToDeleteRows = false;
            this.dgv_Result.BackgroundColor = System.Drawing.Color.Honeydew;
            this.dgv_Result.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("微软雅黑", 9F);
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv_Result.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv_Result.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_Result.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.col_xlh,
            this.col_Swatch,
            this.col_CheckTime,
            this.col_Name,
            this.col_Sex,
            this.col_Age,
            this.col_ProName,
            this.col_Result,
            this.col_Doc,
            this.col_RecvTime});
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("微软雅黑", 9F);
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv_Result.DefaultCellStyle = dataGridViewCellStyle3;
            this.dgv_Result.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgv_Result.Location = new System.Drawing.Point(-1, 0);
            this.dgv_Result.Name = "dgv_Result";
            this.dgv_Result.ReadOnly = true;
            this.dgv_Result.RowHeadersVisible = false;
            this.dgv_Result.RowTemplate.Height = 23;
            this.dgv_Result.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv_Result.Size = new System.Drawing.Size(833, 408);
            this.dgv_Result.TabIndex = 1;
            // 
            // col_xlh
            // 
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.LightGray;
            this.col_xlh.DefaultCellStyle = dataGridViewCellStyle2;
            this.col_xlh.Frozen = true;
            this.col_xlh.HeaderText = "No.";
            this.col_xlh.Name = "col_xlh";
            this.col_xlh.ReadOnly = true;
            this.col_xlh.Width = 40;
            // 
            // col_Swatch
            // 
            this.col_Swatch.HeaderText = "Test NO.";
            this.col_Swatch.Name = "col_Swatch";
            this.col_Swatch.ReadOnly = true;
            this.col_Swatch.Width = 70;
            // 
            // col_CheckTime
            // 
            this.col_CheckTime.HeaderText = "Test Time";
            this.col_CheckTime.Name = "col_CheckTime";
            this.col_CheckTime.ReadOnly = true;
            this.col_CheckTime.Width = 130;
            // 
            // col_Name
            // 
            this.col_Name.HeaderText = "Name";
            this.col_Name.Name = "col_Name";
            this.col_Name.ReadOnly = true;
            // 
            // col_Sex
            // 
            this.col_Sex.HeaderText = "Gender";
            this.col_Sex.Name = "col_Sex";
            this.col_Sex.ReadOnly = true;
            this.col_Sex.Width = 60;
            // 
            // col_Age
            // 
            this.col_Age.HeaderText = "Age";
            this.col_Age.Name = "col_Age";
            this.col_Age.ReadOnly = true;
            this.col_Age.Width = 60;
            // 
            // col_ProName
            // 
            this.col_ProName.HeaderText = "Item Name";
            this.col_ProName.Name = "col_ProName";
            this.col_ProName.ReadOnly = true;
            // 
            // col_Result
            // 
            this.col_Result.HeaderText = "Results";
            this.col_Result.Name = "col_Result";
            this.col_Result.ReadOnly = true;
            this.col_Result.Width = 80;
            // 
            // col_Doc
            // 
            this.col_Doc.HeaderText = "Doctor";
            this.col_Doc.Name = "col_Doc";
            this.col_Doc.ReadOnly = true;
            this.col_Doc.Width = 60;
            // 
            // col_RecvTime
            // 
            this.col_RecvTime.HeaderText = "Receive Time";
            this.col_RecvTime.Name = "col_RecvTime";
            this.col_RecvTime.ReadOnly = true;
            this.col_RecvTime.Width = 130;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label11.Location = new System.Drawing.Point(102, 451);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(56, 17);
            this.label11.TabIndex = 39;
            this.label11.Text = "From";
            // 
            // DTPickerS
            // 
            this.DTPickerS.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.DTPickerS.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DTPickerS.Location = new System.Drawing.Point(162, 448);
            this.DTPickerS.Name = "DTPickerS";
            this.DTPickerS.Size = new System.Drawing.Size(100, 23);
            this.DTPickerS.TabIndex = 38;
            // 
            // DTPickerE
            // 
            this.DTPickerE.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.DTPickerE.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DTPickerE.Location = new System.Drawing.Point(291, 449);
            this.DTPickerE.Name = "DTPickerE";
            this.DTPickerE.Size = new System.Drawing.Size(100, 23);
            this.DTPickerE.TabIndex = 40;
            // 
            // cmb_Doctor
            // 
            this.cmb_Doctor.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.cmb_Doctor.FormattingEnabled = true;
            this.cmb_Doctor.Location = new System.Drawing.Point(625, 450);
            this.cmb_Doctor.Name = "cmb_Doctor";
            this.cmb_Doctor.Size = new System.Drawing.Size(108, 25);
            this.cmb_Doctor.TabIndex = 43;
            this.cmb_Doctor.Tag = "ys";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label8.Location = new System.Drawing.Point(578, 454);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(32, 17);
            this.label8.TabIndex = 42;
            this.label8.Text = "Doctor";
            // 
            // btn_Query
            // 
            this.btn_Query.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.btn_Query.Location = new System.Drawing.Point(745, 417);
            this.btn_Query.Name = "btn_Query";
            this.btn_Query.Size = new System.Drawing.Size(87, 26);
            this.btn_Query.TabIndex = 44;
            this.btn_Query.Text = "Query";
            this.btn_Query.UseVisualStyleBackColor = true;
            this.btn_Query.Click += new System.EventHandler(this.btn_Query_Click);
            // 
            // btn_Export
            // 
            this.btn_Export.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.btn_Export.Location = new System.Drawing.Point(745, 449);
            this.btn_Export.Name = "btn_Export";
            this.btn_Export.Size = new System.Drawing.Size(87, 26);
            this.btn_Export.TabIndex = 45;
            this.btn_Export.Text = "Export to Excel";
            this.btn_Export.UseVisualStyleBackColor = true;
            this.btn_Export.Click += new System.EventHandler(this.btn_Export_Click);
            // 
            // cmb_QyType
            // 
            this.cmb_QyType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmb_QyType.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.cmb_QyType.FormattingEnabled = true;
            this.cmb_QyType.Items.AddRange(new object[] {
            "Test Date",
            "Receive Date"});
            this.cmb_QyType.Location = new System.Drawing.Point(234, 417);
            this.cmb_QyType.Name = "cmb_QyType";
            this.cmb_QyType.Size = new System.Drawing.Size(157, 25);
            this.cmb_QyType.TabIndex = 47;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label2.Location = new System.Drawing.Point(402, 421);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(32, 17);
            this.label2.TabIndex = 48;
            this.label2.Text = "Name";
            // 
            // tb_Name
            // 
            this.tb_Name.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.tb_Name.Location = new System.Drawing.Point(459, 418);
            this.tb_Name.Name = "tb_Name";
            this.tb_Name.Size = new System.Drawing.Size(108, 23);
            this.tb_Name.TabIndex = 49;
            // 
            // tb_Swatch
            // 
            this.tb_Swatch.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.tb_Swatch.Location = new System.Drawing.Point(625, 418);
            this.tb_Swatch.Name = "tb_Swatch";
            this.tb_Swatch.Size = new System.Drawing.Size(108, 23);
            this.tb_Swatch.TabIndex = 51;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label3.Location = new System.Drawing.Point(578, 421);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(44, 17);
            this.label3.TabIndex = 50;
            this.label3.Text = "Sample ID";
            // 
            // cmb_ItemName
            // 
            this.cmb_ItemName.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.cmb_ItemName.FormattingEnabled = true;
            this.cmb_ItemName.Items.AddRange(new object[] {
            "Man",
            "Woman"});
            this.cmb_ItemName.Location = new System.Drawing.Point(461, 450);
            this.cmb_ItemName.Name = "cmb_ItemName";
            this.cmb_ItemName.Size = new System.Drawing.Size(108, 25);
            this.cmb_ItemName.TabIndex = 53;
            this.cmb_ItemName.Tag = "ys";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label4.Location = new System.Drawing.Point(402, 454);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 17);
            this.label4.TabIndex = 52;
            this.label4.Text = "Item Name";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label1.Location = new System.Drawing.Point(268, 451);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(20, 17);
            this.label1.TabIndex = 41;
            this.label1.Text = "To";
            // 
            // Frm_DataQuery
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(834, 482);
            this.Controls.Add(this.cmb_ItemName);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.tb_Swatch);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.tb_Name);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cmb_QyType);
            this.Controls.Add(this.btn_Export);
            this.Controls.Add(this.btn_Query);
            this.Controls.Add(this.cmb_Doctor);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.DTPickerE);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.DTPickerS);
            this.Controls.Add(this.dgv_Result);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Frm_DataQuery";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Advanced Query";
            this.Load += new System.EventHandler(this.Frm_DataQuery_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Result)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgv_Result;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_xlh;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_Swatch;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_CheckTime;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_Name;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_Sex;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_Age;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_ProName;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_Result;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_Doc;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_RecvTime;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.DateTimePicker DTPickerS;
        private System.Windows.Forms.DateTimePicker DTPickerE;
        private System.Windows.Forms.ComboBox cmb_Doctor;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btn_Query;
        private System.Windows.Forms.Button btn_Export;
        private System.Windows.Forms.ComboBox cmb_QyType;
        private System.Windows.Forms.SaveFileDialog sfDialog;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tb_Name;
        private System.Windows.Forms.TextBox tb_Swatch;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmb_ItemName;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label1;


    }
}