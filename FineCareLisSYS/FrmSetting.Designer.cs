namespace FineCareLisSYS
{
    partial class FrmSetting
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmSetting));
            this.lstV1 = new System.Windows.Forms.ListView();
            this.columnHeader0 = new System.Windows.Forms.ColumnHeader();
            this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
            this.columnHeader2 = new System.Windows.Forms.ColumnHeader();
            this.columnHeader3 = new System.Windows.Forms.ColumnHeader();
            this.columnHeader4 = new System.Windows.Forms.ColumnHeader();
            this.columnHeader5 = new System.Windows.Forms.ColumnHeader();
            this.label1 = new System.Windows.Forms.Label();
            this.tb_Name = new System.Windows.Forms.TextBox();
            this.tb_Unit = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tb_Reference = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btn_Save = new System.Windows.Forms.Button();
            this.chk_reference = new System.Windows.Forms.CheckBox();
            this.btn_ExitSetting = new System.Windows.Forms.Button();
            this.tb_PrintName = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btn_AddItem = new System.Windows.Forms.Button();
            this.tb_AddItem = new System.Windows.Forms.TextBox();
            this.tb_Description = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lstV1
            // 
            this.lstV1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader0,
            this.columnHeader1,
            this.columnHeader2,
            this.columnHeader3,
            this.columnHeader4,
            this.columnHeader5});
            this.lstV1.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lstV1.FullRowSelect = true;
            this.lstV1.GridLines = true;
            this.lstV1.Location = new System.Drawing.Point(9, 11);
            this.lstV1.MultiSelect = false;
            this.lstV1.Name = "lstV1";
            this.lstV1.Size = new System.Drawing.Size(593, 404);
            this.lstV1.TabIndex = 0;
            this.lstV1.UseCompatibleStateImageBehavior = false;
            this.lstV1.View = System.Windows.Forms.View.Details;
            this.lstV1.SelectedIndexChanged += new System.EventHandler(this.lstV1_SelectedIndexChanged);
            // 
            // columnHeader0
            // 
            this.columnHeader0.Text = "ID";
            this.columnHeader0.Width = 30;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Item Name";
            this.columnHeader1.Width = 80;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Unit";
            this.columnHeader2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.columnHeader2.Width = 80;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "Reference Range";
            this.columnHeader3.Width = 150;
            // 
            // columnHeader4
            // 
            this.columnHeader4.Text = "Display Name";
            this.columnHeader4.Width = 100;
            // 
            // columnHeader5
            // 
            this.columnHeader5.Text = "Note";
            this.columnHeader5.Width = 132;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label1.Location = new System.Drawing.Point(616, 49);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(44, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "Item Name";
            // 
            // tb_Name
            // 
            this.tb_Name.Enabled = false;
            this.tb_Name.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.tb_Name.Location = new System.Drawing.Point(679, 47);
            this.tb_Name.Name = "tb_Name";
            this.tb_Name.Size = new System.Drawing.Size(126, 23);
            this.tb_Name.TabIndex = 2;
            // 
            // tb_Unit
            // 
            this.tb_Unit.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.tb_Unit.Location = new System.Drawing.Point(679, 83);
            this.tb_Unit.Name = "tb_Unit";
            this.tb_Unit.Size = new System.Drawing.Size(126, 23);
            this.tb_Unit.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label2.Location = new System.Drawing.Point(616, 85);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(32, 17);
            this.label2.TabIndex = 3;
            this.label2.Text = "Unit";
            // 
            // tb_Reference
            // 
            this.tb_Reference.Enabled = false;
            this.tb_Reference.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.tb_Reference.Location = new System.Drawing.Point(619, 194);
            this.tb_Reference.Multiline = true;
            this.tb_Reference.Name = "tb_Reference";
            this.tb_Reference.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tb_Reference.Size = new System.Drawing.Size(186, 80);
            this.tb_Reference.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label3.Location = new System.Drawing.Point(614, 175);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 17);
            this.label3.TabIndex = 5;
            this.label3.Text = "Reference Range";
            // 
            // btn_Save
            // 
            this.btn_Save.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.btn_Save.Location = new System.Drawing.Point(618, 389);
            this.btn_Save.Name = "btn_Save";
            this.btn_Save.Size = new System.Drawing.Size(72, 27);
            this.btn_Save.TabIndex = 7;
            this.btn_Save.Text = "Save";
            this.btn_Save.UseVisualStyleBackColor = true;
            this.btn_Save.Click += new System.EventHandler(this.btn_Save_Click);
            // 
            // chk_reference
            // 
            this.chk_reference.AutoSize = true;
            this.chk_reference.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chk_reference.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.chk_reference.Location = new System.Drawing.Point(614, 148);
            this.chk_reference.Name = "chk_reference";
            this.chk_reference.Size = new System.Drawing.Size(99, 21);
            this.chk_reference.TabIndex = 8;
            this.chk_reference.Text = "Edit reference range";
            this.chk_reference.UseVisualStyleBackColor = true;
            this.chk_reference.CheckedChanged += new System.EventHandler(this.chk_reference_CheckedChanged);
            // 
            // btn_ExitSetting
            // 
            this.btn_ExitSetting.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.btn_ExitSetting.Location = new System.Drawing.Point(733, 389);
            this.btn_ExitSetting.Name = "btn_ExitSetting";
            this.btn_ExitSetting.Size = new System.Drawing.Size(72, 27);
            this.btn_ExitSetting.TabIndex = 9;
            this.btn_ExitSetting.Text = "Cancel";
            this.btn_ExitSetting.UseVisualStyleBackColor = true;
            this.btn_ExitSetting.Click += new System.EventHandler(this.btn_ExitSetting_Click);
            // 
            // tb_PrintName
            // 
            this.tb_PrintName.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.tb_PrintName.Location = new System.Drawing.Point(679, 118);
            this.tb_PrintName.Name = "tb_PrintName";
            this.tb_PrintName.Size = new System.Drawing.Size(126, 23);
            this.tb_PrintName.TabIndex = 11;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label4.Location = new System.Drawing.Point(616, 120);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 17);
            this.label4.TabIndex = 10;
            this.label4.Text = "Display Name";
            // 
            // btn_AddItem
            // 
            this.btn_AddItem.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_AddItem.Location = new System.Drawing.Point(608, 11);
            this.btn_AddItem.Name = "btn_AddItem";
            this.btn_AddItem.Size = new System.Drawing.Size(71, 25);
            this.btn_AddItem.TabIndex = 12;
            this.btn_AddItem.Text = "Add a new item";
            this.btn_AddItem.UseVisualStyleBackColor = true;
            this.btn_AddItem.Click += new System.EventHandler(this.btn_AddItem_Click);
            // 
            // tb_AddItem
            // 
            this.tb_AddItem.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.tb_AddItem.Location = new System.Drawing.Point(679, 12);
            this.tb_AddItem.Name = "tb_AddItem";
            this.tb_AddItem.Size = new System.Drawing.Size(126, 23);
            this.tb_AddItem.TabIndex = 13;
            // 
            // tb_Description
            // 
            this.tb_Description.Enabled = false;
            this.tb_Description.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.tb_Description.Location = new System.Drawing.Point(619, 300);
            this.tb_Description.Multiline = true;
            this.tb_Description.Name = "tb_Description";
            this.tb_Description.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tb_Description.Size = new System.Drawing.Size(186, 80);
            this.tb_Description.TabIndex = 14;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label5.Location = new System.Drawing.Point(614, 280);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(32, 17);
            this.label5.TabIndex = 15;
            this.label5.Text = "Note";
            // 
            // FrmSetting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(812, 428);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.tb_Description);
            this.Controls.Add(this.tb_AddItem);
            this.Controls.Add(this.btn_AddItem);
            this.Controls.Add(this.tb_PrintName);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btn_ExitSetting);
            this.Controls.Add(this.chk_reference);
            this.Controls.Add(this.btn_Save);
            this.Controls.Add(this.tb_Reference);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.tb_Unit);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tb_Name);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lstV1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "FrmSetting";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Parameter Setting";
            this.Load += new System.EventHandler(this.FrmSetting_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView lstV1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tb_Name;
        private System.Windows.Forms.TextBox tb_Unit;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tb_Reference;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btn_Save;
        private System.Windows.Forms.CheckBox chk_reference;
        private System.Windows.Forms.Button btn_ExitSetting;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.ColumnHeader columnHeader0;
        private System.Windows.Forms.TextBox tb_PrintName;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private System.Windows.Forms.Button btn_AddItem;
        private System.Windows.Forms.TextBox tb_AddItem;
        private System.Windows.Forms.TextBox tb_Description;
        private System.Windows.Forms.ColumnHeader columnHeader5;
        private System.Windows.Forms.Label label5;
    }
}