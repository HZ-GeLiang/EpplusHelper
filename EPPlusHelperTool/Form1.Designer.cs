namespace EPPlusHelperTool
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.filePath1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SelectFileBtn1 = new System.Windows.Forms.Button();
            this.GenerateConfiguration = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.filePath2 = new System.Windows.Forms.TextBox();
            this.SelectFileBtn2 = new System.Windows.Forms.Button();
            this.CheckTemplateConfiguration = new System.Windows.Forms.Button();
            this.wsNameOrIndex1 = new System.Windows.Forms.TextBox();
            this.wsNameOrIndex2 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.TitleLine1 = new System.Windows.Forms.TextBox();
            this.TitleLine2 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.GenerateConfigurationCode = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.DelHiddenWs = new System.Windows.Forms.Button();
            this.WScount1 = new System.Windows.Forms.Button();
            this.WScount2 = new System.Windows.Forms.Button();
            this.dataGridViewExcel2 = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewExcel1 = new System.Windows.Forms.DataGridView();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewExcel2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewExcel1)).BeginInit();
            this.SuspendLayout();
            // 
            // filePath1
            // 
            this.filePath1.AllowDrop = true;
            this.filePath1.Location = new System.Drawing.Point(22, 23);
            this.filePath1.Multiline = true;
            this.filePath1.Name = "filePath1";
            this.filePath1.Size = new System.Drawing.Size(366, 66);
            this.filePath1.TabIndex = 4;
            this.filePath1.DragDrop += new System.Windows.Forms.DragEventHandler(this.textBoxDragDrop);
            this.filePath1.DragEnter += new System.Windows.Forms.DragEventHandler(this.textBoxDragEnter);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(149, 12);
            this.label1.TabIndex = 5;
            this.label1.Text = "文件路径(文本框支持拖拽)";
            // 
            // SelectFileBtn1
            // 
            this.SelectFileBtn1.Location = new System.Drawing.Point(394, 23);
            this.SelectFileBtn1.Name = "SelectFileBtn1";
            this.SelectFileBtn1.Size = new System.Drawing.Size(74, 37);
            this.SelectFileBtn1.TabIndex = 6;
            this.SelectFileBtn1.Text = "选择...";
            this.SelectFileBtn1.UseVisualStyleBackColor = true;
            this.SelectFileBtn1.Click += new System.EventHandler(this.btn_SelectExcelFile);
            // 
            // GenerateConfiguration
            // 
            this.GenerateConfiguration.Location = new System.Drawing.Point(125, 19);
            this.GenerateConfiguration.Name = "GenerateConfiguration";
            this.GenerateConfiguration.Size = new System.Drawing.Size(113, 23);
            this.GenerateConfiguration.TabIndex = 7;
            this.GenerateConfiguration.Text = "给Excel填充配置";
            this.GenerateConfiguration.UseVisualStyleBackColor = true;
            this.GenerateConfiguration.Click += new System.EventHandler(this.GenerateConfiguration_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 295);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(113, 12);
            this.label2.TabIndex = 8;
            this.label2.Text = "自动初始化填充配置";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(525, 7);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(149, 12);
            this.label3.TabIndex = 9;
            this.label3.Text = "文件路径(文本框支持拖拽)";
            // 
            // filePath2
            // 
            this.filePath2.AllowDrop = true;
            this.filePath2.Location = new System.Drawing.Point(527, 23);
            this.filePath2.Multiline = true;
            this.filePath2.Name = "filePath2";
            this.filePath2.Size = new System.Drawing.Size(371, 66);
            this.filePath2.TabIndex = 10;
            this.filePath2.DragDrop += new System.Windows.Forms.DragEventHandler(this.textBoxDragDrop);
            this.filePath2.DragEnter += new System.Windows.Forms.DragEventHandler(this.textBoxDragEnter);
            // 
            // SelectFileBtn2
            // 
            this.SelectFileBtn2.Location = new System.Drawing.Point(904, 23);
            this.SelectFileBtn2.Name = "SelectFileBtn2";
            this.SelectFileBtn2.Size = new System.Drawing.Size(74, 37);
            this.SelectFileBtn2.TabIndex = 11;
            this.SelectFileBtn2.Text = "选择...";
            this.SelectFileBtn2.UseVisualStyleBackColor = true;
            this.SelectFileBtn2.Click += new System.EventHandler(this.btn_SelectExcelFile);
            // 
            // CheckTemplateConfiguration
            // 
            this.CheckTemplateConfiguration.Location = new System.Drawing.Point(536, 322);
            this.CheckTemplateConfiguration.Name = "CheckTemplateConfiguration";
            this.CheckTemplateConfiguration.Size = new System.Drawing.Size(90, 23);
            this.CheckTemplateConfiguration.TabIndex = 12;
            this.CheckTemplateConfiguration.Text = "校验模板配置项";
            this.CheckTemplateConfiguration.UseVisualStyleBackColor = true;
            this.CheckTemplateConfiguration.Click += new System.EventHandler(this.CheckTemplateConfiguration_Click);
            // 
            // wsNameOrIndex1
            // 
            this.wsNameOrIndex1.Location = new System.Drawing.Point(394, 129);
            this.wsNameOrIndex1.Name = "wsNameOrIndex1";
            this.wsNameOrIndex1.Size = new System.Drawing.Size(100, 21);
            this.wsNameOrIndex1.TabIndex = 14;
            // 
            // wsNameOrIndex2
            // 
            this.wsNameOrIndex2.Location = new System.Drawing.Point(904, 129);
            this.wsNameOrIndex2.Name = "wsNameOrIndex2";
            this.wsNameOrIndex2.Size = new System.Drawing.Size(100, 21);
            this.wsNameOrIndex2.TabIndex = 15;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(392, 114);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(65, 12);
            this.label6.TabIndex = 13;
            this.label6.Text = "序号或名字";
            // 
            // TitleLine1
            // 
            this.TitleLine1.Location = new System.Drawing.Point(394, 168);
            this.TitleLine1.Name = "TitleLine1";
            this.TitleLine1.Size = new System.Drawing.Size(100, 21);
            this.TitleLine1.TabIndex = 16;
            this.TitleLine1.Text = "1";
            // 
            // TitleLine2
            // 
            this.TitleLine2.Location = new System.Drawing.Point(904, 168);
            this.TitleLine2.Name = "TitleLine2";
            this.TitleLine2.Size = new System.Drawing.Size(100, 21);
            this.TitleLine2.TabIndex = 16;
            this.TitleLine2.Text = "1";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(902, 114);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(65, 12);
            this.label7.TabIndex = 13;
            this.label7.Text = "序号或名字";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(392, 153);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(41, 12);
            this.label8.TabIndex = 17;
            this.label8.Text = "标题行";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(902, 153);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(41, 12);
            this.label9.TabIndex = 17;
            this.label9.Text = "标题行";
            // 
            // GenerateConfigurationCode
            // 
            this.GenerateConfigurationCode.Location = new System.Drawing.Point(243, 19);
            this.GenerateConfigurationCode.Name = "GenerateConfigurationCode";
            this.GenerateConfigurationCode.Size = new System.Drawing.Size(110, 23);
            this.GenerateConfigurationCode.TabIndex = 18;
            this.GenerateConfigurationCode.Text = "生成所有配置项";
            this.GenerateConfigurationCode.UseVisualStyleBackColor = true;
            this.GenerateConfigurationCode.Click += new System.EventHandler(this.GenerateConfigurationCode_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.DelHiddenWs);
            this.panel1.Controls.Add(this.GenerateConfigurationCode);
            this.panel1.Controls.Add(this.GenerateConfiguration);
            this.panel1.Location = new System.Drawing.Point(14, 302);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(374, 55);
            this.panel1.TabIndex = 19;
            // 
            // DelHiddenWs
            // 
            this.DelHiddenWs.Location = new System.Drawing.Point(8, 19);
            this.DelHiddenWs.Name = "DelHiddenWs";
            this.DelHiddenWs.Size = new System.Drawing.Size(98, 23);
            this.DelHiddenWs.TabIndex = 19;
            this.DelHiddenWs.Text = "删除隐藏工作簿";
            this.DelHiddenWs.UseVisualStyleBackColor = true;
            this.DelHiddenWs.Click += new System.EventHandler(this.DelHiddenWs_Click);
            // 
            // WScount1
            // 
            this.WScount1.Location = new System.Drawing.Point(394, 66);
            this.WScount1.Name = "WScount1";
            this.WScount1.Size = new System.Drawing.Size(75, 23);
            this.WScount1.TabIndex = 20;
            this.WScount1.Text = "工作簿分析";
            this.WScount1.UseVisualStyleBackColor = true;
            this.WScount1.Click += new System.EventHandler(this.WScount1_Click);
            // 
            // WScount2
            // 
            this.WScount2.Location = new System.Drawing.Point(904, 66);
            this.WScount2.Name = "WScount2";
            this.WScount2.Size = new System.Drawing.Size(75, 23);
            this.WScount2.TabIndex = 24;
            this.WScount2.Text = "工作簿分析";
            this.WScount2.UseVisualStyleBackColor = true;
            this.WScount2.Click += new System.EventHandler(this.WScount2_Click);
            // 
            // dataGridViewExcel2
            // 
            this.dataGridViewExcel2.AllowUserToAddRows = false;
            this.dataGridViewExcel2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewExcel2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5,
            this.dataGridViewTextBoxColumn6});
            this.dataGridViewExcel2.Location = new System.Drawing.Point(527, 95);
            this.dataGridViewExcel2.Name = "dataGridViewExcel2";
            this.dataGridViewExcel2.RowHeadersVisible = false;
            this.dataGridViewExcel2.RowTemplate.Height = 23;
            this.dataGridViewExcel2.Size = new System.Drawing.Size(354, 170);
            this.dataGridViewExcel2.TabIndex = 25;
            this.dataGridViewExcel2.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewExcel2_CellClick);
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.dataGridViewTextBoxColumn4.Frozen = true;
            this.dataGridViewTextBoxColumn4.HeaderText = "序号";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.ReadOnly = true;
            this.dataGridViewTextBoxColumn4.Width = 54;
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.dataGridViewTextBoxColumn5.Frozen = true;
            this.dataGridViewTextBoxColumn5.HeaderText = "名字";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            this.dataGridViewTextBoxColumn5.ReadOnly = true;
            this.dataGridViewTextBoxColumn5.Width = 54;
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.dataGridViewTextBoxColumn6.HeaderText = "标题行";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            this.dataGridViewTextBoxColumn6.Width = 66;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.dataGridViewTextBoxColumn3.HeaderText = "标题行";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.Width = 64;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.dataGridViewTextBoxColumn2.Frozen = true;
            this.dataGridViewTextBoxColumn2.HeaderText = "名字";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.ReadOnly = true;
            this.dataGridViewTextBoxColumn2.Width = 54;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.dataGridViewTextBoxColumn1.Frozen = true;
            this.dataGridViewTextBoxColumn1.HeaderText = "序号";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            this.dataGridViewTextBoxColumn1.Width = 52;
            // 
            // dataGridViewExcel1
            // 
            this.dataGridViewExcel1.AllowUserToAddRows = false;
            this.dataGridViewExcel1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewExcel1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3});
            this.dataGridViewExcel1.Location = new System.Drawing.Point(22, 95);
            this.dataGridViewExcel1.Name = "dataGridViewExcel1";
            this.dataGridViewExcel1.RowHeadersVisible = false;
            this.dataGridViewExcel1.RowTemplate.Height = 23;
            this.dataGridViewExcel1.Size = new System.Drawing.Size(354, 183);
            this.dataGridViewExcel1.TabIndex = 22;
            this.dataGridViewExcel1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewExcel1_CellClick);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1059, 410);
            this.Controls.Add(this.filePath2);
            this.Controls.Add(this.filePath1);
            this.Controls.Add(this.dataGridViewExcel2);
            this.Controls.Add(this.WScount2);
            this.Controls.Add(this.dataGridViewExcel1);
            this.Controls.Add(this.WScount1);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.TitleLine2);
            this.Controls.Add(this.TitleLine1);
            this.Controls.Add(this.wsNameOrIndex2);
            this.Controls.Add(this.wsNameOrIndex1);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.CheckTemplateConfiguration);
            this.Controls.Add(this.SelectFileBtn2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.SelectFileBtn1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panel1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewExcel2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewExcel1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox filePath1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button SelectFileBtn1;
        private System.Windows.Forms.Button GenerateConfiguration;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox filePath2;
        private System.Windows.Forms.Button SelectFileBtn2;
        private System.Windows.Forms.Button CheckTemplateConfiguration;
        private System.Windows.Forms.TextBox wsNameOrIndex1;
        private System.Windows.Forms.TextBox wsNameOrIndex2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox TitleLine1;
        private System.Windows.Forms.TextBox TitleLine2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button GenerateConfigurationCode;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button WScount1;
        private System.Windows.Forms.Button WScount2;
        private System.Windows.Forms.DataGridView dataGridViewExcel2;
        private System.Windows.Forms.Button DelHiddenWs;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridView dataGridViewExcel1;
    }
}

