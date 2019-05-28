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
            this.label4 = new System.Windows.Forms.Label();
            this.wsNameOrIndex1 = new System.Windows.Forms.TextBox();
            this.wsNameOrIndex2 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.TitleLine1 = new System.Windows.Forms.TextBox();
            this.TitleLine2 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.GenerateConfigurationCode = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // filePath1
            // 
            this.filePath1.AllowDrop = true;
            this.filePath1.Location = new System.Drawing.Point(22, 23);
            this.filePath1.Multiline = true;
            this.filePath1.Name = "filePath1";
            this.filePath1.Size = new System.Drawing.Size(661, 49);
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
            this.SelectFileBtn1.Location = new System.Drawing.Point(712, 26);
            this.SelectFileBtn1.Name = "SelectFileBtn1";
            this.SelectFileBtn1.Size = new System.Drawing.Size(74, 37);
            this.SelectFileBtn1.TabIndex = 6;
            this.SelectFileBtn1.Text = "选择...";
            this.SelectFileBtn1.UseVisualStyleBackColor = true;
            this.SelectFileBtn1.Click += new System.EventHandler(this.btn_SelectExcelFile);
            // 
            // GenerateConfiguration
            // 
            this.GenerateConfiguration.Location = new System.Drawing.Point(25, 106);
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
            this.label2.Location = new System.Drawing.Point(23, 79);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(113, 12);
            this.label2.TabIndex = 8;
            this.label2.Text = "自动初始化填充配置";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(21, 165);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(155, 12);
            this.label3.TabIndex = 9;
            this.label3.Text = "文件路径2(文本框支持拖拽)";
            // 
            // filePath2
            // 
            this.filePath2.AllowDrop = true;
            this.filePath2.Location = new System.Drawing.Point(25, 200);
            this.filePath2.Multiline = true;
            this.filePath2.Name = "filePath2";
            this.filePath2.Size = new System.Drawing.Size(661, 49);
            this.filePath2.TabIndex = 10;
            this.filePath2.DragDrop += new System.Windows.Forms.DragEventHandler(this.textBoxDragDrop);
            this.filePath2.DragEnter += new System.Windows.Forms.DragEventHandler(this.textBoxDragEnter);
            // 
            // SelectFileBtn2
            // 
            this.SelectFileBtn2.Location = new System.Drawing.Point(712, 200);
            this.SelectFileBtn2.Name = "SelectFileBtn2";
            this.SelectFileBtn2.Size = new System.Drawing.Size(74, 37);
            this.SelectFileBtn2.TabIndex = 11;
            this.SelectFileBtn2.Text = "选择...";
            this.SelectFileBtn2.UseVisualStyleBackColor = true;
            this.SelectFileBtn2.Click += new System.EventHandler(this.btn_SelectExcelFile);
            // 
            // CheckTemplateConfiguration
            // 
            this.CheckTemplateConfiguration.Location = new System.Drawing.Point(241, 389);
            this.CheckTemplateConfiguration.Name = "CheckTemplateConfiguration";
            this.CheckTemplateConfiguration.Size = new System.Drawing.Size(90, 23);
            this.CheckTemplateConfiguration.TabIndex = 12;
            this.CheckTemplateConfiguration.Text = "校验模板配置项";
            this.CheckTemplateConfiguration.UseVisualStyleBackColor = true;
            this.CheckTemplateConfiguration.Click += new System.EventHandler(this.CheckTemplateConfiguration_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(25, 274);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(113, 12);
            this.label4.TabIndex = 13;
            this.label4.Text = "模板文件(文件路径)";
            // 
            // wsNameOrIndex1
            // 
            this.wsNameOrIndex1.Location = new System.Drawing.Point(146, 296);
            this.wsNameOrIndex1.Name = "wsNameOrIndex1";
            this.wsNameOrIndex1.Size = new System.Drawing.Size(100, 21);
            this.wsNameOrIndex1.TabIndex = 14;
            // 
            // wsNameOrIndex2
            // 
            this.wsNameOrIndex2.Location = new System.Drawing.Point(452, 296);
            this.wsNameOrIndex2.Name = "wsNameOrIndex2";
            this.wsNameOrIndex2.Size = new System.Drawing.Size(100, 21);
            this.wsNameOrIndex2.TabIndex = 15;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(339, 274);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(119, 12);
            this.label5.TabIndex = 13;
            this.label5.Text = "对比模版(文件路径2)";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(23, 299);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(107, 12);
            this.label6.TabIndex = 13;
            this.label6.Text = "worksheet名或序号";
            // 
            // TitleLine1
            // 
            this.TitleLine1.Location = new System.Drawing.Point(146, 343);
            this.TitleLine1.Name = "TitleLine1";
            this.TitleLine1.Size = new System.Drawing.Size(100, 21);
            this.TitleLine1.TabIndex = 16;
            this.TitleLine1.Text = "1";
            // 
            // TitleLine2
            // 
            this.TitleLine2.Location = new System.Drawing.Point(452, 343);
            this.TitleLine2.Name = "TitleLine2";
            this.TitleLine2.Size = new System.Drawing.Size(100, 21);
            this.TitleLine2.TabIndex = 16;
            this.TitleLine2.Text = "1";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(339, 299);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(107, 12);
            this.label7.TabIndex = 13;
            this.label7.Text = "worksheet名或序号";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(25, 346);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(41, 12);
            this.label8.TabIndex = 17;
            this.label8.Text = "标题行";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(339, 346);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(41, 12);
            this.label9.TabIndex = 17;
            this.label9.Text = "标题行";
            // 
            // GenerateConfigurationCode
            // 
            this.GenerateConfigurationCode.Location = new System.Drawing.Point(171, 106);
            this.GenerateConfigurationCode.Name = "GenerateConfigurationCode";
            this.GenerateConfigurationCode.Size = new System.Drawing.Size(88, 23);
            this.GenerateConfigurationCode.TabIndex = 18;
            this.GenerateConfigurationCode.Text = "生成配置项";
            this.GenerateConfigurationCode.UseVisualStyleBackColor = true;
            this.GenerateConfigurationCode.Click += new System.EventHandler(this.GenerateConfigurationCode_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.GenerateConfigurationCode);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.TitleLine2);
            this.Controls.Add(this.TitleLine1);
            this.Controls.Add(this.wsNameOrIndex2);
            this.Controls.Add(this.wsNameOrIndex1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.CheckTemplateConfiguration);
            this.Controls.Add(this.SelectFileBtn2);
            this.Controls.Add(this.filePath2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.GenerateConfiguration);
            this.Controls.Add(this.SelectFileBtn1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.filePath1);
            this.Name = "Form1";
            this.Text = "Form1";
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
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox wsNameOrIndex1;
        private System.Windows.Forms.TextBox wsNameOrIndex2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox TitleLine1;
        private System.Windows.Forms.TextBox TitleLine2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button GenerateConfigurationCode;
    }
}

