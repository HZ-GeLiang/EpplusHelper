namespace EPPlusTool
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.filePath1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SelectFileBtn1 = new System.Windows.Forms.Button();
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
            this.BtnAnalyze1 = new System.Windows.Forms.Button();
            this.BtnAnalyze2 = new System.Windows.Forms.Button();
            this.dgv2 = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgv1 = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.GenerateConfiguration = new System.Windows.Forms.Button();
            this.GenerateConfigurationCode = new System.Windows.Forms.Button();
            this.DelHiddenWs = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label15 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.TitleCol1 = new System.Windows.Forms.TextBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label16 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.TitleCol2 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.CreateDataTable = new System.Windows.Forms.Button();
            this.diaplayRowAndColumn = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // filePath1
            // 
            this.filePath1.AllowDrop = true;
            this.filePath1.Location = new System.Drawing.Point(12, 23);
            this.filePath1.Multiline = true;
            this.filePath1.Name = "filePath1";
            this.filePath1.Size = new System.Drawing.Size(501, 52);
            this.filePath1.TabIndex = 4;
            this.filePath1.DragDrop += new System.Windows.Forms.DragEventHandler(this.TextBoxDragDrop);
            this.filePath1.DragEnter += new System.Windows.Forms.DragEventHandler(this.TextBoxDragEnter);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(149, 12);
            this.label1.TabIndex = 5;
            this.label1.Text = "文件路径(文本框支持拖拽)";
            // 
            // SelectFileBtn1
            // 
            this.SelectFileBtn1.Location = new System.Drawing.Point(519, 23);
            this.SelectFileBtn1.Name = "SelectFileBtn1";
            this.SelectFileBtn1.Size = new System.Drawing.Size(95, 24);
            this.SelectFileBtn1.TabIndex = 6;
            this.SelectFileBtn1.Text = "选择...";
            this.SelectFileBtn1.UseVisualStyleBackColor = true;
            this.SelectFileBtn1.Click += new System.EventHandler(this.Btn_SelectExcelFile);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(10, 286);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(149, 12);
            this.label3.TabIndex = 9;
            this.label3.Text = "文件路径(文本框支持拖拽)";
            // 
            // filePath2
            // 
            this.filePath2.AllowDrop = true;
            this.filePath2.Location = new System.Drawing.Point(12, 302);
            this.filePath2.Multiline = true;
            this.filePath2.Name = "filePath2";
            this.filePath2.Size = new System.Drawing.Size(501, 50);
            this.filePath2.TabIndex = 10;
            this.filePath2.DragDrop += new System.Windows.Forms.DragEventHandler(this.TextBoxDragDrop);
            this.filePath2.DragEnter += new System.Windows.Forms.DragEventHandler(this.TextBoxDragEnter);
            // 
            // SelectFileBtn2
            // 
            this.SelectFileBtn2.Location = new System.Drawing.Point(519, 302);
            this.SelectFileBtn2.Name = "SelectFileBtn2";
            this.SelectFileBtn2.Size = new System.Drawing.Size(95, 23);
            this.SelectFileBtn2.TabIndex = 11;
            this.SelectFileBtn2.Text = "选择...";
            this.SelectFileBtn2.UseVisualStyleBackColor = true;
            this.SelectFileBtn2.Click += new System.EventHandler(this.Btn_SelectExcelFile);
            // 
            // CheckTemplateConfiguration
            // 
            this.CheckTemplateConfiguration.Location = new System.Drawing.Point(39, 112);
            this.CheckTemplateConfiguration.Name = "CheckTemplateConfiguration";
            this.CheckTemplateConfiguration.Size = new System.Drawing.Size(78, 50);
            this.CheckTemplateConfiguration.TabIndex = 12;
            this.CheckTemplateConfiguration.Text = "行内容";
            this.CheckTemplateConfiguration.UseVisualStyleBackColor = true;
            this.CheckTemplateConfiguration.Click += new System.EventHandler(this.CheckTemplateConfiguration_Click);
            // 
            // wsNameOrIndex1
            // 
            this.wsNameOrIndex1.Location = new System.Drawing.Point(4, 26);
            this.wsNameOrIndex1.Name = "wsNameOrIndex1";
            this.wsNameOrIndex1.ReadOnly = true;
            this.wsNameOrIndex1.Size = new System.Drawing.Size(100, 21);
            this.wsNameOrIndex1.TabIndex = 14;
            // 
            // wsNameOrIndex2
            // 
            this.wsNameOrIndex2.Location = new System.Drawing.Point(392, 391);
            this.wsNameOrIndex2.Name = "wsNameOrIndex2";
            this.wsNameOrIndex2.ReadOnly = true;
            this.wsNameOrIndex2.Size = new System.Drawing.Size(100, 21);
            this.wsNameOrIndex2.TabIndex = 15;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(2, 10);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(65, 12);
            this.label6.TabIndex = 13;
            this.label6.Text = "序号或名字";
            // 
            // TitleLine1
            // 
            this.TitleLine1.Location = new System.Drawing.Point(52, 53);
            this.TitleLine1.Name = "TitleLine1";
            this.TitleLine1.ReadOnly = true;
            this.TitleLine1.Size = new System.Drawing.Size(52, 21);
            this.TitleLine1.TabIndex = 16;
            // 
            // TitleLine2
            // 
            this.TitleLine2.Location = new System.Drawing.Point(47, 53);
            this.TitleLine2.Name = "TitleLine2";
            this.TitleLine2.ReadOnly = true;
            this.TitleLine2.Size = new System.Drawing.Size(52, 21);
            this.TitleLine2.TabIndex = 16;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(392, 372);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(65, 12);
            this.label7.TabIndex = 13;
            this.label7.Text = "序号或名字";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(394, 142);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(41, 12);
            this.label8.TabIndex = 17;
            this.label8.Text = "行位置";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(2, 57);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(41, 12);
            this.label9.TabIndex = 17;
            this.label9.Text = "行位置";
            // 
            // BtnAnalyze1
            // 
            this.BtnAnalyze1.Location = new System.Drawing.Point(519, 53);
            this.BtnAnalyze1.Name = "BtnAnalyze1";
            this.BtnAnalyze1.Size = new System.Drawing.Size(95, 23);
            this.BtnAnalyze1.TabIndex = 20;
            this.BtnAnalyze1.Text = "工作簿分析";
            this.BtnAnalyze1.UseVisualStyleBackColor = true;
            this.BtnAnalyze1.Click += new System.EventHandler(this.LoadDgv);
            // 
            // BtnAnalyze2
            // 
            this.BtnAnalyze2.Location = new System.Drawing.Point(519, 329);
            this.BtnAnalyze2.Name = "BtnAnalyze2";
            this.BtnAnalyze2.Size = new System.Drawing.Size(95, 23);
            this.BtnAnalyze2.TabIndex = 24;
            this.BtnAnalyze2.Text = "工作簿分析";
            this.BtnAnalyze2.UseVisualStyleBackColor = true;
            this.BtnAnalyze2.Click += new System.EventHandler(this.LoadDgv);
            // 
            // dgv2
            // 
            this.dgv2.AllowUserToAddRows = false;
            this.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5,
            this.dataGridViewTextBoxColumn6,
            this.Column2});
            this.dgv2.Location = new System.Drawing.Point(12, 363);
            this.dgv2.Name = "dgv2";
            this.dgv2.RowHeadersVisible = false;
            this.dgv2.RowTemplate.Height = 23;
            this.dgv2.Size = new System.Drawing.Size(369, 183);
            this.dgv2.TabIndex = 25;
            this.dgv2.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellEndEdit);
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
            // Column2
            // 
            this.Column2.HeaderText = "标题列";
            this.Column2.Name = "Column2";
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
            // dgv1
            // 
            this.dgv1.AllowUserToAddRows = false;
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3,
            this.Column1});
            this.dgv1.Location = new System.Drawing.Point(12, 81);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersVisible = false;
            this.dgv1.RowTemplate.Height = 23;
            this.dgv1.Size = new System.Drawing.Size(373, 195);
            this.dgv1.TabIndex = 22;
            this.dgv1.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellEndEdit);
            // 
            // Column1
            // 
            this.Column1.HeaderText = "标题列";
            this.Column1.Name = "Column1";
            // 
            // GenerateConfiguration
            // 
            this.GenerateConfiguration.Location = new System.Drawing.Point(52, 106);
            this.GenerateConfiguration.Name = "GenerateConfiguration";
            this.GenerateConfiguration.Size = new System.Drawing.Size(66, 24);
            this.GenerateConfiguration.TabIndex = 7;
            this.GenerateConfiguration.Text = "填充配置";
            this.GenerateConfiguration.UseVisualStyleBackColor = true;
            this.GenerateConfiguration.Click += new System.EventHandler(this.GenerateConfiguration_Click);
            // 
            // GenerateConfigurationCode
            // 
            this.GenerateConfigurationCode.Location = new System.Drawing.Point(516, 130);
            this.GenerateConfigurationCode.Name = "GenerateConfigurationCode";
            this.GenerateConfigurationCode.Size = new System.Drawing.Size(98, 24);
            this.GenerateConfigurationCode.TabIndex = 18;
            this.GenerateConfigurationCode.Text = "生成所有配置";
            this.GenerateConfigurationCode.UseVisualStyleBackColor = true;
            this.GenerateConfigurationCode.Click += new System.EventHandler(this.GenerateConfigurationCode_Click);
            // 
            // DelHiddenWs
            // 
            this.DelHiddenWs.Location = new System.Drawing.Point(516, 102);
            this.DelHiddenWs.Name = "DelHiddenWs";
            this.DelHiddenWs.Size = new System.Drawing.Size(98, 22);
            this.DelHiddenWs.TabIndex = 19;
            this.DelHiddenWs.Text = "删除所有隐藏";
            this.DelHiddenWs.UseVisualStyleBackColor = true;
            this.DelHiddenWs.Click += new System.EventHandler(this.DelHiddenWs_Click);
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.label15);
            this.panel2.Controls.Add(this.label14);
            this.panel2.Controls.Add(this.label13);
            this.panel2.Controls.Add(this.button1);
            this.panel2.Controls.Add(this.label10);
            this.panel2.Controls.Add(this.TitleCol1);
            this.panel2.Controls.Add(this.wsNameOrIndex1);
            this.panel2.Controls.Add(this.label6);
            this.panel2.Controls.Add(this.GenerateConfiguration);
            this.panel2.Controls.Add(this.TitleLine1);
            this.panel2.Location = new System.Drawing.Point(387, 84);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(123, 192);
            this.panel2.TabIndex = 26;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(6, 126);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(35, 12);
            this.label15.TabIndex = 34;
            this.label15.Text = "Sheet";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(6, 145);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(29, 12);
            this.label14.TabIndex = 33;
            this.label14.Text = "操作";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(7, 107);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(29, 12);
            this.label13.TabIndex = 32;
            this.label13.Text = "单个";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(41, 133);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(77, 24);
            this.button1.TabIndex = 31;
            this.button1.Text = "创建 Class";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.CreateClass_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(5, 83);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(41, 12);
            this.label10.TabIndex = 30;
            this.label10.Text = "列位置";
            // 
            // TitleCol1
            // 
            this.TitleCol1.Location = new System.Drawing.Point(52, 79);
            this.TitleCol1.Name = "TitleCol1";
            this.TitleCol1.ReadOnly = true;
            this.TitleCol1.Size = new System.Drawing.Size(52, 21);
            this.TitleCol1.TabIndex = 30;
            // 
            // panel3
            // 
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.label16);
            this.panel3.Controls.Add(this.label17);
            this.panel3.Controls.Add(this.label11);
            this.panel3.Controls.Add(this.label18);
            this.panel3.Controls.Add(this.TitleCol2);
            this.panel3.Controls.Add(this.TitleLine2);
            this.panel3.Controls.Add(this.label9);
            this.panel3.Controls.Add(this.CheckTemplateConfiguration);
            this.panel3.Location = new System.Drawing.Point(387, 364);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(123, 182);
            this.panel3.TabIndex = 27;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(3, 131);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(35, 12);
            this.label16.TabIndex = 37;
            this.label16.Text = "Sheet";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(3, 150);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(29, 12);
            this.label17.TabIndex = 36;
            this.label17.Text = "比较";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Cursor = System.Windows.Forms.Cursors.SizeAll;
            this.label11.Location = new System.Drawing.Point(2, 87);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(41, 12);
            this.label11.TabIndex = 31;
            this.label11.Text = "列位置";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(4, 112);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(29, 12);
            this.label18.TabIndex = 35;
            this.label18.Text = "上下";
            // 
            // TitleCol2
            // 
            this.TitleCol2.Location = new System.Drawing.Point(47, 83);
            this.TitleCol2.Name = "TitleCol2";
            this.TitleCol2.ReadOnly = true;
            this.TitleCol2.Size = new System.Drawing.Size(52, 21);
            this.TitleCol2.TabIndex = 32;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(404, 78);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 12);
            this.label4.TabIndex = 28;
            this.label4.Text = "Sheet信息";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(405, 358);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 12);
            this.label5.TabIndex = 29;
            this.label5.Text = "Sheet信息";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(517, 84);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 12);
            this.label2.TabIndex = 30;
            this.label2.Text = "批量Sheet操作";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(517, 165);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(0, 12);
            this.label12.TabIndex = 31;
            // 
            // CreateDataTable
            // 
            this.CreateDataTable.Location = new System.Drawing.Point(406, 247);
            this.CreateDataTable.Name = "CreateDataTable";
            this.CreateDataTable.Size = new System.Drawing.Size(99, 24);
            this.CreateDataTable.TabIndex = 35;
            this.CreateDataTable.Text = "创建 DataTable";
            this.CreateDataTable.UseVisualStyleBackColor = true;
            this.CreateDataTable.Click += new System.EventHandler(this.CreateDataTable_Click);
            // 
            // diaplayRowAndColumn
            // 
            this.diaplayRowAndColumn.Location = new System.Drawing.Point(516, 160);
            this.diaplayRowAndColumn.Name = "diaplayRowAndColumn";
            this.diaplayRowAndColumn.Size = new System.Drawing.Size(98, 25);
            this.diaplayRowAndColumn.TabIndex = 36;
            this.diaplayRowAndColumn.Text = "显示所有行和列";
            this.diaplayRowAndColumn.UseVisualStyleBackColor = true;
            this.diaplayRowAndColumn.Click += new System.EventHandler(this.diaplayRowAndColumn_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(619, 555);
            this.Controls.Add(this.diaplayRowAndColumn);
            this.Controls.Add(this.CreateDataTable);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.GenerateConfigurationCode);
            this.Controls.Add(this.DelHiddenWs);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.filePath2);
            this.Controls.Add(this.filePath1);
            this.Controls.Add(this.dgv2);
            this.Controls.Add(this.BtnAnalyze2);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.BtnAnalyze1);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.wsNameOrIndex2);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.SelectFileBtn2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.SelectFileBtn1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel3);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "EPPlusTool Owner By GeLiang ";
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox filePath1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button SelectFileBtn1;
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
        private System.Windows.Forms.Button BtnAnalyze1;
        private System.Windows.Forms.Button BtnAnalyze2;
        private System.Windows.Forms.DataGridView dgv2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridView dgv1;
        private System.Windows.Forms.Button GenerateConfiguration;
        private System.Windows.Forms.Button GenerateConfigurationCode;
        private System.Windows.Forms.Button DelHiddenWs;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox TitleCol1;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox TitleCol2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Button CreateDataTable;
        private System.Windows.Forms.Button diaplayRowAndColumn;
    }
}

