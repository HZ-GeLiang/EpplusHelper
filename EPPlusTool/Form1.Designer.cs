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
            filePath1 = new System.Windows.Forms.TextBox();
            label1 = new System.Windows.Forms.Label();
            SelectFileBtn1 = new System.Windows.Forms.Button();
            label3 = new System.Windows.Forms.Label();
            filePath2 = new System.Windows.Forms.TextBox();
            SelectFileBtn2 = new System.Windows.Forms.Button();
            CheckTemplateConfiguration = new System.Windows.Forms.Button();
            wsNameOrIndex1 = new System.Windows.Forms.TextBox();
            wsNameOrIndex2 = new System.Windows.Forms.TextBox();
            label6 = new System.Windows.Forms.Label();
            TitleLine1 = new System.Windows.Forms.TextBox();
            TitleLine2 = new System.Windows.Forms.TextBox();
            label7 = new System.Windows.Forms.Label();
            label8 = new System.Windows.Forms.Label();
            label9 = new System.Windows.Forms.Label();
            BtnAnalyze1 = new System.Windows.Forms.Button();
            BtnAnalyze2 = new System.Windows.Forms.Button();
            dgv2 = new System.Windows.Forms.DataGridView();
            dgv1 = new System.Windows.Forms.DataGridView();
            GenerateConfiguration = new System.Windows.Forms.Button();
            GenerateConfigurationCode = new System.Windows.Forms.Button();
            DelHiddenWs = new System.Windows.Forms.Button();
            panel2 = new System.Windows.Forms.Panel();
            label15 = new System.Windows.Forms.Label();
            label14 = new System.Windows.Forms.Label();
            label13 = new System.Windows.Forms.Label();
            button1 = new System.Windows.Forms.Button();
            label10 = new System.Windows.Forms.Label();
            TitleCol1 = new System.Windows.Forms.TextBox();
            panel3 = new System.Windows.Forms.Panel();
            label16 = new System.Windows.Forms.Label();
            label17 = new System.Windows.Forms.Label();
            label11 = new System.Windows.Forms.Label();
            label18 = new System.Windows.Forms.Label();
            TitleCol2 = new System.Windows.Forms.TextBox();
            label4 = new System.Windows.Forms.Label();
            label5 = new System.Windows.Forms.Label();
            label2 = new System.Windows.Forms.Label();
            label12 = new System.Windows.Forms.Label();
            CreateDataTable = new System.Windows.Forms.Button();
            diaplayRowAndColumn = new System.Windows.Forms.Button();
            dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)dgv2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dgv1).BeginInit();
            panel2.SuspendLayout();
            panel3.SuspendLayout();
            SuspendLayout();
            // 
            // filePath1
            // 
            filePath1.AllowDrop = true;
            filePath1.Location = new System.Drawing.Point(14, 33);
            filePath1.Margin = new System.Windows.Forms.Padding(4);
            filePath1.Multiline = true;
            filePath1.Name = "filePath1";
            filePath1.Size = new System.Drawing.Size(584, 72);
            filePath1.TabIndex = 4;
            filePath1.DragDrop += TextBoxDragDrop;
            filePath1.DragEnter += TextBoxDragEnter;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(13, 10);
            label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(148, 17);
            label1.TabIndex = 5;
            label1.Text = "文件路径(文本框支持拖拽)";
            // 
            // SelectFileBtn1
            // 
            SelectFileBtn1.Location = new System.Drawing.Point(606, 33);
            SelectFileBtn1.Margin = new System.Windows.Forms.Padding(4);
            SelectFileBtn1.Name = "SelectFileBtn1";
            SelectFileBtn1.Size = new System.Drawing.Size(111, 34);
            SelectFileBtn1.TabIndex = 6;
            SelectFileBtn1.Text = "选择...";
            SelectFileBtn1.UseVisualStyleBackColor = true;
            SelectFileBtn1.Click += Btn_SelectExcelFile;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new System.Drawing.Point(12, 405);
            label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label3.Name = "label3";
            label3.Size = new System.Drawing.Size(148, 17);
            label3.TabIndex = 9;
            label3.Text = "文件路径(文本框支持拖拽)";
            // 
            // filePath2
            // 
            filePath2.AllowDrop = true;
            filePath2.Location = new System.Drawing.Point(14, 428);
            filePath2.Margin = new System.Windows.Forms.Padding(4);
            filePath2.Multiline = true;
            filePath2.Name = "filePath2";
            filePath2.Size = new System.Drawing.Size(584, 69);
            filePath2.TabIndex = 10;
            filePath2.DragDrop += TextBoxDragDrop;
            filePath2.DragEnter += TextBoxDragEnter;
            // 
            // SelectFileBtn2
            // 
            SelectFileBtn2.Location = new System.Drawing.Point(606, 428);
            SelectFileBtn2.Margin = new System.Windows.Forms.Padding(4);
            SelectFileBtn2.Name = "SelectFileBtn2";
            SelectFileBtn2.Size = new System.Drawing.Size(111, 33);
            SelectFileBtn2.TabIndex = 11;
            SelectFileBtn2.Text = "选择...";
            SelectFileBtn2.UseVisualStyleBackColor = true;
            SelectFileBtn2.Click += Btn_SelectExcelFile;
            // 
            // CheckTemplateConfiguration
            // 
            CheckTemplateConfiguration.Location = new System.Drawing.Point(46, 159);
            CheckTemplateConfiguration.Margin = new System.Windows.Forms.Padding(4);
            CheckTemplateConfiguration.Name = "CheckTemplateConfiguration";
            CheckTemplateConfiguration.Size = new System.Drawing.Size(91, 71);
            CheckTemplateConfiguration.TabIndex = 12;
            CheckTemplateConfiguration.Text = "行内容";
            CheckTemplateConfiguration.UseVisualStyleBackColor = true;
            CheckTemplateConfiguration.Click += CheckTemplateConfiguration_Click;
            // 
            // wsNameOrIndex1
            // 
            wsNameOrIndex1.Location = new System.Drawing.Point(5, 37);
            wsNameOrIndex1.Margin = new System.Windows.Forms.Padding(4);
            wsNameOrIndex1.Name = "wsNameOrIndex1";
            wsNameOrIndex1.ReadOnly = true;
            wsNameOrIndex1.Size = new System.Drawing.Size(116, 23);
            wsNameOrIndex1.TabIndex = 14;
            // 
            // wsNameOrIndex2
            // 
            wsNameOrIndex2.Location = new System.Drawing.Point(457, 554);
            wsNameOrIndex2.Margin = new System.Windows.Forms.Padding(4);
            wsNameOrIndex2.Name = "wsNameOrIndex2";
            wsNameOrIndex2.ReadOnly = true;
            wsNameOrIndex2.Size = new System.Drawing.Size(116, 23);
            wsNameOrIndex2.TabIndex = 15;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new System.Drawing.Point(2, 14);
            label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label6.Name = "label6";
            label6.Size = new System.Drawing.Size(68, 17);
            label6.TabIndex = 13;
            label6.Text = "序号或名字";
            // 
            // TitleLine1
            // 
            TitleLine1.Location = new System.Drawing.Point(61, 75);
            TitleLine1.Margin = new System.Windows.Forms.Padding(4);
            TitleLine1.Name = "TitleLine1";
            TitleLine1.ReadOnly = true;
            TitleLine1.Size = new System.Drawing.Size(60, 23);
            TitleLine1.TabIndex = 16;
            // 
            // TitleLine2
            // 
            TitleLine2.Location = new System.Drawing.Point(55, 75);
            TitleLine2.Margin = new System.Windows.Forms.Padding(4);
            TitleLine2.Name = "TitleLine2";
            TitleLine2.ReadOnly = true;
            TitleLine2.Size = new System.Drawing.Size(60, 23);
            TitleLine2.TabIndex = 16;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new System.Drawing.Point(457, 527);
            label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label7.Name = "label7";
            label7.Size = new System.Drawing.Size(68, 17);
            label7.TabIndex = 13;
            label7.Text = "序号或名字";
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Location = new System.Drawing.Point(460, 201);
            label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label8.Name = "label8";
            label8.Size = new System.Drawing.Size(44, 17);
            label8.TabIndex = 17;
            label8.Text = "行位置";
            // 
            // label9
            // 
            label9.AutoSize = true;
            label9.Location = new System.Drawing.Point(2, 81);
            label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label9.Name = "label9";
            label9.Size = new System.Drawing.Size(44, 17);
            label9.TabIndex = 17;
            label9.Text = "行位置";
            // 
            // BtnAnalyze1
            // 
            BtnAnalyze1.Location = new System.Drawing.Point(606, 75);
            BtnAnalyze1.Margin = new System.Windows.Forms.Padding(4);
            BtnAnalyze1.Name = "BtnAnalyze1";
            BtnAnalyze1.Size = new System.Drawing.Size(111, 33);
            BtnAnalyze1.TabIndex = 20;
            BtnAnalyze1.Text = "工作簿分析";
            BtnAnalyze1.UseVisualStyleBackColor = true;
            BtnAnalyze1.Click += LoadDgv;
            // 
            // BtnAnalyze2
            // 
            BtnAnalyze2.Location = new System.Drawing.Point(606, 466);
            BtnAnalyze2.Margin = new System.Windows.Forms.Padding(4);
            BtnAnalyze2.Name = "BtnAnalyze2";
            BtnAnalyze2.Size = new System.Drawing.Size(111, 33);
            BtnAnalyze2.TabIndex = 24;
            BtnAnalyze2.Text = "工作簿分析";
            BtnAnalyze2.UseVisualStyleBackColor = true;
            BtnAnalyze2.Click += LoadDgv;
            // 
            // dgv2
            // 
            dgv2.AllowUserToAddRows = false;
            dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgv2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] { dataGridViewTextBoxColumn4, dataGridViewTextBoxColumn5, dataGridViewTextBoxColumn6, Column2 });
            dgv2.Location = new System.Drawing.Point(14, 514);
            dgv2.Margin = new System.Windows.Forms.Padding(4);
            dgv2.Name = "dgv2";
            dgv2.RowHeadersVisible = false;
            dgv2.RowTemplate.Height = 23;
            dgv2.Size = new System.Drawing.Size(430, 259);
            dgv2.TabIndex = 25;
            dgv2.CellEndEdit += dgv_CellEndEdit;
            // 
            // dgv1
            // 
            dgv1.AllowUserToAddRows = false;
            dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgv1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] { dataGridViewTextBoxColumn1, dataGridViewTextBoxColumn2, dataGridViewTextBoxColumn3, Column1 });
            dgv1.Location = new System.Drawing.Point(14, 115);
            dgv1.Margin = new System.Windows.Forms.Padding(4);
            dgv1.Name = "dgv1";
            dgv1.RowHeadersVisible = false;
            dgv1.RowTemplate.Height = 23;
            dgv1.Size = new System.Drawing.Size(435, 276);
            dgv1.TabIndex = 22;
            dgv1.CellClick += dgv1_CellClick;
            dgv1.CellEndEdit += dgv_CellEndEdit;
            // 
            // GenerateConfiguration
            // 
            GenerateConfiguration.Location = new System.Drawing.Point(61, 150);
            GenerateConfiguration.Margin = new System.Windows.Forms.Padding(4);
            GenerateConfiguration.Name = "GenerateConfiguration";
            GenerateConfiguration.Size = new System.Drawing.Size(77, 34);
            GenerateConfiguration.TabIndex = 7;
            GenerateConfiguration.Text = "填充配置";
            GenerateConfiguration.UseVisualStyleBackColor = true;
            GenerateConfiguration.Click += GenerateConfiguration_Click;
            // 
            // GenerateConfigurationCode
            // 
            GenerateConfigurationCode.Location = new System.Drawing.Point(602, 184);
            GenerateConfigurationCode.Margin = new System.Windows.Forms.Padding(4);
            GenerateConfigurationCode.Name = "GenerateConfigurationCode";
            GenerateConfigurationCode.Size = new System.Drawing.Size(114, 34);
            GenerateConfigurationCode.TabIndex = 18;
            GenerateConfigurationCode.Text = "生成所有配置";
            GenerateConfigurationCode.UseVisualStyleBackColor = true;
            GenerateConfigurationCode.Click += GenerateConfigurationCode_Click;
            // 
            // DelHiddenWs
            // 
            DelHiddenWs.Location = new System.Drawing.Point(602, 144);
            DelHiddenWs.Margin = new System.Windows.Forms.Padding(4);
            DelHiddenWs.Name = "DelHiddenWs";
            DelHiddenWs.Size = new System.Drawing.Size(114, 31);
            DelHiddenWs.TabIndex = 19;
            DelHiddenWs.Text = "删除所有隐藏";
            DelHiddenWs.UseVisualStyleBackColor = true;
            DelHiddenWs.Click += DelHiddenWs_Click;
            // 
            // panel2
            // 
            panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            panel2.Controls.Add(label15);
            panel2.Controls.Add(label14);
            panel2.Controls.Add(label13);
            panel2.Controls.Add(button1);
            panel2.Controls.Add(label10);
            panel2.Controls.Add(TitleCol1);
            panel2.Controls.Add(wsNameOrIndex1);
            panel2.Controls.Add(label6);
            panel2.Controls.Add(GenerateConfiguration);
            panel2.Controls.Add(TitleLine1);
            panel2.Location = new System.Drawing.Point(451, 119);
            panel2.Margin = new System.Windows.Forms.Padding(4);
            panel2.Name = "panel2";
            panel2.Size = new System.Drawing.Size(143, 271);
            panel2.TabIndex = 26;
            // 
            // label15
            // 
            label15.AutoSize = true;
            label15.Location = new System.Drawing.Point(7, 178);
            label15.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label15.Name = "label15";
            label15.Size = new System.Drawing.Size(40, 17);
            label15.TabIndex = 34;
            label15.Text = "Sheet";
            // 
            // label14
            // 
            label14.AutoSize = true;
            label14.Location = new System.Drawing.Point(7, 205);
            label14.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label14.Name = "label14";
            label14.Size = new System.Drawing.Size(32, 17);
            label14.TabIndex = 33;
            label14.Text = "操作";
            // 
            // label13
            // 
            label13.AutoSize = true;
            label13.Location = new System.Drawing.Point(8, 152);
            label13.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label13.Name = "label13";
            label13.Size = new System.Drawing.Size(32, 17);
            label13.TabIndex = 32;
            label13.Text = "单个";
            // 
            // button1
            // 
            button1.Location = new System.Drawing.Point(48, 188);
            button1.Margin = new System.Windows.Forms.Padding(4);
            button1.Name = "button1";
            button1.Size = new System.Drawing.Size(90, 34);
            button1.TabIndex = 31;
            button1.Text = "创建 Class";
            button1.UseVisualStyleBackColor = true;
            button1.Click += CreateClass_Click;
            // 
            // label10
            // 
            label10.AutoSize = true;
            label10.Location = new System.Drawing.Point(6, 118);
            label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label10.Name = "label10";
            label10.Size = new System.Drawing.Size(44, 17);
            label10.TabIndex = 30;
            label10.Text = "列位置";
            // 
            // TitleCol1
            // 
            TitleCol1.Location = new System.Drawing.Point(61, 112);
            TitleCol1.Margin = new System.Windows.Forms.Padding(4);
            TitleCol1.Name = "TitleCol1";
            TitleCol1.ReadOnly = true;
            TitleCol1.Size = new System.Drawing.Size(60, 23);
            TitleCol1.TabIndex = 30;
            // 
            // panel3
            // 
            panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            panel3.Controls.Add(label16);
            panel3.Controls.Add(label17);
            panel3.Controls.Add(label11);
            panel3.Controls.Add(label18);
            panel3.Controls.Add(TitleCol2);
            panel3.Controls.Add(TitleLine2);
            panel3.Controls.Add(label9);
            panel3.Controls.Add(CheckTemplateConfiguration);
            panel3.Location = new System.Drawing.Point(451, 516);
            panel3.Margin = new System.Windows.Forms.Padding(4);
            panel3.Name = "panel3";
            panel3.Size = new System.Drawing.Size(143, 257);
            panel3.TabIndex = 27;
            // 
            // label16
            // 
            label16.AutoSize = true;
            label16.Location = new System.Drawing.Point(4, 186);
            label16.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label16.Name = "label16";
            label16.Size = new System.Drawing.Size(40, 17);
            label16.TabIndex = 37;
            label16.Text = "Sheet";
            // 
            // label17
            // 
            label17.AutoSize = true;
            label17.Location = new System.Drawing.Point(4, 212);
            label17.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label17.Name = "label17";
            label17.Size = new System.Drawing.Size(32, 17);
            label17.TabIndex = 36;
            label17.Text = "比较";
            // 
            // label11
            // 
            label11.AutoSize = true;
            label11.Cursor = System.Windows.Forms.Cursors.SizeAll;
            label11.Location = new System.Drawing.Point(2, 123);
            label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label11.Name = "label11";
            label11.Size = new System.Drawing.Size(44, 17);
            label11.TabIndex = 31;
            label11.Text = "列位置";
            // 
            // label18
            // 
            label18.AutoSize = true;
            label18.Location = new System.Drawing.Point(5, 159);
            label18.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label18.Name = "label18";
            label18.Size = new System.Drawing.Size(32, 17);
            label18.TabIndex = 35;
            label18.Text = "上下";
            // 
            // TitleCol2
            // 
            TitleCol2.Location = new System.Drawing.Point(55, 118);
            TitleCol2.Margin = new System.Windows.Forms.Padding(4);
            TitleCol2.Name = "TitleCol2";
            TitleCol2.ReadOnly = true;
            TitleCol2.Size = new System.Drawing.Size(60, 23);
            TitleCol2.TabIndex = 32;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new System.Drawing.Point(471, 110);
            label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label4.Name = "label4";
            label4.Size = new System.Drawing.Size(64, 17);
            label4.TabIndex = 28;
            label4.Text = "Sheet信息";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new System.Drawing.Point(472, 507);
            label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label5.Name = "label5";
            label5.Size = new System.Drawing.Size(64, 17);
            label5.TabIndex = 29;
            label5.Text = "Sheet信息";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new System.Drawing.Point(603, 119);
            label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label2.Name = "label2";
            label2.Size = new System.Drawing.Size(88, 17);
            label2.TabIndex = 30;
            label2.Text = "批量Sheet操作";
            // 
            // label12
            // 
            label12.AutoSize = true;
            label12.Location = new System.Drawing.Point(603, 234);
            label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label12.Name = "label12";
            label12.Size = new System.Drawing.Size(0, 17);
            label12.TabIndex = 31;
            // 
            // CreateDataTable
            // 
            CreateDataTable.Location = new System.Drawing.Point(474, 350);
            CreateDataTable.Margin = new System.Windows.Forms.Padding(4);
            CreateDataTable.Name = "CreateDataTable";
            CreateDataTable.Size = new System.Drawing.Size(115, 34);
            CreateDataTable.TabIndex = 35;
            CreateDataTable.Text = "创建 DataTable";
            CreateDataTable.UseVisualStyleBackColor = true;
            CreateDataTable.Click += CreateDataTable_Click;
            // 
            // diaplayRowAndColumn
            // 
            diaplayRowAndColumn.Location = new System.Drawing.Point(602, 227);
            diaplayRowAndColumn.Margin = new System.Windows.Forms.Padding(4);
            diaplayRowAndColumn.Name = "diaplayRowAndColumn";
            diaplayRowAndColumn.Size = new System.Drawing.Size(114, 35);
            diaplayRowAndColumn.TabIndex = 36;
            diaplayRowAndColumn.Text = "显示所有行和列";
            diaplayRowAndColumn.UseVisualStyleBackColor = true;
            diaplayRowAndColumn.Click += diaplayRowAndColumn_Click;
            // 
            // dataGridViewTextBoxColumn1
            // 
            dataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewTextBoxColumn1.Frozen = true;
            dataGridViewTextBoxColumn1.HeaderText = "序号";
            dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            dataGridViewTextBoxColumn1.ReadOnly = true;
            dataGridViewTextBoxColumn1.Width = 60;
            // 
            // dataGridViewTextBoxColumn2
            // 
            dataGridViewTextBoxColumn2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewTextBoxColumn2.Frozen = true;
            dataGridViewTextBoxColumn2.HeaderText = "名字";
            dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            dataGridViewTextBoxColumn2.ReadOnly = true;
            dataGridViewTextBoxColumn2.Width = 57;
            // 
            // dataGridViewTextBoxColumn3
            // 
            dataGridViewTextBoxColumn3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewTextBoxColumn3.HeaderText = "标题行";
            dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            // 
            // Column1
            // 
            Column1.HeaderText = "标题列";
            Column1.Name = "Column1";
            // 
            // dataGridViewTextBoxColumn4
            // 
            dataGridViewTextBoxColumn4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewTextBoxColumn4.Frozen = true;
            dataGridViewTextBoxColumn4.HeaderText = "序号";
            dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            dataGridViewTextBoxColumn4.ReadOnly = true;
            dataGridViewTextBoxColumn4.Width = 57;
            // 
            // dataGridViewTextBoxColumn5
            // 
            dataGridViewTextBoxColumn5.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewTextBoxColumn5.Frozen = true;
            dataGridViewTextBoxColumn5.HeaderText = "名字";
            dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            dataGridViewTextBoxColumn5.ReadOnly = true;
            dataGridViewTextBoxColumn5.Width = 57;
            // 
            // dataGridViewTextBoxColumn6
            // 
            dataGridViewTextBoxColumn6.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewTextBoxColumn6.HeaderText = "标题行";
            dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            dataGridViewTextBoxColumn6.Width = 69;
            // 
            // Column2
            // 
            Column2.HeaderText = "标题列";
            Column2.Name = "Column2";
            // 
            // Form1
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(722, 786);
            Controls.Add(diaplayRowAndColumn);
            Controls.Add(CreateDataTable);
            Controls.Add(label12);
            Controls.Add(label2);
            Controls.Add(GenerateConfigurationCode);
            Controls.Add(DelHiddenWs);
            Controls.Add(label4);
            Controls.Add(label5);
            Controls.Add(filePath2);
            Controls.Add(filePath1);
            Controls.Add(dgv2);
            Controls.Add(BtnAnalyze2);
            Controls.Add(dgv1);
            Controls.Add(BtnAnalyze1);
            Controls.Add(label8);
            Controls.Add(wsNameOrIndex2);
            Controls.Add(label7);
            Controls.Add(SelectFileBtn2);
            Controls.Add(label3);
            Controls.Add(SelectFileBtn1);
            Controls.Add(label1);
            Controls.Add(panel2);
            Controls.Add(panel3);
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            Margin = new System.Windows.Forms.Padding(4);
            Name = "Form1";
            Text = "EPPlusTool Owner By GeLiang ";
            ((System.ComponentModel.ISupportInitialize)dgv2).EndInit();
            ((System.ComponentModel.ISupportInitialize)dgv1).EndInit();
            panel2.ResumeLayout(false);
            panel2.PerformLayout();
            panel3.ResumeLayout(false);
            panel3.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
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
        private System.Windows.Forms.DataGridView dgv1;
        private System.Windows.Forms.Button GenerateConfiguration;
        private System.Windows.Forms.Button GenerateConfigurationCode;
        private System.Windows.Forms.Button DelHiddenWs;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label5;
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
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
    }
}

