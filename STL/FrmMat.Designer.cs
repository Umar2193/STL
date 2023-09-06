namespace STL {
    partial class FrmMat {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if(disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.rbtLeq = new System.Windows.Forms.RadioButton();
            this.rbtTeq = new System.Windows.Forms.RadioButton();
            this.btnAddToPlanning = new System.Windows.Forms.Button();
            this.dgvPlanning = new System.Windows.Forms.DataGridView();
            this.cmbComponent = new System.Windows.Forms.ComboBox();
            this.cmbType = new System.Windows.Forms.ComboBox();
            this.chk4x2 = new System.Windows.Forms.CheckBox();
            this.chk5x2 = new System.Windows.Forms.CheckBox();
            this.chk5x = new System.Windows.Forms.CheckBox();
            this.chk4x = new System.Windows.Forms.CheckBox();
            this.chk3x = new System.Windows.Forms.CheckBox();
            this.chk2x2 = new System.Windows.Forms.CheckBox();
            this.chk2x1 = new System.Windows.Forms.CheckBox();
            this.chk2x = new System.Windows.Forms.CheckBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.lblSearch = new System.Windows.Forms.Label();
            this.lblMessage = new System.Windows.Forms.Label();
            this.dgvSource = new System.Windows.Forms.DataGridView();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.dgvRental = new System.Windows.Forms.DataGridView();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.dgvConsumables = new System.Windows.Forms.DataGridView();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.lblBoltError = new System.Windows.Forms.Label();
            this.btnCopy = new System.Windows.Forms.Button();
            this.txtCopy = new System.Windows.Forms.TextBox();
            this.lblResult = new System.Windows.Forms.Label();
            this.btnFind = new System.Windows.Forms.Button();
            this.txtSearchValue = new System.Windows.Forms.TextBox();
            this.lblEnter = new System.Windows.Forms.Label();
            this.btnExcel = new System.Windows.Forms.Button();
            this.btnBackToProject = new System.Windows.Forms.Button();
            this.lblProject = new System.Windows.Forms.Label();
            this.btnSubmit = new System.Windows.Forms.Button();
            this.btnCopyCockpit = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPlanning)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSource)).BeginInit();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvRental)).BeginInit();
            this.tabPage4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvConsumables)).BeginInit();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(12, 35);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1679, 680);
            this.tabControl1.TabIndex = 15;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.rbtLeq);
            this.tabPage1.Controls.Add(this.rbtTeq);
            this.tabPage1.Controls.Add(this.btnAddToPlanning);
            this.tabPage1.Controls.Add(this.dgvPlanning);
            this.tabPage1.Controls.Add(this.cmbComponent);
            this.tabPage1.Controls.Add(this.cmbType);
            this.tabPage1.Controls.Add(this.chk4x2);
            this.tabPage1.Controls.Add(this.chk5x2);
            this.tabPage1.Controls.Add(this.chk5x);
            this.tabPage1.Controls.Add(this.chk4x);
            this.tabPage1.Controls.Add(this.chk3x);
            this.tabPage1.Controls.Add(this.chk2x2);
            this.tabPage1.Controls.Add(this.chk2x1);
            this.tabPage1.Controls.Add(this.chk2x);
            this.tabPage1.Controls.Add(this.textBox1);
            this.tabPage1.Controls.Add(this.lblSearch);
            this.tabPage1.Controls.Add(this.lblMessage);
            this.tabPage1.Controls.Add(this.dgvSource);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabPage1.Size = new System.Drawing.Size(1671, 651);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "LEQ / TEQ";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // rbtLeq
            // 
            this.rbtLeq.AutoSize = true;
            this.rbtLeq.Checked = true;
            this.rbtLeq.Location = new System.Drawing.Point(92, 66);
            this.rbtLeq.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.rbtLeq.Name = "rbtLeq";
            this.rbtLeq.Size = new System.Drawing.Size(54, 20);
            this.rbtLeq.TabIndex = 38;
            this.rbtLeq.TabStop = true;
            this.rbtLeq.Text = "LEQ";
            this.rbtLeq.UseVisualStyleBackColor = true;
            this.rbtLeq.CheckedChanged += new System.EventHandler(this.rbtLeq_CheckedChanged);
            // 
            // rbtTeq
            // 
            this.rbtTeq.AutoSize = true;
            this.rbtTeq.Location = new System.Drawing.Point(167, 66);
            this.rbtTeq.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.rbtTeq.Name = "rbtTeq";
            this.rbtTeq.Size = new System.Drawing.Size(56, 20);
            this.rbtTeq.TabIndex = 36;
            this.rbtTeq.Text = "TEQ";
            this.rbtTeq.UseVisualStyleBackColor = true;
            this.rbtTeq.CheckedChanged += new System.EventHandler(this.rbtTeq_CheckedChanged);
            // 
            // btnAddToPlanning
            // 
            this.btnAddToPlanning.Location = new System.Drawing.Point(487, 101);
            this.btnAddToPlanning.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnAddToPlanning.Name = "btnAddToPlanning";
            this.btnAddToPlanning.Size = new System.Drawing.Size(152, 23);
            this.btnAddToPlanning.TabIndex = 35;
            this.btnAddToPlanning.Text = "Add To Planning";
            this.btnAddToPlanning.UseVisualStyleBackColor = true;
            this.btnAddToPlanning.Click += new System.EventHandler(this.btnAddToPlanning_Click);
            // 
            // dgvPlanning
            // 
            this.dgvPlanning.AllowUserToAddRows = false;
            this.dgvPlanning.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvPlanning.ColumnHeadersVisible = false;
            this.dgvPlanning.Location = new System.Drawing.Point(723, 129);
            this.dgvPlanning.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dgvPlanning.Name = "dgvPlanning";
            this.dgvPlanning.RowHeadersVisible = false;
            this.dgvPlanning.RowHeadersWidth = 51;
            this.dgvPlanning.RowTemplate.Height = 29;
            this.dgvPlanning.Size = new System.Drawing.Size(942, 447);
            this.dgvPlanning.TabIndex = 34;
            this.dgvPlanning.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvPlanning_CellClick);
            this.dgvPlanning.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dgvPlanning_CellPainting);
            this.dgvPlanning.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvPlanning_CellValueChanged);
            // 
            // cmbComponent
            // 
            this.cmbComponent.FormattingEnabled = true;
            this.cmbComponent.Location = new System.Drawing.Point(588, 62);
            this.cmbComponent.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.cmbComponent.Name = "cmbComponent";
            this.cmbComponent.Size = new System.Drawing.Size(207, 24);
            this.cmbComponent.TabIndex = 33;
            this.cmbComponent.Text = "Search by Component";
            this.cmbComponent.SelectedIndexChanged += new System.EventHandler(this.cmbComponent_SelectedIndexChanged);
            // 
            // cmbType
            // 
            this.cmbType.FormattingEnabled = true;
            this.cmbType.Location = new System.Drawing.Point(380, 62);
            this.cmbType.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.cmbType.Name = "cmbType";
            this.cmbType.Size = new System.Drawing.Size(202, 24);
            this.cmbType.TabIndex = 32;
            this.cmbType.Text = "Search by Type";
            this.cmbType.SelectedIndexChanged += new System.EventHandler(this.cmbType_SelectedIndexChanged);
            // 
            // chk4x2
            // 
            this.chk4x2.AutoSize = true;
            this.chk4x2.Location = new System.Drawing.Point(946, 39);
            this.chk4x2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.chk4x2.Name = "chk4x2";
            this.chk4x2.Size = new System.Drawing.Size(91, 20);
            this.chk4x2.TabIndex = 31;
            this.chk4x2.Text = "4X SG-145";
            this.chk4x2.UseVisualStyleBackColor = true;
            this.chk4x2.CheckedChanged += new System.EventHandler(this.chk2x_CheckedChanged);
            // 
            // chk5x2
            // 
            this.chk5x2.AutoSize = true;
            this.chk5x2.Location = new System.Drawing.Point(1158, 39);
            this.chk5x2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.chk5x2.Name = "chk5x2";
            this.chk5x2.Size = new System.Drawing.Size(91, 20);
            this.chk5x2.TabIndex = 26;
            this.chk5x2.Text = "5X SG-170";
            this.chk5x2.UseVisualStyleBackColor = true;
            this.chk5x2.CheckedChanged += new System.EventHandler(this.chk2x_CheckedChanged);
            // 
            // chk5x
            // 
            this.chk5x.AutoSize = true;
            this.chk5x.Location = new System.Drawing.Point(1052, 38);
            this.chk5x.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.chk5x.Name = "chk5x";
            this.chk5x.Size = new System.Drawing.Size(91, 20);
            this.chk5x.TabIndex = 25;
            this.chk5x.Text = "5X SG-155";
            this.chk5x.UseVisualStyleBackColor = true;
            this.chk5x.CheckedChanged += new System.EventHandler(this.chk2x_CheckedChanged);
            // 
            // chk4x
            // 
            this.chk4x.AutoSize = true;
            this.chk4x.Location = new System.Drawing.Point(840, 39);
            this.chk4x.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.chk4x.Name = "chk4x";
            this.chk4x.Size = new System.Drawing.Size(91, 20);
            this.chk4x.TabIndex = 24;
            this.chk4x.Text = "4X SG-132";
            this.chk4x.UseVisualStyleBackColor = true;
            this.chk4x.CheckedChanged += new System.EventHandler(this.chk2x_CheckedChanged);
            // 
            // chk3x
            // 
            this.chk3x.AutoSize = true;
            this.chk3x.Location = new System.Drawing.Point(734, 39);
            this.chk3x.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.chk3x.Name = "chk3x";
            this.chk3x.Size = new System.Drawing.Size(91, 20);
            this.chk3x.TabIndex = 23;
            this.chk3x.Text = "3X SG-132";
            this.chk3x.UseVisualStyleBackColor = true;
            this.chk3x.CheckedChanged += new System.EventHandler(this.chk2x_CheckedChanged);
            // 
            // chk2x2
            // 
            this.chk2x2.AutoSize = true;
            this.chk2x2.Location = new System.Drawing.Point(628, 39);
            this.chk2x2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.chk2x2.Name = "chk2x2";
            this.chk2x2.Size = new System.Drawing.Size(91, 20);
            this.chk2x2.TabIndex = 22;
            this.chk2x2.Text = "2X SG-126";
            this.chk2x2.UseVisualStyleBackColor = true;
            this.chk2x2.CheckedChanged += new System.EventHandler(this.chk2x_CheckedChanged);
            // 
            // chk2x1
            // 
            this.chk2x1.AutoSize = true;
            this.chk2x1.Location = new System.Drawing.Point(503, 39);
            this.chk2x1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.chk2x1.Name = "chk2x1";
            this.chk2x1.Size = new System.Drawing.Size(108, 20);
            this.chk2x1.TabIndex = 21;
            this.chk2x1.Text = "2X SG2.5-114";
            this.chk2x1.UseVisualStyleBackColor = true;
            this.chk2x1.CheckedChanged += new System.EventHandler(this.chk2x_CheckedChanged);
            // 
            // chk2x
            // 
            this.chk2x.AutoSize = true;
            this.chk2x.Location = new System.Drawing.Point(378, 39);
            this.chk2x.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.chk2x.Name = "chk2x";
            this.chk2x.Size = new System.Drawing.Size(108, 20);
            this.chk2x.TabIndex = 20;
            this.chk2x.Text = "2X SG2.1-114";
            this.chk2x.UseVisualStyleBackColor = true;
            this.chk2x.CheckedChanged += new System.EventHandler(this.chk2x_CheckedChanged);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(91, 37);
            this.textBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(255, 22);
            this.textBox1.TabIndex = 18;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // lblSearch
            // 
            this.lblSearch.AutoSize = true;
            this.lblSearch.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblSearch.Location = new System.Drawing.Point(18, 38);
            this.lblSearch.Name = "lblSearch";
            this.lblSearch.Size = new System.Drawing.Size(67, 20);
            this.lblSearch.TabIndex = 17;
            this.lblSearch.Text = "Search : ";
            // 
            // lblMessage
            // 
            this.lblMessage.AutoSize = true;
            this.lblMessage.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblMessage.Location = new System.Drawing.Point(92, 10);
            this.lblMessage.Name = "lblMessage";
            this.lblMessage.Size = new System.Drawing.Size(51, 20);
            this.lblMessage.TabIndex = 16;
            this.lblMessage.Text = "label1";
            this.lblMessage.Visible = false;
            // 
            // dgvSource
            // 
            this.dgvSource.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSource.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgvSource.Location = new System.Drawing.Point(16, 129);
            this.dgvSource.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dgvSource.Name = "dgvSource";
            this.dgvSource.RowHeadersVisible = false;
            this.dgvSource.RowHeadersWidth = 51;
            this.dgvSource.RowTemplate.Height = 29;
            this.dgvSource.Size = new System.Drawing.Size(701, 447);
            this.dgvSource.TabIndex = 15;
            this.dgvSource.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dgvSource_CellBeginEdit);
            this.dgvSource.CellToolTipTextNeeded += new System.Windows.Forms.DataGridViewCellToolTipTextNeededEventHandler(this.dgvSource_CellToolTipTextNeeded);
            this.dgvSource.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvSource_CellValueChanged);
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dgvRental);
            this.tabPage3.Location = new System.Drawing.Point(4, 25);
            this.tabPage3.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(1671, 651);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Rental";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // dgvRental
            // 
            this.dgvRental.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvRental.Location = new System.Drawing.Point(3, 2);
            this.dgvRental.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dgvRental.Name = "dgvRental";
            this.dgvRental.RowHeadersVisible = false;
            this.dgvRental.RowHeadersWidth = 51;
            this.dgvRental.RowTemplate.Height = 29;
            this.dgvRental.Size = new System.Drawing.Size(1315, 538);
            this.dgvRental.TabIndex = 16;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.dgvConsumables);
            this.tabPage4.Location = new System.Drawing.Point(4, 25);
            this.tabPage4.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(1671, 651);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Consumables";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // dgvConsumables
            // 
            this.dgvConsumables.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvConsumables.Location = new System.Drawing.Point(3, 2);
            this.dgvConsumables.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dgvConsumables.Name = "dgvConsumables";
            this.dgvConsumables.RowHeadersVisible = false;
            this.dgvConsumables.RowHeadersWidth = 51;
            this.dgvConsumables.RowTemplate.Height = 29;
            this.dgvConsumables.Size = new System.Drawing.Size(1312, 608);
            this.dgvConsumables.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.lblBoltError);
            this.tabPage2.Controls.Add(this.btnCopy);
            this.tabPage2.Controls.Add(this.txtCopy);
            this.tabPage2.Controls.Add(this.lblResult);
            this.tabPage2.Controls.Add(this.btnFind);
            this.tabPage2.Controls.Add(this.txtSearchValue);
            this.tabPage2.Controls.Add(this.lblEnter);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabPage2.Size = new System.Drawing.Size(1671, 651);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "GP - A9B bolt";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // lblBoltError
            // 
            this.lblBoltError.AutoSize = true;
            this.lblBoltError.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblBoltError.Location = new System.Drawing.Point(222, 136);
            this.lblBoltError.Name = "lblBoltError";
            this.lblBoltError.Size = new System.Drawing.Size(51, 20);
            this.lblBoltError.TabIndex = 6;
            this.lblBoltError.Text = "label1";
            this.lblBoltError.Visible = false;
            // 
            // btnCopy
            // 
            this.btnCopy.Location = new System.Drawing.Point(594, 91);
            this.btnCopy.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(128, 23);
            this.btnCopy.TabIndex = 5;
            this.btnCopy.Text = "Copy";
            this.btnCopy.UseVisualStyleBackColor = true;
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
            // 
            // txtCopy
            // 
            this.txtCopy.Location = new System.Drawing.Point(291, 91);
            this.txtCopy.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtCopy.Name = "txtCopy";
            this.txtCopy.Size = new System.Drawing.Size(250, 22);
            this.txtCopy.TabIndex = 4;
            // 
            // lblResult
            // 
            this.lblResult.AutoSize = true;
            this.lblResult.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblResult.Location = new System.Drawing.Point(212, 94);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(61, 20);
            this.lblResult.TabIndex = 3;
            this.lblResult.Text = "Result :";
            // 
            // btnFind
            // 
            this.btnFind.Location = new System.Drawing.Point(594, 42);
            this.btnFind.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnFind.Name = "btnFind";
            this.btnFind.Size = new System.Drawing.Size(128, 23);
            this.btnFind.TabIndex = 2;
            this.btnFind.Text = "Find";
            this.btnFind.UseVisualStyleBackColor = true;
            this.btnFind.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtSearchValue
            // 
            this.txtSearchValue.Location = new System.Drawing.Point(291, 44);
            this.txtSearchValue.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtSearchValue.Name = "txtSearchValue";
            this.txtSearchValue.Size = new System.Drawing.Size(250, 22);
            this.txtSearchValue.TabIndex = 1;
            // 
            // lblEnter
            // 
            this.lblEnter.AutoSize = true;
            this.lblEnter.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblEnter.Location = new System.Drawing.Point(51, 46);
            this.lblEnter.Name = "lblEnter";
            this.lblEnter.Size = new System.Drawing.Size(222, 20);
            this.lblEnter.TabIndex = 0;
            this.lblEnter.Text = "Enter SAP Material / GP Code :";
            // 
            // btnExcel
            // 
            this.btnExcel.Location = new System.Drawing.Point(165, 8);
            this.btnExcel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(136, 23);
            this.btnExcel.TabIndex = 27;
            this.btnExcel.Text = "Save to Excel";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // btnBackToProject
            // 
            this.btnBackToProject.Location = new System.Drawing.Point(12, 8);
            this.btnBackToProject.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnBackToProject.Name = "btnBackToProject";
            this.btnBackToProject.Size = new System.Drawing.Size(147, 23);
            this.btnBackToProject.TabIndex = 28;
            this.btnBackToProject.Text = "Back to Projects";
            this.btnBackToProject.UseVisualStyleBackColor = true;
            this.btnBackToProject.Click += new System.EventHandler(this.btnBackToProject_Click);
            // 
            // lblProject
            // 
            this.lblProject.AutoSize = true;
            this.lblProject.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblProject.Location = new System.Drawing.Point(451, 11);
            this.lblProject.Name = "lblProject";
            this.lblProject.Size = new System.Drawing.Size(51, 20);
            this.lblProject.TabIndex = 29;
            this.lblProject.Text = "label1";
            // 
            // btnSubmit
            // 
            this.btnSubmit.Location = new System.Drawing.Point(307, 8);
            this.btnSubmit.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.Size = new System.Drawing.Size(136, 23);
            this.btnSubmit.TabIndex = 30;
            this.btnSubmit.Text = "Submit to MM";
            this.btnSubmit.UseVisualStyleBackColor = true;
            this.btnSubmit.Click += new System.EventHandler(this.btnSubmit_Click);
            // 
            // btnCopyCockpit
            // 
            this.btnCopyCockpit.Location = new System.Drawing.Point(1107, 8);
            this.btnCopyCockpit.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnCopyCockpit.Name = "btnCopyCockpit";
            this.btnCopyCockpit.Size = new System.Drawing.Size(136, 23);
            this.btnCopyCockpit.TabIndex = 31;
            this.btnCopyCockpit.Text = "Copy to Cockpit";
            this.btnCopyCockpit.UseVisualStyleBackColor = true;
            this.btnCopyCockpit.Visible = false;
            this.btnCopyCockpit.Click += new System.EventHandler(this.btnCopyCockpit_Click);
            // 
            // FrmMat
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1701, 787);
            this.Controls.Add(this.btnCopyCockpit);
            this.Controls.Add(this.btnSubmit);
            this.Controls.Add(this.lblProject);
            this.Controls.Add(this.btnBackToProject);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.btnExcel);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "FrmMat";
            this.Text = "Materials";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.FrmMat_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPlanning)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSource)).EndInit();
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvRental)).EndInit();
            this.tabPage4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvConsumables)).EndInit();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Button btnDownload;
        private TabControl tabControl1;
        private TabPage tabPage1;
        private Button btnExcel;
        private CheckBox chk5x2;
        private CheckBox chk5x;
        private CheckBox chk4x;
        private CheckBox chk3x;
        private CheckBox chk2x2;
        private CheckBox chk2x1;
        private CheckBox chk2x;
        private TextBox textBox1;
        private Label lblSearch;
        private Label lblMessage;
        private DataGridView dgvSource;
        private TabPage tabPage2;
        private Button btnFind;
        private TextBox txtSearchValue;
        private Label lblEnter;
        private Button btnCopy;
        private TextBox txtCopy;
        private Label lblResult;
        private Label lblBoltError;
        private CheckBox chk4x2;
        private TabPage tabPage3;
        private DataGridView dgvRental;
        private ComboBox cmbComponent;
        private ComboBox cmbType;
        private Button btnBackToProject;
        private Label lblProject;
        private DataGridView dgvPlanning;
        private Button btnAddToPlanning;
        private TabPage tabPage4;
        private DataGridView dgvConsumables;
        private RadioButton rbtLeq;
        private RadioButton rbtTeq;
        private Button btnSubmit;
        private Button btnCopyCockpit;
    }
}