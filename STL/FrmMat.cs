using ClosedXML.Excel;
using Newtonsoft.Json;
using STL.Helpers;
using STL.Models;
using System.Data;
using System.Text.RegularExpressions;

namespace STL {
    public partial class FrmMat : Form {
        private readonly List<string> sourceColumns = new List<string>() { "MatNo", "Object description", "quantity", "Package", "Comment", "2X SG2.1-114", "2X SG2.5-114", "2X SG-126", "3X SG-132", "4X SG-132", "4X SG-145", "5X SG-155", "5X SG-170", "Type", "Component" };
        private readonly List<string> sourceDisplayColumns = new List<string>() { "MatNo", "Object description", "quantity", "Package" };
        private readonly string[] editableSourceColumns = new string[] { "DeliveryDestination", };
        private readonly List<string> rentalColumns = new List<string>() { "MatNo", "QTY", "MatDescription", "RentalCharges", "Package" };
        private readonly List<string> consumablesColumns = new List<string>() { "MatNo", "QTY", "MatDescription", "RentalCharges", "Type" };
        private readonly string[] reportColumns = new string[] { "MatNo", "QTY", "MatDescription", "Customer material", "Demand date ex warehouse", "Date of submission", "EQM Ownership", "Voyage ref.", "Delivery Destination", "Contact person and Details for Delivery", "Comment", "Project Phase", "Reason for late order", "Return date" };
        private readonly string basePath = "";
        private readonly string project = "";
        List<string> types = new List<string>();
        List<string> componentlist = new List<string>();
        List<DeliveryDestination> deliveryDestinations = new List<DeliveryDestination>();
        DataTable dtReport = new DataTable();
        DataTable dtSource = new DataTable();
        Dictionary<string, string> filters = new Dictionary<string, string>();
        List<RentalRate> RentalRates = new();
        Config config = new Config();
        List<string> bladeSources = new List<string>();
        List<string> towerSources = new List<string>();
        List<string> mainSources = new List<string>();
        List<PlanningMaterial> planningMaterial = new List<PlanningMaterial>();
        private string upid = "";
        private string executionPlant = "";
        private string storageLocation = "";
        private string wbsLEQ = "";
        private string wbsTEQ = "";
        private DateTime StartDate = new DateTime();
        private DateTime EndDate = new DateTime();

        private readonly Dictionary<string, string> custMaterial = new Dictionary<string, string>() {
            { "Please Choose", string.Empty},
            { "Mains TEQ", "Mains TEQ" } ,
            { "Blade TEQ", "Blade TEQ" } ,
            { "Tower TEQ", "Tower TEQ" } ,
            { "LEQ", "LEQ" } ,
            { "Others", "Others" } ,
        };

        public FrmMat(string basePath, string project) {
            InitializeComponent();
            this.basePath = basePath;
            this.project = project;
            this.Text = project;
            if(File.Exists("config.json")) {
                using StreamReader r = new StreamReader("config.json");
                string json = r.ReadToEnd();
                config = JsonConvert.DeserializeObject<Config>(json);
                if(config != null) {
                    custMaterial = config.CustomerMaterial;
                }
            }
        }

        private void FrmMat_Load(object sender, EventArgs e) {
            try {
                LoadProjectInfo();
                LoadData("LEQ");
                LoadRentalRates();
            } catch(Exception ex) {
                Logger.Write($"[{ex.Message}] - {ex.StackTrace}");
                MessageBox.Show("Something went wrong. Close the project report file if it is open.", "Error");
            }
        }

        private void LoadData(string name) {
            try {
                //Clear all filters
                chk2x.Checked = false;
                chk2x1.Checked = false;
                chk2x2.Checked = false;
                chk3x.Checked = false;
                chk4x.Checked = false;
                chk4x2.Checked = false;
                chk5x.Checked = false;
                chk5x2.Checked = false;
                cmbComponent.SelectedIndex = -1;
                cmbType.SelectedIndex = -1;

                LoadPlanningExcel($"{name} Overview");
                LoadSourceExcel($"STL {name}");
                LoadRentalExcel($"{name} Rental", dgvRental, rentalColumns);
                LoadRentalExcel($"{name} Consumables", dgvConsumables, consumablesColumns);

                LoadReportExcel($"{name} Order");
            } catch(Exception ex) {
                Logger.Write($"[{ex.Message}] - {ex.StackTrace}");
                MessageBox.Show("Something went wrong. Close the project report file if it is open.", "Error");
            }
        }

        private void LoadProjectInfo() {
            var workbook = new XLWorkbook($"{basePath}\\{project}\\{project}_Report.xlsm");

            var projWorkSheet = workbook?.Worksheets?.FirstOrDefault(c => c.Name == "Project Info");
            if(projWorkSheet != null) {
                StartDate = DateTime.Parse(projWorkSheet.Cell(3, 3)?.Value?.ToString());
                EndDate = DateTime.Parse(projWorkSheet.Cell(4, 3)?.Value?.ToString());
                lblProject.Text = $"Project : {projWorkSheet.Cell(5, 3).Value} (From {StartDate:yyyy-MM-dd} To {EndDate:yyyy-MM-dd})";

                upid = Convert.ToString(projWorkSheet.Cell(27, 3).Value ?? string.Empty);
                storageLocation = Convert.ToString(projWorkSheet.Cell(27, 6).Value ?? string.Empty);
                wbsLEQ = Convert.ToString(projWorkSheet.Cell(28, 3).Value ?? string.Empty);
                wbsTEQ = Convert.ToString(projWorkSheet.Cell(29, 3).Value ?? string.Empty);
                executionPlant = Convert.ToString(projWorkSheet.Cell(30, 3).Value ?? string.Empty);

                for(int i = 16; i < 24; i++) {
                    if(!string.IsNullOrEmpty(projWorkSheet.Cell(i, 6).Value?.ToString())) {
                        deliveryDestinations.Add(new DeliveryDestination() {
                            Name = projWorkSheet.Cell(i, 6).Value?.ToString(),
                            PersonName = projWorkSheet.Cell(i, 12).Value?.ToString(),
                            Address = projWorkSheet.Cell(i, 11).Value?.ToString(),
                            WBS = projWorkSheet.Cell(i, 8).Value?.ToString(),
                            ExecutionPlant = projWorkSheet.Cell(i, 7).Value?.ToString(),
                            FOPO = projWorkSheet.Cell(i, 9).Value?.ToString(),
                            Incoterm = projWorkSheet.Cell(i, 10).Value?.ToString(),
                            IncotermLocation = projWorkSheet.Cell(i, 11).Value?.ToString(),
                        });
                    }
                }

                bladeSources.AddRange(new List<string>() { projWorkSheet.Cell(17, 4)?.Value?.ToString(), projWorkSheet.Cell(21, 4)?.Value?.ToString() });
                towerSources.AddRange(new List<string>() { projWorkSheet.Cell(18, 4)?.Value?.ToString(), projWorkSheet.Cell(22, 4)?.Value?.ToString() });
                mainSources.AddRange(new List<string>() { projWorkSheet.Cell(16, 4)?.Value?.ToString(), projWorkSheet.Cell(20, 4)?.Value?.ToString() });
            }
        }

        private void LoadPlanningExcel(string sheetName) {
            dgvPlanning.Columns.Clear();

            var workbook = new XLWorkbook($"{basePath}\\{project}\\{project}_Report.xlsm");
            var dt = new DataTable();
            dt.Columns.Add("Item");
            dt.Columns.Add("Total");

            foreach(var week in CommonHelper.WeekList(StartDate, EndDate)) {
                dt.Columns.Add(week);
            }

            var worksheet = workbook?.Worksheets?.FirstOrDefault(c => c.Name == sheetName);
            if(worksheet != null) {
                var startIndex = 1;
                if(worksheet.Cell("A1").Value.ToString().ToLower().Contains("total ")) {
                    startIndex = 3;
                }
                foreach(IXLRow row in worksheet.Rows(startIndex, worksheet?.LastRowUsed()?.RowNumber() ?? 0)) {
                    dt.Rows.Add();
                    int i = 0;
                    foreach(IXLCell cell in row.Cells(true)) {
                        try {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                        } catch { }
                        i++;
                    }
                }
            }

            dgvPlanning.DataSource = dt;

            DataGridViewButtonColumn btnDeleteButton = new();
            btnDeleteButton.Name = "Delete";
            btnDeleteButton.Text = "Delete";
            btnDeleteButton.UseColumnTextForButtonValue = true;

            dgvPlanning.Columns.Insert(dgvPlanning.Columns.Count, btnDeleteButton);

            PaintPlanningGrid(dt);
        }

        private void PaintPlanningGrid(DataTable dt) {
            for(int i = 0; i < dt.Rows.Count; i++) {
                int rowToSelect = i - (i % 7);

                if(i % 7 == 0) {
                    rowToSelect = i;
                }

                dgvPlanning.Rows[i].Tag = dgvPlanning.Rows[rowToSelect].Cells[0].Value;
                if((i + 1) % 7 == 1 || (i + 1) % 7 == 4 || (i + 1) % 7 == 5) {
                    DisableRow(i);
                } else if((i + 1) % 7 == 6 || (i + 1) % 7 == 0) {
                    DisableRow(i, false);
                } else {
                    dgvPlanning.Rows[i].Cells[0].Style.BackColor = Color.LightGray;
                    dgvPlanning.Rows[i].Cells[0].ReadOnly = true;
                    dgvPlanning.Rows[i].Cells[1].Style.BackColor = Color.LightGray;
                    dgvPlanning.Rows[i].Cells[1].ReadOnly = true;

                }
                if((i + 1) % 7 == 4) {
                    dgvPlanning.Rows[i].Cells[0].ReadOnly = false;
                    dgvPlanning.Rows[i].Cells[0].Style.BackColor = Color.White;
                }
            }
        }

        private void DisableRow(int index, bool shouldApplyColor = true) {
            dgvPlanning.Rows[index].ReadOnly = true;
            if(shouldApplyColor) {
                dgvPlanning.Rows[index].DefaultCellStyle.BackColor = Color.LightGray;
            }
        }

        private static object[] GetItem(string name, List<string> items, string total) {
            var inrow = new List<object> { name, total };
            inrow.AddRange(items.Select(c => c.ToString()).ToList());
            return inrow.ToArray();
        }

        private void LoadSourceExcel(string name) {
            dgvSource.Columns.Clear();
            dtSource = new DataTable();

            lblMessage.Visible = true;
            lblMessage.Text = "Loading...";

            using(var stream = new FileStream($"{basePath}\\Source.xlsx", FileMode.Open, FileAccess.Read)) {
                using(var workbook = new XLWorkbook(stream)) {

                    if(workbook == null) {
                        dgvSource.Visible = false;
                        lblMessage.Visible = true;
                        lblMessage.Text = "Source file does not exists. Please load correct excel file in the base location.";
                    }
                    var worksheet = workbook?.Worksheets?.FirstOrDefault(c => c.Name == name);

                    if(worksheet == null) {
                        dgvSource.Visible = false;
                        lblMessage.Visible = true;
                        lblMessage.Text = "Invalid source file. Please load correct excel file in the base location.";
                    } else {
                        lblMessage.Visible = false;
                        GetSourceData(worksheet, workbook, name);
                        dgvSource.DataSource = dtSource;

                        dgvSource.Visible = true;
                        SetSourceColumns();
                    }
                }
            }
        }

        private void LoadReportExcel(string sheetName) {
            dtReport.Rows.Clear();
            lblMessage.Visible = true;
            lblMessage.Text = "Loading...";
            var workbook = new XLWorkbook($"{basePath}\\{project}\\{project}_Report.xlsm");
            if(workbook == null) {
                dgvSource.Visible = false;
                lblMessage.Visible = true;
                lblMessage.Text = "Report file does not exists. Please load correct report file in project folder.";
            }
            var worksheet = workbook?.Worksheets?.FirstOrDefault(c => c.Name == sheetName);

            if(worksheet == null) {
                dgvPlanning.Visible = false;
                lblMessage.Visible = true;
                lblMessage.Text = "Invalid report file. Please load correct report file in project folder.";
            } else {
                lblMessage.Visible = false;
                dtReport = GetReportData(worksheet);
                //  dgvPlanning.DataSource = dtSource;
                // dgvPlanning.Visible = true;
                //  SetOrderColumns();
            }
        }

        private void LoadRentalExcel(string sheetName, DataGridView dgv, List<string> defaultColumns) {
            dgv.Columns.Clear();
            var dt = new DataTable();
            lblMessage.Visible = true;
            lblMessage.Text = "Loading...";
            var workbook = new XLWorkbook($"{basePath}\\{project}\\{project}_Report.xlsm");
            if(workbook == null) {
                dgv.Visible = false;
                lblMessage.Visible = true;
                lblMessage.Text = "Report file does not exists. Please load correct report file in project folder.";
                return;
            }
            var worksheet = workbook?.Worksheets?.FirstOrDefault(c => c.Name == sheetName);

            if(worksheet == null) {
                dt.Columns.AddRange(defaultColumns.Select(c => new DataColumn(c)).ToArray());
            } else {
                foreach(IXLRow row in worksheet.Rows()) {
                    if(row.RowNumber() == 1) {
                        foreach(IXLCell cell in row.Cells(1, row.LastCellUsed().Address.ColumnNumber)) {
                            dt.Columns.Add(cell.Value.ToString().Trim());
                        }
                    } else {
                        if(!string.IsNullOrEmpty(row.Cell(1).Value?.ToString())) {
                            dt.Rows.Add();
                            int i = 0;
                            foreach(IXLCell cell in row.Cells(1, dt.Columns.Count)) {
                                try {
                                    dt.Rows[dt.Rows.Count - 1][i] = cell.Value;
                                } catch { }
                                i++;
                            }
                        }
                    }
                }
            }
            dgv.DataSource = dt;
            dgv.Visible = true;
            lblMessage.Visible = false;
        }

        private void LoadRentalRates() {
            var workbook = new XLWorkbook($"{basePath}\\RentalRate.xlsx");
            if(workbook == null) {
                lblMessage.Visible = true;
                lblMessage.Text = "Rental rate file does not exists. Please load correct rental rate file in root folder.";
                return;
            }
            var worksheet = workbook?.Worksheets?.FirstOrDefault();

            if(worksheet != null) {
                foreach(IXLRow row in worksheet.Rows()) {
                    if(row.RowNumber() > 1 && !string.IsNullOrEmpty(row.Cell(4).Value.ToString())) {
                        try {
                            RentalRates.Add(new RentalRate() {
                                MatNo = row.Cell(1).Value.ToString(),
                                Rate = Convert.ToDouble(row.Cell(4).Value.ToString())
                            });
                        } catch(Exception ex) { }
                    }
                }
            }
        }

        private void SetSourceColumns() {
            for(int i = 0; i < dgvSource.Columns.Count; i++) {
                if(!sourceDisplayColumns.Contains(dgvSource.Columns[i].HeaderText)) {
                    dgvSource.Columns[i].Visible = false;
                }
            }

            dgvSource.Columns["Package"].Width = 90;
            dgvSource.Columns["MatNo"].Width = 85;
            dgvSource.Columns["quantity"].Width = 50;
            dgvSource.Columns["quantity"].HeaderText = "Qty";
            dgvSource.Columns.Remove("Comment");

            if(dgvSource.Columns["DeliveryDestination"] != null) {
                dgvSource.Columns.Remove("DeliveryDestination");
            }
            DataGridViewComboBoxColumn deliveryDestCombobox = new DataGridViewComboBoxColumn();
            deliveryDestCombobox.HeaderText = "Delivery Destination";
            deliveryDestCombobox.Name = "DeliveryDestination";
            deliveryDestCombobox.DataPropertyName = "DeliveryDestination";
            deliveryDestCombobox.DataSource = deliveryDestinations.Select(c => c.Name).ToList();

            dgvSource.Columns.Add(deliveryDestCombobox);
            dgvSource.Columns["DeliveryDestination"].Width = 100;

            if(dgvSource.Columns["Source"] != null) {
                dgvSource.Columns.Remove("Source");
            }
            foreach(DataGridViewColumn column in dgvSource.Columns) {
                if(!editableSourceColumns.Contains(column.Name)) {
                    column.ReadOnly = true;
                }
            }

            if(!rbtLeq.Checked) {
                DataGridViewComboBoxColumn sourceCombobox = new DataGridViewComboBoxColumn();
                sourceCombobox.HeaderText = "Source";
                sourceCombobox.Name = "Source";
                sourceCombobox.DataPropertyName = "Source";
                sourceCombobox.DataSource = bladeSources;

                dgvSource.Columns.Add(sourceCombobox);
                dgvSource.Columns["Source"].Width = 100;
            }

            if(dgvSource.Columns["Select"] != null) {
                dgvSource.Columns.Remove("Select");
            }
            DataGridViewCheckBoxColumn checkColumn = new DataGridViewCheckBoxColumn();
            checkColumn.Name = "Select";
            checkColumn.HeaderText = "Select";
            checkColumn.Width = 60;
            checkColumn.ReadOnly = false;
            checkColumn.FillWeight = 10;
            dgvSource.Columns.Add(checkColumn);

            foreach(DataGridViewRow row in dgvSource.Rows) {
                row.DefaultCellStyle.BackColor = ColorTranslator.FromHtml(GetColor(row?.Cells["Color"].Value?.ToString()));
            }
        }

        private void SetOrderColumns() {
            for(int i = 0; i < dgvPlanning.Columns.Count; i++) {
                if(!reportColumns.Contains(dgvPlanning.Columns[i].HeaderText)) {
                    dgvPlanning.Columns[i].Visible = false;
                }
            }

            if(dgvPlanning.Columns["Customermaterial"] == null) {
                DataGridViewComboBoxColumn inputtablecombobox = new DataGridViewComboBoxColumn();
                inputtablecombobox.HeaderText = "Customer material";
                inputtablecombobox.Name = "Customermaterial";
                inputtablecombobox.DataPropertyName = "Customer material";
                inputtablecombobox.DataSource = custMaterial.ToList();
                inputtablecombobox.DisplayMember = "key";
                inputtablecombobox.ValueMember = "value";
                dgvPlanning.Columns.RemoveAt(3);
                dgvPlanning.Columns.Insert(3, inputtablecombobox);
            }

            DataGridViewButtonColumn btnViewButton = new();
            btnViewButton.Name = "Delete";
            btnViewButton.Text = "Delete";
            btnViewButton.UseColumnTextForButtonValue = true;
            dgvPlanning.Columns.Insert(0, btnViewButton);
        }

        private static string GetColor(string color) {
            if(string.IsNullOrEmpty(color)) {
                return "#FFFFFF";
            } else if(color.ToLower() == "transparent") {
                return "#FFFFFF";
            }
            return $"#{color}";
        }

        private void GetSourceData(IXLWorksheet workSheet, IXLWorkbook workbook, string name) {
            var baseIndex = rbtLeq.Checked ? 13 : 15;
            types.Clear();
            componentlist.Clear();

            var columns = workSheet.Row(3).CellsUsed();
            var columnIndexes = new Dictionary<string, int>();
            foreach(var column in columns) {
                if(sourceColumns.Contains(column.GetString())) {
                    columnIndexes.Add(column.GetString(), column.Address.ColumnNumber);

                    dtSource.Columns.Add(column.GetString());
                }
            }
            dtSource.Columns.Add("Color");

            foreach(IXLRow row in workSheet.Rows(1, workSheet?.LastRowUsed()?.RowNumber() ?? 0)) {

                if(row.RowNumber() > 3 && !string.IsNullOrEmpty(row.Cell(1).Value?.ToString()) && (CellColor(row.Cell(1), workbook) == XLColor.FromArgb(255, 87, 47).Color || CellColor(row.Cell(1), workbook) == XLColor.FromArgb(255, 192, 0).Color)) {
                    if(row.FirstCellUsed() != null) {
                        var rowData = new List<string>();
                        var displayData = new List<string>();
                        foreach(var column in columnIndexes) {
                            rowData.Add(row.Cell(column.Value).GetString());
                        }
                        if (rowData.Count > 0)
                        {


                            if (rowData[10] == "BOM CLASSIFIED AS NON-FUNCTION" && !rowData[0].EndsWith("-L"))
                            {
                                rowData[0] = rowData[0] + "-L";
                            }
                           else if (baseIndex == 15 && rowData[12] == "BOM CLASSIFIED AS NON-FUNCTION" && !rowData[0].EndsWith("-L"))
                            {
                                rowData[0] = rowData[0] + "-L";
                            }
                        }
                        dtSource.Rows.Add(rowData.ToArray());
                        dtSource.Rows[dtSource.Rows.Count - 1]["Color"] = row.Cell(1).Style.Fill.BackgroundColor.ColorType == XLColorType.Theme ? "FFFFFF" : row.Cell(1).Style.Fill.BackgroundColor?.Color.Name;
                        types.Add(row.Cell(baseIndex).Value?.ToString());
                        componentlist.Add(row.Cell(baseIndex + 1).Value?.ToString());
                    }
                }
            }
            types = types.Distinct().OrderBy(c => c).ToList();
            types.Insert(0, "--Search by Type--");
            componentlist = componentlist.Distinct().OrderBy(c => c).ToList();
            componentlist.Insert(0, "--Search by Component--");
            cmbType.DataSource = types;
            cmbComponent.DataSource = componentlist;
        }

        private DataTable GetReportData(IXLWorksheet workSheet) {
            var dt = new DataTable();
            foreach(IXLRow row in workSheet.Rows()) {
                if(row.RowNumber() == 3) {
                    foreach(IXLCell cell in row.Cells(1, row.LastCellUsed().Address.ColumnNumber)) {
                        dt.Columns.Add(cell.Value.ToString().Trim());
                    }
                }
                if(row.RowNumber() > 3) {
                    if(!string.IsNullOrEmpty(row.Cell(3).Value?.ToString())) {
                        dt.Rows.Add();
                        int i = 0;
                        foreach(IXLCell cell in row.Cells(1, dt.Columns.Count)) {
                            try {
                                dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            } catch { }
                            i++;
                        }
                    }
                }
            }
            return dt;
        }

        private void chk2x_CheckedChanged(object sender, EventArgs e) {
            var chkbox = (CheckBox)sender;
            if(chkbox.Checked) {
                filters.Add(chkbox.Text, "x");
            } else {
                filters.Remove(chkbox.Text);
            }

            FilterRecords();
        }

        private void FilterRecords() {
            if(filters.Any()) {
                if(filters.Any(c => c.Key == "Component") == true)
                    foreach(DataGridViewRow row in dgvPlanning.Rows) {
                        if(row.Tag != null && row.Tag?.ToString().Contains(filters.FirstOrDefault(c => c.Key == "Component").Value) == true) {
                            row.Visible = true;
                        } else {
                            if(row.Visible) {
                                dgvPlanning.CurrentCell = null;
                                row.Visible = false;
                            }
                        }
                    }
                else {
                    foreach(DataGridViewRow row in dgvPlanning.Rows) {
                        row.Visible = true;
                    }
                }

                var tmpDt = dtSource.Clone();

                foreach(DataRow drtableOld in dtSource.Rows) {
                    tmpDt.ImportRow(drtableOld);
                }
                foreach(var filter in filters) {
                    if(tmpDt.Columns.Contains(filter.Key)) {
                        if(tmpDt.Rows.Count > 0) {
                            if(tmpDt?.AsEnumerable()?.Any(r => r.Field<string>(filter.Key)?.ToString()?.ToLower()?.Contains(filter.Value.ToLower()) ?? false) == true) {
                                tmpDt = tmpDt?.AsEnumerable()?.Where(r => r.Field<string>(filter.Key)?.ToString()?.ToLower()?.Contains(filter.Value.ToLower()) ?? false)?.CopyToDataTable();
                            } else {
                                tmpDt.Rows.Clear();
                            }
                        }
                    }
                }

                dgvSource.DataSource = tmpDt;
            } else {
                dgvSource.DataSource = dtSource;
                foreach(DataGridViewRow row in dgvPlanning.Rows) {
                    row.Visible = true;
                }
            }
            SetSourceColumns();
        }

        private void textBox1_TextChanged(object sender, EventArgs e) {
            filters.Remove("MatNo");
            if(!string.IsNullOrEmpty(textBox1.Text)) {
                filters.Add("MatNo", textBox1.Text);
            }
            FilterRecords();
        }

        private void button1_Click(object sender, EventArgs e) {
            lblBoltError.Visible = true;
            lblBoltError.Text = "Searching...";
            txtCopy.Text = "";
            txtSearchValue.Text = txtSearchValue.Text.Trim();
            var workbook = new XLWorkbook($"{basePath}\\{project}\\{project}_Report.xlsm");
            if(workbook == null) {
                dgvSource.Visible = false;
                lblBoltError.Visible = true;
                lblBoltError.Text = "Report file does not exists. Please load correct excel file.";
            }
            var worksheet = workbook?.Worksheets?.FirstOrDefault(c => c.Name == "GP - A9B bolt replacement list");

            if(worksheet == null) {
                dgvSource.Visible = false;
                lblBoltError.Visible = true;
                lblBoltError.Text = "Invalid report file. Please load correct excel file.";
            } else {
                lblBoltError.Visible = false;

                foreach(IXLRow row in worksheet.Rows()) {
                    if(row.RowNumber() > 1) {
                        if(row.Cell(1)?.Value?.ToString()?.ToLower() == txtSearchValue.Text.ToLower()) {
                            txtCopy.Text = String.IsNullOrEmpty(row.Cell(3)?.Value?.ToString()) ? "?" : row.Cell(3)?.Value?.ToString();
                            break;
                        } else if(row.Cell(3)?.Value?.ToString()?.ToLower() == txtSearchValue.Text.ToLower()) {
                            txtCopy.Text = String.IsNullOrEmpty(row.Cell(1)?.Value?.ToString()) ? "?" : row.Cell(1)?.Value?.ToString();
                            break;
                        } else {
                            txtCopy.Text = "?";
                        }
                    }
                }
            }

            lblBoltError.Visible = false;
        }

        private void btnCopy_Click(object sender, EventArgs e) {
            Clipboard.SetText(txtCopy.Text);
        }

        private void btnExcel_Click(object sender, EventArgs e) {
            try {
                SaveDatatoExcel();
                MessageBox.Show("Excel saved successfully", project);

            } catch(Exception ex) {
                MessageBox.Show("Something went wrong. Close the project report file if it is open.", "Error");
            }
        }

        private void SaveDatatoExcel() {
            var sheetName = rbtLeq.Checked ? "LEQ Overview" : "TEQ Overview";
            var workbook = new XLWorkbook($"{basePath}\\{project}\\{project}_Report.xlsm");
            double total = 0;
            if(workbook == null) {
                dgvSource.Visible = false;
                lblMessage.Visible = true;
                lblMessage.Text = "Report file does not exists. Please load correct order file.";
                return;
            }

            if(workbook.Worksheets.Contains(sheetName)) {
                workbook.Worksheets.Delete(sheetName);
            }
            workbook.AddWorksheet(sheetName);
            var worksheet = workbook?.Worksheets?.FirstOrDefault(c => c.Name == sheetName);

            var dt = (dgvPlanning?.DataSource as DataTable);

            for(int i = 1; i <= dt.Rows.Count; i++) {
                for(int j = 1; j <= dt.Columns.Count; j++) {
                    worksheet.Cell(i, j).Value = dt.Rows[i - 1][j - 1].ToString();
                }
                if(i % 7 != 0 && i % 7 != 6) {
                    worksheet.Row(i).Cells(true).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    worksheet.Row(i).Cells(true).Style.Border.TopBorderColor = XLColor.Black;

                    worksheet.Row(i).Cells(true).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    worksheet.Row(i).Cells(true).Style.Border.BottomBorderColor = XLColor.Black;

                    worksheet.Row(i).Cells(true).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    worksheet.Row(i).Cells(true).Style.Border.LeftBorderColor = XLColor.Black;

                    worksheet.Row(i).Cells(true).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    worksheet.Row(i).Cells(true).Style.Border.RightBorderColor = XLColor.Black;
                }

                if(i % 7 == 1 || i % 7 == 4 || i % 7 == 5) {
                    worksheet.Row(i).Cells(true).Style.Fill.BackgroundColor = XLColor.LightGray;
                    if(i % 7 == 4) {
                        worksheet.Row(i).Cells(true).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                        worksheet.Row(i).Cells(true).Style.Border.BottomBorderColor = XLColor.Black;
                    }
                    if(i % 7 == 1) {
                        worksheet.Row(i).Cells(true).Style.Border.TopBorder = XLBorderStyleValues.Thick;
                        worksheet.Row(i).Cells(true).Style.Border.TopBorderColor = XLColor.Black;
                    }
                } else if(i % 7 == 2 || i % 7 == 3) {
                    worksheet.Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.LightGray;
                }
                if(i % 7 == 5) {
                    total += Convert.ToDouble(worksheet.Row(i).Cell(2).Value);
                }
            }
            //Save Total on top
            worksheet.Row(1).InsertRowsAbove(2);

            worksheet.Row(1).Cell(1).Value = "Total ";
            worksheet.Row(1).Cell(2).Value = total;

            worksheet.Row(1).Cell(1).Style.Fill.BackgroundColor = XLColor.LightGray;
            worksheet.Row(1).Cell(2).Style.Fill.BackgroundColor = XLColor.LightGray;

            worksheet.Row(1).Cell(1).Style.Border.TopBorder = XLBorderStyleValues.Thick;
            worksheet.Row(1).Cell(1).Style.Border.TopBorderColor = XLColor.Black;
            worksheet.Row(1).Cell(2).Style.Border.TopBorder = XLBorderStyleValues.Thick;
            worksheet.Row(1).Cell(2).Style.Border.TopBorderColor = XLColor.Black;

            worksheet.Row(1).Cell(1).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
            worksheet.Row(1).Cell(1).Style.Border.BottomBorderColor = XLColor.Black;
            worksheet.Row(1).Cell(2).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
            worksheet.Row(1).Cell(2).Style.Border.BottomBorderColor = XLColor.Black;

            worksheet.Row(1).Cell(1).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            worksheet.Row(1).Cell(1).Style.Border.LeftBorderColor = XLColor.Black;

            worksheet.Row(1).Cell(1).Style.Border.RightBorder = XLBorderStyleValues.Thin;
            worksheet.Row(1).Cell(1).Style.Border.RightBorderColor = XLColor.Black;

            worksheet.Row(1).Cell(2).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            worksheet.Row(1).Cell(2).Style.Border.LeftBorderColor = XLColor.Black;

            worksheet.Row(1).Cell(2).Style.Border.RightBorder = XLBorderStyleValues.Thin;
            worksheet.Row(1).Cell(2).Style.Border.RightBorderColor = XLColor.Black;

            //Save Order info
            WriteTableToSheet(workbook, rbtLeq.Checked ? "LEQ Order" : "TEQ Order", dtReport, 3, false);

            //Save Rental info
            WriteTableToSheet(workbook, rbtLeq.Checked ? "LEQ Rental" : "TEQ Rental", dgvRental?.DataSource as DataTable);

            //Save Consumable info
            WriteTableToSheet(workbook, rbtLeq.Checked ? "LEQ Consumables" : "TEQ Consumables", dgvConsumables?.DataSource as DataTable);

            workbook.SaveAs($"{basePath}\\{project}\\{project}_Report.xlsm");
        }

        private void WriteTableToSheet(XLWorkbook? workbook, string sheetName, DataTable dt, int startIndex = 1, bool shouldUpdateHeader = true) {
            if(workbook?.Worksheets.Any(c => c.Name == sheetName) == false) {
                workbook.Worksheets.Add(sheetName);
            }
            var worksheet = workbook.Worksheets.FirstOrDefault(c => c.Name == sheetName);

            foreach(IXLRow row in worksheet.Rows()) {
                if(row.RowNumber() > startIndex) {
                    row.Clear(XLClearOptions.AllContents);
                }
            }

            var rowIndex = startIndex;
            if(shouldUpdateHeader) {
                for(int j = 1; j <= dt.Columns.Count; j++) {
                    worksheet.Cell(1, j).Value = dt.Columns[j - 1].ColumnName;
                }
            }

            foreach(DataRow row in dt.Rows) {
                rowIndex++;
                for(int j = 1; j <= dt.Columns.Count; j++) {
                    worksheet.Cell(rowIndex, j).Value = row[j - 1].ToString();
                }
            }
        }

        private Color CellColor(IXLCell cell, IXLWorkbook wb) {
            switch(cell.Style.Fill.BackgroundColor.ColorType) {
                case XLColorType.Color: {
                        return cell.Style.Fill.BackgroundColor.Color;
                    }
                case XLColorType.Theme: {
                        switch(cell.Style.Fill.BackgroundColor.ThemeColor) {
                            case XLThemeColor.Accent1: {
                                    return wb.Theme.Accent1.Color;
                                }
                            case XLThemeColor.Accent2: {
                                    return wb.Theme.Accent2.Color;
                                }
                            case XLThemeColor.Accent3: {
                                    return wb.Theme.Accent3.Color;
                                }
                            case XLThemeColor.Accent4: {
                                    return wb.Theme.Accent4.Color;
                                }
                            case XLThemeColor.Accent5: {
                                    return wb.Theme.Accent5.Color;
                                }
                            case XLThemeColor.Accent6: {
                                    return wb.Theme.Accent6.Color;
                                }
                            case XLThemeColor.Background1: {
                                    return wb.Theme.Background1.Color;
                                }
                            case XLThemeColor.Background2: {
                                    return wb.Theme.Background2.Color;
                                }
                        }

                        return XLColor.White.Color;
                    }
            }

            return XLColor.White.Color;
        }

        private void cmbType_SelectedIndexChanged(object sender, EventArgs e) {
            if(cmbType.IsHandleCreated && cmbType.Focused) {
                if(cmbType.SelectedIndex > 0) {
                    filters.Remove("Type");
                    filters.Add("Type", types[cmbType.SelectedIndex]);
                } else {
                    filters.Remove("Type");
                }

                FilterRecords();
            }
        }

        private void cmbComponent_SelectedIndexChanged(object sender, EventArgs e) {
            if(cmbComponent.IsHandleCreated && cmbComponent.Focused) {
                if(cmbComponent.SelectedIndex > 0) {
                    filters.Remove("Component");

                    filters.Add("Component", componentlist[cmbComponent.SelectedIndex]);
                } else {
                    filters.Remove("Component");
                }
                FilterRecords();
            }
        }

        private void btnBackToProject_Click(object sender, EventArgs e) {
            this.Hide();
            var m = new FrmProject();
            m.ShowDialog();
            this.Close();
        }

        private void dgvPlanning_CellPainting(object sender, DataGridViewCellPaintingEventArgs e) {

            if(e.ColumnIndex == dgvPlanning.ColumnCount - 1 && dgvPlanning.Rows[e.RowIndex].Cells[1]?.Value?.ToString() != "Total") {
                dgvPlanning.Rows[e.RowIndex].Cells[e.ColumnIndex].Style = new DataGridViewCellStyle { Padding = new Padding(dgvPlanning.Rows[e.RowIndex].Cells[e.ColumnIndex].OwningColumn.Width, 0, 0, 0) };
            }
        }

        private void dgvPlanning_CellClick(object sender, DataGridViewCellEventArgs e) {

            if(e.ColumnIndex == dgvPlanning.Columns.Count - 1 && e.RowIndex % 7 == 0) {
                //Delete records from log sheet
                var materialMeta = (dgvPlanning.DataSource as DataTable).Rows[e.RowIndex][0].ToString();
                var packageIdRegex = new Regex("(\\[).+(\\])");
                var packageId = packageIdRegex.Match(materialMeta).Value.Replace("[", "").Replace("]", "");

                dtReport.Rows.Cast<DataRow>().Where(r => r.ItemArray[dtReport.Columns.Count - 1].ToString() == packageId).ToList().ForEach(r => r.Delete());


                for(int i = 0; i < 7; i++) {
                    if(e.RowIndex < dgvPlanning.Rows.Count) {
                        dgvPlanning.Rows.RemoveAt(e.RowIndex);
                    } else {
                        if (dgvPlanning.Rows.Count > 0)
                        {
                            dgvPlanning.Rows.RemoveAt(dgvPlanning.Rows.Count - 1);
                        }
                    }
                    dgvPlanning.Refresh();
                }
            }
        }

        private void btnAddToPlanning_Click(object sender, EventArgs e) {
            var packageName = "";
            var materials = new List<string>();
            var deliveryDestination = "";

            int i = 0;
            double rate = 0;
            dgvSource.EndEdit();
            var packageId = Guid.NewGuid();

            foreach(DataGridViewRow row in dgvSource.Rows) {
                DataGridViewCheckBoxCell chkchecking = row.Cells["Select"] as DataGridViewCheckBoxCell;

                if(Convert.ToBoolean(chkchecking.Value) == true) {
                    int orderQty = 1;
                    //var orderQty = Convert.ToInt32(row.Cells["OrderQty"]?.Value?.ToString() ?? "0");

                    //if(orderQty == 0) {
                    //    MessageBox.Show($"Please provide order qty for material - {row.Cells["MatNo"].Value} ", "Error");
                    //    return;
                    //}

                    if(string.IsNullOrEmpty(row.Cells["DeliveryDestination"]?.Value?.ToString())) {
                        MessageBox.Show($"Please select delivery destination for material - {row.Cells["MatNo"].Value} ", "Error");
                        return;
                    }
                    deliveryDestination = row.Cells["DeliveryDestination"]?.Value?.ToString();

                    if(i > 0 && packageName.ToLower() != row.Cells["Package"].Value.ToString().ToLower()) {
                        MessageBox.Show("Please choose materials from same package.", "Error");
                        return;
                    }
                    packageName = row.Cells["Package"].Value.ToString();
                    i++;

                    materials.Add(row.Cells["MatNo"].Value.ToString());

                    rate += (RentalRates.FirstOrDefault(c => c.MatNo.Replace("-L","").Equals(row.Cells["MatNo"].Value.ToString().Replace("-L", ""), StringComparison.OrdinalIgnoreCase))?.Rate ?? 0) * orderQty;

                    planningMaterial.Add(new PlanningMaterial() {
                        MatNo = row.Cells["MatNo"].Value.ToString(),
                        DeliveryDestination = row.Cells["DeliveryDestination"].Value.ToString(),
                        Description = row.Cells["Object description"].Value.ToString(),
                        Qty = orderQty,
                        Package = packageName,
                        Rate = (RentalRates.FirstOrDefault(c => c.MatNo.Replace("-L", "").Equals(row.Cells["MatNo"].Value.ToString().Replace("-L", ""), StringComparison.OrdinalIgnoreCase))?.Rate ?? 0) * orderQty,
                        PackageId = packageId.ToString(),
                        EqmOwnership = GetEqmOwnership(rbtLeq.Checked ? string.Empty : row.Cells["Source"]?.Value?.ToString(), packageName),
                        Region = GetComponentSourceRegion(rbtLeq.Checked ? string.Empty : row.Cells["Source"]?.Value?.ToString(), packageName)
                    });
                }
            }

            if(packageName.ToLower().StartsWith("c-")) {

                var dtConsumables = dgvConsumables.DataSource as DataTable;

                foreach(var rental in planningMaterial) {
                    dtConsumables.Rows.Add();
                    dtConsumables.Rows[dtConsumables.Rows.Count - 1]["MatNo"] = rental.MatNo;
                    dtConsumables.Rows[dtConsumables.Rows.Count - 1]["QTY"] = rental.Qty;
                    dtConsumables.Rows[dtConsumables.Rows.Count - 1]["MatDescription"] = rental.Description;
                    dtConsumables.Rows[dtConsumables.Rows.Count - 1]["RentalCharges"] = rental.Rate;
                    dtConsumables.Rows[dtConsumables.Rows.Count - 1]["Type"] = rental.Package.ToLower().Contains("tower") ? "Tower" : "Blade";
                }
                dgvConsumables.DataSource = dtConsumables;

            } else {
                if(materials.Any() == true) {
                    var dt = dgvPlanning.DataSource as DataTable;

                    if(dt != null) {
                        if(dgvPlanning.Rows.Count > 0) {
                            dt.Rows.Add(new object[] { });
                            dt.Rows.Add(new object[] { });
                        }
                        dt.Rows.Add(GetItem($"{packageName}({string.Join(", ", materials)})[{packageId}]", CommonHelper.WeekList(StartDate, EndDate), "Total"));
                        dt.Rows.Add(GetItem("Order In", CommonHelper.WeekList(StartDate, EndDate).Select(c => $"0").ToList(), "0"));
                        dt.Rows.Add(GetItem("Order Out", CommonHelper.WeekList(StartDate, EndDate).Select(c => $"0").ToList(), "0"));
                        dt.Rows.Add(GetItem("Quantity on the project [" + deliveryDestination + "]", CommonHelper.WeekList(StartDate, EndDate).Select(c => $"0").ToList(), "0"));
                        dt.Rows.Add(GetItem($"Rate - {rate}", CommonHelper.WeekList(StartDate, EndDate).Select(c => $"0").ToList(), "0"));

                    }


                    dgvPlanning.DataSource = dt;
                    dgvPlanning.Update();
                    dgvPlanning.Refresh();

                    PaintPlanningGrid(dt);
                    foreach(DataGridViewRow row in dgvSource.Rows) {
                        DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[dgvSource.ColumnCount - 1];
                        chk.Value = chk.FalseValue;
                        //row.Cells["OrderQty"].Value = "";
                        row.Cells["DeliveryDestination"].Value = "";
                        if(rbtTeq.Checked) {
                            row.Cells["Source"].Value = "";
                        }
                    }
                    var dtRental = dgvRental.DataSource as DataTable;

                    foreach(var rental in planningMaterial) {
                        dtRental.Rows.Add();
                        dtRental.Rows[dtRental.Rows.Count - 1]["MatNo"] = rental.MatNo;
                        dtRental.Rows[dtRental.Rows.Count - 1]["QTY"] = rental.Qty;
                        dtRental.Rows[dtRental.Rows.Count - 1]["MatDescription"] = rental.Description;
                        dtRental.Rows[dtRental.Rows.Count - 1]["RentalCharges"] = rental.Rate;
                        dtRental.Rows[dtRental.Rows.Count - 1]["Package"] = rental.Package;

                    }
                    dgvRental.DataSource = dtRental;

                } else {
                    MessageBox.Show("Please choose at least one material to add in planning.", "Error");
                    return;
                }
            }
        }

        private string GetEqmOwnership(string source, string package) {
            if(rbtLeq.Checked) {
                return "NE&ME";
            }

            if(string.IsNullOrEmpty(source) || string.IsNullOrEmpty(package)) {
                return string.Empty;
            }

            if(package.ToLower().StartsWith("blade")) {
                return config.BladeSources.FirstOrDefault(c => c.Name.Equals(source)).EQMProvider;
            } else if(package.ToLower().StartsWith("tower")) {
                return config.TowerSources.FirstOrDefault(c => c.Name.Equals(source)).EQMProvider;
            } else {
                return config.MainSources.FirstOrDefault(c => c.Name.Equals(source)).EQMProvider;
            }
        }
        private string GetComponentSourceRegion(string source, string package) {
            if(rbtLeq.Checked) {
                return "NE&ME";
            }
            if(string.IsNullOrEmpty(source) || string.IsNullOrEmpty(package)) {
                return string.Empty;
            }

            if(package.ToLower().StartsWith("blade")) {
                return config.BladeSources.FirstOrDefault(c => c.Name.Equals(source)).Region;
            } else if(package.ToLower().StartsWith("tower")) {
                return config.TowerSources.FirstOrDefault(c => c.Name.Equals(source)).Region;
            } else {
                return config.MainSources.FirstOrDefault(c => c.Name.Equals(source)).Region;
            }
        }

        private void dgvPlanning_CellValueChanged(object sender, DataGridViewCellEventArgs e) {
            var dt = dgvPlanning.DataSource as DataTable;
            double rate;
            var totalIn = 0;
            var totalOut = 0;
            var totalProj = 0;
            var totalRate = 0;

            if((e.RowIndex + 1) % 7 == 2) {// Qty in updated
                rate = Int32.Parse(dt.Rows[e.RowIndex + 3][0].ToString().Split(new string[] { "- " }, StringSplitOptions.None)[1] ?? "0");

                for(int i = 2; i < dgvPlanning.Columns.Count - 1; i++) {
                    totalIn += Int32.Parse(dt.Rows[e.RowIndex][i]?.ToString() ?? "0");
                    totalOut += Int32.Parse(dt.Rows[e.RowIndex + 1][i]?.ToString() ?? "0");
                    if(i >= e.ColumnIndex) {
                        dt.Rows[e.RowIndex + 2][i] = totalIn - totalOut + Int32.Parse(dt.Rows[e.RowIndex + 1][i]?.ToString() ?? "0");
                        dt.Rows[e.RowIndex + 3][i] = (totalIn - totalOut + Int32.Parse(dt.Rows[e.RowIndex + 1][i]?.ToString() ?? "0")) * rate;
                    }
                    totalProj += Int32.Parse(dt.Rows[e.RowIndex + 2][i]?.ToString() ?? "0");
                    totalRate += Int32.Parse(dt.Rows[e.RowIndex + 3][i]?.ToString() ?? "0");
                }
                dt.Rows[e.RowIndex][1] = totalIn;
                dt.Rows[e.RowIndex + 2][1] = totalProj;
                dt.Rows[e.RowIndex + 3][1] = totalRate;

                var materialMeta = dt.Rows[e.RowIndex - 1][0].ToString();
                var packageIdRegex = new Regex("(\\[).+(\\])");
                var materialRegex = new Regex("(\\().+(\\))");
                var currentStartDate = CommonHelper.GetStartDate(dt.Rows[e.RowIndex - 1][e.ColumnIndex].ToString());

                var packageId = packageIdRegex.Match(materialMeta).Value.Replace("[", "").Replace("]", "");
                var materials = materialRegex.Match(materialMeta).Value.Replace("(", "").Replace(")", "").Split(new string[] { ", " }, StringSplitOptions.None);
                if(materials.Any()) {
                    foreach(var material in materials) {

                        if(dtReport.AsEnumerable().Any(row => row.Field<string>("PackageId") == packageId && row.Field<string>("MatNo") == material && row.Field<string>("Demand date ex warehouse") == currentStartDate.Date.ToString())) {
                            var rowIndex = dtReport.Rows.IndexOf(dtReport.AsEnumerable().FirstOrDefault(row => row.Field<string>("PackageId") == packageId && row.Field<string>("MatNo") == material && row.Field<string>("Demand date ex warehouse") == currentStartDate.Date.ToString()));
                            dtReport.Rows[rowIndex]["QTY"] = rbtLeq.Checked ? 1 : dt.Rows[e.RowIndex][e.ColumnIndex];
                        } else {
                            var rental = planningMaterial.FirstOrDefault(c => c.PackageId == packageId && c.MatNo == material);
                            if(rental == null) {
                                var rowIndexOlder = dtReport.Rows.IndexOf(dtReport.AsEnumerable().FirstOrDefault(row => row.Field<string>("PackageId") == packageId && row.Field<string>("MatNo") == material));
                                rental = new PlanningMaterial();
                                rental.Description = dtReport.Rows[rowIndexOlder]["MatDescription"].ToString();
                                rental.DeliveryDestination = dtReport.Rows[rowIndexOlder]["Delivery Destination"].ToString();
                                if(rbtTeq.Checked) {
                                    rental.Region = dtReport.Rows[rowIndexOlder]["Component Sourcing Region"].ToString();
                                }
                                rental.EqmOwnership = dtReport.Rows[rowIndexOlder]["EQM Ownership"].ToString();
                            }
                            string storage = rental.DeliveryDestination;
                            string deliverydest = "";
                            if (!string.IsNullOrEmpty(rental.DeliveryDestination))
                            {
                                char delimiter = '-';
                                if (storage.Contains("-"))
                                {
                                    deliverydest = storage.Split('-')[1];
                                    storage = storage.Split(delimiter)[0];
                                  
                                }

                            }
                            for(int i = 0; i < Convert.ToInt32(dt.Rows[e.RowIndex][e.ColumnIndex]); i++) {
                                dtReport.Rows.Add();
                                var mat = material;
                                if(!string.IsNullOrEmpty(rental.Description) && rental.Description.ToUpper() == "BOM CLASSIFIED AS NON-FUNCTION" && !material.EndsWith("-L"))
                                {
                                    mat = material + "-L";
                                }
                                dtReport.Rows[dtReport.Rows.Count - 1]["Customer material"] = rbtLeq.Checked ? "LEQ" : "TEQ";
                                dtReport.Rows[dtReport.Rows.Count - 1]["MatDescription"] = rental.Description;
                                dtReport.Rows[dtReport.Rows.Count - 1]["MatNo"] = mat;
                                dtReport.Rows[dtReport.Rows.Count - 1]["QTY"] = 1;
                                dtReport.Rows[dtReport.Rows.Count - 1]["Delivery Destination"] = deliverydest;
                                dtReport.Rows[dtReport.Rows.Count - 1]["Contact person and Details for Delivery"] = deliveryDestinations?.FirstOrDefault(c => c.Name == rental.DeliveryDestination)?.PersonName ?? string.Empty;
                                dtReport.Rows[dtReport.Rows.Count - 1]["Plant"] = deliveryDestinations?.FirstOrDefault(c => c.Name == rental.DeliveryDestination)?.ExecutionPlant ?? string.Empty;
                                dtReport.Rows[dtReport.Rows.Count - 1]["Storage"] = storage;
                                dtReport.Rows[dtReport.Rows.Count - 1]["WBS"] = deliveryDestinations?.FirstOrDefault(c => c.Name == rental.DeliveryDestination)?.WBS ?? string.Empty;
                                dtReport.Rows[dtReport.Rows.Count - 1]["Demand date ex warehouse"] = currentStartDate;
                                dtReport.Rows[dtReport.Rows.Count - 1]["Incoterm"] = deliveryDestinations?.FirstOrDefault(c => c.Name == rental.DeliveryDestination)?.Incoterm ?? string.Empty;
                                dtReport.Rows[dtReport.Rows.Count - 1]["Incoterm Location"] = deliveryDestinations?.FirstOrDefault(c => c.Name == rental.DeliveryDestination)?.IncotermLocation ?? string.Empty;
                                if (rbtTeq.Checked) {
                                    dtReport.Rows[dtReport.Rows.Count - 1]["Component Sourcing Region"] = rental.Region;
                                }
                                dtReport.Rows[dtReport.Rows.Count - 1]["EQM Ownership"] = rental.EqmOwnership;
                                dtReport.Rows[dtReport.Rows.Count - 1]["PackageId"] = packageId;
                            }
                        }
                    }
                }
            }
            if((e.RowIndex + 1) % 7 == 3) {// Qty out updated
                rate = Int32.Parse(dt.Rows[e.RowIndex + 2][0].ToString().Split(new string[] { "- " }, StringSplitOptions.None)[1] ?? "0");
                var qty = Convert.ToInt32(dt.Rows[e.RowIndex][e.ColumnIndex].ToString());

                var materialMeta = dt.Rows[e.RowIndex - 2][0].ToString();
                var packageIdRegex = new Regex("(\\[).+(\\])");
                var materialRegex = new Regex("(\\().+(\\))");
                var currentEndDate = CommonHelper.GetEndDate(dt.Rows[e.RowIndex - 2][e.ColumnIndex].ToString());

                var packageId = packageIdRegex.Match(materialMeta).Value.Replace("[", "").Replace("]", "");
                var materials = materialRegex.Match(materialMeta).Value.Replace("(", "").Replace(")", "").Split(new string[] { ", " }, StringSplitOptions.None);
                if(materials.Any()) {
                    foreach(var material in materials) {
                      
                        if (dtReport.AsEnumerable().Any(row => row.Field<String>("PackageId") == packageId && row.Field<String>("MatNo") == material)) {
                            var rowIndex = dtReport.Rows.IndexOf(dtReport.AsEnumerable().FirstOrDefault(row => row.Field<String>("PackageId") == packageId && row.Field<String>("MatNo") == material));
                            do {
                                if(qty == 0) {
                                    break;
                                }
                                if(dtReport.Rows[rowIndex]["PackageId"].ToString() == packageId && dtReport.Rows[rowIndex]["MatNo"].ToString() == material) {
                                    dtReport.Rows[rowIndex]["Return date"] = currentEndDate;
                                    qty -= Convert.ToInt32(dtReport.Rows[rowIndex]["QTY"].ToString());
                                }
                                rowIndex++;
                            } while(rowIndex < dtReport.Rows.Count);
                        }
                    }
                }

                for(int i = 2; i < dgvPlanning.Columns.Count - 1; i++) {
                    totalIn += Int32.Parse(dt.Rows[e.RowIndex - 1][i]?.ToString() ?? "0");
                    totalOut += Int32.Parse(dt.Rows[e.RowIndex][i]?.ToString() ?? "0");
                    if(i > e.ColumnIndex) {
                        dt.Rows[e.RowIndex + 1][i] = totalIn - totalOut + Int32.Parse(dt.Rows[e.RowIndex][i]?.ToString() ?? "0");
                        dt.Rows[e.RowIndex + 2][i] = (totalIn - totalOut + Int32.Parse(dt.Rows[e.RowIndex][i]?.ToString() ?? "0")) * rate;
                    }
                    totalProj += Int32.Parse(dt.Rows[e.RowIndex + 1][i]?.ToString() ?? "0");
                    totalRate += Int32.Parse(dt.Rows[e.RowIndex + 2][i]?.ToString() ?? "0");
                }
                dt.Rows[e.RowIndex][1] = totalOut;
                dt.Rows[e.RowIndex + 1][1] = totalProj;
                dt.Rows[e.RowIndex + 2][1] = totalRate;
            }
            dgvPlanning.DataSource = dt;
            dgvPlanning.Refresh();
            dgvPlanning.Update();
        }

        private void rbtTeq_CheckedChanged(object sender, EventArgs e) {
            LoadData("TEQ");
        }

        private void rbtLeq_CheckedChanged(object sender, EventArgs e) {
            LoadData("LEQ");
        }

        private void btnSubmit_Click(object sender, EventArgs e) {
            for(int i = 0; i < dtReport.Rows.Count; i++) {
                if(string.IsNullOrWhiteSpace(dtReport.Rows[i]["Date of submission"].ToString())) {
                    dtReport.Rows[i]["Date of submission"] = DateTime.Now.ToString(CommonHelper.DateFormat);
                }
            }

            try {
                SaveDatatoExcel();
                var workbook = new XLWorkbook($"{basePath}\\{project}\\{project}_Report.xlsm");

                var worksheet = workbook?.Worksheets?.FirstOrDefault(c => c.Name == (rbtTeq.Checked ? "LEQ Order" : "TEQ Order"));

                var otherOrderDt = GetReportData(worksheet);
                for(int i = 0; i < otherOrderDt.Rows.Count; i++) {
                    if(string.IsNullOrWhiteSpace(otherOrderDt.Rows[i]["Date of submission"].ToString())) {
                        otherOrderDt.Rows[i]["Date of submission"] = DateTime.Now.ToString(CommonHelper.DateFormat);
                    }
                }
                WriteTableToSheet(workbook, rbtLeq.Checked ? "TEQ Order" : "LEQ Order", otherOrderDt, 3, false);
                workbook.SaveAs($"{basePath}\\{project}\\{project}_Report.xlsm");
                MessageBox.Show("MM submission done sucessfully", project);

            } catch(Exception ex) {
                MessageBox.Show("Something went wrong. Close the project report file if it is open.", "Error");
            }
        }

        void setSource(int row, string type) {
            DataGridViewComboBoxCell sourceDropdown =
                                   (DataGridViewComboBoxCell)(dgvSource.Rows[row].Cells["Source"]);
            if(string.IsNullOrEmpty(type)) {
                return;
            }
            if(type.ToLower().StartsWith("blade")) {
                sourceDropdown.DataSource = bladeSources;
            } else if(type.ToLower().StartsWith("tower")) {
                sourceDropdown.DataSource = towerSources;
            } else {
                sourceDropdown.DataSource = mainSources;
            }

            sourceDropdown.Value = sourceDropdown.Items[0];
        }

        private void dgvSource_CellValueChanged(object sender, DataGridViewCellEventArgs e) {

        }

        private void dgvSource_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e) {
            if(rbtTeq.Checked && e.ColumnIndex == dgvSource.Columns["Source"].Index) {
                setSource(e.RowIndex, dgvSource.Rows[e.RowIndex].Cells["Package"].Value.ToString());
            }
        }

        private void dgvSource_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e) {
            if(e.RowIndex >= 0 && e.ColumnIndex == 0) {
                e.ToolTipText = (dgvSource.DataSource as DataTable).Rows[e.RowIndex]["Comment"].ToString();
            }
        }

        private void btnCopyCockpit_Click(object sender, EventArgs e) {
            SaveToCockpit(rbtLeq.Checked ? "LEQ" : "TEQ");
        }

        private void SaveToCockpit(string type) {

            var helper = new SharePointHelper("test", "test", "test", basePath);

            var reportFile = $"{basePath}\\{project}\\{project}_Cockpit_{type.ToUpper()}.xlsx";

            if(File.Exists(reportFile)) {
                File.Delete(reportFile);
            }
            helper.SaveFile($"{basePath}\\CockpitReport.xlsx", reportFile);

            var workbook = new XLWorkbook(reportFile);

            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("BUSINESS SCENARIO", typeof(string));
            dataTable.Columns.Add("VERSION", typeof(string));
            dataTable.Columns.Add("PROJECT UPID", typeof(string));
            dataTable.Columns.Add("RENTAL FO PO", typeof(string));
            dataTable.Columns.Add("MAIN COMPONENT", typeof(string));
            dataTable.Columns.Add("TEQ/TOOL MATERIAL", typeof(string));
            dataTable.Columns.Add("TEQ QUANTITY", typeof(int));
            dataTable.Columns.Add("FREE OF CHARGE(FOC)", typeof(string));
            dataTable.Columns.Add("DELIVERY GROUP", typeof(string));
            dataTable.Columns.Add("PO SPLIT", typeof(string));
            dataTable.Columns.Add("SHIPPING INSTRUCTION HEADER TEXT", typeof(string));
            dataTable.Columns.Add("VENDOR", typeof(string));
            dataTable.Columns.Add("SUPPLY PLANT", typeof(string));
            dataTable.Columns.Add("SUPPLY SLOC", typeof(string));
            dataTable.Columns.Add("DELIVERY DATE", typeof(string));
            dataTable.Columns.Add("RECV PLANT", typeof(string));
            dataTable.Columns.Add("RECV SLOC", typeof(string));
            dataTable.Columns.Add("SHIP TO PARTY", typeof(string));
            dataTable.Columns.Add("VALUATION TYPE", typeof(string));
            dataTable.Columns.Add("ACC ASSGMT CAT", typeof(string));
            dataTable.Columns.Add("WBS", typeof(string));
            dataTable.Columns.Add("COST CENTER", typeof(string));
            dataTable.Columns.Add("PURCHASING GROUP", typeof(string));
            dataTable.Columns.Add("INCOTERMS", typeof(string));
            dataTable.Columns.Add("INCOTERMS LOCATION", typeof(string));

            foreach(DataRow reportRow in dtReport.Rows) {
                DataRow row = dataTable.NewRow();
                row["BUSINESS SCENARIO"] = "20";
                row["VERSION"] = "1";
                row["PROJECT UPID"] = upid;
                row["RENTAL FO PO"] = string.Empty;
                row["MAIN COMPONENT"] = string.Empty;
                row["TEQ/TOOL MATERIAL"] = reportRow["MatNo"] ?? string.Empty;
                row["TEQ QUANTITY"] = reportRow["QTY"] ?? string.Empty;
                row["FREE OF CHARGE(FOC)"] = string.Empty;
                row["DELIVERY GROUP"] = string.Empty;
                row["PO SPLIT"] = string.Empty;
                row["SHIPPING INSTRUCTION HEADER TEXT"] = string.Empty;
                row["VENDOR"] = string.Empty;
                row["SUPPLY PLANT"] = type.ToUpper() == "LEQ" ? "DK66" : "DK64";
                row["SUPPLY SLOC"] = string.Empty;
                row["DELIVERY DATE"] = reportRow["Demand date ex warehouse"] ?? string.Empty;
                row["RECV PLANT"] = executionPlant;
                row["RECV SLOC"] = storageLocation;
                row["SHIP TO PARTY"] = string.Empty;
                row["VALUATION TYPE"] = string.Empty;
                row["ACC ASSGMT CAT"] = string.Empty;
                row["WBS"] = type.ToUpper() == "LEQ" ? wbsLEQ : wbsTEQ;
                row["COST CENTER"] = string.Empty;
                row["PURCHASING GROUP"] = string.Empty;
                row["INCOTERMS"] = string.Empty;
                row["INCOTERMS LOCATION"] = string.Empty;
                dataTable.Rows.Add(row);
            }

            WriteTableToSheet(workbook, "Sheet1", dataTable, 3, false);

            workbook.SaveAs(reportFile);

            MessageBox.Show($"{type.ToUpper()} Cockpit report saved successfully.", project);
        }
    }
}
