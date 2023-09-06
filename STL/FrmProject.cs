using ClosedXML.Excel;
using Newtonsoft.Json;
using STL.Helpers;
using STL.Models;
using System.Data;

namespace STL {
    public partial class FrmProject : Form {

        string rootPath = string.Empty;
        Config config = new Config();

        SharePointHelper helper;

        public FrmProject() {
            InitializeComponent();
        }

        private void ToggleNewProjectSection(bool visible) {
            pnlForm.Visible = visible;
            pnlList.Visible = !visible;
            lblMessage.Visible = false;
            if(visible == false) {
                GetProjects();
                txtSearch.Clear();
            }
        }

        private void btnNew_Click(object sender, EventArgs e) {
            ToggleNewProjectSection(true);
        }

        private void btnCancel_Click(object sender, EventArgs e) {
            ToggleNewProjectSection(false);
            lblProjName.Text = string.Empty;
            ResetForm();
        }

        private void ResetForm() {
            foreach(Control control in pnlForm.Controls) {
                if(control is TextBox) {
                    control.ResetText();
                }
            }
        }

        private void GetProjects(string filter = "") {
            var dtExcelLog = new DataTable();
            dtExcelLog.Columns.Add("Name");
            var listFolders = helper.ReadFolders("");
            if(!string.IsNullOrEmpty(filter)) {
                listFolders = listFolders.Where(c => c.ToLower().Contains(filter));
            }

            for(int i = 0; i < listFolders.Count(); i++) {
                dtExcelLog.Rows.Add();
                dtExcelLog.Rows[i][0] = listFolders.ElementAt(i);
            }

            dgvExcelLog.DataSource = dtExcelLog;

            if(dgvExcelLog.Columns.Count == 1) {
                dgvExcelLog.Columns[0].Width = 400;
                DataGridViewButtonColumn btnOpenButton = new();
                btnOpenButton.Name = "Open";
                btnOpenButton.Text = "Open";
                btnOpenButton.UseColumnTextForButtonValue = true;
                dgvExcelLog.Columns.Insert(1, btnOpenButton);

                DataGridViewButtonColumn btnEditButton = new();
                btnEditButton.Name = "Edit";
                btnEditButton.Text = "Edit";
                btnEditButton.UseColumnTextForButtonValue = true;
                dgvExcelLog.Columns.Insert(2, btnEditButton);
            }
            pnlList.Left = 100;
        }

        private void btnSearch_Click(object sender, EventArgs e) {
            GetProjects(txtSearch.Text.ToLower());
        }

        private void FrmProject_Load(object sender, EventArgs e) {

            if(File.Exists("config.json")) {
                using StreamReader r = new StreamReader("config.json");
                string json = r.ReadToEnd();
                config = JsonConvert.DeserializeObject<Config>(json);
                if(config != null) {
                    cmb1Blade.DataSource = config.BladeSources.Select(c => c.Name).ToList();
                    cmb2Blade.DataSource = config.BladeSources.Select(c => c.Name).ToList();

                    cmb1Tower.DataSource = config.TowerSources.Select(c => c.Name).ToList();
                    cmb2Tower.DataSource = config.TowerSources.Select(c => c.Name).ToList();

                    cmb1Turbine.DataSource = config.MainSources.Select(c => c.Name).ToList();
                    cmb2Turbine.DataSource = config.MainSources.Select(c => c.Name).ToList();
                }
            }

            lblMessage.Text = "";
            pnlList.Left = 100;

            Logger.Write($" OneDrive - {Environment.GetEnvironmentVariable("OneDrive")}, OneDriveCommercial - {Environment.GetEnvironmentVariable("OneDriveCommercial")}, OneDriveConsumer - {Environment.GetEnvironmentVariable("OneDriveConsumer")} ");

            rootPath = CommonHelper.GetRootPath();

            if(string.IsNullOrEmpty(rootPath)) {
                ShowError("One drive installtion not found. Please install One drive to continue.");
                return;
            }
            rootPath = $"{rootPath}\\LOG orders\\Files";
            if(!Directory.Exists(rootPath)) {
                ShowError("Project root folder not found. Please contact admin to share the project folder with you.");
                return;
            }

            helper = new SharePointHelper("test", "test", "test", rootPath);
            GetProjects();
        }

        private void btnSave_Click(object sender, EventArgs e) {
            try {

                lblMessage.Visible = false;
                txtName.Text = txtName.Text.Trim();
                if(string.IsNullOrEmpty(txtName.Text)) {
                    ShowError("Project name should not be empty.");
                    return;
                }

                if(string.IsNullOrEmpty(txtStartDate.Text)) {
                    ShowError("Start date should not be empty.");
                    return;
                }

                if(string.IsNullOrEmpty(txtEndDate.Text)) {
                    ShowError("Start date should not be empty.");
                    return;
                }

                if(txtStartDate.Value > txtEndDate.Value) {
                    ShowError("Start date should be lesser than end date.");
                    return;
                }

                if(txtName.Text != lblProjName.Text && helper.CheckIfFolderExists(txtName.Text)) {
                    ShowError("Project already exists.");
                    return;
                }

                btnSave.Enabled = false;
                btnCancel.Enabled = false;

                if(string.IsNullOrEmpty(lblProjName.Text)) {
                    helper.CreateFolder(txtName.Text);
                    helper.SaveFile($"{rootPath}\\Report.xlsm", $"{rootPath}\\{txtName.Text}\\{txtName.Text}_Report.xlsm");
                } else {
                    if(txtName.Text != lblProjName.Text) {
                        helper.MoveFolder(lblProjName.Text, txtName.Text);
                        helper.moveFile($"{rootPath}\\{txtName.Text}\\{lblProjName.Text}_Report.xlsm", $"{rootPath}\\{txtName.Text}\\{txtName.Text}_Report.xlsm");
                    }
                }
                //Update the project details

                string fileName = $"{rootPath}\\{txtName.Text}\\{txtName.Text}_Report.xlsm";
                var workbook = new XLWorkbook(fileName);
                if(workbook != null) {
                    var worksheet = workbook.Worksheets.FirstOrDefault(c => c.Name == "Project Info");

                    if(worksheet != null) {
                        worksheet.Cell(17, 4).Value = cmb1Blade.SelectedValue;
                        worksheet.Cell(18, 4).Value = cmb1Tower.SelectedValue;
                        worksheet.Cell(16, 4).Value = cmb1Turbine.SelectedValue;
                        worksheet.Cell(21, 4).Value = cmb2Blade.SelectedValue;
                        worksheet.Cell(22, 4).Value = cmb2Tower.SelectedValue;
                        worksheet.Cell(20, 4).Value = cmb2Turbine.SelectedValue;
                        worksheet.Cell(4, 8).Value = txt1Contact.Text;
                        worksheet.Cell(4, 9).Value = txt1Gid.Text;
                        worksheet.Cell(4, 7).Value = txt1Name.Text;
                        worksheet.Cell(5, 8).Value = txt2Contact.Text;
                        worksheet.Cell(5, 9).Value = txt2Gid.Text;
                        worksheet.Cell(5, 7).Value = txt2Name.Text;
                        worksheet.Cell(6, 8).Value = txt3Contact.Text;
                        worksheet.Cell(6, 9).Value = txt3Gid.Text;
                        worksheet.Cell(6, 7).Value = txt3Name.Text;
                        worksheet.Cell(7, 8).Value = txt4Contact.Text;
                        worksheet.Cell(7, 9).Value = txt4Gid.Text;
                        worksheet.Cell(7, 7).Value = txt4Name.Text;
                        worksheet.Cell(8, 8).Value = txt5Contact.Text;
                        worksheet.Cell(8, 9).Value = txt5Gid.Text;
                        worksheet.Cell(8, 7).Value = txt5Name.Text;
                        worksheet.Cell(9, 8).Value = txt6Contact.Text;
                        worksheet.Cell(9, 9).Value = txt6Gid.Text;
                        worksheet.Cell(9, 7).Value = txt6Name.Text;
                        worksheet.Cell(10, 8).Value = txt7Contact.Text;
                        worksheet.Cell(10, 9).Value = txt7Gid.Text;
                        worksheet.Cell(10, 7).Value = txt7Name.Text;
                        worksheet.Cell(11, 8).Value = txt8Contact.Text;
                        worksheet.Cell(11, 9).Value = txt8Gid.Text;
                        worksheet.Cell(11, 7).Value = txt8Name.Text;
                        worksheet.Cell(12, 7).Value = txt9Name.Text;
                        worksheet.Cell(12, 8).Value = txt9Contact.Text;
                        worksheet.Cell(12, 9).Value = txt9Gid.Text;

                        worksheet.Cell(17, 3).Value = txtBladeType1.Text;
                        worksheet.Cell(21, 3).Value = txtBladeType2.Text;
                        worksheet.Cell(8, 3).Value = txtCity.Text;
                        worksheet.Cell(9, 3).Value = txtCountry.Text;
                        worksheet.Cell(10, 3).Value = txtCustomer.Text;
                        worksheet.Cell(4, 3).Value = txtEndDate.Value;
                        worksheet.Cell(7, 3).Value = txtPostalCode.Text;

                        worksheet.Cell(16, 6).Value = txtstrlc1.Text;
                        worksheet.Cell(16, 7).Value = txtEp1.Text;
                        worksheet.Cell(16, 8).Value = txtWBS1.Text;
                        worksheet.Cell(16, 9).Value = txtFOPO1.Text;
                        worksheet.Cell(16, 10).Value = txtincoterm1.Text;
                        worksheet.Cell(16, 11).Value = txtincotermlc1.Text;
                        worksheet.Cell(16, 12).Value = txtcontactperson1.Text;

                        worksheet.Cell(17, 6).Value = txtstrlc2.Text;
                        worksheet.Cell(17, 7).Value = txtEp2.Text;
                        worksheet.Cell(17, 8).Value = txtWBS2.Text;
                        worksheet.Cell(17, 9).Value = txtFOPO2.Text;
                        worksheet.Cell(17, 10).Value = txtincoterm2.Text;
                        worksheet.Cell(17, 11).Value = txtincotermlc2.Text;
                        worksheet.Cell(17, 12).Value = txtcontactperson2.Text;

                        worksheet.Cell(18, 6).Value = txtstrlc3.Text;
                        worksheet.Cell(18, 7).Value = txtEp3.Text;
                        worksheet.Cell(18, 8).Value = txtWBS3.Text;
                        worksheet.Cell(18, 9).Value = txtFOPO3.Text;
                        worksheet.Cell(18, 10).Value = txtincoterm3.Text;
                        worksheet.Cell(18, 11).Value = txtincotermlc3.Text;
                        worksheet.Cell(18, 12).Value = txtcontactperson3.Text;

                        worksheet.Cell(19, 6).Value = txtstrlc4.Text;
                        worksheet.Cell(19, 7).Value = txtEp4.Text;
                        worksheet.Cell(19, 8).Value = txtWBS4.Text;
                        worksheet.Cell(19, 9).Value = txtFOPO4.Text;
                        worksheet.Cell(19, 10).Value = txtincoterm4.Text;
                        worksheet.Cell(19, 11).Value = txtincotermlc4.Text;
                        worksheet.Cell(19, 12).Value = txtcontactperson4.Text;

                        worksheet.Cell(20, 6).Value = txtstrlc5.Text;
                        worksheet.Cell(20, 7).Value = txtEp5.Text;
                        worksheet.Cell(20, 8).Value = txtWBS5.Text;
                        worksheet.Cell(20, 9).Value = txtFOPO5.Text;
                        worksheet.Cell(20, 10).Value = txtincoterm5.Text;
                        worksheet.Cell(20, 11).Value = txtincotermlc5.Text;
                        worksheet.Cell(20, 12).Value = txtcontactperson5.Text;

                        worksheet.Cell(21, 6).Value = txtstrlc6.Text;
                        worksheet.Cell(21, 7).Value = txtEp6.Text;
                        worksheet.Cell(21, 8).Value = txtWBS6.Text;
                        worksheet.Cell(21, 9).Value = txtFOPO6.Text;
                        worksheet.Cell(21, 10).Value = txtincoterm6.Text;
                        worksheet.Cell(21, 11).Value = txtincotermlc6.Text;
                        worksheet.Cell(21, 12).Value = txtcontactperson6.Text;

                        worksheet.Cell(22, 6).Value = txtstrlc7.Text;
                        worksheet.Cell(22, 7).Value = txtEp7.Text;
                        worksheet.Cell(22, 8).Value = txtWBS7.Text;
                        worksheet.Cell(22, 9).Value = txtFOPO7.Text;
                        worksheet.Cell(22, 10).Value = txtincoterm7.Text;
                        worksheet.Cell(22, 11).Value = txtincotermlc7.Text;
                        worksheet.Cell(22, 12).Value = txtcontactperson7.Text;

                        worksheet.Cell(23, 6).Value = txtstrlc8.Text;
                        worksheet.Cell(23, 7).Value = txtEp8.Text;
                        worksheet.Cell(23, 8).Value = txtWBS8.Text;
                        worksheet.Cell(23, 9).Value = txtFOPO8.Text;
                        worksheet.Cell(23, 10).Value = txtincoterm8.Text;
                        worksheet.Cell(23, 11).Value = txtincotermlc8.Text;
                        worksheet.Cell(23, 12).Value = txtcontactperson8.Text;

                       // worksheet.Cell(27, 3).Value = txtRerportFopo.Text;
                        worksheet.Cell(12, 3).Value = txtProjectUPId.Text;

                        //worksheet.Cell(29, 3).Value = txtwbsleq.Text;
                        //worksheet.Cell(30, 3).Value = txtwbsteq.Text;
                        //worksheet.Cell(31, 3).Value = txtexecutionplant.Text;

                        worksheet.Cell(3, 3).Value = txtStartDate.Value;
                        worksheet.Cell(6, 3).Value = txtStreet.Text;
                        worksheet.Cell(18, 3).Value = txtTowerType1.Text;
                        worksheet.Cell(22, 3).Value = txtTowerType2.Text;
                        worksheet.Cell(11, 3).Value = txtTurbineCount.Text;
                        worksheet.Cell(19, 3).Value = txtTurbineCount1.Text;
                        worksheet.Cell(23, 3).Value = txtTurbineCount2.Text;
                        worksheet.Cell(16, 3).Value = txtTurbineType1.Text;
                        worksheet.Cell(20, 3).Value = txtTurbineType2.Text;
                        worksheet.Cell(5, 3).Value = txtName.Text;

                        worksheet.Cell(16, 13).Value = txtMainCWBS.Text;
                        worksheet.Cell(16, 14).Value = txtMainCPO.Text;

                        worksheet.Cell(17, 13).Value = txtMainCWBS_2.Text;
                        worksheet.Cell(17, 14).Value = txtMainCPO_2.Text;

                        worksheet.Cell(18, 13).Value = txtMainCWBS_3.Text;
                        worksheet.Cell(18, 14).Value = txtMainCPO_3.Text;


                    }

                    workbook.SaveAs(fileName);
                }

                ResetForm();
                btnSave.Enabled = true;
                btnCancel.Enabled = true;

                ToggleNewProjectSection(false);
                GetProjects();
            } catch(Exception ex) {
                Logger.Write($"[{ex.Message}] - {ex.StackTrace}");
                MessageBox.Show("Something went wrong. Please try again", "Error");
                btnSave.Enabled = true;
                btnCancel.Enabled = true;
            }
        }

        private void ShowError(string message) {
            MessageBox.Show(message, "Error");
            // lblMessage.Text = message;
            // lblMessage.Visible = true;
        }

        private void dgvExcelLog_CellClick(object sender, DataGridViewCellEventArgs e) {
            var project = dgvExcelLog.CurrentRow.Cells["Name"]?.Value?.ToString();
            if(e.ColumnIndex == dgvExcelLog.Columns["Edit"].Index) {
                try {
                    lblProjName.Text = project;
                    var workbook = new XLWorkbook($"{rootPath}\\\\{project}\\{project}_Report.xlsm");
                    if(workbook == null) {
                        ShowError("Report file does not exists. Please load correct report file in project folder.");
                    }
                    var worksheet = workbook?.Worksheets?.FirstOrDefault(c => c.Name == "Project Info");

                    if(worksheet == null) {
                        ShowError("Report file does not contains the project info sheet. Please load correct report file in project folder.");
                    } else {
                        txtName.Text = worksheet.Cell(5, 3).Value?.ToString();
                        txtStreet.Text = worksheet.Cell(6, 3).Value.ToString();
                        txtPostalCode.Text = worksheet.Cell(7, 3).Value.ToString();
                        txtCity.Text = worksheet.Cell(8, 3).Value.ToString();
                        txtCountry.Text = worksheet.Cell(9, 3).Value.ToString();
                        txtCustomer.Text = worksheet.Cell(10, 3).Value.ToString();
                        txtTurbineCount.Text = worksheet.Cell(11, 3).Value.ToString();

                        txtTurbineType1.Text = worksheet.Cell(16, 3).Value.ToString();
                        txtBladeType1.Text = worksheet.Cell(17, 3).Value.ToString();
                        txtTowerType1.Text = worksheet.Cell(18, 3).Value.ToString();
                        txtTurbineCount1.Text = worksheet.Cell(19, 3).Value.ToString();
                        txtTurbineType2.Text = worksheet.Cell(20, 3).Value.ToString();
                        txtBladeType2.Text = worksheet.Cell(21, 3).Value.ToString();
                        txtTowerType2.Text = worksheet.Cell(22, 3).Value.ToString();
                        txtTurbineCount2.Text = worksheet.Cell(23, 3).Value.ToString();

                        txt1Name.Text = worksheet.Cell(4, 7).Value.ToString();
                        txt1Contact.Text = worksheet.Cell(4, 8).Value.ToString();
                        txt1Gid.Text = worksheet.Cell(4, 9).Value.ToString();

                        txt2Name.Text = worksheet.Cell(5, 7).Value.ToString();
                        txt2Contact.Text = worksheet.Cell(5, 8).Value.ToString();
                        txt2Gid.Text = worksheet.Cell(5, 9).Value.ToString();

                        txt3Name.Text = worksheet.Cell(6, 7).Value.ToString();
                        txt3Contact.Text = worksheet.Cell(6, 8).Value.ToString();
                        txt3Gid.Text = worksheet.Cell(6, 9).Value.ToString();

                        txt4Name.Text = worksheet.Cell(7, 7).Value.ToString();
                        txt4Contact.Text = worksheet.Cell(7, 8).Value.ToString();
                        txt4Gid.Text = worksheet.Cell(7, 9).Value.ToString();

                        txt5Name.Text = worksheet.Cell(8, 7).Value.ToString();
                        txt5Contact.Text = worksheet.Cell(8, 8).Value.ToString();
                        txt5Gid.Text = worksheet.Cell(8, 9).Value.ToString();

                        txt6Name.Text = worksheet.Cell(9, 7).Value.ToString();
                        txt6Contact.Text = worksheet.Cell(9, 8).Value.ToString();
                        txt6Gid.Text = worksheet.Cell(9, 9).Value.ToString();

                        txt7Name.Text = worksheet.Cell(10, 7).Value.ToString();
                        txt7Contact.Text = worksheet.Cell(10, 8).Value.ToString();
                        txt7Gid.Text = worksheet.Cell(10, 9).Value.ToString();

                        txt8Name.Text = worksheet.Cell(11, 7).Value.ToString();
                        txt8Contact.Text = worksheet.Cell(11, 8).Value.ToString();
                        txt8Gid.Text = worksheet.Cell(11, 9).Value.ToString();

                        txtStartDate.Value = DateTime.Parse(worksheet.Cell(3, 3)?.Value?.ToString());
                        txtEndDate.Value = DateTime.Parse(worksheet.Cell(4, 3)?.Value?.ToString());                                     

                        if(!string.IsNullOrEmpty(worksheet.Cell(16, 4).Value.ToString())) {
                            cmb1Turbine.SelectedIndex = cmb1Turbine.FindString(worksheet.Cell(16, 4).Value.ToString());
                        }
                      
                        if(!string.IsNullOrEmpty(worksheet.Cell(17, 4).Value.ToString())) {
                            cmb1Blade.SelectedIndex = cmb1Blade.FindString(worksheet.Cell(17, 4).Value.ToString());
                        }

                        if(!string.IsNullOrEmpty(worksheet.Cell(18, 4).Value.ToString())) {
                            cmb1Tower.SelectedIndex = cmb1Tower.FindString(worksheet.Cell(18, 4).Value.ToString());
                        }

                        if(!string.IsNullOrEmpty(worksheet.Cell(20, 4).Value.ToString())) {
                            cmb2Turbine.SelectedIndex = cmb2Turbine.FindString(worksheet.Cell(20, 4).Value.ToString());
                        }
                        if(!string.IsNullOrEmpty(worksheet.Cell(21, 4).Value.ToString())) {
                            cmb2Blade.SelectedIndex = cmb2Blade.FindString(worksheet.Cell(21, 4).Value.ToString());
                        }

                        if(!string.IsNullOrEmpty(worksheet.Cell(22, 4).Value.ToString())) {
                            cmb2Tower.SelectedIndex = cmb2Tower.FindString(worksheet.Cell(22, 4).Value.ToString());
                        }

                        txtstrlc1.Text = worksheet.Cell(16, 6).Value.ToString();
                        txtEp1.Text = worksheet.Cell(16, 7).Value.ToString();
                        txtWBS1.Text = worksheet.Cell(16, 8).Value.ToString();
                        txtFOPO1.Text = worksheet.Cell(16, 9).Value.ToString();
                        txtincoterm1.Text = worksheet.Cell(16, 10).Value.ToString();
                        txtincotermlc1.Text = worksheet.Cell(16, 11).Value.ToString();
                        txtcontactperson1.Text = worksheet.Cell(16, 12).Value.ToString();

                        txtstrlc2.Text = worksheet.Cell(17, 6).Value.ToString();
                        txtEp2.Text = worksheet.Cell(17, 7).Value.ToString();
                        txtWBS2.Text = worksheet.Cell(17, 8).Value.ToString();
                        txtFOPO2.Text = worksheet.Cell(17, 9).Value.ToString();
                        txtincoterm2.Text = worksheet.Cell(17, 10).Value.ToString();
                        txtincotermlc2.Text = worksheet.Cell(17, 11).Value.ToString();
                        txtcontactperson2.Text = worksheet.Cell(17, 12).Value.ToString();

                        txtstrlc3.Text = worksheet.Cell(18, 6).Value.ToString();
                        txtEp3.Text = worksheet.Cell(18, 7).Value.ToString();
                        txtWBS3.Text = worksheet.Cell(18, 8).Value.ToString();
                        txtFOPO3.Text = worksheet.Cell(18, 9).Value.ToString();
                        txtincoterm3.Text = worksheet.Cell(18, 10).Value.ToString();
                        txtincotermlc3.Text = worksheet.Cell(18, 11).Value.ToString();
                        txtcontactperson3.Text = worksheet.Cell(18, 12).Value.ToString();

                        txtstrlc4.Text = worksheet.Cell(19, 6).Value.ToString();
                        txtEp4.Text = worksheet.Cell(19, 7).Value.ToString();
                        txtWBS4.Text = worksheet.Cell(19, 8).Value.ToString();
                        txtFOPO4.Text = worksheet.Cell(19, 9).Value.ToString();
                        txtincoterm4.Text = worksheet.Cell(19, 10).Value.ToString();
                        txtincotermlc4.Text = worksheet.Cell(19, 11).Value.ToString();
                        txtcontactperson4.Text = worksheet.Cell(19, 12).Value.ToString();
                     
                        txtstrlc5.Text = worksheet.Cell(20, 6).Value.ToString();
                        txtEp5.Text = worksheet.Cell(20, 7).Value.ToString();
                        txtWBS5.Text = worksheet.Cell(20, 8).Value.ToString();
                        txtFOPO5.Text = worksheet.Cell(20, 9).Value.ToString();
                        txtincoterm5.Text = worksheet.Cell(20, 10).Value.ToString();
                        txtincotermlc5.Text = worksheet.Cell(20, 11).Value.ToString();
                        txtcontactperson5.Text = worksheet.Cell(20, 12).Value.ToString();
                       
                        txtstrlc6.Text = worksheet.Cell(21, 6).Value.ToString();
                        txtEp6.Text = worksheet.Cell(21, 7).Value.ToString();
                        txtWBS6.Text = worksheet.Cell(21, 8).Value.ToString();
                        txtFOPO6.Text = worksheet.Cell(21, 9).Value.ToString();
                        txtincoterm6.Text = worksheet.Cell(21, 10).Value.ToString();
                        txtincotermlc6.Text = worksheet.Cell(21, 11).Value.ToString();
                        txtcontactperson6.Text = worksheet.Cell(21, 12).Value.ToString();
                       
                        txtstrlc7.Text = worksheet.Cell(22, 6).Value.ToString();
                        txtEp7.Text = worksheet.Cell(22, 7).Value.ToString();
                        txtWBS7.Text = worksheet.Cell(22, 8).Value.ToString();
                        txtFOPO7.Text = worksheet.Cell(22, 9).Value.ToString();
                        txtincoterm7.Text = worksheet.Cell(22, 10).Value.ToString();
                        txtincotermlc7.Text = worksheet.Cell(22, 11).Value.ToString();
                        txtcontactperson7.Text = worksheet.Cell(22, 12).Value.ToString();

                        txtstrlc8.Text = worksheet.Cell(23, 6).Value.ToString();
                        txtEp8.Text = worksheet.Cell(23, 7).Value.ToString();
                        txtWBS8.Text = worksheet.Cell(23, 8).Value.ToString();
                        txtFOPO8.Text = worksheet.Cell(23, 9).Value.ToString();
                        txtincoterm8.Text = worksheet.Cell(23, 10).Value.ToString();
                        txtincotermlc8.Text = worksheet.Cell(23, 11).Value.ToString();
                        txtcontactperson8.Text = worksheet.Cell(23, 12).Value.ToString();

                        //txtRerportFopo.Text = worksheet.Cell(27, 3).Value.ToString();
                        txtProjectUPId.Text = worksheet.Cell(12, 3).Value.ToString();
                        //txtwbsleq.Text = worksheet.Cell(29, 3).Value.ToString();
                        //txtwbsteq.Text = worksheet.Cell(30, 3).Value.ToString();
                        //txtexecutionplant.Text = worksheet.Cell(31, 3).Value.ToString();

                         txt9Name.Text     = worksheet.Cell(12, 7).Value.ToString();
                         txt9Contact.Text  = worksheet.Cell(12, 8).Value.ToString();
                         txt9Gid.Text = worksheet.Cell(12, 9).Value.ToString();

                        txtMainCWBS.Text = worksheet.Cell(16, 13).Value.ToString();
                        txtMainCPO.Text = worksheet.Cell(16, 14).Value.ToString();

                        txtMainCWBS_2.Text = worksheet.Cell(17, 13).Value.ToString();
                        txtMainCPO_2.Text = worksheet.Cell(17, 14).Value.ToString();

                        txtMainCWBS_3.Text = worksheet.Cell(18, 13).Value.ToString();
                        txtMainCPO_3.Text = worksheet.Cell(18, 14).Value.ToString();
                        ToggleNewProjectSection(true);
                    }

                } catch(Exception ex) {
                    Logger.Write($"[{ex.Message}] - {ex.StackTrace}");

                    lblProjName.Text = "";
                    MessageBox.Show("Something went wrong. Close the project report file if it is open.", "Error");
                }
            } else {
                this.Hide();
                var m = new FrmMat($"{rootPath}\\", project ?? "");
                m.ShowDialog();
                this.Close();
            }
        }

        private void label51_Click(object sender, EventArgs e) {

        }

        private void pnlForm_Paint(object sender, PaintEventArgs e) {

        }
    }
}
