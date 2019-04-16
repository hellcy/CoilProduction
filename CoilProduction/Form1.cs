using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CoilProduction
{
    public partial class CoilProduction : Form
    {
        Panel activePanel, activeJobPanel;
        string COILID, line;
        int rowIndex;
        string name = "Sheet1";
        string constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    @"C:\coilProductionFile\JobinProcess.xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

        int RO1800, RO2100, RO2400, RO2700, RO3000, RE1800, RE2100, RE2400, RE2700, RE3000, RO2370, RO3100, RE2370, RE3100, RO2365, RO2710, RE2365, RE2710;
        string Machine, Operator;

        private BindingSource bindingSource1 = new BindingSource();

        public CoilProduction()
        {
            InitializeComponent();
            activePanel = MenuPanel;
            activeJobPanel = POJobDetailsPanel;
            StartPanel.Visible = false;
            JobsPanel.Visible = false;
            FinishPanel.Visible = false;
            POJobDetailsPanel.Visible = false;
            PLJobDetailsPanel.Visible = false;
            RAJobDetailsPanel.Visible = false;

            DataGridViewDisableButtonColumn buttonColumn = new DataGridViewDisableButtonColumn();
            buttonColumn.Text = "Select";
            buttonColumn.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(buttonColumn);

            DataGridViewRow row = new DataGridViewRow();
            ISJobDetailsGrid.Rows.Add();
            ISJobDetailsGrid.Rows[0].Cells[0].Value = 1490;
            ISJobDetailsGrid.Rows.Add();
            ISJobDetailsGrid.Rows[1].Cells[0].Value = 1790;
            ISJobDetailsGrid.Rows.Add();
            ISJobDetailsGrid.Rows[2].Cells[0].Value = 2090;
            ISJobDetailsGrid.Rows.Add();
            ISJobDetailsGrid.Rows[3].Cells[0].Value = 2390;
        }

        // start button on the menu panel
        private void StartButton_Click(object sender, EventArgs e)
        {
            activePanel.Visible = false;
            activePanel = StartPanel;
            activePanel.Visible = true;
        }

        // Finish button on the menu panel
        private void FinishButton_Click(object sender, EventArgs e)
        {
            try
            {
                OleDbConnection con = new OleDbConnection(constr);
                con.Open();
                readJobs(con);
                con.Close();
            }
            catch (Exception ex)
            {
                StartErrMsg.Text = ex.Message;
            }
            activePanel.Visible = false;
            activePanel = JobsPanel;
            activePanel.Visible = true;
        }

        private void coilScanText_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                CoilIDText.Text = "";
                ColorText.Text = "";
                TypeText.Text = "";

                /*  
                 To do: check the slit coil table first, if not found, then check the master coil table.
                 */

                // take the partial of the string up to the underscore as the master coil ID and run the checking
                COILID = coilScanText.Text.Split('_', '+')[0];

                try
                {
                    if ((System.IO.File.Exists(@"C:\coilProductionFile\COIL_MASTER_20190408.csv")) == true)
                    {
                        foreach (string line in System.IO.File.ReadAllLines(@"C:\coilProductionFile\COIL_MASTER_20190408.csv"))
                        {
                            if (COILID.Equals(line.Split(',')[0])) // found coil in the file
                            {
                                string[] lineArray = line.Split(','); // split the input string by using the delimiter ','

                                CoilIDText.Text = coilScanText.Text.Split('+')[0];
                                TypeText.Text = lineArray[1];
                                ColorText.Text = lineArray[2];
                            }
                        }
                        if (CoilIDText.Text == "")
                        {
                            StartErrMsg.Text = "Coil not found, please fill required information.";
                        }
                    }
                }
                catch (Exception ex)
                {
                    StartErrMsg.Text = ex.Message;
                }

            }
        }

        // Back button on the start panel
        private void BackButton_Click(object sender, EventArgs e)
        {
            clearStartPanel();
            activePanel.Visible = false;
            activePanel = MenuPanel;
            activePanel.Visible = true;
        }

        // start button on the start panel
        private void StartJobButton_Click(object sender, EventArgs e)
        {
            string date = DateTime.Now.ToString("yyyy-MM-dd");
            string time = DateTime.Now.ToString("HH:mm:ss");
            string dateTime = date + " " + time;
            string type = TypeText.Text.ToString().ToUpper();

            if (CoilIDText.Text != "" && ColorText.Text != "" && TypeText.Text != "" && machineText.Text != "" && operatorText.Text != "")
            {
                Machine = machineText.Text.ToString();
                Operator = operatorText.Text.ToString();

                if (type != "PO" && type != "PL" && type != "RA" && type != "SP" && type != "IS")
                {
                    StartErrMsg.Text = "Type '" + TypeText.Text.ToUpper() + "' is not supported.";
                    return;
                }
                try
                {
                    OleDbConnection con = new OleDbConnection(constr);
                    con.Open();

                    // write new job to the file
                    object coilValue = CoilIDText.Text.ToUpper();
                    object typeValue = TypeText.Text.ToUpper();
                    object colorValue = ColorText.Text.ToUpper();
                    object timeValue = dateTime;
                    object machienValue = machineText.Text.ToUpper();
                    object operatorValue = operatorText.Text.ToUpper();

                    var commandText = $"Insert Into [" + name + "$] ([Coil ID], [Type], [Color], [Start Time], [Flag], [Machine], [Operator]) Values (@PropertyOne, @PropertyTwo, @PropertyThree, @PropertyFour, 0, @PropertyFive, @PropertySix)";

                    using (var command = new OleDbCommand(commandText, con))
                    {
                        command.Parameters.AddWithValue("@PropertyOne", coilValue ?? DBNull.Value);
                        command.Parameters.AddWithValue("@PropertyTwo", typeValue ?? DBNull.Value);
                        command.Parameters.AddWithValue("@PropertyThree", colorValue ?? DBNull.Value);
                        command.Parameters.AddWithValue("@PropertyFour", timeValue ?? DBNull.Value);
                        command.Parameters.AddWithValue("@PropertyFive", machienValue ?? DBNull.Value);
                        command.Parameters.AddWithValue("@PropertySix", operatorValue ?? DBNull.Value);
                        command.ExecuteNonQuery();
                    }

                    readJobs(con);
                    con.Close();
                }
                catch (Exception ex)
                {
                    StartErrMsg.Text = ex.Message;
                }

                clearStartPanel();
                activePanel.Visible = false;
                activePanel = JobsPanel;
                activePanel.Visible = true;
            }
            else
            {
                //errorProvider1.SetError(CoilIDText, "Please Fill the Name.");
                //errorProvider1.SetError(TypeText, "Please Fill the Name.");
                //errorProvider1.SetError(ColorText, "Please Fill the Name.");
                StartErrMsg.Text = "Please fill required information.";
            }
        }

        // grid button click
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewDisableButtonColumn &&
                e.RowIndex >= 0)
            {
                DataGridViewDisableButtonCell btn = (DataGridViewDisableButtonCell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                if (btn.Enabled == false) return;

                FcoilIDText.UseMnemonic = false; // display & symbol
                FcoilIDText.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                FtypeText.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                FcolorText.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                FstartTimeText.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                FMachineText.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
                FOperatorText.Text = dataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString();

                rowIndex = e.RowIndex;

                switch (FtypeText.Text)
                {
                    case "PO":
                        activeJobPanel.Visible = false;
                        activeJobPanel = POJobDetailsPanel;
                        activeJobPanel.Visible = true;
                        break;
                    case "PL":
                        activeJobPanel.Visible = false;
                        activeJobPanel = PLJobDetailsPanel;
                        activeJobPanel.Visible = true;
                        break;
                    case "RA":
                        activeJobPanel.Visible = false;
                        activeJobPanel = RAJobDetailsPanel;
                        activeJobPanel.Visible = true;
                        break;
                    case "SP":
                        activeJobPanel.Visible = false;
                        activeJobPanel = POJobDetailsPanel;
                        activeJobPanel.Visible = true;
                        break;
                    case "IS":
                        activeJobPanel.Visible = false;
                        activeJobPanel = ISJobDetailsPanel;
                        activeJobPanel.Visible = true;
                        break;
                    default:
                        break;
                }

                activePanel.Visible = false;
                activePanel = FinishPanel;
                activePanel.Visible = true;
            }
        }

        // Back button on the jobs panel
        private void BackButton2_Click(object sender, EventArgs e)
        {
            activePanel.Visible = false;
            activePanel = MenuPanel;
            activePanel.Visible = true;
        }

        // Back button on the finish panel
        private void BackButton3_Click(object sender, EventArgs e)
        {
            switch (FtypeText.Text)
            {
                case "PO":
                    clearPODetails();
                    break;
                case "PL":
                    clearPLDetails();
                    break;
                case "RA":
                    clearRADetails();
                    break;
                case "SP":
                    clearPODetails();
                    break;
                default:
                    break;
            }

            FinishErrMsg.Text = "";
            activePanel.Visible = false;
            activePanel = JobsPanel;
            activePanel.Visible = true;
        }

        // Finish button on the finish panel
        private void FinishButton2_Click(object sender, EventArgs e)
        {
            string date = DateTime.Now.ToString("yyyy-MM-dd");
            string time = DateTime.Now.ToString("HH:mm:ss");
            string dateTime = date + " " + time;
            string path = "";

            // write finished job to a file
            try
            {

                if (isEmptyPLDetails() && isEmptyPODetails() && isEmptyRADetails())
                {
                    FinishErrMsg.Text = "Please fill job details first.";
                    return;
                }

                switch (FtypeText.Text)
                {
                    case "RA":
                        path = @"C:\coilProductionFile\Rail\RA " + date + ".csv";
                        break;
                    case "PO":
                        path = @"C:\coilProductionFile\Channel Post\PO " + date + ".csv";
                        break;
                    case "PL":
                        path = @"C:\coilProductionFile\Plinth\PL " + date + ".csv";
                        break;
                    case "SP":
                        path = @"C:\coilProductionFile\Smart Post\SP " + date + ".csv";
                        break;
                    default:
                        return;
                }
                FileInfo fInfo = new FileInfo(path);
                TextWriter txt = new StreamWriter(path, true); // true means text will be appended to the file.

                switch (FtypeText.Text)
                {
                    case "RA":
                        if (fInfo.Length == 0)
                        {
                            txt.WriteLine("COILID,TYPE,COLOR,2370,Rejected,3100,Rejected,Total,Total Rejected,Start Time,Finish Time");
                        }
                        line = FcoilIDText.Text + "," + FtypeText.Text + "," + FcolorText.Text + "," + RA2370RolledText.Text + "," + RA2370RejectedText.Text + "," + RA3100RolledText.Text + "," + RA3100RejectedText.Text + "," + Total.Text + "," + TotalRejected.Text + "," + FstartTimeText.Text + "," + dateTime;
                        break;
                    case "PO":
                        if (fInfo.Length == 0)
                        {
                            txt.WriteLine("COILID,TYPE,COLOR,1800,Rejected,2100,Rejected,2400,Rejected,2700,Rejected,3000,Rejected,Total,Total Rejected,Start TIme,Finish Time");
                        }
                        line = FcoilIDText.Text + "," + FtypeText.Text + "," + FcolorText.Text + "," + PO1800RolledText.Text + "," + PO1800RejectedText.Text + "," + PO2100RolledText.Text + "," + PO2100RejectedText.Text + "," + PO2400RolledText.Text + "," +
                            PO2400RejectedText.Text + "," + PO2700RolledText.Text + "," + PO2700RejectedText.Text + "," + PO3000RolledText.Text + "," + PO3000RejectedText.Text + "," + Total.Text + "," + TotalRejected.Text + "," + FstartTimeText.Text + "," + dateTime;
                        break;
                    case "PL":
                        if (fInfo.Length == 0)
                        {
                            txt.WriteLine("COILID,TYPE,COLOR,2365,Rejected,2710,Rejected,Total,Total Rejected,Start Time,Finish Time");
                        }
                        line = FcoilIDText.Text + "," + FtypeText.Text + "," + FcolorText.Text + "," + PL2365RolledText.Text + "," + PL2365RejectedText.Text + "," + PL2710RolledText.Text + "," + PL2710RejectedText.Text + "," + Total.Text + "," + TotalRejected.Text + "," + FstartTimeText.Text + "," + dateTime;
                        break;
                    case "SP":
                        if (fInfo.Length == 0)
                        {
                            txt.WriteLine("COILID,TYPE,COLOR,1800,Rejected,2100,Rejected,2400,Rejected,2700,Rejected,3000,Rejected,Total,Total Rejected,Start TIme,Finish Time");
                        }
                        line = FcoilIDText.Text + "," + FtypeText.Text + "," + FcolorText.Text + "," + PO1800RolledText.Text + "," + PO1800RejectedText.Text + "," + PO2100RolledText.Text + "," + PO2100RejectedText.Text + "," + PO2400RolledText.Text + "," +
                            PO2400RejectedText.Text + "," + PO2700RolledText.Text + "," + PO2700RejectedText.Text + "," + PO3000RolledText.Text + "," + PO3000RejectedText.Text + "," + Total.Text + "," + TotalRejected.Text + "," + FstartTimeText.Text + "," + dateTime;
                        break;
                    default:
                        break;
                }

                txt.WriteLine(line);
                txt.Close();
            }
            catch (Exception ex)
            {
                StartErrMsg.Text = ex.Message;
            }

            OleDbConnection con = new OleDbConnection(constr);
            con.Open();

            // update job's end time
            object EndTime = dateTime;
            object COILID = FcoilIDText.Text;

            var commandText = $"Update [" + name + "$] SET [End Time] = (@PropertyOne), [Flag] = 1 WHERE [Coil ID] = (@PropertyTwo)";

            using (var command = new OleDbCommand(commandText, con))
            {
                command.Parameters.AddWithValue("@PropertyOne", EndTime ?? DBNull.Value);
                command.Parameters.AddWithValue("@PropertyTwo", COILID ?? DBNull.Value);
                command.ExecuteNonQuery();
            }

            // read and display file
            readJobs(con);
            con.Close();

            // clear the textboxes
            switch (FtypeText.Text)
            {
                case "PO":
                    clearPODetails();
                    break;
                case "PL":
                    clearPLDetails();
                    break;
                case "RA":
                    clearRADetails();
                    break;
                case "SP":
                    clearPODetails();
                    break;
                default:
                    break;
            }
            FinishErrMsg.Text = "";

            activePanel.Visible = false;
            activePanel = JobsPanel;
            activePanel.Visible = true;
        }

        // read file and display
        private void readJobs(OleDbConnection con)
        {
            //DataTable data = GetDataTableFromExcel(@"C:\coilProductionFile\JobinProcess.xlsx");
            //DataView view = data.DefaultView;
            OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            sda.Fill(data);
            dataGridView1.DataSource = data;

            dataGridView1.DataSource = bindingSource1;
            DataView view = data.DefaultView;
            view.Sort = "Flag ASC, End Time DESC, Start Time DESC";
            bindingSource1.DataSource = view; //rebind the data source

            dataGridView1.Columns[0].Width = 80; // button
            dataGridView1.Columns[1].Width = 180; // CoilID
            dataGridView1.Columns[2].Width = 85; // Type
            dataGridView1.Columns[3].Width = 85; // Color
            dataGridView1.Columns[4].Width = 220; // Start Time
            dataGridView1.Columns[5].Width = 220; // End Time
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void dataGridView1_RowsAdded(object sender, System.Windows.Forms.DataGridViewRowsAddedEventArgs e)
        {
            for (int index = e.RowIndex; index <= e.RowIndex + e.RowCount - 1; index++)
            {
                if (dataGridView1.Rows[index].Cells[5].Value.ToString() != "")
                {
                    DataGridViewDisableButtonCell btn = (DataGridViewDisableButtonCell)dataGridView1.Rows[index].Cells[0];
                    btn.Enabled = false;
                    dataGridView1.Invalidate();
                }
            }
        }

        public static DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.First();
                DataTable tbl = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return tbl;
            }
        }

        private void clearStartPanel()
        {
            CoilIDText.Text = "";
            TypeText.Text = "";
            ColorText.Text = "";
            coilScanText.Text = "";
            StartErrMsg.Text = "";
            machineText.Text = "";
            operatorText.Text = "";
        }

        private void clearPLDetails()
        {
            PL2365RolledText.Text = "";
            PL2365RejectedText.Text = "";
            PL2710RolledText.Text = "";
            PL2710RejectedText.Text = "";
        }

        private void clearPODetails()
        {
            PO1800RolledText.Text = "";
            PO1800RejectedText.Text = "";
            PO2100RolledText.Text = "";
            PO2100RejectedText.Text = "";
            PO2400RolledText.Text = "";
            PO2400RejectedText.Text = "";
            PO2700RolledText.Text = "";
            PO2700RejectedText.Text = "";
            PO3000RolledText.Text = "";
            PO3000RejectedText.Text = "";
        }

        private void clearRADetails()
        {
            RA2370RolledText.Text = "";
            RA2370RejectedText.Text = "";
            RA3100RolledText.Text = "";
            RA3100RejectedText.Text = "";
        }

        private bool isEmptyPLDetails()
        {
            return PL2365RolledText.Text == "" && PL2365RejectedText.Text == "" && PL2710RolledText.Text == "" && PL2710RejectedText.Text == "";
        }

        private bool isEmptyPODetails()
        {
            return PO1800RolledText.Text == "" && PO1800RejectedText.Text == "" && PO2100RolledText.Text == "" && PO2100RejectedText.Text == "" && PO2400RolledText.Text == ""
                 && PO2400RejectedText.Text == "" && PO2700RolledText.Text == "" && PO2700RejectedText.Text == "" && PO3000RolledText.Text == "" && PO3000RejectedText.Text == "";
        }

        private bool isEmptyRADetails()
        {
            return RA2370RolledText.Text == "" && RA2370RejectedText.Text == "" && RA3100RolledText.Text == "" && RA3100RejectedText.Text == "";
        }

        private void calulateTotal()
        {
            checkValue();
            switch (FtypeText.Text)
            {
                case "PO":
                    Total.Text = (RO1800 + RO2100 + RO2400 + RO2700 + RO3000).ToString();
                    TotalRejected.Text = (RE1800 + RE2100 + RE2400 + RE2700 + RE3000).ToString();
                    break;
                case "RA":
                    Total.Text = (RO2370 + RO3100).ToString();
                    TotalRejected.Text = (RE2370 + RE3100).ToString();
                    break;
                case "PL":
                    Total.Text = (RO2365 + RO2710).ToString();
                    TotalRejected.Text = (RE2365 + RE2710).ToString();
                    break;
                case "SP":
                    Total.Text = (RO1800 + RO2100 + RO2400 + RO2700 + RO3000).ToString();
                    TotalRejected.Text = (RE1800 + RE2100 + RE2400 + RE2700 + RE3000).ToString();
                    break;
                default:
                    break;
            }
        }

        private void checkValue()
        {
            RO1800 = (PO1800RolledText.Text == "") ? 0 : Int32.Parse(PO1800RolledText.Text);
            RO2100 = (PO2100RolledText.Text == "") ? 0 : Int32.Parse(PO2100RolledText.Text);
            RO2400 = (PO2400RolledText.Text == "") ? 0 : Int32.Parse(PO2400RolledText.Text);
            RO2700 = (PO2700RolledText.Text == "") ? 0 : Int32.Parse(PO2700RolledText.Text);
            RO3000 = (PO3000RolledText.Text == "") ? 0 : Int32.Parse(PO3000RolledText.Text);

            RE1800 = (PO1800RejectedText.Text == "") ? 0 : Int32.Parse(PO1800RejectedText.Text);
            RE2100 = (PO2100RejectedText.Text == "") ? 0 : Int32.Parse(PO2100RejectedText.Text);
            RE2400 = (PO2400RejectedText.Text == "") ? 0 : Int32.Parse(PO2400RejectedText.Text);
            RE2700 = (PO2700RejectedText.Text == "") ? 0 : Int32.Parse(PO2700RejectedText.Text);
            RE3000 = (PO3000RejectedText.Text == "") ? 0 : Int32.Parse(PO3000RejectedText.Text);

            RO2370 = (RA2370RolledText.Text == "") ? 0 : Int32.Parse(RA2370RolledText.Text);
            RO3100 = (RA3100RolledText.Text == "") ? 0 : Int32.Parse(RA3100RolledText.Text);
            RE2370 = (RA2370RejectedText.Text == "") ? 0 : Int32.Parse(RA2370RejectedText.Text);
            RE3100 = (RA3100RejectedText.Text == "") ? 0 : Int32.Parse(RA3100RejectedText.Text);

            RO2365 = (PL2365RolledText.Text == "") ? 0 : Int32.Parse(PL2365RolledText.Text);
            RO2710 = (PL2710RolledText.Text == "") ? 0 : Int32.Parse(PL2710RolledText.Text);
            RE2365 = (PL2365RejectedText.Text == "") ? 0 : Int32.Parse(PL2365RejectedText.Text);
            RE2710 = (PL2710RejectedText.Text == "") ? 0 : Int32.Parse(PL2710RejectedText.Text);
        }

        // PO
        private void RA2370RolledText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void RA3100RolledText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void RA2370RejectedText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void RA3100RejectedText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void PL2365RolledText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void PL2710RolledText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void PL2365RejectedText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void PL2710RejectedText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void PO1800RolledText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void PO2100RolledText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void PO2400RolledText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void PO2700RolledText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void PO3000RolledText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void PO1800RejectedText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void PO2100RejectedText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void PO2400RejectedText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void PO2700RejectedText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void PO3000RejectedText_TextChanged(object sender, EventArgs e) { calulateTotal(); }
        private void PO1800RolledText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
        private void PO2100RolledText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
        private void PO2400RolledText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
        private void PO2700RolledText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
        private void PO3000RolledText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
        private void PO1800RejectedText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
        private void PO2100RejectedText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
        private void PO2400RejectedText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
        private void PO2700RejectedText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
        private void PO3000RejectedText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
        // PL
        private void PL2365RolledText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
        private void PL2710RolledText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
        private void PL2365RejectedText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
        private void PL2710RejectedText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }

        //RA
        private void RA2370RejectedText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
        private void RA3100RejectedText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
        private void RA2370RolledText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
        private void RA3100RolledText_KeyPress(object sender, KeyPressEventArgs e) { e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar); }
    }
}

