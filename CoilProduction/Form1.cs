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

            // add the finish button to the grid
            DataGridViewButtonColumn buttonColumn = new DataGridViewButtonColumn() { Name = "Finish Job", DataPropertyName = "NewColumnData" };
            buttonColumn.Text = "Finish";
            buttonColumn.UseColumnTextForButtonValue = true;

            dataGridView1.Columns.Add(buttonColumn);
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

        // start button on the start panel
        private void StartJobButton_Click(object sender, EventArgs e)
        {
            string date = DateTime.Now.ToString("yyyy-MM-dd");
            string time = DateTime.Now.ToString("HH:mm:ss");
            string dateTime = date + " " + time;
            string type = TypeText.Text.ToString().ToUpper();

            if (CoilIDText.Text != "" && ColorText.Text != "" && TypeText.Text != "")
            {
                if (type != "PO" && type != "PL" && type != "RA")
                {
                    StartErrMsg.Text = "Type '" + TypeText.Text.ToUpper() + "' is not supported.";
                    return;
                }
                try
                {
                    // open connection
                    String name = "Sheet1";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    @"C:\coilProductionFile\JobinProcess.xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    con.Open();

                    // write new job to the file
                    object coilValue = CoilIDText.Text.ToUpper();
                    object typeValue = TypeText.Text.ToUpper();
                    object colorValue = ColorText.Text.ToUpper();
                    object timeValue = dateTime;

                    var commandText = $"Insert Into [" + name + "$] ([Coil ID], [Type], [Color], [Start Time]) Values (@PropertyOne, @PropertyTwo, @PropertyThree, @PropertyFour)";

                    using (var command = new OleDbCommand(commandText, con))
                    {
                        command.Parameters.AddWithValue("@PropertyOne", coilValue ?? DBNull.Value);
                        command.Parameters.AddWithValue("@PropertyTwo", typeValue ?? DBNull.Value);
                        command.Parameters.AddWithValue("@PropertyThree", colorValue ?? DBNull.Value);
                        command.Parameters.AddWithValue("@PropertyFour", timeValue ?? DBNull.Value);
                        command.ExecuteNonQuery();
                    }

                    // read file and display
                    OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);

                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable data = new DataTable();
                    sda.Fill(data);
                    dataGridView1.DataSource = data;

                    dataGridView1.Columns[0].Width = 150;
                    dataGridView1.Columns[1].Width = 230;
                    dataGridView1.Columns[2].Width = 100;
                    dataGridView1.Columns[3].Width = 100;
                    dataGridView1.Columns[4].Width = 280;

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
            else StartErrMsg.Text = "Please fill required information.";
        }

        // grid button click
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn &&
                e.RowIndex >= 0)
            {
                //TODO - Button Clicked - Execute Code Here
                FcoilIDText.UseMnemonic = false; // display & symbol
                FcoilIDText.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                FtypeText.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                FcolorText.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                FstartTimeText.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                rowIndex = e.RowIndex;

                switch(FtypeText.Text)
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
                    default:
                        break;
                }

                activePanel.Visible = false;
                activePanel = FinishPanel;
                activePanel.Visible = true;
            }
        }

        /* 
         Button redirect functions
             */

        // start button on the menu panel
        private void StartButton_Click(object sender, EventArgs e)
        {
            activePanel.Visible = false;
            activePanel = StartPanel;
            activePanel.Visible = true;
        }

        // Back button on the start panel
        private void BackButton_Click(object sender, EventArgs e)
        {
            clearStartPanel();
            activePanel.Visible = false;
            activePanel = MenuPanel;
            activePanel.Visible = true;
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
                switch (FtypeText.Text)
                {
                    case "RA":
                        path = @"C:\coilProductionFile\Rail or Smart Post\RA " + date + ".csv";
                        break;
                    case "PO":
                        path = @"C:\coilProductionFile\Channel Post\PO " + date + ".csv";
                        break;
                    case "PL":
                        path = @"C:\coilProductionFile\Plinth\PL " + date + ".csv";
                        break;
                    default:
                        path = @"C:\coilProductionFile\" + date + ".csv";
                        break;
                }
                FileInfo fInfo = new FileInfo(path);
                TextWriter txt = new StreamWriter(path, true); // true means text will be appended to the file.

                switch (FtypeText.Text)
                {
                    case "RA":
                        if (fInfo.Length == 0)
                        {
                            txt.WriteLine("COILID,TYPE,COLOR,2370,Rejected,3100,Rejected,Start Time,Finish Time");
                        }
                        line = FcoilIDText.Text + "," + FtypeText.Text + "," + FcolorText.Text + "," + RA2370RolledText.Text + "," + RA2370RejectedText.Text + "," + RA3100RolledText.Text + "," + RA3100RejectedText.Text + "," + FstartTimeText.Text + "," + dateTime;
                        break;
                    case "PO":
                        if (fInfo.Length == 0)
                        {
                            txt.WriteLine("COILID,TYPE,COLOR,1800,Rejected,2100,Rejected,2400,Rejected,2700,Rejected,3000,Rejected,Start TIme,Finish Time");
                        }
                        line = FcoilIDText.Text + "," + FtypeText.Text + "," + FcolorText.Text + "," + PO1800RolledText.Text + "," + PO1800RejectedText.Text + "," + PO2100RolledText.Text + "," + PO2100RejectedText.Text + "," + PO2400RolledText.Text + "," +
                            PO2400RejectedText.Text + "," + PO2700RolledText.Text + "," + PO2700RejectedText.Text + "," + PO3000RolledText.Text + "," + PO3000RejectedText.Text + "," + FstartTimeText.Text + "," + dateTime;
                        break;
                    case "PL":
                        if (fInfo.Length == 0)
                        {
                            txt.WriteLine("COILID,TYPE,COLOR,2365,Rejected,2710,Rejected,Start Time,Finish Time");
                        }
                        line = FcoilIDText.Text + "," + FtypeText.Text + "," + FcolorText.Text + "," + PL2365RolledText.Text + "," + PL2365RejectedText.Text + "," + PL2710RolledText.Text + "," + PL2710RejectedText.Text + "," + FstartTimeText.Text + "," + dateTime;
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

            // remove finished job from JobinProcess excel file
            try
            {
                //Open the workbook (or create it if it doesn't exist)
                var fi = new FileInfo(@"C:\coilProductionFile\JobinProcess.xlsx");

                using (var p = new ExcelPackage(fi))
                {
                    //Get the Worksheet created in the previous codesample. 
                    var ws = p.Workbook.Worksheets["Sheet1"];
                    ws.DeleteRow(rowIndex + 2, 1, true);
                    p.Save();
                }
            }
            catch (Exception ex)
            {
                StartErrMsg.Text = ex.Message;
            }

            activePanel.Visible = false;
            activePanel = MenuPanel;
            activePanel.Visible = true;
        }

        // Finish button on the menu panel
        private void FinishButton_Click(object sender, EventArgs e)
        {
            try
            {
                // open connection
                String name = "Sheet1";
                String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                @"C:\coilProductionFile\JobinProcess.xlsx" +
                                ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                OleDbConnection con = new OleDbConnection(constr);
                con.Open();

                // read file and display
                OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);

                OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                DataTable data = new DataTable();
                sda.Fill(data);
                dataGridView1.DataSource = data;

                dataGridView1.Columns[0].Width = 150;
                dataGridView1.Columns[1].Width = 230;
                dataGridView1.Columns[2].Width = 100;
                dataGridView1.Columns[3].Width = 100;
                dataGridView1.Columns[4].Width = 280;

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

        private void clearStartPanel()
        {
            CoilIDText.Text = "";
            TypeText.Text = "";
            ColorText.Text = "";
            coilScanText.Text = "";
            StartErrMsg.Text = "";
        }
    }
}
