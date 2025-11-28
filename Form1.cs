using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SQLite;
using System.IO;
using System.Windows.Forms;

namespace CropYieldPrediction
{
    public partial class Form1 : Form
    {
        private string connectionString = "Data Source=crops.db;Version=3;";
        private DataTable importedData = new DataTable();
        private Button btnSignOut = new Button(); // Sign out button

        public Form1()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized; // Open in full screen
            this.FormBorderStyle = FormBorderStyle.None; // Hide title bar
            this.Load += Form1_Load;
            this.Resize += Form1_Resize;
            SetupSignOutButton();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            AlignControls();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            AlignControls();
        }

        private void AlignControls()
        {
            int formWidth = this.ClientSize.Width;
            int formHeight = this.ClientSize.Height;

            int controlWidth = formWidth / 3; // Set width to 1/3 of form
           
            int controlHeight = formHeight / 3;
            txtFilePath.Width = controlWidth;
            btnBrowse.Width = controlWidth / 3;
            btnImport.Width = controlWidth / 3;

            // Manually set positions
            lblFilePath.Location = new System.Drawing.Point(150, 60);
            lblFilePath.AutoSize = true;
            lblFilePath.Font = new System.Drawing.Font("Arial", 18, System.Drawing.FontStyle.Bold);

            txtFilePath.Location = new System.Drawing.Point(400, 60);
            btnBrowse.Location = new System.Drawing.Point(440, 90);
            btnImport.Location = new System.Drawing.Point(620, 90);

            // Increase button sizes
            btnBrowse.Height = 40;
            btnImport.Height = 40;
            btnSignOut.Height = 50;
            btnSignOut.Width += 20;

            // Adjust DataGridView
            dgvData.Size = new System.Drawing.Size(formWidth - 40, formHeight - 200);
            dgvData.Location = new System.Drawing.Point(20, 200);

            btnSignOut.Location = new System.Drawing.Point(formWidth - 230, 20); // Top right corner
        }

        private void SetupSignOutButton()
        {
            btnSignOut.Text = "Sign Out";
            btnSignOut.Width = 100;
            btnSignOut.Height = 40;
            btnSignOut.BackColor = System.Drawing.Color.Red;
            btnSignOut.ForeColor = System.Drawing.Color.White;
            btnSignOut.FlatStyle = FlatStyle.Flat;
            btnSignOut.Font = new System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold);
            btnSignOut.Click += BtnSignOut_Click;
            this.Controls.Add(btnSignOut);
        }

        private void BtnSignOut_Click(object sender, EventArgs e)
        {
            selection ss = new selection();
            ss.Show();
            this.Hide();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|.xlsx;.xls",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = openFileDialog.FileName;
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFilePath.Text) || !File.Exists(txtFilePath.Text))
            {
                MessageBox.Show("Please select a valid Excel file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DataTable dt = ReadExcel(txtFilePath.Text);
            if (dt.Rows.Count > 0)
            {
                InsertDataIntoSQLite(dt);
                importedData = dt;
                MessageBox.Show("Data imported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadLatestDataIntoGrid();
            }
            else
            {
                MessageBox.Show("No data found in the Excel file.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private DataTable ReadExcel(string filePath)
        {
            DataTable dt = new DataTable();
            string fileExtension = Path.GetExtension(filePath);
            string connString;

            if (fileExtension == ".xls")
            {
                connString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={filePath};Extended Properties='Excel 8.0;HDR=YES;'";
            }
            else
            {
                connString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR=YES;'";
            }

            using (OleDbConnection conn = new OleDbConnection(connString))
            {
                try
                {
                    conn.Open();
                    DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (schemaTable == null || schemaTable.Rows.Count == 0) return dt;

                    string sheetName = schemaTable.Rows[0]["TABLE_NAME"].ToString();
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter($"SELECT * FROM [{sheetName}]", conn))
                    {
                        adapter.Fill(dt);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error reading Excel file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            return dt;
        }

        private void InsertDataIntoSQLite(DataTable dt)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();

                // Ensure the table exists before inserting data
                string createTableQuery = @"CREATE TABLE IF NOT EXISTS CropData (
                                            ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                            State TEXT,
                                            District TEXT,
                                            Location TEXT,
                                            SoilType TEXT,
                                            FastMovingRating INTEGER,
                                            Crop TEXT,
                                            PH REAL);";
                using (SQLiteCommand createCmd = new SQLiteCommand(createTableQuery, conn))
                {
                    createCmd.ExecuteNonQuery();
                }

                using (SQLiteTransaction transaction = conn.BeginTransaction())
                {
                    using (SQLiteCommand cmd = new SQLiteCommand(conn))
                    {
                        cmd.CommandText = "INSERT INTO CropData (State, District, Location, SoilType, FastMovingRating, Crop, PH) " +
                                          "VALUES (@State, @District, @Location, @SoilType, @FastMovingRating, @Crop, @PH)";

                        foreach (DataRow row in dt.Rows)
                        {
                            cmd.Parameters.Clear();
                            cmd.Parameters.AddWithValue("@State", row["State"].ToString());
                            cmd.Parameters.AddWithValue("@District", row["District"].ToString());
                            cmd.Parameters.AddWithValue("@Location", row["Location"].ToString());
                            cmd.Parameters.AddWithValue("@SoilType", row["Soil Type"].ToString());
                            cmd.Parameters.AddWithValue("@FastMovingRating", Convert.ToInt32(row["FastMovingRating"]));
                            cmd.Parameters.AddWithValue("@Crop", row["Crop"].ToString());
                            cmd.Parameters.AddWithValue("@PH", Convert.ToDouble(row["PH"]));
                            cmd.ExecuteNonQuery();
                        }
                    }
                    transaction.Commit();
                }
            }
        }

        private void LoadLatestDataIntoGrid()
        {
            dgvData.DataSource = null;
            dgvData.Columns.Clear();
            dgvData.DataSource = importedData;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void btnBrowse_Click_1(object sender, EventArgs e)
        {

        }

        private void btnImport_Click_1(object sender, EventArgs e)
        {

        }

    }
}