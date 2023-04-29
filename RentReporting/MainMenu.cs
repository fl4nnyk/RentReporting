using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace RentReporting
{
    public partial class MainMenu : Form
    {
        private SqlConnection sqlConnection = null;
        private SqlCommandBuilder sqlCommandBuilder = null;
        private SqlDataAdapter sqlDataAdapter = null;
        private DataSet dataSet = null;

        private bool NewRowAdding = false;

        private void LoadData()
        {
            try
            {
                sqlDataAdapter = new SqlDataAdapter("SELECT *, 'Delete' AS [Command] FROM Exchanges", sqlConnection);
                sqlCommandBuilder = new SqlCommandBuilder(sqlDataAdapter);

                sqlCommandBuilder.GetInsertCommand();
                sqlCommandBuilder.GetUpdateCommand();
                sqlCommandBuilder.GetDeleteCommand();

                dataSet = new DataSet();

                sqlDataAdapter.Fill(dataSet, "Exchanges");

                dataGridView.DataSource = dataSet.Tables["Exchanges"];

                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView[8, i] = linkCell;
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ReloadData()
        {
            try
            {
                dataSet.Tables["Exchanges"].Clear();

                sqlDataAdapter.Fill(dataSet, "Exchanges");

                dataGridView.DataSource = dataSet.Tables["Exchanges"];

                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView[8, i] = linkCell;
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public MainMenu()
        {
            InitializeComponent();
        }

        private void MainMenu_Load(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=""C:\Users\dimap\OneDrive\Рабочий стол\KR\RentReporting\RentReporting\Database.mdf"";Integrated Security=True");

            sqlConnection.Open();

            LoadData();
        }

        private void Update_toolStripButton_Click(object sender, EventArgs e)
        {
            ReloadData();

        }

        private void dataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 8)
                {
                    string task = dataGridView.Rows[e.RowIndex].Cells[8].Value.ToString();

                    if (task == "Delete")
                    {
                        if (MessageBox.Show("Ви справді хочете видалити цю строку?", "Видалення", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            int RowIndex = e.RowIndex;

                            dataGridView.Rows.RemoveAt(e.RowIndex);

                            dataSet.Tables["Exchanges"].Rows[RowIndex].Delete();

                            sqlDataAdapter.Update(dataSet, "Exchanges");
                        }
                    }
                    else if (task == "Insert")
                    {
                        int RowIndex = dataGridView.Rows.Count - 2;

                        DataRow row = dataSet.Tables["Exchanges"].NewRow();

                        row["Id"] = dataGridView.Rows[RowIndex].Cells["Id"].Value;
                        row["Agent"] = dataGridView.Rows[RowIndex].Cells["Agent"].Value;
                        row["Address"] = dataGridView.Rows[RowIndex].Cells["Address"].Value;
                        row["Rent"] = dataGridView.Rows[RowIndex].Cells["Rent"].Value;
                        row["Security"] = dataGridView.Rows[RowIndex].Cells["Security"].Value;
                        row["RRO"] = dataGridView.Rows[RowIndex].Cells["RRO"].Value;
                        row["Utillity Company"] = dataGridView.Rows[RowIndex].Cells["Utillity Company"].Value;
                        row["Salary"] = dataGridView.Rows[RowIndex].Cells["Salary"].Value;

                        dataSet.Tables["Exchanges"].Rows.Add(row);

                        dataSet.Tables["Exchanges"].Rows.RemoveAt(dataSet.Tables["Exchanges"].Rows.Count - 1);

                        dataGridView.Rows.RemoveAt(dataGridView.Rows.Count - 2);

                        dataGridView.Rows[e.RowIndex].Cells[8].Value = "Delete";

                        sqlDataAdapter.Update(dataSet, "Exchanges");

                        NewRowAdding = false;
                    }
                    else if (task == "Update")
                    {
                        int row = e.RowIndex;

                        dataSet.Tables["Exchanges"].Rows[row]["Id"] = dataGridView.Rows[row].Cells["Id"].Value;
                        dataSet.Tables["Exchanges"].Rows[row]["Agent"] = dataGridView.Rows[row].Cells["Agent"].Value;
                        dataSet.Tables["Exchanges"].Rows[row]["Address"] = dataGridView.Rows[row].Cells["Address"].Value;
                        dataSet.Tables["Exchanges"].Rows[row]["Rent"] = dataGridView.Rows[row].Cells["Rent"].Value;
                        dataSet.Tables["Exchanges"].Rows[row]["Security"] = dataGridView.Rows[row].Cells["Security"].Value;
                        dataSet.Tables["Exchanges"].Rows[row]["RRO"] = dataGridView.Rows[row].Cells["RRO"].Value;
                        dataSet.Tables["Exchanges"].Rows[row]["Utillity Company"] = dataGridView.Rows[row].Cells["Utillity Company"].Value;
                        dataSet.Tables["Exchanges"].Rows[row]["Salary"] = dataGridView.Rows[row].Cells["Salary"].Value;

                        sqlDataAdapter.Update(dataSet, "Exchanges");

                        dataGridView.Rows[e.RowIndex].Cells[8].Value = "Delete";
                    }

                    ReloadData();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            try
            {
                if (NewRowAdding == false)
                {
                    NewRowAdding = true;

                    int lastRow = dataGridView.Rows.Count - 2;

                    DataGridViewRow row = dataGridView.Rows[lastRow];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView[8, lastRow] = linkCell;

                    row.Cells["Command"].Value = "Insert";
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (NewRowAdding == false)
                {
                    int RowIndex = dataGridView.SelectedCells[0].RowIndex;

                    DataGridViewRow editingRow = dataGridView.Rows[RowIndex];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView[8, RowIndex] = linkCell;

                    editingRow.Cells["Command"].Value = "Update";
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(ColumnKeyPress);

            if (dataGridView.CurrentCell.ColumnIndex != 1 && dataGridView.CurrentCell.ColumnIndex != 2)
            {
                TextBox textBox = e.Control as TextBox;

                if (textBox != null)
                {
                    textBox.KeyPress += new KeyPressEventHandler(ColumnKeyPress);
                }
            }
        }

        private void ColumnKeyPress(object sender, KeyPressEventArgs e)
        {
            if (dataGridView.CurrentCell.ColumnIndex != 1 && dataGridView.CurrentCell.ColumnIndex != 2)
            {

                if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void ExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();

            excelApp.Workbooks.Add();

            Excel.Worksheet worksheet = (Excel.Worksheet)excelApp.ActiveSheet;

            for (int i = 0; i < dataGridView.RowCount - 2; i++)
            {
                for (int j = 0; j < dataGridView.ColumnCount - 1; j++)
                {
                    worksheet.Cells[1, j + 1] = dataGridView.Columns[j].HeaderText.ToString();
                    worksheet.Cells[i + 1, j + 1] = dataGridView[i, j].Value.ToString();
                }
            }

            excelApp.Visible = true;
        }
    }
}
