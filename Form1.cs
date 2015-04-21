using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.IO; // File.Exists()
using System.Data.OleDb; // OleDbConnection, OleDbDataAdapter, OleDbCommandBuilder

namespace MS_Access__accdb__in_CSharp
{
    public partial class Form1 : Form
    {
        string DBPath;

        OleDbConnection conn;
        OleDbDataAdapter adapter;
        DataTable dtMain;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DBPath = Application.StartupPath + "\\Data\\Mediation.accdb";

            // create DB via ADOX if not exists
            if (!File.Exists(DBPath))
            {
                ADOX.Catalog cat = new ADOX.Catalog();
                try
                {
                    cat.Create("Provider=Microsoft.ACE.OLEDB.10.0;Data Source=" + DBPath);
                }
                catch
                {
                    try
                    {
                        cat.Create("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DBPath);
                    }
                    catch
                    {
                        try
                        {
                            cat.Create("Provider=Microsoft.ACE.OLEDB.14.0;Data Source=" + DBPath);
                        }
                        catch
                        {
                            cat.Create("Provider=Microsoft.ACE.OLEDB.15.0;Data Source=" + DBPath);
                        }
                    }
                }
                cat = null;
            }

            // connect to DB
            try
            {
                conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.10.0;Data Source=" + DBPath);
                conn.Open();
            }
            catch
            {
                try
                {
                    conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DBPath);
                    conn.Open();
                }
                catch
                {
                    try
                    {
                        conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.14.0;Data Source=" + DBPath);
                        conn.Open();
                    }
                    catch
                    {
                        conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.15.0;Data Source=" + DBPath);
                        conn.Open();
                    }
                }
            }

            // create table "Table_1" if not exists
            // DO NOT USE SPACES IN TABLE AND COLUMNS NAMES TO PREVENT TROUBLES WITH SAVING, USE _
            // OLEDBCOMMANDBUILDER DON'T SUPPORT COLUMNS NAMES WITH SPACES
            try
            {
                using (OleDbCommand cmd = new OleDbCommand("CREATE TABLE [Table_1] ([id] COUNTER PRIMARY KEY, [text_column] MEMO, [int_column] INT);", conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex) {if (ex != null) ex = null; }

            // get all tables from DB
            using (DataTable dt = conn.GetSchema("Tables"))
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i].ItemArray[dt.Columns.IndexOf("TABLE_TYPE")].ToString() == "TABLE")
                    {
                        comboBoxTables.Items.Add(dt.Rows[i].ItemArray[dt.Columns.IndexOf("TABLE_NAME")].ToString());
                    }
                }
            }
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            if (comboBoxTables.SelectedItem == null) return;

            adapter = new OleDbDataAdapter("SELECT * FROM [" + comboBoxTables.SelectedItem.ToString() + "]", conn);
            
            new OleDbCommandBuilder(adapter);

            dtMain = new DataTable();
            adapter.Fill(dtMain);
            dtMain.Columns[0].ReadOnly = true; // deprecate id field edit to prevent exceptions
            dataGridView1.DataSource = dtMain;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (adapter == null) return;

            adapter.Update(dtMain);
        }

        // show tooltip (not intrusive MessageBox) when user trying to input letters into INT column cell
        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (dtMain.Columns[e.ColumnIndex].DataType == typeof(Int64) ||
                dtMain.Columns[e.ColumnIndex].DataType == typeof(Int32) ||
                dtMain.Columns[e.ColumnIndex].DataType == typeof(Int16))
            {
                Rectangle rectColumn;
                rectColumn = dataGridView1.GetColumnDisplayRectangle(e.ColumnIndex, false);

                Rectangle rectRow;
                rectRow = dataGridView1.GetRowDisplayRectangle(e.RowIndex, false);

                toolTip1.ToolTipTitle = "This field is for numbers only.";
                toolTip1.Show(" ",
                          dataGridView1,
                          rectColumn.Left, rectRow.Top + rectRow.Height);
            }
        }

        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            toolTip1.Hide(dataGridView1);
        }
    }
}
