using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Odbc;
using Bordereaux_SICS_Mapping.BAL;

namespace Bordereaux_SICS_Mapping.Forms
{
    public partial class MacroDB : Form
    {
        HelperV21 objHlpr2 = new HelperV21();

        public MacroDB()
        {
            InitializeComponent();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {

            try
            {
                string strSearch = txtSearch.Text;
                string strMacroDB = cmbxMacro.SelectedItem.ToString();
                string Dbconnection = "DSN=SICS_Postgres_DB;" + "UID=sics;" + "PWD=sics_1";
                string query;
                OdbcConnection cnDB = new OdbcConnection(Dbconnection);

                if (string.IsNullOrEmpty(strSearch))
                {
                    var dialog = MessageBox.Show("Input a data to search","Macro Database Search Failed", MessageBoxButtons.OK);
                }
                else if (strSearch == "*" && !string.IsNullOrEmpty(strMacroDB))
                {
                    if (string.IsNullOrEmpty(strMacroDB))
                    {
                        var dialog = MessageBox.Show("Choose a Database", "Choose Database", MessageBoxButtons.OK);
                    }
                    else
                    {

                        query = "SELECT * FROM " + strMacroDB.ToLower();
                        cnDB.Open();
                        OdbcCommand DbCommand = cnDB.CreateCommand();
                        DbCommand.CommandText = query;
                        OdbcDataReader DbReader = DbCommand.ExecuteReader();

                        listView1.Columns.Clear();
                        listView1.Items.Clear();
                        listView1.GridLines = true;
                        listView1.View = View.Details;
                        listView1.Columns.Add("Cession No");
                        listView1.Columns.Add("URC");
                        listView1.Columns.Add("Policy No");
                        listView1.Columns.Add("Cession Type Code");
                        listView1.Columns.Add("Currency Code");
                        listView1.Columns.Add("Issue Age");
                        listView1.Columns.Add("Issue Date");
                        listView1.Columns.Add("Mortality Rating Code");
                        listView1.Columns.Add("Refunding Code");
                        listView1.Columns.Add("NAME");
                        listView1.Columns.Add("Date of Birth");
                        listView1.Columns.Add("Gender");
                        listView1.Columns.Add("Cover7c");
                        listView1.Columns.Add("Benefit");
                        listView1.Columns.Add("amt7insrd7a");
                        listView1.Columns.Add("amt7reinsrd7a");
                        listView1.Columns.Add("ced7retn7a");
                        listView1.Columns.Add("Company Name");
                        listView1.Columns.Add("Source Code");
                        listView1.AutoResizeColumns((ColumnHeaderAutoResizeStyle.HeaderSize));

                        while (DbReader.Read())
                        {
                            ListViewItem lv = new ListViewItem(DbReader[1].ToString());
                            lv.SubItems.Add(DbReader[2].ToString());
                            lv.SubItems.Add(DbReader[3].ToString());
                            lv.SubItems.Add(DbReader[4].ToString());
                            lv.SubItems.Add(DbReader[5].ToString());
                            lv.SubItems.Add(DbReader[6].ToString());
                            lv.SubItems.Add(DbReader[7].ToString());
                            lv.SubItems.Add(DbReader[8].ToString());
                            lv.SubItems.Add(DbReader[9].ToString());
                            lv.SubItems.Add(DbReader[10].ToString());
                            lv.SubItems.Add(DbReader[11].ToString());
                            lv.SubItems.Add(DbReader[12].ToString());
                            lv.SubItems.Add(DbReader[13].ToString());
                            lv.SubItems.Add(DbReader[14].ToString());
                            lv.SubItems.Add(DbReader[15].ToString());
                            lv.SubItems.Add(DbReader[16].ToString());
                            lv.SubItems.Add(DbReader[17].ToString());
                            lv.SubItems.Add(DbReader[18].ToString());
                            lv.SubItems.Add(DbReader[19].ToString());
                            listView1.Items.Add(lv);

                        }
                        cnDB.Close();
                        var dialog = MessageBox.Show("Search Completed","Completed", MessageBoxButtons.OK);
                    }

                }
                else
                {
                    string strDataType = cmbxDataType.SelectedItem.ToString();
                    query = "SELECT * FROM " + strMacroDB.ToLower() + " WHERE " + strDataType + " LIKE " +"'" + strSearch + "%'";
                    cnDB.Open();
                    OdbcCommand DbCommand = cnDB.CreateCommand();
                    DbCommand.CommandText = query;
                    OdbcDataReader DbReader = DbCommand.ExecuteReader();

                    listView1.Columns.Clear();
                    listView1.Items.Clear();
                    listView1.GridLines = true;
                    listView1.View = View.Details;
                    listView1.Columns.Add("Cession No");
                    listView1.Columns.Add("URC");
                    listView1.Columns.Add("Policy No");
                    listView1.Columns.Add("Cession Type Code");
                    listView1.Columns.Add("Currency Code");
                    listView1.Columns.Add("Issue Age");
                    listView1.Columns.Add("Issue Date");
                    listView1.Columns.Add("Mortality Rating Code");
                    listView1.Columns.Add("Refunding Code");
                    listView1.Columns.Add("NAME");
                    listView1.Columns.Add("Date of Birth");
                    listView1.Columns.Add("Gender");
                    listView1.Columns.Add("Cover7c");
                    listView1.Columns.Add("Benefit");
                    listView1.Columns.Add("amt7insrd7a");
                    listView1.Columns.Add("amt7reinsrd7a");
                    listView1.Columns.Add("ced7retn7a");
                    listView1.Columns.Add("Company Name");
                    listView1.Columns.Add("Source Code");
                    listView1.AutoResizeColumns((ColumnHeaderAutoResizeStyle.HeaderSize));

                    while (DbReader.Read())
                    {
                        ListViewItem lv = new ListViewItem(DbReader[1].ToString());
                        lv.SubItems.Add(DbReader[2].ToString());
                        lv.SubItems.Add(DbReader[3].ToString());
                        lv.SubItems.Add(DbReader[4].ToString());
                        lv.SubItems.Add(DbReader[5].ToString());
                        lv.SubItems.Add(DbReader[6].ToString());
                        lv.SubItems.Add(DbReader[7].ToString());
                        lv.SubItems.Add(DbReader[8].ToString());
                        lv.SubItems.Add(DbReader[9].ToString());
                        lv.SubItems.Add(DbReader[10].ToString());
                        lv.SubItems.Add(DbReader[11].ToString());
                        lv.SubItems.Add(DbReader[12].ToString());
                        lv.SubItems.Add(DbReader[13].ToString());
                        lv.SubItems.Add(DbReader[14].ToString());
                        lv.SubItems.Add(DbReader[15].ToString());
                        lv.SubItems.Add(DbReader[16].ToString());
                        lv.SubItems.Add(DbReader[17].ToString());
                        lv.SubItems.Add(DbReader[18].ToString());
                        lv.SubItems.Add(DbReader[19].ToString());
                        listView1.Items.Add(lv);

                    }
                    cnDB.Close();
                    var dialog = MessageBox.Show("Search Completed", "Completed", MessageBoxButtons.OK);
                }
               
            }
            catch (Exception ex)
            {
                if (string.IsNullOrEmpty(cmbxDataType.SelectedItem.ToString()))
                {
                    var dialog = MessageBox.Show("Choose a Data Type", "Datatype Failed", MessageBoxButtons.OK);
                }
                else
                {
                    var dialog = MessageBox.Show("Choose a Database", "Database Failed", MessageBoxButtons.OK);
                }
                
            }
        }

        private void MacroDB_Load(object sender, EventArgs e)
        {
            listView1.GridLines = true;
            listView1.FullRowSelect = true;
            listView1.LabelEdit = true;
        }

        private void btnClear_Click(object sender, EventArgs e)
        {

            var dialog = MessageBox.Show("Refresh the screen?", "Refresh", MessageBoxButtons.YesNo);
            if (dialog == DialogResult.Yes)
            {
               listView1.Clear();
               txtSearch.Clear();
               cmbxDataType.SelectedIndex = -1;
               cmbxMacro.SelectedIndex = -1;

            }
            else if (dialog == DialogResult.No)
            {
                //do nothing
            }

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}

