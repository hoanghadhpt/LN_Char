using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel;
using System.Globalization;
using System.Threading;

namespace Char2
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
            dt_load();
        }

        private void dt_load()
        {
            DataSet result;
            FileStream fs = File.Open(@"\\172.19.3.129\Program\LN Programs\!!!!SAVE\hh\kt2.xls", FileMode.Open, FileAccess.Read);
            IExcelDataReader rd;
            rd = ExcelReaderFactory.CreateBinaryReader(fs); //read
            rd.IsFirstRowAsColumnNames = true;
            result = rd.AsDataSet();
            dataGridView1.DataSource = result.Tables[0];
            dataGridView1.Columns["NAME"].Frozen = true;
            rd.Close();
            // TOPTEN
            var rcount = dataGridView1.Rows.Count;
            for (int irow = 0;irow< rcount; irow++)
            {
                var rankrow = dataGridView1.Rows[irow].Cells[40].Value;
                if (rankrow.ToString() == "1")
                {
                    stt1.Text = dataGridView1.Rows[irow].Cells[1].Value.ToString();
                }
                if (rankrow.ToString() == "2")
                {
                    stt2.Text = dataGridView1.Rows[irow].Cells[1].Value.ToString();
                }
                if (rankrow.ToString() == "3")
                {
                    stt3.Text = dataGridView1.Rows[irow].Cells[1].Value.ToString();
                }
                if (rankrow.ToString() == "4")
                {
                    stt4.Text = dataGridView1.Rows[irow].Cells[1].Value.ToString();
                }
                if (rankrow.ToString() == "5")
                {
                    stt5.Text = dataGridView1.Rows[irow].Cells[1].Value.ToString();
                }
                if (rankrow.ToString() == "6")
                {
                    stt6.Text = dataGridView1.Rows[irow].Cells[1].Value.ToString();
                }
                if (rankrow.ToString() == "7")
                {
                    stt7.Text = dataGridView1.Rows[irow].Cells[1].Value.ToString();
                }
                if (rankrow.ToString() == "8")
                {
                    stt8.Text = dataGridView1.Rows[irow].Cells[1].Value.ToString();
                }
                if (rankrow.ToString() == "9")
                {
                    stt9.Text = dataGridView1.Rows[irow].Cells[1].Value.ToString();
                }
                if (rankrow.ToString() == "10")
                {
                    stt10.Text = dataGridView1.Rows[irow].Cells[1].Value.ToString();
                }

            }


            // END Topten
        }

        private void tbID1_TextChanged(object sender, EventArgs e)
        {
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("ID LIKE '{0}%'", tbID1.Text);
            txtID.Text = "" + dataGridView1.Rows[0].Cells[0].Value;
            txtName.Text = "" + dataGridView1.Rows[0].Cells[1].Value;
            txtGroup.Text = "" + dataGridView1.Rows[0].Cells[2].Value;
            txtSum.Text = "" + string.Format("{0:n0}", dataGridView1.Rows[0].Cells[37].Value);
            txtSal.Text = "" + string.Format("{0:n0}", dataGridView1.Rows[0].Cells[39].Value);
            txtRank.Text = "" + string.Format("{0:n0}", dataGridView1.Rows[0].Cells[40].Value);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("Group LIKE '{0}%'", comboBox1.Text);
            this.tbID1.Focus();
        }


        private void button1_Click(object sender, EventArgs e)
        {
                if (panel2.Width <= 69)
                {
                    for (int i = 34; i <= 340; i++)
                    { 
                        panel2.Width = i;
                        panel2.Refresh();
                    }
                }
                else
                {
                    for (int i = 340; i >= 34; i--)
                    {
                        panel2.Width = i;
                        panel2.Refresh();
                    }
                }
            
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1) return;
            if (dataGridView1.CurrentCell.ColumnIndex == 0)
            {
                var rowindex = dataGridView1.CurrentCell.RowIndex;
                txtID.Text = "" + dataGridView1.Rows[rowindex].Cells[0].Value;
                txtName.Text = "" + dataGridView1.Rows[rowindex].Cells[1].Value;
                txtGroup.Text = "" + dataGridView1.Rows[rowindex].Cells[2].Value;
                txtSum.Text = "" + string.Format("{0:n0}", dataGridView1.Rows[rowindex].Cells[37].Value);
                txtSal.Text = "" + string.Format("{0:n0}", dataGridView1.Rows[rowindex].Cells[39].Value);
                txtRank.Text = "" + string.Format("{0:n0}", dataGridView1.Rows[rowindex].Cells[40].Value);
            }
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                var rowindex = dataGridView1.CurrentCell.RowIndex;
                txtID.Text = "" + dataGridView1.Rows[rowindex].Cells[0].Value;
                txtName.Text = "" + dataGridView1.Rows[rowindex].Cells[1].Value;
                txtGroup.Text = "" + dataGridView1.Rows[rowindex].Cells[2].Value;
                txtSum.Text = "" + string.Format("{0:n0}", dataGridView1.Rows[rowindex].Cells[37].Value);
                txtSal.Text = "" + string.Format("{0:n0}", dataGridView1.Rows[rowindex].Cells[39].Value);
                txtRank.Text = "" + string.Format("{0:n0}", dataGridView1.Rows[rowindex].Cells[40].Value);
            }
        }

        // End =================================================================
    }
}

