using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FlyLolo.WordReport;

namespace FlyLolo.WordReport.Demo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            init();
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            DataSet dataSource = SqlHelper.ExecuteDataSet("DefaultConn", "DataBaseToWord");
            create(dataSource);
        }

        private void create(DataSet ds)
        {
            WordReportHelper helper = new WordReportHelper();
            int a = cbxType.SelectedIndex;
            string newFileName = "";
            string errorMsg = "";
            helper.CreateReport(@"d:\WordTemplate.docx", ds, out errorMsg, @"d:\", ref newFileName, int.Parse(cbxType.SelectedValue.ToString()));
        }

        private void init()
        {
            DataTable dt = new DataTable();
            DataColumn dc1 = new DataColumn("id");
            DataColumn dc2 = new DataColumn("name");
            dt.Columns.Add(dc1);
            dt.Columns.Add(dc2);

            DataRow dr1 = dt.NewRow();
            dr1["id"] = 16;
            dr1["name"] = "Word2007";

            DataRow dr2 = dt.NewRow();
            dr2["id"] = 0;
            dr2["name"] = "Word97-2003";

            DataRow dr3 = dt.NewRow();
            dr3["id"] = 17;
            dr3["name"] = "PDF";

            dt.Rows.Add(dr1);
            dt.Rows.Add(dr2);
            dt.Rows.Add(dr3);

            cbxType.DataSource = dt;
            cbxType.ValueMember = "id";
            cbxType.DisplayMember = "name";
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
