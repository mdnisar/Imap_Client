using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace imap_client_app
{
    public partial class frmShowExcelData : Form
    {
        clsdata cm = new clsdata();
        public frmShowExcelData()
        {
            InitializeComponent();
        }
       

        private void frmShowExcelData_Load(object sender, EventArgs e)
        {


           // showdata("CAL68X73iCtXLONi3uyeLrdaTFqiEVfG7zppyybh4fdGeGtDO+g@mail.gmail.com");
        }
        public void showdata(string folderid)
        {
            wfrm_ExportAttachment_SaveExcelToDB frm = new wfrm_ExportAttachment_SaveExcelToDB();
            DataTable dt = new DataTable();
            dt = cm.getDataTable("select ad_data1,ad_data2,ad_data3 from attachment_dtl where mm_message_id = '<" + folderid + ">'");
            
            dg.DataSource = dt;
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            ((DataTable)dg.DataSource).DefaultView.RowFilter = "ad_data2 like '%" + txtSearch.Text.Trim() + "%' or ad_data3 like '%" + txtSearch.Text.Trim() + "%' or ad_data1 like '%" + txtSearch.Text.Trim() + "%'";
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
