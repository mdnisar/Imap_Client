using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace imap_client_app
{ 
    class clsdata
    {
        private string sqlstring;
        private SqlCommand com;
        private string dmlString;
        private SqlDataReader sqlDr;
        private SqlConnection sqlCnn = new SqlConnection("Data Source=.;database=del;integrated security=true;");
        public DataTable getDataTable(string qry)
        {
            if (sqlCnn.State != ConnectionState.Open)
                sqlCnn.Open();
            SqlCommand sqlCmd = new SqlCommand(qry, sqlCnn);
            DataTable dt = new DataTable();
            sqlCmd.CommandTimeout = 150000;
            sqlDr = sqlCmd.ExecuteReader();
            dt.Load(sqlDr);
            sqlDr.Close();
            if (sqlCnn.State == ConnectionState.Open)
                sqlCnn.Close();
            return dt;
        }
        public void executeQuery(string strQry)
        {
            if (sqlCnn.State != ConnectionState.Open)
                sqlCnn.Open();
            SqlCommand sqlCmd = new SqlCommand(strQry, sqlCnn);
            sqlCmd.CommandTimeout = 15000;
            sqlCmd.ExecuteNonQuery();
            sqlCnn.Close();
        }
    }
}
