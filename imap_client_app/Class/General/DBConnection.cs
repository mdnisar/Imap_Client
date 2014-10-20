using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace imap_client_app
{
    public class DBConnection
    {
        protected SqlConnection sqlcon;
        public DBConnection()
        {
            sqlcon = new SqlConnection(@"Data Source=.;database=del;integrated security=true;Connect Timeout=120;");
        }
    }
}
