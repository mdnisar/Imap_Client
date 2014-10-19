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
            sqlcon = new SqlConnection(@"Data Source=.;database=NSupport;uid=sa;pwd=n@1;Connect Timeout=120;");
        }
    }
}
