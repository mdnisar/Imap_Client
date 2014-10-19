using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
    

namespace imap_client_app
{
    
    class MailSP:DBConnection
    {
        public bool Client(MainInfo clientadd)
        {
            try
            {
                if (sqlcon.State == ConnectionState.Closed)
                {
                    sqlcon.Open();
                }
                SqlCommand sccmd = new SqlCommand("spNSupport_AddClientmaster",sqlcon);
                sccmd.CommandType = CommandType.StoredProcedure;
                SqlParameter sprmparam = new SqlParameter();
                sprmparam = sccmd.Parameters.Add("@clientName", SqlDbType.VarChar);
                sprmparam.Value = clientadd.clientName;
               
            }
            catch
            {

            }
            finally
            {

            }
            return true;
        }
    }
}
