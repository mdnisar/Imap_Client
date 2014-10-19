using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LumiSoft.Net;
using LumiSoft.Net.Log;
using LumiSoft.Net.IMAP.Client;

namespace imap_client_app
{
    public partial class nisar1 : Form
    {
        private EventHandler<WriteLogEventArgs> m_pLogCallback = null;
        private IMAP_Client m_pIMAP = null;
        public nisar1()
        {

            InitializeComponent();
        }

        private void connect()
        {
            this.Cursor = Cursors.WaitCursor;
            IMAP_Client imap = new IMAP_Client();
            try
            {
                imap.Logger = new Logger();
                imap.Logger.WriteLog += m_pLogCallback;
                imap.Connect("mail.ntier.in", 143, false);
                //if (m_pSecurity.SelectedIndex == 1)
                //{
                //    imap.StartTls();
                //}
                imap.Login("nisar@ntier.in", "nisar1");

                m_pIMAP = imap;
                MessageBox.Show("Connected Sucess");
               // this.DialogResult = DialogResult.OK;
                //this.Close();
            }
            catch (Exception x)
            {
                MessageBox.Show(this, "IMAP server returned: " + x.Message + " !", "Error:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                imap.Dispose();
            }

            this.Cursor = Cursors.Default;
        }

    }
}
