using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Windows.Forms;

using LumiSoft.Net;
using LumiSoft.Net.Log;
using LumiSoft.Net.IMAP.Client;

namespace imap_client_app
{
    /// <summary>
    /// Connect to IMAP window.
    /// </summary>
    public class wfrm_Connect : Form
    {
        private Label         mt_Security = null;
        private ComboBox      m_pSecurity = null;
        private Label         mt_Server   = null;
        private TextBox       m_pServer   = null;
        private NumericUpDown m_pPort     = null;
        private Label         mt_UserName = null;
        private TextBox       m_pUserName = null;
        private Label         mt_Password = null;
        private TextBox       m_pPassword = null;
        private Button        m_pCancel   = null;
        private Button        m_pConnect  = null;

        private EventHandler<WriteLogEventArgs> m_pLogCallback = null;
        private IMAP_Client                     m_pIMAP        = null;

        /// <summary>
        /// Default constructor.
        /// </summary>
        /// <param name="logCallback">Log callback method.</param>
        public wfrm_Connect(EventHandler<WriteLogEventArgs> logCallback)
        {
            m_pLogCallback = logCallback;

            InitUI();
        }

        #region method InitUI

        /// <summary>
        /// Creates and initializes UI.
        /// </summary>
        private void InitUI()
        {
            this.ClientSize = new Size(300,140);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.ShowInTaskbar = false;
            this.Text = "Log on to IMAP server";

            mt_Security = new Label();
            mt_Security.Size = new Size(80,20);
            mt_Security.Location = new Point(0,10);
            mt_Security.TextAlign = ContentAlignment.MiddleRight;
            mt_Security.Text = "Security:";

            m_pSecurity = new ComboBox();
            m_pSecurity.Size = new Size(150,20);
            m_pSecurity.Location = new Point(85,10);
            m_pSecurity.DropDownStyle = ComboBoxStyle.DropDownList;
            m_pSecurity.Items.Add("No TLS/SSL");
            m_pSecurity.Items.Add("Use TLS");
            m_pSecurity.Items.Add("Use SSL");
            m_pSecurity.SelectedIndex = 0;
            m_pSecurity.SelectedIndexChanged += new EventHandler(m_pSecurity_SelectedIndexChanged);            

            mt_Server = new Label();
            mt_Server.Size = new Size(80,20);
            mt_Server.Location = new Point(0,35);
            mt_Server.TextAlign = ContentAlignment.MiddleRight;
            mt_Server.Text = "Server:";

            m_pServer = new TextBox();
            m_pServer.Size = new Size(150,20);
            m_pServer.Location = new Point(85,35);
            m_pServer.Text = "imap.gmail.com";

            m_pPort = new NumericUpDown();
            m_pPort.Size = new Size(50,20);
            m_pPort.Location = new Point(240,35);
            m_pPort.Minimum = 1;
            m_pPort.Maximum = 99999;
            m_pPort.Value = WellKnownPorts.IMAP4;

            mt_UserName = new Label();
            mt_UserName.Size = new Size(80,20);
            mt_UserName.Location = new Point(0,60);
            mt_UserName.TextAlign = ContentAlignment.MiddleRight;
            mt_UserName.Text = "User:";

            m_pUserName = new TextBox();
            m_pUserName.Size = new Size(205,20);
            m_pUserName.Location = new Point(85,60);
            m_pUserName.Text = "_mail";

            mt_Password = new Label();
            mt_Password.Size = new Size(80,20);
            mt_Password.Location = new Point(0,85);
            mt_Password.TextAlign = ContentAlignment.MiddleRight;
            mt_Password.Text = "Password:";

            m_pPassword = new TextBox();
            m_pPassword.Size = new Size(205,20);
            m_pPassword.Location = new Point(85,85);
            m_pPassword.PasswordChar = '*';
            m_pPassword.Text = "_pass";

            m_pCancel = new Button();
            m_pCancel.Size = new Size(70,20);
            m_pCancel.Location = new Point(145,110);
            m_pCancel.Text = "Cancel";
            m_pCancel.Click += new EventHandler(m_pCancel_Click);

            m_pConnect = new Button();
            m_pConnect.Size = new Size(70,20);
            m_pConnect.Location = new Point(220,110);
            m_pConnect.Text = "Connect";
            m_pConnect.Click += new EventHandler(m_pConnect_Click);
            
            this.Controls.Add(mt_Security);
            this.Controls.Add(mt_Server);
            this.Controls.Add(m_pServer);
            this.Controls.Add(m_pPort);
            this.Controls.Add(mt_UserName);
            this.Controls.Add(m_pUserName);
            this.Controls.Add(mt_Password);
            this.Controls.Add(m_pPassword);
            this.Controls.Add(m_pSecurity);
            this.Controls.Add(m_pCancel);
            this.Controls.Add(m_pConnect);
        }
                                                               
        #endregion


        #region Events Handling

        #region method m_pSecurity_SelectedIndexChanged

        private void m_pSecurity_SelectedIndexChanged(object sender,EventArgs e)
        {
            if(m_pSecurity.SelectedIndex == 2){
                m_pPort.Value = WellKnownPorts.IMAP4_SSL;
            }
            else{
                m_pPort.Value = WellKnownPorts.IMAP4;
            }
        }

        #endregion

        #region method m_pCancel_Click

        private void m_pCancel_Click(object sender,EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        #endregion

        #region method m_pConnect_Click

        private void m_pConnect_Click(object sender,EventArgs e)
        {
            if(m_pUserName.Text == ""){
                MessageBox.Show(this,"Please fill user name !","Error:",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }

            this.Cursor = Cursors.WaitCursor;

            IMAP_Client imap = new IMAP_Client();
            try{               
                imap.Logger = new Logger();
                imap.Logger.WriteLog += m_pLogCallback;
                imap.Connect(m_pServer.Text,(int)m_pPort.Value,(m_pSecurity.SelectedIndex == 2));
                if(m_pSecurity.SelectedIndex == 1){
                    imap.StartTls();
                }
                imap.Login(m_pUserName.Text,m_pPassword.Text);

                m_pIMAP = imap;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch(Exception x){
                MessageBox.Show(this,"IMAP server returned: " + x.Message + " !","Error:",MessageBoxButtons.OK,MessageBoxIcon.Error);
                imap.Dispose();
            }

            this.Cursor = Cursors.Default;
        }

        #endregion

        #endregion


        #region Properties Implementation

        /// <summary>
        /// Gets connected IMAP client. Value null means no connected IMAP client.
        /// </summary>
        public IMAP_Client IMAP
        {
            get{ return m_pIMAP; }
        }

        #endregion

    }
}
