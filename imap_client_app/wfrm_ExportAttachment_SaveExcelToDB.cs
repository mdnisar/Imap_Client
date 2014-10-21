using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using imap_client_app.Resources;
using LumiSoft.Net;
using LumiSoft.Net.Log;
using LumiSoft.Net.MIME;
using LumiSoft.Net.Mail;
using LumiSoft.Net.IMAP;
using LumiSoft.Net.IMAP.Client;
using System.Data.OleDb;
using System.Data;

namespace imap_client_app
{
    /// <summary>
    /// Application main form.
    /// </summary>
    public class wfrm_ExportAttachment_SaveExcelToDB : Form
    {
        clsdata clsdt = new clsdata();
        private TabControl m_pTab = null;
        // TabPage mail
        private SplitContainer m_pTabPageMail_FoldersSplitter = null;
        private ToolStrip m_pTabPageMail_FoldersToolbar = null;
        private TreeView m_pTabPageMail_Folders = null;
        private ToolStrip m_pTabPageMail_MessagesToolbar = null;
        private ListView m_pTabPageMail_Messages = null;
        private ListView m_pTabPageMail_MessageAttachments = null;
        private TextBox m_pTabPageMail_MessageText = null;
        //TabPage Report
        private DataGridView m_pTabPageReport_dgDownlodedData = null;
        private DataGridViewTextBoxColumn dgvtc_mm_from = null ;
        private DataGridViewTextBoxColumn dgvtc_mm_sub = null;
        private DataGridViewTextBoxColumn dgvtc_mm_date = null;
        private DataGridViewTextBoxColumn dgvtc_mm_size = null;
        private DataGridViewTextBoxColumn dgvtc_am_name = null;
        

        // TabPage log
        private RichTextBox m_pTabPageLog_LogText = null;

        private IMAP_Client m_pImap = null;
        public string folder_id = "";
        /// <summary>
        /// Default constructor.
        /// </summary>
        public wfrm_ExportAttachment_SaveExcelToDB()
        {
            InitUI();
        }

        #region method Dispose

        /// <summary>
        /// Cleans up any resources being used.
        /// </summary>
        /// <param name="disposing"></param>
        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);

            // Clean up IMAP client.
            if (m_pImap != null)
            {
                m_pImap.Dispose();
                m_pImap = null;
            }
        }

        #endregion

        #region method InitUI

        /// <summary>
        /// Creates and initializes UI.
        /// </summary>
        private void InitUI()
        {
            this.ClientSize = new Size(900, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "IMAP Client demo";
            this.Shown += new EventHandler(wfrm_Main_Shown);

            m_pTab = new TabControl();
            m_pTab.Dock = DockStyle.Fill;

            this.Controls.Add(m_pTab);

            #region TabPage Mail

            m_pTab.TabPages.Add("Mail");
            m_pTab.TabPages[0].ClientSize = new Size(792, 574);

            m_pTabPageMail_FoldersSplitter = new SplitContainer();
            m_pTabPageMail_FoldersSplitter.Dock = DockStyle.Fill;
            m_pTabPageMail_FoldersSplitter.Orientation = Orientation.Vertical;
            m_pTabPageMail_FoldersSplitter.BorderStyle = BorderStyle.FixedSingle;

            m_pTab.TabPages[0].Controls.Add(m_pTabPageMail_FoldersSplitter);
            m_pTabPageMail_FoldersSplitter.SplitterDistance = 200;

            #region Folders pane

            m_pTabPageMail_FoldersToolbar = new ToolStrip();
            m_pTabPageMail_FoldersToolbar.Size = new Size(55, 25);
            m_pTabPageMail_FoldersToolbar.Location = new Point(145, 0);
            m_pTabPageMail_FoldersToolbar.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            m_pTabPageMail_FoldersToolbar.GripStyle = ToolStripGripStyle.Hidden;
            m_pTabPageMail_FoldersToolbar.BackColor = this.BackColor;
            m_pTabPageMail_FoldersToolbar.Renderer = new ToolBarRendererEx();
            m_pTabPageMail_FoldersToolbar.ItemClicked += new ToolStripItemClickedEventHandler(m_pTabPageMail_FoldersToolbar_ItemClicked);
            // Add button
            ToolStripButton button_Add = new ToolStripButton();
            button_Add.Image = ResManager.GetIcon("add.ico").ToBitmap();
            button_Add.Name = "add";
            button_Add.ToolTipText = "Create folder";
            m_pTabPageMail_FoldersToolbar.Items.Add(button_Add);
            // Edit button
            ToolStripButton button_Edit = new ToolStripButton();
            button_Edit.Enabled = false;
            button_Edit.Image = ResManager.GetIcon("edit.ico").ToBitmap();
            button_Edit.Name = "edit";
            button_Edit.ToolTipText = "Rename folder";
            m_pTabPageMail_FoldersToolbar.Items.Add(button_Edit);
            // Delete button
            ToolStripButton button_Delete = new ToolStripButton();
            button_Delete.Enabled = false;
            button_Delete.Image = ResManager.GetIcon("delete.ico").ToBitmap();
            button_Delete.Name = "delete";
            button_Delete.ToolTipText = "Delete folder";
            m_pTabPageMail_FoldersToolbar.Items.Add(button_Delete);

            ImageList folders_ImageList = new ImageList();
            folders_ImageList.Images.Add(ResManager.GetIcon("folder.ico"));
            m_pTabPageMail_Folders = new TreeView();
            m_pTabPageMail_Folders.Size = new Size(200, 545);
            m_pTabPageMail_Folders.Location = new Point(0, 25);
            m_pTabPageMail_Folders.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            m_pTabPageMail_Folders.BorderStyle = BorderStyle.None;
            m_pTabPageMail_Folders.FullRowSelect = true;
            m_pTabPageMail_Folders.HotTracking = true;
            m_pTabPageMail_Folders.HideSelection = false;
            m_pTabPageMail_Folders.ImageList = folders_ImageList;
            m_pTabPageMail_Folders.AfterSelect += new TreeViewEventHandler(m_pTabPageMail_Folders_AfterSelect);

            m_pTabPageMail_FoldersSplitter.Panel1.Controls.Add(m_pTabPageMail_FoldersToolbar);
            m_pTabPageMail_FoldersSplitter.Panel1.Controls.Add(m_pTabPageMail_Folders);

            #endregion

            #region Messages pane

            m_pTabPageMail_MessagesToolbar = new ToolStrip();
            m_pTabPageMail_MessagesToolbar.Dock = DockStyle.None;
            m_pTabPageMail_MessagesToolbar.Location = new Point(480, 5);
            m_pTabPageMail_MessagesToolbar.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            m_pTabPageMail_MessagesToolbar.GripStyle = ToolStripGripStyle.Hidden;
            m_pTabPageMail_MessagesToolbar.BackColor = this.BackColor;
            m_pTabPageMail_MessagesToolbar.Renderer = new ToolBarRendererEx();
            m_pTabPageMail_MessagesToolbar.ItemClicked += new ToolStripItemClickedEventHandler(m_pTabPageMail_MessagesToolbar_ItemClicked);
            // Refresh button
            ToolStripButton tabMail_ToolbarButton_Refresh = new ToolStripButton();
            tabMail_ToolbarButton_Refresh.Enabled = false;
            tabMail_ToolbarButton_Refresh.Image = ResManager.GetIcon("refresh.ico").ToBitmap();
            tabMail_ToolbarButton_Refresh.Name = "refresh";
            tabMail_ToolbarButton_Refresh.ToolTipText = "Refresh";
            m_pTabPageMail_MessagesToolbar.Items.Add(tabMail_ToolbarButton_Refresh);
            // Save button
            ToolStripButton tabMail_ToolbarButton_Save = new ToolStripButton();
            tabMail_ToolbarButton_Save.Enabled = false;
            tabMail_ToolbarButton_Save.Image = ResManager.GetIcon("save.ico").ToBitmap();
            tabMail_ToolbarButton_Save.Name = "save";
            tabMail_ToolbarButton_Save.ToolTipText = "Save";
            m_pTabPageMail_MessagesToolbar.Items.Add(tabMail_ToolbarButton_Save);
            // Delete button
            ToolStripButton tabMail_ToolbarButton_Delete = new ToolStripButton();
            tabMail_ToolbarButton_Delete.Enabled = false;
            tabMail_ToolbarButton_Delete.Image = ResManager.GetIcon("delete.ico").ToBitmap();
            tabMail_ToolbarButton_Delete.Name = "delete";
            tabMail_ToolbarButton_Delete.ToolTipText = "Delete";
            m_pTabPageMail_MessagesToolbar.Items.Add(tabMail_ToolbarButton_Delete);

            ImageList messages_ImageList = new ImageList();
            messages_ImageList.Images.Add(ResManager.GetIcon("message.ico"));
            m_pTabPageMail_Messages = new ListView();
            m_pTabPageMail_Messages.Size = new Size(575, 300);
            m_pTabPageMail_Messages.Location = new Point(5, 30);
            m_pTabPageMail_Messages.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            m_pTabPageMail_Messages.View = View.Details;
            m_pTabPageMail_Messages.HideSelection = false;
            m_pTabPageMail_Messages.FullRowSelect = true;
            m_pTabPageMail_Messages.SmallImageList = messages_ImageList;
            m_pTabPageMail_Messages.Columns.Add("From", "From", 150);
            m_pTabPageMail_Messages.Columns.Add("Subject", "Subject", 290);
            m_pTabPageMail_Messages.Columns.Add("Received", "Received", 120);
            m_pTabPageMail_Messages.Columns.Add("Size", "Size", 60);
            m_pTabPageMail_Messages.SelectedIndexChanged += new EventHandler(m_pTabPageMail_Messages_SelectedIndexChanged);

            ImageList attachments_ImageList = new ImageList();
            attachments_ImageList.Images.Add(ResManager.GetIcon("save.ico"));
            m_pTabPageMail_MessageAttachments = new ListView();
            m_pTabPageMail_MessageAttachments.Size = new Size(575, 40);
            m_pTabPageMail_MessageAttachments.Location = new Point(5, 335);
            m_pTabPageMail_MessageAttachments.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            m_pTabPageMail_MessageAttachments.View = View.SmallIcon;
            m_pTabPageMail_MessageAttachments.SmallImageList = attachments_ImageList;
            m_pTabPageMail_MessageAttachments.MouseClick += new MouseEventHandler(m_pTabPageMail_MessageAttachments_MouseClick);

            m_pTabPageMail_MessageText = new TextBox();
            m_pTabPageMail_MessageText.Size = new Size(575, 185);
            m_pTabPageMail_MessageText.Location = new Point(5, 380);
            m_pTabPageMail_MessageText.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            m_pTabPageMail_MessageText.ScrollBars = ScrollBars.Both;
            m_pTabPageMail_MessageText.Multiline = true;

            m_pTabPageMail_FoldersSplitter.Panel2.Controls.Add(m_pTabPageMail_MessagesToolbar);
            m_pTabPageMail_FoldersSplitter.Panel2.Controls.Add(m_pTabPageMail_Messages);
            m_pTabPageMail_FoldersSplitter.Panel2.Controls.Add(m_pTabPageMail_MessageAttachments);
            m_pTabPageMail_FoldersSplitter.Panel2.Controls.Add(m_pTabPageMail_MessageText);

            #endregion

            #endregion

            #region TabPage Log

            m_pTab.TabPages.Add("Log");
            m_pTab.TabPages[1].ClientSize = new Size(700, 500);

            m_pTabPageLog_LogText = new RichTextBox();
            m_pTabPageLog_LogText.Dock = DockStyle.Fill;
            m_pTabPageLog_LogText.ReadOnly = true;

            m_pTab.TabPages[1].Controls.Add(m_pTabPageLog_LogText);

            #endregion

            #region TabPage Report

            m_pTab.TabPages.Add("Report");
            m_pTab.TabPages[2].ClientSize = new Size(700, 500);

            m_pTabPageReport_dgDownlodedData = new DataGridView();
            m_pTabPageReport_dgDownlodedData.Dock = DockStyle.Fill;
            m_pTabPageReport_dgDownlodedData.AllowUserToAddRows = false;
            m_pTabPageReport_dgDownlodedData.AllowUserToOrderColumns = true;
            m_pTabPageReport_dgDownlodedData.AutoGenerateColumns = true;
            
            //dgvtc_mm_from.DataPropertyName = "mm_from";
            //dgvtc_mm_from.HeaderText = "From";
            //dgvtc_mm_from.Name = "mm_from";
            //dgvtc_mm_from.ReadOnly = true;

            //dgvtc_mm_sub.DataPropertyName = "mm_sub";
            //dgvtc_mm_sub.HeaderText = "Subject";
            //dgvtc_mm_sub.Name = "mm_sub";
            //dgvtc_mm_sub.ReadOnly = true;

            //dgvtc_mm_date.DataPropertyName = "mm_date";
            //dgvtc_mm_date.HeaderText = "Date";
            //dgvtc_mm_date.Name = "mm_date";
            //dgvtc_mm_date.ReadOnly = true;

            //dgvtc_mm_size.DataPropertyName = "mm_size";
            //dgvtc_mm_size.HeaderText = "Size";
            //dgvtc_mm_size.Name = "mm_size";
            //dgvtc_mm_size.ReadOnly = true;

            //dgvtc_am_name.DataPropertyName = "am_name";
            //dgvtc_am_name.HeaderText = "Attachments";
            //dgvtc_am_name.Name = "am_name";
            //dgvtc_am_name.ReadOnly = true;

            //m_pTabPageReport_dgDownlodedData.Columns.Add(dgvtc_mm_from);
            //m_pTabPageReport_dgDownlodedData.Columns.Add(dgvtc_mm_sub);
            //m_pTabPageReport_dgDownlodedData.Columns.Add(dgvtc_mm_date);
            //m_pTabPageReport_dgDownlodedData.Columns.Add(dgvtc_mm_size);
            //m_pTabPageReport_dgDownlodedData.Columns.Add(dgvtc_am_name);

           

           
            m_pTab.TabPages[2].Controls.Add(m_pTabPageReport_dgDownlodedData);

            #endregion

        }

        #endregion


        #region Events Handling
        void loadOldData()
        {
            m_pTabPageReport_dgDownlodedData.DataSource = clsdt.getDataTable("select mm.mm_from as [From],mm.mm_sub as Subject,mm.mm_date as Date,mm.mm_size as Size,am.am_name as Attachment from mail_mst mm inner join attachment_mst am on mm.mm_message_id = am.mm_message_id");
        }
        #region method wfrm_Main_Shown

        private void wfrm_Main_Shown(object sender, EventArgs e)
        {
            // Show connect UI.
            wfrm_Connect frm = new wfrm_Connect(new EventHandler<WriteLogEventArgs>(m_pImap_WriteLog));
            if (frm.ShowDialog(this) == DialogResult.OK)
            {
                m_pImap = frm.IMAP;
                m_pImap.MessageExpunged += new EventHandler<EventArgs<IMAP_r_u_Expunge>>(m_pImap_MessageExpunged);

                LoadFolders();
                loadOldData();
            }
            else
            {
                Dispose();
            }
        }

        #endregion

        #region method m_pTabPageMail_FoldersToolbar_ItemClicked

        private void m_pTabPageMail_FoldersToolbar_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (m_pTabPageMail_Folders.SelectedNode == null)
            {
                return;
            }

            string intialFolder = ObjectToString(m_pTabPageMail_Folders.SelectedNode.Tag);

            if (e.ClickedItem.Name == "add")
            {
                wfrm_Folder frm = new wfrm_Folder(true, "", false);
                if (frm.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        m_pImap.CreateFolder(string.IsNullOrEmpty(intialFolder) ? frm.Folder : intialFolder + "/" + frm.Folder);
                    }
                    catch (Exception x)
                    {
                        MessageBox.Show(this, "Error:" + x.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else if (e.ClickedItem.Name == "edit")
            {
                wfrm_Folder frm = new wfrm_Folder(false, intialFolder, true);
                if (frm.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        m_pImap.RenameFolder(intialFolder, frm.Folder);
                    }
                    catch (Exception x)
                    {
                        MessageBox.Show(this, "Error:" + x.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else if (e.ClickedItem.Name == "delete")
            {
                if (MessageBox.Show(this, "Do you want to delete folder '" + m_pTabPageMail_Folders.SelectedNode.Text + "' ?", "Confirm Delete:", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        m_pImap.DeleteFolder(m_pTabPageMail_Folders.SelectedNode.Tag.ToString());
                    }
                    catch (Exception x)
                    {
                        MessageBox.Show(this, "Error:" + x.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            LoadFolders();
        }

        #endregion

        #region method m_pTabPageMail_Folders_AfterSelect

        private void m_pTabPageMail_Folders_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node == null || e.Node.Tag.ToString().Length == 0)
            {
                return;
            }

            m_pTabPageMail_FoldersToolbar.Items["edit"].Enabled = true;
            m_pTabPageMail_FoldersToolbar.Items["delete"].Enabled = true;
            string kk = e.Node.Tag.ToString();
            LoadFolderMessages(e.Node.Tag.ToString());
        }

        #endregion


        #region method m_pTabPageMail_MessagesToolbar_ItemClicked

        private void m_pTabPageMail_MessagesToolbar_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (m_pTabPageMail_Messages.SelectedItems.Count == 0)
            {
                return;
            }

            if (e.ClickedItem.Name == "refresh")
            {
                LoadFolderMessages(m_pTabPageMail_Folders.SelectedNode.Tag.ToString());
            }
            else if (e.ClickedItem.Name == "save")
            {
                SaveFileDialog dlg = new SaveFileDialog();
                dlg.Filter = "Email message | *.eml";
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    Mail_Message msg = (Mail_Message)m_pTabPageMail_MessageAttachments.Tag;
                    msg.ToFile(dlg.FileName, null, null);
                }
            }
            else if (e.ClickedItem.Name == "delete")
            {
                if (MessageBox.Show(this, "Do you want to delete selected message ?", "Confirm Delete:", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    DeleteMessage((long)m_pTabPageMail_Messages.SelectedItems[0].Tag);
                }
            }
        }

        #endregion

        #region method m_pTabPageMail_Messages_SelectedIndexChanged

        private void m_pTabPageMail_Messages_SelectedIndexChanged(object sender, EventArgs e)
        {
            m_pTabPageMail_MessagesToolbar.Items["save"].Enabled = false;
            m_pTabPageMail_MessagesToolbar.Items["delete"].Enabled = false;
            m_pTabPageMail_MessageText.Text = "";

            if (m_pTabPageMail_Messages.SelectedItems.Count == 0)
            {
                return;
            }

            LoadMessage((long)m_pTabPageMail_Messages.SelectedItems[0].Tag);
        }

        #endregion

        #region method m_pTabPageMail_MessageAttachments_MouseClick

        private void m_pTabPageMail_MessageAttachments_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && m_pTabPageMail_MessageAttachments.SelectedItems.Count > 0)
            {
                ContextMenuStrip menu = new ContextMenuStrip();
                menu.Items.Add(new ToolStripMenuItem("Save", ResManager.GetIcon("save.ico").ToBitmap()));
                menu.ItemClicked += new ToolStripItemClickedEventHandler(m_pTabPageMail_MessageAttachmentsMenu_ItemClicked);
                menu.Show(Control.MousePosition);
            }
        }

        private void m_pTabPageMail_MessageAttachmentsMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            try
            {
                MIME_Entity entity = (MIME_Entity)m_pTabPageMail_MessageAttachments.SelectedItems[0].Tag;
                SaveFileDialog dlg = new SaveFileDialog();
                dlg.FileName = m_pTabPageMail_MessageAttachments.SelectedItems[0].Text;
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    File.WriteAllBytes(dlg.FileName, ((MIME_b_SinglepartBase)entity.Body).Data);
                }
            }
            catch (Exception x)
            {
                MessageBox.Show(this, "Error: " + x.Message, "Error:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion


        #region method m_pImap_MessageExpunged

        /// <summary>
        /// This method is called when IMAP server has expunged specified message.
        /// </summary>
        /// <param name="sender">Sender.</param>
        /// <param name="e">Event data.</param>
        private void m_pImap_MessageExpunged(object sender, EventArgs<IMAP_r_u_Expunge> e)
        {
            if (m_pTabPageMail_Messages.Items.Count >= e.Value.SeqNo)
            {
                m_pTabPageMail_Messages.Items.RemoveAt(e.Value.SeqNo - 1);
            }
        }

        #endregion

        #region method m_pImap_Fetch_MessageItems_UntaggedResponse

        /// <summary>
        /// This method is called when FETCH (Envelope Flags InternalDate Rfc822Size Uid) untagged response is received.
        /// </summary>
        /// <param name="sender">Sender.</param>
        /// <param name="e">Event data.</param>
        private void m_pImap_Fetch_MessageItems_UntaggedResponse(object sender, EventArgs<IMAP_r_u> e)
        {

            /* NOTE: All IMAP untagged responses may be raised from thread pool thread,
                so all UI operations must use Invoke.
             
               There may be other untagged responses than FETCH, because IMAP server
               may send any untagged response to any command.
            */
            try
            {
                if (e.Value is IMAP_r_u_Fetch)
                {
                    IMAP_r_u_Fetch fetchResp = (IMAP_r_u_Fetch)e.Value;
                    this.BeginInvoke(new MethodInvoker(delegate()
                    {
                        try
                        {
                            ListViewItem currentItem = new ListViewItem();
                            currentItem.ImageIndex = 0;
                            currentItem.Tag = fetchResp.UID.UID;

                            string from = "";
                            if (fetchResp.Envelope.From != null)
                            {
                                for (int i = 0; i < fetchResp.Envelope.From.Length; i++)
                                {
                                    // Don't add ; for last item
                                    if (i == fetchResp.Envelope.From.Length - 1)
                                    {
                                        from += fetchResp.Envelope.From[i].ToString();
                                    }
                                    else
                                    {
                                        from += fetchResp.Envelope.From[i].ToString() + ";";
                                    }
                                }
                            }
                            else
                            {
                                from = "<none>";
                            }
                            string strquery = "INSERT INTO mail_mst(mm_from,mm_sub,mm_date,mm_size,mm_message_id) VALUES('" + from + "','" + fetchResp.Envelope.Subject + "','" + fetchResp.InternalDate.Date.ToString("yyyy-MM-dd") + "','" + (((decimal)(fetchResp.Rfc822Size.Size / (decimal)1000)).ToString("f2") + " kb") + "','" + fetchResp.Envelope.MessageID + "')";
                            clsdt.executeQuery(strquery);

                            IMAP_t_MsgFlags status = fetchResp.Flags.Flags;
                            string readstatus = "";
                            string[] statuses = status.ToArray();
                            if (statuses.Length == 0)
                            {
                                //string strquery = "INSERT INTO mail_mst(mm_from,mm_sub,mm_date,mm_size) VALUES('" + from + "','" + fetchResp.Envelope.Subject + "','" + fetchResp.InternalDate.Date.ToString("yyyy-MM-dd") + "','" + (((decimal)(fetchResp.Rfc822Size.Size / (decimal)1000)).ToString("f2") + " kb") + "')";
                                //clsdt.executeQuery(strquery);
                            }
                            else
                            {
                                for (int i = 0; i < statuses.Length; i++)
                                {
                                    string newstatus = statuses[i].Remove(0, 1);
                                    if (newstatus.ToLower() == "seen")
                                    {

                                    }
                                }

                            }
                            currentItem.Text = from;
                            currentItem.SubItems.Add(fetchResp.Envelope.Subject != null ? fetchResp.Envelope.Subject : "<none>");
                            currentItem.SubItems.Add(fetchResp.InternalDate.Date.ToString("dd.MM.yyyy HH:mm"));
                            currentItem.SubItems.Add(((decimal)(fetchResp.Rfc822Size.Size / (decimal)1000)).ToString("f2") + " kb");
                            m_pTabPageMail_Messages.Items.Add(currentItem);
                        }
                        catch (Exception x)
                        {
                            MessageBox.Show("Error: " + x.ToString(), "Error:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }));
                }
            }
            catch (Exception x)
            {
                MessageBox.Show("Error: " + x.ToString(), "Error:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region method m_pImap_Fetch_Message_UntaggedResponse

        /// <summary>
        /// This method is called when FETCH RFC822 untagged response is received.
        /// </summary>
        /// <param name="sender">Sender.</param>
        /// <param name="e">Event data.</param>

        private void m_pImap_Fetch_Message_UntaggedResponse(object sender, EventArgs<IMAP_r_u> e)
        {
            /* NOTE: All IMAP untagged responses may be raised from thread pool thread,
                so all UI operations must use Invoke.
             
               There may be other untagged responses than FETCH, because IMAP server
               may send any untagged response to any command.
            */

            try
            {
                DataTable dataTable = new DataTable();
                string str = "";
                if (e.Value is IMAP_r_u_Fetch)
                {
                    IMAP_r_u_Fetch fetchResp = (IMAP_r_u_Fetch)e.Value;

                    this.BeginInvoke(new MethodInvoker(delegate()
                    {
                        try
                        {
                            fetchResp.Rfc822.Stream.Position = 0;
                            Mail_Message mime = Mail_Message.ParseFromStream(fetchResp.Rfc822.Stream);
                            fetchResp.Rfc822.Stream.Dispose();

                            m_pTabPageMail_MessagesToolbar.Items["save"].Enabled = true;
                            m_pTabPageMail_MessagesToolbar.Items["delete"].Enabled = true;

                            m_pTabPageMail_MessageAttachments.Tag = mime;
                            foreach (MIME_Entity entity in mime.Attachments)
                            {
                                ListViewItem item = new ListViewItem();
                                if (entity.ContentDisposition != null && entity.ContentDisposition.Param_FileName != null)
                                {
                                    item.Text = entity.ContentDisposition.Param_FileName;
                                }
                                else
                                {
                                    item.Text = "untitled";
                                }
                                item.ImageIndex = 0;
                                item.Tag = entity;
                                m_pTabPageMail_MessageAttachments.Items.Add(item);
                                clsdt.executeQuery("insert into attachment_mst(file_code,am_name,mm_message_id) values('','" + item.Text + "','" + mime.MessageID + "')");
                                folder_id = mime.MessageID.Replace('<', ' ').Replace('>', ' ').Trim();
                            }
                            if (mime.BodyText != null)
                            {
                                m_pTabPageMail_MessageText.Text = mime.BodyText;
                            }
                            foreach (ListViewItem item in m_pTabPageMail_MessageAttachments.Items)
                            {
                                MIME_Entity entity = (MIME_Entity)item.Tag;
                                bool exists = Directory.Exists("attachments/" + folder_id);
                                if (!exists)
                                {
                                    Directory.CreateDirectory("attachments/" + folder_id);
                                }
                                string fileType = Path.GetExtension(item.Text);

                                File.WriteAllBytes("attachments/" + folder_id + "/" + item.Text, ((MIME_b_SinglepartBase)entity.Body).Data);
                                if (fileType == ".xlsx")
                                {
                                    OleDbCommand command = new OleDbCommand();
                                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                                    //OleDbConnection connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.GetDirectoryName(Application.ExecutablePath) + "//attachments//" + folder_id + "//" + item.Text + "; Extended Properties='Excel 8.0;HDR=Yes;IMAX=1'");
                                    OleDbConnection connection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path.GetDirectoryName(Application.ExecutablePath) + "//attachments//" + folder_id + "//" + item.Text + "; Extended Properties=Excel 12.0;");
                                    connection.Open();
                                    DataTable oleDbSchemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, null);
                                    if (oleDbSchemaTable.Rows.Count > 0)
                                    {
                                        str = oleDbSchemaTable.Rows[0]["TABLE_NAME"].ToString();
                                    }
                                    command = new OleDbCommand("SELECT * FROM [" + str + "]", connection);
                                    adapter.SelectCommand = command;
                                    adapter.Fill(dataTable);
                                    connection.Close();
                                    for (int i = 0; i < dataTable.Rows.Count; i++)
                                    {
                                        clsdt.executeQuery("insert into attachment_dtl(mm_message_id,ad_data1,ad_data2,ad_data3) values('<" + folder_id + ">','" + dataTable.Rows[i][0] + "','" + dataTable.Rows[i][1] + "','" + dataTable.Rows[i][2] + "')");
                                    }
                                    frmShowExcelData frmSED = new frmShowExcelData();
                                    frmSED.showdata(folder_id);
                                    frmSED.ShowDialog();
                                }
                                else
                                {
                                    MessageBox.Show("Attachment file is not in Excel format");
                                }
                            }
                            
                        }
                        catch (Exception x)
                        {
                            MessageBox.Show("Error: " + x.ToString(), "Error:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }));
                }
            }
            catch (Exception x)
            {
                MessageBox.Show("Error: " + x.ToString(), "Error:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region method m_pImap_WriteLog

        /// <summary>
        /// This method is called when IMAP client has new log entry.
        /// </summary>
        /// <param name="sender">Sender.</param>
        /// <param name="e">Event data.</param>
        private void m_pImap_WriteLog(object sender, WriteLogEventArgs e)
        {
            try
            {
                this.BeginInvoke(new MethodInvoker(delegate()
                {
                    if (e.LogEntry.EntryType == LogEntryType.Read)
                    {
                        m_pTabPageLog_LogText.AppendText(ObjectToString(e.LogEntry.RemoteEndPoint) + " >> " + e.LogEntry.Text + "\r\n");
                    }
                    else if (e.LogEntry.EntryType == LogEntryType.Write)
                    {
                        m_pTabPageLog_LogText.AppendText(ObjectToString(e.LogEntry.RemoteEndPoint) + " << " + e.LogEntry.Text + "\r\n");
                    }
                    else if (e.LogEntry.EntryType == LogEntryType.Text)
                    {
                        m_pTabPageLog_LogText.AppendText(ObjectToString(e.LogEntry.RemoteEndPoint) + " xx " + e.LogEntry.Text + "\r\n");
                    }
                }));
            }
            catch (Exception x)
            {
                MessageBox.Show(x.ToString());
            }
        }

        #endregion

        #endregion


        #region mehtod LoadFolders

        /// <summary>
        /// Loads IMAP server folders to UI.
        /// </summary>
        private void LoadFolders()
        {
            m_pTabPageMail_Folders.Nodes.Clear();

            TreeNode nodeMain = new TreeNode("IMAP folders");
            nodeMain.Tag = "";
            m_pTabPageMail_Folders.Nodes.Add(nodeMain);

            try
            {
                IMAP_r_u_List[] folders = m_pImap.GetFolders(null);

                char folderSeparator = m_pImap.FolderSeparator;
                foreach (IMAP_r_u_List folder in folders)
                {
                    string[] folderPath = folder.FolderName.Split(folderSeparator);

                    // Conatins sub folders.
                    if (folderPath.Length > 1)
                    {
                        TreeNodeCollection nodes = nodeMain.Nodes;
                        string currentPath = "";

                        foreach (string fold in folderPath)
                        {
                            if (currentPath.Length > 0)
                            {
                                currentPath += "/" + fold;
                            }
                            else
                            {
                                currentPath = fold;
                            }

                            TreeNode node = FindNode(nodes, fold);
                            if (node == null)
                            {
                                node = new TreeNode(fold);
                                node.Tag = currentPath;
                                nodes.Add(node);
                            }

                            nodes = node.Nodes;
                        }
                    }
                    else
                    {
                        TreeNode node = new TreeNode(folder.FolderName);
                        node.ImageIndex = 0;
                        node.Tag = folder.FolderName;
                        nodeMain.Nodes.Add(node);
                    }
                }
            }
            catch (Exception x)
            {
                MessageBox.Show(this, "Error:" + x.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            nodeMain.Expand();
        }

        #endregion

        #region method LoadFolderMessages

        /// <summary>
        /// Gets specified folder messages list from IMAP server and adds them to UI.
        /// </summary>
        /// <param name="folder">IMAP folder which messages to load.</param>
        private void LoadFolderMessages(string folder)
        {
            m_pTabPageMail_MessagesToolbar.Items["refresh"].Enabled = true;
            m_pTabPageMail_Messages.Items.Clear();

            this.Cursor = Cursors.WaitCursor;
            try
            {
                m_pImap.SelectFolder(folder);

                // Start fetching.
                m_pImap.Fetch(
                    false,
                    IMAP_t_SeqSet.Parse("1:10"),
                    new IMAP_t_Fetch_i[]{
                        new IMAP_t_Fetch_i_Envelope(),
                        new IMAP_t_Fetch_i_Flags(),
                        new IMAP_t_Fetch_i_InternalDate(),
                        new IMAP_t_Fetch_i_Rfc822Size(),
                        new IMAP_t_Fetch_i_Uid()
                    },
                    this.m_pImap_Fetch_MessageItems_UntaggedResponse
                );
            }
            catch (Exception x)
            {
                MessageBox.Show(x.ToString());
                MessageBox.Show(this, "Error: " + x.Message, "Error:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            this.Cursor = Cursors.Default;
        }

        #endregion

        #region method LoadMessage

        /// <summary>
        /// Load specified IMAP message to UI.
        /// </summary>
        /// <param name="uid">Message IMAP UID value.</param>
        private void LoadMessage(long uid)
        {
            m_pTabPageMail_MessageAttachments.Items.Clear();

            this.Cursor = Cursors.WaitCursor;
            try
            {
                // Start fetching.
                m_pImap.Fetch(
                    true,
                    IMAP_t_SeqSet.Parse(uid.ToString()),
                    new IMAP_t_Fetch_i[]{
                        new IMAP_t_Fetch_i_Rfc822()
                    },
                    this.m_pImap_Fetch_Message_UntaggedResponse
                );
            }
            catch (Exception x)
            {
                MessageBox.Show(this, "Error: " + x.ToString(), "Error:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            this.Cursor = Cursors.Default;
        }

        #endregion

        #region method DeleteMessage

        /// <summary>
        /// Deletes specified message.
        /// </summary>
        /// <param name="uid">Message UID.</param>
        private void DeleteMessage(long uid)
        {
            this.Cursor = Cursors.WaitCursor;
            try
            {
                /* NOTE: In IMAP message deleting is 2 step operation.
                 *  1) You need to mark message deleted, by setting "Deleted" flag.
                 *  2) You need to call Expunge command to force server to dele messages physically.
                */

                IMAP_t_SeqSet sequence_set = IMAP_t_SeqSet.Parse(uid.ToString());
                m_pImap.StoreMessageFlags(true, sequence_set, IMAP_Flags_SetType.Add, new IMAP_t_MsgFlags(new string[] { IMAP_t_MsgFlags.Deleted }));
                m_pImap.Expunge();
            }
            catch (Exception x)
            {
                MessageBox.Show(this, "Error: " + x.Message, "Error:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            this.Cursor = Cursors.Default;
        }

        #endregion

        #region mehtod FindNode

        private TreeNode FindNode(TreeNodeCollection nodes, string nodeName)
        {
            if (nodes != null)
            {
                foreach (TreeNode node in nodes)
                {
                    if (node.Text == nodeName)
                    {
                        return node;
                    }
                }
            }

            return null;
        }

        #endregion

        #region method ObjectToString

        /// <summary>
        /// Calls obj.ToSting() if obj is not null, otherwise returns "".
        /// </summary>
        /// <param name="obj">Object.</param>
        /// <returns>Returns obj.ToSting() if obj is not null, otherwise returns "".</returns>
        private string ObjectToString(object obj)
        {
            if (obj == null)
            {
                return "";
            }
            else
            {
                return obj.ToString();
            }
        }

        #endregion

    }
}
