using System.Windows.Forms;
using System;
using LumiSoft.Net.IMAP.Client;
using LumiSoft.Net.IMAP;
using LumiSoft.Net;

namespace gmail_inbox_to_database
{
    class Program
    {
        //public IMAP_Client m_pImap = null;
        public void LoadFolderMessages()
        {
            try
            {
                using (var m_pImap = new IMAP_Client())
                {
                    m_pImap.Connect("imap.gmail.com", 993, true);
                    m_pImap.Login("gmail@gmail.com", "pass");
                    m_pImap.SelectFolder("INBOX");

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
                    Console.ReadLine();
                }
            }
            catch (Exception x)
            {
                MessageBox.Show(x.ToString());
            }
        }
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
                    try
                    {
                        //ListViewItem currentItem = new ListViewItem();
                        //currentItem.ImageIndex = 0;
                        //currentItem.Tag = fetchResp.UID.UID;

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
                        IMAP_t_MsgFlags status = fetchResp.Flags.Flags;
                        string readstatus = "";
                        string[] statuses = status.ToArray();
                        if (statuses.Length == 0)
                        {
                            string strquery = "INSERT INTO CLN_MST(CLM_MAIL,CLM_SUB,CLM_MLDT) VALUES('" + from + "','" + fetchResp.Envelope.Subject + "','" + fetchResp.InternalDate.Date + "')";
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
                        //Console.WriteLine(AlignCentre(from, 100));
                        Console.WriteLine("From : \t\t" + from);
                        Console.WriteLine("Subject : \t\t" + fetchResp.Envelope.Subject != null ? fetchResp.Envelope.Subject : "<none>");
                        Console.WriteLine("Date : \t\t" + fetchResp.InternalDate.Date.ToString("dd.MM.yyyy HH:mm"));
                        Console.WriteLine("Size : \t\t" + ((decimal)(fetchResp.Rfc822Size.Size / (decimal)1000)).ToString("f2") + " kb");
                        Console.WriteLine();
                        Console.WriteLine();
                        //,
                        //);
                        //currentItem.Text = from;
                        //currentItem.SubItems.Add(fetchResp.Envelope.Subject != null ? fetchResp.Envelope.Subject : "<none>");
                        //currentItem.SubItems.Add(fetchResp.InternalDate.Date.ToString("dd.MM.yyyy HH:mm"));
                        //currentItem.SubItems.Add(((decimal)(fetchResp.Rfc822Size.Size / (decimal)1000)).ToString("f2") + " kb");

                    }
                    catch (Exception x)
                    {
                        MessageBox.Show("Error: " + x.ToString(), "Error:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
            }
            catch (Exception x)
            {
                MessageBox.Show("Error: " + x.ToString(), "Error:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
   
        static void Main(string[] args)
        {
            Program p = new Program();
            p.LoadFolderMessages();
        }
    }
}
