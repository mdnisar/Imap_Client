using System;
using System.Windows.Forms;

namespace imap_client_app
{
	/// <summary>
	/// Application main class.
	/// </summary>
	public class MainX
    {
        #region static method Main

        /// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		public static void Main() 
		{
            try{
                Application.EnableVisualStyles();
                Application.Run(new wfrm_ExportAttachment_SaveExcelToDB());
                //Application.Run(new frmShowExcelData());
            }
            catch{
            }
        }

        #endregion
    }
}
