using System;
using System.Collections.Generic;
using System.Text;

namespace imap_client_app
{
    class MainInfo
    {
        private string _SerialNu;
        private string _name;
        private string _subject;
        private string _from;
        private string _mail;
        private string _date;

        public string clientSI
        {
            get { return _SerialNu; }
            set { _SerialNu = value; }
        }
        public string clientName
        {
            get { return _name; }
            set {_name = value;}
        }
        public string subject
        {
            get { return _subject; }
            set { _subject = value; }
        }
        public string from
        {
            get { return _from; }
            set { _from = value; }
        }
        public string mail
        {
            get { return _mail; }
            set { _mail = value; }
        }
        public string date
        {
            get { return _date; }
            set{_date = value;}
        }
    }

    
}
