using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using System.Xml.Serialization;
using System.IO;

namespace SSLWeb.Models
{
    [Serializable,
        XmlRoot(Namespace = "http//www.collegeboard/sdp/contractsmanagement/SSL/Contact/")]
    public class SSL
    {
        public string firstname = string.Empty;
        public string lastname = string.Empty;
        public string title = string.Empty;
        public string institution = string.Empty;
        public string address = string.Empty;
        public string city = string.Empty;
        public string state = string.Empty;
        public string zipcode = string.Empty;

        private string pathxmlfile = string.Empty;
        private string namexmlfile = string.Empty;

        public SSL()
        {
            pathxmlfile = @"D:\Dev Projects\SSL\Documents\";
            namexmlfile = "SSL.xml";
        }

        public string PathXmlFile { get;}
        public string NameXmlFile { get;}

        public void CreateLetter()
        {
            this.SerializeSSLAsXML();
        }

        /* Serialize to XML
         * Pass XML to Word Document
         * Bind XML to Content Controls
         * Save Word Document
         * Display Word Document
         */

        private void SerializeSSLAsXML()
        {
            XmlSerializer SSLXml = new XmlSerializer(typeof(SSL));
            FileStream fs = new FileStream(pathxmlfile + namexmlfile, FileMode.OpenOrCreate);
            SSLXml.Serialize(fs, this);
            fs.Close();
        }
    }


}