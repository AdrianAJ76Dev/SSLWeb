using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using System.Xml.Serialization;
using System.IO;

namespace SSLWeb.Models
{
    [Serializable,
        XmlRoot(Namespace = "http//www.collegeboard/sdp/contractsmanagement/SSL")]
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
        private string namewordfile = string.Empty;

        public SSL()
        {
            pathxmlfile = @"D:\Dev Projects\SSL\Documents\";
            namexmlfile = "SSL.xml";
        }

       public void CreateLetter()
        {
            /* Serialize to XML
           * Pass XML to Word Document
           */
            pathxmlfile = @"D:\Dev Projects\SSL\Documents\";
            namexmlfile = "SSL.xml";
            namewordfile = "Sole Source Letter v2.dotx";

            this.SerializeSSLAsXML();
            CBDocument SSLDoc = new CBDocument();
            SSLDoc.DocFullName = pathxmlfile + namewordfile;
            SSLDoc.XmlFileFullName = pathxmlfile + namexmlfile;
            /*
             * Bind XML to Content Controls
             * Save Word Document
             *  Display Word Document
             */
           SSLDoc.AddContact();
        }

        private void SerializeSSLAsXML()
        {
            XmlSerializer SSLXml = new XmlSerializer(typeof(SSL));
            FileStream fs = new FileStream(pathxmlfile + namexmlfile, FileMode.Create);
            SSLXml.Serialize(fs, this);
            fs.Close();
        }
    }


}