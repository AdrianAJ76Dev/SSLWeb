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
    public enum ssltype:byte
    {
        k12=0,
        hed=1,
        warranty1=3,
        warranty2=4,
        warranty3=5
    }
    public enum sslsignature:byte
    {
        DavidCMeadeJr=0,
        JaneDapkus=1,
        TrevorPacker=2,
        AuditiChakravarty=3,
        JeremySinger=4,
    }

    public class SSL
    {
        // Defaults
       // Contact
        private string firstname = string.Empty;
        private string lastname = string.Empty;
        private string title = string.Empty;
        private string institution = string.Empty;
        private string address = string.Empty;
        private string city = string.Empty;
        private string state = string.Empty;
        private string zipcode = string.Empty;

        // AutoText
        private string[] autotextlistssltype = { "SSL-K12", "SSL-HED", "Sole Source - Price Warranty", "Sole Source - Price Warranty 2", "Sole Source - Price Warranty 3" };
        private string[] autotextlistsslsignature = { "David C Meade Jr", "Cyndie Schmeiser", "Trevor Packer", "Auditi Chakravarty", "Jeremy Singer"};
        private ssltype lettertype = ssltype.k12;
        private sslsignature signaturechoice = sslsignature.JaneDapkus;

        // Template
        private const string pathxmlfile = @"D:\Dev Projects\SSL\Documents\";
        private const string namexmlfile = "SSL.xml";
        private const string namewordfile = "Sole Source Letter v3.dotx";
        private const string sslnamespace = @"http//www.collegeboard/sdp/contractsmanagement/SSL/Contact/";
        private string dataidlink = string.Empty;

        // Properties
        // Contact fields for the serialization.
        public string FirstName { get { return firstname; } set { firstname = value; } }
        public string LastName { get { return lastname; } set { lastname = value; } }
        public string Title { get { return title; } set { title = value; } }
        public string Institution { get { return institution; } set { institution = value; } }
        public string Address { get { return address; } set { address = value; } }
        public string City { get { return city; } set { city = value; } }
        public string State { get { return state; } set { state = value; } }
        public string ZipCode { get { return zipcode; } set { zipcode = value; } }

        // Sole Source Letter Type i.e. enumeration ssltype
        public ssltype LetterType { get { return lettertype; } set { lettertype = value; } }
        // Sole Source Letter Signature i.e. enumeration sslsignature
        public sslsignature SignatureChoice { get { return signaturechoice; } set { signaturechoice = value; } }

        public string TemplateFullName { get { return pathxmlfile + namewordfile;} }
        public string CustomXMLFileName { get { return pathxmlfile + namexmlfile; } }

        public string ContentControlName
        {
            get
            {
                if (this.LetterType==ssltype.hed || this.LetterType==ssltype.k12)
                {
                    return "Programs";
                }
                else
                {
                    return "LetterBody";
                }
            }
        }

        public void CreateLetter()
        {
 
            /* Serialize to XML
           * Pass XML to Word Document
           */
            this.SerializeSSLAsXML();

           CBDocument SSLDoc = new CBDocument();
            SSLDoc.DocFullName = this.TemplateFullName;
            SSLDoc.XmlFileFullName = this.CustomXMLFileName;
            SSLDoc.GenerateSSL();
            /*
             * Bind XML to Content Controls
             * Save Word Document
             *  Display Word Document
             */
        }

        private void SerializeSSLAsXML()
        {
            XmlSerializer SSLXml = new XmlSerializer(typeof(SSL));
            FileStream fs = new FileStream(this.CustomXMLFileName, FileMode.Create);
            SSLXml.Serialize(fs, this);
            fs.Close();
        }
    }


}