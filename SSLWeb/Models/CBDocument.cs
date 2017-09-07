using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

// Open XML SDK
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Wrd13 = DocumentFormat.OpenXml.Office2013.Word;

using System.IO;
using System.Diagnostics;
using System.Text;

namespace SSLWeb.Models
{
    public class CBDocument
    {
        /* 08/30/2017 These are the internal classes. They are the simpliest expression
         * of the architecture of merging that I learned at Micro-Modeling Associates
         */
        private class ClientInfo { }
        private class AutoText { }
        private class ContentControl { }

        public CBDocument() { }

        public string DocFullName { get; set; }
        public string XmlFileFullName { get; set; }


        public void GenerateSSL()
        {
            ClientInfo CI = new ClientInfo();
            AutoText AT = new AutoText();
            ContentControl CC = new ContentControl();
        }

        public void AddContact()
        {
            /*
            * Bind XML to Content Controls
            * Save Word Document
            *  Display Word Document
            */
            using (WordprocessingDocument SSLDoc = WordprocessingDocument.Open(DocFullName, true))
            {
                MainDocumentPart SSLMain = SSLDoc.MainDocumentPart;

                // Retrieve Databinding ID Code. Assuming 1 set of content controls with binding to ONE customxml part
                Wrd13.DataBinding contentdatabinding = SSLMain.Document.Descendants<Wrd13.DataBinding>().FirstOrDefault();
                string databindingvalue = contentdatabinding.StoreItemId;

                // Retrieve CustomXML
                IEnumerable<CustomXmlPart> WordData = from cxml in SSLMain.CustomXmlParts
                                                      where cxml.CustomXmlPropertiesPart.DataStoreItem.ItemId == databindingvalue
                                                      select cxml;

                foreach (var DocData in WordData)
                {
                    using (var reader = new StreamReader(DocData.GetStream(), Encoding.UTF8))
                    {
                        string value = reader.ReadToEnd();
                        // Do something with the value
                        Debug.WriteLine("CustomXML Found: " + value.ToString());
                    }
                }

                foreach (var DocData in WordData)
                {
                    using (StreamReader SSLDataFS = new StreamReader(XmlFileFullName))
                    {
                        DocData.FeedData(SSLDataFS.BaseStream);
                    }
                }
                SSLMain.Document.Save();
                SSLDoc.Close();
            }
        }
    }
}