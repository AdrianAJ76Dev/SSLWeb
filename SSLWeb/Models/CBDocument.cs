﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using System.IO;
using System.Diagnostics;
using System.Text;

// Open XML SDK
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Wrd13 = DocumentFormat.OpenXml.Office2013.Word;

namespace SSLWeb.Models
{
    public class CBDocument
    {
        private string documentname;
        private string xmlfile;

        private WordprocessingDocument cbtemplate;
        private WordprocessingDocument newssldoc;
        private MainDocumentPart cbtemplatemain;
        private GlossaryDocument GlossaryDoc;

        private CBAutoText docAtx;

        private int autotextcount=0;

        // Constructors
        public CBDocument() { }
        public CBDocument(string DocFullName, string XmlFileFullName)
        {
            documentname = DocFullName;
            xmlfile = XmlFileFullName;

            cbtemplate = WordprocessingDocument.Open(documentname, true); //Open template
            cbtemplatemain = cbtemplate.MainDocumentPart;

            GlossaryDocument GlossaryDoc =
                cbtemplatemain.GetPartsOfType<GlossaryDocumentPart>().FirstOrDefault().GlossaryDocument;

            if (GlossaryDoc!=null)
            {
                autotextcount=GlossaryDoc.DocParts.Count();
            }
        }

        public void CreateDocumentSimple() { }
        public void CreateDocumentFromTemplate() { }

        public MainDocumentPart TemplateMain { get { return cbtemplatemain; } }
        public GlossaryDocument AutoTextDoc { get { return GlossaryDoc; } }
        public int AutoTextTotal { get{ return autotextcount; } set{ autotextcount = value; } }

        private class CBAutoText
        {
            /* autotextname is self-explanatory
             * autotextcategory is how I assign several pieces of AutoText to 
             * a single content control. "Programs" content control for example
             * it can take 1 or more AutoText Entries i.e. they are all placed
             * in the same position in the document.  Signatories will be like this
             */
            private string autotextname;
            private string autotextcategory;
            private SdtContentBlock ccautotext;
            private SdtContentRun ccautotextrun;


            private string autotext;
            private int autotextcount=0;

            private CBDocument parentdoc;

            public CBAutoText() { }
            public CBAutoText(string AutoTextName)
            {
                autotextname = AutoTextName;
            }

            public CBDocument ParentDoc { get { return parentdoc; } set { parentdoc=value; }  }

            public int Total
            {
                get
                {
                    return autotextcount;
                }
            }

            public string Insert
            {
                get
                {
                    //    var gDocPartBodyPrograms = from x in GlossaryDoc.DocParts
                    //                               where x.Descendants<DocPartProperties>().FirstOrDefault().DocPartName.Val == autotextname
                    //                               select x.Descendants<Paragraph>().FirstOrDefault();

                    //    autotext = gDocPartBodyPrograms.FirstOrDefault().InnerXml;
                    return autotext;
                }
            }
        }

        //public int AutoTextTotal()
        //{
        //    CBAutoText at = new CBAutoText();
        //    return at.Total;
        //}

        /* 08/30/2017 These are the internal classes. They are the simpliest expression
          * of the architecture of merging that I learned at Micro-Modeling Associates
          */
        // DocPart from template

        public void AddContact()
        {
            /*
            * Bind XML to Content Controls
            * Save Word Document
            *  Display Word Document
            */
            using (WordprocessingDocument SSLDoc = WordprocessingDocument.Open(documentname, true))
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
                    using (StreamReader SSLDataFS = new StreamReader(xmlfile))
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