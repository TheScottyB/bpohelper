using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Net;
using System.IO;
//using //BitMiracle.Docotic.Pdf;
using System.Diagnostics;
using System.Collections;
using System.Diagnostics;
using System.Threading;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Linq;
using System.Security.Cryptography;


namespace bpohelper
{

    public static class LandSafeSyntax
    {
        public enum PropertyType
        {
            SFR, Coop, PUD, Condo
        }

        public enum PropertyCondition
        {
            Good,Average,Fair,Poor
        }

        public enum PropertyOccupancy
        {
            Owner,Tenant,Vacant
        }

    }

    public  class LandSafeBpo
    {
        static String filename = @"C:\Users\Scott\Documents\My Dropbox\BPOs\35089 N ROSEWOOD AVE\blank.xml";
        static XmlReader reader;

        string salesComp1 = @"<COMPS COMPNUM=" + "\"1\"" + @">
      <DATA>MLS</DATA>
      <CONDITION>
        <DESCRIPTION>Average</DESCRIPTION>
      </CONDITION>
      <MARKETCHANGE>
        <DESCRIPTION>No</DESCRIPTION>
      </MARKETCHANGE>
      <MARKETDAYS>19</MARKETDAYS>
      <UNLABELED>
        <DESCRIPTION>N/A</DESCRIPTION>
      </UNLABELED>
      <YRBUILT>
        <DESCRIPTION>1956</DESCRIPTION>
      </YRBUILT>
      <AMENITIES>
        <DESCRIPTION>Porch, Patio</DESCRIPTION>
      </AMENITIES>
      <SALETYPE>
        <DESCRIPTION>No</DESCRIPTION>
      </SALETYPE>
      <GBASQFT>
        <SQFT>864</SQFT>
      </GBASQFT>
      <PROPTYPE>
        <DESCRIPTION>1 Story</DESCRIPTION>
      </PROPTYPE>
      <SALE>
        <DATE>10/02/2012</DATE>
      </SALE>
      <SALESPRICE>75005</SALESPRICE>
      <PROXIMITY>0.6</PROXIMITY>
      <PROJSIZETYPE>
        <DESCRIPTION>N/A</DESCRIPTION>
      </PROJSIZETYPE>
      <BASEMENT>
        <DESCRIPTION>Full Finished</DESCRIPTION>
      </BASEMENT>
      <DESIGNSTYLE>
        <DESCRIPTION>Aluminum Siding</DESCRIPTION>
      </DESIGNSTYLE>
      <GARAGE>
        <DESCRIPTION>1 Detached</DESCRIPTION>
      </GARAGE>
      <LEASEFEE>
        <DESCRIPTION>Fee simple</DESCRIPTION>
      </LEASEFEE>
      <ADDR>
        <STREET>13 N Lincoln Ave</STREET>
        <CITY>Round Lake</CITY>
        <STATEPROV>IL</STATEPROV>
        <ZIP>60073</ZIP>
      </ADDR>
      <LOTSIZE>0.3</LOTSIZE>
      <ABOVEGRADE>
        <BEDROOMS>5</BEDROOMS>
        <BATH>1</BATH>
      </ABOVEGRADE>
    </COMPS>";




        private Dictionary<string, string> orderInfoAndSubjPropAnlysFields = new Dictionary<string, string>()
        {          
            {"DATASRC", "Tax records"},
            {"CURRENTOCCUPANT", ""},
            {"LOANNUM", ""},
            {"SERVICERLOANNUM", ""}, 
            {"PROPTYPE", ""}
        };


        private Dictionary<string, string> subjectPropertyFields = new Dictionary<string, string>()
        {
            
        
            {"GBASQFT", ""},{"SALETYPE", ""},{"", ""},{"PROXIMITY", ""},{"LEASEFEE", "Fee Simple"},{"DESIGNSTYLE", ""},{"PROJSIZETYPE", "NA"},{"STATEPROV", ""},{"STREET", ""},{"ZIP", ""},
            {"CITY", ""},
            {"GARAGE", ""},
            {"BASEMENT", ""},           
            {"PROPTYPE", ""},
            {"YRBUILT", ""},
            {"LIVINGSQFT", ""},
            {"CONDITION", ""}
                    
        };

        private Dictionary<string, string> compCommonFields = new Dictionary<string, string>()
        {
            {"BEDROOMS", ""},{"BATH", ""},{"LOTSIZE", ""},{"MARKETDAYS", ""},{"GBASQFT", ""},{"SALETYPE", ""},{"MARKETCHANGE", ""},{"PROXIMITY", ""},{"LEASEFEE", "Fee Simple"},{"DESIGNSTYLE", ""},{"PROJSIZETYPE", "NA"},{"STATEPROV", ""},{"STREET", ""},{"ZIP", ""},
            {"CITY", ""},
            {"DATASRC", "MLS"},
            {"GARAGE", ""},
            {"BASEMENT", ""},           
            {"PROPTYPE", ""},
            {"YRBUILT", ""},
            {"LIVINGSQFT", ""},
            {"CONDITION", ""}
                    
        };

      
        public void WriteEnvFile()
        {
            ValidationEventHandler eventHandler = new ValidationEventHandler(LandSafeBpo.ValidationCallback);
            
         
            try
            {
                // Create the validating reader and specify DTD validation.
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.DtdProcessing = DtdProcessing.Prohibit;
                settings.ValidationType = ValidationType.None;
                settings.ValidationEventHandler += eventHandler;

                reader = XmlReader.Create(filename, settings);
     
                // Pass the validating reader to the XML document. 
                // Validation fails due to an undefined attribute, but the  
                // data is still loaded into the document.
                XmlDocument doc = new XmlDocument();
                XDocument anotherDoc = XDocument.Load(XmlReader.Create(filename, settings));
                doc.Load(reader);
                

                //anotherDoc.Element("FORMINFO").Element("SUBJECT").Element("PROPTYPE").SetValue("TTT");

                //salescomp section
                //returns "1"
                var compNum = anotherDoc.Element("FORMINFO").Element("SALESCOMP").Element("COMPS").Attribute("COMPNUM");
                var ttt = anotherDoc.Element("FORMINFO").Element("SALESCOMP").Elements("COMPS");

                foreach (var vv in ttt)
                {
                    vv.Element("DATA").SetValue("MLS");
                    vv.Element("CONDITION").Element("DESCRIPTION").SetValue("AVERAGE");
                    vv.Element("MARKETCHANGE").Element("DESCRIPTION").SetValue("NO");
                    //vv.Element("UNLABELED").Element("DESCRIPTION").SetValue("NA");
                    vv.Element("LEASEFEE").Element("DESCRIPTION").SetValue("Fee Simple");

                }
                
                //anotherDoc.WriteTo(XmlWriter.Create(filename + "_updated"));
                anotherDoc.Save(filename + "updated");

                SHA1CryptoServiceProvider sha1 = new SHA1CryptoServiceProvider();

                UnicodeEncoding uniEncoding = new UnicodeEncoding();

                byte [] xxx;
                //String filename = @"C:\Users\Scott\Documents\My Dropbox\BPOs\35089 N ROSEWOOD AVE\sha-hash-test.xml";
                using (FileStream fs = new FileStream(@"C:\Users\Scott\Documents\My Dropbox\BPOs\35089 N ROSEWOOD AVE\7001BOADB_68951007230799.env", FileMode.Open))
                //using (BufferedStream bs = new BufferedStream(fs))

                using (var cryptoProvider = new SHA1CryptoServiceProvider())
                {
                    string hash = BitConverter
                            .ToString(cryptoProvider.ComputeHash(fs));

                    MessageBox.Show(hash);

                    //do something with hash
                }

                string ReceivedValue = BitConverter.ToString(sha1.ComputeHash(uniEncoding.GetBytes(anotherDoc.ToString())));

               
                string s = compNum.BaseUri;
                //var data = from item in anotherDoc.Descendants("LISTINGCOMP")
                //           select new
                //           {
                //               listComp1 = item.Element("LISTCOMPS").Element("PROPTYPE").Value,
                //               listComp2 = item.Element("LISTCOMPS").Element("PROPTYPE").Value
                //              // listComp3 = item.
                //              // moneySpent = item.Element("PROPTYPE").Value,
                //               //zipCode = item.Element("personalInfo").Element("zip").Value
                //           };
                //foreach (var p in data)
                //{
                //    Console.WriteLine(p.ToString());

                //}
                Console.WriteLine(doc.OuterXml);

                string ggg = ReceivedValue;


                    //XmlWriterSettings ws = new XmlWriterSettings();
                    //ws.Indent = true;
                    //using (XmlWriter writer = XmlWriter.Create(filename + "_updated"))
                    //{

                       

                    //}
                


            }
            finally
            {
                if (reader != null)
                    reader.Close();
            }
        }

        private static void ValidationCallback(object sender, ValidationEventArgs args)
        {
            Console.WriteLine("Validation error loading: {0}", filename);
            Console.WriteLine(args.Message);
        }

        public string LoanNumber
        {
            get { return orderInfoAndSubjPropAnlysFields["LOANNUM"]; }
            set { orderInfoAndSubjPropAnlysFields["LOANNUM"] = value; }
        }

        public string SubjectPropertyCurrentOcc
        {
            get { return orderInfoAndSubjPropAnlysFields["CURRENTOCCUPANT"]; }
            set { orderInfoAndSubjPropAnlysFields["CURRENTOCCUPANT"] = value; }
        }

        public string SubjectPropertyBasement
        {
            get { return subjectPropertyFields["BASEMENT"]; }
            set { subjectPropertyFields["BASEMENT"] = value; }
        }

        public string SubjectPropertyDataSource
        {
            get { return subjectPropertyFields["DATASRC"]; }
            set { subjectPropertyFields["DATASRC"] = value; }
        }
        
        public string SubjectPropertyGla
        {
            get { return subjectPropertyFields["LIVINGSQFT"]; }
            set { subjectPropertyFields["LIVINGSQFT"] = value; }
        }

        public string SubjectPropertyType
        {
            get { return orderInfoAndSubjPropAnlysFields["PROPTYPE"]; }
            set { orderInfoAndSubjPropAnlysFields["PROPTYPE"] = value; }
        }

        public string SubjectPropertyYearBuilt
        {
            get { return orderInfoAndSubjPropAnlysFields["YRBUILT"]; }
            set { orderInfoAndSubjPropAnlysFields["YRBUILT"] = value; }
        }

        public string SubjectPropertyStyle
        {
            get { return subjectPropertyFields["PROPTYPE"]; }
            set { subjectPropertyFields["PROPTYPE"] = value; }
        }

        public string CompPropertyStyle
        {
            get { return compCommonFields["PROPTYPE"]; }
            set { compCommonFields["PROPTYPE"] = value; }
        }
    }
}
