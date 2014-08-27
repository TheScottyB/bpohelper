using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using System.Windows.Forms;

namespace bpohelper
{
    public class RealistReport
    {

        public RealistReport()
        {

        }

         public enum Field
        {
            Township, TownshipRangeSect, OwnerName, OwnerName2, OwnerOccupied, Exterior
        };

        public enum OwnerInfoField
        {
            OwnerName, OwnerName2
        };

        public enum LocationInfoField
        {
            Township, TownshipRangeSect
        };

      

        public class RealistReportField : DataItem
        {
            public RealistReportField(string n)
            {
                Key = n;
            }
        }

        public string Township
        {
            get { return FieldMap[Field.Township].Value; }
            set 
            {
                FieldMap.Remove(Field.Township);
             //   FieldMap.Add(Field.Township, 
            
            }
        }

        private Dictionary<Field, DataItem> FieldMap = new Dictionary<Field, DataItem>()
        {
           {Field.OwnerName, new RealistReportField("Owner Name:")},
           {Field.OwnerOccupied, new RealistReportField("Owner Occupied:")},
           {Field.Township, new RealistReportField("Township:")},
           {Field.Exterior, new RealistReportField("Exterior:")}
           //{"Tax Billing Address:"},
           //{"Tax Billing City & State:"},
           //{"Tax Billing Zip:"},
           //{"Tax Billing Zip+4:"},
           //{"Owner Occupied:"},
           //{"Owner Name 2:"}
        };

        public List<string> locationInfoFieldList = new List<string>()
        {
            {"Township:"},
            {"Township Range Sect:"},
            {"Subdivision:"},
            {"School District Name:"},
            {"School District:"},
            {"Census Tract:"},
            {"Carrier Route:"}
        };

        private Dictionary<OwnerInfoField, string> availableFields = new Dictionary<OwnerInfoField, string>()
        {

        };

        

        private string rawText;
        private readonly IYourForm form;
        public RealistReport(IYourForm form)
        {
            this.form = form;
        }
        public string basementSqFt;
        public string avm;
        public string owner1;
        public string owner2;
        public string school = "Unknown";
        public string pin;
        public string yearBuilt;
        public string lotAcres;
        public string gla_above = "0";
        public string lastSalePrice = "Not Available";
        public string lastSaleDate = "Not Available";
        public string fullAddress;
        public string county;
        public string proximity;
        public string subdivision;

        private string GetFieldValue(string fn)
        {
            //xxx(.*?)xxx|xxx(.*?)\n
            string pattern1 = "xxx(.*?)xxx";
            string pattern2 = "xxx(.*?)\n";
            string returnValue = "NotFound";
            string result = "";

            //result = Regex.Match(rawText, string.Format(@"{0}{1}", fn, pattern1)).Groups[1].Value;
            result =  Regex.Match(rawText, string.Format(@"{0}{1}|{0}{2}", fn, pattern1, pattern2)).Groups[1].Value;

            if (String.IsNullOrWhiteSpace(result))
            {
                returnValue = Regex.Match(rawText, string.Format(@"{0}{1}", fn, pattern2)).Groups[1].Value;
            }
            else
            {
                returnValue = result;
            }

            return returnValue;
                
            //return Regex.Match(rawText, string.Format(@"{0}{1}|{0}{2}", fn, pattern1, pattern2)).Groups[1].Value;
        }

        public void GetSubjectInfo(string s)
        {
            //xxx(.*?)xxx|xxx(.*?)\n
            string pattern1 = "xxx(.*?)xxx:";
            string pattern2 = "xxx(.*?)\n";
            string pattern = "";
            string fieldName = "";
            Match match;
            rawText = s;

            foreach (Field f in FieldMap.Keys)
            {
                FieldMap[f].Value = GetFieldValue(FieldMap[f].Key);

            }
            //{
            //    availableFields.Add(Field.OwnerName, GetFieldValue(f));
            //}


            basementSqFt = GetFieldValue(@"Basement Sq Ft:");

            
           //string ttt = ownerInfoFieldList[FieldName.OwnerName];

            //RealAVM™(1):
            //fieldName = @"RealAVM™(1):";
            //pattern = string.Format("{0}{1}{0}{2}", fieldName, pattern1, fieldName, pattern2);
            //Match match = Regex.Match(s, pattern);
            //avm = match.Groups[1].Value;
            avm = GetFieldValue(@"RealAVM™..1.:");


            //pattern = "Owner Name:x+([^x]+)";
            //match = Regex.Match(s, pattern);
            //owner1 = match.Groups[1].Value;
            owner1 = GetFieldValue(@"Owner Name:");


            //pattern = "Parcel ID:x+([^x]+)";
            //match = Regex.Match(s, pattern);
            //pin = match.Groups[1].Value;
            pin = GetFieldValue(@"Parcel ID:");

            //pattern = @"Subdivision:x+([^x\n]+)";
            //match = Regex.Match(s, pattern);
            //subdivision = match.Groups[1].Value;
            subdivision = GetFieldValue(@"Subdivision:");

            //pattern = "School District:x+([^x\\n]+)";
            //match = Regex.Match(s, pattern);
            //if (match.Success)
            //{
            //    school = match.Groups[1].Value;
            //}
            school = GetFieldValue(@"School District:");


            //listing info in realist
            pattern = "MLS Status:x*Closed";
            match = Regex.Match(s, pattern);
            if (match.Success)
            {
                //MLS Sold Price:xxx$169,900
                pattern = "MLS Sold Price:x*([^x\\n]+)";
                match = Regex.Match(s, pattern);
                lastSalePrice = match.Groups[1].Value;

                //MLS Closed Date:xxx01/25/2006xxx
                pattern = "MLS Closed Date:x*([^x\\n]+)";
                match = Regex.Match(s, pattern);
                lastSaleDate = match.Groups[1].Value;
            }
            else
            {

                pattern = "Sale Price[^\\$\\n]+([^x\\n]+)";
                match = Regex.Match(s, pattern);
                if (match.Success)
                {
                    lastSalePrice = match.Groups[1].Value;
                }

                pattern = "Recording Datex+([^x\\n]+)";
                match = Regex.Match(s, pattern);
                if (match.Success)
                {
                    lastSaleDate = match.Groups[1].Value;
                }
            }
            pattern = "Building Above Grade Sq Ft:x+([^x\\n]+)";
            match = Regex.Match(s, pattern);
            if (match.Success)
            {
                gla_above = match.Groups[1].Value;
            }
            else
            {
                pattern = "Building Sq Ft:x+([^x\\n]+)";
                match = Regex.Match(s, pattern);
                if (match.Success)
                {
                    gla_above = match.Groups[1].Value;
                }
            }

            pattern = "Lot Acres:x+([^x\\n]+)";
            match = Regex.Match(s, pattern);
            lotAcres = match.Groups[1].Value;


            //Year Built:xxxTax: 1999 MLS: 1998xxx
            pattern = "Year Built:x+Tax:\\s(\\d\\d\\d\\d)";
            match = Regex.Match(s, pattern);
            yearBuilt = match.Groups[1].Value;
            if (!match.Success)
            {
                pattern = "Year Built:x+([^x\\n]+)";
                match = Regex.Match(s, pattern);
                yearBuilt = match.Groups[1].Value;
            }



            //gets full line, including county
            pattern = "([^\\n]+ County)";
            match = Regex.Match(s, pattern);
            fullAddress = match.Groups[1].Value;

            county = fullAddress.Substring(fullAddress.LastIndexOf(',') + 2);
            county = county.Remove(county.LastIndexOf("County"));
            county = county.Replace(" ", "");

            fullAddress = fullAddress.Remove(fullAddress.LastIndexOf(','));


            


            form.SubjectPin = pin;
            form.SubjectSchoolDistrict = school;
            form.SubjectLastSaleDate = lastSaleDate;
            form.SubjectLastSalePrice = lastSalePrice;
            form.SubjectAboveGLA = gla_above;
            form.SubjectLotSize = lotAcres;
            form.SubjectYearBuilt = yearBuilt;
            form.SubjectOOR = owner1;
            form.SubjectFullAddress = fullAddress;
            form.SubjectCounty = county;
            form.SubjectSubdivision = subdivision;
            form.SubjectAvm = avm;
            form.SubjectBasementGLA = basementSqFt;
            form.SubjectExteriorFinish = FieldMap[Field.Exterior].Value;
        }
    }
}
