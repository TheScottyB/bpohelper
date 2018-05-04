using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Xml.Serialization;

namespace bpohelper
{
    public class RealProperty
    {
        private string pin;
        private Dictionary<string, MLSListing> mlsListings;
        private RealistReport rr;
     
        protected  Form1 form;
        private string mlsTaxAmout = "-1";
        private MatchCollection rawDataFromPrintedMlsSheet;
        private string county;
        private string dateOfLastSale;
        private string lastSalePrice;
        private string currentPropertyTax;
        private RealEstateTransaction lastSale;
        private string mostCurrentMLSListingNumber;
        private string activeMLSListingNumber;
        private MLSListing currentListing;
        private static string[] activeStates= {"ACTV", "NEW", "TEMP", "CTG"};
        private string rawListingHtml;
        public string mlsStatus;
        public string origListingPrice;

        // Add(System.Text.RegularExpressions.Match)

       
        //
        //Constructors
        //
        #region constructors
        public RealProperty()
        {
            mlsListings = new Dictionary<string, MLSListing>();
        }

        public RealProperty(string parcelID)
            : this()
        {
            pin = parcelID;
        }

        public RealProperty(string parcelID, Form1 f)
            : this()
        {
            pin = parcelID;
            form = f;
        }

        #endregion

        //
        //Internal functions
        //
        #region private
        public virtual string GetFieldValue(string fn)
        {
            //                m.Groups[0].Value
            //"xxxMLS #:09094917xxx"
            //m.Groups[1].Value
            //"MLS #:09094917"
            //m.Groups[2].Value
            //""

            string returnValue = "NotFound";
            string result = "";

            foreach (Match m in rawDataFromPrintedMlsSheet)
            {
                if (m.Groups[1].Value.Contains(fn))
                {
                    returnValue = m.Groups[1].Value.Split(':')[1];
                }
            }

            return returnValue;

            //return Regex.Match(rawText, string.Format(@"{0}{1}|{0}{2}", fn, pattern1, pattern2)).Groups[1].Value;
        }

        public string GetFieldFromPintouts (string fn)
        {
             string returnValue = "NotFound";
                string result = "";

                foreach (Match m in rawDataFromPrintedMlsSheet)
                {
                    if (m.Groups[1].Value.Contains(fn))
                    {
                        returnValue = m.Groups[1].Value.Split(':')[1];
                    }
                }

                return returnValue;
        }

        private bool IsActive()
        {

            if (activeStates.Contains(mlsStatus))
            {
                return true;
            }

            return false;
        }
        #endregion



        public MatchCollection PrintedMlsSheetNameValuePairs
        {
           // get { return rawDataFromPrintedMlsSheet; }
            set { rawDataFromPrintedMlsSheet = value; }
        }

       

        public void AddMlsListing(string mlsnum, MLSListing l)
        {
            mlsListings.Add(mlsnum, l);
        }

       
        public RealEstateTransaction LastSale
        {
            get { return lastSale; }
            set { lastSale = value; }
        }


        public string County
        {
            get { return county; }
            set { county = value; }
        }

        public string DateOfLastSale
        {
            get { return dateOfLastSale; }
            set { dateOfLastSale = value; }
        }

        public string Township
        {
            get { return myRealistReport.Township; }
            set { myRealistReport.Township = value; }
        }
        public string LastSalePrice
        {
            get { return lastSalePrice; }
            set { lastSalePrice = value; }
        }  

        public string Subdivision
        {
            get
            {
                return myRealistReport.subdivision;
                
                //foreach (Match m in rawDataFromPrintedMlsSheet)
                //{
                //    if (m.Value.Contains("Subdivision"))
                //    {
                //        return m.Value;
                     
                //    }
                //}
                //return null;
            
            }
          
        }

        public string ParcelID
        {
            get { return pin; }
            set { pin = value; }
        }

        public RealistReport myRealistReport
        {
            get { return rr; }
            set { rr = value; }
        }
        [XmlIgnoreAttribute]
        public Form1 MainForm
        {
            get { return form; }
            set { form = value; }
        }

        public string MlsType
        {
            get { return rawDataFromPrintedMlsSheet.ToString() ; }
        }

        public string PropertyTax
        {
            get { return mlsTaxAmout; }
            set { mlsTaxAmout = value; }
        }

        public void AddMlsListing(MLSListing m)
        {
            try
            {
                mlsListings.Add(m.MlsNumber, m);
                mostCurrentMLSListingNumber = m.MlsNumber;
            }
            catch
            {

            }
           
        }

        public MLSListing GetCurrentMlsListing()
        {
            return mlsListings[mostCurrentMLSListingNumber];
        }

        public MLSListing GetMlsListing(string id)
        {
            return mlsListings[id];
        }

        public bool IsActiveListing
        {
            get
            {
                return IsActive();
              
            }
        }

        public string ListDate
        {
         get
         {
             MLSListing m = new MLSListing();
             try
             {

       
                     return GetFieldValue(@"List Date:");

                

             } catch
             {

             }
           


             return "";
            //GetMlsListing(mostCurrentMLSListingNumber)  
            
           // m.ListDateString = GetFieldValue(@"List Date:");
         }
            

        }





        public bool ListedInLastYear
        {
            get
            {
                if (string.IsNullOrWhiteSpace(mostCurrentMLSListingNumber))
                {
                    return false;
                }
                else
                {
                    return mlsListings[mostCurrentMLSListingNumber].ListedInLast12Months(); 
                }
                
            }
        }

    }

    public class SubjectProperty:RealProperty
    {

        public SubjectProperty()
            : base()
        {

        }

        public SubjectProperty(string parcelID, Form1 f) : base(parcelID,f)
        {
           
        }

        public string rawtextFromPdfActiveListing;


        public override string GetFieldValue(string fn)
        {
            return helper_GetFieldValue(fn);
        }

        private string helper_GetFieldValue(string fn)
        {
            //:x*(.*?)xxx


            string pattern1 = "x*(.*?)xxx";
            string pattern2 = "x*(.*?)\n";
            string returnValue = "NotFound";
            string result = "";

            //result = Regex.Match(rawText, string.Format(@"{0}{1}", fn, pattern1)).Groups[1].Value;
            result = Regex.Match(rawtextFromPdfActiveListing, string.Format(@"{0}{1}|{0}{2}", fn, pattern1, pattern2)).Groups[1].Value;

            if (String.IsNullOrWhiteSpace(result))
            {
                returnValue = Regex.Match(rawtextFromPdfActiveListing, string.Format(@"{0}{1}", fn, pattern2)).Groups[1].Value;
            }
            else
            {
                returnValue = result;
            }

            return returnValue;

            //return Regex.Match(rawText, string.Format(@"{0}{1}|{0}{2}", fn, pattern1, pattern2)).Groups[1].Value;
        }



        //
        //Address Fields: AddressLine1, City, State, Zip
        //
        #region AddressFields
        public string AddressLine1
        {
            get
            {
                return Regex.Match(form.SubjectFullAddress, @"(^[\w\s]*),").Groups[1].Value;
            }
        }

        public string City
        {
            get
            {
                return Regex.Match(form.SubjectFullAddress, @",\s*(.*),").Groups[1].Value;
            }
        }

        public string Zip
        {
            get
            {
                return Regex.Match(form.SubjectFullAddress, @",\s*IL\s*([\w\s]*)").Groups[1].Value;
            }
        }


        #endregion
        //
        //listing info
        //
        public string  BrokerPhone
        {
            get
            {
                return GetFieldFromPintouts(@"Ph #:");
            }
        }
  //
        //Derived Fields: Age, FullBath, HalfBath, GarageStalls, InspecitonDate, BasementFinishedPercentage
        //
        #region Derived Fields

        public bool AttachedGarage
        {
            get
            {
                if (form.subjectParkingTypeTextbox.Text.ToLower().Contains("gar") && form.subjectParkingTypeTextbox.Text.ToLower().Contains("att"))
                    return true;

                return false;
            }
        }

        public bool DetachedGarage
        {
            get
            {
                if (form.subjectParkingTypeTextbox.Text.ToLower().Contains("gar") && form.subjectParkingTypeTextbox.Text.ToLower().Contains("det"))
                    return true;

                return false;
            }
        }

        public int Age
        {
            get
            {
                DateTime myAge = new DateTime((Convert.ToInt32(form.SubjectYearBuilt)), 1, 1);

                TimeSpan ts = DateTime.Now - myAge;

                return ts.Days / 365;
            }

        }
        public string FullBathCount
        {
            get
            {
                return form.SubjectBathroomCount[0].ToString();
            }
        }
        public string HalfBathCount
        {
            get
            {
                return form.SubjectBathroomCount[2].ToString();
            }
        }
        public string GarageStallCount
        {
            get
            {
                return Regex.Match(form.SubjectParkingType, @"\d+").Value;
            }
        }

        public string TypeOfMlsListing
        {
            get
            {
                if (form.SubjectAttached)
                {
                    return "Attached";
                }
                return "Detached";
            }
        }

        public string InspectionDate()
        {
             string[] fileEntries = Directory.GetFiles(MainForm.SubjectFilePath);

                foreach (string fileName in fileEntries)
                {
                    if (fileName.ToLower().Contains("jpg") && fileName.Length > 6)
                    {
                        return Directory.GetLastWriteTime(fileName).ToString("MM/dd/yyyy");
                    }
                }

                return DateTime.Now.ToString("MM/dd/yyyy");

        }

        public decimal BasementFinishedPercentage
        {
            get
            {
                decimal percentage = 0;
                try
                {
                    int x, y;
                    Int32.TryParse(form.subjectBasementGlaTextbox.Text, out x);
                    Int32.TryParse(form.subjectFinishedBasementGlaTextBox.Text, out y);
                    percentage = y / x * 100;
                }
                catch
                {
                    //something not set right
                }
                return percentage;
            }

        }

           
        #endregion

       
        public bool IsOccupied
        {
            get
            {
                if (GlobalVar.mainWindow.comboBoxSubjectOccupancy.Text.ToLower().Contains("occupied"))
                {
                    return true;
                }

                return false;
            }
        }



    }

    public class Neighborhood
    {
        public string name;

        public int oldestHome;
        public int newestHome;
        public int medianAge;

        public decimal minListPrice;
        public decimal maxListPrice;
        public double medianListPrice;

        public decimal minSalePrice;
        public decimal maxSalePrice;
        public double medianSalePrice;

        public int numberActiveListings;
        public int numberSoldListings;
        public int numberOfCompListings;

        public int numberREOListings;
        public int numberREOSales;

        public int numberShortSales;
        public int numberOfShortSaleListings;

        public int avgDom;
        public int avgDomActv;
        public int avgDomSold;

        
      
        
        public double saleToListRatio = 0.97;
        public double monthlyAppreciationRate = 0.003;

        public decimal AbsorbtionRate
        {
            get
            {
                return numberSoldListings / 12;
            }
        }

        public decimal MonthsSupply
        {
            get
            {
                return Math.Round(numberActiveListings / AbsorbtionRate);
            }
        }

        public double ThreeMonthListPrice
        {
            get
            {
                return medianListPrice * Math.Pow((1 + monthlyAppreciationRate), (12 * .25));
            }
        }

        public double ThreeMonthSalePrice
        {
            get
            {
                return medianSalePrice * Math.Pow((1 + monthlyAppreciationRate), (12 * .25));
            }
        }

        public double SixMonthListPrice
        {
            get
            {
                return medianListPrice * Math.Pow((1 + monthlyAppreciationRate), (12 * .5));
            }
        }

        public double SixMonthSalePrice
        {
            get
            {
                return medianSalePrice * Math.Pow((1 + monthlyAppreciationRate), (12 * .5));
            }
        }
         
    
    }

    public class AssessmentInfo
    {
        public string amount;
        public string frequency;
        public string includes;
        public string managementContact;
        public string managementPhone;
    }

    //property helper classes

    public class RealEstateTransaction
    {
        private DateTime saleDate;
        private decimal salePrice;


        public RealEstateTransaction()
        {

        }

        public RealEstateTransaction(DateTime d, decimal p)
        {
            saleDate = d;
            salePrice = p;
        }

        public string Price
        {
            get { return salePrice.ToString(); }
        }

        public string Date
        {
            get { return saleDate.ToString(); }
        }

    }
  
}
