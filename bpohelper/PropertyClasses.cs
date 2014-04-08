using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{
    public class RealProperty
    {
        private string pin;
        private Dictionary<string, MLSListing> mlsListings;
        private RealistReport rr;
        private Form1 form;
        private string mlsTaxAmout;
        private MatchCollection rawDataFromPrintedMlsSheet;
        private string county;
        private string dateOfLastSale;
        private string lastSalePrice;
        private string currentPropertyTax;
        private RealEstateTransaction lastSale;
        private string mostCurrentMLSListingNumber;
        private string activeMLSListingNumber;

        // Add(System.Text.RegularExpressions.Match)

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

        public Form1 MainForm
        {
            //get { return form; }
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
            }
            catch
            {

            }
           
        }

        public MLSListing GetMlsListing(string id)
        {
            return mlsListings[id];
        }


    }

    public class Neighborhood
    {
        public string name;
        public int oldestHome;
        public int newestHome;
        public int medianAge;
        public int numberOfCompListings;
        public int numberOfSales;
        public decimal minListPrice;
        public decimal highListPrice;
        public int numberOfShortSaleListings;
        public decimal minSalePrice;
        public decimal maxSalePrice;
        public decimal medianSoldPrice;
        public int avgDom;
        public int numberActiveListings;
        public int numberREOListings;
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
