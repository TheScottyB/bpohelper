using System;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using WatiN.Core;
using System.Text.RegularExpressions;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using WatiN;
using WatiN.Core.Native;
using iMacros;
using System.Net;
using System.IO;
using Google.GData.Client;
using Google.GData.Spreadsheets;
//using //BitMiracle.Docotic.Pdf;
using System.Diagnostics;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Collections;
using HtmlAgilityPack;
using XnaFan.ImageComparison;
using Google.Apis.Fusiontables.v1;
using DotNetOpenAuth.OAuth2;
using Google.Apis.Authentication;
using Google.Apis.Authentication.OAuth2;
using Google.Apis.Authentication.OAuth2.DotNetOpenAuth;
using Google.Apis.Samples.Helper;
using System.Threading;
using System.Xml;
using System.Xml.Schema;
using System.Runtime.Serialization;
using System.Web;
using Newtonsoft.Json;


 public interface IYourForm
    {
        string SubjectDom { get; set; }
        string SubjectListingAgent { get; set; }
        string SubjectBrokerPhone { get; set; }
        string SubjectCurrentListPrice { get; set; }
        string SubjectMlsStatus { get; set; } //subjectMlsStatusTextBox
        string SubjectQuickSaleValue { get; set; }
        string SubjectRent { get; set; }
        void AddInfoColor(string value, Color c);
        bool CacheSearch { get; set; }
        Bitmap CompPic { set; }
        string SubjectAvm { get; set; }
        string SubjectPin { get; set; }
        string SubjectSchoolDistrict { get; set; }
        string SubjectLastSaleDate { get; set; }
        string SubjectLastSalePrice { get; set; }
        string SubjectAboveGLA { get; set; }
        string SubjectLotSize { get; set; }
        string SubjectYearBuilt { get; set; }
        string SubjectRoomCount { get; set; }
        string SubjectBedroomCount { get; set; }
        string SubjectBathroomCount { get; set; }
        string SubjectBasementType { get; set; }
        string SubjectBasementDetails { get; set; }
        string SubjectBasementGLA { get; set; }
        string SubjectBasementFinishedGLA { get; set; }
        string SubjectNumberFireplaces { get; set; }
        string SubjectParkingType { get; set; }
        string SubjectOOR { get; set; }
        string SubjectFullAddress { get; set; }
        string SubjectFilePath { get; set; }
        string SubjectCounty { get; set; }
        string SubjectProximityToOffice { get; set; }
        string SubjectStyle { get; set; }
        bpohelper._BPO_SandboxDataSetTableAdapters.RawSFDataTableAdapter RawSFDatatable { get; }
        string SubjectSubdivision { get; set; }
        string SubjectMlsType { get; set; }
        bpohelper.Neighborhood SubjectNeighborhood { get; set; }
        bpohelper.AssessmentInfo SubjectAssessmentInfo { get; set; }
        bool SubjectDetached { get; set; }
        bool SubjectAttached { get; set; }
        Bitmap SubjectPic { set; }
        string SetStatusBar { set; }
        string PicDiffLabel { set; }
        string AddInfo { set; }
        string StatusUpdate { set; }
        string CurrentSearchName { get; set; }
        bool UpdateRealist { get; set; }
        string NumberOfCompsFound { set; }
        string SubjectExteriorFinish { get; set; }
    }

namespace bpohelper
{
    public partial class Form1 : System.Windows.Forms.Form, IYourForm
    {
        iMacros.App iim = new iMacros.App();
        iMacros.App iim2 = new iMacros.App();
        iMacros.Status status;
        public Neighborhood subjectNeighborhood;
        public AssessmentInfo subjectAssessmentInfo;

        Broker dawn = new Broker("Dawn", "471.009163", "7", "60050");
        Broker scott;

        USRESForm usres_webform;


        IE realist;
        IE listingswindow;
        IE mainstreet;
        IE m2m;
        ExtBPO mybpo = new ExtBPO();
        FireFox m2mf;

        MarketStats oneMile = new MarketStats();

        TownshipReport subjectTownshipRecord;


        public Form1()
        {
            InitializeComponent();
            GlobalVar.ccc = GlobalVar.CompCompareSystem.NABPOP;
            bool foundLandsafe = false;



            iMacros.App browser1 = new iMacros.App();
            iMacros.App browser2 = new iMacros.App();
            status = browser1.iimOpen("-ie", false, 60);
            browser1.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string b1Url = browser1.iimGetLastExtract();
            status = browser2.iimOpen("-ie", false, 60);
            browser2.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string b2Url = browser2.iimGetLastExtract();

            if (b1Url.Contains("mredllc"))
            {
                iim = browser1;
                iim2 = browser2;
            }
            else if (b2Url.Contains("mredllc"))
            {
                iim = browser2;
                iim2 = browser1;
            }
            else
            {
                if (b1Url.Contains("about:blank"))
                {
                    iim = browser1;
                    iim2 = browser2;
                }
                else if (b2Url.Contains("about:blank"))
                {
                    iim2 = browser1;
                    iim = browser2;
                }
                else
                {
                    iim = browser1;
                    iim2 = browser2;
                }
                
            }


            if (status == iMacros.Status.sOk)
            {
                browser1.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
                string currentUrl = browser1.iimGetLastExtract();
                b1Url = currentUrl;

                if (currentUrl.ToLower().Contains("dnaforms"))
                {
                    foundLandsafe = true;
                    ieConnectedCheckbox.BackColor = Color.Green;
                    ieConnectedCheckbox.Checked = true;
                    iim2 = browser1;
                    iim = browser2;
                }
                else
                {
                    browser2.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
                    currentUrl = browser2.iimGetLastExtract();
                    b2Url = currentUrl;

                    if (currentUrl.ToLower().Contains("dnaforms"))
                    {
                        foundLandsafe = true;
                        ieConnectedCheckbox.BackColor = Color.Green;
                        ieConnectedCheckbox.Checked = true;
                        iim2 = browser2;
                        iim = browser1;
                    }
                }




            }

            AssessmentInfo subjectAssessmentInfo = new AssessmentInfo();

             Neighborhood subjectNeighborhood = new Neighborhood();
             SubjectNeighborhood = subjectNeighborhood;
            SubjectAssessmentInfo = subjectAssessmentInfo;
            TypeDetachedList tdl = new TypeDetachedList();
            subjectTownshipRecord = new TownshipReport();
            foreach (string key in tdl.mlsTypeDetached.Keys)
            {
                subjectMlsTypecomboBox.Items.Add(tdl.mlsTypeDetached[key]);
            }
        }

        //
        //Helpers
        //
        private void LoadMlsTypeList(TypeDetachedList tdl)
        {
            subjectMlsTypecomboBox.Items.Clear();

            foreach (string key in tdl.mlsTypeDetached.Keys)
            {
                subjectMlsTypecomboBox.Items.Add(tdl.mlsTypeDetached[key]);
            }
        }

        private void LoadMlsTypeList(TypeAttachedList tal)
        {
            subjectMlsTypecomboBox.Items.Clear();
            foreach (string key in tal.mlsTypeAttached.Keys)
            {
                subjectMlsTypecomboBox.Items.Add(tal.mlsTypeAttached[key]);
            }
        }

        private int Age(string yearBuilt)
        {
            DateTime subject_age = new DateTime((Convert.ToInt32(yearBuilt)), 1, 1);

            TimeSpan ts = DateTime.Now - subject_age;

            return ts.Days / 365;
        }

        private void streetnameTextBox_TextChanged(object sender, EventArgs e)
        {
            //this.streetnameTextBox.Text = comboBox1.Text;
        }

        async private void run_script_Click(object sender, EventArgs e)
        {
            MLSListing m;

            // Display the ProgressBar control.
            pBar2.Visible = true;
            // Set Minimum to 1 to represent the first file being copied.
            pBar2.Minimum = 1;
            // Set Maximum to the total number of files to copy.
            pBar2.Maximum = 6;
            // Set the initial value of the ProgressBar.
            pBar2.Value = 1;
            // Set the Step property to a value of 1 to represent each file being copied.
            pBar2.Step = 1;

            // Display the ProgressBar control.
            pBar1.Visible = true;
            // Set Minimum to 1 to represent the first file being copied.
            pBar1.Minimum = 1;
            // Set Maximum to the total number of files to copy.
            pBar1.Maximum = 24;
            // Set the initial value of the ProgressBar.
            pBar1.Value = 1;
            // Set the Step property to a value of 1 to represent each file being copied.
            pBar1.Step = 1;

            StringBuilder macro12 = new StringBuilder();
            macro12.AppendLine(@"FRAME NAME=navpanel");
            macro12.AppendLine(@"TAG POS=1 TYPE=TD ATTR=ID:navtd EXTRACT=TXT");
            macro12.AppendLine(@"FRAME NAME=workspace");
            //// macro12.AppendLine(@"TAG POS=1 TYPE=TABLE FORM=NAME:dc ATTR=CLASS:gridview EXTRACT=TXT");
            // macro12.AppendLine(@"TAG POS=1 TYPE=HTML ATTR=* EXTRACT=HTM ");
            macro12.AppendLine(@"TAG POS=1 TYPE=DIV FORM=NAME:dc ATTR=CLASS:report EXTRACT=HTM");
            iMacros.Status s = iim.iimPlayCode(macro12.ToString());
            string htmlCode = iim.iimGetLastExtract();
            string header = iim.iimGetExtract(0);

            if (!Regex.IsMatch(header, @"showing\s*1\s*of\s*6"))
            {
                var dr = MessageBox.Show("Do you want to start at comp 1?", "Error", MessageBoxButtons.YesNoCancel);

                if (dr == DialogResult.Yes)
                {
                    StringBuilder macro = new StringBuilder();
                    macro.AppendLine(@"SET !ERRORIGNORE YES");
                    macro.AppendLine(@"SET !TIMEOUT_STEP 0");
                    macro.AppendLine(@"FRAME NAME=subheader");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:header ATTR=TXT:List<SP>View");
                    macro.AppendLine(@"FRAME NAME=workspace");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=HREF:javascript:clickedDCID('1');");
                    string macroCode = macro.ToString();
                    var status = iim.iimPlayCode(macroCode, 60);
                } 
                else if (dr == DialogResult.Cancel)
                {
                    return;
                }
                
            }


            if (subjectAttachedRadioButton.Checked)
            {
                m = new AttachedListing(htmlCode);
            }
            else if (subjectDetachedradioButton.Checked)
            {
                m = new DetachedListing(htmlCode);
            }
            else
            {
                m = new MLSListing(htmlCode);
            }

  


            #region check for mred and start read
            iim2.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string currentUrl = iim2.iimGetLastExtract();
            string input_comp_name = streetnameTextBox.Text;

            if (currentUrl.ToLower().Contains("usres"))
            {
                input_comp_name = "usres";
            }

            streetnameTextBox.Text = input_comp_name;
            streetnameTextBox.Update();


            StringBuilder move_through_comps_macro = new StringBuilder();
            move_through_comps_macro.AppendLine(@"SET !TIMEOUT_STEP 0");
            move_through_comps_macro.AppendLine(@"FRAME NAME=navpanel");
            move_through_comps_macro.AppendLine(@"TAG POS=1 TYPE=IMG ATTR=SRC:*/next.gif");
     
            #endregion


            #region setup comp names

            int active_comps = 0;
            int closed_comps = 0;
                
            string [] comp_name_list = {"RECENT_SALE1", "RECENT_SALE2","RECENT_SALE3"};
            string[] active_name_list = { "COMPARABLE1", "COMPARABLE2", "COMPARABLE3" };




            if (currentUrl.ToLower().Contains("solutionstar.gatorsasp.com"))
            {

                comp_name_list[0] = "1";
                comp_name_list[1] = "2";
                comp_name_list[2] = "3";
                active_name_list[0] = "1";
                active_name_list[1] = "2";
                active_name_list[2] = "3";

            }

            if (streetnumTextBox.Text == "emort")
            {

                comp_name_list[0] = "_1";
                comp_name_list[1] = "_2";
                comp_name_list[2] = "_3";
                active_name_list[0] = "_1";
                active_name_list[1] = "_2";
                active_name_list[2] = "_3";

            }
       
            if (currentUrl.ToLower().Contains("bpofulfillment.com"))
            {

                comp_name_list[0] = "2";
                comp_name_list[1] = "3";
                comp_name_list[2] = "4";
                active_name_list[0] = "5";
                active_name_list[1] = "6";
                active_name_list[2] = "7";

            }

            if (streetnumTextBox.Text == "sls")
            {

                comp_name_list[0] = "1";
                comp_name_list[1] = "2";
                comp_name_list[2] = "3";
                active_name_list[0] = "1_2";
                active_name_list[1] = "2_2";
                active_name_list[2] = "3_2";

            }


            if (input_comp_name == "usres")
            {

                comp_name_list[0] = "S1";
                comp_name_list[1] = "S2";
                comp_name_list[2] = "S3";
                active_name_list[0] = "L1";
                active_name_list[1] = "L2";
                active_name_list[2] = "L3";

            }

            if (streetnumTextBox.Text == "vp")
            {

                comp_name_list[0] = "COMP1";
                comp_name_list[1] = "COMP2";
                comp_name_list[2] = "COMP3";
                active_name_list[0] = "LISTCOMP1";
                active_name_list[1] = "LISTCOMP2";
                active_name_list[2] = "LISTCOMP3";

            }

            if (currentUrl.ToLower().Contains("equi-trax"))
            {

                comp_name_list[0] = "Sales1";
                comp_name_list[1] = "Sales2";
                comp_name_list[2] = "Sales3";
                active_name_list[0] = "List1";
                active_name_list[1] = "List2";
                active_name_list[2] = "List3";
            }

            if (currentUrl.ToLower().Contains("marktomarket"))
            {

                comp_name_list[0] = "Sale1";
                comp_name_list[1] = "Sale2";
                comp_name_list[2] = "Sale3";
                active_name_list[0] = "Listing1";
                active_name_list[1] = "Listing2";
                active_name_list[2] = "Listing3";
            }

            if (currentUrl.ToLower().Contains("ort.quandis"))
            {

                comp_name_list[0] = "CompOne";
                comp_name_list[1] = "CompTwo";
                comp_name_list[2] = "CompThree";
                active_name_list[0] = "ListingOne";
                active_name_list[1] = "ListingTwo";
                active_name_list[2] = "ListingThree";
            }

            
            if (currentUrl.ToLower().Contains("goodmandean"))
            {
                comp_name_list[0] = "CS1";
                comp_name_list[1] = "CS2";
                comp_name_list[2] = "CS3";
                active_name_list[0] = "CL1";
                active_name_list[1] = "CL2";
                active_name_list[2] = "CL3";
                
            }

            if (currentUrl.ToLower().Contains("dispo"))
            {
                comp_name_list[0] = "0";
                comp_name_list[1] = "1";
                comp_name_list[2] = "2";
                active_name_list[0] = "0";
                active_name_list[1] = "1";
                active_name_list[2] = "2";

            }

            if (streetnumTextBox.Text == "nvs")
            {

                comp_name_list[0] = "CSale1";
                comp_name_list[1] = "CSale2";
                comp_name_list[2] = "CSale3";
                active_name_list[0] = "CList1";
                active_name_list[1] = "CList2";
                active_name_list[2] = "CList3";

            }

            #endregion

            string [] compPrices ={"","",""};


            for (int i = 0; i < 6; i++)
            {
                #region loop through the 6 comps

                //
                //Read MRED listing sheet
                //
                #region read tables
            

                string sale_or_list_flag = "";

                

                if (m.Status.Contains("CLSD"))
                {
                    input_comp_name = comp_name_list[closed_comps];
                    closed_comps = closed_comps + 1;
                    sale_or_list_flag = "sale";

                }
                else
                {
                    input_comp_name = active_name_list[active_comps];
                    active_comps = active_comps + 1;
                    sale_or_list_flag = "list";
                }
                #endregion
                //
                //
                //

                //
                //Save pics
                //
                #region save comp pic
                //open pic tab
                StringBuilder openTab = new StringBuilder();
                openTab.AppendLine(@"FRAME NAME=workspace");
                openTab.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:dc ATTR=CLASS:*navImage*");
                //openTab.AppendLine(@"FRAME F=0");
                //openTab.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=HREF:javascript:openWindow* EXTRACT=TXT");

                string openTab_macroCode = openTab.ToString();
                status = iim.iimPlayCode(openTab_macroCode, 30);

                openTab.Clear();

                StringBuilder save_pics_macro = new StringBuilder();

                //string filepath = "C:\\Users\\Scott\\Documents\\My<SP>Dropbox\\BPOs\\" + search_address_textbox.Text;
                string filepath = search_address_textbox.Text;
                string filename = "comp.jpg";

                if (sale_or_list_flag == "sale")
                {
                    save_pics_macro.AppendLine(@"ONDOWNLOAD FOLDER=" + filepath.Replace(" ","<SP>") + " FILE=S" + closed_comps.ToString() + ".jpg");
                }
                else
                {
                    save_pics_macro.AppendLine(@"ONDOWNLOAD FOLDER=" + filepath.Replace(" ", "<SP>") + " FILE=L" + active_comps.ToString() + ".jpg");
                }
                
                //line changed for new version of connectmls
                
                //save_pics_macro.AppendLine(@"'New tab opened");
               // save_pics_macro.AppendLine(@"WAIT SECONDS=2");
                save_pics_macro.AppendLine(@"TAB T=2");
                
                //save_pics_macro.AppendLine(@"TAG POS=2 TYPE=IMG FORM=NAME:dc ATTR=HREF:""*.JPEG"" CONTENT=EVENT:SAVEITEM");
                //save_pics_macro.AppendLine(@"WAIT SECONDS=5");
                //also changed for new connectmls
                //save_pics_macro.AppendLine(@"TAG POS=2 TYPE=A FORM=NAME:dc ATTR=TXT:Download<SP>This<SP>Photo CONTENT=EVENT:SAVETARGETAS");
               // save_pics_macro.AppendLine(@"TAG POS=2 TYPE=A FORM=NAME:dc ATTR=TXT:Download<SP>This<SP>Photo CONTENT=EVENT:SAVEITEM");
              
               // iim.iimPlayCode(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=HREF:javascript:openWindow('http%3A%2F%2Ftours.databasedads.com%2F2797441%2F1314-S-W* EXTRACT=TXT");
              

              

                save_pics_macro.AppendLine(@"FRAME F=0");

                if (iim.iimGetLastExtract().Contains("Virtual Tour"))
                {
                   // save_pics_macro.AppendLine(@"TAG POS=3 TYPE=IMG FORM=NAME:dc ATTR=HREF:""*.JPEG"" CONTENT=EVENT:SAVEITEM");
                    save_pics_macro.AppendLine(@"TAG POS=3 TYPE=IMG FORM=NAME:dc ATTR=HREF:* CONTENT=EVENT:SAVEITEM");
                }else
                {
                  //  save_pics_macro.AppendLine(@"TAG POS=2 TYPE=IMG FORM=NAME:dc ATTR=HREF:""*.JPEG"" CONTENT=EVENT:SAVEITEM");
                    save_pics_macro.AppendLine(@"TAG POS=2 TYPE=IMG FORM=NAME:dc ATTR=HREF:* CONTENT=EVENT:SAVEITEM");
                }

                
                //save_pics_macro.AppendLine(@"WAIT SECONDS=2");
                save_pics_macro.AppendLine(@"TAB CLOSE");

                 string save_pics_macroCode = save_pics_macro.ToString();
                status = iim.iimPlayCode(save_pics_macroCode, 30);
                #endregion
                //
                //
                //

                pBar1.PerformStep();
                # region read_mls_sheet
             
                string mlsnum = m.MlsNumber;
                string current_list_price = m.CurrentListPrice.ToString() ;
                string list_date = m.ListDateString;
                string orig_list_price = m.OriginalListPrice.ToString();
                string sold_price = m.SalePrice.ToString();
                string contract_date = m.ContractDate;
                string financingType = m.FinancingMlsString;

                if(sale_or_list_flag == "sale")
                {
                    compPrices[closed_comps -1] = sold_price;
                }
                
                string sale_type = m.TransactionType;
                string address = m.StreetAddress;
            
                string street_number = "";
                string street_name = "";
                string street_postfix = "";
                string city = m.City;
                string zip = m.Zipcode;
                string full_street_address = m.StreetAddress;

                if (zip.Contains("-"))
                {
                    zip = zip.Remove(zip.IndexOf("-"));
                }

                //need to remove unit in attached listing addresses.
                if (full_street_address.Contains("Unit"))
                {
                    full_street_address = full_street_address.Remove(full_street_address.IndexOf("Unit"));
                }




                //2316 W Fairview Ln , McHenry, Illinois 60051
                string pattern = "^(\\d+)\\s+\\w\\s+(\\w+)\\s+(\\w+)";
                Match match = Regex.Match(full_street_address, pattern);

                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;

                }

                //2316 Fairview Ln , McHenry, Illinois 60051
                pattern = "^(\\d+)\\s+(\\w\\w+)\\s+(\\w+)";
                match = Regex.Match(full_street_address, pattern);
                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;
                    if (street_postfix.Contains("Bay"))
                    {
                        street_postfix = "";
                    }

                }

                //1309 N Chapel Hill Rd , McHenry, Illinois 60051
                pattern = "^(\\d+)\\s+\\w\\s+(\\w+\\s+\\w+)\\s+(\\w+)";
                match = Regex.Match(full_street_address, pattern);
                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;
                }


                //769 White Pine Cir , Lake In The Hills, Illinois 60156
                pattern = "^(\\d+)\\s+(\\w\\w+\\s+\\w+)\\s+(\\w+)[^\\w+]";
                match = Regex.Match(full_street_address, pattern);
                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;
                }

                // 21727 NORTH Hickory Hill Dr 
                pattern = "^(\\d+)\\s+NORTH\\s+(\\w\\w+\\s+\\w+)\\s+(\\w+)";
                match = Regex.Match(full_street_address, pattern);
                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;
                }
                string dom = m.DOM;
                string closed_date = m.SalesDate.ToShortDateString();
                string year_built = m.YearBuilt.ToString();
                string county = m.County;
                string room_count = m.TotalRoomCount.ToString();
                string full_bath = m.FullBathCount;
                string half_bath = m.HalfBathCount;
                string bedrooms = m.BedroomCount;
                string number_of_firplaces = m.NumberOfFireplaces;
                bool hasFireplace = false;
                
                if (number_of_firplaces != "")
                    hasFireplace = true;

                string basement = m.BasementType;
                 bool fullBasement = false;
                 bool partialBasement = false;

                 if (basement.Contains("Full") || basement.Contains("English"))
                     fullBasement = true;
                 if (basement.Contains("Partial"))
                 {
                     partialBasement = true;
                 }

                 string parking = m.MredParkingString;
                 string mls_number_of_spaces = m.NumberGarageStalls();
                 string mls_subdivision = m.Subdivision;
                 
                if (string.IsNullOrWhiteSpace(mls_subdivision))
                {
                    mls_subdivision  = "Unk/NA";
                }

                string financing = m.FinancingMlsString;
                string type = m.Type; //mls type field
                string mls_garage_type = m.GarageType();
                string mls_garage_spaces = m.NumberGarageStalls().ToString();

                bool finished_basement = false;
                if (m.FinishedBasement)
                {
                    finished_basement = true;
                }

                string exteriorDetails = m.ExteriorMlsString;
                string pin = m.ParcelNumber;
                string mls_gla = m.GLA.ToString();

                mls_gla = mls_gla.Replace(",", "");
                string mls_lot_size = "0";
                mls_lot_size = m.Lotsize.ToString();
                string listing_Agent = m.ListingAgentName;
                string broker = m.ListingBrokerageName;
              string phone = m.ListingAgentPhone();
                string additionalSalesInfo = m.AdditionalSalesInfo();
              
                #endregion
           
                pBar1.PerformStep();
                #region Price Change extraction
                //
                //Price Change extraction
                //
                StringBuilder macro_get_price_changes = new StringBuilder();

                macro_get_price_changes.AppendLine(@"FRAME NAME=workspace");
                macro_get_price_changes.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:Additional<SP>Information");
                macro_get_price_changes.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:Listing<SP>&<SP>Property<SP>History");
                macro_get_price_changes.AppendLine(@"'New tab opened");
                //macro_get_price_changes.AppendLine(@"TAB T=2");
                macro_get_price_changes.AppendLine(@"TAG POS=1 TYPE=TABLE FORM=NAME:NoFormName ATTR=CLASS:innergridview EXTRACT=TXT");
                macro_get_price_changes.AppendLine(@"FRAME NAME=subheader");
               // macro_get_price_changes.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:header ATTR=ID:closeWinLink");
                macro_get_price_changes.AppendLine(@"TAB CLOSE");
                string macroCode = macro_get_price_changes.ToString();

                status = iim.iimPlayCode(macroCode, 60);

                string price_change_table = iim.iimGetLastExtract(1);

                //STATUS: PCHG -> ACTV#NEXT#ACTV#NEXT#$449,900 #NEXT#05/21/2012#NEXT#
                DateTime lastPriceChangeDate = new DateTime();
                DateTime firstPriceChangeDate = new DateTime();

                match = Regex.Match(price_change_table, @"STATUS: PCHG -> ACTV#NEXT#ACTV#NEXT#\$[\d,]+ #NEXT#(\d+.\d+.\d+)#NEXT#");
                MatchCollection mc = Regex.Matches(price_change_table, @"STATUS: PCHG -> ACTV#NEXT#ACTV#NEXT#\$[\d,]+ #NEXT#(\d+.\d+.\d+)#NEXT#");


                string domAtCurrentListPrice;
                string domAtOriginalListPrice;
                try
                {
                    lastPriceChangeDate = DateTime.Parse(match.Groups[1].Value);
                    firstPriceChangeDate = DateTime.Parse(mc[mc.Count - 1].Groups[1].Value);

                    domAtCurrentListPrice = (DateTime.Now - lastPriceChangeDate).Days.ToString();
                    domAtOriginalListPrice = (firstPriceChangeDate - DateTime.Parse(list_date)).Days.ToString();
                }
                catch
                {
                    lastPriceChangeDate = DateTime.Parse(list_date);
                    firstPriceChangeDate = lastPriceChangeDate;
                    domAtCurrentListPrice = dom;
                    domAtOriginalListPrice = dom;
                }

                  
           

                // domAtCurrentListPrice = (DateTime.Now - lastPriceChangeDate).Days.ToString();
               //  domAtOriginalListPrice = (DateTime.Parse(list_date) - firstPriceChangeDate).Days.ToString();

               // MessageBox.Show(domAtCurrentListPrice + ", " + domAtOriginalListPrice);


                int count = new Regex("-> PCHG").Matches(price_change_table).Count;



                //string searchTerm = "PCHG";
                ////Convert the string into an array of words
                //string[] source = price_change_table.Split(new char[] { '.', '?', '!', ' ', ';', ':', ',' }, StringSplitOptions.RemoveEmptyEntries);

                //// Create and execute the query. It executes immediately 
                //// because a singleton value is produced.
                //// Use ToLowerInvariant to match "data" and "Data" 
                //var matchQuery = from word in source
                //                 where word.ToLowerInvariant() == searchTerm.ToLowerInvariant()
                //                 select word;

                //// Count the matches.
                //int wordCount = matchQuery.Count();


                #endregion
                pBar1.PerformStep();

                #region realist extraction
                GoogleFusionTable realist_bpohelper = new GoogleFusionTable("1UKrOVmhPWrgLP5d5bDCsiW9whMIK8aLxKhcyOaI");
                await realist_bpohelper.helper_OAuthFusion();

                string realist_rawHtml = "";
                string realist_gla = "0";
                string realist_subdivision = "";
                string censusTract = "";
                string realist_schoolDistrict = "";
                string realist_lotAcres = "0";
                string realist_address = "";
            
              
                //StringBuilder realist_extraction_macro = new StringBuilder();
                //realist_extraction_macro.AppendLine(@"SET !TIMEOUT_STEP 10");
                //realist_extraction_macro.AppendLine(@"FRAME NAME=workspace");
                //realist_extraction_macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:Realist<SP>Tax<SP>Report");
                //realist_extraction_macro.AppendLine(@"'New tab opened");
                ////realist_extraction_macro.AppendLine(@"TAB T=2");
                //realist_extraction_macro.AppendLine(@"FRAME F=0");
                //realist_extraction_macro.AppendLine(@"SET !TIMEOUT_STEP 10");
                //realist_extraction_macro.AppendLine(@"TAG POS=4 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");
                //realist_extraction_macro.AppendLine(@"TAG POS=2 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");
                //realist_extraction_macro.AppendLine(@"TAB CLOSE");
                //string realist_extraction_macro_code = realist_extraction_macro.ToString();

                //status = iim.iimPlayCode(realist_extraction_macro_code, 60);

                //string realist_char_table = iim.iimGetLastExtract(1);
                //string realist_location_table = iim.iimGetLastExtract(2);
                
                ////
                ////Location Information
                ////
                //pattern = "Subdivision:#NEXT#([^#]+)";
                //match = Regex.Match(realist_location_table, pattern);
                // realist_subdivision = "";
                //realist_subdivision = match.Groups[1].Value;

                ////
                ////Characteristics
                ////
                ////Lot Acres:#NEXT#0.1257#NEXT#
                //pattern = "Lot Acres:#NEXT#([^#]+)";
                //match = Regex.Match(realist_char_table, pattern);
            
                //string realist_lot_size = "0";
                //realist_lot_size = match.Groups[1].Value;

                ////MessageBox.Show("MLS Lot: " + mls_lot_size + " " + "Realist Lot Size:" + realist_lot_size);

                //if (mls_lot_size == "0" | mls_lot_size == "")
                //{
                //    mls_lot_size = realist_lot_size;
                //}

                //pattern = "Building Above Grade Sq Ft:#NEXT#([^#]+)";
                //match = Regex.Match(realist_char_table, pattern);

                // realist_gla = "0";
                //realist_gla = match.Groups[1].Value;

                ////MessageBox.Show("MLS GLA: " + mls_gla + " " + "Realist GLA:" + realist_gla);

                //if (mls_gla == "0" | mls_gla == "")
                //{
                //    mls_gla = realist_gla.Replace(",","");
                //}

                //pattern = "Year Built:#NEXT#([^#]+)";
                //match = Regex.Match(realist_char_table, pattern);

                //string realist_yearbuilt = "";
                //realist_yearbuilt = match.Groups[1].Value;

                ////MessageBox.Show("MLS GLA: " + mls_gla + " " + "Realist GLA:" + realist_gla);

                ////MessageBox.Show("MLS YB: " + year_built + " " + "Realist YB:" + realist_yearbuilt);
             



                DateTime subject_age = new DateTime((Convert.ToInt32(SubjectYearBuilt)), 1, 1);

                TimeSpan ts = DateTime.Now - subject_age;

                int age = ts.Days / 365;


                if (!realist_bpohelper.RecordExists(pin))
                {
                    #region realist extraction update/add

                    #region open realist and extract data script

                    StringBuilder realist_extraction_macro = new StringBuilder();

                    //open realist
                    //realist_extraction_macro.AppendLine(@"SET !ERRORIGNORE YES");
                    realist_extraction_macro.AppendLine(@"SET !TIMEOUT_STEP 0");
                    realist_extraction_macro.AppendLine(@"FRAME NAME=workspace");
                    realist_extraction_macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:Realist<SP>Tax<SP>Report");

                    string realist_extraction_macro_code = realist_extraction_macro.ToString();
                    status = iim.iimPlayCode(realist_extraction_macro_code, 60);
                    realist_extraction_macro.Clear();

                    if (status == Status.sOk)
                    {
                        //see what we have opened
                        s = iim.iimPlayCode("SET !TIMEOUT_STEP 12 \r\n TAG POS=1 TYPE=TABLE ATTR=* EXTRACT=TXT");
                        realist_rawHtml = iim.iimGetLastExtract();
                        //realist window is open but no properties found so skip the steps below
                        if (!realist_rawHtml.Contains("No properties were found."))
                        {
                            //gets all html
                            s = iim.iimPlayCode("SET !TIMEOUT_STEP 12 \r\n TAG POS=1 TYPE=DIV ATTR=ID:htmlApplication EXTRACT=HTM");
                            realist_rawHtml = iim.iimGetLastExtract();
                            //gets the address
                            s = iim.iimPlayCode("SET !TIMEOUT_STEP 0 \r\n TAG POS=1 TYPE=DIV ATTR=ID:headerText EXTRACT=TXT");
                            realist_address = iim.iimGetLastExtract();
                        }

                    }


                    if (realist_rawHtml.Contains("Parcel"))
                    {

                        #region goodRecord
                        //Owner Information
                        realist_extraction_macro.AppendLine(@"TAG POS=1 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");
                        realist_extraction_macro.AppendLine(@"SET !TIMEOUT_STEP 1");
                        //Location Information
                        realist_extraction_macro.AppendLine(@"TAG POS=2 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");

                        //Tax Information Table
                        realist_extraction_macro.AppendLine(@"TAG POS=3 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");

                        //Assessment & Tax
                        //Assessment
                        realist_extraction_macro.AppendLine(@"TAG POS=1 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");
                        //Tax
                        realist_extraction_macro.AppendLine(@"TAG POS=2 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");


                        //Characteristics
                        realist_extraction_macro.AppendLine(@"TAG POS=4 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");


                        //Estimated Value
                        realist_extraction_macro.AppendLine(@"TAG POS=5 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");

                        //Listing Information
                        //realist_extraction_macro.AppendLine(@"TAG POS=6 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");
                        //most recent mls listing
                        //realist_extraction_macro.AppendLine(@"TAG POS=3 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");

                        //Last Market Sale & Sales History
                        //first row
                        //realist_extraction_macro.AppendLine(@"TAG POS=4 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");
                        //second row
                        //realist_extraction_macro.AppendLine(@"TAG POS=5 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");

                        //Mortgage History
                        //first row
                        //realist_extraction_macro.AppendLine(@"TAG POS=6 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");
                        //second row
                        //  realist_extraction_macro.AppendLine(@"TAG POS=7 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");

                        //Foreclosure History
                        //   realist_extraction_macro.AppendLine(@"TAG POS=8 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");



                        //realist_extraction_macro.AppendLine(@"TAB CLOSE");
                        realist_extraction_macro_code = realist_extraction_macro.ToString();

                        status = iim.iimPlayCode(realist_extraction_macro_code, 60);



                        string realist_owner_table = iim.iimGetLastExtract(1);
                        string realist_location_table = iim.iimGetLastExtract(2);
                        string realist_taxinfo_table = iim.iimGetLastExtract(3);
                        string realist_assessment_table = iim.iimGetLastExtract(4);
                        string realist_tax_table = iim.iimGetLastExtract(5);
                        string realist_char_table = iim.iimGetLastExtract(6);
                        string realist_estimatedvalue_table = iim.iimGetLastExtract(7);
                        string realist_listinginfo1_table = iim.iimGetLastExtract(8);
                        string realist_listinginfo2_table = iim.iimGetLastExtract(9);
                        string realist_saleshisotry1_table = iim.iimGetLastExtract(10);
                        string realist_saleshisotry2_table = iim.iimGetLastExtract(11);
                        string realist_morthisotry1_table = iim.iimGetLastExtract(12);
                        string realist_morthisotry2_table = iim.iimGetLastExtract(13);
                        string realist_foreclosure_table = iim.iimGetLastExtract(14);


                        //  s = iim.iimPlayCode(@"TAG POS=1 TYPE=DIV ATTR=ID:headerText EXTRACT=TXT");
                        //  realist_address = iim.iimGetLastExtract();

                        //  s = iim.iimPlayCode(@"TAG POS=1 TYPE=DIV ATTR=ID:htmlApplication EXTRACT=HTM");
                        //  realist_rawHtml = iim.iimGetLastExtract();

                        s = iim.iimPlayCode(@"TAB CLOSE");

                        //          MessageBox.Show(realist_address + realist_rawHtml);



                        string allTables = realist_owner_table + "#NEXT#" + realist_location_table + "#NEXT#" + realist_taxinfo_table + "#NEXT#" + realist_char_table + "#NEXT#" + realist_estimatedvalue_table;



                        allTables.Replace("# #", "##");


                        //Township:#NEXT#Grafton Twp#NEXT#Census Tract:#NEXT#8711.04#NEXT##NEWLINE#Township Range Sect:#NEXT#43N-7E-27#NEXT#Carrier Route:#NEXT#R010#NEXT##NEWLINE#Subdivision:#NEXT#Pasquinellis Huntley Mdws#NEXT##NEXT# 

                        string[] sep = { "#NEWLINE#" };


                        string[] splitonnewline = allTables.Split(sep, StringSplitOptions.RemoveEmptyEntries);


                        string valuePattern = @"#NEXT#([^#]+)#NEXT#";
                        string namePattern = @"^([^#]+)#NEXT#|#NEXT#([^#]+)#NEXT#";
                        string fieldName;
                        string fieldValue;
                        string parid = Regex.Match(allTables, @"Parcel ID:#NEXT#([^#]+)").Groups[1].Value;

                        Dictionary<string, string> realistReportNameValuePairs = new Dictionary<string, string>();
                        #endregion

                        #region parse data and add/update fusion table
                        //if we found a parcel id, which is the key for the record
                        if (!String.IsNullOrEmpty(parid))
                        {

                            Stopwatch stopWatch = new Stopwatch();
                            stopWatch.Start();
                            string[] gFormatedLocation = new string[3];


                            gFormatedLocation = realist_address.Split(',');




                            if (!realist_bpohelper.RecordExists(parid))
                            {
                                //if realist is missing addess, split will only have 2 fields, not 3
                                try
                                {
                                    realist_bpohelper.AddRecord(parid, gFormatedLocation[0] + " " + gFormatedLocation[1] + " " + gFormatedLocation[2]);
                                }
                                catch
                                {
                                    realist_bpohelper.AddRecord(parid, address.Replace(",", ""));
                                }
                            }
                            else
                            {
                                //add address to pending updates, incase it's missing or needs updating in fusion table
                                realistReportNameValuePairs.Add("Location", gFormatedLocation[0] + " " + gFormatedLocation[1] + " " + gFormatedLocation[2]);

                            }





                            //
                            //Check and add new colums
                            //
                            foreach (string valuePairLine in splitonnewline)
                            {


                                string modStr = Regex.Replace(valuePairLine, @"# ", "Number").Replace("Garage #", "Garage Number");


                                MatchCollection mc1 = Regex.Matches(modStr, namePattern);
                                foreach (Match m1 in mc1)
                                {


                                    string v = m1.Value.Replace("#NEXT#", "").Trim().Replace("(", @"\(").Replace(")", @"\)") + "#NEXT#([^#]+)";
                                    if (!v.Contains("NODATA"))
                                    {
                                        if (realist_bpohelper.ColumnExsists(m1.Value.Replace("#NEXT#", "").Trim()))
                                        {


                                            realistReportNameValuePairs.Add(m1.Value.Replace("#NEXT#", "").Trim(), Regex.Match(modStr, v).Groups[1].Value);
                                            //realist_bpohelper.UpdateRecord(parid, m1.Value.Replace("#NEXT#", "").Trim(), Regex.Match(modStr, v).Groups[1].Value);
                                            //MessageBox.Show(m1.Value + "exsists.");
                                        }
                                        else
                                        {
                                            if (m1.Value != "NEXT")
                                            {
                                                realist_bpohelper.AddColumn(m1.Value.Replace("#NEXT#", "").Trim());

                                                try
                                                {
                                                    realistReportNameValuePairs.Add(m1.Value.Replace("#NEXT#", "").Trim(), Regex.Match(modStr, v).Groups[1].Value);
                                                }
                                                catch
                                                {
                                                    MessageBox.Show("problem adding realist nvp.");
                                                }


                                            }
                                        }
                                    }

                                }

                            }


                            //
                            //Update record
                            //

                            realistReportNameValuePairs.Concat(realist_bpohelper.Geocode(address));

                            realist_bpohelper.UpdateRecord(parid, realistReportNameValuePairs);

                            realist_bpohelper.m_rowid = "";


                            // MessageBox.Show(splitonnewline[0]);

                        }



                        //
                        //Location Information
                        //
                        pattern = "Subdivision:#NEXT#([^#]+)";
                        match = Regex.Match(realist_location_table, pattern);

                        realist_subdivision = match.Groups[1].Value;

                        //School District:
                        pattern = "School District:#NEXT#([^#]+)";
                        match = Regex.Match(realist_location_table, pattern);

                        realist_schoolDistrict = match.Groups[1].Value;

                        //Lot Acres:#NEXT#0.1257#NEXT#
                        pattern = "Lot Acres:#NEXT#([^#]+)";
                        match = Regex.Match(realist_char_table, pattern);

                        string realist_lot_size = "0";
                        realist_lot_size = match.Groups[1].Value;
                        realist_lotAcres = realist_lot_size;

                        //MessageBox.Show("MLS Lot: " + mls_lot_size + " " + "Realist Lot Size:" + realist_lot_size);

                        if (mls_lot_size == "0" | mls_lot_size == "")
                        {
                            mls_lot_size = realist_lot_size;
                        }

                        pattern = "Building Above Grade Sq Ft:#NEXT#([^#]+)";
                        match = Regex.Match(realist_char_table, pattern);


                        realist_gla = match.Groups[1].Value;

                        //MessageBox.Show("MLS GLA: " + mls_gla + " " + "Realist GLA:" + realist_gla);
                        if (match.Success)
                        {
                            if (mls_gla == "0" | mls_gla == "")
                            {
                                mls_gla = realist_gla;
                            }
                        }
                        else
                        {
                            realist_gla = "-1";
                        }
                        pattern = "Year Built:#NEXT#([^#]+)";
                        match = Regex.Match(realist_char_table, pattern);

                        string realist_yearbuilt = "";
                        realist_yearbuilt = Regex.Match(match.Groups[1].Value, @"(\d\d\d\d)").Value;

                        //MessageBox.Show("MLS GLA: " + mls_gla + " " + "Realist GLA:" + realist_gla);

                        //MessageBox.Show("MLS YB: " + year_built + " " + "Realist YB:" + realist_yearbuilt);
                        if (string.IsNullOrEmpty(year_built) || year_built.ToLower() == "unk" || year_built == "0")
                        {
                            if (!string.IsNullOrEmpty(realist_yearbuilt))
                            {
                                year_built = realist_yearbuilt;
                            }
                            else
                            {
                                year_built = "1950";
                            }
                        }



                         subject_age = new DateTime((Convert.ToInt32(year_built)), 1, 1);

                         ts = DateTime.Now - subject_age;

                         age = ts.Days / 365;


                        pattern = @"Census Tract:#NEXT#([^#]+)";
                        match = Regex.Match(realist_char_table, pattern);

                        if (match.Success)
                        {
                            censusTract = match.Groups[1].Value;
                        }
                        else
                        {
                            match = Regex.Match(realist_location_table, pattern);
                            censusTract = match.Groups[1].Value;
                        }


                    }
                        #endregion //cccc

                    #endregion  //ccccdsddd

                    else
                    {
                        //unusable record OR
                        //No realist pop-up ie Wisconson address

                        s = iim.iimPlayCode("SET !TIMEOUT_STEP 0 \r\n TAB T=2 \r\n TAB CLOSE");
                        if (string.IsNullOrEmpty(year_built) || year_built.ToLower() == "unk" || year_built == "0")
                        {
                            year_built = "1950";
                        }



                         subject_age = new DateTime((Convert.ToInt32(year_built)), 1, 1);

                         ts = DateTime.Now - subject_age;

                         age = ts.Days / 365;

                    }




                }
                    #endregion

                else
                {
                    #region realist read from fusion table
                    try
                    {
                        realist_lotAcres = realist_bpohelper.curRec["Lot Acres:"];
                    }
                    catch
                    {

                    }

                    if (mls_lot_size == "0" | mls_lot_size == "")
                    {
                        mls_lot_size = realist_lotAcres;
                        double x;
                        Double.TryParse(realist_lotAcres, out x);
                        m.Lotsize = x;
                    }

                    realist_gla = realist_bpohelper.curRec["Building Above Grade Sq Ft:"];
                    if (String.IsNullOrEmpty(realist_gla))
                    {
                        realist_gla = "-1";
                    }
                    else
                    {
                        if (mls_gla == "0" | mls_gla == "")
                        {
                            mls_gla = realist_gla;
                            int x;
                            Int32.TryParse("mls_gla", out x);
                            m.GLA = x;
                        }
                    }
                    realist_subdivision = realist_bpohelper.curRec["Subdivision:"];
                    censusTract = realist_bpohelper.curRec["Census Tract:"];
                    realist_schoolDistrict = realist_bpohelper.curRec["School District:"];
                    if (string.IsNullOrEmpty(year_built) || year_built.ToLower() == "unk" || year_built == "0")
                    {
                        if (realist_bpohelper.curRec["Year Built:"] != "")
                        {
                            if (realist_bpohelper.curRec["Year Built:"].Contains("Tax:"))
                            {
                                year_built = Regex.Match(realist_bpohelper.curRec["Year Built:"], @"Tax:\s*(\d\d\d\d)").Groups[1].Value;

                            }
                            else if (realist_bpohelper.curRec["Year Built:"].Contains("MLS:"))
                            {
                                year_built = Regex.Match(realist_bpohelper.curRec["Year Built:"], @"MLS:\s*(\d\d\d\d)").Groups[1].Value;
                            }
                            else
                            {
                                year_built = realist_bpohelper.curRec["Year Built:"];
                            }

                        }
                        else
                        {
                            year_built = "1950";
                        }
                    }
                    realist_bpohelper.Geocode(address);

                    #endregion
                }

                string propertyLatitude = "";
                string propertyLongitude = "";

                int xx = 0;
                Int32.TryParse(realist_gla.Replace(",", ""), out xx);
                m.RealistGLA = xx;

                double dd = 0;
                Double.TryParse(realist_lotAcres, out dd);
                m.RealistLotSize = dd;

                try
                {

                    propertyLatitude = realist_bpohelper.curRec["Latitude"];
                    propertyLongitude = realist_bpohelper.curRec["Longitude"];

                    GeoPoint destPoint;
                    try
                    {
                        destPoint.Latitude = Convert.ToDouble(propertyLatitude);
                        destPoint.Longitude = Convert.ToDouble(propertyLongitude);
                    }
                    catch
                    {
                        destPoint = GlobalVar.subjectPoint;
                    }

                    MlsReportDriver x = new MlsReportDriver();

                    m.proximityToSubject = Math.Round(x.Get_Distance(GlobalVar.subjectPoint, destPoint), 2);

                }
                catch
                {

                }

                #endregion
                pBar1.PerformStep();
                

                // Perform the increment on the ProgressBar.


                //
                //REOcBPO - First Pass
                //
                #region REOcBPO
                if (currentUrl.ToLower().Contains("reo-central"))
                {
                    #region code

                    s = iim.iimPlayCode(macro12.ToString());
                    htmlCode = iim.iimGetLastExtract();
                    if (subjectAttachedRadioButton.Checked)
                    {
                        m = new AttachedListing(htmlCode);
                    }
                    else if (subjectDetachedradioButton.Checked)
                    {
                        m = new DetachedListing(htmlCode);
                    }
                    else
                    {
                        m = new MLSListing(htmlCode);
                    }

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    REOcBPO bpoform = new REOcBPO(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion

                    #region basementlogic




                    #endregion

                    #region garagelogic

                    #endregion

                    #region fireplacelogic

                    #endregion

                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                }

                #endregion  

                //
                //solutionstar - First Pass
                //
                #region solutionstar
                if (currentUrl.ToLower().Contains("solutionstar"))
                {
                    #region code

                    s = iim.iimPlayCode(macro12.ToString());
                    htmlCode = iim.iimGetLastExtract();
                    if (subjectAttachedRadioButton.Checked)
                    {
                        m = new AttachedListing(htmlCode);
                    }
                    else if (subjectDetachedradioButton.Checked)
                    {
                        m = new DetachedListing(htmlCode);
                    }
                    else
                    {
                        m = new MLSListing(htmlCode);
                    }

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    SolutionStar bpoform = new SolutionStar(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion

                    #region basementlogic




                    #endregion

                    #region garagelogic

                    #endregion

                    #region fireplacelogic

                    #endregion

                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                }

                #endregion  

                //
                //AVM - First Pass
                //
                #region AVM
                if (currentUrl.ToLower().Contains("avm.assetval.com"))
                {
                    #region code

                    s = iim.iimPlayCode(macro12.ToString());
                    htmlCode = iim.iimGetLastExtract();
                    if (subjectAttachedRadioButton.Checked)
                    {
                        m = new AttachedListing(htmlCode);
                    }
                    else if (subjectDetachedradioButton.Checked)
                    {
                        m = new DetachedListing(htmlCode);
                    }
                    else
                    {
                        m = new MLSListing(htmlCode);
                    }

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    AVM bpoform = new AVM(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion

                    #region basementlogic




                    #endregion

                    #region garagelogic

                    #endregion

                    #region fireplacelogic

                    #endregion

                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                }

                #endregion  

                //
                //M2M Fannie Mae Form
                //
                #region M2M Fannie Mae
                if (currentUrl.ToLower().Contains("ordereditwizardfnma"))
                {
                    #region code
                    
                    s = iim.iimPlayCode(macro12.ToString());
                    htmlCode = iim.iimGetLastExtract();
                    if (subjectAttachedRadioButton.Checked)
                    {
                        m = new AttachedListing(htmlCode);
                    }
                    else if (subjectDetachedradioButton.Checked)
                    {
                        m = new DetachedListing(htmlCode);
                    }
                    else
                    {
                        m = new MLSListing(htmlCode);
                    }

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    M2M bpoform = new M2M(m);

                    fieldList.Add("filepath", SubjectFilePath);
                 
                    #endregion

                    #region basementlogic

                    


                    #endregion

                    #region garagelogic
                  
                    #endregion

                    #region fireplacelogic
               
                    #endregion

                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                 
                }

                #endregion  

                //
                //Dispo
                //
                #region Dispo
                if (currentUrl.ToLower().Contains("dispo"))
                {
                    #region code
                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    Dispo bpoform = new Dispo();

                    fieldList.Add("Address_StreetAddress", full_street_address);
                    fieldList.Add("Address_City", city);
                    fieldList.Add("*Address_State", "IL");
                    fieldList.Add("Address_ZipCode", zip);
                    fieldList.Add("Address_County", county);
                    fieldList.Add("ProximityToSubject_EstimatedDistance", Get_Distance(SubjectFullAddress.Replace("," ," "), full_street_address + " " + city + " IL " + zip ));
                    fieldList.Add("SubdivisionName", realist_subdivision);

                    fieldList.Add("ListingHistory_OriginalListPrice", orig_list_price);
                    fieldList.Add("ListingHistory_CurrentListPrice", current_list_price);
                    fieldList.Add("ListingHistory_CurrentListPriceDate", lastPriceChangeDate.ToShortDateString());
                    fieldList.Add("ListingHistory_DaysOnMarket", dom);

                    fieldList.Add("SaleHistory_OriginalListPrice", orig_list_price);
                    fieldList.Add("SaleHistory_FinalListPrice", current_list_price);
                    fieldList.Add("SaleHistory_SalePrice", sold_price);
                    fieldList.Add("ClosingDetails_UnderContractDate", contract_date);
                    fieldList.Add("ClosingDetails_ClosingDate", closed_date);
                    fieldList.Add("SaleHistory_DaysOnMarket", dom);

                    if (m.TransactionType == "REO")
                    {
                        fieldList.Add("*TransactionType", "1");
                    }
                    else if (m.TransactionType == "Short Sale")
                    {
                        fieldList.Add("*TransactionType", "2");
                    } else
                    {
                        fieldList.Add("*TransactionType", "3");
                    }

             



                    if (SubjectMlsType.ToLower().Contains("1 story"))
                    {
                        fieldList.Add("*Structure_StyleType", "1<SP>Story");
                    }
                    else if (SubjectMlsType.ToLower().Contains("2 stories"))
                    {
                        fieldList.Add("*Structure_StyleType", "2<SP>Story");
                    }
                    else if (SubjectMlsType.ToLower().Contains("raised ranch"))
                    {
                        fieldList.Add("*Structure_StyleType", "Multilevel");
                    }
                    else if (SubjectMlsType.ToLower().Contains("split") || SubjectMlsType.ToLower().Contains("other"))
                    {
                        fieldList.Add("*Structure_StyleType", "Multilevel");
                    }
                    else if (SubjectMlsType.ToLower().Contains("1.5"))
                    {
                        fieldList.Add("*StyleType", "1.5<SP>Story");
                    }

                    fieldList.Add("*PropertyType", "1");
                    fieldList.Add("Site_LotSize", mls_lot_size);

                    fieldList.Add("Structure_YearBuilt", year_built);
                    fieldList.Add("Structure_LivingAreaSquareFeet", mls_gla);
                    fieldList.Add("Structure_GradedRooms_TotalRoomCount", room_count);
                    fieldList.Add("Structure_GradedRooms_BedroomCount", bedrooms);
                    fieldList.Add("Structure_GradedRooms_FullBathCount", full_bath);
                    fieldList.Add("Structure_GradedRooms_PowderRoomCount", half_bath);

                   

                    #endregion

                    #region basementlogic

                    //Basement 
                    if (basement.ToLower().Contains("none"))
                    {
                        fieldList.Add("*SquareFeet", "0");
                        fieldList.Add("Structure*Basement*PercentFinished", "0");
                        fieldList.Add("BasementType", "crawl");
                    }
                    else
                    {
                        if (fullBasement)
                        {
                            fieldList.Add("BasementType", "Full");
                            if (type == "1 Story")
                            {
                                fieldList.Add("Structure*Basement*SquareFeet", mls_gla);
                            }
                            else if (type == "2 Stories")
                            {
                                fieldList.Add("Structure*Basement*SquareFeet", (Convert.ToInt64(mls_gla) / 2).ToString());
                            }
                            else
                            {
                                fieldList.Add("Structure*Basement*SquareFeet", (Convert.ToInt64(mls_gla) / 3).ToString());
                            }
                        }

                        if (partialBasement)
                        {
                            fieldList.Add("BasementType", "Partial");
                            if (type == "1 Story")
                            {
                                fieldList.Add("Structure*Basement*SquareFeet", (Convert.ToInt64(mls_gla) / 2).ToString());
                            }
                            else if (type == "2 Stories")
                            {
                                fieldList.Add("Structure*Basement*SquareFeet", (Convert.ToInt64(mls_gla) / 4).ToString());
                            }
                            else
                            {
                                fieldList.Add("Structure*Basement*SquareFeet", (Convert.ToInt64(mls_gla) / 6).ToString());
                            }
                        }

                        if (finished_basement)
                        {
                            fieldList.Add("Structure*Basement*PercentFinished", "100");
                        }
                        else
                        {
                            fieldList.Add("Structure*Basement*PercentFinished", "0");
                        }
                    }


                    #endregion

                    #region garagelogic
                    string garStr = "";
                    if (mls_garage_spaces == "1")
                    {
                        garStr = "1<SP>Car";
                    } else
                    {
                         garStr = mls_garage_spaces + "<SP>Cars";
                    }


                    if (mls_garage_type.ToLower().Contains("att"))
                    {
                        fieldList.Add("*Structure_Garages_AttachedCarCountType", garStr);
                    }
                    if (mls_garage_type.ToLower().Contains("det"))
                    {
                        fieldList.Add("*Structure_Garages_DetachedCarCountType", garStr);
                    }
                    #endregion

                    #region fireplacelogic
                    fieldList.Add("Fireplace", hasFireplace.ToString());
                    #endregion

                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                }

                #endregion  
                //
                //Goodman Dean
                //
                #region goodmandean code

                if (currentUrl.ToLower().Contains("goodmandean"))
                {

                    GoodmanDean bpoform = new GoodmanDean();
                    Dictionary<string, string> fieldList = new Dictionary<string, string>();

                    fieldList.Add("Address", full_street_address);
                    fieldList.Add("City", city);
                    fieldList.Add("Zip", zip);
                    fieldList.Add("Distance", Get_Distance(full_street_address + ", " + zip, SubjectFullAddress));

                    //fieldList.Add("PRICEREVDT", lastPriceChangeDate.ToString("d", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                    //fieldList.Add("ReductionCount", count.ToString());
                    fieldList.Add("Design", type);
                    fieldList.Add("TotalRooms", room_count);
                    fieldList.Add("Bed", bedrooms);
                    //fieldList.Add("Bath", full_bath + "." + half_bath);
                    fieldList.Add("Bath", full_bath);
                    fieldList.Add("HalfBath", half_bath);
                    fieldList.Add("MLS", mlsnum);

                    //
                    //bpoform10
                    //
                    if (currentUrl.ToLower().Contains("bpoform10"))
                    {

                        //property characteristics
                        fieldList.Add("HOA", "0");
                        fieldList.Add("Location", "Suburban");
                        fieldList.Add("Other", realist_subdivision);
                        fieldList.Add("*Condition", "Average");
                        fieldList.Add("*Overall", "Similar");
                        fieldList.Add("Comparison", "Similar");
                        fieldList.Add("*PropPool", "0");

                        //sale characteristics
                        fieldList.Add("OriginalList", orig_list_price);
                        
                        fieldList.Add("View", "Residential");
                        fieldList.Add("Units", "1");

                        //fields that need logic
                        if (string.IsNullOrEmpty(financingType))
                        {
                            fieldList.Add("LoanType", "NA");
                        }
                        else
                        {
                            fieldList.Add("LoanType", financingType);
                        }
                        if (fullBasement && finished_basement)
                        {
                            fieldList.Add("*BasementType", "Full");
                            fieldList.Add("BsmtSqFt", mls_gla);
                            fieldList.Add("BsftFin", "100");
                        }
                        else if (fullBasement && !finished_basement)
                        {
                            fieldList.Add("*BasementType", "Partial");
                            fieldList.Add("BsmtSqFt", mls_gla);
                            fieldList.Add("BsftFin", "0");
                        }
                        else if (partialBasement && !finished_basement)
                        {
                            fieldList.Add("*BasementType", "Partial");
                            fieldList.Add("BsmtSqFt", (Convert.ToInt64(mls_gla) / 2).ToString());
                            fieldList.Add("BsftFin", "0");
                        }
                        else if (partialBasement && finished_basement)
                        {
                            fieldList.Add("*BasementType", "Partial");
                            fieldList.Add("BsmtSqFt", (Convert.ToInt64(mls_gla) / 2).ToString());
                            fieldList.Add("BsftFin", "100");
                        }
                        else
                        {
                            fieldList.Add("*BasementType", "None");
                            fieldList.Add("BsmtSqFt", "0");
                            fieldList.Add("BsftFin", "0");
                        }

                        if (sale_type == "Normal")
                        {
                            fieldList.Add("*DistressedSale", "No");
                        }
                        else if (sale_type == "Short")
                        {
                            fieldList.Add("*DistressedSale", "Short<SP>Sale");
                        }
                        else if (sale_type == "REO")
                        {
                            fieldList.Add("*DistressedSale", "REO");
                        }
                    }

                    if (parking.ToLower().Contains("gar"))
                    {
                        if (mls_garage_type.ToLower().Contains("att"))
                        {
                            //attached gar
                            fieldList.Add("ParkingType", "Gar<SP>Attached");
                            fieldList.Add("ParkingStalls", mls_garage_spaces);
                        }
                        else
                        {
                            //we dont care, well call it det (if we have type gar and spaces and its not attached)
                            fieldList.Add("ParkingType", "Gar<SP>Dettached");
                            fieldList.Add("ParkingStalls", mls_garage_spaces);
                        }

                    }
                    else
                    {
                        fieldList.Add("ParkingType", "None");
                        fieldList.Add("ParkingStalls", "0");
                    }


                    if (sale_type == "Normal")
                    {
                        fieldList.Add("SaleType", "NO");
                    }
                    else if (sale_type == "Short")
                    {
                        fieldList.Add("SaleType", "SHORTSALE");
                    }
                    else if (sale_type == "REO")
                    {
                        fieldList.Add("SaleType", "REO");
                    }
                    //sale_type


                    //fieldList.Add("PROXIMITY", this.Get_Distance(SubjectFullAddress, address));
                    //fieldList.Add("PROPTYPE", type);

                    if (sale_or_list_flag == "sale")
                    {
                        fieldList.Add("List", current_list_price);
                        fieldList.Add("Sale", sold_price);
                        fieldList.Add("ListDate", closed_date);
                        fieldList.Add("SaleDate", contract_date);
                    }
                    else
                    {
                        fieldList.Add("List", orig_list_price);
                        fieldList.Add("SalePrice", current_list_price);
                        fieldList.Add("ListDate", list_date);
                        fieldList.Add("SaleDate", closed_date);
                    }
                    
                    fieldList.Add("DOM", dom);


                    fieldList.Add("Lot", mls_lot_size);
                    fieldList.Add("YearBuilt", year_built);
                    fieldList.Add("SqFt", mls_gla);

                    // fieldList.Add("LIVINGSQFT", mls_gla);

                    //fieldList.Add("*LeaseHoldFeeSimple", "Fee Simple");
                    //fieldList.Add("*State", "IL");
                   // fieldList.Add("*NumUnits", "1 SFR");
                    
                   
                   // fieldList.Add("*DesignAppeal", "Average");
                   // fieldList.Add("*QualityOfConstruction", "Average");
               
                  //  fieldList.Add("*FunctionalUtility", "Average");


                    fieldList.Add("Concessions", "NA");
                  //  fieldList.Add("HeatingCooling", "GFA.C");
                   // fieldList.Add("EnergyEfficientItems", "na");
                  //  fieldList.Add("OtherItems", "na");






                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                }

                #endregion  
                //
                //Old Republic
                //
                #region oldrepublic code

                if (currentUrl.ToLower().Contains("ort.quandis"))
                {

                    OldRep bpoform = new OldRep();
                    Dictionary<string, string> fieldList = new Dictionary<string, string>();

                    fieldList.Add("Address", full_street_address);
                    fieldList.Add("City", city);
                    fieldList.Add("PostalCode", zip);

                    
                    //fieldList.Add("PRICEREVDT", lastPriceChangeDate.ToString("d", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                    fieldList.Add("ReductionCount", count.ToString());

                    fieldList.Add("Rooms", room_count);

                  
                    fieldList.Add("Bedrooms", bedrooms);
                    fieldList.Add("Bathrooms", full_bath + "." + half_bath);
                 
                    //needs logic for selection list
                    //needs logic for selection list
                    
                    if (fullBasement && finished_basement)
                    {
                        fieldList.Add("*Basement", "90% Finished");
                    }
                    else if (fullBasement || partialBasement && !finished_basement) 
                    {
                        fieldList.Add("*Basement", "Unfinished");
                    }
                    else if (partialBasement && finished_basement)
                    {
                        fieldList.Add("*Basement", "50% Finished");
                    }
                    else
                    {
                        fieldList.Add("*Basement", "None");
                    }


                    if (parking.ToLower().Contains("gar"))
                    {
                        if (mls_garage_type.ToLower().Contains("att"))
                        {
                            //attached gar
                            fieldList.Add("*Parking", mls_garage_spaces + "CarAtt");
                        }
                        else
                        {
                            //we dont care, well call it det (if we have type gar and spaces and its not attached)
                            fieldList.Add("*Parking", mls_garage_spaces + "CarDet");
                        }
                   
                    }
                    else
                    {
                        fieldList.Add("*Parking", "None");
                    }


                    if (sale_type == "Normal")
                    {
                        fieldList.Add("*SaleType", "Standard");
                    }
                    else if (sale_type == "Short")
                    {
                        fieldList.Add("*SaleType", "Short Sale");
                    }
                    else if (sale_type == "REO")
                    {
                        fieldList.Add("*SaleType", "REO");
                    }
                    //sale_type
                
                   
                    //fieldList.Add("PROXIMITY", this.Get_Distance(SubjectFullAddress, address));
                    //fieldList.Add("PROPTYPE", type);

                    if (sale_or_list_flag == "sale")
                    {
                        fieldList.Add("ListPrice", current_list_price);
                        fieldList.Add("SalePrice", sold_price);
                    }
                    else
                    {
                        fieldList.Add("ListPrice", orig_list_price);
                        fieldList.Add("SalePrice", current_list_price);
                    }
                    
                    fieldList.Add("SaleDate", closed_date);
                    fieldList.Add("DOM", dom);
                    
                   
                    fieldList.Add("Site", mls_lot_size);
                    fieldList.Add("Age", year_built);
                    fieldList.Add("SquareFeet", mls_gla);
                   // fieldList.Add("LIVINGSQFT", mls_gla);

                    fieldList.Add("*LeaseHoldFeeSimple", "Fee Simple");
                    fieldList.Add("*State", "IL");
                    fieldList.Add("*NumUnits", "1 SFR");
                    fieldList.Add("*LocationCondition", "Suburban");
                    fieldList.Add("*View", "Average");
                    fieldList.Add("*DesignAppeal", "Average");
                    fieldList.Add("*QualityOfConstruction", "Average");
                    fieldList.Add("*Condition", "Average");
                    fieldList.Add("*FunctionalUtility", "Average");
                    fieldList.Add("*DataSource", "MLS");

                    fieldList.Add("Incentive", "na");
                    fieldList.Add("HeatingCooling", "GFA.C");
                    fieldList.Add("EnergyEfficientItems", "na");
                    fieldList.Add("OtherItems", "na");
                    fieldList.Add("*Style", bpoform.StyleString(type));





                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                }

                #endregion  
                //
                //Landsafe
                //
                #region landsafe
                if (currentUrl.ToLower().Contains(@"collateraldna.com/forms"))
                {
                    LandSafe bpoform = new LandSafe();
                    Dictionary<string, string> fieldList = new Dictionary<string, string>();

                    s = iim.iimPlayCode(macro12.ToString());
                    htmlCode = iim.iimGetLastExtract();
                    if (subjectAttachedRadioButton.Checked)
                    {
                        m = new AttachedListing(htmlCode);
                    }
                    else if (subjectDetachedradioButton.Checked)
                    {
                        m = new DetachedListing(htmlCode);
                    }
                    else
                    {
                        m = new MLSListing(htmlCode);
                    }

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    fieldList.Add("ADDRSTREET", full_street_address);
                    fieldList.Add("STREET", full_street_address);
                    fieldList.Add("ADDRCITY", city);
                    fieldList.Add("CITY", city);
                    fieldList.Add("ADDRSTATEPROV", "IL");
                    fieldList.Add("STATEPROV", "IL");
                    fieldList.Add("ADDRZIP", zip);
                    fieldList.Add("ZIP", zip);
                    fieldList.Add("LISTINGPRICE", current_list_price);
                    fieldList.Add("PRICEREVDT", lastPriceChangeDate.ToString("d",System.Globalization.DateTimeFormatInfo.InvariantInfo));
                    fieldList.Add("BASEMENT", basement);
                    fieldList.Add("BEDROOMS", bedrooms);
                    fieldList.Add("BATH", full_bath + "." + half_bath);
                    fieldList.Add("GARAGE", mls_garage_type);
                    fieldList.Add("PROXIMITY", this.Get_Distance(SubjectFullAddress, address));
                    fieldList.Add("SALESPRICE", sold_price);
                    fieldList.Add("SALEDT", m.SalesDate.ToString("d", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                    fieldList.Add("MARKETDAYS", dom);
                    fieldList.Add("PROPTYPE", type);
                    fieldList.Add("LOTSIZE", mls_lot_size);
                    fieldList.Add("YRBUILT", year_built);
                    fieldList.Add("GBASQFT", mls_gla);
                    fieldList.Add("LIVINGSQFT", mls_gla);
                    fieldList.Add("DESIGNSTYLE", m.ExteriorMlsString);

                    if(m.TransactionType == "REO" || m.TransactionType == "ShortSale")
                    {
                            fieldList.Add("*SALETYPE", "YES");
                    }
                    else
                    {
                            fieldList.Add("*SALETYPE", "NO");
                    }
                

                    

                    //try
                    //{
                    //    if (Convert.ToInt32(half_bath) != 0)
                    //    {
                    //        fieldList.Add("*BA", full_bath + full_bath);
                    //    }
                    //    else
                    //    {
                    //        fieldList.Add("*BA", full_bath);
                    //    }
                    //}

                    //catch
                    //{
                    //    fieldList.Add("*BA", full_bath);
                    //}



                    //fieldList.Add("*Type", "s");



                    ////fieldList.Add("*Style", "2 Story Conv");


                    //fieldList.Add("*Loc", "r");

                    //fieldList.Add("GLASqFt", mls_gla);
                    //fieldList.Add("EstValSqFt", "0");
                    //if (finished_basement)
                    //{
                    //    fieldList.Add("BsmtPerFin", "100");
                    //}
                    //else
                    //{
                    //    fieldList.Add("BsmtPerFin", "0");
                    //}

                    //fieldList.Add("EstValBsmtPerFin", "0");
                    //fieldList.Add("YrBlt", year_built);
                    //fieldList.Add("EstValYrBlt", "0");
                    //fieldList.Add("LotSize", mls_lot_size + "ac");
                    //fieldList.Add("EstValLotSize", "0");
                    //fieldList.Add("DOM", dom);
                    //fieldList.Add("EstValDOM", "0");
                    //fieldList.Add("Other", "n/a");
                    //fieldList.Add("EstValOther", "0");
                    //if (parking.ToLower().Contains("gar"))
                    //{
                    //    fieldList.Add("*Gar", mls_garage_spaces);
                    //}
                    //else
                    //{
                    //    fieldList.Add("*Gar", "n");
                    //}
                    //fieldList.Add("EstValGar", "0");
                    //fieldList.Add("*Pool", "n");


                    //fieldList.Add("ExtAdd", "0");
                    //fieldList.Add("EstValExt", "0");
                    //fieldList.Add("*Land", "r");
                    //fieldList.Add("Fin", "n/a");
                    //fieldList.Add("EstValFin", "0");
                    //fieldList.Add("*Cond", "a");
                    //fieldList.Add("EstValCond", "0");
                    //fieldList.Add("Comm", "Equal to subject in age, size, type and location");
                    //fieldList.Add("OLPrice", orig_list_price);
                    //fieldList.Add("OLDate", list_date);
                    //fieldList.Add("CLPrice", current_list_price);
                    //fieldList.Add("CLDate", lastPriceChangeDate.ToShortDateString());
                    //fieldList.Add("SPrice", sold_price);
                    //fieldList.Add("SDate", closed_date);
                    //fieldList.Add("SLPrice", current_list_price);
                    //fieldList.Add("SaleType", sale_type);
                    //fieldList.Add("Source", "MLS");
                    //fieldList.Add("SourceID", mlsnum);
                    //fieldList.Add("SDistrict", "Unknown");
                    //fieldList.Add("SDivision", "n/a");
                    //fieldList.Add("MLSArea", "n/a");
                    //fieldList.Add("*Construct", "a");




                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList, this);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                }
                #endregion
                //
                //Equitrax
                //
                #region equitrax
                //form Z or form X?, completely different systems
                //form z and n, very similar
                if (currentUrl.ToLower().Contains("equi-trax"))
                {
                    bool formX = true;
                    
                    if (formX)
                    {
                        #region formx
                        {
                            Equitrax bpoform = new Equitrax();
                            Dictionary<string, string> fieldList = new Dictionary<string, string>();
                            fieldList.Add("BsmtSqFt", m.BasementGLA());
                            fieldList.Add("BsmtFinSqFt", m.BasementFinishedGLA());
                            fieldList.Add("MLS", mlsnum);

                            fieldList.Add("filepath", SubjectFilePath);
                            fieldList.Add("Address", full_street_address);
                            fieldList.Add("City", city);
                            fieldList.Add("*State", "ii");
                            fieldList.Add("Zip", zip);
                            fieldList.Add("*TR", room_count);
                            fieldList.Add("*BR", bedrooms);
                            fieldList.Add("*REO", m.DistressedSaleYesNo());
                            
                            try
                            {
                                if (Convert.ToInt32(half_bath) != 0)
                                {
                                    fieldList.Add("*BA", full_bath + full_bath);
                                }
                                else
                                {
                                    fieldList.Add("*BA", full_bath);
                                }
                            }

                            catch
                            {
                                fieldList.Add("*BA", full_bath);
                            }



                            fieldList.Add("*Type", "s");



                            fieldList.Add("*Style", m.Type.Replace(" ", "<SP>"));



                            fieldList.Add("*Loc", "r");

                            fieldList.Add("GLASqFt", mls_gla);
                            fieldList.Add("EstValSqFt", "0");
                            if (finished_basement)
                            {
                                fieldList.Add("BsmtPerFin", "100");
                            }
                            else
                            {
                                fieldList.Add("BsmtPerFin", "0");
                            }

                            fieldList.Add("EstValBsmtPerFin", "0");
                            fieldList.Add("YrBlt", year_built);
                            fieldList.Add("EstValYrBlt", "0");
                            fieldList.Add("LotSize", mls_lot_size + "ac");
                            fieldList.Add("EstValLotSize", "0");
                            fieldList.Add("DOM", dom);
                            fieldList.Add("EstValDOM", "0");
                            fieldList.Add("Other", "n/a");
                            fieldList.Add("EstValOther", "0");
                            if (parking.ToLower().Contains("gar"))
                            {
                                fieldList.Add("*Gar", mls_garage_spaces);
                            }
                            else
                            {
                                fieldList.Add("*Gar", "n");
                            }
                            fieldList.Add("EstValGar", "0");
                            fieldList.Add("*Pool", "n");


                            fieldList.Add("ExtAdd", "0");
                            fieldList.Add("EstValExt", "0");
                            fieldList.Add("*Land", "r");
                            fieldList.Add("Fin", m.FinancingMlsString);
                            fieldList.Add("EstValFin", "0");
                            fieldList.Add("*Cond", "a");
                            fieldList.Add("EstValCond", "0");
                            fieldList.Add("Comm", "Equal to subject in age, size, type and location");
                            fieldList.Add("OLPrice", orig_list_price);
                            fieldList.Add("OLDate", list_date);
                            fieldList.Add("CLPrice", current_list_price);
                            fieldList.Add("CLDate", lastPriceChangeDate.ToShortDateString());
                            fieldList.Add("SPrice", sold_price);
                            fieldList.Add("SDate", closed_date);
                            fieldList.Add("SLPrice", current_list_price);
                            fieldList.Add("SaleType", sale_type);
                            fieldList.Add("Source", "MLS");
                            fieldList.Add("SourceID", mlsnum);
                            fieldList.Add("SDistrict", "Unknown");
                            fieldList.Add("Subdivision", mls_subdivision.Replace(" ", "<SP>"));
                            fieldList.Add("MLSArea", "n/a");
                            fieldList.Add("*Construct", "a");
                            fieldList.Add("Amenities1", "n/a");
                            fieldList.Add("Amenities2", "n/a");

                        #endregion
                            //
                            //Form Z additions
                            //
                            if (bpoform.ReportType(iim2) == "z" || bpoform.ReportType(iim2) == "n")
                            {
                                #region form Z and N additions
                                fieldList.Add("LDate", list_date);
                                fieldList.Add("LaSPrice", current_list_price);
                                fieldList.Add("SqFt", mls_gla);
                                fieldList.Add("EnergyEff", "N/A");
                                fieldList.Add("Ext", "N/A");
                                fieldList.Add("*REOCorp", "No");

                                fieldList.Add("FencePool", "N/A");
                                if (SubjectAttached)
                                {
                                    fieldList.Add("Site", "0ac");
                                }
                                else
                                {
                                    fieldList.Add("Site", mls_lot_size + "ac");
                                }
                                
                                fieldList.Add("AC", "GFA.C");

                                //Bath
                                #region bathroom
                                fieldList.Remove("*BA");
                                fieldList.Add("*BA", full_bath);
                                fieldList.Add("*BAH", half_bath);

                                //string tBath = fieldList["*BA"];
                                //try
                                //{
                                //    fieldList.Add("*BAH", half_bath);
                                //    //if there is a half bath, assuming /d./d pattern, hence the try block, incase /d format 
                                //    if (Convert.ToInt16(tBath[2].ToString()) > 0)
                                //    {
                                //        tBath = tBath[0].ToString() + ".5";
                                //    }
                                //    else
                                //    {
                                //        tBath = tBath[0].ToString();
                                //    }
                                //    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropBA CONTENT=" + tBath);
                                //    fieldList.Remove("*BA");
                                //    fieldList.Add("*BA", tBath);

                                //}

                                //catch
                                //{

                                //}

                                  #endregion

                                //Design Appeal
                                #region type selection
                                fieldList.Remove("*Type");
                                fieldList.Add("*Type", bpoform.TypeString(m));
                              

                                if (type.ToLower().Contains("1 story") || type.ToLower().Contains("ranch"))
                                {
                                    fieldList.Add("*Appeal", "Single<SP>Story");
                                }
                                else if (type.ToLower().Contains("2 stories") || type.ToLower().Contains("townhome"))
                                {
                                    fieldList.Add("*Appeal", "2-Story<SP>Conv");
                                }
                                else if (type.ToLower().Contains("raised ranch"))
                                {
                                    fieldList.Add("*Appeal", "Split/Bi-Level");
                                }
                                else if (type.ToLower().Contains("split") || type.ToLower().Contains("other"))
                                {
                                    fieldList.Add("*Appeal", "Tri/Muilti-Level");
                                }
                                else if (type.ToLower().Contains("1.5"))
                                {
                                    fieldList.Add("*Appeal", "Cape");
                                }
                                else
                                {
                                    fieldList.Add("*Appeal", "2-Story<SP>Conv");
                                }
                                #endregion
                                //Garage
                                #region garage

                                fieldList.Remove("*Gar");
                                fieldList.Add("*Gar", bpoform.GarageString(m));
                                //if (SubjectParkingType.ToLower().Contains("gar"))
                                //{

                                //    string contentString = "";

                                //    if (SubjectParkingType.ToLower().Contains("att"))
                                //    {
                                //        contentString = (Regex.Match(SubjectParkingType, @"\d").Value + "<SP>Attached");
                                //    }
                                //    else if (SubjectParkingType.ToLower().Contains("det"))
                                //    {
                                //        contentString = (Regex.Match(SubjectParkingType, @"\d").Value + "<SP>Detached");
                                //    }
                                //    fieldList.Add("*Gar", contentString);

                                //}
                                //else
                                //{
                                //    fieldList.Add("*Gar", "None");
                                //}
                                #endregion

                                
                                fieldList.Add("BasementSF", m.BasementGLA());
                                fieldList.Add("Basement", m.BasementFinishedPercentage());

                            //    //Basement 
                            //    if (basement.ToLower().Contains("none"))
                            //    {
                            //        fieldList.Add("BasementSF", m.BasementGLA());
                            //        fieldList.Add("Basement", "0");
                            //    }
                            //    else
                            //    {
                            //        try
                            //        {
                            //            if (fullBasement)
                            //            {
                            //                if (type == "1 Story")
                            //                {
                            //                    fieldList.Add("BasementSF", mls_gla);
                            //                }
                            //                else if (type == "2 Stories")
                            //                {
                            //                    fieldList.Add("BasementSF", (Convert.ToInt64(mls_gla) / 2).ToString());
                            //                }
                            //                else
                            //                {
                            //                    fieldList.Add("BasementSF", (Convert.ToInt64(mls_gla) / 3).ToString());
                            //                }
                            //            }

                            //            if (partialBasement)
                            //            {
                            //                if (type == "1 Story")
                            //                {
                            //                    fieldList.Add("BasementSF", (Convert.ToInt64(mls_gla) / 2).ToString());
                            //                }
                            //                else if (type == "2 Stories")
                            //                {
                            //                    fieldList.Add("BasementSF", (Convert.ToInt64(mls_gla) / 4).ToString());
                            //                }
                            //                else
                            //                {
                            //                    fieldList.Add("BasementSF", (Convert.ToInt64(mls_gla) / 6).ToString());
                            //                }
                            //            }

                            //            if (finished_basement)
                            //            {
                            //                fieldList.Add("Basement", "100");
                            //            }
                            //            else
                            //            {
                            //                fieldList.Add("Basement", "0");
                            //            }
                            //        }
                            //        catch { }
                                //        }
                                #endregion
                            }

                            bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                            status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                            iim.iimPlayCode(macro12.ToString());
                            htmlCode = iim.iimGetLastExtract();

                            // Perform the increment on the ProgressBar.
                            pBar2.PerformStep();

                            if (subjectAttachedRadioButton.Checked)
                            {
                                m = new AttachedListing(htmlCode);
                            }
                            else if (subjectDetachedradioButton.Checked)
                            {
                                m = new DetachedListing(htmlCode);
                            }
                            else
                            {
                                m = new MLSListing(htmlCode);
                            }
                        }
                    }
                    else
                    {
                        #region not currently used
                        Equitrax bpoform = new Equitrax();
                        Dictionary<string, string> fieldList = new Dictionary<string, string>();

                        fieldList.Add("Address", full_street_address);
                        fieldList.Add("City", city);
                        //
                        //proximity button
                        //
                        //
                        //Reo Y/N dropdown
                        //
                        fieldList.Add("SPrice", sold_price);
                        fieldList.Add("SDate", closed_date);
                        fieldList.Add("OLPrice", orig_list_price);
                        fieldList.Add("LDate", list_date);
                        fieldList.Add("LaSPrice", current_list_price);
                        fieldList.Add("*BR", bedrooms);
                        try
                        {
                            if (Convert.ToInt32(half_bath) != 0)
                            {
                                fieldList.Add("*BA", full_bath + full_bath);
                            }
                            else
                            {
                                fieldList.Add("*BA", full_bath);
                            }
                        }

                        catch
                        {
                            fieldList.Add("*BA", full_bath);
                        }
                        fieldList.Add("*TR", room_count);
                        fieldList.Add("*Type", "s");
                        fieldList.Add("SqFt", mls_gla);
                        fieldList.Add("*Loc", "g");
                        fieldList.Add("YrBlt", year_built);
                        if (parking.ToLower().Contains("gar"))
                        {
                            fieldList.Add("*Gar", mls_garage_spaces);
                        }
                        else
                        {
                            fieldList.Add("*Gar", "n");
                        }
                        //
                        //Patio/Porches text field
                        //
                        //
                        //fences/pools txt field
                        //
                        fieldList.Add("Fin", "n/a");
                        fieldList.Add("*Leasehold", "f");
                        fieldList.Add("LotSize", mls_lot_size + "ac");
                        fieldList.Add("*View", "Residential");
                        //
                        //Design / Appeal dropdown
                        //
                        //
                        //bsmt gla txt box
                        //
                        //
                        //bsmt fin %
                        //
                        fieldList.Add("*QualConst", "g");
                        fieldList.Add("*FcnUtil", "y");
                        //
                        //hvac text box
                        //
                        fieldList.Add("EnergyEff", "N/A");
                        fieldList.Add("Other", "N/A");
                        fieldList.Add("Source", "MLS");
                        fieldList.Add("*Cond", "g");
                      
                        //if (finished_basement)
                        //{
                        //    fieldList.Add("BsmtPerFin", "100");
                        //}
                        //else
                        //{
                        //    fieldList.Add("BsmtPerFin", "0");
                        //}

                       
                        fieldList.Add("DOM", dom);
                      
                        
                       
                       
        
            
                       
                  
                        fieldList.Add("CLDate", lastPriceChangeDate.ToShortDateString());
                        
                       
                        fieldList.Add("CLPrice", current_list_price);
                        fieldList.Add("SaleType", sale_type);
                        
                        fieldList.Add("SourceID", mlsnum);
                        fieldList.Add("SDistrict", "Unknown");
                        fieldList.Add("SDivision", "n/a");
                        fieldList.Add("MLSArea", "n/a");
                        fieldList.Add("*Construct", "a");




                        bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                        status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                        #endregion
                    }
                }
                
                #endregion  
                //
                //lres
                //
                #region lres
                if (currentUrl.ToLower().Contains("lres"))
                {
                    Lres lres = new Lres(this);
                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    fieldList.Add("Address", full_street_address);
                    fieldList.Add("City", city);
                    fieldList.Add("Zip", zip);
                    fieldList.Add("State", "%IL");
                    fieldList.Add("OLP", orig_list_price);
                    fieldList.Add("DOMOLP", domAtOriginalListPrice);
                    fieldList.Add("ListPrice", current_list_price);
                    fieldList.Add("DOM", dom);
                    fieldList.Add("SalePrice", sold_price);
                    fieldList.Add("SaleDate", closed_date);
                    fieldList.Add("RoomCount", room_count);
                    fieldList.Add("BedrmCount", bedrooms);
                    fieldList.Add("BathrmCount", full_bath + "." + half_bath.Replace("1", "5"));
                    fieldList.Add("LivArea", mls_gla);
                    fieldList.Add("LotSize", mls_lot_size);
                    fieldList.Add("YearBuilt", year_built);
                    fieldList.Add("Style", "%Cntmp");
                    fieldList.Add("ConstructionType", "%Other");
                    if (fullBasement)
                    {
                         fieldList.Add("BasementFeature", "%Full");
                    } else   if (partialBasement)
                    {
                         fieldList.Add("BasementFeature", "%Partial");
                    } else
                    {
                        fieldList.Add("BasementFeature", "%Crawl" );
                    }

                    if (finished_basement)
                    {
                        fieldList.Add("BasementFinish", "%Fully<SP>Finished");
                    }
                    else
                    {
                        fieldList.Add("BasementFinish", "%Unfinished");
                    }

                    fieldList.Add("Location", "%Average");


                    fieldList.Add("Condition", "%Average");

                    string lresGarageStr = "None";
                    //string numSpaces = Regex.Match(form.SubjectParkingType, @"\d").Value;
                    string att_det = "";

                    if (!string.IsNullOrEmpty(mls_garage_spaces))
                    {
                        if (mls_garage_type.ToLower().Contains("att"))
                        {
                            att_det = "Attached";
                        }
                        else if (mls_garage_type.ToLower().Contains("det"))
                        {
                            att_det = "Detached";
                        }

                        switch (mls_garage_spaces)
                        {
                            case "1":
                                lresGarageStr = "1<SP>Car<SP>" + att_det;
                                break;
                            case "2":
                                lresGarageStr = "2<SP>Car<SP>" + att_det;
                                break;
                            default:
                                lresGarageStr = "2+<SP>Car<SP>" + att_det;
                                break;
                        }

                    }
                    fieldList.Add("Garage", "%" + lresGarageStr);
                    fieldList.Add("Compared", "%Equal");
                    fieldList.Add("InfoSource", "%MLS");
                    fieldList.Add("TransType", "%Fair<SP>Market");

            
                   
                    lres.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                }
                #endregion
                //
                //Emort
                //
                #region emort
                if (currentUrl.ToLower().Contains("emortgage"))
                {
                    Emortgage bpoform = new Emortgage();
                    Dictionary<string, string> fieldList = new Dictionary<string, string>();

                    //newest form
                    bool formE = true;


                    //common fields
                    fieldList.Add("Address", full_street_address);
                    fieldList.Add("City", city);
                    fieldList.Add("Zip", zip);


                    if (formE)
                    {
                        
                            fieldList.Add("filepath", SubjectFilePath);

                           
                            //fieldList.Add("*State", "ii");
                            fieldList.Add("*TR", room_count);
                            fieldList.Add("*BR", bedrooms);

                            //Baths
                            #region bathroom
                            string tBath = full_bath + "." + half_bath;
                            try
                            {
                                //if there is a half bath, assuming /d./d pattern, hence the try block, incase /d format 
                                if (Convert.ToInt16(tBath[2].ToString()) > 0)
                                {
                                    tBath = tBath[0].ToString() + ".5";
                                }
                                else
                                {
                                    tBath = tBath[0].ToString();
                                }
                                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropBA CONTENT=" + tBath);
                                fieldList.Add("*BA", tBath);
                            }
                            catch
                            {

                            }

                            #endregion

                            if (hasFireplace)
                            {
                                fieldList.Add("*Fireplace", "Yes");
                            } else
                            {
                                fieldList.Add("*Fireplace", "No");

                            }


                           // fieldList.Add("*Type", "s");



                            //fieldList.Add("*Style", "2 Story Conv");


                           // fieldList.Add("*Loc", "r");

                            fieldList.Add("SqFt", mls_gla);

                            if (fullBasement || partialBasement)
                            {
                                fieldList.Add("*Bsmt", "Yes");
                                fieldList.Add("*BsmtSqFt", (Convert.ToInt16(mls_gla)/2).ToString());

                                if (finished_basement)
                                {
                                    fieldList.Add("BsmtPerFin", "100");
                                }
                                else
                                {
                                    fieldList.Add("BsmtPerFin", "0");
                                }
                            } else
                            {
                                fieldList.Add("*Bsmt", "No"); 
                            }

                           

                     fieldList.Add("Appeal", "Average");
                            fieldList.Add("YrBlt", year_built);
                      //      fieldList.Add("EstValYrBlt", "0");
                            fieldList.Add("LotSize", mls_lot_size + "ac");
                        //    fieldList.Add("EstValLotSize", "0");
                            fieldList.Add("DOM", dom);
                        //    fieldList.Add("EstValDOM", "0");
                            fieldList.Add("Other", "n/a");
                        //    fieldList.Add("EstValOther", "0");
                            //if (parking.ToLower().Contains("gar"))
                            //{
                            //    fieldList.Add("*Gar", mls_garage_spaces);
                            //}
                            //else
                            //{
                            //    fieldList.Add("*Gar", "n");
                            //}
                         //   fieldList.Add("EstValGar", "0");
                         //   fieldList.Add("*Pool", "n");


                           // fieldList.Add("ExtAdd", "0");
                          //  fieldList.Add("EstValExt", "0");
                          //  fieldList.Add("*Land", "r");
                          //  fieldList.Add("Fin", "n/a");
                          //  fieldList.Add("EstValFin", "0");
                          //  fieldList.Add("*Cond", "a");
                          //  fieldList.Add("EstValCond", "0");
                            fieldList.Add("Comm", "Equal to subject in age, size, type and location");
                            fieldList.Add("OLPrice", orig_list_price);
                            fieldList.Add("OLDate", list_date);
                            fieldList.Add("CLPrice", current_list_price);
                            fieldList.Add("LRDate", lastPriceChangeDate.ToShortDateString());
                            fieldList.Add("SPrice", sold_price);
                            fieldList.Add("SDate", closed_date);
                            fieldList.Add("SLPrice", current_list_price);
                            fieldList.Add("SaleType", sale_type);
                            fieldList.Add("Source", "MLS");
                            fieldList.Add("SourceID", mlsnum);
                            fieldList.Add("SDistrict", "Unknown");
                            fieldList.Add("SDivision", "n/a");
                            fieldList.Add("MLSArea", "n/a");
                           // fieldList.Add("*Construct", "a");

                            //Bath
                            #region bathroom
                           // fieldList.Remove("*BA");
                            //fieldList.Add("*BA", full_bath);
                    //        fieldList.Add("*BAH", half_bath);

                            //string tBath = fieldList["*BA"];
                            //try
                            //{
                            //    fieldList.Add("*BAH", half_bath);
                            //    //if there is a half bath, assuming /d./d pattern, hence the try block, incase /d format 
                            //    if (Convert.ToInt16(tBath[2].ToString()) > 0)
                            //    {
                            //        tBath = tBath[0].ToString() + ".5";
                            //    }
                            //    else
                            //    {
                            //        tBath = tBath[0].ToString();
                            //    }
                            //    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropBA CONTENT=" + tBath);
                            //    fieldList.Remove("*BA");
                            //    fieldList.Add("*BA", tBath);

                            //}

                            //catch
                            //{

                            //}

                            #endregion

                            //Design Appeal
                            #region type selection
                           // fieldList.Remove("*Type");
                            fieldList.Add("*Type", bpoform.TypeString(m));


                            if (type.ToLower().Contains("1 story") || type.ToLower().Contains("ranch"))
                            {
                                fieldList.Add("*Style", "Single<SP>Story");
                            }
                            else if (type.ToLower().Contains("2 stories") || type.ToLower().Contains("townhome"))
                            {
                                fieldList.Add("*Style", "2-Story<SP>Conv");
                            }
                            else if (type.ToLower().Contains("raised ranch"))
                            {
                                fieldList.Add("*Style", "Split/Bi-Level");
                            }
                            else if (type.ToLower().Contains("split") || type.ToLower().Contains("other"))
                            {
                                fieldList.Add("*Style", "Tri/Multi-Level");
                            }
                            else if (type.ToLower().Contains("1.5"))
                            {
                                fieldList.Add("*Style", "Cape");
                            }
                            else
                            {
                                fieldList.Add("*Style", "2-Story<SP>Conv");
                            }
                            #endregion
                            //Garage
                            #region garage

                            //fieldList.Remove("*Gar");
                           fieldList.Add("*Gar", bpoform.GarageString(m));
                            //if (SubjectParkingType.ToLower().Contains("gar"))
                            //{

                            //    string contentString = "";

                            //    if (SubjectParkingType.ToLower().Contains("att"))
                            //    {
                            //        contentString = (Regex.Match(SubjectParkingType, @"\d").Value + "<SP>Attached");
                            //    }
                            //    else if (SubjectParkingType.ToLower().Contains("det"))
                            //    {
                            //        contentString = (Regex.Match(SubjectParkingType, @"\d").Value + "<SP>Detached");
                            //    }
                            //    fieldList.Add("*Gar", contentString);

                            //}
                            //else
                            //{
                            //    fieldList.Add("*Gar", "None");
                            //}
                            #endregion

                          

                            //fieldList.Add("BasementSF", m.BasementGLA());
                            //fieldList.Add("Basement", m.BasementFinishedPercentage());

                             

                               

                            
                    }
                    else  // original code/form
                    {
                        
                        fieldList.Add("TR", room_count);
                        fieldList.Add("BR", bedrooms);
                        fieldList.Add("BA", full_bath + "." + half_bath.Replace("1", "5"));
                        fieldList.Add("EstValBABR", "0");
                        fieldList.Add("Type", "SF Detached");
                        fieldList.Add("EstValType", "0");


                        fieldList.Add("Style", "2 Story Conv");
                        fieldList.Add("EstValStyle", "0");

                        fieldList.Add("Loc", "Good");
                        fieldList.Add("EstValLoc", "0");
                        fieldList.Add("SqFt", mls_gla);
                        fieldList.Add("EstValSqFt", "0");
                        if (finished_basement)
                        {
                            fieldList.Add("BsmtPerFin", "100");
                        }
                        else
                        {
                            fieldList.Add("BsmtPerFin", "0");
                        }

                        fieldList.Add("EstValBsmtPerFin", "0");
                        fieldList.Add("YrBlt", year_built);
                        fieldList.Add("EstValYrBlt", "0");
                        fieldList.Add("LotSize", mls_lot_size + "ac");
                        fieldList.Add("EstValLotSize", "0");
                        fieldList.Add("DOM", dom);
                        fieldList.Add("EstValDOM", "0");
                        fieldList.Add("Other", "n/a");
                        fieldList.Add("EstValOther", "0");
                        if (parking.ToLower().Contains("gar"))
                        {
                            fieldList.Add("Gar", mls_garage_spaces + " " + mls_garage_type);
                        }
                        else
                        {
                            fieldList.Add("Gar", "None");
                        }
                        fieldList.Add("EstValGar", "0");
                        fieldList.Add("Pool", "No");


                        fieldList.Add("ExtAdd", "0");
                        fieldList.Add("EstValExt", "0");
                        fieldList.Add("Land", "Good");
                        fieldList.Add("Fin", "n/a");
                        fieldList.Add("EstValFin", "0");
                        fieldList.Add("Cond", "Good");
                        fieldList.Add("EstValCond", "0");
                        fieldList.Add("Comm", "Equal to subject in age, size, type and location");
                        fieldList.Add("OLPrice", orig_list_price);
                        fieldList.Add("SPrice", sold_price);
                        fieldList.Add("SDate", closed_date);
                        fieldList.Add("SLPrice", current_list_price);
                        fieldList.Add("SaleType", sale_type);
                        fieldList.Add("Source", "MLS");
                        fieldList.Add("SourceID", mlsnum);
                        fieldList.Add("SDistrict", "Unknown");
                        fieldList.Add("SDivision", "n/a");
                        fieldList.Add("MLSArea", "n/a");

                    }


                   bpoform.CompFill(iim2,sale_or_list_flag, input_comp_name, fieldList);
                   status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                }
                #endregion  
                //
                //Mainstreet aka BPOfulfillment aka Redbell
                //
                #region mainstreet
                if (currentUrl.ToLower().Contains("bpofulfillment.com"))
                {
                    StringBuilder macro = new StringBuilder();
                    macro.AppendLine(@"SET !ERRORIGNORE YES");
                    macro.AppendLine(@"SET !TIMEOUT_STEP 0");
                    
                                     //TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=findbpo&assetid=3394712&orderid=2067826 ATTR=NAME:BPO$Comparables$txtAddress22 CONTENT=rer
                                     //TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtAddress22 CONTENT=thenewvalue2
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtAddress" + input_comp_name + " CONTENT=" + full_street_address.Replace(" ","<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtCity" + input_comp_name + " CONTENT=" + city.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtCounty" + input_comp_name + " CONTENT=" + county.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtZip" + input_comp_name + " CONTENT=" + zip);
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtProximity" + input_comp_name + " CONTENT=0.00");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtSubdivision" + input_comp_name + " CONTENT=" + mls_subdivision.Replace(" ","<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtAreaName" + input_comp_name + " CONTENT=" + mls_subdivision.Replace(" ", "<SP>"));
                 
                    
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboDataSource" + input_comp_name + " CONTENT=%305");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtCurrentListDate" + input_comp_name + "$txtMonth CONTENT=" + list_date.Substring(0,2));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtCurrentListDate" + input_comp_name + "$txtDay CONTENT=" + list_date.Substring(3, 2));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtCurrentListDate" + input_comp_name + "$txtYear CONTENT=" + list_date.Substring(6, 4));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtCurrentListPrice" + input_comp_name + " CONTENT=" + current_list_price.Replace("$", "").Replace(",",""));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtListDate" + input_comp_name + "$txtMonth CONTENT=" + list_date.Substring(0, 2));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtListDate" + input_comp_name + "$txtDay CONTENT=" + list_date.Substring(3, 2));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtListDate" + input_comp_name + "$txtYear CONTENT=" + list_date.Substring(6, 4));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtFinalListPrice" + input_comp_name + " CONTENT="+ current_list_price.Replace("$", "").Replace(",",""));

                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboConstruction" + input_comp_name + " CONTENT=%104");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboLandscaping" + input_comp_name + " CONTENT=%112");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboPropertyType" + input_comp_name + " CONTENT=%764");

                    if (closed_date != "")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtSalePrice" + input_comp_name + " CONTENT=" + sold_price.Replace("$", "").Replace(",", ""));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtDateOfSale" + input_comp_name + "$txtMonth CONTENT=" + closed_date.Substring(0, 2));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtDateOfSale" + input_comp_name + "$txtDay CONTENT=" + closed_date.Substring(3, 2));

                        //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtConcessions" + input_comp_name + " CONTENT=0.00");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtDateOfSale" + input_comp_name + "$txtYear CONTENT=" + closed_date.Substring(6, 4));
                    }
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtConcessions" + input_comp_name + " CONTENT=0.00");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtNumUnits" + input_comp_name + " CONTENT=1");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboPropertyType" + input_comp_name + " CONTENT=%142");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboPropertyStyle" + input_comp_name + " CONTENT=%0");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboPropertyStyle" + input_comp_name + " CONTENT=%0");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboCondition" + input_comp_name + " CONTENT=%348");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboView" + input_comp_name + " CONTENT=%485");
                    macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                    double x = -1;
                    string st = mls_lot_size;
                    Double.TryParse(st, out x);
                    x = Math.Round(x, 2);

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtYearBuilt" + input_comp_name + " CONTENT=" + year_built);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtLotSize" + input_comp_name + " CONTENT=" + x.ToString());

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtSite" + input_comp_name + " CONTENT=Level/100%");

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtTotalSQ" + input_comp_name + " CONTENT=" + mls_gla);
                    macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtTotalRooms" + input_comp_name + " CONTENT=" + room_count);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtBedrooms" + input_comp_name + " CONTENT=" + bedrooms);
                    macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                  //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtBedrooms2 CONTENT=3");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtFullBath" + input_comp_name + " CONTENT=" + full_bath);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtHalfBath" + input_comp_name + " CONTENT=" + half_bath);
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboBasement" + input_comp_name + " CONTENT=%330");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtBasementFinished" + input_comp_name + " CONTENT=0.00");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboGarageCarport" + input_comp_name + " CONTENT=%156");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtPoolSpaFirplace" + input_comp_name + " CONTENT=na");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtAmenities" + input_comp_name + " CONTENT=na");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboSuperior" + input_comp_name + " CONTENT=%496");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboSaleType" + input_comp_name + " CONTENT=%445");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtMLS" + input_comp_name + "");
                    macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*TxtMulSales1");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtMLS" + input_comp_name + " CONTENT=" + mlsnum);
                 //   macro.AppendLine(@"'TAG POS=0 TYPE=TEXTAREA FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*TxtMulSales1 CONTENT=EASY<SP>LIVING<SP>ALL<SP>ON<SP>ONE<SP>LEVEL!<SP>LAWN<SP>CARE<SP>&<SP>SNOW<SP>REMOVAL<SP>INCLUDED!<SP>LIVING<SP>ROOM<SP>W/<SP>VAULTED<SP>CEILING,<SP>CHANDELIER<SP>FOR<SP>DINING<SP>TABLE,<SP>CORNER<SP>GAS<SP>FIREPLACE<SP>W/CERAMIC<SP>TILE<SP>SURROUND<SP>&<SP>NICHE<SP>ABOVE<SP>FOR<SP>FLAT-SCREEN<SP>TV!<SP>9'<SP>CEILINGS!<SP>MASTER<SP>SUITE<SP>W/PRIVATE<SP>BATH<SP>&<SP>WALK-IN<SP>CLOSET.<SP>LARGE<SP>EAT-IN<SP>KITCHEN<SP>W/PANTRY<SP>&<SP>WINDOW<SP>SEAT!<SP>MUDROOM<SP>OFF<SP>GARAGE<SP>WITH<SP>LAUNDRY.<SP>2.5<SP>CAR<SP>GARAGE<SP>W/PLENTY<SP>OF<SP>STORAGE<SP>SPACE.<SP>CUSTOM<SP>BLINDS<SP>STAY!<SP>HUGE<SP>BACKYARD!<SP>");

                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboViewFactor1" + input_comp_name + "  CONTENT=%485");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*PropertyStyle" + input_comp_name + "  CONTENT=%93");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*PropertyType" + input_comp_name + "  CONTENT=%764");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*LeaseHold" + input_comp_name + "  CONTENT=%415");

                    //string macroCode = macro.ToString();
                    status = iim2.iimPlayCode(macro.ToString(), 60);

                      s = iim.iimPlayCode(macro12.ToString());
                    htmlCode = iim.iimGetLastExtract();
                    if (subjectAttachedRadioButton.Checked)
                    {
                        m = new AttachedListing(htmlCode);
                    }
                    else if (subjectDetachedradioButton.Checked)
                    {
                        m = new DetachedListing(htmlCode);
                    }
                    else
                    {
                        m = new MLSListing(htmlCode);
                    }

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    BPOFulfillment bpoform = new BPOFulfillment(m);

                    fieldList.Add("filepath", SubjectFilePath);

                

                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                

                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);


                }

                #endregion


                #region valuationPartners
                if (streetnumTextBox.Text == "vp")
                {

                    StringBuilder macro = new StringBuilder();


                    if (!input_comp_name.Contains("LIST"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "ADDRSTREET CONTENT=" + full_street_address.Replace(" ", "<SP>"));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "ADDRCITY CONTENT=" + city.Replace(" ", "<SP>"));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "ADDRZIP CONTENT=" + zip);

                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "SALEDT CONTENT=" + closed_date);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "SALESPRICE CONTENT=" + sold_price);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "SRCFUNDS CONTENT=Conventional");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "ORIGPRICE CONTENT=" + orig_list_price);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "LISTPRICE CONTENT=" + current_list_price);

                    }
                    else
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "STREET CONTENT=" + full_street_address.Replace(" ", "<SP>"));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "CITY CONTENT=" + city.Replace(" ", "<SP>"));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "ZIP CONTENT=" + zip);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "ORGPRICE CONTENT=" + orig_list_price);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "LISTINGPRICE CONTENT=" + current_list_price);
                    }

                     macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "DATA CONTENT=MLS");

                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "PROXIMITY");

                    
                   
                    
                   macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "MARKETDAYS CONTENT=" + dom);
                  
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "SCONCESSION CONTENT=0");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "DISTRESSEDSALE CONTENT=No");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "HOAASSESSMENT CONTENT=0");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "LIVINGSQFT CONTENT=" + mls_gla);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "AGEYRS CONTENT=" + age);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "LOTSIZE CONTENT=" + mls_lot_size);
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "SITE");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "ROOMTOT CONTENT=" + room_count);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "BEDROOMS CONTENT=" + bedrooms);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "BATHFULL CONTENT=" + full_bath);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "BATHHALF CONTENT=" + half_bath);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "DESIGNSTYLE CONTENT=" + type.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "VIEW CONTENT=Residential");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "VIEWCOMPARISON CONTENT=Equal");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "NUMUNITS CONTENT=1");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "BASEMENTTYPE CONTENT=" + basement.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "BASEMENT CONTENT=");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "BASEMENTFIN CONTENT=100");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "GARAGE CONTENT=Garage");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "GARAGENUMCARS CONTENT=");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "POOL CONTENT=N/A");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "CONDITION CONTENT=Good");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:ddl" + input_comp_name + "MOSTCOMPARABLE CONTENT=%Equal");

                        
                        macroCode = macro.ToString();

                    status = iim2.iimPlayCode(macroCode, 30);



                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                
                }


                #endregion

                #region m2m

                if (currentUrl.ToLower().Contains(@"ordereditwizard/"))
                {
                    #region code

                    s = iim.iimPlayCode(macro12.ToString());
                    htmlCode = iim.iimGetLastExtract();
                    if (subjectAttachedRadioButton.Checked)
                    {
                        m = new AttachedListing(htmlCode);
                    }
                    else if (subjectDetachedradioButton.Checked)
                    {
                        m = new DetachedListing(htmlCode);
                    }
                    else
                    {
                        m = new MLSListing(htmlCode);
                    }

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    M2MStandard bpoform = new M2MStandard(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion

                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                }

                #endregion

                #region eval
                if (currentUrl.ToLower().Contains("evalonline"))
                {

                    
                    
                 
                    StringBuilder macro = new StringBuilder();

                    if (input_comp_name == "COMPARABLE1")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Competitive<SP>Listing<SP>1");
                    }
                    if (input_comp_name == "COMPARABLE2")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Competitive<SP>Listing<SP>2");
                    } 
                    if (input_comp_name == "COMPARABLE3")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Competitive<SP>Listing<SP>3");
                    } 
                    
                    if (input_comp_name == "RECENT_SALE1")
                    {
                         macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Comparable<SP>Sales<SP>1");
                    }

                    if (input_comp_name == "RECENT_SALE2")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Comparable<SP>Sales<SP>2");
                    }

                    if (input_comp_name == "RECENT_SALE3")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Comparable<SP>Sales<SP>3");
                    }

                    macro.AppendLine(@"WAIT SECONDS=2");
            
                   
                   
                    //macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Competitive<SP>Listing<SP>1");
                    //macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Competitive<SP>Listing<SP>2");
                    //macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Competitive<SP>Listing<SP>3");

                    // C# snippet generated by iMacros Editor.
                    // See http://wiki.imacros.net/Web_Scripting for details on how to use the iMacros Scripting Interface.

                    // iMacros.AppClass iim = new iMacros.AppClass();
                    // iMacros.Status status = iim.iimOpen("", true, timeout);
                    //StringBuilder macro = new StringBuilder();


                    //
                    //Common Fields w/ matching var names
                    //

                    //lot size sf
                    if (mls_lot_size == "")
                        mls_lot_size = "0";

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10634 CONTENT=" + (Convert.ToDecimal(mls_lot_size) * 43560).ToString());
                    //macro.AppendLine(@"WAIT SECONDS=3");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10636 CONTENT=" + mls_lot_size);

                    //garage type and spaces
                    //garage type 1 = attached, 4 = detached
                    if (mls_garage_type.ToLower().Contains("att"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10630 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10632 CONTENT=" + mls_garage_spaces);
                    }
                    else if (mls_garage_type.ToLower().Contains("det"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10630 CONTENT=%4");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10632 CONTENT=" + mls_garage_spaces);
                    }
                    else
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10630 CONTENT=%3");
                    }


                    //att=1 or det=3
                    if (subjectAttachedRadioButton.Checked)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10604 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10610 CONTENT=%7");
                    }
                    else
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10604 CONTENT=%3");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10610 CONTENT=%1");

                    }
                    //mlstype
                    switch (type)
                    {
                        case "1 Story":
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%1");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Ranch");
                            if (fullBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + mls_gla);
                            }
                            else if (partialBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + (Convert.ToInt16(mls_gla.Replace(",", "")) / 2));
                            }
                            else
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=0");
                            }
                            break;
                        case "1.5 Story":
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%3");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Conventional");
                             if (fullBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + (Convert.ToInt16(mls_gla.Replace(",", "")) / 2));
                            }
                            else if (partialBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + (Convert.ToInt16(mls_gla.Replace(",", "")) / 4));
                            }
                            else
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=0");
                            }
                          
                            break;
                        case "2 Stories":
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%4");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Contemporary");
                            if (fullBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + (Convert.ToInt16(mls_gla.Replace(",", "")) / 2));
                            }
                            else if (partialBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + (Convert.ToInt16(mls_gla.Replace(",", "")) / 4));
                            }
                            else
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=0");
                            }
                            break;
                        case "Raised Ranch":
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%9");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Raised<SP>Ranch");
                             macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%1");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Ranch");
                            if (fullBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + mls_gla);
                            }
                            else if (partialBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + (Convert.ToInt16(mls_gla.Replace(",", "")) / 2));
                            }
                            else
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=0");
                            }
                            break;
                           
                        case "Split Level":
                        case @"Split Level w/Sub":
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%25");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Split<SP>Level");
                             macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%4");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Contemporary");
                            if (fullBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + (Convert.ToInt16(mls_gla.Replace(",", "")) / 2));
                            }
                            else if (partialBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + (Convert.ToInt16(mls_gla.Replace(",", "")) / 4));
                            }
                            else
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=0");
                            }
                            break;
                    }

                    //basement
                    if (basement.ToLower().Contains("full"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10624 CONTENT=%3");
                        if (finished_basement)
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10626 CONTENT=%5");
                        }
                        else
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10626 CONTENT=%1");
                        }
                    }
                    else
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10624 CONTENT=%2");
                    }
                    //patio, porch, deck, 3=no
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10640 CONTENT=%3");
                    //fireplace, 3=no
                    if (hasFireplace)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10642 CONTENT=%1");
                    }
                    else
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10642 CONTENT=%3");
                    }
                    
                    string content = "%Other<SP>(Explain)";
                    //exterior walls
                    if (exteriorDetails.Contains("Vinyl Siding"))
                    {
                        content = "%Vinyl<SP>Siding";
                    }
                    else if (exteriorDetails.Contains("Brick"))
                    {
                        content = "%Brick";
                    } 
                    else if (exteriorDetails.Contains("Cedar"))
                    {
                        content = "%Wood";
                    }

                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10648 CONTENT=" + content);
                    if (content == "Other<SP>(Explain)")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10822 CONTENT=" + exteriorDetails);
                    }

                    //
                    //Distance logic
                    //
                    string proximity = this.Get_Distance(subjectFullAddressTextbox.Text, address);

     
                    if (Convert.ToDouble(proximity) > 4 && Convert.ToDouble(proximity) <= 7)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%11");

                    }
                    else if (Convert.ToDouble(proximity) > 0 && Convert.ToDouble(proximity) <= .125)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%3");
                    }
                    else if (Convert.ToDouble(proximity) > .125 && Convert.ToDouble(proximity) <= .3)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%4");
                    }
                    else if (Convert.ToDouble(proximity) > .3 && Convert.ToDouble(proximity) <= .425)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%5");
                    }
                    else if (Convert.ToDouble(proximity) > .425 && Convert.ToDouble(proximity) <= .55)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%6");
                    }
                    else if (Convert.ToDouble(proximity) > .55 && Convert.ToDouble(proximity) <= .675)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%7");
                    }
                    else if (Convert.ToDouble(proximity) > .675 && Convert.ToDouble(proximity) <= .8)
                    {
                         macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%7");
                    }
                    else if (Convert.ToDouble(proximity) > .8 && Convert.ToDouble(proximity) <= 2)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%7");
                    }
                    else if (Convert.ToDouble(proximity) > 2 && Convert.ToDouble(proximity) <= 4)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%11");
                    }
                    //
                 

                    if (input_comp_name.Contains("SALE"))
                    {
                        #region salecomps

                       

                        if (input_comp_name == "RECENT_SALE2")
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10916 CONTENT=%3");
                        }
                        else
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10916 CONTENT=%1");
                        }
                        //financing
                        switch (financing)
                        {
                            case "FHA" :
                                 macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10906 CONTENT=%1");
                                break;
                            case "VA" :
                                 macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10906 CONTENT=%1");
                                break;
                            case "Conventional" :
                                 macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10906 CONTENT=%3");
                                break;
                            case "Cash" :
                                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10906 CONTENT=%5");
                                break;

                        }
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10906 CONTENT=%3");
                        //condition
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10900 CONTENT=%3");
                        
                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10918 CONTENT=Equal");

                        if (sale_type == "Normal")
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10905 CONTENT=%Arms<SP>Length");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10908 CONTENT=%11");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10910 CONTENT=%3");
                        }
                        else if (sale_type == "Short")
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10905 CONTENT=%Short<SP>Sale");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10908 CONTENT=%11");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10910 CONTENT=%1");
                        }
                        else if (sale_type == "REO")
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10905 CONTENT=%REO");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10908 CONTENT=%13");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10910 CONTENT=%1");
                        }
                     
                        

                        
                       

                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10500 CONTENT=" + full_street_address.Replace(" ", "<SP>"));

                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10504 CONTENT=" + city.Replace(" ", "<SP>"));

                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10506 CONTENT=%IL");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10508 CONTENT=" + zip);
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10510 CONTENT=%4");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10512 CONTENT=%3");

                        //prox to subject, 6=within 1 mi
                       // macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%6");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10516 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10518 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10518 CONTENT=%3");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10520 CONTENT=na");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10522 CONTENT=na");

                      
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10600 CONTENT=" + age.ToString());
                        
                        
                        //number of units = 1
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10602 CONTENT=1");


                     
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10604 CONTENT=%2");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%25");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Split<SP>Level");
                        //legal land use = residential
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10608 CONTENT=%1");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10610 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10611 CONTENT=%Framed");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10612 CONTENT=" + mls_gla);
                        //GLA source 1=tax records
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10614 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10616 CONTENT=" + room_count);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10618 CONTENT=" + bedrooms);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10620 CONTENT=" + full_bath);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10622 CONTENT=" + half_bath);
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10624 CONTENT=%3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10626 CONTENT=%5");
                       
                        
                        
                        // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10634");
                        // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10636 CONTENT=.2");
                        // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:formInput ATTR=ID:submit&&VALUE:SAVE<SP>&<SP>VALIDATE");
                        macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10636 CONTENT=.2");

                        //pool = no
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10638 CONTENT=%3");

                        // macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10640 CONTENT=%3");
                       
                        
                        //guest house = no
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10644 CONTENT=%3");
                        
                        
                        //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10646");
                     
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10700");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10702");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10704");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10706");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10708 CONTENT=" + orig_list_price);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10710 CONTENT=" + sold_price);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10712 CONTENT=" + closed_date);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10714 CONTENT=" + dom);
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10716");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10717 CONTENT=%1");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10718");
                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10720 CONTENT=See<SP>subject<SP>comments.");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10722 CONTENT=" + subjectRentTextbox.Text);
                        
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10902 CONTENT=$--Choose<SP>Option--");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10904");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10905 CONTENT=%Arms<SP>Length");
                        //macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=ID:formInput ATTR=TXT:--Choose<SP>Option--FHA<SP>VA<SP>Conventional<SP>Owner-financed<SP>Cash<SP>Other<SP>-Explain<SP>in<SP>comm*");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10906 CONTENT=%1");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10908 CONTENT=%11");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10910 CONTENT=%3");
                        
                        //compare to subject = equal
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10914 CONTENT=%2");


                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10916 CONTENT=%1");
                     
                        
                        //view = residential
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10921 CONTENT=%10");

                        //view effect = equal
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10922 CONTENT=%3");
                        //macro.AppendLine(@"TAG POS=19 TYPE=TD FORM=ID:formInput ATTR=CLASS:formLabelAlignRight");
                        //personally inspect property; 3 = no
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10926 CONTENT=%3");
                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10918 CONTENT=Equal");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:formInput ATTR=ID:submit&&VALUE:SAVE<SP>&<SP>VALIDATE");
                        macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:InputForm");
                        macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:output");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:formInput ATTR=ID:submit&&VALUE:SAVE<SP>&<SP>VALIDATE");
                        ////macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:InputForm");
                        ////macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:output");
                        ////string macroCode = macro.ToString();
                        ////status = iim.iimPlayCode(macroCode, timeout);
                        #endregion
                    }
                    else
                    {

                        #region listcomps
                        if (input_comp_name == "COMPARABLE2")
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10816 CONTENT=%3");
                        }
                        else
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10816 CONTENT=%1");
                        }
                        if (additionalSalesInfo.Contains("Short"))
                        {
                            //  macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10805 CONTENT=%Short<SP>Sale");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10808 CONTENT=%11");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10810 CONTENT=%1");
                           
                        }
                        else if (additionalSalesInfo.Contains("REO"))
                        {
                           // macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10805 CONTENT=%REO");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10808 CONTENT=%13");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10810 CONTENT=%1");
                        }
                        else
                        {
                             //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10805 CONTENT=%Arms<SP>Length");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10808 CONTENT=%11");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10810 CONTENT=%3");
                          
                        }

                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10500 CONTENT=" + full_street_address.Replace(" ", "<SP>"));

                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10504 CONTENT=" + city.Replace(" ", "<SP>"));

                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10506 CONTENT=%IL");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10508 CONTENT=" + zip);
                        
                        //pud = no
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10510 CONTENT=%4");
                        
                        //corner = no
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10512 CONTENT=%3");


                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%1");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%5");

                        //location = same 
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10516 CONTENT=%1");


                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10518 CONTENT=%3");


                        //dev name  - sub name
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10520 CONTENT=na");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10522 CONTENT=na");


                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10600 CONTENT=" + age.ToString());
                        //units = 1
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10602 CONTENT=1");
                        
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10604 CONTENT=%3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%27");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%29");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%26");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%25");
                        ////macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Split<SP>Level");

                        //legal land use = res
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10608 CONTENT=%1");
                        
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10610 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10611 CONTENT=%Framed");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10612 CONTENT=" + mls_gla);
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10614 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10616 CONTENT=" + room_count);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10618 CONTENT=" + bedrooms);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10620 CONTENT=" + full_bath);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10622 CONTENT=" + half_bath);
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10624 CONTENT=%3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10626 CONTENT=%5");

                        //basement sf
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=");


                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10630 CONTENT=%1");
                        //macro.AppendLine(@"WAIT SECONDS=3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10632 CONTENT=2");
                        //macro.AppendLine(@"WAIT SECONDS=3");
                        
                       

                        //macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                        //macro.AppendLine(@"WAIT SECONDS=3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=B FORM=ID:formInput ATTR=TXT:Convert<SP>to<SP>Sqft");
                        //macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                        //macro.AppendLine(@"ONDIALOG POS=2 BUTTON=YES");
                        //macro.AppendLine(@"ONDIALOG POS=3 BUTTON=YES");
                        //macro.AppendLine(@"ONDIALOG POS=4 BUTTON=YES");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10636 CONTENT=.2");
                        //macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                        //macro.AppendLine(@"ONDIALOG POS=2 BUTTON=YES");
                        //macro.AppendLine(@"ONDIALOG POS=3 BUTTON=YES");
                        //macro.AppendLine(@"ONDIALOG POS=4 BUTTON=YES");
                        
                        //pool = no
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10638 CONTENT=%3");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10614 CONTENT=%8");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10642 CONTENT=%1");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10642 CONTENT=%3");
                        //guest house = no
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10644 CONTENT=%3");

                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10518 CONTENT=%3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%7");

                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10646");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10648 CONTENT=%Brick");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10648 CONTENT=%Vinyl<SP>Siding");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10750");
                        //agent name
                      
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10752 CONTENT=" + listing_Agent.Replace(" ","<SP>"));

                        //agency address
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10754 CONTENT=" + broker.Replace(" ", "<SP>"));


                        //agency phone
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10756 CONTENT=" + phone.Replace(" ", "<SP>"));
                        
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10758 CONTENT=" + list_date);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10760 CONTENT=" + current_list_price);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10762 CONTENT=" + list_date);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10764 CONTENT=" + current_list_price);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10766 CONTENT=" + dom);
                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10768 CONTENT=" + mlsnum);
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10769 CONTENT=%0");
                        //offered conssessions
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10770 CONTENT=0");
                        //rent different field number for sale comps
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10772 CONTENT=" + subjectRentTextbox.Text);

                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10767 CONTENT=See<SP>subject<SP>comments.");

                        //condition = avg
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10800 CONTENT=%3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10802 CONTENT=$--Choose<SP>Option--");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10804");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10805 CONTENT=$--Choose<SP>Option--");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10806 CONTENT=$--Choose<SP>Option--");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10808 CONTENT=%11");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10810 CONTENT=%3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10812");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10810 CONTENT=%1");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10812");
                        //compare to subject = equal
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10814 CONTENT=%2");

                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10816 CONTENT=$--Choose<SP>Option--");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10808 CONTENT=%13");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10810 CONTENT=%1");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10812");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10814 CONTENT=%2");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10816 CONTENT=%3");
                        
                        //personally inspect = no
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10818 CONTENT=%3");
                        
                        //lsiting view = res
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10823 CONTENT=%10");

                        //view affect, same
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10824 CONTENT=%3");
                       
                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10820 CONTENT=Equal");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:formInput ATTR=ID:submit&&VALUE:SAVE<SP>&<SP>VALIDATE");
                        macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:InputForm");
                        macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:output");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:formInput ATTR=ID:submit&&VALUE:SAVE<SP>&<SP>VALIDATE");
                        macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:InputForm");
                        macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:output");

                        #endregion

                    }
             
                    
                     macroCode = macro.ToString();
                    status = iim2.iimPlayCode(macroCode, 30);
                   


                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                    //MessageBox.Show("Comp:" + input_comp_name + " finished.");

                }
#endregion 

                #region fill usres
                if (currentUrl.ToLower().Contains("usres"))
                {

                   

                    StringBuilder macro = new StringBuilder();
                   
                    if (input_comp_name.Contains("L"))
                    {
                        
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:IMAGE FORM=NAME:InputForm ATTR=NAME:btnSave");
                        //macro.AppendLine(@"WAIT SECONDS=5");
                        macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:InputForm ATTR=SRC:http://valuations.usres.com/BpoImages/bpo_pg2_btn.gif");
                        //macro.AppendLine(@"WAIT SECONDS=5");
                       


                    }


                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abAddr" + input_comp_name + " CONTENT=" + full_street_address.Replace(" ", "<SP>"));

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abCity" + input_comp_name + " CONTENT=" + city.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abState" + input_comp_name + " CONTENT=%IL");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abZip" + input_comp_name + " CONTENT=" + zip);

                    //
                    //Distance logic
                    //
                    string proximity = this.Get_Distance(subjectFullAddressTextbox.Text, address);

                    if (Convert.ToDouble(proximity) > 4 && Convert.ToDouble(proximity) <= 7)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%2Mile");
                    }
                    else if (Convert.ToDouble(proximity) > 0 && Convert.ToDouble(proximity) <= .125)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%1Blk");
                    }
                    else if (Convert.ToDouble(proximity) > .125 && Convert.ToDouble(proximity) <= .3)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%2Blk");
                    }
                    else if (Convert.ToDouble(proximity) > .3 && Convert.ToDouble(proximity) <= .425)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%3Blk");
                    }
                    else if (Convert.ToDouble(proximity) > .425 && Convert.ToDouble(proximity) <= .55)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%4Blk");
                    }
                    else if (Convert.ToDouble(proximity) > .55 && Convert.ToDouble(proximity) <= .675)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%5Blk");
                    }
                    else if (Convert.ToDouble(proximity) > .675 && Convert.ToDouble(proximity) <= .8)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%6Blk");
                    }
                    else if (Convert.ToDouble(proximity) > .8 && Convert.ToDouble(proximity) <= .9125)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%7Blk");
                    }
                    else if (Convert.ToDouble(proximity) > .9125 && Convert.ToDouble(proximity) <= 4)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%1Mile");
                    }
                    //
                
                    if (input_comp_name.Contains("S"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSp" + input_comp_name + " CONTENT=" + sold_price.Replace("$", "").Replace(",", ""));
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abStype1 CONTENT=%t");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSaleDt" + input_comp_name + " CONTENT=" + closed_date);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abLp" + input_comp_name + " CONTENT=" + current_list_price.Replace("$", "").Replace(",", ""));
                    }

                    if (input_comp_name.Contains("L"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abLpOrig" + input_comp_name + " CONTENT=" + orig_list_price.Replace("$", "").Replace(",", ""));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSp" + input_comp_name + " CONTENT=" + current_list_price.Replace("$", "").Replace(",", ""));
                    }



                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abLpRed" + input_comp_name + " CONTENT=" + count);
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abSrce" + input_comp_name + " CONTENT=%MLS");

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abDom" + input_comp_name + " CONTENT=" + dom);
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abLoc" + input_comp_name + " CONTENT=%Suburban");

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abRight" + input_comp_name + " CONTENT=Fee<SP>Simple");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSite" + input_comp_name + " CONTENT=" + mls_lot_size);

                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abUnits" + input_comp_name + " CONTENT=%1");

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abView" + input_comp_name + " CONTENT=Neighborhood");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abDesign" + input_comp_name + " CONTENT=%Avg");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abAge" + input_comp_name + " CONTENT=" + year_built);
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abCond" + input_comp_name + " CONTENT=%Avg");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abRooms" + input_comp_name + " CONTENT=" + room_count);

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBeds" + input_comp_name + " CONTENT=" + bedrooms);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBaths" + input_comp_name + " CONTENT=" + full_bath + "." + half_bath);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abGla" + input_comp_name + " CONTENT=" + mls_gla);

                    //Basement and finished rooms below grade
                    if (finished_basement)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBsmt" + input_comp_name + " CONTENT=" + basement.Replace(" ", "<SP>") + "<SP>/<SP>Yes");
                    }
                    else
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBsmt" + input_comp_name + " CONTENT=" + basement.Replace(" ", "<SP>") + "<SP>/<SP>No");
                    }

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abHeat" + input_comp_name + " CONTENT=Gas/Central");


                    if (mls_garage_type.ToLower().Contains("att"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abGarage" + input_comp_name + " CONTENT=%" + mls_garage_spaces +"CA");
                    }
                    else if (mls_garage_type.ToLower().Contains("det"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abGarage" + input_comp_name + " CONTENT=%" + mls_garage_spaces + "CD");
                    }
                    else
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abGarage CONTENT=%N");
                    }
                    
                    
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:IMAGE FORM=NAME:InputForm ATTR=NAME:btnSave");
                    macro.AppendLine(@"WAIT SECONDS=5");

                    if (input_comp_name.Contains("L"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:IMAGE FORM=NAME:InputForm ATTR=NAME:btnSave");
                        macro.AppendLine(@"TAG POS=2 TYPE=IMG FORM=NAME:InputForm ATTR=SRC:http://valuations.usres.com/BpoImages/bpo_pg1_btn.gif");
                       


                    }
                   
                    macroCode = macro.ToString();
                    status = iim2.iimPlayCode(macroCode, 60);


                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                    //MessageBox.Show(input_comp_name);

                    //break;

                }

                #endregion

                StringBuilder macro3 = new StringBuilder();
                string macroCode3 = "";
                #region fill NVS

                if (streetnumTextBox.Text == "nvs")
                {
                    
                    macro3.AppendLine(@"SET !ERRORIGNORE YES");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "DistToSubj CONTENT=" + Get_Distance(subjectFullAddressTextbox.Text, address ));
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    //
                    //TAG POS=1 TYPE=TD FORM=NAME:bpoform ATTR=TXT:?<SP>Urban<SP>Suburban<SP>Rural
                    // 3=rural
                    //

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Location_id CONTENT=%2");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Condition CONTENT=%Average");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Rooms CONTENT=" + room_count);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Beds CONTENT=" + bedrooms);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "FullBaths CONTENT=" + full_bath);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "HalfBaths CONTENT=" + half_bath);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    if (mls_gla == "")
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "SqFeet CONTENT=" + subjectAboveGlaTextbox.Text.Replace(",", ""));
                    }
                    else
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "SqFeet CONTENT=" + mls_gla.Replace(",", ""));
                    }
                    
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "LotSize CONTENT=" + mls_lot_size);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    //
                    //specific logic to form
                    //
                    if (year_built == "")
                    {
                        year_built = "1900";
                    }
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "YearBuilt CONTENT=" + year_built);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "ListPrice CONTENT=" + current_list_price.Replace("$", "").Replace(",", ""));
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "DOM CONTENT=" + dom);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    //
                    //need select list logic
                    if (type.ToLower().Contains("1 story"))
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Style CONTENT=%Ranch");
                    }
                    else
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Style CONTENT=%Contemporary");
                    }
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    //
                    //

                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Construction_type CONTENT=%1");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

                     if (number_of_firplaces == "" | number_of_firplaces == "0")
                     {
                          macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Fireplace CONTENT=%0");
                     }
                     else
                     {
                           macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Fireplace CONTENT=%1");
                     }

                  
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Pool CONTENT=%N");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Carport CONTENT=%0");

                    if (parking.Contains("Gar"))
                    {
                          if (mls_garage_type.Contains("Att"))
                          {
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Garage CONTENT=%A" + mls_garage_spaces);
                          }
                          else
                          {
                              macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Garage CONTENT=%D" + mls_garage_spaces);
                          }
                     }
                      else
                      {
                         macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Garage CONTENT=%N");
                      }

                    //macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Garage CONTENT=%A2");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Porch CONTENT=%N");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Patio CONTENT=%N");
                    

                    //
                    //Addition data for sales
                    //
                    if (input_comp_name.Contains("Sale"))
                    {

                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Landscaping CONTENT=%Average");
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "SaleDate CONTENT=" + closed_date);
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "SalePrice CONTENT=" + sold_price.Replace("$", "").Replace(",", ""));
                        //macro3.AppendLine(@"");

                    }


                    //crawl
                    if (subjectBasementDetailsTextbox.Text.Contains("Crawl"))
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%C");
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Basement CONTENT=%C");

                    }

                    //none
                    if (subjectBasementTypeTextbox.Text == "None" | subjectBasementTypeTextbox.Text == "")   //.Text.Contains("None"))
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%N");
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Basement CONTENT=%N");
                    }


                    if (subjectBasementTypeTextbox.Text.Contains("Full"))
                    {
                        if (subjectBasementDetailsTextbox.Text.Contains("Finished"))
                        {
                            //Full finished                    
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%F");
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Basement CONTENT=%F");
                        }
                        else
                        {
                            //full unfinished  
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%Y");
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Basement CONTENT=%Y");
                        }
                    }



                    if (subjectBasementTypeTextbox.Text.Contains("Partial"))
                    {
                        if (subjectBasementDetailsTextbox.Text.Contains("Finished"))
                        {
                            //partial finish
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%P");
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Basement CONTENT=%P");
                        }
                        else
                        {
                            //partial unfinished
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%X");
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Basement CONTENT=%X");
                        }
                    }



                    macro3.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Explain CONTENT=Same<SP>size<SP>GLA<SP>and<SP>lot,<SP>same<SP>area<SP>type.");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:bpoform ATTR=VALUE:Go<SP>to<SP>Next<SP>Section<SP>-<SP>Save");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:bpoform ATTR=VALUE:Go<SP>to<SP>Next<SP>Section<SP>-<SP>Save");
                    macroCode3 = macro3.ToString();
                    status = iim2.iimPlayCode(macroCode3, 60);

                    //MessageBox.Show(status.ToString());







                    //
                    //TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:bpoform ATTR=VALUE:Go<SP>to<SP>Next<SP>Section<SP>-<SP>Save
                    //TAG POS=1 TYPE=TEXTAREA FORM=NAME:bpoform ATTR=NAME:CList1Explain CONTENT=Same<SP>size<SP>GLA<SP>and<SP>lot,<SP>same<SP>area<SP>type.
                    //ONDIALOG POS=1 BUTTON=NO

                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                    //return;

                }

                #endregion

                #region fill sls
                if (streetnumTextBox.Text == "sls")
                {


                    StringBuilder macro = new StringBuilder();
                    //macro.AppendLine(@"VERSION BUILD=8021952");
                    //macro.AppendLine(@"TAB T=1");
                    //macro.AppendLine(@"TAB CLOSEALLOTHERS");
                    //macro.AppendLine(@"URL GOTO=https://portal.nreis.com/VendorWorkflow/VendorNewForm.aspx?OrderID=1684158&SelectionID=1873443&MCID=17&ContactID=74134&SkipPipeline=Y&VendorNewForm=Y");

                    macro.AppendLine(@"SET !ERRORIGNORE YES");
                    macro.AppendLine(@"TAG POS=3 TYPE=TD FORM=ID:form1 ATTR=CLASS:BpoGrid");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompAdr" + input_comp_name + " CONTENT=" + full_street_address.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompCty" + input_comp_name + " CONTENT=" + city.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlCompSta" + input_comp_name + " CONTENT=%IL");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompZip" + input_comp_name + " CONTENT=" + zip);

                    //wait for proximity to load
                    macro.AppendLine(@"WAIT SECONDS=3");


                    // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompProximity1");
                    macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompZip1 CONTENT=60098");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompProximity1");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompSalePrice2 CONTENT=565656");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompOrigPrice" + input_comp_name + " CONTENT=" + orig_list_price.Replace("$", "").Replace(",", ""));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompSalePrice" + input_comp_name + " CONTENT=" + sold_price.Replace("$", "").Replace(",", ""));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompListPrice" + input_comp_name + " CONTENT=" + current_list_price.Replace("$", "").Replace(",", ""));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompDataSrc" + input_comp_name + " CONTENT=MLS");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompSaleDate" + input_comp_name + " CONTENT=" + closed_date);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompListDate" + input_comp_name + " CONTENT=" + list_date);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompConcessions" + input_comp_name + " CONTENT=NA");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlCompLoc" + input_comp_name + " CONTENT=%2");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlCompLeaseFee" + input_comp_name + " CONTENT=%1");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompSize" + input_comp_name + " CONTENT=" + mls_lot_size.Replace(",", ""));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompView" + input_comp_name + " CONTENT=Residential");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompDesign" + input_comp_name + " CONTENT=" + type.Replace(" ", "<SP>") + "/Avg");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlCompQuality" + input_comp_name + " CONTENT=%3");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompAge" + input_comp_name + " CONTENT=" + age.ToString());
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlCompCond" + input_comp_name + " CONTENT=%3");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompRmCountTotal" + input_comp_name + " CONTENT=" + room_count);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompRmCountBdms" + input_comp_name + " CONTENT=" + bedrooms);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompRmCountBaths" + input_comp_name + " CONTENT=" + full_bath);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompRmCountHBath" + input_comp_name + " CONTENT=" + half_bath);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompLivArea" + input_comp_name + " CONTENT=" + mls_gla.Replace(",", ""));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompBaseRms" + input_comp_name + "");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompPercFin" + input_comp_name + "");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlCompUtility" + input_comp_name + " CONTENT=%3");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompHeatAC" + input_comp_name + " CONTENT=GASFA/Central");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompEnergy" + input_comp_name + " CONTENT=" + "NA");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompGarage" + input_comp_name + " CONTENT=" + mls_garage_spaces + mls_garage_type );
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompDeckEtc" + input_comp_name + " CONTENT=" + "NA");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompPoolEtc" + input_comp_name + " CONTENT=" + "NA");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompOther" + input_comp_name + " CONTENT=" + "NA");


                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompConcessionsAdj1");
                    macroCode = macro.ToString();
                    status = iim2.iimPlayCode(macroCode, 60);

                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                    //return;

                }
                #endregion
               
                #region imort dialog control
                if (currentUrl.ToLower().Contains("propertysmart"))
                {
                    macro3.AppendLine(@"SET !ERRORIGNORE YES");
                    macro3.AppendLine(@"SET !TIMEOUT_STEP 0");
                    macro3.AppendLine(@"SET !ERRORIGNORE YES");
                    macro3.AppendLine(@"FRAME NAME=_MAIN");
                    macro3.AppendLine(@"ONWEBPAGEDIALOG KEYS={CLOSE}");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    #endregion

                    #region comp input
                    //macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/APN CONTENT=" + subjectpin_textbox.Text);
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/School_District CONTENT=school");

                   // macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:Form1 ATTR=NAME:PS_FORM/RECENT_SALE1/Datasource CONTENT=%MLS");

                 
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Street_Address1_Number CONTENT=" + street_number);
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Street_Address1_Text CONTENT=" + street_name.Replace(" ", "<SP>"));
                    //
                    //postfix fixing
                    //
                    if (street_postfix.ToUpper() == "TER")
                    {
                        street_postfix = "TERRACE";
                    }

                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Street_Address1_Ext CONTENT=%" + street_postfix);
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/City CONTENT=" + city.Replace(" ", "<SP>"));
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/State CONTENT=IL");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/State CONTENT=%IL");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Zip_Code CONTENT=" + zip);

                    macro3.AppendLine(@"WAIT SECONDS=1");
                    macro3.AppendLine(@"DS cmd=KEY X={{!TAGX}} Y={{!TAGY}}  CONTENT=1{BACKSPACE}");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Unit_Number");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");


                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");

                    //fnma
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Miles CONTENT=" + this.Get_Distance(subjectFullAddressTextbox.Text, address) );
                    //
                    //fnma
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Location CONTENT=Suburban");
                    //
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Data_Source CONTENT=MLS");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Location CONTENT=Suburban");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Concessions  CONTENT=None");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Financing_Concessions CONTENT=None");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Utility CONTENT=Typical");



                    if (SubjectDetached)
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Property_Type CONTENT=%SFR<SP>Detached");
                    }


                    if (string.IsNullOrWhiteSpace(mls_subdivision))
                    {
                         macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Sub_Neighborhood CONTENT=Unk");
                    }
                    else 
                    {
                         macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Sub_Neighborhood CONTENT=" + mls_subdivision.Replace(" ", "<SP>"));
                    }
                   




                    if (input_comp_name.Contains("SALE"))
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Num_Reductions CONTENT=" + count);
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/List_At_Sale CONTENT=" + current_list_price.Replace("$", "").Replace(",", ""));
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");






                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Sale_Date CONTENT=" + closed_date);
                        //fnma
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Sales_Date CONTENT=" + closed_date);
                        //
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");






                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        //fnma
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Sales_Price CONTENT=" + sold_price.Replace("$", "").Replace(",", ""));
                      
                        //
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Sale_Price CONTENT=" + sold_price.Replace("$", "").Replace(",", ""));
                        
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");






                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    }
                    else
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Datasource CONTENT=MLS");
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Price_Reductions CONTENT=" + count);
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/List_Price CONTENT=" + current_list_price.Replace("$", "").Replace(",", ""));

                       

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

                    }

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Days_On_Market CONTENT=" + dom);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");


                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                  
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Location CONTENT=%Suburban");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");


                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    //fnma
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Leasehold CONTENT=Fee<SP>Simple");
                    //
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Fee_Simple CONTENT=Fee<SP>Simple");
                      macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Fee_Simple CONTENT=%Fee<SP>Simple");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");


                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    //fnam
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Site CONTENT=" + mls_lot_size);
                    //
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Lot_Size CONTENT=" + mls_lot_size);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");


                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/View CONTENT=Residential");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/View CONTENT=%Residential");
                    //fnma
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Design CONTENT=" + type.Replace(" ", "<SP>") + "/Average");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Construction CONTENT=Average");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Condition CONTENT=Average");
                    //
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Style_Design CONTENT=%Average");

                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Year_Built CONTENT=" + year_built);
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Condition CONTENT=%Average");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");


                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");


                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");






                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Total_Rooms CONTENT=" + room_count);
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Bedrooms CONTENT=" + bedrooms);
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Bathrooms CONTENT=" + full_bath);
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Half_Bathrooms CONTENT=" + half_bath);

                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");


                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Living_Square_Feet CONTENT=" + mls_gla.Replace(",", ""));

                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Heating_Cooling CONTENT=Gas<SP>FA/Central");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Heating_Cooling CONTENT=%FWA");

                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Basement CONTENT=%Yes");

                    if (basement.Contains("None"))
                    {
                        

                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Basement CONTENT=%No");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Basement_Finished CONTENT=%No");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    }

                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Fireplace CONTENT=%Yes");
                    if (number_of_firplaces == "")
                    {
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Fireplace CONTENT=%No");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    }

                    if (parking != "Garage")
                    {

                        //TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/RECENT_SALE1/Garage_Type CONTENT=%Attached
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Garage CONTENT=%NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");

                    }

                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Garage_Carport CONTENT=" + subjectParkingTypeTextbox.Text.Replace(" ", "<SP>"));
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Design_Style CONTENT=" + type.Replace(" ", "<SP>"));

                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Fireplace_Adj EXTRACT=TXT");
                    macro3.AppendLine(@"DS cmd=CLICK X={{!TAGX}} Y={{!TAGY}}  ");


                    //ONDIALOG POS=1 BUTTON=NO
                    //TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/RECENT_SALE1/Design_Style CONTENT=2<SP>Story
                    //TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/RECENT_SALE1/Design_Style CONTENT=2<SP>Story<SP>Split
                    //ONDIALOG POS=1 BUTTON=NO


                    //TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/RECENT_SALE1/Basement_Adj
                    //TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/RECENT_SALE1/Basement_Finished_Adj














                    //TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/RECENT_SALE1/Street_Address1_Number CONTENT=1508

                    //macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSpS1 CONTENT=4321");


                    macro3.AppendLine(@"FRAME NAME=_MAIN");
                  // macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:Form1 ATTR=NAME:PS_FORM/RECENT_SALE1/Datasource CONTENT=%MLS");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Orig_List_Date CONTENT=" + list_date);
                   
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Orig_List_Price CONTENT=" + orig_list_price.Replace("$", "").Replace(",", ""));
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Concessions_Type CONTENT=%None");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Land_Value CONTENT=5000");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Datasource CONTENT=%MLS");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Fee_Simple CONTENT=%Fee<SP>Simple");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Num_Units CONTENT=1");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/View CONTENT=%Residential");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/View_Comparison CONTENT=%Equal");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Other CONTENT=%Average");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Overall_Comp CONTENT=%Equal");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Heating_Cooling CONTENT=%FWA");


                    macroCode3 = macro3.ToString();
                    status = iim2.iimPlayCode(macroCode3, 60);

                    #endregion

                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);


                }

                s = iim.iimPlayCode(macro12.ToString());
                 htmlCode = iim.iimGetLastExtract();
                 header = iim.iimGetExtract(0);

            if (subjectAttachedRadioButton.Checked)
            {
                m = new AttachedListing(htmlCode);
            }
            else if (subjectDetachedradioButton.Checked)
            {
                m = new DetachedListing(htmlCode);
            }
            else
            {
                m = new MLSListing(htmlCode);
            }
              
                #endregion
            }

            //
            //post processing
            //

            #region equi-trax post processing
            if (streetnumTextBox.Text == "equi-trax")
            {
                StringBuilder macro = new StringBuilder();
                macro.AppendLine(@"");
                macro.AppendLine(@"FRAME NAME=main");
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:b_sv");
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:dmsg_close_btn");

                macro.AppendLine(@"");
                macro.AppendLine(@"FRAME NAME=main");
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:b_ap");
                macro.AppendLine(@"FRAME NAME=iFileMan");
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=CLASS:s_button");
                macro.AppendLine(@"FRAME NAME=main");
                macro.AppendLine(@"TAG POS=1 TYPE=IFRAME ATTR=ID:iFileMan");
                macro.AppendLine(@"FRAME NAME=iFileMan");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f6 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S1.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f7 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S2.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f8 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S3.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f9 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L1.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f10 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L2.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f11 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L3.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:HMSG");
                macro.AppendLine(@"TAG POS=18 TYPE=A ATTR=CLASS:s_button");
                macro.AppendLine(@"TAG POS=1 TYPE=NOBR ATTR=TXT:Save<SP>Changes");
                macro.AppendLine(@"FRAME NAME=main");
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:b_ManageFiles_close");
                macro.AppendLine(@"TAG POS=11 TYPE=A ATTR=CLASS:s_button");



                string macroCode = macro.ToString();
                status = iim2.iimPlayCode(macroCode, 30);

            }
            #endregion

            if (streetnumTextBox.Text == "usres")
            {
                //
                //load comp pics
                //
                StringBuilder macro = new StringBuilder();
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:InputForm ATTR=CLASS:auditlink");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:/BPOAGENTUPLOADPROC:* ATTR=NAME:bpo_*_65 CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") +  "\\L1.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:/BPOAGENTUPLOADPROC:* ATTR=NAME:bpo_*_70 CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") +  "\\L2.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:/BPOAGENTUPLOADPROC:* ATTR=NAME:bpo_*_75 CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") +  "\\L3.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:/BPOAGENTUPLOADPROC:* ATTR=NAME:bpo_*_85 CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") +  "\\S1.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:/BPOAGENTUPLOADPROC:* ATTR=NAME:bpo_*_90 CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") +  "\\S2.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:/BPOAGENTUPLOADPROC:* ATTR=NAME:bpo_*_95 CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") +  "\\S3.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:/BPOAGENTUPLOADPROC:* ATTR=VALUE:Upload<SP>Pictures");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=ACTION:/BPOAGENTUPLOADPROC:* ATTR=TXT:Assigned");
                string macroCode = macro.ToString();
                 status = iim2.iimPlayCode(macroCode, 60);

            }


            //if (streetnumTextBox.Text == "")
            //{
            //    StringBuilder macro = new StringBuilder();
               
            //    macro.AppendLine(@"FRAME NAME=_MAIN");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:Form1 ATTR=ID:fname_Listing1_Front CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") + "\\L1.jpg");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:Form1 ATTR=ID:fname_Listing2_Front CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") + "\\L2.jpg");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:Form1 ATTR=ID:fname_Listing3_Front CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") + "\\L3.jpg");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:Form1 ATTR=ID:fname_Sale1_Front CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") + "\\S1.jpg");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:Form1 ATTR=ID:fname_Sale2_Front CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") + "\\S2.jpg");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:Form1 ATTR=ID:fname_Sale3_Front CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") + "\\S3.jpg");
            //    macro.AppendLine(@"FRAME NAME=_TOP_MENU");
            //    macro.AppendLine(@"TAG POS=1 TYPE=B FORM=ID:Form1 ATTR=TXT:SAVE");

            //    string macroCode = macro.ToString();
            //    status = iim2.iimPlayCode(macroCode, 30);
            //}
        }

        private void subjectBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.subjectBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this._BPO_SandboxDataSet);

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the '_BPO_SandboxDataSet.RawSFData' table. You can move, or remove it, as needed.
            this.rawSFDataTableAdapter.Fill(this._BPO_SandboxDataSet.RawSFData);
            
            // TODO: This line of code loads data into the '_BPO_SandboxDataSet.subject' table. You can move, or remove it, as needed.
            this.subjectTableAdapter.Fill(this._BPO_SandboxDataSet.subject);
            if (!string.IsNullOrEmpty(search_address_textbox.Text))
            {
                this.importSubjectInfoButton(this, new EventArgs());
            }
            else
            {
                // search_address_textbox
                //everyones dropbox folder is different, hence this helper function
                //DropBoxFolder

                string baseDir = DropBoxFolder + @"\BPOs\";
                string[] dirs = Directory.GetDirectories(baseDir);
                string mostRecentAccess;
                TimeSpan ts = new TimeSpan();

                foreach (string d in dirs)
                {
                       ts = Directory.GetLastAccessTime(d) - DateTime.Now;
                     if (ts.Days > -5)
                     {
                          comboBox3.Items.Add(d);
                     }

                }
                TypeDetachedList tdl = new TypeDetachedList();

                foreach (string key in tdl.mlsTypeDetached.Keys)
                {
                    comboBox4.Items.Add(tdl.mlsTypeDetached[key]);
                }

                TypeAttachedList tal = new TypeAttachedList();
                foreach (string key in tal.mlsTypeAttached.Keys)
                {
                    comboBox4.Items.Add(tal.mlsTypeAttached[key]);
                }
            }

        }

        private void order_prefill_button_Click(object sender, EventArgs e)
        {

            StringBuilder macro = new StringBuilder();
            DataTable subject_table = subjectTableAdapter.GetData();
            //iMacros.App bpoForm = new iMacros.App();
            int timeout = 15;
            string filter = "MainPin = '" + subjectpin_textbox.Text.ToString() + "'";

            DataRow[] foundRows;
            foundRows = subject_table.Select(filter);

            iim2.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            
            string currentUrl = iim2.iimGetLastExtract();

         //   MessageBox.Show(currentUrl);

            #region solutionstar
            if (currentUrl.ToLower().Contains("solutionstar.gatorsasp.com"))
            {
                SolutionStar bpoform = new SolutionStar();
                streetnumTextBox.Text = "SolutionStar";
                bpoform.Prefill(iim2, this);
            }
            #endregion

            #region bpofullfillment aka mainstreet aka redbell
            if (currentUrl.ToLower().Contains("bpofulfillment"))
            {
                //AVM bpoform = new AVM();
                //bpoform.Prefill(iim2, this);
                macro.AppendLine(@"SET !REPLAYSPEED MEDIUM");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$txtInspectionDate$txtMonth CONTENT=04");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$txtInspectionDate$txtDay CONTENT=10");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$txtInspectionDate$txtYear CONTENT=2014");
               
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$cboInfoSource CONTENT=%483");
                //764 = single gamily residence
                //149 = condo
                if (SubjectDetached)
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$cboPropertyType CONTENT=%764");
                }

                if (SubjectAttached)
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$cboPropertyType CONTENT=%149");
                }

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$txtAssessorParcel CONTENT=" + SubjectPin);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$txtAssessorParcel CONTENT=" + SubjectPin);

                if (GlobalVar.theSubjectProperty.ListedInLastYear)
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$cboListedLast12Months CONTENT=%113");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$cboCurrentlyListed CONTENT=%113");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$cboMultipleListings CONTENT=%113");
                }

                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$cboSubjectVisibility CONTENT=%208");
           
          
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$cboPropertyVacant CONTENT=%209");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$cboSecured CONTENT=%208");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$cboPropertyView CONTENT=%485");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$txtTaxes CONTENT=" + GlobalVar.theSubjectProperty.PropertyTax.Replace(",", "").Replace("$", ""));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$txtOwnerPubRec CONTENT=" + SubjectOOR.Replace(" ", "<SP>"));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$txtLandValue CONTENT=1");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$txtDelinquentTaxes CONTENT=0");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$cboPredominantOccupancy CONTENT=%497");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$SubjectHistory$LegalDesc CONTENT=UNK");

                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=ACTION:Bpo.aspx* ATTR=TXT:Next");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$txtCounty1 CONTENT=" + SubjectCounty.Replace(" ", "<SP>"));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$txtSubdivision1 CONTENT=" + SubjectSubdivision.Replace(" ", "<SP>"));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$txtAreaName1 CONTENT=" + SubjectSubdivision.Replace(" ", "<SP>"));

               

                double x = -1;
                string s = SubjectLotSize;
                Double.TryParse(s, out x);

               x = Math.Round(x, 2);


               macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");


                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$txtLotSize1 CONTENT=" +  x.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$txtSite1 CONTENT=level");
              //needs logic
                //  macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$cboPropertyStyle1 CONTENT=%81");
         
              //  macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$cboConstruction1 CONTENT=%103");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$cboConstruction1 CONTENT=%104");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$cboCondition1 CONTENT=%348");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$txtNumUnits1 CONTENT=1");
           
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$txtYearBuilt1 CONTENT=" + SubjectYearBuilt);
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$cboLandscaping1 CONTENT=%112");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$txtBasementFinished1 CONTENT=0");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$txtBasementSqft1 CONTENT=" + SubjectBasementGLA);

                //needs logic
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$cboGarageCarport1 CONTENT=%531");
                
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$cboLeaseHold1 CONTENT=%415");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$txtPoolSpaFirplace1 CONTENT=unk");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$txtTotalSQ1 CONTENT=" + SubjectAboveGLA.Replace(",", ""));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$txtTotalRooms1 CONTENT=" + SubjectRoomCount);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$txtBedrooms1 CONTENT=" + SubjectBedroomCount);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$txtFullBath1 CONTENT=" + SubjectBathroomCount[0]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$txtHalfBath1 CONTENT=" + SubjectBathroomCount[2]);
                
                
                //need logic
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$Comparables$cboBasement1 CONTENT=%330");


                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=ACTION:Bpo.aspx* ATTR=TXT:Next");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboPrideOfOwnerShip CONTENT=%45");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboCrimeVandalRisk CONTENT=%105");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboNeighborhoodTrend CONTENT=%321");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboHomeValues CONTENT=%322");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboHomeValues CONTENT=%0");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboEnviromentalIssues CONTENT=%209");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboEnviromentalIssues CONTENT=%210");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboHomeValues CONTENT=%322");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$txtRateOfPerMonth CONTENT=1");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$txtOwnerOccupied CONTENT=95");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboSupply CONTENT=%322");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboDemand CONTENT=%324");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboSupply CONTENT=%324");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboDemand CONTENT=%324");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboPredominantBuyer CONTENT=%536");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboREOTrend CONTENT=%324");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboMarketingTrend CONTENT=%324");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboZoningCompliance CONTENT=%687");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboZoningCompliance CONTENT=%691");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboNeigConstruction CONTENT=%113");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$txtPctList CONTENT=93.0000");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$BoardedHomes CONTENT=0");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboEmpConditions CONTENT=%324");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboRentControl CONTENT=%113");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboIndustrialWithIn CONTENT=%113");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboDisaster CONTENT=%113");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx* ATTR=NAME:BPO$NeighborhoodInfo$cboSearchCriteria CONTENT=%0");
                macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=ACTION:Bpo.aspx* ATTR=TXT:<Unassigned><SP>GLA<SP>Subdivision<SP>Zip<SP>Code<SP>Map<SP>Grid<SP>Price<SP>Room<SP>Count<SP>Other");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=ACTION:Bpo.aspx* ATTR=TXT:Save");
              

            }
            #endregion


            #region AVM
            if (currentUrl.ToLower().Contains("avm.assetval.com"))
            {
                            AVM bpoform = new AVM();
                            streetnumTextBox.Text = "AVM";
                            bpoform.Prefill(iim2, this);
            }
            #endregion


            #region dispo
            if (currentUrl.ToLower().Contains("disposolutions"))
            {
                Dispo bpoform = new Dispo();
                streetnumTextBox.Text = "dispo";
                bpoform.Prefill(iim2, this);
            }
            #endregion  

            #region goodmandean
            if (currentUrl.ToLower().Contains("goodmandean"))
            {

                GoodmanDean bpoform = new GoodmanDean();
                streetnumTextBox.Text = "goodmandean";
                bpoform.Prefill(iim2, this);
            }
            #endregion  

            #region oldrep
            if (currentUrl.ToLower().Contains("ort.quan"))
            {
                
                OldRep bpoform = new OldRep();
                streetnumTextBox.Text = "oldrepublic";
                bpoform.Prefill(iim2, this);
            }
            #endregion  

            #region landsafe
            if (currentUrl.ToLower().Contains(@".collateraldna.com/forms"))
            {
                LandSafe bpoform = new LandSafe();
                streetnumTextBox.Text = "landsafe";
                bpoform.Prefill(iim2, this);
            }
            #endregion  


            #region equitrax
            if (currentUrl.ToLower().Contains("equi-trax"))
            {
                Equitrax bpoform = new Equitrax();
                streetnumTextBox.Text = "equitrax";
                bpoform.Prefill(iim2, this);
            }
            #endregion  

            //
            //LRES - Lighthouse
            //
            #region lres
            if (currentUrl.ToLower().Contains("lres.com"))
            {
                Lres lres = new Lres(this);
                streetnumTextBox.Text = "lres";
                lres.Prefill(iim2);
            }


            #endregion

            //
            //Emort
            //
            #region emort
            if (currentUrl.ToLower().Contains("emortgage"))
            {
                Emortgage emort = new Emortgage();
                streetnumTextBox.Text = "emort";
                emort.Prefill(iim2, this);
            }
            #endregion  

            //
            //BrokerPriceOpinion.com PCR
            //
            #region bpo.com
            if (streetnumTextBox.Text == "bpo.com")
            {
                string content = "";
                //
                //Page 1
                //
                StringBuilder page1 = new StringBuilder();
                page1.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_IEULAAPPR CONTENT=%True");
                page1.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:form ATTR=VALUE:Save<SP>and<SP>Continue<SP>>>");
                status = iim2.iimPlayCode(page1.ToString(), timeout);

                //
                //Page 2
                //
                macro.Clear();
                //Inspection Date
                string month  = DateTime.Now.Month.ToString();
                string day = DateTime.Now.Day.ToString();
                string year = DateTime.Now.Year.ToString();

                if (month.Length == 1)
                {
                    month = "0" + month;
                }

                if (day.Length == 1)
                {
                    day = "0" + day;
                }
                content =  month + day + year;

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form ATTR=ID:Data_SID CONTENT=" + content);
                //Property Type
                if (subjectDetachedradioButton.Checked)
                {
                    content = "%SFR";
                }
                else
                {
                    content = "%Condo";
                   
                }
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_MDSUBPT CONTENT=" + content);
                //style
                content = "";

                switch (subjectMlsTypecomboBox.Text)
                {
                    case "1 Story" :
                        content = "%Ranch";
                        break;
                    case "1.5 Story" :
                        content = "%1.5<SP>Story";
                        break;
                    case "Raised Ranch" :
                        content = "%Raised<SP>Ranch";
                        break;
                    case "2 Stories":
                        content = "%2<SP>Story";
                        break; 
                }
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_MDSUBSTYLE CONTENT=" + content);
                content = "";

                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SLOC CONTENT=%Suburban");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SOCC CONTENT=%Occupied");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SPAS CONTENT=%Yes");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_MDSUBCOND CONTENT=%Average<SP>or<SP>Better");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SSAPPEAL CONTENT=%Typical<SP>for<SP>Area");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_MDSUBSSC CONTENT=%Typical<SP>for<SP>Area");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_MDSUBGLACOMPARISON CONTENT=%Typical<SP>for<SP>Area");
                //garage CONTENT=%2<SP>Cars<SP>Att
                if (subjectParkingTypeTextbox.Text.ToLower().Contains("gar"))
                {
                    string spaces = Regex.Match(subjectParkingTypeTextbox.Text, "\\d").Value;
                    if (spaces == "")
                    {
                        spaces = "2";
                    }

                    if (Convert.ToInt16(spaces) > 5)
                    {
                        spaces = "5+";
                    }
                    string garType = "Att";
                    if (subjectParkingTypeTextbox.Text.ToLower().Contains("det"))
                    {
                        garType = "Det";
                    }

                    if (spaces == "1")
                    {
                        content = "%" + spaces + "<SP>Car<SP>" + garType;
                    }
                    else
                    {
                        content = "%" + spaces + "<SP>Cars<SP>" + garType;
                    }
                    

                }
                else
                {
                    content = "%None";
                }
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SGARAGE CONTENT=" + content);
                content = "";

                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SGARCOMPARISON CONTENT=%Typical<SP>for<SP>Area");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SLANDSCAPING CONTENT=%Typical<SP>for<SP>Area");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form ATTR=ID:Data_RESTEXTCOST CONTENT=0");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SMAINTENANCE CONTENT=%Average");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SCONFORMITY CONTENT=%Yes");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SCL CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SISNEW CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_RRR CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SRDAMAGE CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SWINDAMAGE CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SDOORDAMAGE CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SSIDINGDAMAGE CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SSTRUCTUREDAMAGE CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SFIREDAMAGE CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_SWATERDAMAGE CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:form ATTR=ID:Data_SVNAR CONTENT=Subject<SP>is<SP>maintained<SP>and<SP>landscaped.<SP>Located<SP>within<SP>an<SP>area<SP>of<SP>maintained<SP>homes,<SP>subject<SP>conforms.");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:form ATTR=VALUE:Save<SP>and<SP>Continue<SP>>>");
                status = iim2.iimPlayCode(macro.ToString(), timeout);
                macro.Clear();

                //
                //Page 3
                //
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_NMTDDL CONTENT=%Stable");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form ATTR=ID:Data_NADP CONTENT=0");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form ATTR=ID:Data_NMKTM CONTENT=120");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_NGATED CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form ATTR=ID:Data_NCLS CONTENT=5");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_NNEARBYVAC CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_NNEARBYBOARDED CONTENT=%No");
                macro.AppendLine(@"TAG POS=8 TYPE=DIV FORM=ID:form ATTR=CLASS:ZZ_Value");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_NNEARBYPOWERLINES CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_NNEARBYRAILROAD CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_NNEARBYHIGHWAY CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_NNEARBYFLIGHTPATH CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_NNEARBYCOMMPROP CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_NNEARBYWASTEFAC CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_NNEARBYVANDALISM CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=DIV FORM=ID:form ATTR=CLASS:VTITLE");
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:form ATTR=ID:Data_NVNAR CONTENT=Describe<SP>the<SP>neighborhood,<SP>including<SP>vandalism<SP>risk,<SP>type<SP>of<SP>homes,<SP>age<SP>of<SP>neighborhood,<SP>proximity<SP>to<SP>linkages<SP>and<SP>schools.");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form ATTR=ID:Data_NNEARBYWASTEFAC CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:form ATTR=ID:Data_NVNAR CONTENT=Established<SP>Residential<SP>Area.<SP>Low<SP>risk<SP>of<SP>vandalism.<SP>80s<SP>and<SP>90s<SP>era<SP>tract<SP>homes.<SP>Close<SP>proximity<SP>to<SP>all<SP>amenities:(shopping,<SP>school,<SP>freeway)");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:form ATTR=VALUE:Save<SP>and<SP>Continue<SP>>>");
               
                status = iim2.iimPlayCode( macro.ToString(), timeout);
                macro.Clear();

                //
                //Page 4
                //
                
                //subject photos
                content = search_address_textbox.Text.Replace(" ", "<SP>");
             
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:form ATTR=NAME:file CONTENT=" + content + @"\front.jpg");
                macro.AppendLine(@"WAIT SECONDS=10");
                macro.AppendLine(@"TAG POS=2 TYPE=INPUT:FILE FORM=ID:form ATTR=NAME:file CONTENT=" + content + @"\street.jpg");
                macro.AppendLine(@"WAIT SECONDS=10");
                macro.AppendLine(@"TAG POS=3 TYPE=INPUT:FILE FORM=ID:form ATTR=NAME:file CONTENT=" + content + @"\address.jpg");
                macro.AppendLine(@"WAIT SECONDS=10");
                macro.AppendLine(@"TAG POS=4 TYPE=INPUT:FILE FORM=ID:form ATTR=NAME:file CONTENT=" + content + @"\side.jpg");
                macro.AppendLine(@"WAIT SECONDS=10");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:form ATTR=VALUE:Save<SP>and<SP>Continue<SP>>>");
               
                status = iim2.iimPlayCode( macro.ToString(), timeout);
                macro.Clear();
                //TBD:
                //optional photos
                //

                //
                //Page 5
                //
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form ATTR=ID:Data_XSIGNOFF CONTENT=Dawn");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:form ATTR=VALUE:Submit<SP>and<SP>Return<SP>to<SP>Inbox<SP>>>");
              
                status = iim2.iimPlayCode(macro.ToString(), timeout);
                macro.Clear();
            }

            #endregion


            #region sls
            if (streetnumTextBox.Text == "sls")
            {
                DateTime subject_age = new DateTime((Convert.ToInt32(subjectYearBuiltTextbox.Text)), 1, 1);

                TimeSpan ts = DateTime.Now - subject_age;

                int age = ts.Days / 365;

                //
                //Header
                //
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtInspDate CONTENT=" + DateTime.Now.ToShortDateString());
                macro.AppendLine(@"TAG POS=1 TYPE=LABEL FORM=ID:form1 ATTR=TXT:Occupied");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtParcelNum CONTENT=" + subjectpin_textbox.Text);

                //
                //Market conditions
                //
                macro.AppendLine(@"TAG POS=1 TYPE=LABEL FORM=ID:form1 ATTR=TXT:Stable");
                macro.AppendLine(@"TAG POS=2 TYPE=LABEL FORM=ID:form1 ATTR=TXT:Stable");
                macro.AppendLine(@"TAG POS=1 TYPE=LABEL FORM=ID:form1 ATTR=TXT:Remained<SP>stable");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form1 ATTR=ID:BPOForm1_chkMarketHist_2&&VALUE:on CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtOwnerPercent CONTENT=90");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtTenantPercent CONTENT=10");
                macro.AppendLine(@"TAG POS=1 TYPE=LABEL FORM=ID:form1 ATTR=TXT:Oversupply");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form1 ATTR=ID:BPOForm1_chkListingSupply_1&&VALUE:on CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtBlockedHomes CONTENT=0");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlOwnerPride CONTENT=%2");
                //macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:form1 ATTR=ID:BPOForm1_txtMarketCondComments");

                //
                //Subject Marketability
                //
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtValueRangeLow CONTENT=" + oneMile.soldPrice[3].Replace("$","").Replace(",",""));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtValueRangeHigh CONTENT=" + oneMile.soldPrice[0].Replace("$", "").Replace(",", ""));
                macro.AppendLine(@"TAG POS=1 TYPE=LABEL FORM=ID:form1 ATTR=TXT:Appropriate<SP>improvement<SP>for<SP>the<SP>neighborhood");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form1 ATTR=ID:BPOForm1_chkImprovement_2&&VALUE:on CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtMarketTime CONTENT=180");
                macro.AppendLine(@"TAG POS=1 TYPE=LABEL FORM=ID:form1 ATTR=TXT:Yes");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form1 ATTR=ID:BPOForm1_chkAvailFinancing_0&&VALUE:on CONTENT=YES");
                macro.AppendLine(@"TAG POS=2 TYPE=LABEL FORM=ID:form1 ATTR=TXT:No");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form1 ATTR=ID:BPOForm1_chkMarketPeriod_1&&VALUE:on CONTENT=YES");
                macro.AppendLine(@"TAG POS=3 TYPE=TD FORM=ID:form1 ATTR=TXT:YesNo");
                macro.AppendLine(@"TAG POS=3 TYPE=LABEL FORM=ID:form1 ATTR=TXT:No");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form1 ATTR=ID:BPOForm1_chkIsListed_1&&VALUE:on CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=LABEL FORM=ID:form1 ATTR=TXT:Single<SP>family<SP>attached");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form1 ATTR=ID:BPOForm1_chkUnitType_0&&VALUE:on CONTENT=YES");

                //
                //Subject info
                //
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlSubjLoc1 CONTENT=%2");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlSubjLeaseFee1 CONTENT=%1");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjSize1 CONTENT=.2");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjView1 CONTENT=Residential");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjDesign1 CONTENT=2<SP>Story<SP>/<SP>Avg");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlSubjQuality1 CONTENT=%3");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjAge1 CONTENT=" + age.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlSubjCond1 CONTENT=%3");
                macro.AppendLine(@"TAG POS=148 TYPE=TD FORM=ID:form1 ATTR=CLASS:BpoGrid");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjRmCountTotal1 CONTENT=" + subjectRoomCountTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjRmCountBdms1 CONTENT=" + subjectBedroomTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjRmCountBaths1 CONTENT=" + subjectBathroomTextbox.Text[0]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjRmCountHBath1 CONTENT=" + subjectBathroomTextbox.Text[2]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjLivArea1 CONTENT=" + subjectAboveGlaTextbox.Text);
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjBaseRms1 CONTENT=1111");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjPercFin1 CONTENT=100");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlSubjUtility1 CONTENT=%3");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjHeatAC1 CONTENT=Gas/Central");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjEnergy1 CONTENT=NA");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjGarage1 CONTENT=2GA");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjDeckEtc1 CONTENT=Unknown");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjPoolEtc1 CONTENT=NA");

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjOther1 CONTENT=NA");
                
                //
                //footer
                //
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSignature CONTENT=Scott<SP>Beilfuss,<SP>C-REPS");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtLicense CONTENT=471.009163");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtAsIsMarketVal CONTENT=" + valueTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtAsIsSugListPrice CONTENT=" + listTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtRepairedMarketVal CONTENT=" + valueTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtRepairedSugListPrice CONTENT=" + listTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtFairMarketRent CONTENT=" + subjectRentTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtAsIsMarketValQuick CONTENT=" + quickSaleTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtAsIsSugListPriceQuick CONTENT=" + quickSaleTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtRepairedMarketValQuick CONTENT=" + quickSaleTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtRepairedSugListPriceQuick CONTENT=" + quickSaleTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtSubjectLandValue CONTENT=5000");

            }
            #endregion

            #region nvs
            if (streetnumTextBox.Text == "nvs")
            {
                //
                //Page 1
                //

                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Type CONTENT=%1");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Units CONTENT=%1");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:OccupancyStatus CONTENT=%1");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:HOAFees CONTENT=NA");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:Years_Experience CONTENT=5");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:PreparingAgent CONTENT=Scott<SP>Beilfuss");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:bpoform ATTR=VALUE:Go<SP>to<SP>Next<SP>Section<SP>-<SP>Save");
               

                //
                //Page 2
                //
            
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:DateInspected CONTENT=" + DateTime.Now.ToShortDateString());
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Condition CONTENT=%Average");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:Beds CONTENT=" + subjectBedroomTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:FullBaths CONTENT=" + subjectBathroomTextbox.Text[0]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:HalfBaths CONTENT=" + subjectBathroomTextbox.Text[2]);
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:Rooms CONTENT=" + subjectRoomCountTextbox.Text);
                
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:YearBuilt CONTENT=" + subjectYearBuiltTextbox.Text);
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:SqFeet CONTENT=" + subjectAboveGlaTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:LotSize CONTENT=" + subjectLotSizeTextbox.Text);
                if (subjectStyleTextbox.Text.ToLower().Contains("1 story"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Style CONTENT=%Ranch");
                }
                else
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Style CONTENT=%Contemporary");
                }
                
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Construction_type CONTENT=%1");

                //crawl
                if (subjectBasementDetailsTextbox.Text.Contains("Crawl"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%C");
                }
               
                //none
                if (subjectBasementTypeTextbox.Text == "None" | subjectBasementTypeTextbox.Text == "")   //.Text.Contains("None"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%N");
                }
                
                
                if (subjectBasementTypeTextbox.Text.Contains("Full"))
                {
                    if (subjectBasementDetailsTextbox.Text.Contains("Finished"))
                    {
                        //Full finished                    
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%F");
                    }
                    else
                    {
                        //full unfinished  
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%Y");
                    }
                }
                

                
                if (subjectBasementTypeTextbox.Text.Contains("Partial"))
                {
                    if (subjectBasementDetailsTextbox.Text.Contains("Finished"))
                    {
                        //partial finish
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%P");
                    }
                    else
                    {
                        //partial unfinished
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%X");
                    }
                }
                    
                
                if (subjectNumFireplacesTextbox.Text == "")
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Fireplace CONTENT=%0");
                }
                else
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Fireplace CONTENT=%1");
                }
                
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Pool CONTENT=%N");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Carport CONTENT=%0");

                if (subjectParkingTypeTextbox.Text.Contains("Gar"))
                {
                    string [] numSpaces = subjectParkingTypeTextbox.Text.Split(' ');
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Garage CONTENT=%D" + numSpaces[1]);
                }
                else
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Garage CONTENT=%N");
                }
                
                
                
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Porch CONTENT=%N");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Patio CONTENT=%N");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:bpoform ATTR=VALUE:Go<SP>to<SP>Next<SP>Section<SP>-<SP>Save");
                
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");


            }


            #endregion


            #region m2m



            if (currentUrl.ToLower().Contains("marktomarket"))
            {
                StringBuilder m = new StringBuilder();
                //m.AppendLine(@"TAG POS=1 TYPE=TABLE FORM=ACTION:/Order/OrderEditWizardAHM/Step1/11106449 ATTR=TXT:* EXTRACT=TXT");
                //m.AppendLine(@"TAG POS=1 TYPE=A FORM=ACTION:/Order/OrderEditWizardAHM/Step1/11106449 ATTR=HREF:http://www.marktomarket.us/Order/OrderMgmt/Detail/11106449 EXTRACT=TXT");

                m.AppendLine(@"TAG POS=1 TYPE=A FORM=ACTION:/Order/OrderEditWizard*/* ATTR=HREF:http://www.marktomarket.us/Order/OrderMgmt/Detail/* EXTRACT=TXT");
                string mc = m.ToString();
                status = iim2.iimPlayCode(mc, 60);
                string orderNumber = iim2.iimGetLastExtract(1);

                macro.AppendLine(@"SET !ERRORIGNORE YES");
                macro.AppendLine(@"SET !TIMEOUT_STEP 0");
                macro.AppendLine(@"SET !REPLAYSPEED FAST");

                //
                //Page1
                //

                //OOR
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizardAHM/Step1/" + orderNumber + " ATTR=ID:Subject_Owner CONTENT=" + subjectOorTextbox.Text.Replace(" ","<SP>"));
                //county
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizardAHM/Step1/" + orderNumber + " ATTR=ID:Subject_County CONTENT=%" + subjectCountyTextbox.Text);
                //office proximity
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizardAHM/Step1/" + orderNumber + " ATTR=ID:DistanceAgentSubject CONTENT=" + subjectProximityToOfficeTextbox.Text);
                //sales history
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizardAHM/Step1/" + orderNumber + " ATTR=ID:DateListed1 CONTENT=1/1/12");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizardAHM/Step1/" + orderNumber + " ATTR=ID:DateSold1 CONTENT=" + subjectLastSaleDateTextbox.Text.Replace(" ", "<SP>"));
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizardAHM/Step1/" + orderNumber + " ATTR=ID:ListPrice1-input-text CONTENT=Enter<SP>value");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizardAHM/Step1/" + orderNumber + " ATTR=ID:SalePrice1-input-text CONTENT=" + subjectLastSalePriceTextbox.Text.Replace(" ", "<SP>"));
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizardAHM/Step1/" + orderNumber + " ATTR=ID:Comments1");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:/Order/OrderEditWizardAHM/Step1/" + orderNumber + " ATTR=ID:SaveButton&&VALUE:Save");
                

                //
                //Page2
                //
                //TO DO: Get market data.
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:/Order/OrderEditWizardAHM/Step1/" + orderNumber + " ATTR=ID:step2_submit&&VALUE:2");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:Subject_PropertyType CONTENT=%Single<SP>Family<SP>Detached");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:NoUnits CONTENT=1");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:Subject_Basement CONTENT=Unk");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:OccupancyStatus CONTENT=%Unknown");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:QualityOfConstruction CONTENT=%Average");
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:ZoningClassification CONTENT=Single<SP>Familty<SP>Residential");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:CurrentUseBestUse&&VALUE:True CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:PotentialRentAmt-input-text CONTENT=" + subjectRentTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:LocationType CONTENT=%Suburban");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:CompetitiveListings CONTENT=" + oneMile.totalActive);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:LowerAreaPriceRange-input-text CONTENT=" + oneMile.listPrice[3]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:UpperAreaPriceRange-input-text CONTENT=" + oneMile.listPrice[0]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:NumberOfComparableListings CONTENT=" + oneMile.totalSold );
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:NumberOfComparableSalesLow-input-text CONTENT=" + oneMile.soldPrice[3]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:NumberOfComparableSalesHigh-input-text CONTENT=" + oneMile.soldPrice[0]);
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:DemandSupply CONTENT=%Low");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:/Order/OrderEditWizardAHM/Step2/" + orderNumber + " ATTR=ID:SaveButton&&VALUE:Save");

                //
                //Page 3, subject info
                //
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON ATTR=ID:step3_submit&&VALUE:3");

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:Subject_Gla CONTENT=" + SubjectAboveGLA.Replace(",", ""));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form0 ATTR=ID:Subject_AgtBrokerInspected&&VALUE:False CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form0 ATTR=ID:Subject_AgtBrokerInspected&&VALUE:True CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:Subject_YearBuilt CONTENT=" + SubjectYearBuilt);
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:Subject_Condition CONTENT=$Select<SP>a<SP>Condition<SP>>>");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:Subject_Condition CONTENT=%3-Average");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:Subject_TotalRooms CONTENT=" + SubjectRoomCount);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:Subject_Bedrooms CONTENT=" + SubjectBedroomCount);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:Subject_Bathrooms CONTENT=" + SubjectBathroomCount);
                //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:Subject_GarageCarportDescription CONTENT=%3<SP>Stall<SP>Attached");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:Subject_SiteSize CONTENT=" + SubjectLotSize);

                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/* ATTR=NAME:Subject.PropertyType CONTENT=%Single<SP>Family<SP>Detached");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/* ATTR=NAME:Subject.Style CONTENT=%1<SP>1/2<SP>Story");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/* ATTR=NAME:Subject.Exterior CONTENT=%Metal/Vinyl");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/* ATTR=NAME:Subject.YearBuilt CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectYearBuilt);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/* ATTR=NAME:Subject.AboveGradeSf CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectAboveGLA);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/* ATTR=NAME:Subject.FinishedSf CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectAboveGLA);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/* ATTR=NAME:Subject.TotalRooms CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectRoomCount);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/* ATTR=NAME:Subject.Bedrooms CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectBedroomCount);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/* ATTR=NAME:Subject.Bathrooms CONTENT=" + GlobalVar.theSubjectProperty.FullBathCount);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/* ATTR=NAME:Subject.HalfBaths CONTENT=" + GlobalVar.theSubjectProperty.HalfBathCount);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/* ATTR=NAME:Subject.BasementRooms CONTENT=0");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/* ATTR=NAME:Subject.BasementSQFT CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectBasementGLA);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/* ATTR=NAME:Subject.PercentageBasementFinished CONTENT=0");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/* ATTR=NAME:Subject.Garage CONTENT=%1<SP>ATTACHED");

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/* ATTR=NAME:Subject.SiteSize CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectLotSize);

                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/* ATTR=NAME:Subject.Pool CONTENT=%None");


            }

            #endregion

            #region vp

            if (streetnumTextBox.Text == "vp")
            {
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtREVIEWDESC CONTENT=Exterior");
               
                
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtCURRENTOWNER CONTENT=" + subjectOorTextbox.Text.Replace(" ", "<SP>"));
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtCOUNTY CONTENT=LAKE");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtASSESPARCELNUM CONTENT=" + subjectpin_textbox.Text);
                //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:ddlPROJTYPE CONTENT=%SFR");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtPROJTYPEDESC");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtHOMEOWNERASSNFEE");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form1 ATTR=ID:rbGENERALNEWCONSTRUCT1&&VALUE:YES CONTENT=NO");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form1 ATTR=ID:rbGENERALNEWCONSTRUCT2&&VALUE:NO CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form1 ATTR=ID:rbSUBJECTDISASTERRESPONSE1&&VALUE:YES CONTENT=NO");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form1 ATTR=ID:rbSUBJECTDISASTERRESPONSE2&&VALUE:NO CONTENT=YES");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTDISASTERDT");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:ddlSUBJECTSOLDLISTED CONTENT=%NO");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtLISTSUBJECTMARKETDAYS");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form1 ATTR=ID:rbSALESHISTORYRESPONSE1&&VALUE:YES CONTENT=NO");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form1 ATTR=ID:rbSALESHISTORYRESPONSE2&&VALUE:NO CONTENT=YES");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTSOLDLISTEDPREVPRICE");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTSOLDLISTEDPREVDATE");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtLISTSUBJECTORGPRICE");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtLISTSUBJECTORGLISTINGDT");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtLISTSUBJECTLISTINGPRICE");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtLISTSUBJECTPRICEREVDT");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtLISTBROKER");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtLISTBROKERCONAME");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtLISTBROKERPHONE");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:ddlCURRENTOCCUPANT CONTENT=%Owner");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtREPORTDT CONTENT=7/3/12");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form1 ATTR=ID:rbREVIEWTYPE1&&VALUE:EXTERIOR CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:form1 ATTR=ID:txtREVIEWSUBJECTCOMMENTS CONTENT=Located<SP>within<SP>an<SP>area<SP>of<SP>maintained<SP>homes,<SP>subject<SP>conforms.<SP>No<SP>repairs<SP>noted<SP>from<SP>drive-by.<SP>Typical<SP>construction<SP>for<SP>area.");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:ddlLOCATION CONTENT=%Suburban");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form1 ATTR=ID:rbOCCUPANCYOWNER1&&VALUE:Owner CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:ddlPROPVALUES CONTENT=%Stable");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtMARKETINGTIME CONTENT=180");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form1 ATTR=ID:rbLANDUSEINDUSTRIAL2&&VALUE:NO CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:ddlOCCUPANCYVACANTPERCENT CONTENT=%0TO5");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form1 ATTR=ID:rbNBRHOODNEWCONSTRUCTION1&&VALUE:YES CONTENT=NO");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form1 ATTR=ID:rbNBRHOODNEWCONSTRUCTION2&&VALUE:NO CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form1 ATTR=ID:rbNBRHOODDISASTERRESPONSE1&&VALUE:YES CONTENT=NO");
                macro.AppendLine(@"TAG POS=6 TYPE=LABEL FORM=ID:form1 ATTR=TXT:No");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form1 ATTR=ID:rbNBRHOODDISASTERRESPONSE2&&VALUE:NO CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtNBRHOODDISASTERDT");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtMEDIANRENT CONTENT=2000");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtCOMMENTREOSALESEFFECT CONTENT=Increasing");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtMARKETINGTIMEDESC CONTENT=Stable");

                //
                //subject info
                //

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTLIVINGSQFT CONTENT=" + subjectAboveGlaTextbox.Text);
               
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTAGEYRS CONTENT=" );
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTLOTSIZE CONTENT=" + subjectLotSizeTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTROOMTOT CONTENT=" + subjectRoomCountTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTBEDROOMS CONTENT=" + subjectBedroomTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTBATHFULL CONTENT=" + subjectBathroomTextbox.Text[0].ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTBATHHALF CONTENT=" + subjectBathroomTextbox.Text[2].ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTDESIGNSTYLE CONTENT=");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTVIEW CONTENT=Residential");

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTNUMUNITS CONTENT=1");

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTBASEMENTTYPE CONTENT=" + subjectBasementTypeTextbox.Text.Replace(" ", "<SP>"));
               // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTBASEMENT");
               // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTBASEMENTFIN");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTGARAGE CONTENT=" + subjectParkingTypeTextbox.Text.Replace(" ", "<SP>"));
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTGARAGENUMCARS");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTCONDITION");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtNUMCOMPLISTINGS CONTENT=2");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtCOMPLISTINGPRICELOW");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtNUMCOMPLISTINGS CONTENT=3");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtCOMPLISTINGPRICELOW CONTENT=550000");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtCOMPLISTINGPRICEHIGH CONTENT=899900");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtNUMCOMPSALES CONTENT=1");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtCOMPSALESPRICELOW CONTENT=650000");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtCOMPSALESPRICEHIGH CONTENT=650000");
                //macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:form1 ATTR=ID:txtREVIEWNBRHOODDESC CONTENT=Secluded/Isolated<SP>Area.<SP><SP>Rural<SP>like,<SP>but<SP>close<SP>to<SP>amenitites.");
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:form1 ATTR=ID:txtMARKETCONDITIONS CONTENT=Area<SP>of<SP>limited<SP>buyer<SP>demand<SP>with<SP>high<SP>invertory.");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:form1 ATTR=ID:SaveWork11&&VALUE:Save<SP>Work");

            }

            #endregion vp


            #region usres prefill

            if (currentUrl.ToLower().Contains("usres"))
            {

                StringBuilder read_header = new StringBuilder();

              
                read_header.AppendLine(@"TAG POS=4 TYPE=TABLE FORM=NAME:InputForm ATTR=CLASS:form_txt_blk EXTRACT=TXT");
                string read_headerCode = read_header.ToString();
                status = iim2.iimPlayCode(read_headerCode, 30);


                string table1 = iim2.iimGetLastExtract(1);


                string pattern = "IL.\\s+(\\d\\d\\d\\d\\d)";
                Match match = Regex.Match(iim2.iimGetLastExtract(1), pattern);
                string orderzip = match.Groups[1].Value;



                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abParcel CONTENT=");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBkrDist CONTENT=" + this.Get_Distance(orderzip));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBkrLic CONTENT=471.009163");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abAgExp CONTENT=7");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:d CONTENT=YES");
                macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:s CONTENT=YES");
                macro.AppendLine(@"TAG POS=3 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:s CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:s CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:o CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=NAME:InputForm ATTR=TXT:(*)<SP>(must<SP>be<SP>a<SP>number)");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMcOwnerPerc CONTENT=90");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMcTenentPerc CONTENT=10");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMcCompUnits CONTENT=" + oneMile.totalActive);
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMcCompLists CONTENT=");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMcBoards CONTENT=0");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abNbBound CONTENT=1<SP>mile<SP>radius");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMkLow CONTENT=" + oneMile.soldPrice[3]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMkHigh CONTENT=" + oneMile.soldPrice[0]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abPropSold CONTENT=" + oneMile.totalSold);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMkListLow CONTENT=" + oneMile.listPrice[3]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMkListHigh CONTENT=" + oneMile.listPrice[0]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abPropList CONTENT=" + oneMile.totalActive);
                
                macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:o CONTENT=NO");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:u CONTENT=NO");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:a CONTENT=YES");
               
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMkTime");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:u CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMkTime CONTENT=180");
                macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:n CONTENT=YES");
                macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:y CONTENT=YES");
                macro.AppendLine(@"TAG POS=4 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:n CONTENT=YES");
                macro.AppendLine(@"TAG POS=5 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:n CONTENT=YES");
                macro.AppendLine(@"TAG POS=4 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:d CONTENT=YES");
                macro.AppendLine(@"TAG POS=4 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:o CONTENT=YES");
                macro.AppendLine(@"'TAG POS=0 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:u CONTENT=YES");
                
                //
                //subject info
                //

                
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abParcel CONTENT=" + subjectpin_textbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abSrce CONTENT=%MLS");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abLoc CONTENT=%Suburban");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abRight CONTENT=Fee<SP>Simple");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSite CONTENT=" + subjectLotSizeTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abUnits CONTENT=%1");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abView CONTENT=Neighborhood");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abDesign CONTENT=%Avg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abAge CONTENT=" + subjectYearBuiltTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abCond CONTENT=%Avg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abRooms CONTENT=" + subjectRoomCountTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBeds CONTENT=" + subjectBedroomTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBaths CONTENT=" + subjectBathroomTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abGla CONTENT=" + subjectAboveGlaTextbox.Text);

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBsmt CONTENT=" + subjectBasementDetailsTextbox.Text.Replace(" ", "<SP>"));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abHeat CONTENT=Gas<SP>FA/Central");
                //
                //garage
                //

                 //string [] numSpaces = subjectParkingTypeTextbox.Text.Split(' ');
                string numSpaces = Regex.Match(subjectParkingTypeTextbox.Text, "(\\d+)").Value;

                 if (subjectParkingTypeTextbox.Text.ToLower().Contains("att"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abGarage CONTENT=%" + numSpaces + "CA");
                }
                 else if (subjectParkingTypeTextbox.Text.ToLower().Contains("det"))
                {

                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abGarage CONTENT=%" + numSpaces + "CD");
                }
                else
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abGarage CONTENT=%N");
                }
                    
                //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abGarage CONTENT=%2CA");
                //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abGarage CONTENT=%3CA");
                //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abGarage CONTENT=%3CD");
                //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abGarage CONTENT=%N");


                //
                //Page 2
                //
                 macro.AppendLine(@"TAG POS=1 TYPE=INPUT:IMAGE FORM=NAME:InputForm ATTR=NAME:btnSave");
                 //macro.AppendLine(@"WAIT SECONDS=5");
                 macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:InputForm ATTR=SRC:http://valuations.usres.com/BpoImages/bpo_pg2_btn.gif");
                 //macro.AppendLine(@"WAIT SECONDS=5");
                 macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abRent CONTENT=" + subjectRentTextbox.Text);
                 macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:a CONTENT=YES");
                 macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:o CONTENT=YES");
                 macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:o CONTENT=YES");
                 macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMvAsIs CONTENT=" + valueTextbox.Text);
                 macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMvAsIsSlp CONTENT=" + listTextbox.Text);
                 macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMvRep CONTENT=" + valueTextbox.Text);
                 macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMvRepSlp CONTENT=" + listTextbox.Text);
                 macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMvQuick CONTENT=" + quickSaleTextbox.Text);
                 macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMvSugg CONTENT=" + quickSaleTextbox.Text);
                 macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMvLand CONTENT=5000");

                 macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:InputForm ATTR=ID:txtComments CONTENT=No<SP>adverse<SP>conditions<SP>were<SP>noted<SP>at<SP>the<SP>time<SP>of<SP>inspection<SP>based<SP>on<SP>exterior<SP>observations.");
//                 macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:InputForm ATTR=NAME:txtAddendum CONTENT=Competitive<SP>style<SP>used<SP>due<SP>to<SP>diverse<SP>styles<SP>in<SP>area,<SP>best<SP>available<SP>utilized.");
                 macro.AppendLine(@"TAG POS=1 TYPE=INPUT:IMAGE FORM=NAME:InputForm ATTR=NAME:btnSave");
                 macro.AppendLine(@"TAG POS=2 TYPE=IMG FORM=NAME:InputForm ATTR=SRC:http://valuations.usres.com/BpoImages/bpo_pg1_btn.gif");

                //
                //Save CMA and map screenshots
                //
                 StringBuilder macro2 = new StringBuilder();
                 macro2.AppendLine(@"FRAME NAME=workspace");
                 macro2.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=ID:Map<SP>Results*");
                 macro2.AppendLine(@"WAIT SECONDS=1");
                 macro2.AppendLine(@"SAVEAS TYPE=PNG FOLDER=" + search_address_textbox.Text.Replace(" ", "<SP>") + " FILE=" + "map.jpg");
                 macro2.AppendLine(@"WAIT SECONDS=1");
                macro2.AppendLine(@"FRAME NAME=subheader");
                macro2.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:header ATTR=ID:mapon");
                macro2.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:header ATTR=ID:report_type CONTENT=%1linegrid");
               macro2.AppendLine(@"WAIT SECONDS=1");
                macro2.AppendLine(@"SAVEAS TYPE=PNG FOLDER=" + search_address_textbox.Text.Replace(" ", "<SP>") + " FILE=" + "cma.jpg");
       
                 string macroCode2 = macro2.ToString();
                 status = iim.iimPlayCode(macroCode2, 60);

           

            }

            #endregion



            #region imort prefill
            if (currentUrl.ToLower().Contains("propertysmart"))
            {

                StringBuilder read_header = new StringBuilder();
                read_header.AppendLine(@"SET !ERRORIGNORE YES");
                read_header.AppendLine(@"ONSCRIPTERROR CONTINUE=YES");
                read_header.AppendLine(@"FRAME NAME=_MAIN");
                read_header.AppendLine(@"TAG POS=7 TYPE=TABLE FORM=ID:Form1 ATTR=TXT:* EXTRACT=TXT");
                string read_headerCode = read_header.ToString();
                status = iim2.iimPlayCode(read_headerCode, 30);


                string table1 = iim2.iimGetLastExtract(1);
              

                string pattern = "Zip: #NEXT#(\\d+)";
                Match match = Regex.Match(iim2.iimGetLastExtract(1), pattern);
                string orderzip = match.Groups[1].Value;



                //
                //Header
                //
                macro.AppendLine(@"SET !ERRORIGNORE YES");
                macro.AppendLine(@"ONSCRIPTERROR CONTINUE=YES");
                macro.AppendLine(@"SET !TIMEOUT_STEP 0");
                macro.AppendLine(@"FRAME NAME=_MAIN");


                ///

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Inspection_Type[@Field_Type='checkbox']&&VALUE:Exterior<SP>Only CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Vendor_Email CONTENT=scott@okrealtyplus.com");
                macro.AppendLine(@"TAG POS=25 TYPE=TD FORM=ID:Form1 ATTR=CLASS:lblhd");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Property_Values[@Field_Type='checkbox']&&VALUE:Stable CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Property_Values[@Field_Type='checkbox']&&VALUE:Slow CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Property_Values[@Field_Type='checkbox']&&VALUE:Stable CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Employment_Conditions[@Field_Type='checkbox']&&VALUE:Stable CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Market_Price[@Field_Type='checkbox']&&VALUE:Stable CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Owners CONTENT=90");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Tenants CONTENT=10");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Comparable_Listing_Supply[@Field_Type='checkbox']&&VALUE:normal<SP>supply CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Comparable_Listing_Supply[@Field_Type='checkbox']&&VALUE:over<SP>supply CONTENT=YES");
              
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Subject_Improvement[@Field_Type='checkbox']&&VALUE:shortage CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Marketing_Time CONTENT=180");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Financing_Available[@Field_Type='checkbox'] CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Recently_Listed[@Field_Type='checkbox']&&VALUE:No CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/GENERAL_CONDITIONS/Property_Type[@Field_Type='checkbox']&&VALUE:single<SP>family<SP>detached CONTENT=YES");
              
              
               









                ///





                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Broker_Lic_Num CONTENT=471.009162");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Broker_Years_Exp CONTENT=5");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Broker_Dist_Subject CONTENT=" + this.Get_Distance(orderzip));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/APN CONTENT=" + subjectpin_textbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/School_District CONTENT=" + subject_school_textbox.Text.Replace(" ","<SP>"));
           
                //
                //market data
                //

                //fnma form
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Comparable_Listings CONTENT=" + oneMile.totalOnMarket);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Competing_Listings CONTENT=" + (Convert.ToInt64(oneMile.totalOnMarket) / 3).ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Boarded_Homes CONTENT=0");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Value_Range_Low CONTENT=" + oneMile.soldPrice[3]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Value_Range_High CONTENT=" + oneMile.soldPrice[0]);
                //

                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Neighborhood_Boundaries CONTENT=1<SP>mile<SP>radius");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/NeighborhoodSales1 CONTENT=" + oneMile.soldPrice[3]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/NeighborhoodSales2 CONTENT=" + oneMile.soldPrice[0]);

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/NeighborhoodSalesNum CONTENT=" + oneMile.totalSold);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/NeighborhoodListings1 CONTENT=" + oneMile.listPrice[3]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/NeighborhoodListings2 CONTENT=" + oneMile.listPrice[0]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/NeighborhoodListingsNum CONTENT=" + oneMile.totalOnMarket);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Marketing_Time_Low CONTENT=10");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Marketing_Time_High CONTENT=300");
                //macro.AppendLine(@"TAG POS=34 TYPE=TD FORM=ID:Form1 ATTR=CLASS:lblhd");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Cur_Market_Cond&&VALUE:IMPROVING CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Cur_Market_Cond&&VALUE:STABLE CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Employment_Cond&&VALUE:STABLE CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Market_Price_This_Type&&VALUE:STABLE CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/OwnerOccupantPercent CONTENT=90");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/CompSupply&&VALUE:OVERSUPPLY CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/CompSupply&&VALUE:INBALANCE CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/CompSupply&&VALUE:OVERSUPPLY CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Number_Boarded CONTENT=0");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Number_Comps");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Improvements&&VALUE:APPROPRIATE CONTENT=YES");
                //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Improvements");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/NeighborhoodSales1 CONTENT=" + oneMile.soldPrice[3]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/NeighborhoodSales2 CONTENT=" + oneMile.soldPrice[0]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/NeighborhoodListings1 CONTENT=" + oneMile.listPrice[3]);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/NeighborhoodListings2 CONTENT=" + oneMile.listPrice[0]);
               
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Financing");
                macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Past_12_Months");
                macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Currently_Listed");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Unit_Type");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Last_Sale_Date CONTENT=" + subjectLastSaleDateTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Last_Sale_Listing CONTENT=na");

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Last_Sale_Price CONTENT=" + subjectLastSalePriceTextbox.Text.Replace(" ","<SP>"));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Est_Rent CONTENT="  + subjectRentTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Occupancy_Status");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Est_Rent_As");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Most_Likely_Buyer");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKETING_STRATEGY/Land_Value CONTENT=5000");
                //macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Comments CONTENT=Unincorporated<SP>area<SP>with<SP>McHerny.");
               

                //
                //subject fields
                //

                //fmna
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Design CONTENT=" + subjectMlsTypecomboBox.Text.Replace(" ", "<SP>"));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Construction CONTENT=Average");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Condition  CONTENT=Average");

               
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Utility");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Energy_Efficient");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Garage_Carport CONTENT=" + subjectParkingTypeTextbox.Text.Replace(" ", "<SP>"));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Amenities");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Fence_Pool");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Other");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Marketing_Strategy[@Field_Type='checkbox']&&VALUE:As-Is CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Buyer[@Field_Type='checkbox']&&VALUE:Owner<SP>occupant CONTENT=YES");
               


                //



                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Location CONTENT=Suburban");
               
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Location CONTENT=%Urban");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Location CONTENT=%Suburban");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Location CONTENT=Suburban");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Leasehold CONTENT=Fee<SP>Simple");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Fee_Simple CONTENT=Fee<SP>Simple");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Site CONTENT=" + subjectLotSizeTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Lot_Size CONTENT=" + subjectLotSizeTextbox.Text);
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Num_Units CONTENT=%1");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/View CONTENT=Neighborhood");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Style_Design CONTENT=%Average");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Year_Built CONTENT=" + subjectYearBuiltTextbox.Text);
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Condition CONTENT=%Average");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Total_Rooms CONTENT=" + subjectRoomCountTextbox.Text);
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Bedrooms CONTENT=" + subjectBedroomTextbox.Text);
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Bathrooms CONTENT=" + subjectBathroomTextbox.Text);
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Living_Square_Feet CONTENT=" + subjectAboveGlaTextbox.Text.Replace(",",""));
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                //default yes on form
                //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Basement CONTENT=%Yes");
                if (subjectBasementTypeTextbox.Text.Contains("None") | subjectBasementTypeTextbox.Text == "")
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Basement CONTENT=%No");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Basement_Finished CONTENT=%No");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Basement CONTENT=%No");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Basement_Finished CONTENT=NA");
                }
                else
                {
                    //there is a basement
                    //is it finished?
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Basement CONTENT=%Yes");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Basement_Finished CONTENT=Unknown");

                    if (!subjectBasementDetailsTextbox.Text.Contains("Finished"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Basement_Finished CONTENT=%No");
                    }
                }
                    
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Heating_Cooling CONTENT=GasFA/Central");

                //default yes/attached  on form
                if (subjectParkingTypeTextbox.Text.Contains("Gar"))
                {
                    if (subjectParkingTypeTextbox.Text.Contains("Det"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Garage_Type CONTENT=%Detached");
                    }

                }
                else
                {
                    //no garage
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Garage CONTENT=%NO");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Garage_Type CONTENT=%Parking<SP>Space");

                }
                
                
               //default on form yes
                if (subjectNumFireplacesTextbox.Text == "0" | subjectNumFireplacesTextbox.Text == "")
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Fireplace CONTENT=%No");
                }
                //design
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Design_Style CONTENT=" + subjectMlsTypecomboBox.Text.Replace(" ", "<SP>"));
             
                //
                //Pricing
                //

                //fnma
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Market_As_Is CONTENT=" + valueTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Repaired_As_Is CONTENT=" + valueTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Market_Suggested CONTENT=" + listTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/Repaired_Suggested CONTENT=" + listTextbox.Text);

                //

                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
macro.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
macro.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKETING_STRATEGY/Sale_As_Is_Value CONTENT=" + valueTextbox.Text);
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKETING_STRATEGY/Suggested_List_As_Is_Value CONTENT=" + listTextbox.Text);
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
              //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKETING_STRATEGY/Suggested_Repaired_Value CONTENT=" + listTextbox.Text);
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKETING_STRATEGY/Quick_Sale_Repaired_Value CONTENT=" + quickSaleTextbox.Text);
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                
       
            }
            #endregion

            #region eval
            if (currentUrl.ToLower().Contains("evalonline"))
            {

                string comments = "Searched the following: distance of at least 1 mile, gla +//- 20% sqft, lot size 30% +//- sq ft, and age 30% +//- yrs up to 12 months in time.   Results:  No other sales data that matched gla, lot size, age or condition were considered applicable in regards to distance to subject, 3 and 6 month date of sale parameters, 90 DOM requirement, and still be within 15% tolerance range.   The comparables selected were considered to be the best available.";


                //
                //page1
                //
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Subject<SP>Property<SP>Ownership<SP>Info");
                macro.AppendLine(@"WAIT SECONDS=1");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10100 CONTENT=" + DateTime.Now.Date.ToShortDateString());
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10102 CONTENT=%1");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10104 CONTENT=%1");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10200 CONTENT=%2");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10202 CONTENT=%3");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10214 CONTENT=%3");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10216 CONTENT=%3");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10218 CONTENT=" + subjectpin_textbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10220 CONTENT=%2");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10222 CONTENT=na");
     
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10224 CONTENT=na");
                
   
                //macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10226 CONTENT=GrassLakeRd");
               
                //macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10228 CONTENT=Woods");
            
                //macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10230 CONTENT=GrandAve");

                //macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10232 CONTENT=Lake");

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10250");
                
                //pud = no
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10262 CONTENT=%4");

                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10263 CONTENT=%Yes");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10269 CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10277 CONTENT=%No");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                
                //last sale info
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10264 CONTENT=" + subjectLastSaleDateTextbox.Text.Replace(" ", "<SP>"));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10265 CONTENT=" + subjectLastSalePriceTextbox.Text.Replace(" ", "<SP>"));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10266");
                
                //property listed
                //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10300 CONTENT=%1");
                macro.AppendLine(@"TAG POS=214 TYPE=TD FORM=ID:formInput ATTR=*");
                macro.AppendLine(@"TAG POS=164 TYPE=TD FORM=ID:formInput ATTR=*");
                //macro.AppendLine(@"TAG POS=19 TYPE=TD FORM=ID:formInput ATTR=CLASS:formLabelAlignRightRed");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10277 CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10278");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10300 CONTENT=%3");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:formInput ATTR=ID:submit&&VALUE:SAVE<SP>&<SP>VALIDATE");
                //end page 1

                //
                //page 2 - Subject info
                //
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Subject<SP>Property<SP>Details");
                macro.AppendLine(@"WAIT SECONDS=1");
                //age
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10000 CONTENT=" + Age(subjectYearBuiltTextbox.Text));
                //units
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10002 CONTENT=1");
                //att=1 or det=3
                if (subjectAttachedRadioButton.Checked)
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10004 CONTENT=%1");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10010 CONTENT=%7");
                }
                else
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10004 CONTENT=%3");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10010 CONTENT=%1");

                } 
                //mlstype
                switch (subjectMlsTypecomboBox.Text)
                {
                    case "1 Story" :
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10006 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10007 CONTENT=%Ranch");
                        break;
                    case "1.5 Story" :
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10006 CONTENT=%3");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10007 CONTENT=%Conventional");
                        break;
                    case "2 Stories":
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10006 CONTENT=%4");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10007 CONTENT=%Contemporary");
                        break;
                    case "Raised Ranch":
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10006 CONTENT=%9");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10007 CONTENT=%Raised<SP>Ranch");
                        break;
                    case "Split Level":
                    case @"Split Level w/Sub":
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10006 CONTENT=%25");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10007 CONTENT=%Split<SP>Level");
                        break;
                }
              

                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10008 CONTENT=%1");
              
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10011 CONTENT=%Framed");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10012 CONTENT=%1");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10016 CONTENT=SFR");
                //gla
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10018 CONTENT=" + subjectAboveGlaTextbox.Text.Replace(",",""));
                //data source tax = 1
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10020 CONTENT=%1");
                //room count
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10022 CONTENT=" + subjectRoomCountTextbox.Text); 
                //beds
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10024 CONTENT=" + subjectBedroomTextbox.Text);
                //full baths
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10026 CONTENT=" + subjectBathroomTextbox.Text[0]);
                //half baths
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10028 CONTENT=" + subjectBathroomTextbox.Text[2]);

                //basement
                if (SubjectBasementType.ToLower().Contains("full"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10030 CONTENT=%3");
                    if (subjectBasementDetailsTextbox.Text.ToLower().Contains("finished"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10032 CONTENT=%5");
                    }
                    else
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10032 CONTENT=%1");
                    }
                }
                else
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10030 CONTENT=%2");
                }
                
                //garage type and spaces
                //garage type 1 = attached, 4 = detached
                if (subjectParkingTypeTextbox.Text.ToLower().Contains("att"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10036 CONTENT=%1");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10038 CONTENT=" + Regex.Match(subjectParkingTypeTextbox.Text,@"\d"));
                }
                else if (subjectParkingTypeTextbox.Text.ToLower().Contains("det"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10036 CONTENT=%4");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10038 CONTENT=" + Regex.Match(subjectParkingTypeTextbox.Text, @"\d"));
                }
                else
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10036 CONTENT=%3");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10038 CONTENT=0");
                }
                try
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10040 CONTENT=" + (Convert.ToDecimal(subjectLotSizeTextbox.Text) * 43560).ToString());
                }
                catch 
                { 
                
                }
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10042 CONTENT=" + subjectLotSizeTextbox.Text);
                //macro.AppendLine(@"WAIT SECONDS=3");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10044 CONTENT=%3");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10046 CONTENT=%1");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10048 CONTENT=%2");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10050 CONTENT=%3");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10054 CONTENT=%Vinyl<SP>Siding");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10850 CONTENT=%3");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10852 CONTENT=%3");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10854 CONTENT=%4");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10858 CONTENT=%3");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10862 CONTENT=%10");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10864 CONTENT=%3");
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10856 CONTENT=" + comments.Replace(" ", "<SP>"));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:formInput ATTR=ID:submit&&VALUE:SAVE<SP>&<SP>VALIDATE");
                macro.AppendLine(@"'TAG POS=0 TYPE=SELECT ATTR=NAME:fieldid_10864");


                //
                //page 9  Subject Property Valuation
                //
                //Neighborhood info
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Subject<SP>Property<SP>Valuation");
                macro.AppendLine(@"WAIT SECONDS=1");

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10950 CONTENT=" + subjectNeighborhood.oldestHome.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10952 CONTENT=" + subjectNeighborhood.newestHome.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10954 CONTENT=" + subjectNeighborhood.medianAge.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10956 CONTENT=" + subjectNeighborhood.numberOfCompListings.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10958 CONTENT=" + subjectNeighborhood.minListPrice.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10960 CONTENT=" + subjectNeighborhood.highListPrice.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10962 CONTENT=" + subjectNeighborhood.numberOfSales.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10964 CONTENT=" + subjectNeighborhood.minSalePrice.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10966 CONTENT=" + subjectNeighborhood.maxSalePrice.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10968 CONTENT=" + subjectNeighborhood.numberOfShortSaleListings.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10970 CONTENT=" + subjectNeighborhood.medianSoldPrice.ToString());

                //pricing

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_11000 CONTENT=" + quickSaleTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_11002 CONTENT=" + quickSaleTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_11004 CONTENT=" + valueTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_11006 CONTENT=" + valueTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_11008 CONTENT=120");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_11010 CONTENT=" + valueTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_11012 CONTENT=" + valueTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_11014 CONTENT=5000");


                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_11050 CONTENT=%2");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_11100 CONTENT=" + (Convert.ToDecimal(quickSaleTextbox.Text) + 900).ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_11102 CONTENT=" + (Convert.ToDecimal(quickSaleTextbox.Text) + 900).ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_11104 CONTENT=" + listTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_11106 CONTENT=" + listTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_11108 CONTENT=" + listTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_11110 CONTENT=" + listTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_11112 CONTENT=" + subjectRentTextbox.Text);
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_11114 CONTENT=%First<SP>time<SP>Buyer");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_11116 CONTENT=%3");



                //
                //Page 10 - Subject Marketibility
                //
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Subject<SP>Marketability");
                macro.AppendLine(@"WAIT SECONDS=1");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_11200 CONTENT=%Average");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_11202 CONTENT=%OwnerOccupied");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:formInput ATTR=ID:fieldid_11204 CONTENT=80");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:formInput ATTR=ID:fieldid_11206 CONTENT=90");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:formInput ATTR=ID:fieldid_11208 CONTENT=10");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_11210 CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_11250 CONTENT=%Declining");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_11252 CONTENT=%Over<SP>Supplied");
                macro.AppendLine(@"TAG POS=62 TYPE=TD ATTR=*");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:formInput ATTR=ID:fieldid_11254 CONTENT=-1");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_11256 CONTENT=%3");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_11260 CONTENT=%3");
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:formInput ATTR=NAME:fieldid_11258 CONTENT=na");
                macro.AppendLine(@"");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_11262 CONTENT=%FullyDeveloped");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_11264 CONTENT=%Over75%");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_11300 CONTENT=%No");
                macro.AppendLine(@"TAG POS=83 TYPE=TD ATTR=*");
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:formInput ATTR=NAME:fieldid_11302 CONTENT=Strong<SP>investor<SP>market,<SP>will<SP>sell<SP>at<SP>right<SP>price.");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_11304 CONTENT=%Equal(Same)");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_11308 CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:formInput ATTR=NAME:fieldid_11312 CONTENT=None.");
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:formInput ATTR=NAME:fieldid_11316 CONTENT=None.");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_11314 CONTENT=%ListAsIs");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:formInput ATTR=ID:submit");


                //
                //Page 11 Exterior repairs
                //
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Exterior<SP>Repair<SP>Addendum");
                macro.AppendLine(@"WAIT SECONDS=1");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7000 CONTENT=%Average");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7004 CONTENT=%Average");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7008 CONTENT=%Average");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7012 CONTENT=%Average");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7016 CONTENT=%Average");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7020 CONTENT=%Average");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7025 CONTENT=%Average");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7029 CONTENT=%Not<SP>Applicable");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7037 CONTENT=%Unable<SP>to<SP>Determine");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7045 CONTENT=%Not<SP>Applicable");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7001 CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7005 CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7009 CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7013 CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7017 CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7021 CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7026 CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7030 CONTENT=%Not<SP>Applicable");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7034 CONTENT=%Not<SP>Applicable");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7038 CONTENT=%Not<SP>Applicable");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7046 CONTENT=%Not<SP>Applicable");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_7042 CONTENT=%Not<SP>Applicable");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_8050 CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_8052 CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_8054 CONTENT=%Not<SP>Applicable");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_8051 CONTENT=%No");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_8053 CONTENT=%Yes");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:formInput ATTR=NAME:fieldid_8054 CONTENT=%Not<SP>Applicable");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:formInput ATTR=ID:submit");



            }
            #endregion

            string macroCode = macro.ToString();
             status = iim2.iimPlayCode(macroCode, 60);
        }

        private void mls_search_button_Click(object sender, EventArgs e)
        {
            importSubjectInfoButton(sender, e);
        }

        public void GoogleApiCall(string s)
        {
            googleReqResRichTextBox.AppendText(s + "\r\n");
        }

        //public void StatusUpdate(string s)
        //{
        //    programStatusRichTextBox.AppendText(s + "\r\n");
        //}

        public string Get_Distance(string zip)
        {
            //Google distance matric webservice
            //http://maps.googleapis.com/maps/api/distancematrix/xml?origins=60050&destinations=60002&sensor=false&units=imperial
            //
            //string zip = "60085";

            string googlestr = "http://maps.googleapis.com/maps/api/distancematrix/xml?origins=60050&destinations=" + zip + "&sensor=false&units=imperial";

            // Create a request for the URL. 		
            WebRequest request = WebRequest.Create(googlestr);

            // Get the response.
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();


            // Get the stream containing content returned by the server.
            Stream dataStream = response.GetResponseStream();
            // Open the stream using a StreamReader for easy access.
            StreamReader reader = new StreamReader(dataStream);
            // Read the content.
            string responseFromServer = reader.ReadToEnd();
            // Display the content.
         
            //MessageBox.Show(responseFromServer);

            //<distance>
    //<value>42315</value>
    //<text>26.3 mi</text>
   //</distance>

            string pattern = "(\\d+.\\d+) mi";
            Match match = Regex.Match(responseFromServer, pattern);
            string distance = match.Groups[1].Value;

            //MessageBox.Show(distance);
            //Console.WriteLine(responseFromServer);
            // Cleanup the streams and the response.
            reader.Close();
            dataStream.Close();
            response.Close();

            return distance;

        }

        public string Get_Distance(string o, string d)
        {
            //Google distance matric webservice
            //http://maps.googleapis.com/maps/api/distancematrix/xml?origins=60050&destinations=60002&sensor=false&units=imperial
            //
            //string zip = "60085";

            string googlestr = "http://maps.googleapis.com/maps/api/distancematrix/xml?origins=" + o + "&destinations=" + d + "&sensor=false&units=imperial";

            // Create a request for the URL. 		
            WebRequest request = WebRequest.Create(googlestr);

            // Get the response.
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();


            // Get the stream containing content returned by the server.
            Stream dataStream = response.GetResponseStream();
            // Open the stream using a StreamReader for easy access.
            StreamReader reader = new StreamReader(dataStream);
            // Read the content.
            string responseFromServer = reader.ReadToEnd();
            // Display the content.

            //MessageBox.Show(responseFromServer);

            //<distance>
            //<value>42315</value>
            //<text>26.3 mi</text>
            //</distance>

            string pattern = "(\\d+.\\d*) [mf]";
            Match match = Regex.Match(responseFromServer, pattern);
            string distance = match.Groups[1].Value;
            if (responseFromServer.Contains(" ft</text>"))
            {
                distance = "0.05";
            }
            //MessageBox.Show(distance);
            //Console.WriteLine(responseFromServer);
            // Cleanup the streams and the response.
            reader.Close();
            dataStream.Close();
            response.Close();

            return distance;

        }

        private void importSubjectInfoButton(object sender, EventArgs e)
        {

            RealistReport realistSubject = new RealistReport(this);
            MlsReportDriver mlsSubject = new MlsReportDriver(this);
            // TownshipReport subjectTownshipRecord = new TownshipReport(this);

            //string[] pages = { "", "", "", "", "", "", "", "" };

            //to do:
            //get a list of pdfs
            //
            string[] filePaths = Directory.GetFiles(search_address_textbox.Text, "*.pdf");
            string filename = "";
            string currentText = "";

            IEnumerable<System.Windows.Forms.TextBox> query1 = this.groupBox1.Controls.OfType<System.Windows.Forms.TextBox>();

            foreach (System.Windows.Forms.TextBox t in query1)
            {
                t.Clear();
            }

            try
            {
                bpoCommentsTextBox.Clear();
                bpoCommentsTextBox.LoadFile(SubjectFilePath + "\\bpocomments.rtf");
            }
            catch
            {

            }


            if (File.Exists(SubjectFilePath + "\\subjectinfo.txt"))
            {
                #region load data from subjectinfo file

                //load other fields not displayed
                foreach (string f in filePaths)
                {
                    if (f.ToLower().Contains("realist") || f.ToLower().Contains("property detail report"))
                    {
                        StringBuilder pages = new StringBuilder();
                        PdfReader pdfReader = new PdfReader(f);
                        ITextExtractionStrategy strat = new PdfHelper.LocationTextExtractionStrategyEx();
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            pages.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, strat));
                        }

                        //string fullReport = "";
                        //foreach (string p in pages)
                        //{
                        //    fullReport = fullReport + p;
                        //}

                        realistSubject.GetSubjectInfo(pages.ToString());
                        GlobalVar.theSubjectProperty.myRealistReport = realistSubject;
                        this.statusTextBox.AppendText("Realist Loaded...");
                    }

                    if (f.ToLower().Contains("report-") || f.ToLower().Contains("connectmls"))
                    {
                        StringBuilder pages = new StringBuilder();
                        PdfReader pdfReader = new PdfReader(f);
                        ITextExtractionStrategy strat = new PdfHelper.LocationTextExtractionStrategyEx();
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            pages.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, strat));
                        }

                        mlsSubject.GetSubjectInfo(pages.ToString());

                        this.statusTextBox.AppendText("MLS listing Loaded...");
                    }
                }

                //overwrite with what is saved; leave the rest alone
                string line = "";
                string[] splitLine;
                using (System.IO.StreamReader file = new System.IO.StreamReader(this.SubjectFilePath + "\\" + "subjectinfo.txt"))
                {
                    while (!file.EndOfStream)
                    {
                        line = file.ReadLine();
                        splitLine = line.Split(';');


                        System.Windows.Forms.Control[] c = this.Controls.Find(splitLine[0], true);

                        if (c[0] is TextBox || c[0] is ComboBox)
                        {
                            c[0].Text = splitLine[1];
                        }

                        if (c[0] is System.Windows.Forms.RadioButton)
                        {
                            System.Windows.Forms.RadioButton t = (System.Windows.Forms.RadioButton)c[0];

                            t.Checked = Convert.ToBoolean(splitLine[1]);
                        }


                        //foreach (System.Windows.Forms.TextBox t in query1)
                        //{

                        //}                       
                    }
                    
                }

                #endregion
            }
            else
            {
                #region Load Date from Pdf files
                foreach (string f in filePaths)
                {
                    if (f.ToLower().Contains("realist") || f.ToLower().Contains("property detail report"))
                    {
                        StringBuilder pages = new StringBuilder();
                        PdfReader pdfReader = new PdfReader(f);
                        ITextExtractionStrategy strat = new PdfHelper.LocationTextExtractionStrategyEx();
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            pages.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, strat));
                        }

                        //string fullReport = "";
                        //foreach (string p in pages)
                        //{
                        //    fullReport = fullReport + p;
                        //}

                        realistSubject.GetSubjectInfo(pages.ToString());
                        GlobalVar.theSubjectProperty.myRealistReport = realistSubject;
                        this.statusTextBox.AppendText("Realist Loaded...");
                    }

                    if (f.ToLower().Contains("report-") || f.ToLower().Contains("connectmls"))
                    {
                        StringBuilder pages = new StringBuilder();
                        PdfReader pdfReader = new PdfReader(f);
                        ITextExtractionStrategy strat = new PdfHelper.LocationTextExtractionStrategyEx();
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            pages.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, strat));
                        }

                        mlsSubject.GetSubjectInfo(pages.ToString());
                      
                        this.statusTextBox.AppendText("MLS listing Loaded...");
                    }

                    if (f.ToLower().Contains("algonquin"))
                    {
                        StringBuilder pages = new StringBuilder();
                        PdfReader pdfReader = new PdfReader(f);
                        ITextExtractionStrategy strat = new PdfHelper.LocationTextExtractionStrategyEx();
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            pages.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, strat));
                        }

                        // AlgonquinTownshipReport ar = new AlgonquinTownshipReport(this, fullReport);
                        // subjectTownshipRecord = ar.DeepCopy();

                        subjectTownshipRecord = new AlgonquinTownshipReport(this, pages.ToString());
                        subjectBasementGlaTextbox.Text = subjectTownshipRecord.basementGla.ToString();
                        subjectFinishedBasementGlaTextBox.Text = subjectTownshipRecord.finishedBasementGla.ToString();
                        this.statusTextBox.AppendText("Assesor Record Loaded...");
                    }

                    if (f.ToLower().Contains("lakecountyil") || f.ToLower().Contains("lake county"))
                    {
                        StringBuilder pages = new StringBuilder();
                        PdfReader pdfReader = new PdfReader(f);
                        ITextExtractionStrategy strat = new PdfHelper.LocationTextExtractionStrategyEx();
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            pages.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, strat));
                        }


                        subjectTownshipRecord = new LakeCountyTownshipReport(this, pages.ToString());
                        subjectBasementGlaTextbox.Text = subjectTownshipRecord.basementGla.ToString();
                        subjectFinishedBasementGlaTextBox.Text = subjectTownshipRecord.finishedBasementGla.ToString();
                        this.statusTextBox.AppendText("Assesor Record Loaded...");
                    }

                    if (f.ToLower().Contains("mchenrytownship"))
                    {
                        StringBuilder pages = new StringBuilder();
                        PdfReader pdfReader = new PdfReader(f);
                        ITextExtractionStrategy strat = new PdfHelper.LocationTextExtractionStrategyEx();
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            pages.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, strat));
                        }


                        subjectTownshipRecord = new McHenryTownshipReport(this, pages.ToString());
                        subjectBasementGlaTextbox.Text = subjectTownshipRecord.basementGla.ToString();
                        subjectFinishedBasementGlaTextBox.Text = subjectTownshipRecord.finishedBasementGla.ToString();
                        this.statusTextBox.AppendText("Assesor Record Loaded...");
                    }

                    if (f.ToLower().Contains("elgin") && f.ToLower().Contains("township"))
                    {
                        StringBuilder pages = new StringBuilder();
                        PdfReader pdfReader = new PdfReader(f);
                        ITextExtractionStrategy strat = new PdfHelper.LocationTextExtractionStrategyEx();
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            pages.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, strat));
                        }


                        subjectTownshipRecord = new ElginTownshipReport(this, pages.ToString());
                        subjectBasementGlaTextbox.Text = subjectTownshipRecord.basementGla.ToString();
                        subjectFinishedBasementGlaTextBox.Text = subjectTownshipRecord.finishedBasementGla.ToString();
                        this.statusTextBox.AppendText("Assesor Record Loaded...");
                    }







                }
      

                try
                {
                    //
                    //check for missing data
                    //
                    if (subjectAboveGlaTextbox.Text == "0")
                    {
                        subjectAboveGlaTextbox.Text = subjectTownshipRecord.aboveGradeGla.ToString();
                    }

                    if (subjectYearBuiltTextbox.Text == "")
                    {
                        subjectYearBuiltTextbox.Text = subjectTownshipRecord.yearBuilt.ToString();
                    }
                }
                catch
                {
                    //no township record loaded.
                }

                subjectTownshipTextBox.Text = realistSubject.Township;
                subjectTaxAmountTextBox.Text = GlobalVar.theSubjectProperty.PropertyTax;

                #endregion
            }

            try
            {
                decimal dec;
                Decimal.TryParse(ndAvgDomTextBox.Text, out  dec );
                SubjectNeighborhood.avgDom = Decimal.ToInt32(dec);
                SubjectNeighborhood.minSalePrice = Convert.ToDecimal(ndMinSalePriceTextBox.Text);
                SubjectNeighborhood.maxSalePrice = Convert.ToDecimal(ndMaxSalePriceTextBox.Text);
                SubjectNeighborhood.numberActiveListings = Convert.ToInt32(ndNumberOfActiveListingTextBox.Text);
                SubjectNeighborhood.numberREOListings = Convert.ToInt32(ndNumberActiveReoListingsTextBox.Text);
            }
            catch (Exception ex)
            {

            }
            SubjectBrokerPhone = subjectBrokerPhoneTextBox.Text;
            SubjectListingAgent = subjectListingAgentTextBox.Text;
            SubjectCurrentListPrice = subjectCurrentListPriceTextBox.Text;
            SubjectMlsStatus = subjectMlsStatusTextBox.Text;

            SubjectAssessmentInfo.managementContact = subjectHoaContactTextBox.Text;
            SubjectAssessmentInfo.frequency = subjectHoaFrequencyTextBox.Text; 
             SubjectAssessmentInfo.includes = subjectHoaIncludesTextBox.Text;
            SubjectAssessmentInfo.managementPhone =subjectHoaPhoneTextBox.Text  ;
            SubjectAssessmentInfo.amount = subjectHoaTextBox.Text;
            SubjectDom = subjectDomTextBox.Text;


            subjectProximityToOfficeTextbox.Text = Get_Distance(subjectFullAddressTextbox.Text);

            if (!string.IsNullOrEmpty(subjectFullAddressTextbox.Text))
            {
                GlobalVar.subjectPoint = Geocode(subjectFullAddressTextbox.Text);
            }



            GlobalVar.theSubjectProperty.County = SubjectCounty;

            GlobalVar.theSubjectProperty.MainForm = this;



            //DataTable subject_table = subjectTableAdapter.GetData();
            //MessageBox.Show(realist_lot_acres + " " + pin + " " + realist_percent_improved + " " + realist_school_district + " " + realist_census_tract + " " + realist_carrier_route + " " + realist_subdivision);

            //try
            //{
            //    int x = this.subjectTableAdapter.InsertQuery(pin, "", realist_census_tract, address, city, county, realist_township, null, null, null, null);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}

        }
       
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            decimal lotSize = 0;
            try
            {
                lotSize = Convert.ToDecimal(subjectLotSizeTextbox.Text.Replace(",", ""));

            }
            catch (Exception)
            {
                // throw;
            }

            lowLotSizetextBox.Text = (lotSize * (1- lotSizeUpDown.Value)).ToString();
            highLotSizetextBox.Text = (lotSize * (1 + lotSizeUpDown.Value)).ToString();
        }

        private void runMredScriptButton_Click(object sender, EventArgs e)
        {
            recheckLastSearchbutton.Enabled = false;
            MRED_Driver d = new MRED_Driver(this);
            iMacros.App mred = new iMacros.App();

           // GoogleFusionTable gt = new GoogleFusionTable("1cU8NHmtVbE3KWpGa2qSTkvIpwhC5KIJX3hcvlIg");

            //using a known shared table to avoid access issues to start auth
            //below key is for bpo table 
            //GoogleFusionTable gt = new GoogleFusionTable("15O-Cw9hrgob0ETCDM8Fk33qc_B-GtPdL84COXZ8");

            iim.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string iimCurrentUrl = iim.iimGetLastExtract();
            string iim2CurrentUrl;

            if (iimCurrentUrl.ToLower().Contains("connectmls"))
            {
                mred = iim;
            }
            else
            {
                iim2.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
                iim2CurrentUrl = iim2.iimGetLastExtract();
                if (iim2CurrentUrl.ToLower().Contains("connectmls"))
                {
                    mred = iim2;
                }
                else
                {
                    //neither are on mred, now what??
                    //is one on the imacro start page?
                    if (iimCurrentUrl.Contains(@"iOpus/iMacros/Start.html"))
                    {
                        mred = iim;
                        iim.iimPlay(@"MRED\#start_mls.iim"); 
                    }
                    else if (iim2CurrentUrl.Contains(@"iOpus/iMacros/Start.html"))
                    {
                        mred = iim2;
                    }
                    else
                    {
                        MessageBox.Show("No open or connected MLS.");
                        return;
                    }
                }
            }
             
            switch (comboBox2.Text)
            {
                case "New Map Search":
                    d.NewMapSearch(mred);
                    break;
                case "Save 1 MI Stats":
                    d.Save1MiSearch(mred);
                    break;
                case "SaveStats":
                    d.SaveStats(mred, oneMile);
                    break;
                case "FindComps":
                    d.FindComps(mred, true);
                    recheckLastSearchbutton.Enabled = true;
                    break;
                case "FindComps - Current Search":
                    d.FindComps(mred, false);
                    recheckLastSearchbutton.Enabled = true;
                    break;
                default:
                    //Console.WriteLine("Invalid selection. Please select 1, 2, or 3.");
                    break;
            }
            infoWindowrichTextBox.SaveFile(SubjectFilePath + "\\" + "searchSummary_list.rtf");


            //
            //Set neighborhood data fields
            //
            try
            {

                ndAvgDomTextBox.Text = SubjectNeighborhood.avgDom.ToString();
                ndMaxSalePriceTextBox.Text = SubjectNeighborhood.maxSalePrice.ToString();
                ndMinSalePriceTextBox.Text = SubjectNeighborhood.minSalePrice.ToString();
                ndNumberOfActiveListingTextBox.Text = SubjectNeighborhood.numberActiveListings.ToString();
                ndNumberActiveReoListingsTextBox.Text = SubjectNeighborhood.numberREOListings.ToString();
                ndNumberActiveShortListingsTextBox.Text = SubjectNeighborhood.numberOfShortSaleListings.ToString();
                ndNumberOfSoldListingTextBox.Text = SubjectNeighborhood.numberOfSales.ToString();
                ndNumberSoldReoListingsTextBox.Text = SubjectNeighborhood.numberREOSales.ToString();
                ndNumberSoldShortListingsTextBox.Text = SubjectNeighborhood.numberShortSales.ToString();

            }
            catch
            {

            }

        }

        private void textBox2_TextChanged_2(object sender, EventArgs e)
        {

            try
            {
                this.listTextbox.Text = (Math.Round((Convert.ToInt64(valueTextbox.Text) * 1.03))).ToString();
                this.quickSaleTextbox.Text = (Math.Round((double)(((Convert.ToInt64(valueTextbox.Text) * Convert.ToInt64(quickSalePercentageTextBox.Text)))))).ToString();
            }
            catch
            {
                try
                {
                    this.quickSaleTextbox.Text = (Math.Round(Convert.ToInt64(valueTextbox.Text) * .85)).ToString();
                }
                catch
                {
                    this.quickSaleTextbox.Text = "";
                }
            }

        }

        private void button12_Click(object sender, EventArgs e)
        {
            //this change image from 1.8MB to 74kb
            Bitmap bmp = new Bitmap(@"C:\Users\Scott\Documents\My Dropbox\BPOs\510 Forest Glen\SAM_9047.JPG");
            Bitmap bmp2 = new Bitmap(bmp, 640,480);
            bmp2.Save(@"C:\Users\Scott\Documents\My Dropbox\BPOs\510 Forest Glen\small_SAM_9047.JPG", System.Drawing.Imaging.ImageFormat.Jpeg);
        }

        private void subjectRentTextbox_TextChanged(object sender, EventArgs e)
        {
        http://www.zillow.com/webservice/GetSearchResults.htm?zws-id=<ZWSID>&address=2114+Bigelow+Ave&citystatezip=Seattle%2C+WA

            string[] address = subjectFullAddressTextbox.Text.Split(',');

        try
        {
            string googlestr = @"http://www.zillow.com/webservice/GetSearchResults.htm?zws-id=X1-ZWz1degk2gu3gr_799rh&address=" + address[0].Replace(" ", "+") + "&citystatezip=" + address[1].Replace(" ", "+") + "%2C," + address[2].Replace(" ", "+") + "&rentzestimate=true";
            //// Create a request for the URL. 		
            WebRequest request = WebRequest.Create(googlestr);

            //// Get the response.
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();


            // Get the stream containing content returned by the server.
            Stream dataStream = response.GetResponseStream();
            // Open the stream using a StreamReader for easy access.
            StreamReader reader = new StreamReader(dataStream);
            // Read the content.
            string responseFromServer = reader.ReadToEnd();
            // Display the content.
            //MessageBox.Show(response.ContentLength.ToString());
            //MessageBox.Show(responseFromServer);
            ////Console.WriteLine(responseFromServer);
            //// Cleanup the streams and the response.
            reader.Close();
            dataStream.Close();
            response.Close();


            string pattern = @"<rentzestimate><amount currency=.USD.>(\d+)</amount>";
            Match match = Regex.Match(responseFromServer, pattern);
            string rent = match.Groups[1].Value;
            subjectRentTextbox.Text = rent;

            //<zestimate><amount currency="USD">127222</amount>
            pattern = @"<zestimate><amount currency=.USD.>(\d+)</amount>";
            match = Regex.Match(responseFromServer, pattern);
            string zest = match.Groups[1].Value;
            subjectZestimateTextbox.Text = zest;
        }
        catch
        {

        }
            


        }

        private void subjectAttachedRadioButton_CheckedChanged(object sender, EventArgs e)
        {

            if (subjectAttachedRadioButton.Checked)
            {
                this.LoadMlsTypeList(new TypeAttachedList());
            }
       
        }

        private void subjectDetachedradioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (subjectDetachedradioButton.Checked)
            {
                this.LoadMlsTypeList(new TypeDetachedList());
            }
        }

        private void streetnumTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //infoWindowrichTextBox.SaveFile(this.SubjectFilePath + "\\" + "comp_list.rtf");
            //try
            //{
            //    IEnumerable<System.Windows.Forms.TextBox> query1 = this.groupBox1.Controls.OfType<System.Windows.Forms.TextBox>();


            //    using (System.IO.StreamWriter file = new System.IO.StreamWriter(this.SubjectFilePath + "\\" + "subjectinfo.txt"))
            //    {
            //        foreach (System.Windows.Forms.TextBox t in query1)
            //        {
            //            file.WriteLine("{0};{1}", t.Name, t.Text);
            //        }
            //    }
            //}
            //catch
            //{
            //}

        
        }

        private void saveSubjectInfoButton_Click(object sender, EventArgs e)
        {
            IEnumerable<System.Windows.Forms.TextBox> query1 = this.groupBox1.Controls.OfType<System.Windows.Forms.TextBox>();
            IEnumerable<System.Windows.Forms.TextBox> query2 = this.neighborhoodDataGroupBox.Controls.OfType<System.Windows.Forms.TextBox>();

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(this.SubjectFilePath + "\\" + "subjectinfo.txt"))
            {
                foreach (System.Windows.Forms.TextBox t in query1)
                {
                    file.WriteLine("{0};{1}", t.Name, t.Text);
                }
                foreach (System.Windows.Forms.TextBox t in query2)
                {
                    file.WriteLine("{0};{1}", t.Name, t.Text);
                }
                
               file.WriteLine("{0};{1}", subjectDetachedradioButton.Name, subjectDetachedradioButton.Checked.ToString());
               file.WriteLine("{0};{1}", subjectAttachedRadioButton.Name, subjectAttachedRadioButton.Checked.ToString());
               file.WriteLine("{0};{1}", subjectMlsTypecomboBox.Name, subjectMlsTypecomboBox.Text);
               bpoCommentsTextBox.SaveFile(this.SubjectFilePath + "\\" + "bpocomments.rtf");

            }
        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void search_address_textbox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //search_address_textbox.Text = DropBoxFolder;
            Process.Start(SubjectFilePath);


        }

        private void subjectBathroomTextbox_Validating(object sender, CancelEventArgs e)
        {
            if (!subjectBathroomTextbox.Text.Contains("."))
            {
                subjectBathroomTextbox.AppendText(".0");
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            var form2 = new equiTraxView();
            form2.pin = subjectpin_textbox.Text;
            form2.style = subjectMlsTypecomboBox.Text;
            form2.Visible = true;
            
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            GlobalVar.listingsFromLastSearch.Clear();
            search_address_textbox.Text = comboBox3.Text;
            importSubjectInfoButton(sender, e);
        }

        private void statusTextBox_TextChanged(object sender, EventArgs e)
        {
            // search_address_textbox
            //everyones dropbox folder is different, hence this helper function
            //DropBoxFolder

            string baseDir = DropBoxFolder + @"\BPOs\";
            string[] dirs = Directory.GetDirectories(baseDir);
            string mostRecentAccess;
            TimeSpan ts = new TimeSpan();



            foreach (string d in dirs)
            {
                ts = Directory.GetLastAccessTime(d) - DateTime.Now;
                if (ts.Days > -5)
                {
                    comboBox3.Items.Add(d);
                }

            }
        }

        private void subjectAboveGlaTextbox_TextChanged(object sender, EventArgs e)
        {
             int gla = 0;
            try
            {
                gla = Convert.ToInt32(subjectAboveGlaTextbox.Text.Replace(",",""));
            }
            catch (Exception)
            {
               // throw;
            }
          
            lowGlaTextBox.Text = (gla * .85).ToString("f0");
            highGlaTextBox.Text = (gla * 1.15).ToString("f0");
        }

        private void comboBox3_Click(object sender, EventArgs e)
        {
            
        }

        private void glaUpDown_ValueChanged(object sender, EventArgs e)
        {
            int gla = 0;
            try
            {
                gla = Convert.ToInt32(subjectAboveGlaTextbox.Text.Replace(",", ""));
            }
            catch (Exception)
            {
                // throw;
            }

            lowGlaTextBox.Text = (gla * Math.Abs((1 - glaUpDown.Value))).ToString("f0");
            highGlaTextBox.Text = (gla * (1 + glaUpDown.Value)).ToString("f0");
        }

        private void lotSizeUpDown_ValueChanged(object sender, EventArgs e)
        {
            decimal lotSize = 0;
            try
            {
                lotSize = Convert.ToDecimal(subjectLotSizeTextbox.Text.Replace(",", ""));
            }
            catch (Exception)
            {
                // throw;
            }
            
            lowLotSizetextBox.Text = (lotSize * (1-lotSizeUpDown.Value)).ToString();
            highLotSizetextBox.Text = (lotSize * (1 + lotSizeUpDown.Value)).ToString();
        }

        private void useCustomradioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (useCustomradioButton.Checked) 
                GlobalVar.ccc = GlobalVar.CompCompareSystem.USER;
        }

        private void useNabpopradioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (useCustomradioButton.Checked)
                GlobalVar.ccc = GlobalVar.CompCompareSystem.NABPOP;
        }

        private void highGlaTextBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                GlobalVar.upperGLA = Convert.ToDecimal(highGlaTextBox.Text);
            }
            catch
            {

            }

        }

        private void lowGlaTextBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                GlobalVar.lowerGLA = Convert.ToDecimal(lowGlaTextBox.Text);
            }
            catch
            {

            }
        }

        private void highLotSizetextBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                GlobalVar.upperLotSize = Convert.ToDecimal(highLotSizetextBox.Text);
            }
            catch
            {

            }
        }

        private void lowLotSizetextBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                GlobalVar.lowerLotSize = Convert.ToDecimal(lowLotSizetextBox.Text);
            }
            catch
            {

            }
        }

        private void subjectBasementDetailsTextbox_DoubleClick(object sender, EventArgs e)
        {
            BasementForm BasementForm = new BasementForm(this);
            BasementForm.Visible = true;
        }

        private void subjectBasementTypeTextbox_TextChanged(object sender, EventArgs e)
        {
           

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void recheckLastSearchbutton_Click(object sender, EventArgs e)
        {
             MRED_Driver d = new MRED_Driver(this);
             d.FindComps(iim, false);
        }

        private void button16_Click(object sender, EventArgs e)
        {
            StringBuilder output = new StringBuilder();

            String xmlString =
                    @"<FORMINFO FORMNUM='7001BOADB' FILENUM='27270824' CASE_NO='19539340' DOCID='20130207-02681-1' FORMVERSION='6-2010' MAINFORM='7001BOADB' VENDOR='CDNA'>
  <SUBJECT>
    <LOANNUM>27270824</LOANNUM>
    <SERVICERLOANNUM>27270824</SERVICERLOANNUM>
    <APPRAISAL>
      <PURPOSE>
        <RESPONSE>SERVICING</RESPONSE>
        <DESCRIPTION>
        </DESCRIPTION>
      </PURPOSE>
    </APPRAISAL>
    <BORROWER>JOE L MCKENZIE</BORROWER>
    <ADDR>
      <STREET>317  BRIERHILL DR</STREET>
      <CITY>ROUND LAKE</CITY>
      <STATEPROV>IL</STATEPROV>
      <ZIP>60073-3429</ZIP>
    </ADDR>
    <UNITNUM>
    </UNITNUM>
    <COUNTY>Lake</COUNTY>
    <PROPTYPE>SFR</PROPTYPE>
    <HOMEOWNERASSNFEE>
    </HOMEOWNERASSNFEE>
    <SOLDLISTED VALUE='NO' />
    <CURRENTOCCUPANT>OWNER</CURRENTOCCUPANT>
  </SUBJECT>
  <LENDERCLIENT>
    <COMPANY>
      <NAME>Bank of America-Executive</NAME>
    </COMPANY>
    <ADDR>
      <STREET>7105 Corporate Dr</STREET>
      <STREET2>
      </STREET2>
      <CITY>Plano</CITY>
      <STATEPROV>TX</STATEPROV>
      <ZIP>75024-4100</ZIP>
    </ADDR>
  </LENDERCLIENT>
  <BROKER>
    <COMPANY>
      <NAME>O.K. &amp; Associates, Realty Plus</NAME>
    </COMPANY>
    <NAME>Scott Beilfuss</NAME>
    <PHONE>815-315-0203</PHONE>
  </BROKER>
  <LISTINGCOMP>
    <SUBJECT>
      <ADDR>
        <STREET>317  BRIERHILL DR</STREET>
        <CITY>ROUND LAKE</CITY>
        <STATEPROV>IL</STATEPROV>
        <ZIP>60073-3429</ZIP>
      </ADDR>
      <ORIGINALLISTINGDATE>n/a</ORIGINALLISTINGDATE>
      <ORIGINALPRICE>n/a</ORIGINALPRICE>
      <LISTINGPRICE>n/a</LISTINGPRICE>
      <MARKETDAYS>n/a</MARKETDAYS>
    </SUBJECT>
    <REPAIR>
      <TOTALCOST>
      </TOTALCOST>
    </REPAIR>
    <REPAIR NUM='1'>
      <ESTCOST>0</ESTCOST>
    </REPAIR>
    <REPAIR NUM='2'>
      <ESTCOST>0</ESTCOST>
    </REPAIR>
    <REPAIR NUM='3'>
      <ESTCOST>0</ESTCOST>
    </REPAIR>
    <REPAIR NUM='4'>
      <ESTCOST>0</ESTCOST>
    </REPAIR>
    <REPAIR NUM='5'>
      <ESTCOST>0</ESTCOST>
    </REPAIR>
    <REPAIR NUM='6'>
      <ESTCOST>0</ESTCOST>
    </REPAIR>
  </LISTINGCOMP>
  <SALESCOMP>
    <SUBJECT>
      <ADDR>
        <STREET>317  BRIERHILL DR</STREET>
        <CITY>ROUND LAKE</CITY>
        <STATEPROV>IL</STATEPROV>
        <ZIP>60073-3429</ZIP>
      </ADDR>
      <NUMUNITS>1</NUMUNITS>
    </SUBJECT>
  </SALESCOMP>
  <REVIEW>
    <ANALYSIS>
      <NETADJUSTMENT>
      </NETADJUSTMENT>
      <GROSSADJUSTMENT>
      </GROSSADJUSTMENT>
    </ANALYSIS>
    <REPORTSECTION>
      <IMPROVEMENTS>
        <INTERIOR>No adverse conditions were noted at the time of inspection based on exterior observations. No repairs noted from drive-by.</INTERIOR>
      </IMPROVEMENTS>
      <NBRHOOD>
        <COMMENTS>#closed: 79 #active: 74 #pending: 5
Average/Mean $/Above GLA, sale: 59.56 Median $/Above GLA, sale: 58.84
Min dom: 2, Max dom: 1061, Average dom: 140, Median dom: 69
Min Age: 1, Max Age: 113, Average Age: 46, Median Age: 43
REO Sold: 53, REO Active: 16, Short Sold: 11, Short Active: 41</COMMENTS>
      </NBRHOOD>
    </REPORTSECTION>
  </REVIEW>
  <COMMENT>
    <CONDITIONIMPROVE>No adverse conditions were noted at the time of inspection based on exterior observations. No repairs noted from drive-by.</CONDITIONIMPROVE>
  </COMMENT>
  <SITE>
    <SALESHISTORY>
      <RESPONSE>NO</RESPONSE>
    </SALESHISTORY>
  </SITE>
  <NBRHOOD>
    <PROPVALUES>STABLE</PROPVALUES>
    <PROPSIMILARITY>SOMEWHATSIMILAR</PROPSIMILARITY>
    <BUILTUP>25-75%</BUILTUP>
    <SINGLEFAM>
      <PRICE>
        <LOW>17500</LOW>
        <HIGH>215000</HIGH>
      </PRICE>
    </SINGLEFAM>
    <LOCATION>SUBURBAN</LOCATION>
  </NBRHOOD>
  <IMPROVEMENTS>
    <EVIDENCEREPAIRS VALUE='NO' />
    <GENERAL>
      <UNITS>
        <NUMRESPONSE>
        </NUMRESPONSE>
      </UNITS>
    </GENERAL>
    <RATINGS>
      <INTERIORAPPEAL>AVERAGE</INTERIORAPPEAL>
    </RATINGS>
  </IMPROVEMENTS>
  <PROJECTANALYSIS>
    <VALUETREND>
      <MKTAREA>
        <PROPVALUES>
        </PROPVALUES>
        <SOURCE>MLS</SOURCE>
      </MKTAREA>
    </VALUETREND>
  </PROJECTANALYSIS>
</FORMINFO>";

            // Create an XmlReader
            using (XmlReader reader = XmlReader.Create(new StringReader(xmlString)))
            {
                XmlWriterSettings ws = new XmlWriterSettings();
                ws.Indent = true;
                using (XmlWriter writer = XmlWriter.Create(output, ws))
                {

                    // Parse the file and display each of the nodes.
                    while (reader.Read())
                    {
                        switch (reader.NodeType)
                        {
                            case XmlNodeType.Element:
                                writer.WriteStartElement(reader.Name);
                                if (reader.Name == "PROPTYPE")
                                {
                                     
                                    writer.WriteValue("GGG");
                                }
                                break;
                            case XmlNodeType.Text:
                                writer.WriteString(reader.Value);
                                break;
                            case XmlNodeType.XmlDeclaration:
                            case XmlNodeType.ProcessingInstruction:
                                writer.WriteProcessingInstruction(reader.Name, reader.Value);
                                break;
                            case XmlNodeType.Comment:
                                writer.WriteComment(reader.Value);
                                break;
                            case XmlNodeType.EndElement:
                                writer.WriteFullEndElement();
                                break;
                        }
                    }

                }
            }
            //infoWindowrichTextBox.Text = output.ToString();

            
            // Create an XmlReader
            using (XmlReader reader = XmlReader.Create(new StringReader(xmlString)))
            {
                reader.ReadToFollowing("BORROWER");
                output.AppendLine("Content of the title element: " + reader.ReadElementContentAsString());
                

                reader.ReadToFollowing("CURRENTOCCUPANT");
                output.AppendLine("Content of the title element: " + reader.ReadElementContentAsString());
            }

            infoWindowrichTextBox.Text = output.ToString();

            XmlDocumentSample.MyMethod();

           

            


//            StringBuilder output = new StringBuilder();

//            String xmlString =
//                @"<bookstore>
//        <book genre='autobiography' publicationdate='1981-03-22' ISBN='1-861003-11-0'>
//            <title>The Autobiography of Benjamin Franklin</title>
//            <author>
//                <first-name>Benjamin</first-name>
//                <last-name>Franklin</last-name>
//            </author>
//            <price>8.99</price>
//        </book>
//    </bookstore>";

//            // Create an XmlReader
//            using (XmlReader reader = XmlReader.Create(new StringReader(xmlString)))
//            {
//                reader.ReadToFollowing("book");
//                reader.MoveToFirstAttribute();
//                string genre = reader.Value;
//                output.AppendLine("The genre value: " + genre);

//                reader.ReadToFollowing("title");
//                output.AppendLine("Content of the title element: " + reader.ReadElementContentAsString());
//            }

//            OutputTextBlock.Text = output.ToString();




            LandSafeBpo b = new LandSafeBpo();

            b.SubjectPropertyType = LandSafeSyntax.PropertyType.SFR.ToString() ;
            b.WriteEnvFile();




            //   LandSafeBpo.PropertyType.SFR




        }

        private void subjectpin_textbox_TextChanged(object sender, EventArgs e)
        {
            if (GlobalVar.theSubjectProperty.ParcelID != (subjectpin_textbox.Text))
            {
                GlobalVar.theSubjectProperty = new SubjectProperty(subjectpin_textbox.Text, this);
            }
        
        }

        private void subjectpin_textbox_Validated(object sender, EventArgs e)
        {
  
        }

        private void subjectTaxAmountTextBox_MouseClick(object sender, MouseEventArgs e)
        {
            subjectTaxAmountTextBox.Text = GlobalVar.theSubjectProperty.PropertyTax;
        }

        private void subjectSubdivisionTextbox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            subjectSubdivisionTextbox.Text = GlobalVar.theSubjectProperty.Subdivision;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"TAG POS=1 TYPE=IMG ATTR=ID:myFavs");
            macro.AppendLine(@"FRAME NAME=workspace");
            macro.AppendLine(@"TAG POS=4 TYPE=SPAN FORM=NAME:dc ATTR=CLASS:columnHeader");
            macro.AppendLine(@"FRAME F=0");
            macro.AppendLine(@"FRAME NAME=workspace");
            macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:" + SubjectFullAddress.Substring(0,10) + "*");
            macro.AppendLine(@"TAG POS=22 TYPE=TD FORM=NAME:dc ATTR=TXT:* EXTRACT=TXT");
            macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:{{!EXTRACT}}");

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 20);
        }

        private void subjectLastSaleDateTextbox_TextChanged(object sender, EventArgs e)
        {
            GlobalVar.theSubjectProperty.DateOfLastSale = subjectLastSaleDateTextbox.Text;
        }

        private void subjectTaxAmountTextBox_TextChanged(object sender, EventArgs e)
        {
            GlobalVar.theSubjectProperty.PropertyTax = subjectTaxAmountTextBox.Text;
            
        }

        
        
        private void button18_Click(object sender, EventArgs e)
        {
           
           
            if (GlobalVar.listingsFromLastSearch.Count == 0)
            {

                // Create an instance of the XmlSerializer.
                System.Xml.Serialization.XmlSerializer serializer =
                new System.Xml.Serialization.XmlSerializer(typeof(List<MLSListing>));
                // Reading the XML document requires a FileStream.
                //Stream reader = new FileStream(@"E:\Dropbox\BPOs\1043 Spafford St\listingFromLastSearch3.xml", FileMode.Open);
                Stream reader = new FileStream(SubjectFilePath + @"\listingFromLastSearch.xml", FileMode.Open);
                // Declare an object variable of the type to be deserialized.
                List<MLSListing> i;

               

                // Call the Deserialize method to restore the object's state.
                i = (List<MLSListing>)serializer.Deserialize(reader);

                reader.Close();

                foreach (MLSListing m in i)
                {
                    m.ReloadValues();
                }
                GlobalVar.listingsFromLastSearch = i;
                MessageBox.Show("Loaded " + i.Count.ToString());

            }
       

           // i[0].Fields[MLSListing.ListingSheetFieldName.TypeDetached].value;

           // dataGridView1.DataSource = i;

         

            //var queryListings = from l in GlobalVar.listingsFromLastSearch
            //                    where l.GLA < .25 && l.PropertyType() == SubjectMlsType
            //                    select l;

            //var queryListings = from l in GlobalVar.listingsFromLastSearch
            //                    where l.proximityToSubject < .25 && l.PropertyType() == SubjectMlsType
            //                    select l;
            
            //dataGridView1.DataSource = queryListings.ToList();




            ////DataTable table = queryListings.CopyToDataTable();
            
            

            //var myR = GlobalVar.listingsFromLastSearch.Where(p => p.PropertyType() == "1 Story");

            //MessageBox.Show(myR.Count().ToString());

          
          //  ParameterExpression pe = Expression.Parameter(typeof(string), "company");
            //Expression left = Expression.Call(pe, typeof(string).GetMethod("ToLower", System.Type.EmptyTypes));
            //Expression right = Expression.Constant("coho winery");
            //Expression e1 = Expression.Equal(left, right);

            // Combine the expression trees to create an expression tree that represents the 
            // expression '(company.ToLower() == "coho winery" || company.Length > 16)'.
            //Expression predicateBody = Expression.OrElse(e1, e2);
            //int x = 0;
            //int y = 0;
            //Int32.TryParse(lowGlaTextBox.Text, out x);
            //Int32.TryParse(highGlaTextBox.Text, out y);
            
            
            
            ////dataGridView1.DataSource = results.ToList();
            ////MessageBox.Show("There are " + results.Count().ToString() + " of type " + comboBox4.Text);
            ////var l = GlobalVar.listingsFromLastSearch.Where(p => p.GLA >= x && p.GLA <= y);
            //Double d = -1;
            //Double.TryParse(ccProximityTextBox.Text, out d);
            //var l = GlobalVar.listingsFromLastSearch.Where(p => p.ProximityToSubject <= d);
            dataGridView1.DataSource = GlobalVar.listingsFromLastSearch;
            GlobalVar.listingsSearchResults = GlobalVar.listingsFromLastSearch;
            //MessageBox.Show("There are " + l.Count().ToString() + " of comparable GLA");

          

        }

        private void label38_Click(object sender, EventArgs e)
        {

        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (e.NewValue == CheckState.Unchecked)
            {
           
                checkedListBox1.ClearSelected();
              //  checkedListBox1.Refresh();
                checkedListBox1.Update();
                dataGridView1.DataSource = GlobalVar.listingsFromLastSearch;
                GlobalVar.listingsSearchResults = GlobalVar.listingsFromLastSearch;
            }
            else
            {
                

                CheckedListBox.ObjectCollection cb = checkedListBox1.Items;

                if (e.Index == checkedListBox1.FindString("PriceRange"))
                {
                    Double lo;
                    Double hi;
                    Double target;

                    Double.TryParse(ccTargetPriceTextBox.Text, out target);
                    lo = .5 * target;
                    hi = 1.5 * target;

                    var m = GlobalVar.listingsFromLastSearch.Where(p => p.SalePrice >= lo && p.SalePrice <= hi);

                    if (dataGridView1.RowCount == 0)
                    {

                        GlobalVar.listingsSearchResults = m;
                        dataGridView1.DataSource = m.ToList();
                    }
                    else
                    {

                        m = GlobalVar.listingsSearchResults.Where(p => p.SalePrice >= lo && p.SalePrice <= hi);
                        GlobalVar.listingsSearchResults = m;
                        dataGridView1.DataSource = m.ToList();
                    }

                    MessageBox.Show("There are " + m.Count().ToString() + " within price range:" + lo.ToString() + " to " + hi.ToString());

                }

                if (e.Index == checkedListBox1.FindString("Closed"))
                {


                    var m = GlobalVar.listingsFromLastSearch.Where(p => p.Status == "CLSD" && p.SalesDate <= DateTime.Now.AddMonths(Convert.ToInt32(-1 * ccMonthsBacknumericUpDown.Value)));

                    if (dataGridView1.RowCount == 0)
                    {

                        GlobalVar.listingsSearchResults = m;
                        dataGridView1.DataSource = m.ToList();
                    }
                    else
                    {

                        m = GlobalVar.listingsSearchResults.Where(p => p.Status == "CLSD" && p.SalesDate <= DateTime.Now.AddMonths(Convert.ToInt32(-1 * ccMonthsBacknumericUpDown.Value)));
                        
                        GlobalVar.listingsSearchResults = m;
                        dataGridView1.DataSource = m.ToList();
                    }
                    MessageBox.Show("There are " + m.Count().ToString() + " Closed.");

                }

                if (e.Index == checkedListBox1.FindString("RealistLotSize"))
                {
                    Double lo;
                    Double hi;

                    Double.TryParse(lowLotSizetextBox.Text, out lo);
                    Double.TryParse(highLotSizetextBox.Text, out hi);

                    var m = GlobalVar.listingsFromLastSearch.Where(p => p.Lotsize >= lo && p.Lotsize <= hi);

                    if (dataGridView1.RowCount == 0)
                    {

                        GlobalVar.listingsSearchResults = m;
                        dataGridView1.DataSource = m.ToList();
                    }
                    else
                    {

                        m = GlobalVar.listingsSearchResults.Where(p => p.Lotsize >= lo && p.Lotsize <= hi);
                        GlobalVar.listingsSearchResults = m;
                        dataGridView1.DataSource = m.ToList();
                    }
                    MessageBox.Show("There are " + m.Count().ToString() + " of comparable lotsize");

                }

                if (e.Index == checkedListBox1.FindString("YearBuilt"))
                {
                    int x = 0;
                    int y = 0;
                    //Int32.TryParse(lowGlaTextBox.Text, out x);
                    //Int32.TryParse(highGlaTextBox.Text, out y);
                    Int32.TryParse(SubjectYearBuilt, out x);
                    Int32.TryParse(SubjectYearBuilt, out y);
                

                    if (dataGridView1.RowCount == 0)
                    {

                        var k = GlobalVar.listingsFromLastSearch.Where(p => p.YearBuilt >= x-25 && p.YearBuilt <= y+25);
                        dataGridView1.DataSource = k.ToList();
                        MessageBox.Show("There are " + k.Count().ToString() + " of comparable age.");
                    }
                    else
                    {
                        var k = GlobalVar.listingsSearchResults.Where(p => p.YearBuilt >= x - 25 && p.YearBuilt <= y + 25);
                        GlobalVar.listingsSearchResults = k;
                        dataGridView1.DataSource = k.ToList();
                        MessageBox.Show("There are " + k.Count().ToString() + " of comparable GLA");
                    }
                }

                if (e.Index == checkedListBox1.FindString("RealistGLA"))
                {
                    int x = 0;
                    int y = 0;
                    Int32.TryParse(lowGlaTextBox.Text, out x);
                    Int32.TryParse(highGlaTextBox.Text, out y);

                    if (dataGridView1.RowCount == 0)
                    {

                        var k = GlobalVar.listingsFromLastSearch.Where(p => p.RealistGLA >= x && p.RealistGLA <= y);
                        dataGridView1.DataSource = k.ToList();
                        MessageBox.Show("There are " + k.Count().ToString() + " of comparable GLA");
                    }
                    else
                    {
                        var k = GlobalVar.listingsSearchResults.Where(p => p.RealistGLA >= x && p.RealistGLA <= y);
                        GlobalVar.listingsSearchResults = k;
                        dataGridView1.DataSource = k.ToList();
                        MessageBox.Show("There are " + k.Count().ToString() + " of comparable GLA");
                    }
                }

                switch (e.Index)
                {
                    case 0:
                         Double d = -1;
                        Double.TryParse(ccProximityTextBox.Text, out d);

                         var l = GlobalVar.listingsFromLastSearch.Where(p => p.ProximityToSubject <= d);
                        
                        if (dataGridView1.RowCount == 0)
                        {
                           
                            GlobalVar.listingsSearchResults = l;
                            dataGridView1.DataSource = l.ToList();
                        }
                        else
                        {

                             l = GlobalVar.listingsSearchResults.Where(p => p.ProximityToSubject <= d);
                            GlobalVar.listingsSearchResults = l;
                            dataGridView1.DataSource = l.ToList();
                        }

                        
                       
                     

                        MessageBox.Show("There are " + l.Count().ToString() + " within " + ccProximityTextBox.Text);


                        break;

                    case 1:

                        if (dataGridView1.RowCount == 0)
                        {
                           
                            var k = GlobalVar.listingsFromLastSearch.Where(p => p.PropertyType().Contains(comboBox4.Text));
                            dataGridView1.DataSource = k.ToList();
                            MessageBox.Show("There are " + k.Count().ToString() + " of " + comboBox4.Text);
                        }
                        else
                        {
                             var k = GlobalVar.listingsSearchResults.Where(p => p.PropertyType().Contains(comboBox4.Text));
                              GlobalVar.listingsSearchResults = k;
                            dataGridView1.DataSource = k.ToList();
                            MessageBox.Show("There are " + k.Count().ToString() + " of " + comboBox4.Text);
                        }
                        break;
                    case 2:

                        int x = 0;
                        int y = 0;
                        Int32.TryParse(lowGlaTextBox.Text, out x);
                        Int32.TryParse(highGlaTextBox.Text, out y);

                        if (dataGridView1.RowCount == 0)
                        {

                            var k = GlobalVar.listingsFromLastSearch.Where(p => p.GLA >= x && p.GLA <= y);
                            dataGridView1.DataSource = k.ToList();
                            MessageBox.Show("There are " + k.Count().ToString() + " of comparable GLA");
                        }
                        else
                        {
                            var k = GlobalVar.listingsSearchResults.Where(p => p.GLA >= x && p.GLA <= y);
                            GlobalVar.listingsSearchResults = k;
                            dataGridView1.DataSource = k.ToList();
                            MessageBox.Show("There are " + k.Count().ToString() + " of comparable GLA");
                        }

                        ////var l = 


                        break;

                    case 3:
                        Double lo;
                        Double hi;
                        
                        Double.TryParse(lowLotSizetextBox.Text, out lo);
                        Double.TryParse(highLotSizetextBox.Text, out hi);

                        var m = GlobalVar.listingsFromLastSearch.Where(p => p.Lotsize >= lo && p.Lotsize <= hi);

                        if (dataGridView1.RowCount == 0)
                        {

                            GlobalVar.listingsSearchResults = m;
                            dataGridView1.DataSource = m.ToList();
                        }
                        else
                        {

                            m = GlobalVar.listingsSearchResults.Where(p => p.Lotsize >= lo && p.Lotsize <= hi);
                            GlobalVar.listingsSearchResults = m;
                            dataGridView1.DataSource = m.ToList();
                        }
                        MessageBox.Show("There are " + m.Count().ToString() + " of comparable lotsize");
                       break;

                }

            }
        }

        private void subjectMlsTypecomboBox_TextChanged(object sender, EventArgs e)
        {
            comboBox4.Text = subjectMlsTypecomboBox.Text;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            button18_Click(sender, e);

            var results = GlobalVar.listingsFromLastSearch.Where(p => p.rawData.ToLower().Contains(queryRichTextBox.Text.ToLower()));
            dataGridView1.DataSource = results.ToList();


            //// The IQueryable data to query.
            ////  IQueryable<List<MLSListing>> queryableData = i;
            //var queryableData = GlobalVar.listingsFromLastSearch.AsQueryable<MLSListing>();

            //string str = "fff";


            ////queryableData.Where(p => p.rawData.Contains(str));

            //// Compose the expression tree that represents the parameter to the predicate.
            //ParameterExpression pe = Expression.Parameter(typeof(MLSListing), "p");

            //// Create an expression tree that represents the expression 'p.PropertyType() == "1 Story"'.
            //Expression left = Expression.PropertyOrField(
            ////Expression right = Expression.Constant(queryRichTextBox.Text);
            //Expression e1 = Expression.IsTrue(right);

            //// Combine the expression trees to create an expression tree that represents the

            ////Expression predicateBody = Expression.OrElse(e1, e2);
            //Expression predicateBody = e1;

            //// Create an expression tree that represents the expression 
            //// 'i.Where(p => p.PropertyType() == "1 Story")'
            //MethodCallExpression whereCallExpression = Expression.Call(typeof(Queryable),
            //    "Where",
            //     new Type[] { queryableData.ElementType },
            //     queryableData.Expression,
            //    Expression.Lambda<Func<MLSListing, bool>>(predicateBody, new ParameterExpression[] { pe }));

            //// ***** End Where ***** 

            //// Create an executable query from the expression tree.
            //IQueryable<MLSListing> results = queryableData.Provider.CreateQuery<MLSListing>(whereCallExpression);

          
            
        }

        private void button20_Click(object sender, EventArgs e)
        {
            var rockers = new List<BpoOrder>();

            SpreadsheetsService myService = new SpreadsheetsService("gserve");
            myService.setUserCredentials("beilsco@gmail.com", "Google6079");

            SpreadsheetQuery query = new SpreadsheetQuery();
            SpreadsheetFeed feed = myService.Query(query);

            var campaign = (from x in feed.Entries where x.Title.Text.Contains("bpo tracking") select x).First();

            // GET THE first WORKSHEET from that sheet
            AtomLink link = campaign.Links.FindService(GDataSpreadsheetsNameTable.WorksheetRel, null);
            WorksheetQuery query2 = new WorksheetQuery(link.HRef.ToString());
            WorksheetFeed feed2 = myService.Query(query2);

            var campaignSheet = feed2.Entries.First();


            // GET THE CELLS

            AtomLink cellFeedLink = campaignSheet.Links.FindService(GDataSpreadsheetsNameTable.CellRel, null);
            CellQuery query3 = new CellQuery(cellFeedLink.HRef.ToString());
            CellFeed feed3 = myService.Query(query3);
            uint lastRow = 1;

            BpoOrder rocker = new BpoOrder();
            

            foreach (CellEntry curCell in feed3.Entries)
            {

                if (curCell.Cell.Row > lastRow && lastRow != 1)
                { //When we've moved to a new row, save our BpoOrder
                  
                        rockers.Add(rocker);
                        rocker = new BpoOrder();
           
                }

                //Console.WriteLine("Row {0} Column {1}: {2}", curCell.Cell.Row, curCell.Cell.Column, curCell.Cell.Value);

                switch (curCell.Cell.Column)
                {
                    case 1 : //timestamp
                        DateTime ts;
                        DateTime.TryParse( curCell.Cell.Value, out ts);

                         rocker.timestamp = ts;
                        break;
                    case 4: //address
                        rocker.address = curCell.Cell.Value;
                        Regex rgx = new Regex("[^a-zA-Z0-9]"); //Save a alphanumeric only version
                        //rocker.strippedSite = rgx.Replace(rocker.pics, "");
                        break;
                    case 5: //city
                        rocker.url = curCell.Cell.Value;
                        break;
                    case 6: //sub info
                        rocker.twitter = curCell.Cell.Value;
                        break;
                    case 7: //pics
                        rocker.pics = curCell.Cell.Value;
                        break;
                    case 9: //comps
                        rocker.twitter = curCell.Cell.Value;
                        break;
                    case 11: //due
                        rocker.duedate = curCell.Cell.Value;
                        break;
                    case 12: //completed
                        rocker.completed = curCell.Cell.Value;
                        break;
                    case 13: //billed
                        rocker.twitter = curCell.Cell.Value;
                        break;
                    case 16: //ordernum
                        rocker.ordernum = curCell.Cell.Value;
                        break;

                }
                lastRow = curCell.Cell.Row;
            }

           
         

            var sortedRockers = rockers.Where(x => x.timestamp >= DateTime.Now.AddDays(-5)).ToList();

            dataGridView1.DataSource = sortedRockers;

        }

        private void button21_Click(object sender, EventArgs e)
        {
            string baseDir = DropBoxFolder + @"\BPOs\";
            string subjectAddress = streetnameTextBox.Text;
            string subjecAddressShort = subjectAddress.Replace(" RD", "").Replace(" ", "<SP>");
              



            string newPath = baseDir + subjectAddress;

            CreateDirectory(newPath);

            search_address_textbox.Text = newPath;

            iMacros.App browser = new iMacros.App();

            //need to open an imacro browser
            browser.iimOpen("", true, 30);

            //this should be a unused sec account
            browser.iimPlay(@"MRED\#start_mls_robo.iim");
            browser.iimPlay(@"MRED\mred-open-realist.iim");

            // iMacros.AppClass iim = new iMacros.AppClass();
            // iMacros.Status status = iim.iimOpen("", true, timeout);
            StringBuilder macro = new StringBuilder();
        
           
            macro.AppendLine(@"SET !TIMEOUT_STEP 20");
            macro.AppendLine(@"TAB T=2");
            macro.AppendLine(@"IMAGECLICK POS=1 IMAGE=20130920_1822.png CONFIDENCE=95 CONTENT=" + subjectAddress.Replace(" ","<SP>"));
            macro.AppendLine(@"IMAGECLICK POS=1 IMAGE=20130920_1924.png CONFIDENCE=95");
            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"IMAGESEARCH POS=1 IMAGE=propert-suggestion-box.png CONFIDENCE=95");
            macro.AppendLine(@"IMAGECLICK POS=1 IMAGE=checkbox-1.png CONFIDENCE=95");
            macro.AppendLine(@"IMAGECLICK POS=1 IMAGE=20130920_1919.png CONFIDENCE=95");
            macro.AppendLine(@"IMAGECLICK POS=1 IMAGE=20130920_1921.png CONFIDENCE=95"); macro.AppendLine(@"VERSION BUILD=9012597");
            macro.AppendLine(@"ONDOWNLOAD FOLDER=" + newPath.Replace(" ","<SP>") + " FILE=* WAIT=YES");
            macro.AppendLine(@"IMAGECLICK POS=1 IMAGE=20130920_1909.png CONFIDENCE=95");
            macro.AppendLine(@"IMAGECLICK POS=1 IMAGE=20130920_1910.png CONFIDENCE=95");
            macro.AppendLine(@"WAIT SECONDS=15");


            macro.AppendLine(@"TAB T=1");
            macro.AppendLine(@"TAB CLOSEALLOTHERS");
            macro.AppendLine(@"TAG POS=1 TYPE=NOBR ATTR=TXT:My<SP>MLS");
            macro.AppendLine(@"FRAME NAME=main");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:listings ATTR=NAME:searchField CONTENT=" + subjecAddressShort);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:listings ATTR=NAME:Find");
            macro.AppendLine(@"FRAME NAME=workspace");
            macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:0*");
            macro.AppendLine(@"FRAME NAME=subheader");
            macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:header ATTR=ID:pdf");
            macro.AppendLine(@"'New tab opened");
            macro.AppendLine(@"TAB T=2");
            macro.AppendLine(@"'New tab opened");
            macro.AppendLine(@"TAB T=3");
            macro.AppendLine(@"FRAME F=0");
            //macro.AppendLine(@"ONDOWNLOAD FOLDER=* FILE=* WAIT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:click<SP>here");


            string macroCode = macro.ToString();
             status = browser.iimPlayCode(macroCode, 90);


             importSubjectInfoButton(sender, e);









        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                streetnameTextBox.Text = dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString();
            }
            catch
            {
            }
        }

        private void label41_Click(object sender, EventArgs e)
        {

        }

        private void label43_Click(object sender, EventArgs e)
        {
            


        }

       

        private void button_imort_pics_Click2(object sender, EventArgs e)
        {

            
            
            
            this.button_imort_pics_Click(sender, e);

            
            //StringBuilder macro = new StringBuilder();
            //macro.AppendLine(@"SET !TIMEOUT_STEP 200");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$lnkManualImages");
            //macro.AppendLine(@"FRAME NAME=sb-player");
            //macro.AppendLine(@"'TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:PhotoUploader.aspx?m=1&type=bpo&oid=* ATTR=NAME:dgManualPhotos$ctl04$UploadFile");
            //macro.AppendLine(@"TAB T=1");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:PhotoUploader.aspx?m=1&type=bpo&oid=* ATTR=NAME:dgManualPhotos$ctl02$UploadFile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\L1.jpg");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:PhotoUploader.aspx?m=1&type=bpo&oid=* ATTR=NAME:dgManualPhotos$ctl03$UploadFile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\l2.jpg");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:PhotoUploader.aspx?m=1&type=bpo&oid=* ATTR=NAME:dgManualPhotos$ctl04$UploadFile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\l3.jpg");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:PhotoUploader.aspx?m=1&type=bpo&oid=* ATTR=NAME:dgManualPhotos$ctl05$UploadFile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\s1.jpg");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:PhotoUploader.aspx?m=1&type=bpo&oid=* ATTR=NAME:dgManualPhotos$ctl06$UploadFile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\s2.jpg");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:PhotoUploader.aspx?m=1&type=bpo&oid=* ATTR=NAME:dgManualPhotos$ctl07$UploadFile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\s3.jpg");

            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:PhotoUploader.aspx?m=1&type=bpo&oid=* ATTR=NAME:dgManualPhotos$ctl08$UploadFile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\address.jpg");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:PhotoUploader.aspx?m=1&type=bpo&oid=* ATTR=NAME:dgManualPhotos$ctl09$UploadFile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\front.jpg");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:PhotoUploader.aspx?m=1&type=bpo&oid=* ATTR=NAME:dgManualPhotos$ctl10$UploadFile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\left.jpg");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:PhotoUploader.aspx?m=1&type=bpo&oid=* ATTR=NAME:dgManualPhotos$ctl11$UploadFile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\right.jpg");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:PhotoUploader.aspx?m=1&type=bpo&oid=* ATTR=NAME:dgManualPhotos$ctl12$UploadFile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\street.jpg");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:PhotoUploader.aspx?m=1&type=bpo&oid=* ATTR=NAME:dgManualPhotos$ctl13$UploadFile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\street2.jpg");

            //macro.AppendLine(@"");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:PhotoUploader.aspx?m=1&type=bpo&oid=* ATTR=NAME:cmdUpload");
            //macro.AppendLine(@"'FRAME F=0");
            //macro.AppendLine(@"'TAG POS=1 TYPE=INPUT ATTR=ID:dgManualPhotos_ctl04_UploadFile");
          
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$dlImages$ctl00$cboImage CONTENT=%8");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$dlImages$ctl01$cboImage CONTENT=%9");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$dlImages$ctl02$cboImage CONTENT=%10");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$dlImages$ctl03$cboImage CONTENT=%5");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$dlImages$ctl04$cboImage CONTENT=%6");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$dlImages$ctl05$cboImage CONTENT=%7");

            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$dlImages$ctl06$cboImage CONTENT=%2");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$dlImages$ctl07$cboImage CONTENT=%1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$dlImages$ctl08$cboImage CONTENT=%24");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$dlImages$ctl09$cboImage CONTENT=%25");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$dlImages$ctl10$cboImage CONTENT=%3");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$dlImages$ctl11$cboImage CONTENT=%27");

            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$cboRecommended CONTENT=%113");
            //macro.AppendLine(@"");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$cboNewConstruction CONTENT=%113");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$cboDisaster CONTENT=%113");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:Bpo$Repairs$cboResaleProblem CONTENT=%113");

            //macro.AppendLine(@"TAG POS=1 TYPE=A FORM=ACTION:Bpo.aspx?control=* ATTR=TXT:Save");
            //string macroCode = macro.ToString();
            //iim2.iimPlayCode(macroCode, 60);
        }

        private void button22_Click(object sender, EventArgs e)
        {

            bool stillComps = true;
            string compNumber;
            string fileName = "Error";


            StringBuilder macro12 = new StringBuilder();
            macro12.AppendLine(@"FRAME NAME=navpanel");
            macro12.AppendLine(@"TAG POS=1 TYPE=TD ATTR=ID:navtd EXTRACT=TXT");
            macro12.AppendLine(@"FRAME NAME=workspace");
            //// macro12.AppendLine(@"TAG POS=1 TYPE=TABLE FORM=NAME:dc ATTR=CLASS:gridview EXTRACT=TXT");
            // macro12.AppendLine(@"TAG POS=1 TYPE=HTML ATTR=* EXTRACT=HTM ");
            macro12.AppendLine(@"TAG POS=1 TYPE=DIV FORM=NAME:dc ATTR=CLASS:report EXTRACT=HTM");
            iMacros.Status s = iim.iimPlayCode(macro12.ToString());
            string htmlCode = iim.iimGetLastExtract();
            string header = iim.iimGetExtract(0);


            Stack<string> saleComps = new Stack<string>();
            Stack<string> listComps = new Stack<string>();


            saleComps.Push("S3");
            saleComps.Push("S2");
            saleComps.Push("S1");

            listComps.Push("L3");
            listComps.Push("L2");
            listComps.Push("L1");

            StringBuilder openTab = new StringBuilder();
            //openTab.AppendLine(@"ONDOWNLOAD FOLDER=" +  SubjectFilePath.Replace(" ", "<SP>") + " FILE=S1.pdf WAIT=YES");



            string filepath = search_address_textbox.Text;

            if (!Regex.IsMatch(header, @"showing\s*1\s*of\s*6"))
            {
                if (MessageBox.Show("Do you want to start at comp 1?", "Error", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    return;
                }
            }





            while (stillComps)
            {

                if (Regex.IsMatch(htmlCode, @"ACTV|CTG|PEND|TMP"))
                {
                    fileName = listComps.Pop();
                }

                if (Regex.IsMatch(htmlCode, @"CLSD"))
                {
                    fileName = saleComps.Pop();
                }

                if (Regex.IsMatch(htmlCode, @"showing\s*1\s*of\s*6"))
                {
                    compNumber = "1";
                }

                if (Regex.IsMatch(htmlCode, @"showing\s*2\s*of\s*6"))
                {
                    compNumber = "2";
                }
                if (Regex.IsMatch(htmlCode, @"showing\s*3\s*of\s*6"))
                {
                    compNumber = "3";
                }
                if (Regex.IsMatch(htmlCode, @"showing\s*4\s*of\s*6"))
                {
                    compNumber = "4";
                }
                if (Regex.IsMatch(htmlCode, @"showing\s*5\s*of\s*6"))
                {
                    compNumber = "5";
                }
                if (Regex.IsMatch(htmlCode, @"showing\s*6\s*of\s*6"))
                {
                    compNumber = "6";
                    stillComps = false;
                }
                openTab.Clear();

                openTab.AppendLine(@"FRAME NAME=subheader");
                // openTab.AppendLine(@"ONDOWNLOAD FOLDER=" + SubjectFilePath.Replace(" ", "<SP>") + " FILE=" + fileName + "WAIT=YES");
                openTab.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:header ATTR=ID:pdf");
                openTab.AppendLine(@"ONDOWNLOAD FOLDER=" + SubjectFilePath.Replace(" ", "<SP>") + " FILE=" + fileName + ".pdf WAIT=YES");
                openTab.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:click<SP>here");

                openTab.AppendLine(@"WAIT SECONDS = 10");
                openTab.AppendLine(@"TAB CLOSE");
                openTab.AppendLine(@"FRAME NAME=navpanel");
                openTab.AppendLine(@"TAG POS=1 TYPE=IMG ATTR=SRC:http://connectmls*.mredllc.com/images/next.gif");

                s = iim.iimPlayCode(openTab.ToString());
                s = iim.iimPlayCode(macro12.ToString());
                htmlCode = iim.iimGetLastExtract();


            }
            StringBuilder macro2 = new StringBuilder();
            macro2.AppendLine(@"TAB T=1");
            macro2.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:BPO$Comparables$MLSPdfUpload2 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\S1.pdf");
            macro2.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:BPO$Comparables$MLSPdfUpload3 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\S2.pdf");
            macro2.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:BPO$Comparables$MLSPdfUpload4 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\S3.pdf");
            macro2.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:BPO$Comparables$MLSPdfUpload5 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\L1.pdf");
            macro2.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:BPO$Comparables$MLSPdfUpload6 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\L2.pdf");
            macro2.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:BPO$Comparables$MLSPdfUpload7 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + "\\L3.pdf");
            macro2.AppendLine(@"TAG POS=1 TYPE=A FORM=ACTION:Bpo.aspx?control=* ATTR=TXT:Save");
            iim2.iimPlayCode(macro2.ToString(), 60);


           // if (Regex.IsMatch(htmlCode, @"ACTV|CTG|PEND|TMP") )
           // {
           //     save_pics_macro.AppendLine(@"ONDOWNLOAD FOLDER=" + filepath.Replace(" ", "<SP>") + " FILE=S" + closed_comps.ToString() + ".jpg");
           // }
           // else
           // {
           //     save_pics_macro.AppendLine(@"ONDOWNLOAD FOLDER=" + filepath.Replace(" ", "<SP>") + " FILE=L" + active_comps.ToString() + ".jpg");
           // }

           //      s = iim.iimPlayCode(macro12.ToString());
           //      htmlCode = iim.iimGetLastExtract();
               
            


            //#region save comp pic
            ////open pic tab
            //StringBuilder openTab = new StringBuilder();
            //openTab.AppendLine(@"FRAME NAME=subheader");

            //openTab.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:header ATTR=ID:pdf");
            ////openTab.AppendLine(@"FRAME F=0");
            //openTab.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=HREF:javascript:openWindow* EXTRACT=TXT");

            //string openTab_macroCode = openTab.ToString();
            //status = iim.iimPlayCode(openTab_macroCode, 30);

            //openTab.Clear();

            //StringBuilder save_pics_macro = new StringBuilder();

            ////string filepath = "C:\\Users\\Scott\\Documents\\My<SP>Dropbox\\BPOs\\" + search_address_textbox.Text;
            //string filepath = search_address_textbox.Text;
            //string filename = "comp.jpg";

            //if (sale_or_list_flag == "sale")
            //{
            //    save_pics_macro.AppendLine(@"ONDOWNLOAD FOLDER=" + filepath.Replace(" ", "<SP>") + " FILE=S" + closed_comps.ToString() + ".jpg");
            //}
            //else
            //{
            //    save_pics_macro.AppendLine(@"ONDOWNLOAD FOLDER=" + filepath.Replace(" ", "<SP>") + " FILE=L" + active_comps.ToString() + ".jpg");
            //}

            ////line changed for new version of connectmls

            ////save_pics_macro.AppendLine(@"'New tab opened");
            //// save_pics_macro.AppendLine(@"WAIT SECONDS=2");
            //save_pics_macro.AppendLine(@"TAB T=2");

            ////save_pics_macro.AppendLine(@"TAG POS=2 TYPE=IMG FORM=NAME:dc ATTR=HREF:""*.JPEG"" CONTENT=EVENT:SAVEITEM");
            ////save_pics_macro.AppendLine(@"WAIT SECONDS=5");
            ////also changed for new connectmls
            ////save_pics_macro.AppendLine(@"TAG POS=2 TYPE=A FORM=NAME:dc ATTR=TXT:Download<SP>This<SP>Photo CONTENT=EVENT:SAVETARGETAS");
            //// save_pics_macro.AppendLine(@"TAG POS=2 TYPE=A FORM=NAME:dc ATTR=TXT:Download<SP>This<SP>Photo CONTENT=EVENT:SAVEITEM");

            //// iim.iimPlayCode(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=HREF:javascript:openWindow('http%3A%2F%2Ftours.databasedads.com%2F2797441%2F1314-S-W* EXTRACT=TXT");




            //save_pics_macro.AppendLine(@"FRAME F=0");

            //if (iim.iimGetLastExtract().Contains("Virtual Tour"))
            //{
            //    // save_pics_macro.AppendLine(@"TAG POS=3 TYPE=IMG FORM=NAME:dc ATTR=HREF:""*.JPEG"" CONTENT=EVENT:SAVEITEM");
            //    save_pics_macro.AppendLine(@"TAG POS=3 TYPE=IMG FORM=NAME:dc ATTR=HREF:* CONTENT=EVENT:SAVEITEM");
            //}
            //else
            //{
            //    //  save_pics_macro.AppendLine(@"TAG POS=2 TYPE=IMG FORM=NAME:dc ATTR=HREF:""*.JPEG"" CONTENT=EVENT:SAVEITEM");
            //    save_pics_macro.AppendLine(@"TAG POS=2 TYPE=IMG FORM=NAME:dc ATTR=HREF:* CONTENT=EVENT:SAVEITEM");
            //}


            ////save_pics_macro.AppendLine(@"WAIT SECONDS=2");
            //save_pics_macro.AppendLine(@"TAB CLOSE");

            //string save_pics_macroCode = save_pics_macro.ToString();
            //status = iim.iimPlayCode(save_pics_macroCode, 30);
            //#endregion
        }

        private async void button23_Click(object sender, EventArgs e)
        {
            GoogleCloudDatastore gcds = new GoogleCloudDatastore();
              await gcds.DataStoreTester();

              MessageBox.Show(GlobalVar.sandbox);
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }

        private void button24_Click(object sender, EventArgs e)
        {
            //webBrowser1.Navigate(new Uri("https://www.rapidclose.com"));
            HttpWebRequest myReq =
                (HttpWebRequest)WebRequest.Create("https://www.rapidclose.com/sfs/swappraiser/AppraiserLogin.jsp");
            myReq.CookieContainer = new CookieContainer();

            HttpWebResponse x = (HttpWebResponse) myReq.GetResponse();
            Stream y = x.GetResponseStream();

         //   MessageBox.Show( x.ToString());
           // MessageBox.Show(y.ToString());
            StreamReader reader = new StreamReader(y);
            string responseFromServer = reader.ReadToEnd();
            //MessageBox.Show(responseFromServer.ToString());
            reader.Close();
            foreach (System.Net.Cookie cook in x.Cookies)
            {
                //Console.WriteLine("Cookie:");
                MessageBox.Show(cook.Name + " = " + cook.Value);

                //Console.WriteLine("Domain: {0}", cook.Domain);
                //Console.WriteLine("Path: {0}", cook.Path);
                //Console.WriteLine("Port: {0}", cook.Port);
                //Console.WriteLine("Secure: {0}", cook.Secure);

                //Console.WriteLine("When issued: {0}", cook.TimeStamp);
                //Console.WriteLine("Expires: {0} (expired? {1})",
                //    cook.Expires, cook.Expired);
                //Console.WriteLine("Don't save: {0}", cook.Discard);
                //Console.WriteLine("Comment: {0}", cook.Comment);
                //Console.WriteLine("Uri for comments: {0}", cook.CommentUri);
                //Console.WriteLine("Version: RFC {0}", cook.Version == 1 ? "2109" : "2965");

                //// Show the string representation of the cookie.
                //Console.WriteLine("String: {0}", cook.ToString());
            }

            HttpWebRequest myReq2 = (HttpWebRequest)WebRequest.Create("https://www.rapidclose.com/sfs/swcommon/AuthenticateAppraiser");

            myReq2.Host = @"www.rapidclose.com";   
            myReq2.CookieContainer = new CookieContainer();

            myReq2.CookieContainer = myReq.CookieContainer;
            myReq2.Method = "POST";

            string postData = @"userAgent=Mozilla%2F5.0+%28Windows+NT+6.1%3B+WOW64%29+AppleWebKit%2F537.36+%28KHTML%2C+like+Gecko%29+Chrome%2F35.0.1916.153+Safari%2F537.36&appVersion=5.0+%28Windows+NT+6.1%3B+WOW64%29+AppleWebKit%2F537.36+%28KHTML%2C+like+Gecko%29+Chrome%2F35.0.1916.153+Safari%2F537.36&vendor=Google+Inc.&product=Gecko&javascriptEnabled=true&cookiesEnabled=true&guid=&login=beilsco%40gmail.com&password=27692scott&submit=Login";
            byte[] byteArray = Encoding.UTF8.GetBytes(postData);
            myReq2.ContentType = @"application/x-www-form-urlencoded";
            myReq2.ContentLength = byteArray.Length;
            // Get the request stream.
            Stream dataStream = myReq2.GetRequestStream();
            // Write the data to the request stream.
            dataStream.Write(byteArray, 0, byteArray.Length);
            // Close the Stream object.
            dataStream.Close();
            // Get the response.
            WebResponse response = myReq2.GetResponse();
            // Display the status.
            Console.WriteLine(((HttpWebResponse)response).StatusDescription);
            // Get the stream containing content returned by the server.
            dataStream = response.GetResponseStream();
            // Open the stream using a StreamReader for easy access.
            StreamReader reader2 = new StreamReader(dataStream);
            // Read the content.
            string responseFromServer2 = reader2.ReadToEnd();
            // Display the content.
            Console.WriteLine(responseFromServer2);
            // Clean up the streams.
            reader.Close();
            dataStream.Close();
            response.Close();
          
            var x2 = myReq2.GetResponse();

            MessageBox.Show(x2.ToString());

            y.Close();
            //myReq2.CookieContainer.Add(new CookieCollection mycookies());
                



         

        }

        private void label45_Click(object sender, EventArgs e)
        {

        }

        //private void textBox2_TextChanged_1(object sender, EventArgs e)
        //{

        //}

        private void ndNumberActiveReoListingsTextBox_TextChanged(object sender, EventArgs e)
        {
            Decimal activeReoListings = 0;
            Decimal totalActiveListings = 0;
            Decimal activeShortListings = 0;

            Decimal.TryParse(this.ndNumberActiveReoListingsTextBox.Text, out activeReoListings);
            Decimal.TryParse(this.ndNumberOfActiveListingTextBox.Text, out totalActiveListings);
            Decimal.TryParse(this.ndNumberActiveShortListingsTextBox.Text, out activeShortListings);

            try
            {
                percentActiveListingLabel.Text = (activeReoListings / totalActiveListings).ToString("p") + "Active Listings - REO";
                percentDistressedActiveListingsLabel.Text = ((activeReoListings + activeShortListings) / totalActiveListings).ToString("p") + "Active Listings - Distressed";
            }
            catch
            { }

        }

        private void neighborhoodDataGroupBox_Enter(object sender, EventArgs e)
        {

        }

    

        private void ndNumberActiveShortListingsTextBox_TextChanged(object sender, EventArgs e)
        {
            Decimal activeReoListings = 0;
            Decimal totalActiveListings = 0;
            Decimal activeShortListings = 0;

            Decimal.TryParse(this.ndNumberActiveReoListingsTextBox.Text, out activeReoListings);
            Decimal.TryParse(this.ndNumberOfActiveListingTextBox.Text, out totalActiveListings);
            Decimal.TryParse(this.ndNumberActiveShortListingsTextBox.Text, out activeShortListings);

            try
            {
                percentActiveShortListingLabel.Text = (activeShortListings / totalActiveListings).ToString("p") + "Active Listings - Short";
                percentDistressedActiveListingsLabel.Text = ((activeReoListings + activeShortListings) / totalActiveListings).ToString("p") + "Active Listings - Distressed";
            }
            catch
            { }
        }

        private void ndNumberSoldReoListingsTextBox_TextChanged(object sender, EventArgs e)
        {
            Decimal soldReoListings = 0;
            Decimal totalSoldListings = 0;
            Decimal soldShortListings = 0;

            Decimal.TryParse(this.ndNumberSoldReoListingsTextBox.Text, out soldReoListings);
            Decimal.TryParse(this.ndNumberOfSoldListingTextBox.Text, out totalSoldListings);
            Decimal.TryParse(this.ndNumberSoldShortListingsTextBox.Text, out soldShortListings);

            try
            {
                percentSoldReoListingLabel.Text = (soldReoListings / totalSoldListings).ToString("p") + "Sold Listings - REO";
                percentDistressedSoldListingsLabel.Text = ((soldReoListings + soldShortListings) / totalSoldListings).ToString("p") + "Sold Listings - Distressed";
            }
            catch
            { }
        }

        private void ndNumberSoldShortListingsTextBox_TextChanged(object sender, EventArgs e)
        {
            Decimal soldReoListings = 0;
            Decimal totalSoldListings = 0;
            Decimal soldShortListings = 0;

            Decimal.TryParse(this.ndNumberSoldReoListingsTextBox.Text, out soldReoListings);
            Decimal.TryParse(this.ndNumberOfSoldListingTextBox.Text, out totalSoldListings);
            Decimal.TryParse(this.ndNumberSoldShortListingsTextBox.Text, out soldShortListings);

            try
            {
                percentSoldShortListingLabel.Text = (soldShortListings / totalSoldListings).ToString("p") + "Sold Listings - Short";
                percentDistressedSoldListingsLabel.Text = ((soldReoListings + soldShortListings) / totalSoldListings).ToString("p") + "Sold Listings - Distressed";
            }
            catch
            { }
        }

        private void REOcSandboxButton_Click(object sender, EventArgs e)
        {
            iMacros.App reocFireFoxBrowser = new iMacros.App();
            status = reocFireFoxBrowser.iimOpen("-fx", false, 60);
            reocFireFoxBrowser.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string currentUrl = reocFireFoxBrowser.iimGetLastExtract();

            if (currentUrl.Contains("reo-central"))
            {
                iim2 = reocFireFoxBrowser;
            }


        }

        private void buttonSaveOrderInfo_Click(object sender, EventArgs e)
        {
            iim2.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string currentUrl = iim2.iimGetLastExtract();



            if (currentUrl.Contains("solutionstar"))
            {

                StringBuilder macro12 = new StringBuilder();
                StringBuilder macro = new StringBuilder();

                macro12.AppendLine(@"TAG POS=1 TYPE=HTML ATTR=* EXTRACT=HTM ");

                iMacros.Status s = iim2.iimPlayCode(macro12.ToString());
                string htmlCode = iim2.iimGetLastExtract();
                string header = iim2.iimGetExtract(0);

                var ttt = Regex.Matches(htmlCode, @"orderNo=(\d\d-\d\d\d\d\d\d\d\d)");

           //     MessageBox.Show(ttt.Count.ToString());


                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(htmlCode);

                var x = doc.DocumentNode.SelectSingleNode("//*[@id=\"appraisalOrderList\"]/tbody/tr[1]/td[1]/a/font");
                int orderNumberIndex = 1;
                int addressIndex = 4;
                int cityIndex = 5;

                //
                //TBD:  Calculate correct due date by extracting date + standard 3 days.
                //
                for (int i = 1; i < ttt.Count + 1; i++)
                {
                    macro.AppendLine(@"VERSION BUILD=10022823");
                    macro.AppendLine(@"TAB T=1");
                    macro.AppendLine(@"TAB CLOSEALLOTHERS");
                    macro.AppendLine(@"URL GOTO=https://docs.google.com/spreadsheet/viewform?usp=drive_web&formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ#gid=20");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.2.single CONTENT=" + doc.DocumentNode.SelectSingleNode("//*[@id=\"appraisalOrderList\"]/tbody/tr[" + i.ToString() + "]/td[" + addressIndex.ToString() + "]").InnerText.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.3.single CONTENT=" + doc.DocumentNode.SelectSingleNode("//*[@id=\"appraisalOrderList\"]/tbody/tr[" + i.ToString() + "]/td[" + cityIndex.ToString() + "]").InnerText.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.14.single CONTENT=" + doc.DocumentNode.SelectSingleNode("//*[@id=\"appraisalOrderList\"]/tbody/tr[" + i.ToString() + "]/td[" + orderNumberIndex.ToString() + "]/a/font").InnerText);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.9.single CONTENT=" + DateTime.Now.ToShortDateString());
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.11.single CONTENT=50");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.1.single CONTENT=%Solutionstar");
                   macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:submit");
                    string macroCode = macro.ToString();
                    iim.iimPlayCode(macroCode, 60);
                    macro.Clear();
                
                }
            
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            GenerateVersions(@"F:\Dropbox\BPOs\3703 Jacobson\address-view.jpg");

        }

        private void subjectSubdivisionTextbox_TextChanged(object sender, EventArgs e)
        {

        }

       
       

  

        //private void label6_Click(object sender, EventArgs e)
        //{

        //}
    }

    public class MlsReportDriver
    {
        public MlsReportDriver()
        {
            
        }
        private readonly IYourForm form;
        public MlsReportDriver (IYourForm form)
        {
            this.form = form;
        
        }
        string listPrice;
        public string mlsStatus;
        public string assessmentIncludes;
        public string assessmentContactPhone;
        public string assessmentContact;
        public string assessmentFrequncy;
        public string assessmentAmount;
        public string roomCount;
        public string bedroomCount;
        public string bathCount;
        public string fullBathCount;
        public string halfBathCount = "0";
        public string basementType;
        public string basementDetails;
        public string numFireplaces = "0";
        public string parking;
        public string style;
        public string type;
        public bool attached;
        public bool detached;
        public string listAgent;
        public string brokerPhone;
        public string dom;
        public string mlsTaxAmount;

        private string rawText;
        private string GetFieldValue(string fn)
        {
            //:x*(.*?)xxx


            string pattern1 = "x*(.*?)xxx";
            string pattern2 = "x*(.*?)\n";
            string returnValue = "NotFound";
            string result = "";

            //result = Regex.Match(rawText, string.Format(@"{0}{1}", fn, pattern1)).Groups[1].Value;
            result = Regex.Match(rawText, string.Format(@"{0}{1}|{0}{2}", fn, pattern1, pattern2)).Groups[1].Value;

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



        public string Get_Distance(string o, string d)
        {


            //
            //new way to reduce google map webservice calls
            //

            //Position pos1 = new Position();
            //pos1.Latitude = Convert.ToDouble(hotelx.latitude);
            //pos1.Longitude = Convert.ToDouble(hotelx.longitude);
            //Position pos2 = new Position();
            //pos2.Latitude = Convert.ToDouble(lat);
            //pos2.Longitude = Convert.ToDouble(lng);
            //Haversine calc = new Haversine();

            //double result = calc.Distance(pos1, pos2, DistanceType.Miles);
           
            
            //
            //end new way
            //\\

            //Google distance matric webservice
            //http://maps.googleapis.com/maps/api/distancematrix/xml?origins=60050&destinations=60002&sensor=false&units=imperial
            //
            //string zip = "60085";

           



            string distance = "0";
                                                                                                     //Lot#104 --> Lot 104
            string googlestr = "http://maps.googleapis.com/maps/api/distancematrix/xml?origins=" + o.Replace("#", " ") + "&destinations=" + d + "&sensor=false&units=imperial";

            // Create a request for the URL. 		
            WebRequest request = WebRequest.Create(googlestr);

            // Get the response.
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();


            // Get the stream containing content returned by the server.
            Stream dataStream = response.GetResponseStream();
            // Open the stream using a StreamReader for easy access.
            StreamReader reader = new StreamReader(dataStream);
            // Read the content.
            string responseFromServer = reader.ReadToEnd();
            // Display the content.

            //MessageBox.Show(responseFromServer);

            //<distance>
            //<value>42315</value>
            //<text>26.3 mi</text>
            //</distance>

            string pattern = "(\\d+.\\d*) [mf]";
            Match match = Regex.Match(responseFromServer, pattern);
            distance = match.Groups[1].Value;
            if (responseFromServer.Contains(" ft</text>"))
            {
                distance = "0.05";
            }
            //MessageBox.Show(distance);
            //Console.WriteLine(responseFromServer);
            // Cleanup the streams and the response.
            reader.Close();
            dataStream.Close();
            response.Close();

            return distance;
        }
        public double Get_Distance(GeoPoint o, GeoPoint d)
        {
            //
            //new way to reduce google map webservice calls
            //and return "crow flys" distance
            //

            //Position pos1 = new Position();
            //pos1.Latitude = Convert.ToDouble(hotelx.latitude);
            //pos1.Longitude = Convert.ToDouble(hotelx.longitude);
            //Position pos2 = new Position();
            //pos2.Latitude = Convert.ToDouble(lat);
            //pos2.Longitude = Convert.ToDouble(lng);
            Haversine calc = new Haversine();

            double result = calc.Distance(o, d, DistanceType.Miles);

            return result;
        }
        private int Age(string yearBuilt)
        {
            DateTime subject_age = new DateTime((Convert.ToInt32(yearBuilt)), 1, 1);

            TimeSpan ts = DateTime.Now - subject_age;

            return ts.Days / 365;
        }
        public void GetSubjectInfo(string s)
        {
            MLSListing m = new MLSListing();
            rawText =  s;


            string masterPattern = @"xxx(.*?)xxx|xxx(.*?)\n";
            MatchCollection x = Regex.Matches(s, masterPattern);

            GlobalVar.theSubjectProperty.PrintedMlsSheetNameValuePairs = x;


            string pattern = @"Ax*mount:\$(\d+,*\d*.\d*)";
            MatchCollection mc = Regex.Matches(s, pattern);
            assessmentAmount = mc[0].Groups[1].Value;
            try
            {
                mlsTaxAmount = mc[1].Groups[1].Value;
            }
            catch
            {
                mlsTaxAmount = "-1";
            }

            //pattern = "Ph #:x*([^\\nx]+)|Ph #:x*([^\\n]+)";
            pattern = @"xxxPh #:(.*?)xxx|xxxPh #:(.*?)\n";
            mc = Regex.Matches(s, pattern);
            string oorPhone = mc[0].Groups[1].Value;
            brokerPhone = mc[1].Groups[1].Value;
            string coListerPhone = mc[2].Groups[1].Value;



         
            pattern = "Lst. Mkt. Time:x*([^\\nx]+)|Lst. Mkt. Time:x*([^\\n]+)";
            Match match = Regex.Match(s, pattern);
            dom = match.Groups[1].Value;

            pattern = "List Agent:x*([^\\nx]+)|List Agent:x*([^\\n]+)";
            match = Regex.Match(s, pattern);
            listAgent = match.Groups[1].Value;

            

            pattern = "List Price:x*([^\\nx]+)|List Price:x*([^\\n]+)";
            match = Regex.Match(s, pattern);
            listPrice = match.Groups[1].Value;

            pattern = @"(ACTV|CTG|CLSD|TEMP|CANC|EXP)";
            match = Regex.Match(s, pattern);
            mlsStatus = match.Groups[1].Value;

            

            ////pattern = "Status:x*([^\\nx]+)|Status:x*([^\\n]+)";
            ////match = Regex.Match(s, pattern);
            ////mlsStatus = match.Groups[1].Value;



            pattern = "Asmt Incl:x*([^\\nx]+)|Asmt Incl:x*([^\\n]+)";
            match = Regex.Match(s, pattern);
            assessmentIncludes = match.Groups[1].Value;

            pattern = "Phone:x*([^\\nx]+)|Phone:x*([^\\n]+)";
            match = Regex.Match(s, pattern);
            assessmentContactPhone = match.Groups[1].Value;

            pattern = "Contact Name:([^x]+)";
            match = Regex.Match(s, pattern);
            assessmentContact = match.Groups[1].Value;

            pattern = "Frequency:([^ ]+)";
            match = Regex.Match(s, pattern);
            assessmentFrequncy = match.Groups[1].Value;

             pattern = "Rooms:([^x]+)";
             match = Regex.Match(s, pattern);
            roomCount = match.Groups[1].Value;


            pattern = "Detached Single";
            match = Regex.Match(s, pattern);
            if (match.Success)
            {
                m = new DetachedListing();
                detached = true;
                attached = false;
            }

            pattern = "Ax*ttached Single";
            match = Regex.Match(s, pattern);
            if (match.Success)
            {
                m = new AttachedListing();
                detached = false;
                attached = true;
            }


            //Bathrooms2 /xxx
            //Bathrooms 1 / 1
            pattern = "Bathrooms\\s*(\\d+)";
            match = Regex.Match(s, pattern);
            fullBathCount = match.Groups[1].Value;

            pattern = "Bathrooms\\s*\\d\\s.\\s(\\d+)";
            match = Regex.Match(s, pattern);
            if (match.Success)
            {
                halfBathCount = match.Groups[1].Value;
            }

            bathCount = fullBathCount + "." + halfBathCount;

            //Bedrooms:3xxx
            pattern = "Bedrooms:(\\d+)";
            match = Regex.Match(s, pattern);
            bedroomCount = match.Groups[1].Value;

            pattern = "Basement:([^x]+)";
            match = Regex.Match(s, pattern);
            basementType = match.Groups[1].Value;
            //Basement Details: Partially Finished
            //Basement Details:Partially Finished

            //this works, but one saved match will be ""
            //pattern = "Basement Details:x*([^\\n]+)xxx|Basement Details:x*([^\\n]+)";
            
            pattern = "Basement Details:x*([^\\nx]+)|Basement Details:x*([^\\n]+)";
            
            match = Regex.Match(s, pattern);
            if (match.Groups[1].Value != "")
            {
                basementDetails = match.Groups[1].Value;
            }
            else
            {
                basementDetails = match.Groups[2].Value;
            }
          

            //Fireplaces:xxx1
            //Fireplaces:0
            pattern = "Fireplaces:\\w*(\\d+)";
            match = Regex.Match(s, pattern);
            numFireplaces = match.Groups[1].Value;
            if (string.IsNullOrWhiteSpace(numFireplaces))
            {
                numFireplaces = "0";
            }



            //Parking:Garage
            //Parking:Exterior Space(s)
            pattern = "Parking:x*([^\\n]+)";
            match = Regex.Match(s, pattern);
            parking = match.Groups[1].Value;

            if (match.Success)
            {


                //Spaces:Gar:2
                pattern = "Spaces:Gar:(\\d+)";
                match = Regex.Match(s, pattern);
                parking = parking + " " + match.Groups[1].Value;

                //pattern = "Garage Type:x*([^x\\n]+)xxx";
                pattern = "Garage Type:x*(Attached|Detached)";
                match = Regex.Match(s, pattern);
                parking = parking + " " + match.Groups[1].Value;



            } else

            {
                pattern = "Spaces:Gar:(\\d+)";
                match = Regex.Match(s, pattern);
                parking = parking + " " + match.Groups[1].Value;


                //Parking Type:xxxParking Avail
                //Parking Type:xxxDetached Garage
                pattern = "Parking Type:x*([^x\\n]+)";
                match = Regex.Match(s, pattern);
                parking = parking = parking + " " + match.Groups[1].Value;
            }

            //pattern = @"xxxStyle:(.*?)xxx|xxxStyle:(.*?)\n";
            pattern = @"Style:x*([^x\nG]+)x*";
            match = Regex.Match(s, pattern);
            style = match.Groups[1].Value;

            //required field, except in archive listings, ie
            //Type:xxxGarage Ownership:xxxSewer:
            //pattern = "Type:x*([^x\\n]+)xxx";
            //above patern fails
            pattern = @"Type:x*([^x\nG]+)x*";
            
            match = Regex.Match(s, pattern);
            string type = match.Groups[1].Value;

            if (!match.Success)
            {
                //bad print from pdf convertor
                pattern = @"Tx*yx*px*e:x*(Tx*ownhouse)";
                match = Regex.Match(s, pattern);
                type = match.Groups[1].Value.Replace("x", "");
            }
            
       
            form.SubjectBathroomCount = bathCount;
            form.SubjectBedroomCount = bedroomCount;
            form.SubjectRoomCount = roomCount;
            form.SubjectBasementType = basementType;
            form.SubjectBasementDetails = basementDetails;
            form.SubjectNumberFireplaces = numFireplaces;
            form.SubjectParkingType = parking;
            form.SubjectStyle = style;
            form.SubjectMlsType = type;
            form.SubjectAttached = attached;
            form.SubjectDetached = detached;
            form.SubjectAssessmentInfo.amount = assessmentAmount;
            form.SubjectAssessmentInfo.frequency = assessmentFrequncy;
            form.SubjectAssessmentInfo.includes = assessmentIncludes;
            form.SubjectAssessmentInfo.managementContact = assessmentContact;
            form.SubjectAssessmentInfo.managementPhone = assessmentContactPhone;
            form.SubjectMlsStatus = mlsStatus;
            form.SubjectCurrentListPrice = listPrice;
            form.SubjectListingAgent = listAgent;
            form.SubjectBrokerPhone = brokerPhone;
            form.SubjectDom = dom;
            GlobalVar.theSubjectProperty.PropertyTax = mlsTaxAmount;

            m.MlsNumber = GetFieldValue(@"MLS #:");
            m.ListDateString = GetFieldValue(@"List Date:");
            m.OffMarketDateString = GetFieldValue(@"Off Market:");

            GlobalVar.theSubjectProperty.AddMlsListing(m);

            GlobalVar.theSubjectProperty.ParcelID = Regex.Match(GetFieldValue("PIN:"), @"\d+").Value;

           

            form.SubjectPin = GlobalVar.theSubjectProperty.ParcelID;

           

           //xxxxxxx

         
        }


        public async void ReadMlsSheets(iMacros.App d)
        {


            form.StatusUpdate = this.ToString() + "-->ReadMlsSheets";


            // NumberFormatInfo provider = new NumberFormatInfo( );
           
            
            GoogleFusionTable realist_bpohelper = new GoogleFusionTable("1UKrOVmhPWrgLP5d5bDCsiW9whMIK8aLxKhcyOaI");

            await realist_bpohelper.helper_OAuthFusion();

            
            
            //GoogleFusionTable mlsListingCache = new GoogleFusionTable("1EDFPE91a2_6oohUmEG-2OEr_k2xFOG7c9x3fqZE");

            GoogleCloudDatastore gcds = new GoogleCloudDatastore();
            await gcds.DataStoreTester();


         
           
            List<string> comps = new List<String>();
            List<MLSListing> listings = new List<MLSListing>();
            Status status = iMacros.Status.sOk;
            StringBuilder move_through_comps_macro = new StringBuilder();
            RealProperty currentProperty = new RealProperty();
            Neighborhood currentNeighborhood = new Neighborhood();
            List<RealProperty> rpl = new List<RealProperty>();

            //extracts  all the html
            form.SetStatusBar = "Reading MLS sheet...";

            StringBuilder macro12 = new StringBuilder();
            macro12.AppendLine(@"FRAME NAME=workspace");
            macro12.AppendLine(@"TAG POS=1 TYPE=DIV FORM=NAME:dc ATTR=CLASS:report EXTRACT=HTM");

            string htmlCode;// = d.iimGetLastExtract();

            MLSListing currentListing; // = new MLSListing(htmlCode);
            
           


            List<decimal> activePriceList = new List<decimal>();
            List<decimal> soldPriceList = new List<decimal>();
            List<decimal> picDiffList = new List<decimal>();
            List<int> domList = new List<int>();
            
            int shortSales = 0;
            int reoSales = 0;
            int shortActive = 0;
            int reoActive = 0;
            int numberSameTypeAsSubject = 0;
            int numberComparableGla = 0;
            int numberComparableAge = 0;
            int numberComparableLot = 0;

          


            move_through_comps_macro.AppendLine(@"FRAME NAME=navpanel");
            move_through_comps_macro.AppendLine(@"TAG POS=1 TYPE=IMG ATTR=SRC:http://connectmls*.mredllc.com/images/next.gif");
            //move_through_comps_macro.AppendLine(@"FRAME NAME=workspace");
            //move_through_comps_macro.AppendLine(@"TAG POS=1 TYPE=DIV FORM=NAME:dc ATTR=CLASS:report EXTRACT=HTM");
          
         
           

            string min_sale;
            string max_sale;
            string min_list;
            string max_list;
            int active_comps = 0;
            int closed_comps = 0;
            int pending_comps = 0;
            ArrayList list = new ArrayList();
            //IEnumerable<decimal> myEnumerable;
            List<decimal> pricePerSfList = new List<decimal>();
            List<int> ageList = new List<int>();
            Dictionary<string,int> subdivisionList = new Dictionary<string,int>();
            Dictionary<string, int> mlsTypeList = new Dictionary<string, int>();
            Dictionary<string, int> listingAgentList = new Dictionary<string, int>();

            string attOrDet = "Unknown";

            bool stillRecords = true;
            bool searchCache = false;
            int count = 0;


            GlobalVar.searchCacheMlsListings.Clear();
            GlobalVar.searchCacheMlsListings = GlobalVar.listingsFromLastSearch;
           


            //if (GlobalVar.searchCacheMlsListings.Count > 0)
            //{
            //    searchCache = true;
            //}

            if (form.CacheSearch)
            {
                searchCache = true;
            }

            iMacros.Status s = status;

            //first record
            //s = d.iimPlayCode(macro12.ToString());

            //while (status == iMacros.Status.sOk)
            //for (int i = 0; i < 6; i++)
            while(stillRecords)
            {
               

                //
                //this gets the all html code for current listing page
                //
                if (searchCache)
                {
                    htmlCode = GlobalVar.searchCacheMlsListings[count].rawData; 
                    
                }
                else
                {
                    s = d.iimPlayCode(macro12.ToString());
                    htmlCode = d.iimGetLastExtract();
                    if (htmlCode.Contains("#EANF#"))
                    {
                        break;
                    }
                }
                
                //

                //
                //this creates correct listing object and fills/sets known field/value pairs based on above html
                //
                 if (form.SubjectAttached)
                 {
                     currentListing = new AttachedListing(htmlCode);
                     //currentListing.proximityToSubject = Convert.ToDouble(Get_Distance(currentListing.mlsHtmlFields["address"].value, form.SubjectFullAddress));
                     listings.Add(currentListing);
                 }
                 else if (form.SubjectDetached)
                 {
                     //constructor parses and sets all field/value pairs
                     currentListing = new DetachedListing(htmlCode);
                     listings.Add(currentListing);
                     ////do we have the realist record for the parcel number, if so, do we have lng lat pair stored?
                     //if (realist_bpohelper.RecordExists(currentListing.mlsHtmlFields["Parcel ID:"].value))
                     //{
                        
                     //}
                     //else
                     //{

                     //    try
                     //    {
                     //        currentListing.proximityToSubject = Convert.ToDouble(Get_Distance(currentListing.mlsHtmlFields["address"].value, form.SubjectFullAddress));
                     //    }
                     //    catch
                     //    {
                     //        currentListing.proximityToSubject = -1;
                     //    }
                     //    listings.Add(currentListing);
                     //}
                 }
                 else
                 {
                     currentListing = new MLSListing(htmlCode);
                     currentListing.proximityToSubject = Convert.ToDouble(Get_Distance(currentListing.mlsHtmlFields["address"].value, form.SubjectFullAddress));
                     listings.Add(currentListing);
                 }

                 //htmlcode is too big
                 //mlsListingCache.AddMlsRecord(currentListing.MlsNumber, htmlCode);
                
                 currentProperty.AddMlsListing(currentListing);


               //  string tempdata = currentListing.rawData;

                // currentListing.rawData = "";


               

                

                
                 form.StatusUpdate = "Reading--> " + currentListing.MlsNumber;
                #region read_mls_sheet
                #region Table1
                //
                //Table 1
                //
                string pattern = "";
                Match match;
                //// Detached Single#NEXT#MLS #:#NEXT#07933784#NEXT#List Price:#NEXT#$150,000#NEXT##NEWLINE#Status:#NEXT#CLSD#NEXT#List Date:#NEXT#10/27/2011#NEXT#Orig List Price:#NEXT#$150,000#NEXT##NEWLINE#Area:#NEXT#50  #NEXT#List Dt Rec:#NEXT#10/28/2011#NEXT#Sold Price:#NEXT#$140,000#NEXT##NEWLINE#Address:#NEXT#2316 W Fairview Ln , McHenry, Illinois 60051#NEXT##NEWLINE#Directions:#NEXT#Rt. 120 to River Rd N. to Lincoln E to Hillside to Fairview E#NEXT##NEWLINE#Sold by:#NEXT#Ed Kanabay (16107) / Results Realty USA (1886) #NEXT#Lst. Mkt. Time:#NEXT#18#NEXT##NEWLINE#Closed:#NEXT#02/15/2012#NEXT#Contract:#NEXT#11/13/2011#NEXT#Points:#NEXT##NEXT##NEWLINE#Off Market:#NEXT#11/13/2011#NEXT#Financing:#NEXT#Conventional#NEXT#Contingency:#NEXT##NEXT##NEWLINE#Year Built:#NEXT#1977#NEXT#Blt Before 78:#NEXT#Yes#NEXT#Curr. Leased:#NEXT#No#NEXT##NEWLINE#Dimensions:#NEXT#80X150#NEXT##NEWLINE#Ownership:#NEXT#Fee Simple#NEXT#Subdivision:#NEXT#Eastwood Manor#NEXT#Model:#NEXT##NEXT##NEWLINE#Corp Limits:#NEXT#Unincorporated#NEXT#Township:#NEXT#Mchenry#NEXT#County:#NEXT#Mc Henry#NEXT##NEWLINE#Coordinates:#NEXT#N:33 W:31 #NEXT## Fireplaces:#NEXT##NEXT##NEWLINE#Rooms:#NEXT#8#NEXT#Bathrooms (full/half):#NEXT#3 / 0#NEXT#Parking:#NEXT#Garage#NEXT##NEWLINE#Bedrooms:#NEXT#3#NEXT#Master Bath:#NEXT#Full#NEXT## Spaces:#NEXT#Gar:2 #NEXT##NEWLINE#Basement:#NEXT#None#NEXT#Bsmnt. Bath:#NEXT#No#NEXT#Parking Incl. In Price:#NEXT#Yes
                //string pattern = "MLS #:#NEXT#(\\d+)#NEXT#";
                //Match match = Regex.Match(table1, pattern);
                //string mlsnum = match.Groups[1].Value;
                //
                //pattern = "Subdivision:#NEXT#([^#]+)";
                //match = Regex.Match(table1, pattern);
                //string mls_subdivision = match.Groups[1].Value;
                //doesnt match $1,000,000
                //pattern = "List Price:#NEXT#(.[0-9]+.[0-9]+)";
                //pattern = "List Price:#NEXT#([^#]+)";
                //match = Regex.Match(table1, pattern);
                //string current_list_price = match.Groups[1].Value;
                //
                //pattern = "List Date:#NEXT#(\\d+.\\d+.\\d+)#NEXT#";
                //match = Regex.Match(table1, pattern);
                //string list_date = match.Groups[1].Value;
                //

                //pattern = "Orig List Price:#NEXT#([^#]+)";
                //match = Regex.Match(table1, pattern);
                //string orig_list_price = match.Groups[1].Value;

                //pattern = "Sold Price:#NEXT#([^#]+)";
                //match = Regex.Match(table1, pattern);
                //string sold_price = match.Groups[1].Value;
                //Old way above
                //
                //New Way below
                //

                string mlsnum = currentListing.MlsNumber;
                string mls_subdivison = currentListing.mlsHtmlFields["subdivision"].value;
                string current_list_price = currentListing.mlsHtmlFields["listPrice"].value;
                string list_date = currentListing.mlsHtmlFields["listDate"].value;
                string orig_list_price = currentListing.mlsHtmlFields["origListPrice"].value;


                string sold_price = currentListing.mlsHtmlFields["soldPrice"].value;
                if (String.IsNullOrWhiteSpace(sold_price))
                    sold_price = "0";

                string sale_type = "Normal";
                if (sold_price.Contains("(S)"))
                {
                    sale_type = "Short";
                    sold_price = sold_price.Replace("(S)", "");
                    shortSales++;
                }

                if (sold_price.Contains("(F)"))
                {
                    sale_type = "REO";
                    sold_price = sold_price.Replace("(F)", "");
                    reoSales++;
                }

                if (sold_price.Contains("(C)"))
                {
                    sale_type = "Court";
                    sold_price = sold_price.Replace("(C)", "");
                    reoSales++;
                }

                #region address

                //fuill line address
                //Address:#NEXT#2627 Sycamore Dr , Waukegan, Illinois 60085#NEXT#

               // pattern = "Address:#NEXT#([^#]+)#NEXT#";
               // match = Regex.Match(table1, pattern);
                //string address = match.Groups[1].Value;
                string address = currentListing.mlsHtmlFields["address"].value;
                string[] tempstrarry = address.Split(',');
                string street_number = "";
                string street_name = "";
                string street_postfix = "";
                string city = tempstrarry[1];
                string zip = tempstrarry[2].Split(' ')[2];
                string full_street_address = tempstrarry[0];
                string[] street_combo = tempstrarry[0].Split(' ');


                ////
                ////geocode address
                ////ex:
                ////http://maps.googleapis.com/maps/api/geocode/json?address=1600+Amphitheatre+Parkway,+Mountain+View,+CA&sensor=true_or_false


                //string googlestr = @" http://maps.googleapis.com/maps/api/geocode/json?address=" + address.Replace(" ", "+") + "&sensor=false";


                ////// Create a request for the URL. 		
                //WebRequest request = WebRequest.Create(googlestr);

                ////// Get the response.
                //HttpWebResponse response = (HttpWebResponse)request.GetResponse();


                //// Get the stream containing content returned by the server.
                //Stream dataStream = response.GetResponseStream();
                //// Open the stream using a StreamReader for easy access.
                //StreamReader reader = new StreamReader(dataStream);
                //// Read the content.
                //string responseFromServer = reader.ReadToEnd();
                //// Display the content.
                ////MessageBox.Show(response.ContentLength.ToString());
                ////MessageBox.Show(responseFromServer);
                //////Console.WriteLine(responseFromServer);
                ////// Cleanup the streams and the response.

                //reader.Close();
                //dataStream.Close();
                //response.Close();


                ////ex response:
                ////  "formatted_address" : "107 Glenwood Dr, Round Lake Beach, IL 60073, USA",
                ////  "geometry" : {
                ////   "location" : {
                ////   "lat" : 42.3653120,
                ////   "lng" : -88.08547999999999

                //pattern = @"lat. : (\d+.\d+)";
                //match = Regex.Match(responseFromServer, pattern);
                //string propertyLatitude =  match.Groups[1].Value;

                // pattern = @"lng. : (-\d+.\d+)";
                // match = Regex.Match(responseFromServer, pattern);
                //string propertyLongitude =  match.Groups[1].Value;

               // string propertyLatitude = realist_bpohelper.curRec["Latitude"];
              //  string propertyLongitude = realist_bpohelper.curRec["Longitude"];





                //need to remove unit in attached listing addresses.
                if (tempstrarry[0].Contains("Unit"))
                {
                    tempstrarry[0] = tempstrarry[0].Remove(tempstrarry[0].IndexOf("Unit"));
                }




                //2316 W Fairview Ln , McHenry, Illinois 60051
                pattern = "^(\\d+)\\s+\\w\\s+(\\w+)\\s+(\\w+)";
                match = Regex.Match(tempstrarry[0], pattern);

                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;

                }

                //2316 Fairview Ln , McHenry, Illinois 60051
                pattern = "^(\\d+)\\s+(\\w\\w+)\\s+(\\w+)";
                match = Regex.Match(tempstrarry[0], pattern);
                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;
                    if (street_postfix.Contains("Bay"))
                    {
                        street_postfix = "";
                    }

                }

                //1309 N Chapel Hill Rd , McHenry, Illinois 60051
                pattern = "^(\\d+)\\s+\\w\\s+(\\w+\\s+\\w+)\\s+(\\w+)";
                match = Regex.Match(tempstrarry[0], pattern);
                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;
                }


                //769 White Pine Cir , Lake In The Hills, Illinois 60156
                pattern = "^(\\d+)\\s+(\\w\\w+\\s+\\w+)\\s+(\\w+)[^\\w+]";
                match = Regex.Match(tempstrarry[0], pattern);
                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;
                }

                // 21727 NORTH Hickory Hill Dr 
                pattern = "^(\\d+)\\s+NORTH\\s+(\\w\\w+\\s+\\w+)\\s+(\\w+)";
                match = Regex.Match(tempstrarry[0], pattern);
                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;
                }



                //if (street_combo.Length == 4)
                //{
                //    //2316 W Fairview Ln , McHenry, Illinois 60051
                //    pattern = "^(\\d+)\\s+\\w+\\s+(\\w+)\\s+(\\w+)\\s+,\\s+(\\w+),\\s+\\w+\\s+(\\d+)";
                //    match = Regex.Match(address, pattern);
                //    street_number = match.Groups[1].Value;
                //     street_name = match.Groups[2].Value;
                //     street_postfix = match.Groups[3].Value;
                //     city = match.Groups[4].Value;
                //     zip = match.Groups[5].Value;
                //}
                //else if (street_combo.Length == 5)
                //{
                //    //1309 N Chapel Hill Rd , McHenry, Illinois 60051
                //    pattern = "^(\\d+)\\s+\\w+\\s+(\\w+\\s+\\w+)\\s+(\\w+)\\s+,\\s+(\\w+),\\s+\\w+\\s+(\\d+)";
                //    match = Regex.Match(address, pattern);
                //    street_number = match.Groups[1].Value;
                //    street_name = match.Groups[2].Value;
                //    street_postfix = match.Groups[3].Value;
                //    city = match.Groups[4].Value;
                //    zip = match.Groups[5].Value;

                //}
                //else
                //{

                //    //2316 Fairview Ln , McHenry, Illinois 60051
                //    pattern = "^(\\d+)\\s+(\\w+)\\s+(\\w+)\\s+,\\s+(\\w+),\\s+\\w+\\s+(\\d+)";
                //    match = Regex.Match(address, pattern);
                //    street_number = match.Groups[1].Value;
                //    street_name = match.Groups[2].Value;
                //    street_postfix = match.Groups[3].Value;
                //    city = match.Groups[4].Value;
                //    zip = match.Groups[5].Value;
                //}

                #endregion

                string dom = currentListing.mlsHtmlFields["daysOnMarket"].value;
                string closed_date = currentListing.mlsHtmlFields["closedDate"].value;
                string year_built = currentListing.mlsHtmlFields["yearBulit"].value;
                if (year_built == "NEW" || year_built == "0000" || year_built == "NC")
                {
                    try
                    {
                        year_built = currentListing.SalesDate.Year.ToString();
                    } 
                    catch
                    {
                        year_built = DateTime.Now.Year.ToString();
                    }

                    currentListing.mlsHtmlFields["yearBulit"].value = year_built;
                }
                else if (year_built == "UNK" || year_built == "UNKN"|| String.IsNullOrWhiteSpace(year_built) || !Regex.IsMatch(year_built,@"\d\d\d\d"))
                {
                    //blank it, for realist extract logic to work below
                    currentListing.mlsHtmlFields["yearBulit"].value = "";
                    year_built = "";
                }
                string room_count = currentListing.mlsHtmlFields["roomCount"].value;
    
                pattern = @"(\d+)\s*.\s*(\d+)";
                match = Regex.Match(currentListing.mlsHtmlFields["bathrooms"].value, pattern);
                string full_bath = match.Groups[1].Value;
                string half_bath = match.Groups[2].Value;
    
                string bedrooms = currentListing.mlsHtmlFields["bedrooms"].value;
                pattern = "\\s*(\\d+)";
                match = Regex.Match(bedrooms, pattern);
  

                if (!match.Success)
                {
                    //3 + 1 below grade
                    pattern = "\\s*(\\d+)\\s+\\d[^#]+";
                    match = Regex.Match(bedrooms, pattern);
                    bedrooms = match.Groups[1].Value;
                }

                string   number_of_firplaces = currentListing.mlsHtmlFields["numFireplaces"].value;
                string basement = "None";
                basement = currentListing.mlsHtmlFields["basement"].value;
                string parking = currentListing.mlsHtmlFields["parking"].value;

                string mls_number_of_spaces = currentListing.mlsHtmlFields["numSpaces"].value;

                //pattern = "Lst. Mkt. Time:#NEXT#(\\d+)#NEXT#";
               // match = Regex.Match(table1, pattern);
               // string dom = match.Groups[1].Value;

               // pattern = "Closed:#NEXT#(\\d+.\\d+.\\d+)#NEXT#";
               // match = Regex.Match(table1, pattern);
               // string closed_date = match.Groups[1].Value;


                //pattern = "Year Built:#NEXT#(\\d+)#NEXT#";
               // match = Regex.Match(table1, pattern);
               // string year_built = match.Groups[1].Value;
               


                //pattern = "Rooms:#NEXT#(\\d+)#NEXT#";
                //match = Regex.Match(table1, pattern);
                //string room_count = match.Groups[1].Value;

                ////attached 0 space
                ////Bathrooms (Full/Half):#NEXT#1/1#NEXT#
                ////Bathrooms (Full/Half):#NEXT#1/1#NEXT#
                //pattern = "Bathrooms \\([Ff]ull\\/[Hh]alf\\):#NEXT#(\\d+)\\s*.\\s*(\\d+)#NEXT#";
                //match = Regex.Match(table1, pattern);
                //string full_bath = match.Groups[1].Value;
                //string half_bath = match.Groups[2].Value;

                //pattern = "Bedrooms:#NEXT#(\\d+)";
                //match = Regex.Match(table1, pattern);
                //string bedrooms = match.Groups[1].Value;

                //if (!match.Success)
                //{
                //    //3 + 1 below grade
                //    pattern = "Bedrooms:#NEXT#(\\d+)\\s+\\d[^#]+";
                //    match = Regex.Match(table1, pattern);
                //    bedrooms = match.Groups[1].Value;
                //}
                //pattern = "# Fireplaces:#NEXT#(\\d+)#NEXT#";
                //match = Regex.Match(table1, pattern);
                //string number_of_firplaces = match.Groups[1].Value;

                ////Full, English
                //pattern = "Basement:#NEXT#([^#]+)";
                //match = Regex.Match(table1, pattern);
                //string basement = "None";
                //basement = match.Groups[1].Value;

                //pattern = "Parking:#NEXT#(\\w+)#NEXT#";
                //match = Regex.Match(table1, pattern);
                //string parking = match.Groups[1].Value;


                //pattern = "# Spaces:#NEXT#\\w+:(\\d)";  //# Spaces:#NEXT#Gar:2 #NEXT#
                //match = Regex.Match(table1, pattern);
                //string mls_number_of_spaces = match.Groups[1].Value;
                


                #endregion


                //
                //Table4:  Details
                //
                 
               // pattern = "#Type:(\\d*\\s*\\w+)#NEXT#";
                string type = currentListing.mlsHtmlFields["mlsType"].value;
                string mls_garage_type = currentListing.mlsHtmlFields["garageType"].value;
                
                pattern = "Gar:(\\d)";
                match = Regex.Match(mls_number_of_spaces, pattern);
                string mls_garage_spaces = match.Groups[1].Value;
                string basementDetails = currentListing.mlsHtmlFields["basementDetails"].value;

                //pattern = "#Type:([^#]+)";
                //match = Regex.Match(table4, pattern);
                //string type = match.Groups[1].Value;

                //pattern = "Garage Type:(([^#]+))";  //pattern = "Lot Acres:#NEXT#(([^#]+))";
                //match = Regex.Match(table4, pattern);
                //string mls_garage_type = match.Groups[1].Value;

              

                pattern = "Finished";  //pattern = "Lot Acres:#NEXT#(([^#]+))";
                match = Regex.Match(basementDetails, pattern);
                bool finished_basement = false;
                if (match.Success)
                {
                    finished_basement = true;
                }


                //
                //end table 4
                //
               

                pattern = "\\s*(\\d+)\\s*";
                match = Regex.Match(currentListing.mlsHtmlFields["pin"].value, pattern);
                //string pin = table2.Substring(table2.IndexOf("PIN:") + 4, 10);
                string pin = match.Groups[1].Value;

                string mls_gla = currentListing.mlsHtmlFields["mlsGla"].value;

                if (string.IsNullOrEmpty(mls_gla))
                {
                    mls_gla = "0";
                }


                //pattern = "Appx SF:#NEXT#(\\d+)#NEXT#";
                //match = Regex.Match(table2, pattern);
                ////string pin = table2.Substring(table2.IndexOf("PIN:") + 4, 10);
                //string mls_gla = match.Groups[1].Value;

                ////Attached MLS sheet in table 1
                //if (!match.Success)
                //{
                //    match = Regex.Match(table1, pattern);
                //    if (match.Success)
                //    {
                //        mls_gla = match.Groups[1].Value;
                //    }
                //    else
                //    {
                //        mls_gla = "0";
                //    }

                //}{

                string mls_lot_size = null;

                try
                {
                     mls_lot_size = currentListing.mlsHtmlFields["acerage"].value;
                }
                catch
                {
                   
                }
               
                if (string.IsNullOrEmpty(mls_lot_size))
                {
                    mls_lot_size = "0";
                }


                //pattern = "Acreage:#NEXT#(\\d+.\\d+)";
                //match = Regex.Match(table2, pattern);
                ////string pin = table2.Substring(table2.IndexOf("PIN:") + 4, 10);
                //string mls_lot_size = "0";
                //mls_lot_size = match.Groups[1].Value;

                //if (!match.Success)
                //{
                //    pattern = "Acreage:#NEXT#(\\d+)";
                //    match = Regex.Match(table2, pattern);
                //    mls_lot_size = match.Groups[1].Value;

                //}

                //
                //listing agent info
                //


                string listing_Agent = currentListing.mlsHtmlFields["listAgent"].value;

                //pattern = "List Agent:#NEXT#([^\\(]+)";
                //match = Regex.Match(table5, pattern);
                //string listing_Agent = "";
                //listing_Agent = match.Groups[1].Value;

                if (listingAgentList.ContainsKey(listing_Agent))
                {
                    listingAgentList[listing_Agent]++;
                }
                else
                {
                    try
                    {
                        listingAgentList.Add(listing_Agent, 1);
                    }
                    catch
                    {
                        MessageBox.Show("Problem adding to listing agent list: " + listing_Agent);
                    }
                }

                string broker = currentListing.mlsHtmlFields["broker"].value;
               
                //pattern = "Broker:#NEXT#([^\\(]+)";
                //match = Regex.Match(table5, pattern);
                //string broker = "";
                //broker = match.Groups[1].Value;
                string phone = currentListing.mlsHtmlFields["phone"].value;

                //pattern = "Ph #:#NEXT#([^#]+)";
                //match = Regex.Match(table5, pattern);
                //string phone = "";
                //phone = match.Groups[1].Value;
                string additionalSalesInfo = currentListing.mlsHtmlFields["additionalSalesInfo"].value;

                ////pattern = "Additional Sales Information:#NEXT#([^#]+)|Addl. Sales Info.:#NEXT#([^#]+)";
                ////match = Regex.Match(table5, pattern);
                ////string additionalSalesInfo = "";
                ////additionalSalesInfo = match.Groups[1].Value;

                //if (additionalSalesInfo == "")
                //{
                //    additionalSalesInfo = match.Groups[2].Value;
                //}


                string mls_status = currentListing.mlsHtmlFields["status"].value;

                //pattern = @"Status:#NEXT#(\w+)#NEXT#";
                //match = Regex.Match(table1, pattern);
                //string mls_status = match.Groups[1].Value;

                if (mls_status != "CLSD" && additionalSalesInfo.Contains("Short Sale"))
                {
                    shortActive++;
                }

                if (mls_status != "CLSD" && additionalSalesInfo.Contains("REO"))
                {
                    reoActive++;
                }

                ////Google distance matric webservice
                ////http://maps.googleapis.com/maps/api/distancematrix/xml?origins=60050&destinations=60002&sensor=false&units=imperial
                ////

                //string googlestr = "http://maps.googleapis.com/maps/api/distancematrix/xml?origins=60050&destinations=" + zip + "&sensor=false&units=imperial";

                //// Create a request for the URL. 		
                //WebRequest request = WebRequest.Create(googlestr);

                //// Get the response.
                //HttpWebResponse response = (HttpWebResponse)request.GetResponse();


                //// Get the stream containing content returned by the server.
                //Stream dataStream = response.GetResponseStream();
                //// Open the stream using a StreamReader for easy access.
                //StreamReader reader = new StreamReader(dataStream);
                //// Read the content.
                //string responseFromServer = reader.ReadToEnd();
                //// Display the content.
                ////MessageBox.Show(response.ContentLength.ToString());
                ////MessageBox.Show(responseFromServer);
                ////Console.WriteLine(responseFromServer);
                //// Cleanup the streams and the response.
                //reader.Close();
                //dataStream.Close();
                //response.Close();







                #endregion
                form.StatusUpdate = "Done.";

                #region Price Change extraction
                //
                //Price Change extraction
                //
                //StringBuilder macro_get_price_changes = new StringBuilder();

                //macro_get_price_changes.AppendLine(@"FRAME NAME=workspace");
                //macro_get_price_changes.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:Additional<SP>Information");
                //macro_get_price_changes.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:Listing<SP>&<SP>Property<SP>History");
                //macro_get_price_changes.AppendLine(@"'New tab opened");
                //macro_get_price_changes.AppendLine(@"TAB T=2");
                //macro_get_price_changes.AppendLine(@"TAG POS=1 TYPE=TABLE FORM=NAME:NoFormName ATTR=CLASS:innergridview EXTRACT=TXT");
                //macro_get_price_changes.AppendLine(@"FRAME NAME=subheader");
                //macro_get_price_changes.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:header ATTR=ID:closeWinLink");
                //// macro_get_price_changes.AppendLine(@"TAB CLOSE");
                //string macroCode = macro_get_price_changes.ToString();

                //d.iimPlayCode(macroCode, 60);

                //string price_change_table = d.iimGetLastExtract(1);


                //int count = new Regex("-> PCHG").Matches(price_change_table).Count;



                ////string searchTerm = "PCHG";
                //////Convert the string into an array of words
                ////string[] source = price_change_table.Split(new char[] { '.', '?', '!', ' ', ';', ':', ',' }, StringSplitOptions.RemoveEmptyEntries);

                ////// Create and execute the query. It executes immediately 
                ////// because a singleton value is produced.
                ////// Use ToLowerInvariant to match "data" and "Data" 
                ////var matchQuery = from word in source
                ////                 where word.ToLowerInvariant() == searchTerm.ToLowerInvariant()
                ////                 select word;

                ////// Count the matches.
                ////int wordCount = matchQuery.Count();


                #endregion


                #region image_compare

                //StringBuilder save_pics_macro = new StringBuilder();

                ////string filepath = "C:\\Users\\Scott\\Documents\\My<SP>Dropbox\\BPOs\\" + search_address_textbox.Text;
                //string filepath = form.SubjectFilePath;
                //string filename = "comp.jpg";

               
                //    save_pics_macro.AppendLine(@"ONDOWNLOAD FOLDER=" + filepath.Replace(" ", "<SP>") + " FILE=comp.jpg");
                
                //save_pics_macro.AppendLine(@"FRAME NAME=workspace");
                //save_pics_macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:dc ATTR=SRC:*/PICS/*");
                //save_pics_macro.AppendLine(@"'New tab opened");
                ////save_pics_macro.AppendLine(@"WAIT SECONDS=5");
                //save_pics_macro.AppendLine(@"TAB T=2");
                //save_pics_macro.AppendLine(@"FRAME F=0");
                //save_pics_macro.AppendLine(@"TAG POS=2 TYPE=IMG FORM=NAME:dc ATTR=HREF:""*.JPEG"" CONTENT=EVENT:SAVEITEM");
                ////save_pics_macro.AppendLine(@"WAIT SECONDS=5");
                //save_pics_macro.AppendLine(@"TAB CLOSE");

                //string save_pics_macroCode = save_pics_macro.ToString();
                //status = d.iimPlayCode(save_pics_macroCode, 30);

                ////get the full path of the images
                //string image1Path = Path.Combine(filepath, "front.JPG");
                
                //string image2Path = Path.Combine(filepath, "comp.jpg");

                ////compare the two
                ////Console.Write("Comparing: " + bmp1 + " and " + bmp2 + ", with a threshold of " + threshold);
                //Bitmap firstBmp = (Bitmap)System.Drawing.Image.FromFile(image1Path);
                //Bitmap secondBmp = (Bitmap)System.Drawing.Image.FromFile(image2Path);
                //form.SubjectPic = firstBmp;
                //form.CompPic = new Bitmap(secondBmp);
                //form.PicDiffLabel = (firstBmp.PercentageDifference(secondBmp, 3) * 100).ToString();
                //picDiffList.Add(Convert.ToDecimal(firstBmp.PercentageDifference(secondBmp, 3) * 100));


                //using (secondBmp)
                //{
                //    //save the diffgram
                //    firstBmp.GetDifferenceImage(secondBmp, true).Save(image1Path + "_diff.png");
                //    //MessageBox.Show((string.Format("Difference: {0:0.0} %", firstBmp.PercentageDifference(secondBmp, 3) * 100)));

                //}
                #endregion

                form.StatusUpdate = "Reading Realist--> " + currentListing.MlsNumber;
                #region realist extraction

                string realist_rawHtml = "";
                string realist_gla = "0";
                string realist_subdivision = "";
                string censusTract = "";
                string realist_schoolDistrict = "";
                string realist_lotAcres = "0";
                string realist_address = "";

                if (form.UpdateRealist || !realist_bpohelper.RecordExists(pin))
                {
                    #region realist extraction update/add

                    #region open realist and extract data script

                    StringBuilder realist_extraction_macro = new StringBuilder();

                    //open realist
                    //realist_extraction_macro.AppendLine(@"SET !ERRORIGNORE YES");
                    realist_extraction_macro.AppendLine(@"SET !TIMEOUT_STEP 0");
                    realist_extraction_macro.AppendLine(@"FRAME NAME=workspace");
                    realist_extraction_macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:Realist<SP>Tax<SP>Report");

                    string realist_extraction_macro_code = realist_extraction_macro.ToString();
                    status = d.iimPlayCode(realist_extraction_macro_code, 60);
                    realist_extraction_macro.Clear();

                    if (status == Status.sOk)
                    {
                        //see what we have opened
                        s = d.iimPlayCode("SET !TIMEOUT_STEP 12 \r\n TAG POS=1 TYPE=TABLE ATTR=* EXTRACT=TXT");
                        realist_rawHtml = d.iimGetLastExtract();
                        //realist window is open but no properties found so skip the steps below
                        if (!realist_rawHtml.Contains("No properties were found."))
                        {
                            //gets all html
                            s = d.iimPlayCode("SET !TIMEOUT_STEP 12 \r\n TAG POS=1 TYPE=DIV ATTR=ID:htmlApplication EXTRACT=HTM");
                            realist_rawHtml = d.iimGetLastExtract();
                            //gets the address
                            s = d.iimPlayCode("SET !TIMEOUT_STEP 0 \r\n TAG POS=1 TYPE=DIV ATTR=ID:headerText EXTRACT=TXT");
                            realist_address = d.iimGetLastExtract();
                        }
                       
                    }

                   
                    if (realist_rawHtml.Contains("Parcel"))
                    {

                        #region goodRecord
                        //Owner Information
                        realist_extraction_macro.AppendLine(@"TAG POS=1 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");
                        realist_extraction_macro.AppendLine(@"SET !TIMEOUT_STEP 1");
                        //Location Information
                        realist_extraction_macro.AppendLine(@"TAG POS=2 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");

                        //Tax Information Table
                        realist_extraction_macro.AppendLine(@"TAG POS=3 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");

                        //Assessment & Tax
                        //Assessment
                        realist_extraction_macro.AppendLine(@"TAG POS=1 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");
                        //Tax
                        realist_extraction_macro.AppendLine(@"TAG POS=2 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");


                        //Characteristics
                        realist_extraction_macro.AppendLine(@"TAG POS=4 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");


                        //Estimated Value
                        realist_extraction_macro.AppendLine(@"TAG POS=5 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");

                        //Listing Information
                        //realist_extraction_macro.AppendLine(@"TAG POS=6 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");
                        //most recent mls listing
                        //realist_extraction_macro.AppendLine(@"TAG POS=3 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");

                        //Last Market Sale & Sales History
                        //first row
                        //realist_extraction_macro.AppendLine(@"TAG POS=4 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");
                        //second row
                        //realist_extraction_macro.AppendLine(@"TAG POS=5 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");

                        //Mortgage History
                        //first row
                        //realist_extraction_macro.AppendLine(@"TAG POS=6 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");
                        //second row
                        //  realist_extraction_macro.AppendLine(@"TAG POS=7 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");

                        //Foreclosure History
                        //   realist_extraction_macro.AppendLine(@"TAG POS=8 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");



                        //realist_extraction_macro.AppendLine(@"TAB CLOSE");
                        realist_extraction_macro_code = realist_extraction_macro.ToString();

                        status = d.iimPlayCode(realist_extraction_macro_code, 60);



                        string realist_owner_table = d.iimGetLastExtract(1);
                        string realist_location_table = d.iimGetLastExtract(2);
                        string realist_taxinfo_table = d.iimGetLastExtract(3);
                        string realist_assessment_table = d.iimGetLastExtract(4);
                        string realist_tax_table = d.iimGetLastExtract(5);
                        string realist_char_table = d.iimGetLastExtract(6);
                        string realist_estimatedvalue_table = d.iimGetLastExtract(7);
                        string realist_listinginfo1_table = d.iimGetLastExtract(8);
                        string realist_listinginfo2_table = d.iimGetLastExtract(9);
                        string realist_saleshisotry1_table = d.iimGetLastExtract(10);
                        string realist_saleshisotry2_table = d.iimGetLastExtract(11);
                        string realist_morthisotry1_table = d.iimGetLastExtract(12);
                        string realist_morthisotry2_table = d.iimGetLastExtract(13);
                        string realist_foreclosure_table = d.iimGetLastExtract(14);


                        //  s = d.iimPlayCode(@"TAG POS=1 TYPE=DIV ATTR=ID:headerText EXTRACT=TXT");
                        //  realist_address = d.iimGetLastExtract();

                        //  s = d.iimPlayCode(@"TAG POS=1 TYPE=DIV ATTR=ID:htmlApplication EXTRACT=HTM");
                        //  realist_rawHtml = d.iimGetLastExtract();

                        s = d.iimPlayCode(@"TAB CLOSE");

                        //          MessageBox.Show(realist_address + realist_rawHtml);



                        string allTables = realist_owner_table + "#NEXT#" + realist_location_table + "#NEXT#" + realist_taxinfo_table + "#NEXT#" + realist_char_table + "#NEXT#" + realist_estimatedvalue_table;



                        allTables.Replace("# #", "##");


                        //Township:#NEXT#Grafton Twp#NEXT#Census Tract:#NEXT#8711.04#NEXT##NEWLINE#Township Range Sect:#NEXT#43N-7E-27#NEXT#Carrier Route:#NEXT#R010#NEXT##NEWLINE#Subdivision:#NEXT#Pasquinellis Huntley Mdws#NEXT##NEXT# 

                        string[] sep = { "#NEWLINE#" };


                        string[] splitonnewline = allTables.Split(sep, StringSplitOptions.RemoveEmptyEntries);


                        string valuePattern = @"#NEXT#([^#]+)#NEXT#";
                        string namePattern = @"^([^#]+)#NEXT#|#NEXT#([^#]+)#NEXT#";
                        string fieldName;
                        string fieldValue;
                        string parid = Regex.Match(allTables, @"Parcel ID:#NEXT#([^#]+)").Groups[1].Value;

                        Dictionary<string, string> realistReportNameValuePairs = new Dictionary<string, string>();
                        #endregion

                        #region parse data and add/update fusion table
                        //if we found a parcel id, which is the key for the record
                        if (!String.IsNullOrEmpty(parid))
                        {

                            Stopwatch stopWatch = new Stopwatch();
                            stopWatch.Start();
                            string[] gFormatedLocation = new string[3];


                            gFormatedLocation = realist_address.Split(',');

                          
                           

                                if (!realist_bpohelper.RecordExists(parid))
                                {
                                    //if realist is missing addess, split will only have 2 fields, not 3
                                    try
                                    {
                                        realist_bpohelper.AddRecord(parid, gFormatedLocation[0] + " " + gFormatedLocation[1] + " " + gFormatedLocation[2]);
                                    }
                                    catch
                                    {
                                        realist_bpohelper.AddRecord(parid, address.Replace(",",""));
                                    }
                                }
                                else
                                {
                                    //add address to pending updates, incase it's missing or needs updating in fusion table
                                    realistReportNameValuePairs.Add("Location", gFormatedLocation[0] + " " + gFormatedLocation[1] + " " + gFormatedLocation[2]);

                                }
                      
                           

                            stopWatch.Stop();
                            // Get the elapsed time as a TimeSpan value.
                            TimeSpan sts = stopWatch.Elapsed;

                            // Format and display the TimeSpan value. 
                            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                                sts.Hours, sts.Minutes, sts.Seconds,
                                sts.Milliseconds / 10);

                            form.SetStatusBar = elapsedTime;



                            //
                            //Check and add new colums
                            //
                            foreach (string valuePairLine in splitonnewline)
                            {


                                string modStr = Regex.Replace(valuePairLine, @"# ", "Number").Replace("Garage #", "Garage Number");


                                MatchCollection mc = Regex.Matches(modStr, namePattern);
                                foreach (Match m in mc)
                                {


                                    string v = m.Value.Replace("#NEXT#", "").Trim().Replace("(", @"\(").Replace(")", @"\)") + "#NEXT#([^#]+)";
                                    if (!v.Contains("NODATA"))
                                    {
                                        if (realist_bpohelper.ColumnExsists(m.Value.Replace("#NEXT#", "").Trim()))
                                        {


                                            realistReportNameValuePairs.Add(m.Value.Replace("#NEXT#", "").Trim(), Regex.Match(modStr, v).Groups[1].Value);
                                            //realist_bpohelper.UpdateRecord(parid, m.Value.Replace("#NEXT#", "").Trim(), Regex.Match(modStr, v).Groups[1].Value);
                                            //MessageBox.Show(m.Value + "exsists.");
                                        }
                                        else
                                        {
                                            if (m.Value != "NEXT")
                                            {
                                                realist_bpohelper.AddColumn(m.Value.Replace("#NEXT#", "").Trim());
                                               
                                                try
                                                {
                                                realistReportNameValuePairs.Add(m.Value.Replace("#NEXT#", "").Trim(), Regex.Match(modStr, v).Groups[1].Value);
                                                }
                                                catch{
                                                    MessageBox.Show("problem adding realist nvp.");
                                                }
                                               

                                            }
                                        }
                                    }

                                }

                            }


                            //
                            //Update record
                            //

                            realistReportNameValuePairs.Concat(realist_bpohelper.Geocode(address));

                            realist_bpohelper.UpdateRecord(parid, realistReportNameValuePairs);

                            realist_bpohelper.m_rowid = "";


                            // MessageBox.Show(splitonnewline[0]);

                        }



                        //
                        //Location Information
                        //
                        pattern = "Subdivision:#NEXT#([^#]+)";
                        match = Regex.Match(realist_location_table, pattern);

                        realist_subdivision = match.Groups[1].Value;

                        //School District:
                        pattern = "School District:#NEXT#([^#]+)";
                        match = Regex.Match(realist_location_table, pattern);

                        realist_schoolDistrict = match.Groups[1].Value;

                        //Lot Acres:#NEXT#0.1257#NEXT#
                        pattern = "Lot Acres:#NEXT#([^#]+)";
                        match = Regex.Match(realist_char_table, pattern);
                         
                        string realist_lot_size = "0";
                        realist_lot_size = match.Groups[1].Value;
                        realist_lotAcres = realist_lot_size;

                        //MessageBox.Show("MLS Lot: " + mls_lot_size + " " + "Realist Lot Size:" + realist_lot_size);

                        if (mls_lot_size == "0" | mls_lot_size == "")
                        {
                            mls_lot_size = realist_lot_size;
                        }

                        pattern = "Building Above Grade Sq Ft:#NEXT#([^#]+)";
                        match = Regex.Match(realist_char_table, pattern);


                        realist_gla = match.Groups[1].Value;

                        //MessageBox.Show("MLS GLA: " + mls_gla + " " + "Realist GLA:" + realist_gla);
                        if (match.Success)
                        {
                            if (mls_gla == "0" | mls_gla == "")
                            {
                                mls_gla = realist_gla;
                            }
                        }
                        else
                        {
                            realist_gla = "-1";
                        }
                        pattern = "Year Built:#NEXT#([^#]+)";
                        match = Regex.Match(realist_char_table, pattern);

                        string realist_yearbuilt = "";
                        realist_yearbuilt = Regex.Match(match.Groups[1].Value, @"(\d\d\d\d)").Value;

                        //MessageBox.Show("MLS GLA: " + mls_gla + " " + "Realist GLA:" + realist_gla);

                        //MessageBox.Show("MLS YB: " + year_built + " " + "Realist YB:" + realist_yearbuilt);
                        if (string.IsNullOrEmpty(year_built) || year_built.ToLower() == "unk" || year_built == "0")
                        {
                            if (!string.IsNullOrEmpty(realist_yearbuilt))
                            {
                                year_built = realist_yearbuilt;
                            }
                            else
                            {
                                year_built = "1950";
                            }
                        }



                        DateTime subject_age = new DateTime((Convert.ToInt32(year_built)), 1, 1);

                        TimeSpan ts = DateTime.Now - subject_age;

                        int age = ts.Days / 365;


                        pattern = @"Census Tract:#NEXT#([^#]+)";
                        match = Regex.Match(realist_char_table, pattern);

                        if (match.Success)
                        {
                            censusTract = match.Groups[1].Value;
                        }
                        else
                        {
                            match = Regex.Match(realist_location_table, pattern);
                            censusTract = match.Groups[1].Value;
                        }


                    }
                        #endregion //cccc

                    #endregion  //ccccdsddd

                    else
                    {
                        //unusable record OR
                        //No realist pop-up ie Wisconson address

                        s = d.iimPlayCode("SET !TIMEOUT_STEP 0 \r\n TAB T=2 \r\n TAB CLOSE");
                        if (string.IsNullOrEmpty(year_built) || year_built.ToLower() == "unk" || year_built == "0")
                        {
                                year_built = "1950";   
                        }



                        DateTime subject_age = new DateTime((Convert.ToInt32(year_built)), 1, 1);

                        TimeSpan ts = DateTime.Now - subject_age;

                        int age = ts.Days / 365;

                    }




                }
                    #endregion

                else
                {
                    #region realist read from fusion table
                    try
                    {
                        realist_lotAcres = realist_bpohelper.curRec["Lot Acres:"];
                    }
                    catch
                    {

                    }
                    
                    if (mls_lot_size == "0" | mls_lot_size == "")
                    {
                        mls_lot_size = realist_lotAcres;
                        double x;
                        Double.TryParse(realist_lotAcres, out x);
                        currentListing.Lotsize = x;
                    }

                    realist_gla = realist_bpohelper.curRec["Building Above Grade Sq Ft:"];
                    if (String.IsNullOrEmpty(realist_gla))
                    {
                        realist_gla = "-1";
                    }
                    else
                    {
                        if (mls_gla == "0" | mls_gla == "")
                        {
                            mls_gla = realist_gla;
                            int x;
                            Int32.TryParse("mls_gla", out x);
                            currentListing.GLA = x;
                        }
                    }
                    realist_subdivision = realist_bpohelper.curRec["Subdivision:"];
                    censusTract = realist_bpohelper.curRec["Census Tract:"];
                    realist_schoolDistrict = realist_bpohelper.curRec["School District:"];
                    if (string.IsNullOrEmpty(year_built) || year_built.ToLower() == "unk" || year_built == "0")
                    {
                        if (realist_bpohelper.curRec["Year Built:"] != "")
                        {
                            if (realist_bpohelper.curRec["Year Built:"].Contains("Tax:"))
                            {
                                year_built = Regex.Match(realist_bpohelper.curRec["Year Built:"], @"Tax:\s*(\d\d\d\d)").Groups[1].Value;

                            }
                            else if (realist_bpohelper.curRec["Year Built:"].Contains("MLS:"))
                            {
                                year_built = Regex.Match(realist_bpohelper.curRec["Year Built:"], @"MLS:\s*(\d\d\d\d)").Groups[1].Value;
                            }
                            else
                            {
                                year_built = realist_bpohelper.curRec["Year Built:"];
                            }

                        }
                        else
                        {
                            year_built = "1950";
                        }
                    }
                    realist_bpohelper.Geocode(address);

                    #endregion
                }

                string propertyLatitude = "";
                string propertyLongitude = "";

                int xx = 0;
                Int32.TryParse(realist_gla.Replace(",",""), out xx);
                currentListing.RealistGLA = xx;

                double dd = 0;
                Double.TryParse(realist_lotAcres, out dd);
                currentListing.RealistLotSize = dd;

                try
                {

                     propertyLatitude = realist_bpohelper.curRec["Latitude"];
                     propertyLongitude = realist_bpohelper.curRec["Longitude"];

                    GeoPoint destPoint;
                    try
                    {
                        destPoint.Latitude = Convert.ToDouble(propertyLatitude);
                        destPoint.Longitude = Convert.ToDouble(propertyLongitude);
                    }
                    catch
                    {
                        destPoint = GlobalVar.subjectPoint;
                    }

                    currentListing.GeoPointGd = propertyLatitude + ", " + propertyLongitude;
                    currentListing.proximityToSubject = Math.Round(Get_Distance(GlobalVar.subjectPoint, destPoint), 2);

                }
                catch
                {

                }

                #endregion
                form.StatusUpdate = "Done.";

                RealistReport currentRealistReport = new RealistReport();

                currentRealistReport.subdivision = realist_subdivision;

                currentProperty.myRealistReport = currentRealistReport;

                rpl.Add(currentProperty);

                //
                //Collect stats
                //
                 #region reporting
                 domList.Add(Convert.ToInt16(dom));

                if (Convert.ToDecimal(realist_gla) > 0 && Convert.ToDecimal(mls_gla) > Convert.ToDecimal(realist_gla))
                {
                    mls_gla = realist_gla;
                }

                if (sold_price == " ")
                {
                    sold_price = current_list_price;
                }

                decimal[,] multiDimensionalArray1 = new decimal[2, 2];

                multiDimensionalArray1[0, 0] = Convert.ToDecimal(sold_price.Replace("$", "").Replace(",", ""));

                multiDimensionalArray1[0, 1] = Convert.ToDecimal(mls_gla.Replace(",", ""));

                ageList.Add(this.Age(year_built));
                int c = 0;

                if (realist_subdivision != "")
                {
                    try
                    {
                        c = subdivisionList[realist_subdivision];
                        subdivisionList[realist_subdivision] = c + 1;

                    }
                    catch
                    {
                        try
                        {
                            subdivisionList.Add(realist_subdivision, 1);
                        }
                        catch
                        {
                            MessageBox.Show("Failed adding to subdivision list: " + realist_subdivision);
                        }

                    }
                }
                
                 if (type != "")
                {
                    try
                    {
                        c = mlsTypeList[type];
                        mlsTypeList[type] = c + 1;

                    }
                    catch
                    {
                        try
                        {

                            mlsTypeList.Add(type, 1);
                        }
                        catch
                        {
                            MessageBox.Show("Failed adding to mls type list: " + mlsTypeList);
                        }
                    }
                }
                

               // subdivisionList.Add(mls_subdivision + ", " + realist_subdivision, 1);


                list.Add(mls_status + ", " + sold_price.Replace("$", "").Replace(",", "") + ", " + mls_gla.Replace(",", ""));

                switch (mls_status)
                {
                    case "CLSD":
                        closed_comps++;
                        soldPriceList.Add(Convert.ToDecimal(sold_price.Replace("$", "").Replace(",", "")));

                        if (mls_gla != "0")
                        {
                            pricePerSfList.Add(Convert.ToDecimal(sold_price.Replace("$", "").Replace(",", "")) / Convert.ToDecimal(mls_gla));
                            try
                            {



                                int x = form.RawSFDatatable.Insert(Convert.ToDouble(propertyLongitude), Convert.ToDouble(propertyLatitude), Convert.ToDecimal(sold_price.Replace("$", "").Replace(",", "")), Convert.ToInt16(mls_gla.Replace(",", "")), zip, city, censusTract, mlsnum);
                            }
                            catch (Exception ex)
                            {
                                try
                                {
                                    int x = form.RawSFDatatable.Update(Convert.ToDouble(propertyLongitude), Convert.ToDouble(propertyLatitude), Convert.ToDecimal(sold_price.Replace("$", "").Replace(",", "")), Convert.ToInt16(mls_gla.Replace(",", "")), zip, city, censusTract, mlsnum);
                                }
                                catch
                                {
                                }

                               // MessageBox.Show(ex.Message);
                            }
                        }
                        break;
                    case "CTG":
                    case "ACTV":
                        active_comps++;
                        activePriceList.Add(Convert.ToDecimal(current_list_price.Replace("$", "").Replace(",", "")));
                        break;
                    case "PEND":
                        pending_comps++;
                        activePriceList.Add(Convert.ToDecimal(current_list_price.Replace("$", "").Replace(",", "")));
                        break;

                    default:
                        //Console.WriteLine("Invalid selection. Please select 1, 2, or 3.");
                        break;
                }


                 #endregion

                //
                //compare to subject
                //

                #region compare

                CompCriteria cc = new CompCriteria();
                List<string> compNotes = new List<string>();
                int ageDiff = Math.Abs(Convert.ToInt16(form.SubjectYearBuilt) - Convert.ToInt16(year_built));


                int glaDiff = Math.Abs(Convert.ToInt16(form.SubjectAboveGLA.Replace(",", "")) - Convert.ToInt16(mls_gla.Replace(",", "")));

                bool typeMatch = false;
                bool glaMatch = false;
                bool ageMatch = false;
                bool lotMatch = false;
                bool basementMatch = false;
                bool mredBasement = false;
                bool subjectBasement = false;

                // if anything other than none = it has a basement of some type
                if (currentListing.mlsHtmlFields["basement"].value != "None")
                {
                    mredBasement = true;
                }
            
                //same with subject
                if (form.SubjectBasementType.ToLower() != "none")
                {
                    subjectBasement = true;
                }


                if (currentListing.mlsHtmlFields["basement"].value.Contains(form.SubjectBasementType))
                {
                    basementMatch = true;
                }
                else if (subjectBasement && mredBasement)
                {
                    //for now, any type of basement will match
                    basementMatch = true;
                }


                if (form.SubjectMlsType.ToLower().Contains("duplex") && type.ToLower().Contains("duplex"))
                {
                    typeMatch = true;
                }

                //for townhome searches
                if (form.SubjectMlsType.ToLower().Contains("townhouse") && type.ToLower().Contains("townhouse"))
                {
                    typeMatch = true;
                }

                //for condo searches (multi-type selected on listing, as long as on is condo, we have to consider it)
                if (form.SubjectMlsType.ToLower().Contains("condo") && type.ToLower().Contains("condo"))
                {
                    typeMatch = true;
                }

                //for 1.5 story 

                if (form.SubjectMlsType.ToLower().Contains("1.5 story") && type.ToLower().Contains("2 stories"))
                {
                    typeMatch = true;
                }

                if (form.SubjectMlsType.ToLower().Contains("townhouse") && type.ToLower().Contains("condo"))
                {
                    typeMatch = true;
                }

                //for splits

                if (form.SubjectMlsType.ToLower().Contains("split level") && type.ToLower().Contains("split level"))
                {
                    typeMatch = true;
                }

                if (type == form.SubjectMlsType)
                    typeMatch = true;

                if (GlobalVar.ccc == GlobalVar.CompCompareSystem.NABPOP)
                {
                    if (glaDiff < Convert.ToInt16(form.SubjectAboveGLA.Replace(",", "")) * .2)
                        glaMatch = true;
                } else if (GlobalVar.ccc == GlobalVar.CompCompareSystem.USER)
                {
                    if (GlobalVar.lowerGLA <= Convert.ToDecimal(mls_gla.Replace(",", "")) && Convert.ToDecimal(mls_gla.Replace(",", "")) <= GlobalVar.upperGLA)
                    {
                        glaMatch = true;
                    }
                }

               
                if (cc.Age(form.SubjectYearBuilt, year_built))
                    ageMatch = true;

                if (form.SubjectAttached)
                {
                    lotMatch = true;
                }
                else if (cc.LotSize(form.SubjectLotSize, mls_lot_size))
                {
                    lotMatch = true;
                }
               

                if (typeMatch)
                {
                    numberSameTypeAsSubject++;
                }

                if (glaMatch)
                {
                    numberComparableGla++;
                }

                if (ageMatch)
                {
                    numberComparableAge++;
                }

                if (lotMatch)
                {
                    numberComparableLot++;
                }



                //form.SetStatusBar = "SubjectLotSize--> " + form.SubjectLotSize + "; CompLotSize--> " + mls_lot_size + "; Comparable = " + lotMatch.ToString();
                //string str = string.Format("MLS: {0} is {1} from subject.", currentListing.mlsHtmlFields["mlsNumber"].value, currentListing.proximityToSubject);
                //form.SetStatusBar = str;

                //
                //Format and display comp info to user
                //
                Color txtColor = new Color();


                form.AddInfoColor(currentListing.mlsHtmlFields["mlsNumber"].value, Color.Black);

                form.AddInfo = "  ";

                if (currentListing.proximityToSubject <= 1) 
                {
                    form.AddInfoColor(currentListing.proximityToSubject.ToString(), Color.Green);
                } else
                {
                    form.AddInfoColor(currentListing.proximityToSubject.ToString(), Color.Yellow);
                }

                form.AddInfo = "\t";

                if (typeMatch)
                {
                    form.AddInfoColor(currentListing.mlsHtmlFields["mlsType"].value, Color.Green);
                }
                else
                {
                    form.AddInfoColor(currentListing.mlsHtmlFields["mlsType"].value, Color.Red);
                }


                //Type:2 Stories, Hillside = 19 chars, needs 2 tabs
                if (currentListing.mlsHtmlFields["mlsType"].value.Length >= 20)
                {
                    form.AddInfo = "\t";
                }
                else if (currentListing.mlsHtmlFields["mlsType"].value.Length >= 11)
                {
                    form.AddInfo = "\t\t";
                }
                else
                {
                    form.AddInfo = "\t\t\t";
                }

              

                if (glaMatch)
                {
                    //something we don't use the gla from mls listing or its blank
                    //form.AddInfoColor(currentListing.mlsHtmlFields["mlsGla"].value, Color.Green);
                    txtColor = Color.Green;
                }
                else
                {
                    txtColor = Color.Red;
                }

                form.AddInfoColor(mls_gla, txtColor);

                form.AddInfo = "\t";

                if (ageMatch)
                {
                    //form.AddInfoColor(currentListing.mlsHtmlFields["yearBulit"].value, Color.Green);
                    form.AddInfoColor(year_built, Color.Green);
                    
                }
                else
                {
                    //form.AddInfoColor(currentListing.mlsHtmlFields["yearBulit"].value, Color.Red);
                    form.AddInfoColor(year_built, Color.Red);
                }

                form.AddInfo = "\t";

                if (form.SubjectDetached)
                {
                    if (lotMatch)
                    {
                        //something we don't use the lot size from mls listing or its blank
                        //form.AddInfoColor(currentListing.mlsHtmlFields["mlsGla"].value, Color.Green);
                        txtColor = Color.Green;
                    }
                    else
                    {
                        txtColor = Color.Red;
                    }

                    form.AddInfoColor(mls_lot_size, txtColor);
                    //if (lotMatch)
                    //{
                    //    form.AddInfoColor(currentListing.mlsHtmlFields["acerage"].value, Color.Green);
                    //}
                    //else
                    //{
                    //    form.AddInfoColor(currentListing.mlsHtmlFields["acerage"].value, Color.Red);
                    //}
                }

                form.AddInfo = "\t";

                if (basementMatch)
                {
                    txtColor = Color.Green;
                }
                else
                {
                    txtColor = Color.Red;
                }
                form.AddInfoColor(currentListing.mlsHtmlFields["basement"].value, txtColor);

                //Type:2 Stories, Hillside = 19 chars, needs 2 tabs
                if (currentListing.mlsHtmlFields["basement"].value.Length >= 20)
                {
                    form.AddInfo = "\t";
                }
                else if (currentListing.mlsHtmlFields["basement"].value.Length >= 11)
                {
                    form.AddInfo = "\t\t";
                }
                else
                {
                    form.AddInfo = "\t\t\t";
                }
           


                if (typeMatch && glaMatch && ageMatch && lotMatch && basementMatch)
                {
                    form.AddInfo = "Yes";
                }
                else if (typeMatch && glaMatch || glaMatch && ageMatch)
                {
                    form.AddInfo = "Maybe";
                }
                else
                {
                    form.AddInfo = "No";
                }

                form.AddInfo = "\r\n";






               // string rtfDisplay = string.Format("{0} {1} {2} {3} {4} {5} {6}", currentListing.mlsHtmlFields["mlsNumber"].value, currentListing.proximityToSubject.ToString()., );


               //MessageBox.Show("Type: " + type + " ->" + typeMatch.ToString() + "GLA: " + mls_gla + " ->" + glaMatch.ToString() + " YearBuilt: " + year_built + " ->" + ageDiff.ToString());
               //MessageBox.Show("YearBuilt: " + year_built + " ->" + ageDiff.ToString() + " " + ageMatch.ToString());

                if (typeMatch && glaMatch && ageMatch && lotMatch && basementMatch)
                {

                   
                    
                        StringBuilder macro = new StringBuilder();
                        if (!searchCache)
                        {
                            macro.AppendLine(@"FRAME NAME=workspace");
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:dcid* CONTENT=YES");

                            status = d.iimPlayCode(macro.ToString(), 30);

                        }

                        comps.Add(mlsnum);
                        if (realist_subdivision == form.SubjectSubdivision)
                        {
                            compNotes.Add("Same subdivision as subject");
                        }
                        if (realist_schoolDistrict == form.SubjectSchoolDistrict)
                        {
                            compNotes.Add("Same school district as subject");
                        }

                        using (System.IO.StreamWriter file = new System.IO.StreamWriter(form.SubjectFilePath + "\\" + mlsnum + "notes.txt"))
                        {
                            foreach (object o in compNotes)
                            {
                                file.WriteLine(o.ToString());
                            }
                        }
                 //form = comps.Count.ToString();

                        form.NumberOfCompsFound = comps.Count.ToString();



                }

                #endregion

                //save order info to DB
                //ataTable BPOtable = this.form.bpoTableAdapter1.GetData();

               // string tempstore = currentListing.rawData;

               // currentListing.rawData = "";

                string json = JsonConvert.SerializeObject(currentListing);
                string url = "https://active-century-477.appspot.com/api/mred/v1/mredlistings/";

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "POST";

                System.Text.UTF8Encoding encoding = new System.Text.UTF8Encoding();
                Byte[] byteArray = encoding.GetBytes(json);

                request.ContentLength = byteArray.Length;
                request.ContentType = @"application/json";

                using (Stream dataStream = request.GetRequestStream())
                {
                    dataStream.Write(byteArray, 0, byteArray.Length);
                }
                long length = 0;
                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        length = response.ContentLength;
                       // form.AddInfo = response.StatusDescription;
                    }
                }
                catch (WebException ex)
                {
                    // Log exception and throw as for GET example above
                }



                //currentListing.rawData = tempstore;

                

                // gcds.StoreMredListing(currentListing);



              

                if (searchCache)
                {
                    count++;
                    if (count >= GlobalVar.searchCacheMlsListings.Count)
                    {
                        stillRecords = false;
                    }
                }

                else
                {
                    //GlobalVar.searchCacheMlsListings.Add(currentListing);
                    status = d.iimPlayCode(move_through_comps_macro.ToString(), 60);
                    if (status != Status.sOk)
                    {
                        stillRecords = false;
                    }
                }

            } //end while loop
 
            try
            {


                currentNeighborhood.highListPrice = activePriceList.Max();
                currentNeighborhood.maxSalePrice = soldPriceList.Max();
                currentNeighborhood.medianAge = ageList.Median();
                currentNeighborhood.medianSoldPrice = soldPriceList.Median();
                currentNeighborhood.minListPrice = activePriceList.Min();
                currentNeighborhood.minSalePrice = soldPriceList.Min();
                currentNeighborhood.newestHome = ageList.Min();
                currentNeighborhood.numberOfCompListings = comps.Count;
                currentNeighborhood.numberOfSales = soldPriceList.Count;
                currentNeighborhood.numberOfShortSaleListings = shortActive;
                currentNeighborhood.oldestHome = ageList.Max();
                currentNeighborhood.avgDom = Convert.ToInt32(domList.Average());
                currentNeighborhood.numberActiveListings = activePriceList.Count;
                currentNeighborhood.numberREOListings = reoActive;
                currentNeighborhood.numberREOSales = reoSales;
                currentNeighborhood.numberShortSales = shortSales;

                form.SubjectNeighborhood = currentNeighborhood;
            }
            catch (Exception e)
            {
                //tbd
            }
          

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(form.SubjectFilePath + "\\" + DateTime.Now.Ticks.ToString() + "_clsd_stats.txt"))
            {


            foreach (string s1 in list)
            {
                if (s1.Contains("CLSD"))
                {
                    file.WriteLine(s1);
                }
            }
            }

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(form.SubjectFilePath + "\\" + DateTime.Now.Ticks.ToString() + "_active_stats.txt"))
            {


                foreach (string s1 in list)
                {
                    if (!s1.Contains("CLSD"))
                    {
                        file.WriteLine(s1);
                    }
                }
            }


            string searchName = "";

            if (string.IsNullOrEmpty(form.CurrentSearchName))
            {
                searchName = form.SubjectFilePath + "\\" + DateTime.Now.Ticks.ToString() + "_summary.txt";
            } else
            {
                searchName = form.SubjectFilePath + "\\" + form.CurrentSearchName + "_" + DateTime.Now.Ticks.ToString() + "_summary.txt";
            }

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(searchName))
            {


                int numSameSubdivision = 0;
                if (subdivisionList.ContainsKey("form.SubjectSubdivision"))
                {
                    numSameSubdivision = subdivisionList[form.SubjectSubdivision];
                }
                try
                {
                    file.WriteLine("#closed: " + closed_comps.ToString() + " #active: " + active_comps.ToString() + " #pending: " + pending_comps.ToString());
                    file.WriteLine("Average/Mean $/Above GLA, sale: " + Decimal.Round(pricePerSfList.Average(), 2).ToString() + " Median $/Above GLA, sale: " + Decimal.Round(pricePerSfList.Median(), 2).ToString());
                    file.WriteLine("Min Sale: {0}, Max Sale: {1}, Average Sale: {2}, Median Sale: {3}", soldPriceList.Min(), soldPriceList.Max(), Convert.ToInt32(soldPriceList.Average()), soldPriceList.Median());
                    file.WriteLine("Min List: {0}, Max List: {1}, Average List: {2}, Median List: {3}", activePriceList.Min(), activePriceList.Max(), Convert.ToInt32(activePriceList.Average()), activePriceList.Median());
                    file.WriteLine("Min dom: {0}, Max dom: {1}, Average dom: {2}, Median dom: {3}", domList.Min(), domList.Max(), Convert.ToInt32(domList.Average()), domList.Median());
                    file.WriteLine("Min Age: {0}, Max Age: {1}, Average Age: {2}, Median Age: {3}", ageList.Min(), ageList.Max(), Convert.ToInt32(ageList.Average()), ageList.Median());
                    file.WriteLine("REO Sold: {0}, REO Active: {1}, Short Sold: {2}, Short Active: {3}", reoSales, reoActive, shortSales, shortActive);
                    file.WriteLine("# Same Type: {0}, # Same Subdivision {1}, # Comparable Age: {2},# Comparable aGLA: {3}, Comparable LotSize: {4}, ", numberSameTypeAsSubject.ToString(), numSameSubdivision.ToString(), numberComparableAge.ToString(), numberComparableGla.ToString(), numberComparableLot.ToString());
                }
                catch
                {

                }
                    //picDiffList
               // file.WriteLine("Min {0}, Max {1}, Average {2}, Median {3}", picDiffList.Min(), picDiffList.Max(), picDiffList.Average(), picDiffList.Median());
                foreach (MLSListing m in listings)
                {
                    file.WriteLine("MLS: {0} is {1} from subject.", m.mlsHtmlFields["mlsNumber"].value, m.proximityToSubject);
                }
                
                
                foreach (object o in mlsTypeList)
                {
                        file.WriteLine(o.ToString());
                }
                
                foreach (object o in subdivisionList)
                {
                        file.WriteLine(o.ToString());
                }



                foreach (object o in listingAgentList)
                {
                    file.WriteLine(o.ToString());
                }
                foreach (MLSListing m in listings)
                {
                    file.WriteLine(m.mlsHtmlFields["remarks"].value);
                }

               
            }



          

            var queryRealProperties = from prop in rpl
                                      where !string.IsNullOrWhiteSpace(prop.Subdivision)
                                      select prop;



            System.Xml.Serialization.XmlSerializer writer =
      new System.Xml.Serialization.XmlSerializer(typeof(RealProperty));

            System.IO.StreamWriter xmlTestfile = new System.IO.StreamWriter(form.SubjectFilePath +
                @"\SerializationOverview.xml");

            foreach (RealProperty rp in queryRealProperties)
            {

                writer.Serialize(xmlTestfile, rp);
            }

            xmlTestfile.Close();

            var queryListings = from cust in listings
                                       where cust.AttachedGarage()
                                       select cust;

            var queryListingsRemarks = from cust in listings
                                where cust.mlsHtmlFields["remarks"].value.ToLower().Contains("charming")
                                select cust;

            var queryListingsProximity = from l in listings
                                       where l.proximityToSubject <.25
                                       select l;

            MessageBox.Show(queryListingsRemarks.Count().ToString() + " listing used the word *charming*");

            
            System.Xml.Serialization.XmlSerializer writer2 =
            new System.Xml.Serialization.XmlSerializer(typeof(List<MLSListing>));


            string filename = @"\listingsFromSearchRunAt-" + DateTime.Now.ToString("dd-MM-yy-HH-mm-ss-ffff") + ".xml";
            System.IO.StreamWriter xmlTestfile2 = new System.IO.StreamWriter(form.SubjectFilePath + filename);
            System.IO.StreamWriter xmlTestfile3 = new System.IO.StreamWriter(form.SubjectFilePath +  @"\listingFromLastSearch.xml");

            writer2.Serialize(xmlTestfile2, listings);
            writer2.Serialize(xmlTestfile3, listings);
            //foreach (MLSListing l in listings)
            //{
            //    writer2.Serialize(xmlTestfile2, l);
            //}

            xmlTestfile2.Close();

            GlobalVar.listingsFromLastSearch.Clear();
            GlobalVar.listingsFromLastSearch = listings;

           // MessageBox.Show(queryRealProperties.Count().ToString());

            //MessageBox.Show(GlobalVar.searchCacheMlsListings.Count.ToString());
            
            MessageBox.Show("Comps found: " + comps.Count.ToString());


           
      





        }
    }

   

    public class TownshipReport : ICloneable
    {
         protected string rawData;
         protected  IYourForm form;
         protected Dictionary<string, string> dataExtractionMap;

         public TownshipReport(IYourForm form)
        {
            this.form = form;
        }

         public TownshipReport()
         {

         }
          public TownshipReport(IYourForm form, string s)
         {
             this.form = form;
             rawData = s;
             ProcessRawData();
         }




          protected virtual void ProcessRawData()
         {
             pin = SetParameter(dataExtractionMap["pin"]);
             yearBuilt = SetParameter(dataExtractionMap["yearBuilt"]);
             totalAssessedValue = Convert.ToDouble(SetParameter(dataExtractionMap["totalAssessedValue"]));
             aboveGradeGla = Convert.ToInt16(SetParameter(dataExtractionMap["aboveGradeGla"]));
             basementGla = Convert.ToInt16(SetParameter(dataExtractionMap["basementGla"]));
             finishedBasementGla = Convert.ToInt16(SetParameter(dataExtractionMap["finishedBasementGla"]));
         }

         #region ICloneable Members

         public object Clone()
         {
             return this.MemberwiseClone();
         }

         #endregion

         protected string SetParameter(string p)
         {
             Match match = Regex.Match(rawData, p);
             if (match.Success)
             {
                 return match.Groups[1].Value.Replace(",", "");
             }
             else
             {
                 return null;
             }
         }

        

         public string pin;
         public string yearBuilt;
         public double totalAssessedValue;
         public int aboveGradeGla;
         public int basementGla;
         public int finishedBasementGla;
    }

    public class AlgonquinTownshipReport : TownshipReport
    {
        public AlgonquinTownshipReport(IYourForm form, string s) : base(form, s)
        {
           
        }


      
        override protected void ProcessRawData()
        {
            //PARCEL NUMBERxxx19-13-103-032
            pin = SetParameter(@"PARCEL NUMBERx+([^\n]+)");

            //YEAR BUILTxxx1973 BUILDING STYLExxxSPLIT LEVEL 
            yearBuilt = SetParameter(@"YEAR BUILTx+(\d+)");

            //TOTAL ASSESSED 63,358
            try
            {
                totalAssessedValue = Convert.ToDouble(SetParameter(@"TOTAL ASSESSED ([\d,]+)"));
            }
            catch
            {
                totalAssessedValue = -1;
            }

            //Owner NamexxxTOTAL LIVING AREAxxx924 ENCLOSED PORCH 
            
            
            try
            {
                aboveGradeGla = Convert.ToInt16(SetParameter(@"TOTAL LIVING AREAxxx([\d,]+)"));
                if (aboveGradeGla == 0)
                {
                    try
                    {
                        //condo
                        aboveGradeGla = Convert.ToInt16(SetParameter(@"CONDO UNIT Sq. Ft.x*([\d,]+)"));
                    }
                    catch
                    {
                        aboveGradeGla = -1;
                    }
                }
            }
            catch
            {
                    aboveGradeGla = -1;
            }


            //TOTAL BASEMENT AREAxxx924 FINISHED BASEMENT 528 
            try
            {
                basementGla = Convert.ToInt16(SetParameter(@"TOTAL BASEMENT AREAxxx([\d,]+)"));
            }
            catch
            {
                basementGla = -1;
            }

            //TOTAL BASEMENT AREAxxx924 FINISHED BASEMENT 528 
            try
            {
                finishedBasementGla = Convert.ToInt16(SetParameter(@"FINISHED BASEMENTx*\s*([\d,]+)"));
            }
            catch
            {
                finishedBasementGla = -1;
            }





        }
    }

    public class LakeCountyTownshipReport : TownshipReport
    {
        public LakeCountyTownshipReport(IYourForm form, string s)
        {
            this.form = form;
            rawData = s;
            dataExtractionMap = new Dictionary<string, string>()
            {
                {"pin", @"Pin: ([\d-]+)"},
                {"yearBuilt", @"Year Built / Effective Age: (\d+)"},
                {"totalAssessedValue", @"Total Amount: \$([\d,]+)"},
                {"aboveGradeGla", @"Above Ground Living Areax*([\d,]+)"},
                {"basementGla", @"Basement Area \(Square Feet\):x*\s*([\d,]+)"},
                {"finishedBasementGla", @"Finished Basement:x*\s*([\d,]+)"}

            };
            ProcessRawData();
        }
    }
       public class McHenryTownshipReport : TownshipReport
        {
        public McHenryTownshipReport(IYourForm form, string s)
        {
            this.form = form;
            rawData = s;
            dataExtractionMap = new Dictionary<string, string>()  
            {
                {"pin", @"Property Index Numberx*([\d-]+)"},
                {"yearBuilt", @"Year Builtx*(\d+)"},
                {"totalAssessedValue", @"1xxx\d\d\d\dxxx\wxxx\$[\d,]+xxx\$[\d,]+xxx\$[\d,]+xxx\$([\d,]+)"},
                {"aboveGradeGla", @"Total Building Sq Ftx*(\d+)"},
                {"basementGla", @"Basement Sq Ftx*(\d+)"},
                {"finishedBasementGla", @"Lower Level Sq Ftx*(\d+)"}

            };
            ProcessRawData();
        }

    }

       public class ElginTownshipReport : TownshipReport
       {
           public ElginTownshipReport(IYourForm form, string s)
           {
               this.form = form;
               rawData = s;
               dataExtractionMap = new Dictionary<string, string>()  
            {
                {"pin", @"PIN #:\s*x*([\d-]+)"},
                {"yearBuilt", @"Year Built:x*(\d+)"},
                {"totalAssessedValue", @"\d\d\d\dxxxNormalxxx[\d,]+xxx[\d,]+xxx([\d,]+)"},
                {"aboveGradeGla", @"Total Building Sqft:x*(\d+)"},
                {"basementGla", @"- \w+:x*([\d,]+)"},
                {"finishedBasementGla", @"- \w+:x*([\d,]+)"}

            };
               ProcessRawData();
           }

       }
    
    public class USRESForm
    {

        public void prefill()
        {
            StringBuilder macro = new StringBuilder();
            
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abParcel CONTENT=02064100080000");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBkrDist CONTENT=21");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBkrLic CONTENT=471");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abAgExp CONTENT=7");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:d CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:s CONTENT=YES");
            macro.AppendLine(@"TAG POS=3 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:s CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:s CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:o CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=NAME:InputForm ATTR=TXT:(*)<SP>(must<SP>be<SP>a<SP>number)");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMcOwnerPerc CONTENT=90");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMcTenentPerc CONTENT=10");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMcCompUnits CONTENT=10");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMcCompLists CONTENT=4");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMcBoards CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abNbBound CONTENT=1<SP>mile<SP>radius");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMkLow CONTENT=120000");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMkHigh CONTENT=180000");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abPropSold CONTENT=6");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMkListLow CONTENT=129900");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMkListHigh CONTENT=199000");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abPropList CONTENT=10");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:o CONTENT=NO");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMkTime");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:u CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abMkTime CONTENT=180");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:n CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:y CONTENT=YES");
            macro.AppendLine(@"TAG POS=4 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:n CONTENT=YES");
            macro.AppendLine(@"TAG POS=5 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:n CONTENT=YES");
            macro.AppendLine(@"TAG POS=4 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:d CONTENT=YES");
            macro.AppendLine(@"TAG POS=4 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:o CONTENT=YES");
            macro.AppendLine(@"'TAG POS=0 TYPE=INPUT:RADIO FORM=NAME:InputForm ATTR=VALUE:u CONTENT=YES");
            string macroCode = macro.ToString();
            // status = iim.iimPlayCode(macroCode, timeout);
        }

        public void subjectfill()
        {
            StringBuilder macro = new StringBuilder();
        
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abSrce CONTENT=%MLS");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abLoc CONTENT=%Suburban");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abRight CONTENT=Fee<SP>Simple");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSite CONTENT=8,881");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSiteS1");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abUnits CONTENT=%1");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abView CONTENT=Neighborhood");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abDesign CONTENT=%Avg");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abAge CONTENT=2005");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abCond CONTENT=%Avg");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abRooms CONTENT=9");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abRoomsS1");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBeds CONTENT=3");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBaths CONTENT=2.1");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abGla CONTENT=2700");
            macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=NAME:InputForm ATTR=TXT::Above<SP>Grade<SP>Gross<SP>Living<SP>Area<SP>(SqFt)");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBsmt CONTENT=None<SP>/<SP>0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abHeat CONTENT=GasFA/Central");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abGarage CONTENT=%2CA");
            string macroCode = macro.ToString();
            // status = iim.iimPlayCode(macroCode, timeout);
        }

        public void compfill()
        {
            StringBuilder macro = new StringBuilder();
      
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abAddrS1 CONTENT=373<SP>Kennedy<SP>Dr<SP>");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abCityS1 CONTENT=Antioch");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abStateS1 CONTENT=%IL");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abZipS1 CONTENT=60002");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProxS1 CONTENT=%2Blk");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSpS1");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:InputForm ATTR=NAME:abCorpS1 CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSpS1 CONTENT=156000");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abStype1 CONTENT=%r");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abLps1 CONTENT=159900");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abLpReds1 CONTENT=2");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abSrceS1 CONTENT=%MLS");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSaleDtS1 CONTENT=7/22/11");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abDomS1 CONTENT=112");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abLocS1 CONTENT=%Suburban");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abRightS1 CONTENT=Fee<SP>Simple");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSiteS1 CONTENT=9000");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abUnitsS1 CONTENT=%1");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abViewS1 CONTENT=Neighborhood");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abDesignS1 CONTENT=%Avg");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abAgeS1 CONTENT=2005");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abCondS1 CONTENT=%Avg");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abRoomsS1 CONTENT=7");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBedsS1 CONTENT=4");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBathsS1 CONTENT=2.1");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abGlaS1 CONTENT=2620");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBsmtS1 CONTENT=Full<SP>/<SP>0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abHeat CONTENT=GasFA/Central");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abHeatS1 CONTENT=GasFA/Central");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abGarageS1 CONTENT=%2CA");
            string macroCode = macro.ToString();
            // status = iim.iimPlayCode(macroCode, timeout);

        }


    }

    public class PricingAI
    {
        


    }

    public class Broker
    {
        public string name;
        public string licensenumber;
        public string yearsexp;
        public string officezip;

        public Broker(string n, string ln, string ye, string oz)
        {
            name = n;
            licensenumber = ln;
            yearsexp = ye;
            officezip = oz;
        }

    }

    public class ExtBPO
    {
        public string subjectaddress;
        public string mlsurl;
        public string bpocompany;

        IE realist;
        IE  listingswindow;
        IE mainstreet;
        IE m2m;

        WatiN.Core.Form wf;
        WatiN.Core.Table maintable;

        string rawaddress;
        string address;
        string city;
        string zip;
        string gla;
        string ownername;
        string sellerconcessions = "0";
        string origlistprice;
        string other;
        string saleprice;
        string saledate;
        string proximity;
        string daysonmarket;
        string style;
        string yearbuilt;
        string condition;
        string totalrooms;
        string bedrooms;
        string baths;
        string parking;
        string parkingtype;
        string parkingaord;
        string lotsize;
        string comments;
        string comparison;
        string grosslivingarea = "";
        string exterior = "";
        string currentlistprice;
        string listdate;
        string mlsnumber = "";


        WatiN.Core.Element ElementFromFrames(string elementId, IList<INativeDocument> frames)
        {
            foreach (var f in frames)
            {
                var e = f.Body.AllDescendants.GetElementsById(elementId).FirstOrDefault();
                if (e != null) return new WatiN.Core.Element(listingswindow, e);

                if (f.Frames.Count > 0)
                {
                    var ret = ElementFromFrames(elementId, f.Frames);
                    if (ret != null) return ret;
                }
            }

            return null;
        } 

        public ExtBPO()
        {
            subjectaddress = "";
        }

        public string subjectstreetnumber()
        {
            string[] tempstrarry = subjectaddress.Split(' ');;
            return tempstrarry[0];
        }

        public string subjectstreetname()
        {
            string[] tempstrarry = subjectaddress.Split(' '); ;
            return tempstrarry[1];
        }

        private void MoveNextRecord()
        {
            //later
        }
   


        private void Attach()
        {
            Regex url;

            if (bpocompany == "m2m")
            {
                //m2m
                 url = new Regex("https://m2m.aspengrove.net/Order/OrderEditWizard");
            }
            else
            {
                //emort
                 url = new Regex("https://www.emortgagelogic.com/main.html");
            }
           
            
            //ie.Frame(Find.ByName("main")).Frame(Find.ByName("workspace")).Form(Find.ByName("dc")).Link(Find.ByUrl(addressurl)).Text;

            //listingswindow = new IE();
            listingswindow = IE.AttachTo<IE>(Find.ByUrl(mlsurl));
            //listingswindow = IE.AttachTo<IE>(Find.ByUrl(new Regex("mls.jsp.module.search")));

            var xxx = listingswindow.Title;

            m2m = IE.AttachTo<IE>(Find.ByUrl(url));
            //realist = IE.AttachTo<IE>(Find.ByUrl("http://realist2.firstamres.com/propertydetail.jsp"));


            var ssds = listingswindow.hWnd;

       //     IWebDriver driver;

         

            //string mslnum = m2m.TextField(Find.BySelector("//div[@id='listingspane']/table/tbody/tr/td/div/table/tbody/tr/td[2]/table/tbody/tr/td[3]")).Text;

           // Element sand.E

            

            //foreach (WatiN.Core.Link l in listingswindow.Frame(Find.ByName("main")).Frame(Find.ByName("navpanel")).Links)
            //{
            //    l.Flash();
            //}
            var abc = listingswindow.ToString();
            var myframes =  listingswindow.Frames;
            var myforms = listingswindow.Forms;

           
            

           // Element test = this.ElementFromFrames("main", listingswindow.NativeDocument.Frames);

         //   var d = test.ClassName;

            Frame targetframe = listingswindow.Frame(Find.ByName("main"));

            

            myframes = targetframe.Frames;

            Frame targetframe2 = targetframe.Frame(Find.ByName("workspace"));

            //
            
            //var targetframe1 = listingswindow.Frame(Find.ByName("workspace"));
           wf = listingswindow.Frame(Find.ByName("main")).Frame(Find.ByName("workspace")).Form(Find.ByName("dc"));



            //string temp = listingswindow.Div("listingpane").Table(Find.ByIndex(1))



           // index 8 is multipicutre, 7 single or no 
            // table count 31 for attached 1 pic listing,
            // index 7 for 1 pic 
            if (wf.TableCell(Find.ByText("Attached Single")).Exists)
            {
                if (wf.Tables.Count == 31)
                {
                    maintable = wf.Table(Find.ByIndex(7));
                }
                else
                {
                    maintable = wf.Table(Find.ByIndex(8));
                }
                //set attached mls listings flag

                

            }
            else if (wf.TableCell(Find.ByText("Detached Single")).Exists)
            {
                if (wf.Tables.Count == 22)
                {
                    maintable = wf.Table(Find.ByIndex(8));
                }
                else
                {
                    maintable = wf.Table(Find.ByIndex(9));
                }
            }
            else
            {
                if (wf.Tables.Count == 29)
                {
                    maintable = wf.Table(Find.ByIndex(7));
                }
                else
                {
                    maintable = wf.Table(Find.ByIndex(8));
                }
            }

           
            
           wf.Table(Find.ByIndex(6)).Link(Find.ByIndex(0)).Click();

           wf.Table(Find.ByIndex(7)).Flash();
           wf.Table(Find.ByIndex(8)).Flash();
           wf.Table(Find.ByIndex(9)).Flash();
           wf.Table(Find.ByIndex(10)).Flash();
           wf.Table(Find.ByIndex(11)).Flash();
           Regex addressurl = new Regex("photoBrowser");

           IE photopage = IE.AttachTo<IE>(Find.ByUrl(addressurl));




           WatiN.Core.Image myimage = photopage.Image(Find.ByIndex(1));



           //myimage.Click();

//
  //         IE temp = new IE(photopage.Image(Find.ByIndex(1)).Src);


  //         photopage.CaptureWebPageToFile("test.bmp");

   //         photopage.Image("re").Cl


        }

        private void GetMLSData()
        {
            // get data in row order of mls listing page

            //target cell
            WatiN.Core.TableCell tcell;

            


            //mlsnumber = maintable.TableRow(Find.ByIndex(0)).TableCell(Find.ByIndex(2)).Text;
            mlsnumber = maintable.TableCell(Find.ByText(new Regex("MLS\\s#:"))).NextSibling.Text;
            //currentlistprice = maintable.TableRow(Find.ByIndex(0)).TableCell(Find.ByIndex(4)).Text;
            currentlistprice = maintable.TableCell(Find.ByText(new Regex("List\\sPrice"))).NextSibling.Text;
            //listdate = maintable.TableRow(Find.ByIndex(1)).TableCell(Find.ByIndex(3)).Text;
            listdate = maintable.TableCell(Find.ByText(new Regex("List\\sDate"))).NextSibling.Text;
            //origlistprice = maintable.TableRow(Find.ByIndex(1)).TableCell(Find.ByIndex(5)).Text;
            origlistprice = maintable.TableCell(Find.ByText(new Regex("Orig\\sList\\sPrice"))).NextSibling.Text;
            //saleprice = maintable.TableRow(Find.ByIndex(2)).TableCell(Find.ByIndex(5)).Text;
            saleprice = maintable.TableCell(Find.ByText(new Regex("Sold\\sPrice"))).NextSibling.Text;
            //old way
            //rawaddress = maintable.TableRow(Find.ByIndex(3)).TableCell(Find.ByIndex(1)).Text;
            //new way
            rawaddress = maintable.TableCell(Find.ByText("Address:")).NextSibling.Text;




            //maintable.TableRow(Find.ByIndex(3)).TableCell(Find.ByIndex(1)).Highlight(true);

            //get rid of (F) and (S)
            if (saleprice.IndexOf('(') > 0)
            {
                saleprice = saleprice.Remove(saleprice.IndexOf('('));

            }



            string[] splitaddress = rawaddress.Split(',');
            //string[] splitcityzip = splitaddress[1].Split(' ');

            address = splitaddress[0];
            if (splitaddress[1].Contains('-'))
            {
                splitaddress[1] = splitaddress[1].Remove(splitaddress[1].LastIndexOf('-'));
            }

            //oringinal mls
            //city = splitaddress[1].Substring(0, splitaddress[1].Length - 6);
            //current mls 9/25/2011
            city = splitaddress[1];

            
            zip = splitaddress[2].Substring(splitaddress[2].Length - 5, 5);


            //daysonmarket = maintable.TableRow(Find.ByIndex(5)).TableCell(Find.ByIndex(3)).Text;

            daysonmarket = maintable.TableCell(Find.ByText(new Regex("Lst.\\sMkt.\\sTime"))).NextSibling.Text;

            //saledate = maintable.TableRow(Find.ByIndex(6)).TableCell(Find.ByIndex(3)).Text;

            //attached
            //saledate = maintable.TableCell(Find.ByText(new Regex("Contract\\sDate"))).NextSibling.Text;
            //detached
            saledate = maintable.TableCell(Find.ByText(new Regex("Contract"))).NextSibling.Text;

            //sellerconcessions = maintable.TableRow(Find.ByIndex(6)).TableCell(Find.ByIndex(5)).Text;

            sellerconcessions = maintable.TableCell(Find.ByText("Points:")).NextSibling.Text;

            if (sellerconcessions == null)
            {
                sellerconcessions = "0";
            }




            //yearbuilt = maintable.TableRow(Find.ByIndex(8)).TableCell(Find.ByIndex(1)).Text;

            yearbuilt = maintable.TableCell(Find.ByText(new Regex("Year\\sBuilt"))).NextSibling.Text;

            Regex year = new Regex("\\d\\d\\d\\d");

            if (!year.IsMatch(yearbuilt))
            {
                yearbuilt = "";

            }

            //Regex dim = new Regex("\\d+X\\d+\\n");

            //   wf.Table(Find.ByIndex(10)).Table(Find.ByText(new Regex("Acreage"))).TableCell(Find.ByText(new Regex("Acreage"))).Flash();
            //wf.Span(Find.ByText("Miscellaneous")).NextSibling.Flash();
            //wf.Span(Find.ByText("Miscellaneous")).NextSibling.TableCell(Find.ByText(new Regex("Acreage"))).NextSibling
            //WatiN.Core.Element e = wf.Span(Find.ByText("Miscellaneous")).NextSibling;
            //WatiN.Core.Table t = m2m.Elements.


            // t.TableCell(Find.ByText(new Regex("Acreage"))).NextSibling.Flash();
            //  lotsize = t.TableCell(Find.ByText(new Regex("Acreage"))).NextSibling.Text;
            if (wf.Table(Find.ByIndex(10)).Table(Find.ByText(new Regex("Acreage"))).TableCell(Find.ByText(new Regex("Acreage"))).Exists)
                lotsize = wf.Table(Find.ByIndex(10)).Table(Find.ByText(new Regex("Acreage"))).TableCell(Find.ByText(new Regex("Acreage"))).NextSibling.Text;
            else
                lotsize = "0";



            //lotsize = maintable.TableRow(Find.ByIndex(10)).TableCell(Find.ByIndex(3)).Text;

            //if (!dim.IsMatch(lotsize))
            //{
            //    lotsize = "";

            //}
            //else
            //{
            //    string[] dims = lotsize.Split('X');
            //    int length = Convert.ToInt32(dims[0]);
            //    int width = Convert.ToInt32(dims[1]);

            //    float temp2 = length * width;

            //    float temp = temp2 / 43560;

            //    lotsize = temp.ToString("F");


            //}

            //totalrooms = maintable.TableRow(Find.ByIndex(13)).TableCell(Find.ByIndex(1)).Text;
            
            if (maintable.TableCell(Find.ByText("Rooms:")).Exists)
            {
                //dettached and attached
                totalrooms = maintable.TableCell(Find.ByText("Rooms:")).NextSibling.Text;
            }
            else
            {
                //2-4 units
                totalrooms = maintable.TableCell(Find.ByText("Total Rooms:")).NextSibling.Text;
            }

            if (maintable.TableCell(Find.ByText("Bedrooms:")).Exists)
            {
                //dettached and attached
                bedrooms = maintable.TableCell(Find.ByText("Bedrooms:")).NextSibling.Text;
                bedrooms = bedrooms.Substring(0, 1);
            }
            else
            {
                //2-4 units
                bedrooms = maintable.TableCell(Find.ByText("Total Bedrooms:")).NextSibling.Text;
                bedrooms = bedrooms.Substring(0, 1);
            }


            //bedrooms = maintable.TableRow(Find.ByIndex(13)).TableCell(Find.ByIndex(3)).Text;
            //this finds the map link
            // Regex addressurl = new Regex("dynaconn.mapping.MapClient");
            //ie.Frame(Find.ByName("main")).Frame(Find.ByName("workspace")).Form(Find.ByName("dc")).Link(Find.ByUrl(addressurl)).Text;



            //baths = maintable.TableRow(Find.ByIndex(13)).TableCell(Find.ByIndex(5)).Text;
            if (maintable.TableCell(Find.ByText("Bathrooms (full/half):")).Exists)
            {
                //dettached and attached
                baths = maintable.TableCell(Find.ByText("Bathrooms (full/half):")).NextSibling.Text;
            }
            else
            {
                //2-4 units
                baths = maintable.TableCell(Find.ByText("Bathrooms (full/half):")).NextSibling.Text;
            }

            baths = baths.Replace(" / ", ".");


            // string parkingtype = maintable.TableRow(Find.ByIndex(15)).TableCell(Find.ByIndex(1)).Text;
            parkingtype = maintable.TableCell(Find.ByText("Parking:")).NextSibling.Text;
            parkingaord = wf.TableCell(Find.ByText("# Spaces:")).NextSibling.Text;
            string[] splitparking = parkingaord.Split(':');


            parkingaord = splitparking[0];

            parking = splitparking[1];
            parking = parking.TrimEnd(' ');

            //parking = maintable.TableRow(Find.ByIndex(15)).TableCell(Find.ByIndex(3)).Text;
            //parking = maintable.TableCell(Find.ByText("Cars:")).NextSibling.Text;


            //comments = "From Listing: " + wf.Table(Find.ByIndex(24)).TableRow(Find.ByIndex(0)).Text;


            comments = "From Listing:: " + wf.Span(Find.ByText("Remarks:")).NextSibling.Text;




            //from realist property detail report

            wf.Link(Find.ByText("Realist Tax")).Click();
            realist = IE.AttachTo<IE>(Find.ByUrl("http://realist2.firstamres.com/propertydetail.jsp"));

            ownername = this.CompleteStepOne(realist.Table("Owner Info").TableRow(Find.ByTextInColumn("Owner Name:", 1)).TableCell(Find.ByIndex(realist.Table("Owner Info").TableCell(Find.ByText("Owner Name:")).Index + 2)).Text);

            if (realist.Table("Characteristics").Exists)
            {

                WatiN.Core.TableCellCollection tcc = realist.Table("Characteristics").TableCells;

              //  foreach (WatiN.Core.TableCell tc in tcc)
              //      MessageBox.Show(tc.Text);


                if (tcc.First(Find.ByText("Year Built:")) != null & yearbuilt == "")
                {
                    yearbuilt = tcc.First(Find.ByText("Year Built:")).NextSibling.NextSibling.Text;
                }

                if (tcc.First(Find.ByText("Lot Acres:")) != null & (lotsize == "0" || lotsize == ""))
                {
                    lotsize = tcc.First(Find.ByText("Lot Acres:")).NextSibling.NextSibling.Text;
                }



                try
                {
                    if (realist.Table("Characteristics").TableRow(Find.ByTextInColumn("Building Sq Ft:", 1)).Exists)
                    {
                        grosslivingarea = realist.Table("Characteristics").TableRow(Find.ByTextInColumn("Building Sq Ft:", 1)).TableCell(Find.ByIndex(realist.Table("Characteristics").TableCell(Find.ByText("Building Sq Ft:")).Index + 2)).Text;
                    }
                    else if (realist.Table("Characteristics").TableRow(Find.ByTextInColumn("Building Sq Ft:", 4)).Exists)
                    {
                        grosslivingarea = realist.Table("Characteristics").TableRow(Find.ByTextInColumn("Building Sq Ft:", 4)).TableCell(Find.ByIndex(realist.Table("Characteristics").TableCell(Find.ByText("Building Sq Ft:")).Index + 2)).Text;
                    }
                    else if (realist.Table("Characteristics").TableRow(Find.ByTextInColumn("Above Gnd Sq Ft:", 1)).Exists)
                    {
                        grosslivingarea = realist.Table("Characteristics").TableRow(Find.ByTextInColumn("Above Gnd Sq Ft:", 1)).TableCell(Find.ByIndex(realist.Table("Characteristics").TableCell(Find.ByText("Finished SF (Building):")).Index + 2)).Text;
                    }
                    else
                    {
                        grosslivingarea = "0";
                    }
                    //if there is a number in realist and one in the mls listing for sf, both will appear in this field. 
                    //Realist number will be first.
                    if (grosslivingarea.Contains('|'))
                    {
                        string[] glasplit = grosslivingarea.Split('|');
                        grosslivingarea = glasplit[0];
                        
                    }

                }
                catch
                {
                    grosslivingarea = "0";
                }

                try
                {
                    if (realist.Table("Characteristics").TableRow(Find.ByTextInColumn("Exterior:", 1)).Exists)
                    {
                        exterior = realist.Table("Characteristics").TableRow(Find.ByTextInColumn("Exterior:", 1)).TableCell(Find.ByIndex(realist.Table("Characteristics").TableCell(Find.ByText("Exterior:")).Index + 2)).Text;
                    }
                    else if (realist.Table("Characteristics").TableRow(Find.ByTextInColumn("Exterior:", 4)).Exists)
                    {
                        exterior = realist.Table("Characteristics").TableRow(Find.ByTextInColumn("Exterior:", 4)).TableCell(Find.ByIndex(realist.Table("Characteristics").TableCell(Find.ByText("Exterior:")).Index + 2)).Text;
                    }
                    else
                    {
                        exterior = "Metal/Vinyl";
                    }

                }
                catch
                {
                    exterior = "Metal/Vinyl";
                }

                Regex brick = new Regex("Brick");




                if (exterior == "Concrete")
                {
                    exterior = "Hardboard/Masonite";

                }



                if (brick.IsMatch(exterior))
                {
                    exterior = "Brick";
                }

                if (exterior.Contains("Aluminum"))
                {
                    exterior = "Metal/Vinyl";

                }

                if (exterior.Contains("Frame"))
                {
                    exterior = "Wood";

                }

                if (exterior.Contains("Stucco"))
                {
                    exterior = "Stucco";

                }

               

            }
            realist.Close();

            if (listingswindow.Frame(Find.ByName("main")).Frame(Find.ByName("navpanel")).Links.Count > 1)
            {
                listingswindow.Frame(Find.ByName("main")).Frame(Find.ByName("navpanel")).Link(Find.ByIndex(1)).Click();
            }
            else
            {
                listingswindow.Frame(Find.ByName("main")).Frame(Find.ByName("navpanel")).Link(Find.ByIndex(0)).Click();
            }



          






        }

        public void Enter_Data_emort(string property, string prefix)
        {

           

            //based on form G
            WatiN.Core.Frame mainframe = m2m.Frame(Find.ByName("main"));
            WatiN.Core.Form bpoform = mainframe.Form(Find.ByName("BPOForm"));
            WatiN.Core.TextFieldCollection tfc = bpoform.TextFields;
            WatiN.Core.TextField tf;




            //street address
            tf = tfc.First(Find.ByName(prefix + "Address" + property));
            tf.TypeText(address);

            //city
            tfc.First(Find.ByName(prefix + "City" + property)).TypeText(city);

            //state
            //  bpoform.Link(Find.By("href", "https://www.emortgagelogic.com/broker/dialog/forms/B/FormBg.html?OrderID=1371480&Form=g&OrderType=B&open_again=bpo#")).Click();

            //         tfc.First(Find.ByName("s" + prefix + "State" + property)).Click();


            WatiN.Core.Div dv = mainframe.Div(Find.ById("dsBox"));
            
            //form g, not on form i
            

            if (tfc.First(Find.ByName("s" + prefix + "State" + property)) != null)
            {
                tfc.First(Find.ByName("s" + prefix + "State" + property)).Click();
                dv.Link(Find.By("sval", "IL")).Click();
            }
          


            //hit prox button
            bpoform.Link(Find.ByClass("getprox")).Click();

            //reo y/n
            //
            //impliment later
            //

            //Sales price
            if (prefix == "CSales" && saleprice != "")
            {
                tfc.First(Find.ByName(prefix + "SPrice" + property)).TypeText(saleprice.TrimStart('$'));

                //sales date
                tfc.First(Find.ByName(prefix + "SDate" + property)).TypeText(saledate);
            }

            //original list price
            tfc.First(Find.ByName(prefix + "OLPrice" + property)).TypeText(origlistprice.TrimStart('$'));

            //current list, CListCLPrice_1
            if (prefix == "CList")
            {
                tfc.First(Find.ByName(prefix + "CLPrice" + property)).TypeText(currentlistprice.TrimStart('$'));
            }
            

            //HOA fee
            //to bo improved later
            //default sfr for now
            //form g, not on i
            //form g, not on form i


            if (tfc.First(Find.ByName(prefix + "HOA" + property)) != null)
            {
                tfc.First(Find.ByName(prefix + "HOA" + property)).TypeText("0/mo");
            }
            

            //type sCSalesType_1
            // same as subject
            
            tfc.First(Find.ByName("s" + prefix + "Type" + property)).Click();
            dv = mainframe.Div(Find.ById("dsBox"));
            //this works
            //dv.Link(Find.By("sval", "S")).Click();
            dv.Link(Find.ByTitle(tfc.First(Find.ByName("sPropType")).Text)).Click();

            //zip, CSalesZip_1
            if (tfc.First(Find.ByName(prefix + "Zip" + property)) != null)
            {
                tfc.First(Find.ByName(prefix + "Zip" + property)).TypeText(zip);
            }


            //style, sCListStyle_2
            //default same as subject, sPropStyle

            tfc.First(Find.ByName("s" + prefix + "Style" + property)).Click();
            dv = mainframe.Div(Find.ById("dsBox"));
            dv.Link(Find.ByTitle(tfc.First(Find.ByName("sPropStyle")).Text)).Click();

            

            //unit, sCSalesUnits_1
            if (tfc.First(Find.ByName("s" + prefix + "Units" + property)) != null)
            {
                tfc.First(Find.ByName("s" + prefix + "Units" + property)).Click();
                mainframe.Div(Find.ById("dsBox")).Link(Find.ByIndex(0)).Click();
            }
            //sale type
            //
            //imp later
            //

            //pool sCSalesPool_1
            //default no
            if (tfc.First(Find.ByName("s" + prefix + "Pool" + property)) != null)
            {
                tfc.First(Find.ByName("s" + prefix + "Pool" + property)).Click();
                mainframe.Div(Find.ById("dsBox")).Link(Find.ByIndex(1)).Click();
            }
            //spa, sCSalesSpa_1
            //default no
            if (tfc.First(Find.ByName("s" + prefix + "Spa" + property)) != null)
            {
                tfc.First(Find.ByName("s" + prefix + "Spa" + property)).Click();
                mainframe.Div(Find.ById("dsBox")).Link(Find.ByIndex(1)).Click();
            }
            //beds
            tfc.First(Find.ByName("s" + prefix + "BR" + property)).Click();
            mainframe.Div(Find.ById("dsBox")).Link(Find.ByIndex(Convert.ToInt16(bedrooms) -1)).Click();

            //baths, sCSalesBA_1
            if (baths.Length > 1)
            {
                baths = baths.Replace(".1", ".5");
                baths = baths.Replace(".2", ".5");
                baths = baths.Replace(".3", ".5");

            }
            else
            {
                baths = baths + ".0";
            }


            tfc.First(Find.ByName("s" + prefix + "BA" + property)).Click();
            dv.Link(Find.By("sval", baths)).Click();

            //string[] temparr = baths.Split('.');
            //if (temparr.Length < 2)
            //{
            //    tfc.First(Find.ByName("s" + prefix + "BA" + property)).Click();
            //    mainframe.Div(Find.ById("dsBox")).Link(Find.ByIndex(Convert.ToInt16(temparr[0]) - 1)).Click();
             
            //}
            //else
            //{
            //    tfc.First(Find.ByName("s" + prefix + "BA" + property)).Click();
            //    mainframe.Div(Find.ById("dsBox")).Link(Find.ByIndex(Convert.ToInt16(temparr[0]))).Click();
            //}

            //rooms, sCSalesTR_1
            tfc.First(Find.ByName("s" + prefix + "TR" + property)).Click();
            mainframe.Div(Find.ById("dsBox")).Link(Find.ByIndex(Convert.ToInt16(totalrooms)-1)).Click();


            //Sq. Ft., CSalesSqFt_1
            tfc.First(Find.ByName(prefix + "SqFt" + property)).TypeText(grosslivingarea.Replace(",", ""));

            //year built, CSalesYrBlt_1
            tfc.First(Find.ByName(prefix + "YrBlt" + property)).TypeText(yearbuilt);

            //garage, sCSalesGar_1
            //

            


            string tmp = parking + "\\s+.*" + parkingaord;
            Regex ptype = new Regex(tmp, RegexOptions.IgnoreCase);
            tfc.First(Find.ByName("s" + prefix + "Gar" + property)).Click();
            dv = mainframe.Div(Find.ById("dsBox"));
            


            if (parkingtype == "Garage")
            {
                
                switch (parking)
                {
                    case "1":
                        dv.Link(Find.ByTitle(ptype)).Click();
                        break;

                    case "2":
                        dv.Link(Find.ByTitle(ptype)).Click();
                        break;
                    case "3":
                        dv.Link(Find.ByTitle("3 Detached")).Click();
                        break;
                }
            }

            //fin incentives, CSalesFin_1
            if (tfc.First(Find.ByName(prefix + "Fin" + property)) != null)
            {
                tfc.First(Find.ByName(prefix + "Fin" + property)).TypeText(sellerconcessions);
            }

            //lot size, CSalesLotSize_1
            tfc.First(Find.ByName(prefix + "LotSize" + property)).TypeText(lotsize + "ac");

            //view, CSalesView_1
            //default to Neighborhood
            if (tfc.First(Find.ByName(prefix + "View" + property)) != null)
            {
                tfc.First(Find.ByName(prefix + "View" + property)).TypeText("Neighborhood");
            }
            //Basement SF, CSalesBsmtSqFt_1
            //
            //tbi
            //

            //% finished, CSalesBsmtPerFin_1
            //
            //tbi
            //

            //design/appeal, CSalesAppeal_1
            // default good
            if (tfc.First(Find.ByName(prefix + "Appeal" + property)) != null)
            {
                tfc.First(Find.ByName(prefix + "Appeal" + property)).TypeText("Good");
            }
            //Overall Condition, sCSalesCond_1
            tfc.First(Find.ByName("s" + prefix + "Cond" + property)).Click();
            mainframe.Div(Find.ById("dsBox")).Link(Find.ByIndex(1)).Click();

            //Source, sCSalesSource_1
            tfc.First(Find.ByName("s" + prefix + "Source" + property)).Click();
            mainframe.Div(Find.ById("dsBox")).Link(Find.ByIndex(0)).Click();

            //Source ID, CSalesSourceID_1
            tfc.First(Find.ByName(prefix + "SourceID" + property)).TypeText(mlsnumber);

            //other (comments), CSalesOther_1
            if (tfc.First(Find.ByName(prefix + "Other" + property)) != null)
            {
                tfc.First(Find.ByName(prefix + "Other" + property)).TypeText("na");
            }

            //dom, CSalesDOM_1
            if (tfc.First(Find.ByName(prefix + "DOM" + property)) != null)
            {
                tfc.First(Find.ByName(prefix + "DOM" + property)).TypeText(daysonmarket);
            }

            //location, sCSalesLoc_1

             if (tfc.First(Find.ByName("s" + prefix + "Loc" + property)) != null)
             {
                 tfc.First(Find.ByName("s" + prefix + "Loc" + property)).Click();
                 mainframe.Div(Find.ById("dsBox")).Link(Find.ByIndex(1)).Click();
             }

            //pool,patio,..., CSalesExt_1
             if (tfc.First(Find.ByName(prefix + "Ext" + property)) != null)
             {
                 tfc.First(Find.ByName(prefix + "Ext" + property)).TypeText("na");
             }

            //list price at at sale, CSalesSLPrice_1
             if (tfc.First(Find.ByName(prefix + "SLPrice" + property)) != null)
             {
                 tfc.First(Find.ByName(prefix + "SLPrice" + property)).TypeText(currentlistprice);
             }

            //adjustment boxes
            //

             //sCSalesStyleAdj_1, sCSalesLocAdj_1, sCSalesYrBltAdj_1,sCSalesLotSizeAdj_1, sCSalesGarAdj_1, sCSalesExtAdj_1,sCSalesLandAdj_1
             //sCSalesFinAdj_1, sCSalesCondAdj_1, sCSalesEstValAdj_1

            string [] strarr = {"StyleAdj", "LocAdj", "YrBltAdj", "LotSizeAdj", "GarAdj", "ExtAdj", "LandAdj", "FinAdj", "CondAdj", "EstValAdj" };
            foreach (string s in strarr)
            {
             if (tfc.First(Find.ByName("s" + prefix + s + property)) != null)
             {
                 tfc.First(Find.ByName("s" + prefix + s + property)).Click();
                 mainframe.Div(Find.ById("dsBox")).Link(Find.ByIndex(1)).Click();
             }
            }

            //Overall adj, CSalesEstVal_1
            if (tfc.First(Find.ByName(prefix + "EstVal" + property)) != null)
            {
                tfc.First(Find.ByName(prefix + "EstVal" + property)).TypeText("0");
            }

            //Landscaping, sCSalesLand_2
            if (tfc.First(Find.ByName("s" + prefix + "Land" + property)) != null)
            {
                tfc.First(Find.ByName("s" + prefix + "Land" + property)).Click();
                mainframe.Div(Find.ById("dsBox")).Link(Find.ByIndex(1)).Click();
            }

            //Origlist, CListOLDate_2
            if (tfc.First(Find.ByName(prefix + "OLDate" + property)) != null)
            {
                tfc.First(Find.ByName(prefix + "OLDate" + property)).TypeText(listdate);
            }

        }


        public void Enter_Data(string property)
        {

            

           



            
            
            if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxMLSNumber").Exists)
            {

                m2m.Form("formOrderEditWizard").TextField(property + "_tbxMLSNumber").TypeText(mlsnumber);
            }

            m2m.Form("formOrderEditWizard").TextField(property + "_tbxAddress1").TypeText(address);
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxCity").TypeText(city);
            m2m.Form("formOrderEditWizard").SelectList(property + "_ddlState").Select("IL");
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxZip").TypeText(zip);

            if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxGLA").Exists)
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxGLA").TypeText(grosslivingarea.Replace(",", ""));
            }
            else if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxFinishedSF").Exists)
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxFinishedSF").TypeText(grosslivingarea.Replace(",", ""));
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxAboveGradeSF").TypeText(grosslivingarea.Replace(",", ""));
            }

            if (m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGLASource").Exists)
            {
                m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGLASource").Select("MLS (Multiple Listing System)");
            }

            if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxOwner").Exists)
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxOwner").TypeText(ownername);
            }

            if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxSellerConcession").Exists)
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxSellerConcession").TypeText(sellerconcessions);
            }
            else if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxAdjustment").Exists)
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxAdjustment").TypeText(sellerconcessions);
            }


            if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxOriginalListDate").Exists)
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxOriginalListDate").TypeText(listdate);
            }

            m2m.Form("formOrderEditWizard").TextField(property + "_tbxOriginalListPrice").TypeText(origlistprice.TrimStart('$'));
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxSalePrice").TypeText(saleprice.TrimStart('$'));
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxSaleDate").TypeText(saledate);
            if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxCalculatedDistance").Text != null)
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxDistance").TypeText(m2m.Form("formOrderEditWizard").TextField(property + "_tbxCalculatedDistance").Text);
            }
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxDaysOnMarket").TypeText(daysonmarket);
            m2m.Form("formOrderEditWizard").SelectList(property + "_ddlStyle").Select(m2m.Form("formOrderEditWizard").SelectList("cstSubject_ddlStyle").SelectedItem);
            m2m.Form("formOrderEditWizard").SelectList(property + "_ddlExterior").Select(exterior);
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxYearBuilt").TypeText(yearbuilt);
            m2m.Form("formOrderEditWizard").SelectList(property + "_ddlCondition").Select(m2m.Form("formOrderEditWizard").SelectList("cstSubject_ddlCondition").SelectedItem);
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxTotalRooms").TypeText(totalrooms);
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxBedrooms").TypeText(bedrooms);

            if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxHalfBaths").Exists)
            {
                string[] temparr = baths.Split('.');
                if (temparr.Length < 2)
                {
                    m2m.Form("formOrderEditWizard").TextField(property + "_tbxHalfBaths").TypeText("0");
                    m2m.Form("formOrderEditWizard").TextField(property + "_tbxBathrooms").TypeText(temparr[0]);
                }
                else
                {
                    m2m.Form("formOrderEditWizard").TextField(property + "_tbxHalfBaths").TypeText(temparr[1]);
                    m2m.Form("formOrderEditWizard").TextField(property + "_tbxBathrooms").TypeText(temparr[0]);
                }
            }
            else
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxBathrooms").TypeText(baths);
            }



            System.Collections.Specialized.StringCollection contents = m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGarage").AllContents;


            string tmp = parking + "\\s+.*" + parkingaord;
            Regex ptype = new Regex(tmp, RegexOptions.IgnoreCase);

            if (parkingtype == "Garage")
            {
                switch (parking)
                {
                    case "1":

                        //m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGarage").SelectByValue(ptype);
                        m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGarage").Select("1 Stall - Attached");
                        break;

                    case "2":
                        m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGarage").Select("2 Stall - Attached");
                        break;
                    case "3":
                        m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGarage").Select("3 Stall - Attached");
                        break;
                }
            }
            else
            {
                if (contents.Contains("ON_SITE"))
                {
                    m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGarage").Select("ON_SITE");
                }

            }


            m2m.Form("formOrderEditWizard").TextField(property + "_tbxSiteSize").TypeText(lotsize);
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxComment").TypeText(comments);

            if (m2m.Form("formOrderEditWizard").SelectList(property + "_ddlComparison").Exists)
            {
                m2m.Form("formOrderEditWizard").SelectList(property + "_ddlComparison").Select("Equal");
            }

            if (m2m.Form("formOrderEditWizard").SelectList(property + "_tbxNoBasementRooms").Exists)
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxNoBasementRooms").TypeText("0");
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxPercentageBasementFinished").TypeText("0");
            } 



        }

        public void Fill_Sale1()
        {
            //attach to propery windows
            Attach();
            GetMLSData();
            
            //fill open m2m form
            //sale1

            if (bpocompany == "m2m")
                Enter_Data("cstSale1");
            else
                Enter_Data_emort("_1" , "CSales");


       
        }

        public void Fill_Sale2()
        {
            //attach to propery windows
            Attach();
            GetMLSData();

            //fill open m2m form
            //sale1
            if (bpocompany == "m2m")
                Enter_Data("cstSale2");
            else
                Enter_Data_emort("_2", "CSales");

            

        }

        public void Fill_Sale3()
        {
            //attach to propery windows
            Attach();
            GetMLSData();

            //fill open m2m form
            //sale3
            if (bpocompany == "m2m")
                Enter_Data("cstSale3");
            else
                Enter_Data_emort("_3", "CSales");
            
        }


        public void Enter_List_Comps(string property)
        {

            if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxMLSNumber").Exists)
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxMLSNumber").TypeText(mlsnumber);
            }

            m2m.Form("formOrderEditWizard").TextField(property + "_tbxAddress1").TypeText(address);
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxCity").TypeText(city);
            m2m.Form("formOrderEditWizard").SelectList(property + "_ddlState").Select("IL");
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxZip").TypeText(zip);

            if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxGLA").Exists)
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxGLA").TypeText(grosslivingarea.Replace(",", ""));
            }
            else if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxFinishedSF").Exists)
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxFinishedSF").TypeText(grosslivingarea.Replace(",", ""));
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxAboveGradeSF").TypeText(grosslivingarea.Replace(",", ""));
            }

            if (m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGLASource").Exists)
            {
                m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGLASource").Select("MLS (Multiple Listing System)");
            }

            if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxOwner").Exists)
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxOwner").TypeText(ownername);
            }

            if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxSellerConcession").Exists)
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxSellerConcession").TypeText(sellerconcessions);
            }
            else if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxAdjustment").Exists)
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxAdjustment").TypeText(sellerconcessions);
            }


            m2m.Form("formOrderEditWizard").TextField(property + "_tbxCurrentListPrice").TypeText(currentlistprice.TrimStart('$'));



            m2m.Form("formOrderEditWizard").TextField(property + "_tbxOriginalListDate").TypeText(listdate);
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxOriginalListPrice").TypeText(origlistprice.TrimStart('$'));

            if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxCalculatedDistance").Text != null)
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxDistance").TypeText(m2m.Form("formOrderEditWizard").TextField(property + "_tbxCalculatedDistance").Text);
            }
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxDaysOnMarket").TypeText(daysonmarket);
            m2m.Form("formOrderEditWizard").SelectList(property + "_ddlStyle").Select(m2m.Form("formOrderEditWizard").SelectList("cstSubject_ddlStyle").SelectedItem);
            m2m.Form("formOrderEditWizard").SelectList(property + "_ddlExterior").Select(exterior);
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxYearBuilt").TypeText(yearbuilt);
            m2m.Form("formOrderEditWizard").SelectList(property + "_ddlCondition").Select(m2m.Form("formOrderEditWizard").SelectList("cstSubject_ddlCondition").SelectedItem);
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxTotalRooms").TypeText(totalrooms);
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxBedrooms").TypeText(bedrooms);

            if (m2m.Form("formOrderEditWizard").TextField(property + "_tbxHalfBaths").Exists)
            {
                string[] temparr = baths.Split('.');
                if (temparr.Length < 2)
                {
                    m2m.Form("formOrderEditWizard").TextField(property + "_tbxHalfBaths").TypeText("0");
                    m2m.Form("formOrderEditWizard").TextField(property + "_tbxBathrooms").TypeText(temparr[0]);
                }
                else
                {
                    m2m.Form("formOrderEditWizard").TextField(property + "_tbxHalfBaths").TypeText(temparr[1]);
                    m2m.Form("formOrderEditWizard").TextField(property + "_tbxBathrooms").TypeText(temparr[0]);
                }

            }
            else
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxBathrooms").TypeText(baths);
            }




            //switch (parking)
            //{
            //    case "1":
            //        m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGarage").Select("1 Stall - Attached");
            //        break;

            //    case "2":
            //        m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGarage").Select("2 Stall - Attached");
            //        break;
            //    case "3":
            //        m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGarage").Select("3 Stall - Attached");
            //        break;
            //}

            System.Collections.Specialized.StringCollection contents = m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGarage").AllContents;


            string tmp = parking + "\\s+.*" + parkingaord;
            Regex ptype = new Regex(tmp, RegexOptions.IgnoreCase);

            if (parkingtype == "Garage")
            {
                switch (parking)
                {
                    case "1":

                        m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGarage").SelectByValue(ptype);

                        break;

                    case "2":
                        m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGarage").Select("2 Stall - Attached");
                        break;
                    case "3":
                        m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGarage").Select("3 Stall - Attached");
                        break;
                }
            }
            else
            {
                if (contents.Contains("ON_SITE"))
                {
                    m2m.Form("formOrderEditWizard").SelectList(property + "_ddlGarage").Select("ON_SITE");
                }

            }

            m2m.Form("formOrderEditWizard").TextField(property + "_tbxSiteSize").TypeText(lotsize);
            m2m.Form("formOrderEditWizard").TextField(property + "_tbxComment").TypeText(comments);

            if (m2m.Form("formOrderEditWizard").SelectList(property + "_ddlComparison").Exists)
            {
                m2m.Form("formOrderEditWizard").SelectList(property + "_ddlComparison").Select("Equal");
            }

            if (m2m.Form("formOrderEditWizard").SelectList(property + "_tbxNoBasementRooms").Exists)
            {
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxNoBasementRooms").TypeText("0");
                m2m.Form("formOrderEditWizard").TextField(property + "_tbxPercentageBasementFinished").TypeText("0");
            } 



        }

        public void Fill_List1()
        {
            //attach to propery windows
            Attach();
            GetMLSData();

            //fill open m2m form
            //list1

            if (bpocompany == "m2m")
                Enter_List_Comps("cstListing1");
            else
                Enter_Data_emort("_1", "CList");
           


        }

        public void Fill_List2()
        {
            //attach to propery windows
            Attach();
            GetMLSData();

            //fill open m2m form
            //list2
            if (bpocompany == "m2m")
                Enter_List_Comps("cstListing2");
            else
                Enter_Data_emort("_2", "CList");


        }

        public void Fill_List3()
        {
            //attach to propery windows
            Attach();
            GetMLSData();

            //fill open m2m form
            //list3
            if (bpocompany == "m2m")
                Enter_List_Comps("cstListing3");
            else
                Enter_Data_emort("_3", "CList");

            




        }

        public string CompleteStepOne (string ownername)

        {
            string[] splitname = ownername.Split(' ');
            string owner1lastname = splitname[0];
            string owner1firstname = "";
            string ownermiddleinitial = "";

            //we have name in this format: Last First M
            if (splitname.Length == 3 && splitname[2].Length == 1)
            {
                owner1firstname = splitname[1];
                ownermiddleinitial = splitname[2];
            }

            else if (splitname.Length == 3 && splitname[1].Length == 1)
            {
                //format Last M First
                owner1firstname = splitname[2];
                ownermiddleinitial = splitname[1];
            }
            else if (splitname.Length == 2)
            {
                owner1firstname = splitname[1];
                

            }

                // most likely this form Lastowner1 F Lastowner2 F
            else if (splitname.Length == 4)
            {
                owner1firstname = splitname[1];
                //owner2name = splitname[3] + " " + splitname[2];
            }

            return owner1firstname + " " + ownermiddleinitial + " " + owner1lastname;



        }


    }

    namespace PdfHelper
    {
        /// <summary>
        /// Taken from http://www.java-frameworks.com/java/itext/com/itextpdf/text/pdf/parser/LocationTextExtractionStrategy.java.html
        /// </summary>
        class LocationTextExtractionStrategyEx : LocationTextExtractionStrategy
        {
            private List<TextChunk> m_locationResult = new List<TextChunk>();
            private List<TextInfo> m_TextLocationInfo = new List<TextInfo>();
            public List<TextChunk> LocationResult
            {
                get { return m_locationResult; }
            }
            public List<TextInfo> TextLocationInfo
            {
                get { return m_TextLocationInfo; }
            }

            /// <summary>
            /// Creates a new LocationTextExtracationStrategyEx
            /// </summary>
            public LocationTextExtractionStrategyEx()
            {
            }

            /// <summary>
            /// Returns the result so far
            /// </summary>
            /// <returns>a String with the resulting text</returns>
            public override String GetResultantText()
            {
                m_locationResult.Sort();

                StringBuilder sb = new StringBuilder();
                TextChunk lastChunk = null;
                TextInfo lastTextInfo = null;
                foreach (TextChunk chunk in m_locationResult)
                {
                    if (lastChunk == null)
                    {
                        sb.Append(chunk.Text);
                        lastTextInfo = new TextInfo(chunk);
                        m_TextLocationInfo.Add(lastTextInfo);
                    }
                    else
                    {
                        if (chunk.sameLine(lastChunk))
                        {
                            float dist = chunk.distanceFromEndOf(lastChunk);

                            if (dist < -chunk.CharSpaceWidth)
                            {
                                sb.Append("xxx");
                                lastTextInfo.addSpace();
                            }
                            //append a space if the trailing char of the prev string wasn't a space && the 1st char of the current string isn't a space
                            else if (dist > chunk.CharSpaceWidth / 2.0f && chunk.Text[0] != ' ' && lastChunk.Text[lastChunk.Text.Length - 1] != ' ')
                            {
                                sb.Append("xxx");
                                lastTextInfo.addSpace();
                            }
                            sb.Append(chunk.Text);
                            lastTextInfo.appendText(chunk);
                        }
                        else
                        {
                            
                            sb.Append('\n');
                            sb.Append(chunk.Text);
                            lastTextInfo = new TextInfo(chunk);
                            m_TextLocationInfo.Add(lastTextInfo);
                        }
                    }
                    lastChunk = chunk;
                }
                return sb.ToString();
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="renderInfo"></param>
            public override void RenderText(TextRenderInfo renderInfo)
            {
                LineSegment segment = renderInfo.GetBaseline();
                TextChunk location = new TextChunk(renderInfo.GetText(), segment.GetStartPoint(), segment.GetEndPoint(), renderInfo.GetSingleSpaceWidth(), renderInfo.GetAscentLine(), renderInfo.GetDescentLine());
                m_locationResult.Add(location);
            }

            public class TextChunk : IComparable, ICloneable
            {
                string m_text;
                Vector m_startLocation;
                Vector m_endLocation;
                Vector m_orientationVector;
                int m_orientationMagnitude;
                int m_distPerpendicular;
                float m_distParallelStart;
                float m_distParallelEnd;
                float m_charSpaceWidth;

                public LineSegment AscentLine;
                public LineSegment DecentLine;

                public object Clone()
                {
                    TextChunk copy = new TextChunk(m_text, m_startLocation, m_endLocation, m_charSpaceWidth, AscentLine, DecentLine);
                    return copy;
                }

                public string Text
                {
                    get { return m_text; }
                    set { m_text = value; }
                }
                public float CharSpaceWidth
                {
                    get { return m_charSpaceWidth; }
                    set { m_charSpaceWidth = value; }
                }
                public Vector StartLocation
                {
                    get { return m_startLocation; }
                    set { m_startLocation = value; }
                }
                public Vector EndLocation
                {
                    get { return m_endLocation; }
                    set { m_endLocation = value; }
                }

                /// <summary>
                /// Represents a chunk of text, it's orientation, and location relative to the orientation vector
                /// </summary>
                /// <param name="txt"></param>
                /// <param name="startLoc"></param>
                /// <param name="endLoc"></param>
                /// <param name="charSpaceWidth"></param>
                public TextChunk(string txt, Vector startLoc, Vector endLoc, float charSpaceWidth, LineSegment ascentLine, LineSegment decentLine)
                {
                    m_text = txt;
                    m_startLocation = startLoc;
                    m_endLocation = endLoc;
                    m_charSpaceWidth = charSpaceWidth;
                    AscentLine = ascentLine;
                    DecentLine = decentLine;

                    m_orientationVector = m_endLocation.Subtract(m_startLocation).Normalize();
                    m_orientationMagnitude = (int)(Math.Atan2(m_orientationVector[Vector.I2], m_orientationVector[Vector.I1]) * 1000);

                    // see http://mathworld.wolfram.com/Point-LineDistance2-Dimensional.html
                    // the two vectors we are crossing are in the same plane, so the result will be purely
                    // in the z-axis (out of plane) direction, so we just take the I3 component of the result
                    Vector origin = new Vector(0, 0, 1);
                    m_distPerpendicular = (int)(m_startLocation.Subtract(origin)).Cross(m_orientationVector)[Vector.I3];

                    m_distParallelStart = m_orientationVector.Dot(m_startLocation);
                    m_distParallelEnd = m_orientationVector.Dot(m_endLocation);
                }

                /// <summary>
                /// true if this location is on the the same line as the other text chunk
                /// </summary>
                /// <param name="textChunkToCompare">the location to compare to</param>
                /// <returns>true if this location is on the the same line as the other</returns>
                public bool sameLine(TextChunk textChunkToCompare)
                {
                    if (m_orientationMagnitude != textChunkToCompare.m_orientationMagnitude) return false;
                    if (m_distPerpendicular != textChunkToCompare.m_distPerpendicular) return false;
                    return true;
                }

                /// <summary>
                /// Computes the distance between the end of 'other' and the beginning of this chunk
                /// in the direction of this chunk's orientation vector.  Note that it's a bad idea
                /// to call this for chunks that aren't on the same line and orientation, but we don't
                /// explicitly check for that condition for performance reasons.
                /// </summary>
                /// <param name="other"></param>
                /// <returns>the number of spaces between the end of 'other' and the beginning of this chunk</returns>
                public float distanceFromEndOf(TextChunk other)
                {
                    float distance = m_distParallelStart - other.m_distParallelEnd;
                    return distance;
                }

                /// <summary>
                /// Compares based on orientation, perpendicular distance, then parallel distance
                /// </summary>
                /// <param name="obj"></param>
                /// <returns></returns>
                public int CompareTo(object obj)
                {
                    if (obj == null) throw new ArgumentException("Object is now a TextChunk");

                    TextChunk rhs = obj as TextChunk;
                    if (rhs != null)
                    {
                        if (this == rhs) return 0;

                        int rslt;
                        rslt = m_orientationMagnitude - rhs.m_orientationMagnitude;
                        if (rslt != 0) return rslt;

                        rslt = m_distPerpendicular - rhs.m_distPerpendicular;
                        if (rslt != 0) return rslt;

                        // note: it's never safe to check floating point numbers for equality, and if two chunks
                        // are truly right on top of each other, which one comes first or second just doesn't matter
                        // so we arbitrarily choose this way.
                        rslt = m_distParallelStart < rhs.m_distParallelStart ? -1 : 1;

                        return rslt;
                    }
                    else
                    {
                        throw new ArgumentException("Object is now a TextChunk");
                    }
                }
            }

            public class TextInfo
            {
                public Vector TopLeft;
                public Vector BottomRight;
                private string m_Text;

                public string Text
                {
                    get { return m_Text; }
                }

                /// <summary>
                /// Create a TextInfo.
                /// </summary>
                /// <param name="initialTextChunk"></param>
                public TextInfo(TextChunk initialTextChunk)
                {
                    TopLeft = initialTextChunk.AscentLine.GetStartPoint();
                    BottomRight = initialTextChunk.DecentLine.GetEndPoint();
                    m_Text = initialTextChunk.Text;
                }

                /// <summary>
                /// Add more text to this TextInfo.
                /// </summary>
                /// <param name="additionalTextChunk"></param>
                public void appendText(TextChunk additionalTextChunk)
                {
                    BottomRight = additionalTextChunk.DecentLine.GetEndPoint();
                    m_Text += additionalTextChunk.Text;
                }

                /// <summary>
                /// Add a space to the TextInfo.  This will leave the endpoint out of sync with the text.
                /// The assumtion is that you will add more text after the space which will correct the endpoint.
                /// </summary>
                public void addSpace()
                {
                    m_Text += ' ';
                }


            }
        }
    }

    public class MarketStats
    {

        public string totalSold;
        public string totalActive;
        public string totalPending;
        public string totalOnMarket;

        public string [] listPrice = {"max", "average", "median", "min"};
        public string [] soldPrice = { "max", "average", "median", "min" };
 
        
    }

    public class CompCriteria
    {
        
        //nabpop criteria
        public bool Age(string subjectYearBuilt, string compYearBuilt)
        {
                bool result = false;
                DateTime st = new DateTime((Convert.ToInt32(subjectYearBuilt)), 1, 1);
                TimeSpan ts = DateTime.Now - st;
                int subjectAge = ts.Days / 365;

                DateTime ct = new DateTime((Convert.ToInt32(compYearBuilt)), 1, 1);
                 ts = DateTime.Now - ct;
                int compAge = ts.Days / 365;

            int intSubjectYearBuilt = Convert.ToInt16(subjectYearBuilt);
            int intCompYearBuilt = Convert.ToInt16(compYearBuilt);

            int ageDiff =  (Math.Abs( subjectAge - compAge ));


            if (GlobalVar.ccc == GlobalVar.CompCompareSystem.NABPOP)
            {
                if (subjectAge <= 10 && ageDiff <= 5)
                {
                    return true;
                }
                else if (subjectAge > 10 && subjectAge < 20 && ageDiff <= subjectAge / 2)
                {
                    return true;
                }
                else if (subjectAge >= 20 && subjectAge <= 30 && ageDiff <= 10)
                {
                    return true;
                }
                else if (subjectAge > 30 && subjectAge <= 50 && ageDiff <= 15)
                {
                    return true;
                }
                else if (subjectAge > 50 && subjectAge <= 75 && ageDiff <= 20)
                {
                    return true;
                }
                else if (subjectAge > 75 && ageDiff <= 25)
                {
                    return true;
                }
            } else if (GlobalVar.ccc == GlobalVar.CompCompareSystem.USER)
            {
                if (subjectAge <= 10 && ageDiff <= 5)
                {
                    return true;
                }
                else if (subjectAge > 10 && subjectAge < 20 && ageDiff <= subjectAge / 2)
                {
                    return true;
                }
                else if (subjectAge >= 20 && subjectAge <= 30 && ageDiff <= 15)
                {
                    return true;
                }
                else if (subjectAge > 30 && subjectAge <= 50 && ageDiff <= 20)
                {
                    return true;
                }
                else if (subjectAge > 50 && subjectAge <= 75 && ageDiff <= 25)
                {
                    return true;
                }
                else if (subjectAge > 75 && ageDiff <= 30)
                {
                    return true;
                }
            }
            return result;

        }

        public bool LotSize(string subjectLotSize, string compLotSize)
        {
            bool result = false;

            if (compLotSize == "")
            {
                compLotSize = "0";
            }


            double sLotSize = Convert.ToDouble(subjectLotSize);
            double cLotSize = Convert.ToDouble(compLotSize);



            if (GlobalVar.ccc == GlobalVar.CompCompareSystem.NABPOP)
            {

                //if (
                //condos
                if (sLotSize == 0 && cLotSize < 0.1)
                {
                    return true;
                }

                if (sLotSize <= 1)
                {
                    if (cLotSize * 0.7 <= sLotSize && sLotSize <= cLotSize * 1.3)
                    {
                        return true;
                    }
                }
                else if (sLotSize > 1 && sLotSize < 3)
                {
                    if (cLotSize - 0.5 <= sLotSize && sLotSize <= cLotSize + 0.5)
                    {
                        return true;
                    }
                }
                else if (sLotSize >= 3 && sLotSize < 6)
                {
                    if (cLotSize - 1 <= sLotSize && sLotSize <= cLotSize + 1)
                    {
                        return true;
                    }
                }
                else if (sLotSize >= 6 && sLotSize < 11)
                {
                    if (cLotSize - 2 <= sLotSize && sLotSize <= cLotSize + 2)
                    {
                        return true;
                    }
                }
                else if (sLotSize >= 11)
                {
                    if (cLotSize * 0.8 <= sLotSize && sLotSize <= cLotSize * 1.2)
                    {
                        return true;
                    }
                }
            }
            else if (GlobalVar.ccc == GlobalVar.CompCompareSystem.USER)
            {
                if (GlobalVar.lowerLotSize <= Convert.ToDecimal(cLotSize) && Convert.ToDecimal(cLotSize) <= GlobalVar.upperLotSize)
                {
                    return true;
                }
            }

            return result;

        }

    }

    public class MRED_Driver
    {
        private readonly IYourForm form;
        public MRED_Driver(IYourForm form)
        {
            this.form = form;
        }
        //private iMacros.App driver;
        private int timeout = 30;

        public  void NewMapSearch( iMacros.App d)
        {

            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !TIMEOUT_STEP 30");
            macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=TXT:Search");
            macro.AppendLine(@"FRAME NAME=workspace");
            macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:Click<SP>here<SP>to<SP>select<SP>boundaries");
            macro.AppendLine(@"TAG POS=1 TYPE=SPAN FORM=NAME:dc ATTR=TXT:Center<SP>On...");
            macro.AppendLine(@"wait seconds=1");
            
            //hardcoded         
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:dc ATTR=ID:address CONTENT=172<SP>Morningside<SP>Ln<SP>W,<SP>Buffalo<SP>Grove,<SP>IL<SP>60089");

            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:dc ATTR=ID:address CONTENT=" + form.SubjectFullAddress.Replace(" ","<SP>"));
            
            macro.AppendLine(@"wait seconds=1");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=VALUE:Go");
            macro.AppendLine(@"wait seconds=1");
            macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:dc ATTR=SRC:*/circle.png");
            macro.AppendLine(@"wait seconds=1");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:dc ATTR=ID:distanceInput CONTENT=1");
            macro.AppendLine(@"wait seconds=1");

            macro.AppendLine(@"SAVEAS TYPE=PNG FOLDER=" + form.SubjectFilePath.Replace(" ", "<SP>") + " FILE=search_map_" + DateTime.Now.Ticks.ToString() + ".png");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=ID:OK&&VALUE:OK");
        
            string macroCode = macro.ToString();
           // d.iimOpen("-ie", false, 30);
            d.iimPlayCode(macroCode, 30);
        }

        public void Save1MiSearch(iMacros.App d)
        {
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"FRAME NAME=workspace");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:dc ATTR=ID:searchButtonTop&&VALUE:View<SP>Results");
            macro.AppendLine(@"FRAME NAME=subheader");
            macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:header ATTR=ID:STATS_BLOCK");
            macro.AppendLine(@"FRAME F=5");
            macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=ID:printLink");
            macro.AppendLine(@"FRAME NAME=workspace");
            macro.AppendLine(@"ONPRINT P=8 BUTTON=PRINT");
            macro.AppendLine(@"TAG POS=1 TYPE=IFRAME FORM=NAME:dc ATTR=ID:statisticsFrame");
            macro.AppendLine(@"FRAME NAME=subheader");
            macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:header ATTR=ID:refineSearchLink");
            string macroCode = macro.ToString();
            d.iimPlayCode(macroCode, timeout);
        }

        public void SaveStats(iMacros.App d, MarketStats m)
        {
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"FRAME NAME=workspace");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:dc ATTR=ID:searchButtonTop&&VALUE:View<SP>Results");
            macro.AppendLine(@"FRAME NAME=subheader");
            macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:header ATTR=ID:STATS_BLOCK");
            macro.AppendLine(@"WAIT SECONDS=1");

            macro.AppendLine(@"FRAME F=5");
            macro.AppendLine(@"TAG POS=3 TYPE=TABLE FORM=NAME:dc ATTR=TXT:* EXTRACT=TXT");
           // macro.AppendLine(@"SAVEAS TYPE=EXTRACT FOLDER=" + form.SubjectFilePath.Replace(" ", "<SP>") +" FILE=mytable_{{!NOW:yymmdd_hhnnss}}.csv");
            macro.AppendLine(@"TAG POS=4 TYPE=TABLE FORM=NAME:dc ATTR=TXT:* EXTRACT=TXT");
           // macro.AppendLine(@"SAVEAS TYPE=EXTRACT FOLDER=" + form.SubjectFilePath.Replace(" ", "<SP>") + " FILE=mytable_{{!NOW:yymmdd_hhnnss}}.csv");
            macro.AppendLine(@"");
            macro.AppendLine(@"WAIT SECONDS=3");

            string macroCode = macro.ToString();
            d.iimPlayCode(macroCode, timeout);

            string t1 = d.iimGetExtract(1);
            string t2 = d.iimGetExtract(2);

            string pattern = "Sold#NEXT##NEXT#(\\d+)#NEXT#";
            Match match = Regex.Match(t1, pattern);
            m.totalSold = match.Groups[1].Value;

            pattern = "Under Contract#NEXT##NEXT#(\\d+)#NEXT#";
            match = Regex.Match(t1, pattern);
            m.totalPending = match.Groups[1].Value;

            pattern = "Active#NEXT##NEXT#(\\d+)#NEXT#";
            match = Regex.Match(t1, pattern);
            m.totalActive = match.Groups[1].Value;

            m.totalOnMarket = (Convert.ToInt16(m.totalActive) + Convert.ToInt16(m.totalPending)).ToString();


            //listprice max, avg,med,min
             pattern = "#NEWLINE#LP\\s*[^#]+#NEXT##NEXT#([^#]+)#NEXT#([^#]+)#NEXT#([^#]+)#NEXT#([^#]+)#NEXT#";
             match = Regex.Match(t2, pattern);
            m.listPrice[0] = match.Groups[1].Value;
            m.listPrice[1] = match.Groups[2].Value;
            m.listPrice[2] = match.Groups[3].Value;
            m.listPrice[3] = match.Groups[4].Value;

            pattern = "#NEWLINE#SP\\s*[^#]+#NEXT##NEXT#([^#]+)#NEXT#([^#]+)#NEXT#([^#]+)#NEXT#([^#]+)";
            match = Regex.Match(t2, pattern);
            m.soldPrice[0] = match.Groups[1].Value;
            m.soldPrice[1] = match.Groups[2].Value;
            m.soldPrice[2] = match.Groups[3].Value;
            m.soldPrice[3] = match.Groups[4].Value;

            
            
            //MessageBox.Show(t1);

        }

        private void SetCompSearch(string a)
        {

        }

        public void FindComps(iMacros.App d, bool newSearch)
        {
            int timeout = 90;
            StringBuilder macro = new StringBuilder();
            string test = "";

            //are we on the search screen
            macro.AppendLine(@"SET !TIMEOUT_STEP 30");
            macro.AppendLine(@"FRAME NAME=workspace");
            macro.AppendLine(@"TAG POS=10 TYPE=TD FORM=NAME:dc ATTR=TXT:* EXTRACT=TXT");
            d.iimPlayCode(macro.ToString(), timeout);
            test = d.iimGetLastExtract();

            if (test.Contains("* Status"))
            {
                if (newSearch)
                {
                    macro.AppendLine(@"SET !ERRORIGNORE YES");
                    macro.AppendLine(@"FRAME NAME=workspace");
                    macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:dc ATTR=ID:picklisticon");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST0&&VALUE:ACTV CONTENT=YES");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST1&&VALUE:AUCT CONTENT=YES");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST2&&VALUE:BOMK CONTENT=YES");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST3&&VALUE:CTG CONTENT=YES");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST4&&VALUE:NEW CONTENT=YES");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST5&&VALUE:PCHG CONTENT=YES");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST6&&VALUE:RACT CONTENT=YES");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST7&&VALUE:TEMP CONTENT=YES");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST1&&VALUE:CLSD CONTENT=YES");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST3&&VALUE:PEND CONTENT=YES");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=VALUE:OK");



                    //  macro.AppendLine(@"FRAME NAME=workspace");
                    // macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=NAME:dc ATTR=TXT:1<SP>Month2<SP>Months3<SP>Months4<SP>Months5<SP>Months6<SP>Months9<SP>Months12<SP>Months24<SP>MonthsAll<SP>Mo*");
                    macro.AppendLine(@"TAG POS=6 TYPE=IMG FORM=NAME:dc ATTR=ID:picklisticon");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBMONTHS_BACK5&&VALUE:6<SP>Months CONTENT=YES");
                    macro.AppendLine(@"TAG POS=6 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=VALUE:OK");

                    //macro.AppendLine(@"TAG POS=6 TYPE=IMG FORM=NAME:dc ATTR=ID:picklisticon");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBMONTHS_BACK7&&VALUE:12<SP>Months CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=1 TYPE=SPAN FORM=NAME:dc ATTR=ID:1DISP");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBMONTHS_BACK0&&VALUE:6<SP>Months CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=6 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=VALUE:OK");


                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:dc ATTR=ID:MONTHS_BACKID CONTENT=12<SP>Months");
                    //macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:ui-active-menuitem");
                    // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=ID:countButtonTop&&VALUE:Count");
                    // macro.AppendLine(@"TAG POS=0 TYPE=INPUT:TEXT FORM=NAME:dc ATTR=ID:MONTHS_BACKID CONTENT=1<SP>Month");

                    //
                    //attached search
                    //
                    macro.AppendLine(@"TAG POS=5 TYPE=IMG FORM=NAME:dc ATTR=ID:picklisticon");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBMONTHS_BACK5&&VALUE:6<SP>Months CONTENT=YES");
                    macro.AppendLine(@"TAG POS=5 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=VALUE:OK");
                    macro.AppendLine(@"SAVEAS TYPE=PNG FOLDER=" + form.SubjectFilePath.Replace(" ", "<SP>") + " FILE=search_page_" + DateTime.Now.Ticks.ToString() + ".png");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:dc ATTR=ID:searchButtonTop&&VALUE:View<SP>Results");
                    macro.AppendLine(@"TAG POS=11 TYPE=SPAN FORM=NAME:dc ATTR=CLASS:columnHeader");
                    macro.AppendLine(@"FRAME NAME=subheader");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:header ATTR=ID:report_type CONTENT=%agentfull");
                }
                else
                {
                    macro.AppendLine(@"SAVEAS TYPE=PNG FOLDER=" + form.SubjectFilePath.Replace(" ", "<SP>") + " FILE=search_page_" + DateTime.Now.Ticks.ToString() + ".png");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:dc ATTR=ID:searchButtonTop&&VALUE:View<SP>Results");
                    macro.AppendLine(@"TAG POS=11 TYPE=SPAN FORM=NAME:dc ATTR=CLASS:columnHeader");
                    macro.AppendLine(@"FRAME NAME=subheader");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:header ATTR=ID:report_type CONTENT=%agentfull");
                    //macro.AppendLine(@"TAG POS=2 TYPE=SPAN FORM=NAME:dc ATTR=CLASS:Label EXTRACT=TXT");
                }

                string macroCode = macro.ToString();
                d.iimPlayCode(macroCode, timeout);
            }
            else
            {
                macro.Clear();
                macro.AppendLine(@"FRAME NAME=subheader");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:header ATTR=ID:report_type EXTRACT=TXT");
                d.iimPlayCode(macro.ToString(), timeout);

                test = d.iimGetLastExtract();
                //is the list of properties already open
                if (!test.Contains("Full - Agent"))
                {
                    //if not then break
                    MessageBox.Show("Search Not Setup");
                    return;
                }
            }


            MlsReportDriver r = new MlsReportDriver(form);
            r.ReadMlsSheets(d);
            
            

          
    

          

          // MessageBox.Show("done");

            //this works, resets months and rescans comps
           //macro.Clear();
           //macro.AppendLine(@"FRAME NAME=subheader");
           //macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:header ATTR=ID:refineSearchLink");

           //macro.AppendLine(@"FRAME NAME=workspace");
           //macro.AppendLine(@"TAG POS=6 TYPE=IMG FORM=NAME:dc ATTR=ID:picklisticon");
           //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBMONTHS_BACK0&&VALUE:1<SP>Month CONTENT=NO");
           //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBMONTHS_BACK1&&VALUE:2<SP>Months CONTENT=YES");
           //macro.AppendLine(@"TAG POS=6 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=VALUE:OK");



        

          

         
           //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:dc ATTR=ID:searchButtonTop&&VALUE:View<SP>Results");
           //macro.AppendLine(@"FRAME NAME=subheader");
           //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:header ATTR=ID:report_type CONTENT=%agentfull");
           //d.iimPlayCode(macro.ToString(), timeout);
           //r.ReadMlsSheets(d);

            //
            //Active listings
            //
          // macro.Clear();
          // macro.AppendLine(@"FRAME NAME=workspace");
          // macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:dc ATTR=ID:picklisticon");
          //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST0&&VALUE:ACTV CONTENT=YES");
          //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST1&&VALUE:AUCT CONTENT=YES");
          //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST2&&VALUE:BOMK CONTENT=YES");
          //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST3&&VALUE:CTG CONTENT=YES");
          //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST4&&VALUE:NEW CONTENT=YES");
          //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST5&&VALUE:PCHG CONTENT=YES");
          //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST6&&VALUE:RACT CONTENT=YES");
          //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST7&&VALUE:TEMP CONTENT=YES");
          //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST1&&VALUE:CLSD CONTENT=NO");
          //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST3&&VALUE:PEND CONTENT=YES");
          //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:allACTIVES&&VALUE:on CONTENT=YES");
          // //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST0&&VALUE:ACTV CONTENT=NO");
          // //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST1&&VALUE:AUCT CONTENT=NO");
          // //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST2&&VALUE:BOMK CONTENT=NO");
          // //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST3&&VALUE:CTG CONTENT=NO");
          // //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST4&&VALUE:NEW CONTENT=NO");
          // //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST5&&VALUE:PCHG CONTENT=NO");
          // //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST6&&VALUE:RACT CONTENT=NO");
          // //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST7&&VALUE:TEMP CONTENT=NO");
          // //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:allACTIVES&&VALUE:on CONTENT=NO");
          // //macro.AppendLine(@"TAG POS=1 TYPE=DIV FORM=NAME:dc ATTR=ID:STFR");
          // //macro.AppendLine(@"TAG POS=1 TYPE=DIV FORM=NAME:dc ATTR=ID:STdiv");
          //// macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST3&&VALUE:PEND CONTENT=YES");
          // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=VALUE:OK");
          // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:dc ATTR=ID:searchButtonTop&&VALUE:View<SP>Results");
          // macro.AppendLine(@"FRAME NAME=subheader");
          // macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:header ATTR=ID:report_type CONTENT=%agentfull");


          //  d.iimPlayCode(macro.ToString(), timeout);
          //  r.ReadMlsSheets(d);

        }

    }

    public static class LINQExtension
    {
        public static decimal Median(this IEnumerable<decimal> source)
        {
            if (source.Count() == 0)
            {
                throw new InvalidOperationException("Cannot compute median for an empty set.");
            }

            var sortedList = from number in source
                             orderby number
                             select number;

            int itemIndex = (int)sortedList.Count() / 2;

            if (sortedList.Count() % 2 == 0)
            {
                // Even number of items.
                return (sortedList.ElementAt(itemIndex) + sortedList.ElementAt(itemIndex - 1)) / 2;
            }
            else
            {
                // Odd number of items.
                return sortedList.ElementAt(itemIndex);
            }
        }

        public static int Median(this IEnumerable<int> source)
        {
            if (source.Count() == 0)
            {
                throw new InvalidOperationException("Cannot compute median for an empty set.");
            }

            var sortedList = from number in source
                             orderby number
                             select number;

            int itemIndex = (int)sortedList.Count() / 2;

            if (sortedList.Count() % 2 == 0)
            {
                // Even number of items.
                return (sortedList.ElementAt(itemIndex) + sortedList.ElementAt(itemIndex - 1)) / 2;
            }
            else
            {
                // Odd number of items.
                return sortedList.ElementAt(itemIndex);
            }
        }
    }
}

