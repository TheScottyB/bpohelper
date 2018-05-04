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
//using DotNetOpenAuth.OAuth2;
//using Google.Apis.Authentication;
//using Google.Apis.Authentication.OAuth2;
//using Google.Apis.Authentication.OAuth2.DotNetOpenAuth;
//using Google.Apis.Samples.Helper;
using System.Xml;
using System.Xml.Schema;
using System.Runtime.Serialization;
using System.Web;
using Newtonsoft.Json;
using ImageResizer;


 public interface IYourForm
    {
        bool IgnoreAge { get; }
        bool IgnoreType { get; }
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
        bpohelper._BPO_SandboxDataSet1TableAdapters.RawSFDataTableAdapter RawSFDatatable { get; }
        string SubjectSubdivision { get; set; }
        string SubjectMlsType { get; set; }
        bpohelper.Neighborhood SubjectNeighborhood { get; set; }
        bpohelper.Neighborhood SetOfComps { get; set; }
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
        string SearchMapRadius { get;  }
        string SubjectAssessmentValue { get; set; }
        string SubjectLandValue { get; set; }
        string SubjectMarketValue { get; set; }
        bool PerserveNeighorhoodData { get; }
        bool PerserveCompSetData { get; }
    }

namespace bpohelper
{
    public partial class Form1 : System.Windows.Forms.Form, IYourForm
    {
        iMacros.App iim = new iMacros.App();
        iMacros.App iim2 = new iMacros.App();
        iMacros.Status status;
        public Neighborhood subjectNeighborhood;
        public Neighborhood setOfComps;
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

            GlobalVar.mainWindow = this;

            iMacros.App browser1 = new iMacros.App();
            iMacros.App browser2 = new iMacros.App();
       //     iMacros.App imBrowser = new iMacros.App();

            status = browser1.iimOpen("-ie", false, 60);
          //  status = imBrowser.iimOpen("", false, 60);

            //if (status == Status.sOk)
            //{
            //    browser1 = imBrowser;
            //}

            browser1.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");

           



            string b1Url = browser1.iimGetLastExtract();

            //we want iim to be linked to connectMLS

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
              //  DialogResult response = AlertMessageWithCustomHelpWindow();

             //   MessageBox.Show(response.ToString());
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
             Neighborhood setOfComps = new Neighborhood();
             SubjectNeighborhood = subjectNeighborhood;
             SetOfComps = setOfComps;
            SubjectAssessmentInfo = subjectAssessmentInfo;
            TypeDetachedList tdl = new TypeDetachedList();
            subjectTownshipRecord = new TownshipReport();
            foreach (string key in tdl.mlsTypeDetached.Keys)
            {
                subjectMlsTypecomboBox.Items.Add(tdl.mlsTypeDetached[key]);
            }

            string line = "";
            string[] splitLine;
            try
            {
                using (System.IO.StreamReader file = new System.IO.StreamReader(".config"))
                {
                    while (!file.EndOfStream)
                    {
                        line = file.ReadLine();
                        splitLine = line.Split(';');
                        if (splitLine[0] == "lastOpenedSubject")
                        {
                            var directions = MessageBox.Show("Load Prior Subject?", "Sup...Yo", MessageBoxButtons.YesNo);
                            if (directions == System.Windows.Forms.DialogResult.Yes)
                            {
                                search_address_textbox.Text = splitLine[1];
                                importSubjectInfoButton(this, new EventArgs());
                            }

                        }

                    }
                }
            }
            catch { }

        }

        //
        //Helpers
        //

        // Display a message box with a Help button. Show a custom Help window
        // by handling the HelpRequested event.
        private DialogResult AlertMessageWithCustomHelpWindow()
        {
            // Handle the HelpRequested event for the following message.
            this.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.Form1_HelpRequested);

            this.Tag = "Message with Help button.";

            // Show a message box with OK and Help buttons.
            DialogResult r = MessageBox.Show("Message with Help button.",
                                              "Help Caption", MessageBoxButtons.OK,
                                              MessageBoxIcon.Question,
                                              MessageBoxDefaultButton.Button1,
                                              0, true);

            // Remove the HelpRequested event handler to keep the event
            // from being handled for other message boxes.
            this.HelpRequested -= new System.Windows.Forms.HelpEventHandler(this.Form1_HelpRequested);

            // Return the dialog box result.
            return r;
        }

        private void Form1_HelpRequested(System.Object sender, System.Windows.Forms.HelpEventArgs hlpevent)
        {
            // Create a custom Help window in response to the HelpRequested event.
            System.Windows.Forms.Form helpForm = new System.Windows.Forms.Form();

            // Set up the form position, size, and title caption.
            helpForm.StartPosition = FormStartPosition.Manual;
            helpForm.Size = new Size(200, 400);
            helpForm.DesktopLocation = new Point(this.DesktopBounds.X +
                                                  this.Size.Width,
                                                  this.DesktopBounds.Top);
            helpForm.Text = "Help Form";

            // Create a label to contain the Help text.
            System.Windows.Forms.Label helpLabel = new System.Windows.Forms.Label();

            // Add the label to the form and set its text.
            helpForm.Controls.Add(helpLabel);
            helpLabel.Dock = DockStyle.Fill;

            // Use the sender parameter to identify the context of the Help request.
            // The parameter must be cast to the Control type to get the Tag property.
            System.Windows.Forms.Control senderControl = sender as System.Windows.Forms.Control;

            helpLabel.Text = "Help information shown in response to user action on the '" +
                              (string)senderControl.Tag + "' message.";

            // Set the Help form to be owned by the main form. This helps
            // to ensure that the Help form is disposed of.
            this.AddOwnedForm(helpForm);

            // Show the custom Help window.
            helpForm.Show();

            // Indicate that the HelpRequested event is handled.
            hlpevent.Handled = true;
        }



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

    

        private void subjectBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.subjectBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this._BPO_SandboxDataSet);

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

            //MessageBox.Show(currentUrl);


            if (currentUrl.ToLower().Contains("swbcls-forms"))
            {
                {
                    SWBC bpoform = new SWBC();
                    streetnumTextBox.Text = "swbc";
                    bpoform.Prefill(iim2, this);
                }
            }

            #region sandcastle
            if (currentUrl.ToLower().Contains("sandcastlefs"))

           
            {
                Sandcastle bpoform = new Sandcastle();
                streetnumTextBox.Text = "sandcastlefs";
                bpoform.Prefill(iim2, this);
            }
            #endregion

            #region pyramid
            if (currentUrl.ToLower().Contains("pyramidplatform"))

            //  if (currentUrl.ToLower().Contains("amoservices"))
            {
                Pyramid bpoform = new Pyramid();
                streetnumTextBox.Text = "Pyramid";
                bpoform.Prefill(iim2, this);
            }
            #endregion

            #region AMO-swbc
            if (currentUrl.ToLower().Contains("amoservices"))

          //  if (currentUrl.ToLower().Contains("amoservices"))
            {
                AMO bpoform = new AMO();
                streetnumTextBox.Text = "amoservices";
                bpoform.Prefill(iim2, this);
            }
            #endregion

            #region ClearCap
            if (currentUrl.ToLower().Contains("clearcapital"))
            {
                ClearCap bpoform = new ClearCap();
                streetnumTextBox.Text = "ClearCap";
                bpoform.Prefill(iim2, this);
            }
            #endregion

            #region inside valuation
            if (currentUrl.ToLower().Contains("insidevaluation"))
            {
                InsideValuation bpoform = new InsideValuation();
                streetnumTextBox.Text = "InsideVal";
                bpoform.Prefill(iim2, this);
            }
            #endregion

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
            //oringinally dispo
            #region excelerasSummit
            if (currentUrl.ToLower().Contains("exceleras"))
            {
                //Dispo bpoform = new Dispo();
                Exceleras bpoform = new Exceleras();
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
            if (currentUrl.ToLower().Contains("lres"))
            {
                Lres lres = new Lres();
                streetnumTextBox.Text = "lres";
                lres.Prefill(iim2, this);
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

            if (currentUrl.ToLower().Contains("valuationpartners"))
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
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtREPORTDT CONTENT=" + DateTime.Now.ToShortDateString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form1 ATTR=ID:rbREVIEWTYPE1&&VALUE:EXTERIOR CONTENT=YES");
                string subjectComments = "The subject is a conforming home within the neighborhood. No adverse conditions were noted at the time of inspection based on exterior observations. Unable to determine interior condition due to exterior inspection only, so subject was assumed to be in average condition for this report.".Replace(" ", "<SP>");
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:form1 ATTR=ID:txtREVIEWSUBJECTCOMMENTS CONTENT=" + subjectComments);
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
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTBASEMENT CONTENT=" + SubjectBasementGLA);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTBASEMENTFIN CONTENT=" + SubjectBasementFinishedGLA);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTGARAGE CONTENT=" + subjectParkingTypeTextbox.Text.Replace(" ", "<SP>"));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTGARAGENUMCARS CONTENT=" + Regex.Match(SubjectParkingType, @"/d").Value);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtSUBJECTCONDITION CONTENT=Avg");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtNUMCOMPLISTINGS CONTENT=2");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtCOMPLISTINGPRICELOW");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtNUMCOMPLISTINGS CONTENT=3");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtCOMPLISTINGPRICELOW CONTENT=550000");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtCOMPLISTINGPRICEHIGH CONTENT=899900");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtNUMCOMPSALES CONTENT=1");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtCOMPSALESPRICELOW CONTENT=650000");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txtCOMPSALESPRICEHIGH CONTENT=650000");
                //macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:form1 ATTR=ID:txtREVIEWNBRHOODDESC CONTENT=Secluded/Isolated<SP>Area.<SP><SP>Rural<SP>like,<SP>but<SP>close<SP>to<SP>amenitites.");
                
                string subjectMarketComments = "Stable market with about 20% REO sales mixed with short sales and traditional. High demand under 150k. Seller concessions are not typical for the area.".Replace(" ","<SP>");
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:form1 ATTR=ID:txtMARKETCONDITIONS CONTENT=" + subjectMarketComments);

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:form1 ATTR=ID:SaveWork11&&VALUE:Save<SP>Work");

            }

            #endregion vp

            #region usres prefill

            if (currentUrl.ToLower().Contains("usres") || currentUrl.ToLower().Contains("res.net"))
            {
                if (currentUrl.Contains("FormStaticFifthThirdView"))
                {
                    Resnet bpoform = new Resnet();
                    bpoform.Prefill(iim2, this);
                }
                else
                {
                    #region orignal usres flow
                    StringBuilder read_header = new StringBuilder();


                    read_header.AppendLine(@"TAG POS=4 TYPE=TABLE FORM=NAME:InputForm ATTR=CLASS:form_txt_blk EXTRACT=TXT");
                    string read_headerCode = read_header.ToString();
                    status = iim2.iimPlayCode(read_headerCode, 30);


                    string table1 = iim2.iimGetLastExtract(1);


                    string pattern = "IL.\\s+(\\d\\d\\d\\d\\d)";
                    Match match = Regex.Match(iim2.iimGetLastExtract(1), pattern);
                    string orderzip = match.Groups[1].Value;

                    macro.AppendLine(@"SET !ERRORIGNORE YES");
                    macro.AppendLine(@"SET !TIMEOUT_STEP 0");

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


                    //// C# snippet generated by iMacros Editor.
                    //// See http://wiki.imacros.net/Web_Scripting for details on how to use the iMacros Scripting Interface.

                    //// iMacros.AppClass iim = new iMacros.AppClass();
                    //// iMacros.Status status = iim.iimOpen("", true, timeout);
                    //StringBuilder macro = new StringBuilder();
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:frmUpload ATTR=NAME:BpoPic_46369435_35 CONTENT=D:\DropBox\Dropbox\Listing\01-ACTV\27<SP>Beachview\reports\BPO\20170613\L1.jpg");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:frmUpload ATTR=NAME:BpoPic_46369435_40 CONTENT=D:\DropBox\Dropbox\Listing\01-ACTV\27<SP>Beachview\reports\BPO\20170613\L2.jpg");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:frmUpload ATTR=NAME:BpoPic_46369435_45 CONTENT=D:\DropBox\Dropbox\Listing\01-ACTV\27<SP>Beachview\reports\BPO\20170613\L3.jpg");
                    //macro.AppendLine(@"");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:frmUpload ATTR=NAME:BpoPic_46369435_50 CONTENT=D:\DropBox\Dropbox\Listing\01-ACTV\27<SP>Beachview\reports\BPO\20170613\S1.jpg");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:frmUpload ATTR=NAME:BpoPic_46369435_55 CONTENT=D:\DropBox\Dropbox\Listing\01-ACTV\27<SP>Beachview\reports\BPO\20170613\S2.jpg");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:frmUpload ATTR=NAME:BpoPic_46369435_60 CONTENT=D:\DropBox\Dropbox\Listing\01-ACTV\27<SP>Beachview\reports\BPO\20170613\S3.jpg");
                    //macro.AppendLine(@"");
                    //macro.AppendLine(@"");
                    //macro.AppendLine(@"'6 ext");
                    //macro.AppendLine(@"'3 req front, rear, street scene");
                    //macro.AppendLine(@"'3 optional right, left, misc");
                    //macro.AppendLine(@"");w
                    //macro.AppendLine(@"'12 int");
                    //macro.AppendLine(@"'5 req");
                    //string macroCode = macro.ToString();
                    //// status = iim.iimPlayCode(macroCode, timeout);


                    //
                    //Save CMA and map screenshots
                    //
                    //  StringBuilder macro2 = new StringBuilder();
                    //  macro2.AppendLine(@"FRAME NAME=workspace");
                    //  macro2.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=ID:Map<SP>Results*");
                    //  macro2.AppendLine(@"WAIT SECONDS=1");
                    //  macro2.AppendLine(@"SAVEAS TYPE=PNG FOLDER=" + search_address_textbox.Text.Replace(" ", "<SP>") + " FILE=" + "map.jpg");
                    //  macro2.AppendLine(@"WAIT SECONDS=1");
                    // macro2.AppendLine(@"FRAME NAME=subheader");
                    // macro2.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:header ATTR=ID:mapon");
                    // macro2.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:header ATTR=ID:report_type CONTENT=%1linegrid");
                    //macro2.AppendLine(@"WAIT SECONDS=1");
                    // macro2.AppendLine(@"SAVEAS TYPE=PNG FOLDER=" + search_address_textbox.Text.Replace(" ", "<SP>") + " FILE=" + "cma.jpg");

                    string macroCode2 = macro.ToString();
                    status = iim.iimPlayCode(macroCode2, 60);
                    #endregion
                }
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
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10960 CONTENT=" + subjectNeighborhood.maxListPrice.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10962 CONTENT=" + subjectNeighborhood.numberSoldListings.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10964 CONTENT=" + subjectNeighborhood.minSalePrice.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10966 CONTENT=" + subjectNeighborhood.maxSalePrice.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10968 CONTENT=" + subjectNeighborhood.numberOfShortSaleListings.ToString());
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10970 CONTENT=" + subjectNeighborhood.medianSalePrice.ToString());

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
            string distance = "0";
           // string pattern = "(\\d+.\\d*) [mf]";
            string pattern = @"(\d+.\d*) [m]i[^n]";
            Match match = Regex.Match(responseFromServer, pattern);
            distance = match.Groups[1].Value;
            if (responseFromServer.Contains(" ft</text>"))
            {
                distance = "0.05";
            }

            if (String.IsNullOrWhiteSpace(distance))
            {
                distance = "0";
            }
            //MessageBox.Show(distance);
            //Console.WriteLine(responseFromServer);
            // Cleanup the streams and the response.
            reader.Close();
            dataStream.Close();
            response.Close();

            return distance;

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
            iim2.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");

           
            string iimCurrentUrl = iim.iimGetLastExtract();
            string iim2CurrentUrl = iim2.iimGetLastExtract();
            bool connected = false;

            if (iimCurrentUrl.ToLower().Contains("connectmls"))
            {
                mred = iim;
                connected = true;

            }
            else if (iim2CurrentUrl.ToLower().Contains("connectmls"))
            {
                mred = iim2;
                connected = true;
            }
            else
            {
                connected = false;
            }

            if (!connected)
            {
                if (iimCurrentUrl.ToLower().Contains("blank"))
                {
                    iim.iimPlay(@"MRED\#start_mls_55426.iim");
                    mred = iim;
                    connected = true;
                }
                else if (iim2CurrentUrl.ToLower().Contains("blank"))
                {
                    iim2.iimPlay(@"MRED\#start_mls_55426.iim");
                    mred = iim2;
                    connected = true;
                }
                else
                {
                    MessageBox.Show("No open or connected MLS.");
                    return;
                }
            }
                 
           
            //else
            //{
            //    iim2.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            //    iim2CurrentUrl = iim2.iimGetLastExtract();
            //    if (iim2CurrentUrl.ToLower().Contains("connectmls"))
            //    {
            //        mred = iim2;
            //    }
            //    else
            //    {
            //        //neither are on mred, now what??
            //        //is one on the imacro start page?
            //        if (iimCurrentUrl.Contains(@"iOpus/iMacros/Start.html"))
            //        {
            //            mred = iim;
            //            iim.iimPlay(@"MRED\#start_mls.iim"); 
            //        }
            //        else if (iim2CurrentUrl.Contains(@"iOpus/iMacros/Start.html"))
            //        {
            //            mred = iim2;
            //        }
            //        else
            //        {
            //            MessageBox.Show("No open or connected MLS.");
            //            return;
            //        }
            //    }
            //}
             
            switch (mredCmdComboBox.Text)
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
                    if (!PerserveNeighorhoodData)
                    {
                        ndAvgDomTextBox.Text = SubjectNeighborhood.avgDom.ToString();

                        ndNewestHomeTextBox.Text = SubjectNeighborhood.newestHome.ToString();
                        ndMedianAgeTextBox.Text = SubjectNeighborhood.medianAge.ToString();
                        ndOldestHomeTextBox.Text = SubjectNeighborhood.oldestHome.ToString();

                        ndMinListPriceTextBox.Text = SubjectNeighborhood.minListPrice.ToString();
                        ndMedianListPriceTextBox.Text = SubjectNeighborhood.medianListPrice.ToString();
                        ndMaxListPriceTextBox.Text = SubjectNeighborhood.maxListPrice.ToString();

                        ndMinSalePriceTextBox.Text = SubjectNeighborhood.minSalePrice.ToString();
                        ndMedianSalePriceTextBox.Text = subjectNeighborhood.medianSalePrice.ToString();
                        ndMaxSalePriceTextBox.Text = SubjectNeighborhood.maxSalePrice.ToString();
                        
                        ndNumberOfActiveListingTextBox.Text = SubjectNeighborhood.numberActiveListings.ToString();
                        ndNumberActiveReoListingsTextBox.Text = SubjectNeighborhood.numberREOListings.ToString();
                        ndNumberActiveShortListingsTextBox.Text = SubjectNeighborhood.numberOfShortSaleListings.ToString();
                        
                        ndNumberOfSoldListingTextBox.Text = SubjectNeighborhood.numberSoldListings.ToString();
                        ndNumberSoldReoListingsTextBox.Text = SubjectNeighborhood.numberREOSales.ToString();
                        ndNumberSoldShortListingsTextBox.Text = SubjectNeighborhood.numberShortSales.ToString();
                    }

                    if (!PerserveCompSetData)
                    {
                        cdAvgDomTextBox.Text = SetOfComps.avgDom.ToString();

                        cdNewestHomeTextBox.Text = SetOfComps.newestHome.ToString();
                        cdMedianAgeTextBox.Text = SetOfComps.medianAge.ToString();
                        cdOldestHomeTextBox.Text = SetOfComps.oldestHome.ToString();

                        cdMinListPriceTextBox.Text = SetOfComps.minListPrice.ToString();
                        cdMedianListPriceTextBox.Text = SetOfComps.medianListPrice.ToString();
                        cdMaxListPriceTextBox.Text = SetOfComps.maxListPrice.ToString();

                        cdMinSalePriceTextBox.Text = SetOfComps.minSalePrice.ToString();
                        cdMedianSalePriceTextBox.Text = SetOfComps.medianSalePrice.ToString();
                        cdMaxSalePriceTextBox.Text = SetOfComps.maxSalePrice.ToString();

                        cdNumberOfActiveListingTextBox.Text = SetOfComps.numberActiveListings.ToString();
                        cdNumberActiveReoListingsTextBox.Text = SetOfComps.numberREOListings.ToString();
                        cdNumberActiveShortListingsTextBox.Text = SetOfComps.numberOfShortSaleListings.ToString();

                        cdNumberOfSoldListingTextBox.Text = SetOfComps.numberSoldListings.ToString();
                        cdNumberSoldReoListingsTextBox.Text = SetOfComps.numberREOSales.ToString();
                        cdNumberSoldShortListingsTextBox.Text = SetOfComps.numberShortSales.ToString();
                    }

                   

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
                this.quickSaleTextbox.Text = (Math.Round((double)(((Convert.ToInt64(valueTextbox.Text) * Convert.ToDecimal(quickSalePercentageTextBox.Text)))))).ToString();
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

        private void button13_Click(object sender, EventArgs e)
        {
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"");
            macro.AppendLine(@"FRAME NAME=subheader");
            macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:header ATTR=SRC:*favorites_small.gif");
       //     macro.AppendLine(@"'New tab opened");
            macro.AppendLine(@"TAB T=2");
            macro.AppendLine(@"FRAME F=0");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:dc ATTR=NAME:selHeading CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:dc ATTR=NAME:enterNewHeading CONTENT=" + SubjectFullAddress.Replace(" ", "<SP>") + "-" + DateTime.Now.ToShortDateString());
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:dc ATTR=NAME:<SP>Add<SP>");

            string macroCode = macro.ToString();
            status = iim.iimPlayCode(macroCode, 60); 
        }

        private void subjectRentTextbox_TextChanged(object sender, EventArgs e)
        {
         //    http://www.zillow.com/webservice/GetSearchResults.htm?zws-id=<ZWSID>&address=2114+Bigelow+Ave&citystatezip=Seattle%2C+WA

            string[] address = subjectFullAddressTextbox.Text.Replace(@" ", " ").Split(',');

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
            IEnumerable<System.Windows.Forms.TextBox> query3 = this.compDataGroupBox.Controls.OfType<System.Windows.Forms.TextBox>();
            IEnumerable<System.Windows.Forms.ComboBox> query4 = this.groupBox1.Controls.OfType<System.Windows.Forms.ComboBox>();
   

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
                foreach (System.Windows.Forms.TextBox t in query3)
                {
                    file.WriteLine("{0};{1}", t.Name, t.Text);
                }
                foreach (System.Windows.Forms.ComboBox t in query4)
                {
                    file.WriteLine("{0};{1}", t.Name, t.Text);
                }

      
                
               file.WriteLine("{0};{1}", subjectDetachedradioButton.Name, subjectDetachedradioButton.Checked.ToString());
               file.WriteLine("{0};{1}", subjectAttachedRadioButton.Name, subjectAttachedRadioButton.Checked.ToString());
               //file.WriteLine("{0};{1}", subjectMlsTypecomboBox.Name, subjectMlsTypecomboBox.Text);
               file.WriteLine("{0};{1}", dateTimePickerInspectionDate.Name, dateTimePickerInspectionDate.Value.ToUniversalTime());
               file.WriteLine("{0};{1}", dateTimePickerSubjectCurrentListDate.Name, dateTimePickerSubjectCurrentListDate.Value.ToUniversalTime());
               //file.WriteLine("{0};{1}", dateTimePickerSubjectCurrentListDate.Name, dateTimePickerSubjectCurrentListDate.Value.ToUniversalTime());
               //dateTimePickerInspectionDate
               file.WriteLine("{0};{1}", comboBoxBpoType.Name, comboBoxBpoType.Text);


               richTextBoxNeighborhoodComments.SaveFile(this.SubjectFilePath + "\\" + "neighborhood-comments.rtf");
              
               bpoCommentsTextBox.SaveFile(this.SubjectFilePath + "\\" + "bpocomments.rtf");
            }
             using (System.IO.StreamWriter file = new System.IO.StreamWriter(".config"))
             {
                 file.WriteLine("{0};{1}", "lastOpenedSubject", search_address_textbox.Text);
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
                openTab.AppendLine(@"TAG POS=2 TYPE=DIV ATTR=onclick:show*");

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

            #region southwest
            if (currentUrl.Contains("rapidclose"))
            {
                StringBuilder macro11 = new StringBuilder();
                StringBuilder macro = new StringBuilder();
                List<string> records = new List<string>();
                DateTime dueDate = new DateTime();

                dueDate = DateTime.Now.AddDays(2);

               

                

                macro11.AppendLine(@"TAG POS=2 TYPE=TABLE ATTR=CLASS:tblvieworder EXTRACT=TXT");

                iMacros.Status s = iim2.iimPlayCode(macro11.ToString());
                string extractedTable = iim2.iimGetLastExtract();
                string header = iim2.iimGetExtract(0);

                string[] sep = { "#NEWLINE#" };
                string[] theTable = extractedTable.Split(sep, StringSplitOptions.RemoveEmptyEntries);
                //if (records.Count == 0)
                //{
                //    records.Add(theTable[0].Replace("#NEXT#", ",").Substring(4).TrimEnd(','));  //add the header
                //}
                string[] sep2 = { "#NEXT#" };

                string[] theRecord = { };
                string line = "";

                try
                {

                    using (StreamReader r = File.OpenText("sw-orders.txt"))
                    {
                        while ((line = r.ReadLine()) != null)
                        {
                            records.Add(line);
                        }
                    }
                }
                catch
                {

                }

                string address;
          
                string city;

                for (int i = 0; i < theTable.GetLength(0); i++)
                {

                    theRecord = theTable[i].Split(sep2, StringSplitOptions.RemoveEmptyEntries);

                    if (theRecord.Length > 5 && Regex.IsMatch(theRecord[0], @"\d\d\d\d\d\d\d\d"))
                    {
                        if (!records.Contains(theRecord[0]))
                        {
                            records.Add(theRecord[0]);
                            city = Regex.Match(theRecord[3], (@"\r\n(.*),")).Groups[1].Value;
                            address = Regex.Match(theRecord[3], (@"(.*)\r\n")).Groups[1].Value;

                            macro.AppendLine(@"URL GOTO=https://docs.google.com/forms/d/12BCMedwyV_CpkdwwG10wQvl8mPHlEq0Mbig8PAsb46g/viewform?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ#gid=20");

                           // macro.AppendLine(@"URL GOTO=https://docs.google.com/forms/d/12BCMedwyV_CpkdwwG10wQvl8mPHlEq0Mbig8PAsb46g/viewform?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ#gid=20");macro.AppendLine(@"TAG POS=1 TYPE=DIV FORM=ACTION:https://docs.google.com/forms/d/12BCMedwyV_CpkdwwG10wQvl8mPHlEq0Mbig8PAsb46g/formResponse ATTR=CLASS:docssharedWizSelectPaperselectDropDown");
                     //       macro.AppendLine(@"TAG POS=2 TYPE=DIV FORM=ACTION:https://docs.google.com/forms/d/12BCMedwyV_CpkdwwG10wQvl8mPHlEq0Mbig8PAsb46g/formResponse ATTR=TXT:PCR");
                           //address
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/forms/d* ATTR=NAME:bpo<SP>tracking<SP>form CONTENT=");
                            macro.AppendLine(@"DS cmd=CLICK X={{!TAGX}} Y={{!TAGY}}");
                            macro.AppendLine(@"DS cmd=KEY CONTENT=" + address.Replace(" ", "<SP>"));
                            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/forms/d* ATTR=NAME:bpo<SP>tracking<SP>form CONTENT=" + address.Replace(" ", "<SP>"));
                            
                            //city
                            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/forms/d* ATTR=NAME:bpo<SP>tracking<SP>form CONTENT=");
                            macro.AppendLine(@"DS cmd=CLICK X={{!TAGX}} Y={{!TAGY}}");
                            macro.AppendLine(@"DS cmd=KEY CONTENT=" + city.Replace(" ", "<SP>"));

                          

                            //ordernumber
                            macro.AppendLine(@"TAG POS=3 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/forms/d* ATTR=NAME:bpo<SP>tracking<SP>form CONTENT=");
                            macro.AppendLine(@"DS cmd=CLICK X={{!TAGX}} Y={{!TAGY}}");
                            macro.AppendLine(@"DS cmd=KEY CONTENT=" + theRecord[0]);

                            //duedate
                            //macro.AppendLine(@"TAG POS=4 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/forms/d* ATTR=CLASS:quantumWizTextinputPaperinputInput<SP>exportInput CONTENT=");
                           // macro.AppendLine(@"TAG POS=4 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/forms/d/12BCMedwyV_CpkdwwG10wQvl8mPHlEq0Mbig8PAsb46g/formResponse ATTR=CLASS:quantumWizTextinputPaperinputInput<SP>exportInput");
                            macro.AppendLine(@"TAG POS=5 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/forms/d* ATTR=CLASS:quantumWizTextinputPaperinputInput<SP>exportInput CONTENT=");
                            macro.AppendLine(@"DS cmd=CLICK X={{!TAGX}} Y={{!TAGY}}");
                            macro.AppendLine(@"DS cmd=KEY CONTENT=" + dueDate.Month.ToString() + "{Enter}");
                            macro.AppendLine(@"WAIT SECONDS=5");

                            macro.AppendLine(@"TAG POS=5 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/forms/d* ATTR=CLASS:quantumWizTextinputPaperinputInput<SP>exportInput CONTENT=");
                            macro.AppendLine(@"DS cmd=CLICK X={{!TAGX}} Y={{!TAGY}}");
                            macro.AppendLine(@"DS cmd=KEY CONTENT=" + dueDate.Day.ToString() );
                            macro.AppendLine(@"WAIT SECONDS=1");

                            //macro.AppendLine(@"TAG POS=6 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/forms/d/12BCMedwyV_CpkdwwG10wQvl8mPHlEq0Mbig8PAsb46g/formResponse ATTR=CLASS:quantumWizTextinputPaperinputInput<SP>exportInput CONTENT=");
                            //macro.AppendLine(@"DS cmd=CLICK X={{!TAGX}} Y={{!TAGY}}");
                            //macro.AppendLine(@"DS cmd=KEY CONTENT=" + dueDate.ToShortDateString());
                            //macro.AppendLine(@"WAIT SECONDS=1");
                     
                            //billed amount
                            macro.AppendLine(@"TAG POS=4 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/forms/d/12BCMedwyV_CpkdwwG10wQvl8mPHlEq0Mbig8PAsb46g/formResponse ATTR=NAME:bpo<SP>tracking<SP>form CONTENT=");
                            macro.AppendLine(@"DS cmd=CLICK X={{!TAGX}} Y={{!TAGY}}");
                            macro.AppendLine(@"DS cmd=KEY CONTENT=25");


                         //   macro.AppendLine(@"TAG POS=5 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/forms/d/12BCMedwyV_CpkdwwG10wQvl8mPHlEq0Mbig8PAsb46g/formResponse ATTR=NAME:bpo<SP>tracking<SP>form");
                       //     macro.AppendLine(@"TAG POS=1 TYPE=CONTENT FORM=ACTION:https://docs.google.com/forms/d/12BCMedwyV_CpkdwwG10wQvl8mPHlEq0Mbig8PAsb46g/formResponse ATTR=TXT:Southwest");
                       
                           //// macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:https://docs.google.com/forms/d/* ATTR=NAME:submit");
                            string macroCode = macro.ToString();
                            iim.iimPlayCode(macroCode, 60);
                            macro.Clear();
                        }
                    }
                    
                   


                }


                using (System.IO.StreamWriter file = File.AppendText(("sw-orders.txt")))
                {


                    foreach (string r in records)
                    {
                        file.WriteLine(r);
                    }
                    //foreach (System.Windows.Forms.TextBox t in query2)
                    //{
                    //    file.WriteLine("{0};{1}", t.Name, t.Text);
                    //}

                    //file.WriteLine("{0};{1}", subjectDetachedradioButton.Name, subjectDetachedradioButton.Checked.ToString());
                    //file.WriteLine("{0};{1}", subjectAttachedRadioButton.Name, subjectAttachedRadioButton.Checked.ToString());
                    //file.WriteLine("{0};{1}", subjectMlsTypecomboBox.Name, subjectMlsTypecomboBox.Text);
                    //bpoCommentsTextBox.SaveFile(this.SubjectFilePath + "\\" + "bpocomments.rtf");

                }


                //var ttt = Regex.Matches(htmlCode, @"orderNo=(\d\d\d\d\d\d)");

                //     MessageBox.Show(ttt.Count.ToString());


                //HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                //doc.LoadHtml(htmlCode);

                //var x = doc.DocumentNode.SelectSingleNode("//*[@id=\"appraisalOrderList\"]/tbody/tr[1]/td[1]/a/font");
                //int orderNumberIndex = 1;
                //int addressIndex = 4;
                //int cityIndex = 5;

                ////
                ////TBD:  Calculate correct due date by extracting date + standard 3 days.
                ////
                //for (int i = 1; i < 80 ; i++)
                //{
                //    macro.AppendLine(@"VERSION BUILD=10022823");
                //    macro.AppendLine(@"TAB T=1");
                //    macro.AppendLine(@"TAB CLOSEALLOTHERS");
                //    macro.AppendLine(@"URL GOTO=https://docs.google.com/spreadsheet/viewform?usp=drive_web&formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ#gid=20");
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.2.single CONTENT=" + doc.DocumentNode.SelectSingleNode("//*[@id=\"appraisalOrderList\"]/tbody/tr[" + i.ToString() + "]/td[" + addressIndex.ToString() + "]").InnerText.Replace(" ", "<SP>"));
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.3.single CONTENT=" + doc.DocumentNode.SelectSingleNode("//*[@id=\"appraisalOrderList\"]/tbody/tr[" + i.ToString() + "]/td[" + cityIndex.ToString() + "]").InnerText.Replace(" ", "<SP>"));
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.14.single CONTENT=" + doc.DocumentNode.SelectSingleNode("//*[@id=\"appraisalOrderList\"]/tbody/tr[" + i.ToString() + "]/td[" + orderNumberIndex.ToString() + "]/a/font").InnerText);
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.9.single CONTENT=" + DateTime.Now.ToShortDateString());
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.11.single CONTENT=50");
                //    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.1.single CONTENT=%Solutionstar");
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:submit");
                //    string macroCode = macro.ToString();
                //    iim.iimPlayCode(macroCode, 60);
                //    macro.Clear();
                //}
            }
            #endregion

            #region equitrax
            if (currentUrl.Contains("equi-trax"))
            {
                 StringBuilder macro11 = new StringBuilder();
                StringBuilder macro = new StringBuilder();
                List<string> records = new List<string>();

                macro11.AppendLine(@"FRAME NAME=main");
                macro11.AppendLine(@"TAG POS=1 TYPE=TABLE ATTR=ID:NewBPOList_tbl EXTRACT=TXT");

                iMacros.Status s = iim2.iimPlayCode(macro11.ToString());
                string extractedTable = iim2.iimGetLastExtract();
                string header = iim2.iimGetExtract(0);

               string[] sep = { "#NEWLINE#" };
                string[] theTable = extractedTable.Split(sep, StringSplitOptions.RemoveEmptyEntries);
                //if (records.Count == 0)
                //{
                //    records.Add(theTable[0].Replace("#NEXT#", ",").Substring(4).TrimEnd(','));  //add the header
                //}
               string[] sep2 = { "#NEXT#" };

               string[] theRecord = { };
               string line = "";

               using (StreamReader r = File.OpenText("et-orders.txt"))
               {
                   while ((line = r.ReadLine()) != null)
                   {
                       records.Add(line);
                   }
               }
               string address;
                string city;
                string orderNum = "";
                string ttt = "";
                for (int i = 0; i < theTable.GetLength(0); i++)
                {

                    theRecord = theTable[i].Split(sep2, StringSplitOptions.RemoveEmptyEntries);
                    //ttt = theTable[i].Replace("#NEXT#", ",").Replace("\r\n"," ");
                   // orderNum = 
                    if (!records.Contains(theRecord[2]))
                    {
                        records.Add(theRecord[2]);
                        city = Regex.Match(theRecord[4], (@"\r\n(.*),")).Groups[1].Value;
                        address = Regex.Match(theRecord[4], (@"(.*)\r\n")).Groups[1].Value;
                        macro.AppendLine(@"VERSION BUILD=10022823");
                        macro.AppendLine(@"TAB T=1");
                        macro.AppendLine(@"TAB CLOSEALLOTHERS");
                        macro.AppendLine(@"URL GOTO=https://docs.google.com/spreadsheet/viewform?usp=drive_web&formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ#gid=20");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:https://docs.google.com/forms/d/* ATTR=NAME:entry.2710153 CONTENT=%Exterior");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/forms/d/* ATTR=NAME:entry.1000002 CONTENT=" + address.Replace(" ", "<SP>"));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/forms/d/* ATTR=NAME:entry.1000003 CONTENT=" + city.Replace(" ", "<SP>"));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/forms/d/* ATTR=NAME:entry.1000014 CONTENT=" + theRecord[2]);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/forms/d/* ATTR=NAME:entry.1000017 CONTENT=" + DateTime.Now.AddDays(2).ToShortDateString());
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/forms/d/* ATTR=NAME:entry.1000011 CONTENT=" + theRecord[8]);
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:https://docs.google.com/forms/d/* ATTR=NAME:entry.1000009 CONTENT=%SWBC");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:https://docs.google.com/forms/d/* ATTR=NAME:submit");
                        string macroCode = macro.ToString();
                        iim.iimPlayCode(macroCode, 60);
                        macro.Clear();
                    }

                   
                }


                using (System.IO.StreamWriter file = File.AppendText(("et-orders.txt")))
                {


                    foreach (string r in records)
                    {
                        file.WriteLine(r);
                    }
                    //foreach (System.Windows.Forms.TextBox t in query2)
                    //{
                    //    file.WriteLine("{0};{1}", t.Name, t.Text);
                    //}

                    //file.WriteLine("{0};{1}", subjectDetachedradioButton.Name, subjectDetachedradioButton.Checked.ToString());
                    //file.WriteLine("{0};{1}", subjectAttachedRadioButton.Name, subjectAttachedRadioButton.Checked.ToString());
                    //file.WriteLine("{0};{1}", subjectMlsTypecomboBox.Name, subjectMlsTypecomboBox.Text);
                    //bpoCommentsTextBox.SaveFile(this.SubjectFilePath + "\\" + "bpocomments.rtf");

                }


                //var ttt = Regex.Matches(htmlCode, @"orderNo=(\d\d\d\d\d\d)");

           //     MessageBox.Show(ttt.Count.ToString());


                //HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                //doc.LoadHtml(htmlCode);

                //var x = doc.DocumentNode.SelectSingleNode("//*[@id=\"appraisalOrderList\"]/tbody/tr[1]/td[1]/a/font");
                //int orderNumberIndex = 1;
                //int addressIndex = 4;
                //int cityIndex = 5;

                ////
                ////TBD:  Calculate correct due date by extracting date + standard 3 days.
                ////
                //for (int i = 1; i < 80 ; i++)
                //{
                //    macro.AppendLine(@"VERSION BUILD=10022823");
                //    macro.AppendLine(@"TAB T=1");
                //    macro.AppendLine(@"TAB CLOSEALLOTHERS");
                //    macro.AppendLine(@"URL GOTO=https://docs.google.com/spreadsheet/viewform?usp=drive_web&formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ#gid=20");
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.2.single CONTENT=" + doc.DocumentNode.SelectSingleNode("//*[@id=\"appraisalOrderList\"]/tbody/tr[" + i.ToString() + "]/td[" + addressIndex.ToString() + "]").InnerText.Replace(" ", "<SP>"));
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.3.single CONTENT=" + doc.DocumentNode.SelectSingleNode("//*[@id=\"appraisalOrderList\"]/tbody/tr[" + i.ToString() + "]/td[" + cityIndex.ToString() + "]").InnerText.Replace(" ", "<SP>"));
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.14.single CONTENT=" + doc.DocumentNode.SelectSingleNode("//*[@id=\"appraisalOrderList\"]/tbody/tr[" + i.ToString() + "]/td[" + orderNumberIndex.ToString() + "]/a/font").InnerText);
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.9.single CONTENT=" + DateTime.Now.ToShortDateString());
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.11.single CONTENT=50");
                //    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:entry.1.single CONTENT=%Solutionstar");
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:https://docs.google.com/spreadsheet/formResponse?formkey=dEd4ZVJiWVdKRVk4SWZtN1lDOENCQkE6MQ&ifq ATTR=NAME:submit");
                //    string macroCode = macro.ToString();
                //    iim.iimPlayCode(macroCode, 60);
                //    macro.Clear();
                //}
            }
            #endregion

            #region solutionstar
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
            #endregion
        }

        private void button25_Click(object sender, EventArgs e)
        {
            GenerateVersions(SubjectFilePath);
        }

        private void subjectSubdivisionTextbox_TextChanged(object sender, EventArgs e)
        {

        }

        private void button26_Click(object sender, EventArgs e)
        {
            //iMacros.App mred = new iMacros.App();

            //extract data as a REAL table

            List<string> records = new List<string>();
            Status s = new Status();
            s = Status.sOk;

            while (s == Status.sOk)
            {
                StringBuilder macro = new StringBuilder();
                macro.AppendLine(@"VERSION BUILD=10.2.26.4235");
                macro.AppendLine(@"SET !TIMEOUT_STEP 30");
                
                //macro.AppendLine(@"TAB T=1");
                //macro.AppendLine(@"TAB CLOSEALLOTHERS");
                //macro.AppendLine(@"URL GOTO=http://connectmls*/mls.jsp?module=search&encurl=search/search_index.jsp?uri=search/search.jsp&switch_type=OFFICE");
                macro.AppendLine(@"FRAME NAME=workspace");
                macro.AppendLine(@"TAG POS=1 TYPE=TABLE ATTR=CLASS:gridview EXTRACT=TXT");
                string macroCode = macro.ToString();
                iim.iimPlayCode(macroCode);
                string extractedTable = iim.iimGetLastExtract();
                string[] sep = { "#NEWLINE#" };
                string[] theTable = extractedTable.Split(sep, StringSplitOptions.RemoveEmptyEntries);
                if (records.Count == 0)
                {
                    records.Add(theTable[0].Replace("#NEXT#", ",").Substring(4).TrimEnd(','));  //add the header
                }
               



                //click on each hidden email link in the table
                //which opens a new table.
                //extract all txt on the card then
                //find and store the actual email address text

                for (int i = 1; i < theTable.GetLength(0); i++)
                {

                    macro.Clear();
                    macro.AppendLine(@"FRAME NAME=workspace");

                    macro.AppendLine(@"TAG POS=" + i.ToString() + @" TYPE=IMG FORM=NAME:dc ATTR=SRC:http://connectmls*/images/newMsg.gif");
                    macro.AppendLine(@"'New tab opened");
                    macro.AppendLine(@"TAB T=2");
                    macro.AppendLine(@"FRAME F=0");
                    macro.AppendLine(@"TAG POS=4 TYPE=TR ATTR=* EXTRACT=TXT"); //header
                    macro.AppendLine(@"TAG POS=4 TYPE=TABLE ATTR=* EXTRACT=TXT");  //whole card

                    macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=HREF:mailto:* EXTRACT=TXT"); //email
                    macro.AppendLine(@"TAB CLOSE");
                    macroCode = macro.ToString();
                    status = iim.iimPlayCode(macroCode);

                    //frm immediate window
                    //iim.iimGetExtract(0);
                    //"Thomas Killoren[EXTRACT]tkilloren@aol.com[EXTRACT]"
                    // iim.iimGetExtract(1);
                    //"Thomas Killoren"
                    // iim.iimGetExtract(2);
                    //"tkilloren@aol.com"


                    string mredBusinessCard = iim.iimGetExtract(2);
                    string emailAddressText = iim.iimGetExtract(3);
                    if (emailAddressText.Contains(";"))
                    {
                        emailAddressText = emailAddressText.Split(';')[0];
                    }

                    if (emailAddressText.Contains("#"))
                    {
                        emailAddressText = "";
                    }

                   // MessageBox.Show(emailAddressText);

                   // records.Add(theTable[i].Replace("#NEXT#", ",").Substring(5).TrimEnd(',') + "," + emailAddressText);
                   
                  //  records.Add(theTable[i].Replace("#NEXT#", ",").Substring(5).TrimEnd(',').Insert(theTable[i].IndexOf(",",0,3),emailAddressText));

                    //string tempStr = theTable[i].Replace("#NEXT#", ",").Substring(5).Trim(',');
                    string[] sep2 = { "#NEXT#" };

                    string[] fields = theTable[i].Split(sep2,StringSplitOptions.None);
                    string finalStr = "";
                 
                    int fieldIndex = 0;
                    foreach (string str in fields)
                    {

                        if (fieldIndex == 6)
                        {
                            finalStr = finalStr + emailAddressText + ",";
                        
                        }
                        else if (fieldIndex == 8)
                        {
                            finalStr = finalStr + "\"" + str.Trim() + "\",";
                        }
                        else if (fieldIndex == 0 || fieldIndex == 1 || fieldIndex == 14)
                        {

                        }
                        else
                        {
                            finalStr = finalStr + str.Trim() + ",";
                        }

                        fieldIndex++;
                    }
                   
                    records.Add(finalStr.Trim(',').Replace("[EXTRACT]", ""));
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"F:\Dropbox\Dev\WCR Email CRM project\master.csv",true))
                    {
                        file.WriteLine(finalStr.Trim(',').Replace("[EXTRACT]", ""));

                    //    //foreach (object o in records)
                    //    //{
                    //    //    file.WriteLine(o.ToString());
                    //    //}
                    }
                   // records.Add(theTable[i].Replace("#NEXT#", ",").Substring(5).TrimEnd(',').Insert(x, emailAddressText));
                   
                }

                macro.Clear();
                macro.AppendLine(@"FRAME NAME=navpanel");
                macro.AppendLine(@"TAG POS=2 TYPE=DIV ATTR=onclick:show*");
                macroCode = macro.ToString();
                 s = iim.iimPlayCode(macroCode);

                 //using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"F:\Dropbox\Dev\WCR Email CRM project\master.csv",true))
                 //{
                 //    file.WriteLine(o.ToString());

                 //    //foreach (object o in records)
                 //    //{
                 //    //    file.WriteLine(o.ToString());
                 //    //}
                 //}

                //MessageBox.Show(extractedTable);
            }

            //using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"F:\Dropbox\Dev\WCR Email CRM project\master.csv"))
            //{
            //    #region writefile

            //    foreach (object o in records)
            //    {
            //        file.WriteLine(o.ToString());
            //    }

                //try
                //{
                //    file.WriteLine("#closed: " + closed_comps.ToString() + " #active: " + active_comps.ToString() + " #pending: " + pending_comps.ToString());
                //    file.WriteLine("Average/Mean $/Above GLA, sale: " + Decimal.Round(pricePerSfList.Average(), 2).ToString() + " Median $/Above GLA, sale: " + Decimal.Round(pricePerSfList.Median(), 2).ToString());
                //    file.WriteLine("Min Sale: {0}, Max Sale: {1}, Average Sale: {2}, Median Sale: {3}", soldPriceList.Min(), soldPriceList.Max(), Convert.ToInt32(soldPriceList.Average()), soldPriceList.Median());
                //    file.WriteLine("Min List: {0}, Max List: {1}, Average List: {2}, Median List: {3}", activePriceList.Min(), activePriceList.Max(), Convert.ToInt32(activePriceList.Average()), activePriceList.Median());
                //    file.WriteLine("Min dom: {0}, Max dom: {1}, Average dom: {2}, Median dom: {3}", domList.Min(), domList.Max(), Convert.ToInt32(domList.Average()), domList.Median());
                //    file.WriteLine("Min Age: {0}, Max Age: {1}, Average Age: {2}, Median Age: {3}", ageList.Min(), ageList.Max(), Convert.ToInt32(ageList.Average()), ageList.Median());
                //    file.WriteLine("REO Sold: {0}, REO Active: {1}, Short Sold: {2}, Short Active: {3}", reoSales, reoActive, shortSales, shortActive);
                //    file.WriteLine("# Same Type: {0}, # Same Subdivision {1}, # Comparable Age: {2},# Comparable aGLA: {3}, Comparable LotSize: {4}, ", numberSameTypeAsSubject.ToString(), numSameSubdivision.ToString(), numberComparableAge.ToString(), numberComparableGla.ToString(), numberComparableLot.ToString());
                //}
                //catch
                //{

                //}
                //picDiffList
                // file.WriteLine("Min {0}, Max {1}, Average {2}, Median {3}", picDiffList.Min(), picDiffList.Max(), picDiffList.Average(), picDiffList.Median());
                //foreach (MLSListing m in listings)
                //{
                //    file.WriteLine("MLS: {0} is {1} from subject.", m.mlsHtmlFields["mlsNumber"].value, m.proximityToSubject);
                //}


               

                //foreach (object o in subdivisionList)
                //{
                //    file.WriteLine(o.ToString());
                //}



                //foreach (object o in listingAgentList)
                //{
                //    file.WriteLine(o.ToString());
                //}
                //foreach (MLSListing m in listings)
                //{
                //    file.WriteLine(m.mlsHtmlFields["remarks"].value);
                //}
                //#endregion
            //}
        }

        private void label47_Click(object sender, EventArgs e)
        {

        }

        private void button27_Click(object sender, EventArgs e)
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

           // MessageBox.Show(currentUrl);

            #region equitrax
            if (currentUrl.ToLower().Contains("equi-trax"))
            {
                Equitrax bpoform = new Equitrax();
                streetnumTextBox.Text = "equitrax";
                bpoform.QA(iim2, this);
            }
            #endregion  
        }

        private void button28_Click(object sender, EventArgs e)
        {
            mredCmdComboBox.SelectedItem = "New Map Search";
            this.runMredScriptButton_Click(this, new EventArgs());
            mredCmdComboBox.SelectedItem = "FindComps - Current Search";
            this.runMredScriptButton_Click(this, new EventArgs());
            this.order_prefill_button_Click(this, new EventArgs());
            this.button13_Click(this, new EventArgs());
        }

        private void button29_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> versions = new Dictionary<string, string>();
            //Define the versions to generate and their filename suffixes.
            //versions.Add("_thumb", "width=100&height=100&crop=auto&format=jpg"); //Crop to square thumbnail
            //versions.Add("_small", "maxwidth=200&maxheight=200format=jpg"); //Fit inside 400x400 area, jpeg
            versions.Add("_upload", "maxwidth=640&maxheight=480format=jpg"); //Fit inside 400x400 area, jpeg
            // versions.Add("_large", "maxwidth=1900&maxheight=1900&format=jpg"); //Fit inside 1900x1200 area


            //dateTimePickerInspectionDate
            string dateStamp = dateTimePickerInspectionDate.Value.ToShortDateString().Replace("/", "-");

            //string basePath = ImageResizer.Util.PathUtils.RemoveExtension(original);
            string basePath = SubjectFilePath;
            string writeOutDir =  "stamped_" + dateStamp;
            Directory.SetCurrentDirectory(basePath);
            Directory.CreateDirectory(writeOutDir);
           Directory.SetCurrentDirectory(basePath + @"\" + writeOutDir);
           //Directory.CreateDirectory("StampedOnly");
           Directory.CreateDirectory("ShrunkAndStamped");


            //To store the list of generated paths
            List<string> generatedFiles = new List<string>();
           // System.Drawing.Image img = System.Drawing.Image.FromFile("Brush Tail Possum.jpg");
           // System.Drawing.Image img = System.Drawing.Image.FromFile("Brush Tail Possum.jpg");
           // System.Drawing.Image imgOverlay = System.Drawing.Image.FromFile("overlay.png");
           

           
            System.Drawing.Color color = System.Drawing.Color.FromArgb(128, 255, 31, 31);
            System.Drawing.Color black = System.Drawing.Color.FromArgb(128, 0, 0, 0);

            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;
            stringFormat.LineAlignment = StringAlignment.Center;

            int x = 0;
            int y = 0;

            System.Drawing.SolidBrush myBrush = new System.Drawing.SolidBrush(System.Drawing.Color.Black);
            //System.Drawing.Graphics formGraphics = this.CreateGraphics();
           // formGraphics.FillRectangle(myBrush, new System.Drawing.Rectangle(0, 0, 200, 300));


            //gr.SmoothingMode = Graphics.SmoothingMode.AntiAlias;

          //  gr.DrawImage(imgOverlay, new System.Drawing.Point(img.Width - 78, img.Height - 25));
           // gr.DrawString(DateTime.Now.ToShortDateString(), font, new System.Drawing.SolidBrush(color), new System.Drawing.Point(img.Width - 40, img.Height - 15), stringFormat);

           // MemoryStream outputStream = new MemoryStream();
            //img.Save("Brush Tail Possum2.jpg");
            int i = 1;
            foreach (string pic in Directory.GetFiles(basePath, "*.jp*g"))
            {
                string baseFileName = ImageResizer.Util.PathUtils.RemoveExtension(Path.GetFileName(pic));
            
               //add date/time to pic
                System.Drawing.Image img = System.Drawing.Image.FromFile(pic);
                Graphics gr = Graphics.FromImage(img);
              //  gr.DrawString(DateTime.Now.ToShortDateString(), font, new System.Drawing.SolidBrush(color), new System.Drawing.Point(img.Width - 40, img.Height - 15), stringFormat);
                x = img.Width;
                y = img.Height;

                int fontSize = 96;

                if (y <= 480)
                {
                    fontSize = 32;
                }


                
                System.Drawing.Font font = new System.Drawing.Font("Times New Roman", (float)fontSize, System.Drawing.FontStyle.Regular);

                gr.FillRectangle(myBrush, new System.Drawing.Rectangle(0, y-200, 900, 350));
               // gr.DrawString("2017-06-12", font, new System.Drawing.SolidBrush(black), new System.Drawing.Point((img.Width / 2) + 1 , (img.Height / 2) + 1), stringFormat);
              //  gr.DrawString("2017-06-12", font, new System.Drawing.SolidBrush(color), new System.Drawing.Point(img.Width / 2, img.Height / 2), stringFormat);

                gr.DrawString(dateStamp, font, new System.Drawing.SolidBrush(System.Drawing.Color.White), new System.Drawing.Point(350, img.Height - 50), stringFormat);

                MemoryStream outputStream = new MemoryStream();
                img.Save(baseFileName + "_stamped.jpg");
                //Generate each version

                
                foreach (string suffix in versions.Keys)
                { 
                    //Let the image builder add the correct extension based on the output file type
                    generatedFiles.Add(ImageBuilder.Current.Build(baseFileName + "_stamped.jpg", @"ShrunkAndStamped\picture_" + i.ToString() + "_stamped" + suffix,
                     new ResizeSettings(versions[suffix]), false, true));
                    //generatedFiles.Add(ImageBuilder.Current.Build(img.Tag, basePath + suffix,
                   //  new ResizeSettings(versions[suffix]), false, true)); 
                       
                }

                img.Dispose();
                gr.Dispose();
                font.Dispose();
                outputStream.Close();

                i++; 

            }
            

           
        }

        private void button30_Click(object sender, EventArgs e)
        {
            
            switch (comboBoxBpoVendor.Text)
            {
                case "AVM" :
                     iim2.iimPlay(@"AVM\avm-login.iim");
                     break;
                case "Old Republic" :
                     iim2.iimPlay(@"Old Republic\or-login.iim");
                     break;
                default:
                     break;

            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {

            string retsLogin = @"http://connectmls-rets.mredllc.com/rets/server/login";
            // string retsString = @"http://connectmls-rets.mredllc.com/rets/server/search?SearchType=Property&Class=ResidentialProperty&QueryType=DMQL2&Format=COMPACT&StandardNames=1&Select=ListingID,ListPrice&Query=(ListPrice=300000%2B)&Count=1&Limit=10";

            string retsString = @"http://connectmls-rets.mredllc.com/rets/server/search?SearchType=Property&Class=ResidentialProperty&QueryType=DMQL2&Format=COMPACT&StandardNames=1&Select=ListingID,ListPrice&Query=(ParcelNumber=" + SubjectPin + ")";

            WebRequest loginRequest = WebRequest.Create(retsLogin);
            WebRequest queryRequest = WebRequest.Create(retsString);

            //WebRequest request = WebRequest.Create(retsString);
            //   loginRequest.Headers.Add("Authorization: Digest username=\"RETS_O_74601_6\", realm=\"connectmls-rets.mredllc.com\", nonce=\"0d4194327a1c825d2672e5bf273cfdb6\", uri=\"/rets/server/search?SearchType=Property&Class=ResidentialProperty&QueryType=DMQL2&Format=COMPACT&StandardNames=1&Select=ListingID,ListPrice&Query=(ListPrice=300000%2B)&Count=1&Limit=12\", response=\"197d466fe08aa541be16efe37ff24058\", opaque=\"5ccdef346870ab04ddfe0412367fccba\", qop=auth, nc=00000008, cnonce=\"bf43690177da68f5\"");
           // loginRequest.Headers.Add("Cookie:RETS-Session-ID=24C99698CE44123BF98B40E98CB9888F; __utmt=1; __utmb=118572885.1.10.1462996245; JSESSIONID=F94B7B6B18567676CE1C05CE138A8B6A; __utma=118572885.1876298590.1455414260.1460149404.1462996245.5; __utmc=118572885; __utmz=118572885.1462996245.5.4.utmcsr=google|utmccn=(organic)|utmcmd=organic|utmctr=(not%20provided); JSESSIONID=F94B7B6B18567676CE1C05CE138A8B6A; __utma=118572885.1876298590.1455414260.1460149404.1462996245.5; __utmc=118572885; __utmz=118572885.1462996245.5.4.utmcsr=google|utmccn=(organic)|utmcmd=organic|utmctr=(not%20provided); SERVERID=ws17r");

            //   Authorization:Digest username="RETS_O_74601_6", realm="connectmls-rets.mredllc.com", nonce="0d4194327a1c825d2672e5bf273cfdb6", uri="/rets/server/login", response="298d09bba177478ae5c54d1a2e7d4baa", opaque="5ccdef346870ab04ddfe0412367fccba", qop=auth, nc=00000013, cnonce="d2cef0abe8d87899"
            //RETS-Session-ID=24C99698CE44123BF98B40E98CB9888F; __utmt=1; __utmb=118572885.1.10.1462996245; JSESSIONID=F94B7B6B18567676CE1C05CE138A8B6A; __utma=118572885.1876298590.1455414260.1460149404.1462996245.5; __utmc=118572885; __utmz=118572885.1462996245.5.4.utmcsr=google|utmccn=(organic)|utmcmd=organic|utmctr=(not%20provided); JSESSIONID=133A6FF1B75101A4E6BA7E9DB79B41B0; __utma=118572885.1876298590.1455414260.1460149404.1462996245.5; __utmc=118572885; __utmz=118572885.1462996245.5.4.utmcsr=google|utmccn=(organic)|utmcmd=organic|utmctr=(not%20provided); SERVERID=ws17r
            loginRequest.Headers.Add("Authorization:Digest username=\"RETS_O_74601_6\", realm=\"connectmls-rets.mredllc.com\", nonce=\"0d4194327a1c825d2672e5bf273cfdb6\", uri=\"/rets/server/login\", response=\"298d09bba177478ae5c54d1a2e7d4baa\", opaque=\"5ccdef346870ab04ddfe0412367fccba\", qop=auth, nc=00000013, cnonce=\"d2cef0abe8d87899\"");
           queryRequest.Headers = loginRequest.Headers;
            loginRequest.Headers.Add(@"Cookie:RETS-Session-ID=24C99698CE44123BF98B40E98CB9888F; __utmt=1; __utmb=118572885.1.10.1462996245; JSESSIONID=F94B7B6B18567676CE1C05CE138A8B6A; __utma=118572885.1876298590.1455414260.1460149404.1462996245.5; __utmc=118572885; __utmz=118572885.1462996245.5.4.utmcsr=google|utmccn=(organic)|utmcmd=organic|utmctr=(not%20provided); JSESSIONID=133A6FF1B75101A4E6BA7E9DB79B41B0; __utma=118572885.1876298590.1455414260.1460149404.1462996245.5; __utmc=118572885; __utmz=118572885.1462996245.5.4.utmcsr=google|utmccn=(organic)|utmcmd=organic|utmctr=(not%20provided); SERVERID=ws17r");

            //curl "http://connectmls-rets.mredllc.com/rets/server/login" -H "Pragma: no-cache" -H "Accept-Encoding: gzip, deflate, sdch" -H "Accept-Language: en-US,en;q=0.8" -H "Upgrade-Insecure-Requests: 1" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.94 Safari/537.36" -H "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8" -H "Cache-Control: no-cache" -H "Authorization: Digest username=""RETS_O_74601_6"", realm=""connectmls-rets.mredllc.com"", nonce=""0d4194327a1c825d2672e5bf273cfdb6"", uri=""/rets/server/login"", response=""298d09bba177478ae5c54d1a2e7d4baa"", opaque=""5ccdef346870ab04ddfe0412367fccba"", qop=auth, nc=00000013, cnonce=""d2cef0abe8d87899""" -H "Cookie: RETS-Session-ID=24C99698CE44123BF98B40E98CB9888F; __utmt=1; __utmb=118572885.1.10.1462996245; JSESSIONID=F94B7B6B18567676CE1C05CE138A8B6A; __utma=118572885.1876298590.1455414260.1460149404.1462996245.5; __utmc=118572885; __utmz=118572885.1462996245.5.4.utmcsr=google|utmccn=(organic)|utmcmd=organic|utmctr=(not"%"20provided); JSESSIONID=133A6FF1B75101A4E6BA7E9DB79B41B0; __utma=118572885.1876298590.1455414260.1460149404.1462996245.5; __utmc=118572885; __utmz=118572885.1462996245.5.4.utmcsr=google|utmccn=(organic)|utmcmd=organic|utmctr=(not"%"20provided); SERVERID=ws17r" -H "Connection: keep-alive" --compressed

            try
            {
                //// Get the response.
                HttpWebResponse loginResponse;
                loginResponse = (HttpWebResponse)loginRequest.GetResponse();


                queryRequest.Headers.Add("Cookie:" + loginResponse.GetResponseHeader("Set-Cookie"));


                // Get the stream containing content returned by the server.
                Stream dataStream = loginResponse.GetResponseStream();
                // Open the stream using a StreamReader for easy access.
                StreamReader reader = new StreamReader(dataStream);
                // Read the content.
                string responseFromServer = reader.ReadToEnd();
                // Display the content.
               // MessageBox.Show(loginResponse.ContentLength.ToString());
               // MessageBox.Show(responseFromServer);
                ////Console.WriteLine(responseFromServer);
                //// Cleanup the streams and the response.
              

                //reader.close();
                //datastream.close();
                //response.close();


                HttpWebResponse queryResponse;
                queryResponse = (HttpWebResponse)queryRequest.GetResponse();
                dataStream = queryResponse.GetResponseStream();
                reader = new StreamReader(dataStream);
                responseFromServer = reader.ReadToEnd();
                MessageBox.Show(responseFromServer);


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void buttonSetupSearch_Click(object sender, EventArgs e)
        {
    
             mredCmdComboBox.Text = "New Map Search";
             runMredScriptButton_Click(sender, e);
        }

        private void buttonRunSearch_Click(object sender, EventArgs e)
        {
            mredCmdComboBox.Text = "FindComps";
            runMredScriptButton_Click(sender, e);
        }

        private void subjectBrokerPhoneTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {

        }

      
       
       

  

        //private void label6_Click(object sender, EventArgs e)
        //{

        //}
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

