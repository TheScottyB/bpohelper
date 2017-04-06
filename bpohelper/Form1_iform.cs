using System;
using System.Collections.Generic;
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
//using System.Diagnostics;
using System.Threading;
using System.Xml;
using System.Xml.Schema;

namespace bpohelper
{
    public partial class Form1 : System.Windows.Forms.Form, IYourForm
    {
        //
        //iform properties
        //

        //
        //Subject Related - Exterior Features
        //

        //subjectLandValueTextBox


        public bool PerserveNeighorhoodData
        {
            get { return ndPreserveData.Checked; }
        }

        public bool PerserveCompSetData
        {
            get { return cdPreserveData.Checked; }
        }

        public string SubjectLandValue
        {
            get { return subjectLandValueTextBox.Text; }
            set { subjectLandValueTextBox.Text = value; }

        }

        public bool IgnoreAge
        {
            get { return ignoreAgeCheckBox.Checked; }

        }

        public bool IgnoreType
        {
            get { return ignoreTypeCheckBox.Checked; }

        }

        public string SubjectAssessmentValue
        {
            get { return subjectAssessmentTextbox.Text; }
            set { subjectAssessmentTextbox.Text = value; }

        }

        public string SearchMapRadius
        {
            get { return numericUpDownRadius.Value.ToString(); }
        }

        public string SubjectExteriorFinish

        {
            get { return subjectExteriorFinishTextbox.Text; }
            set { subjectExteriorFinishTextbox.Text = value; }
        }

        public string StatusUpdate
        {
            set { programStatusRichTextBox.AppendText(value + "\r\n"); }
        }

        public bool CacheSearch
        {
            get { return recheckLastSearchbutton.Enabled; }
            set { recheckLastSearchbutton.Enabled = value; }
        }

        public string SubjectBrokerPhone
        {
            get { return subjectBrokerPhoneTextBox.Text; }
            set { subjectBrokerPhoneTextBox.Text = value; }
        }

        public string SubjectListingAgent
        {
            get
            {
                string pattern = @"(.*)\(";
                Match m = Regex.Match(subjectListingAgentTextBox.Text, pattern); 
                if (m.Success)
                {
                    return m.Groups[1].Value;
                }


                return subjectListingAgentTextBox.Text;
            }
            set { subjectListingAgentTextBox.Text = value; }
        }

        public string SubjectDom
        {
            get { return subjectDomTextBox.Text; }
            set { subjectDomTextBox.Text = value; }
        }
        //subjectCurrentListPriceTextBox
        public string SubjectCurrentListPrice
        {
            get { return subjectCurrentListPriceTextBox.Text; }
            set { subjectCurrentListPriceTextBox.Text = value; }
        }
        //

        public string SubjectMlsStatus
        {
            get { return subjectMlsStatusTextBox.Text; }
            set { subjectMlsStatusTextBox.Text = value; }
        }


        public string SubjectRent
        {
            get { return subjectRentTextbox.Text; }
            set { subjectRentTextbox.Text = value; }
        }

        public string SubjectAvm
        {
            get { return subjectAvmTextBox.Text; }
            set { subjectAvmTextBox.Text = value; }
        }


        public bool UpdateRealist
        {
            get { return realistUpdateCheckBox.Checked; }
            set { realistUpdateCheckBox.Checked = value; }
        }

        public string CurrentSearchName
        {
            get { return streetnameTextBox.Text; }
            set { streetnameTextBox.Text = value; }
        }

        public string AddInfo
        {
            set { infoWindowrichTextBox.AppendText(value); }
        }


        public string NumberOfCompsFound
        {
            set
            { 
                richTextBoxNumberCompsFound.Text = value;
               
            }
        }

        public void AddInfoColor(string value, Color c)
        {
            infoWindowrichTextBox.AppendText(value, c);
        }

        public string SetStatusBar
        {
            set { //toolStripStatusLabel1.Text = value; 
            }
        }

        public string PicDiffLabel
        {
            set { picDiffLabel.Text = value; }
        }

        public Bitmap SubjectPic
        {
            set { subjectPictureBox.Image = value; }
        }

        public Bitmap CompPic
        {
            set { compPictureBox.Image = value; }
        }

        public bool SubjectAttached
        {
            get { return subjectAttachedRadioButton.Checked; }
            set { subjectAttachedRadioButton.Checked = value; }
        }

        public bool SubjectDetached
        {
            get { return subjectDetachedradioButton.Checked; }
            set { subjectDetachedradioButton.Checked = value; }
        }


        public AssessmentInfo SubjectAssessmentInfo
        {
            get { return subjectAssessmentInfo; }
            set { subjectAssessmentInfo = value; }
        }

        public Neighborhood SubjectNeighborhood
        {
            get { return subjectNeighborhood; }
            set { subjectNeighborhood = value; }
        }
        public Neighborhood SetOfComps
        {
            get { return setOfComps; }
            set { setOfComps = value; }
        }
        

        public string SubjectMlsType
        {
            get { return subjectMlsTypecomboBox.Text; }
            set { subjectMlsTypecomboBox.Text = value; }
        }

        public string SubjectPin
        {
            get { return subjectpin_textbox.Text; }
            set { subjectpin_textbox.Text = value; }
        }

        public string SubjectSchoolDistrict
        {
            get { return subject_school_textbox.Text; }
            set { subject_school_textbox.Text = value; }
        }

        public string SubjectLastSalePrice
        {
            get { return subjectLastSalePriceTextbox.Text; }
            set { subjectLastSalePriceTextbox.Text = value; }
        }

        public string SubjectLastSaleDate
        {
            get { return subjectLastSaleDateTextbox.Text; }
            set { subjectLastSaleDateTextbox.Text = value; }
        }

        public string SubjectAboveGLA
        {
            get { return subjectAboveGlaTextbox.Text; }
            set { subjectAboveGlaTextbox.Text = value; }
        }

        public string SubjectLotSize
        {
            get { return subjectLotSizeTextbox.Text; }
            set { subjectLotSizeTextbox.Text = value; }
        }

        public string SubjectYearBuilt
        {
            get { return subjectYearBuiltTextbox.Text; }
            set { subjectYearBuiltTextbox.Text = value; }
        }

        public string SubjectRoomCount
        {
            get { return subjectRoomCountTextbox.Text; }
            set { subjectRoomCountTextbox.Text = value; }
        }

        public string SubjectBedroomCount
        {
            get { return subjectBedroomTextbox.Text; }
            set { subjectBedroomTextbox.Text = value; }
        }

        public string SubjectBathroomCount
        {
            get { return subjectBathroomTextbox.Text; }
            set { subjectBathroomTextbox.Text = value; }
        }

        public string SubjectBasementType
        {
            get { return subjectBasementTypeTextbox.Text; }
            set { subjectBasementTypeTextbox.Text = value; }
        }

        public string SubjectBasementDetails
        {
            get { return subjectBasementDetailsTextbox.Text; }
            set { subjectBasementDetailsTextbox.Text = value; }
        }


        public string SubjectBasementGLA
        {
            get { return subjectBasementGlaTextbox.Text; }
            set { subjectBasementGlaTextbox.Text = value; }
        }

        public string SubjectBasementFinishedGLA
        {
            get { return subjectFinishedBasementGlaTextBox.Text; }
            set { subjectFinishedBasementGlaTextBox.Text = value; }
        }

        //SubjectBasementFinishedGLA

        public string SubjectNumberFireplaces
        {
            get { return subjectNumFireplacesTextbox.Text; }
            set { subjectNumFireplacesTextbox.Text = value; }
        }

        public string SubjectParkingType
        {
            get { return subjectParkingTypeTextbox.Text; }
            set { subjectParkingTypeTextbox.Text = value; }
        }

        public string SubjectOOR
        {
            get { return subjectOorTextbox.Text; }
            set { subjectOorTextbox.Text = value; }
        }

        public string SubjectFullAddress
        {
            get { return subjectFullAddressTextbox.Text; }
            set { subjectFullAddressTextbox.Text = value; }
        }

        public string SubjectFilePath
        {
            get { return search_address_textbox.Text; }
            set { search_address_textbox.Text = value; }
        }

        public string SubjectCounty
        {
            get { return subjectCountyTextbox.Text; }
            set { subjectCountyTextbox.Text = value; }
        }

        public string SubjectStyle
        {
            get { return subjectStyleTextbox.Text; }
            set { subjectStyleTextbox.Text = value; }
        }

        public string SubjectProximityToOffice
        {
            get { return subjectProximityToOfficeTextbox.Text; }
            set { subjectProximityToOfficeTextbox.Text = value; }
        }

        public _BPO_SandboxDataSet1TableAdapters.RawSFDataTableAdapter RawSFDatatable
        {
            get { return this.rawSFDataTableAdapter; }
        }

        public string SubjectMarketValue
        {
            get { return valueTextbox.Text; }
            set { valueTextbox.Text = value; }
        }

        public string SubjectMarketValueList
        {
            get { return listTextbox.Text; }
            set { listTextbox.Text = value; }
        }

        public string SubjectQuickSaleValue
        {
            get { return quickSaleTextbox.Text; }
            set { quickSaleTextbox.Text = value; }
        }

        public string SubjectSubdivision
        {
            get { return subjectSubdivisionTextbox.Text; }
            set { subjectSubdivisionTextbox.Text = value; }
        }
    }
}
