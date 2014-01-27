using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{

    

    class OldRep    
    {
        private void WriteScript(string path, string filename, StringBuilder script)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(path + "\\" + filename))
            {
                file.Write(script);
            }


        }

        private Form1 callingForm;
        public Dictionary<string, string> StyleMlsTypeMap = new Dictionary<string, string>()
        {
            {"1 Story",        "%1<SP>Story"}, 
            {"1.5 Story",      "%1.5<SP>Stories"},
            {"2 Stories",      "%2<SP>Stories"},
            {"3 Stories",      "%Other"},
            {"4+ Stories", "%Other"},
            {"Coach House", "%1<SP>Story"},
            {"Earth", "%Other"},
            {"Hillside", "%Split<SP>Level"},
            {"Raised Ranch", "%Split<SP>Level"},
            {"Split Level", "%Split<SP>Level"},
            {@"Split Level w/ Sub", "%Split<SP>Level"},
            {"Other", "%Other"},
            {"Tear Down", "%Other"}
        };

        public string StyleString(string t)
        {

            return StyleMlsTypeMap[t];

            //string r = "";
            //switch (t)
            //{
            //    case "1 Story":
            //    case "Coach House":
            //        r = "%1<SP>Story";
            //        break;
            //    case "2 Stories":
            //        r = "%2<SP>Stories";
            //        break;
            //    case "1.5 Story":
            //        r = "%1.5<SP>Stories";
            //        break;
            //    case "Split Level":
            //    case "Hillside":
            //    case "Raised Ranch":
            //        numberOfAboveGradeLevels = 1;
            //        basementGLADivisionFactor = 3;
            //        break;
            //    case @"Split Level w/Sub":
            //        numberOfAboveGradeLevels = 1;
            //        basementGLADivisionFactor = 2;
            //        break;
                    
            //    case "Other":
            //    case "Tear Down":
            //    case "3 Stories":
            //    case "4+ Stories":
            //    default:
            //      r = "%Other";
            //}
            //return r;
        }

        private string BasementString()
        {
            string s = "%None";
           if (callingForm.SubjectBasementDetails.ToLower().Contains("unfinished"))
           {
               s = "%Unfinished";
           }
           if ((callingForm.SubjectBasementType.ToLower().Contains("full") || callingForm.SubjectBasementType.ToLower().Contains("partial") && !callingForm.SubjectBasementDetails.ToLower().Contains("unfinished")))
            {
                s = "%90%<SP>Finished";
            }
            return s;
        }

        private string GarageString()
        {

            string s = "%";
            string stalls = Regex.Match(callingForm.SubjectParkingType, @"\d").Value;
            string ad = Regex.Match(callingForm.SubjectParkingType, @"det|att", RegexOptions.IgnoreCase).Value;
             
            return s + stalls + "Car" + ad[0].ToString().ToUpper() + ad.Substring(1);
        }
            

        public void Prefill(iMacros.App iim, Form1 form)
        {
            callingForm = form;
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 1");
        

            Dictionary<string, string> fieldList = new Dictionary<string, string>();

            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentBPOSubType CONTENT=%Exterior<SP>Only");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentNeighborhoodImpact CONTENT=%Appropriate<SP>Improvement");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentFinancingAvailable CONTENT=%Yes");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentSubjectLeaseHoldFeeSimple CONTENT=%Fee<SP>Simple");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentSubjectView CONTENT=%Average");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentLocationCondition CONTENT=%Rural");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentLocationCondition CONTENT=%Suburban");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentSubjectDesignAppeal CONTENT=%Average");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentSubjectQualityOfConstruction CONTENT=%Average");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentSubjectFunctionalUtility CONTENT=%Average");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentEmploymentConditions CONTENT=%Stable");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentAreaType CONTENT=$Select<SP>a<SP>Value>>");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentHousingSupply CONTENT=%Stable");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=ID:ifcContentrbNewResConstructionNo&&VALUE:ifcContentrbNewResConstructionNo CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=ID:ifcContentrbNewCommConstructionNo&&VALUE:ifcContentrbNewCommConstructionNo CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentNeighborhoodCrime CONTENT=%Low<SP>Risk");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentNeighborhoodCrime CONTENT=%Minimal<SP>Risk");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentNeighborhoodMarketCondition CONTENT=%Improving");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentNeighborhoodPropertyValues CONTENT=%Increasing");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentTrendRate CONTENT=1");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentTrendRateMonths CONTENT=3");
            macro.AppendLine(@"'TAG POS=0 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentTrendRate CONTENT=3");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=ID:ifcContentrbMarketAsIs&&VALUE:ifcContentrbMarketAsIs CONTENT=YES");
           
            //subject info
            //needs selection box logic
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentStyle CONTENT=" + StyleString(form.SubjectMlsType));
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentBasementComplete CONTENT=" + BasementString());
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentParking CONTENT=" + GarageString());
            
            if (form.SubjectDetached)
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentPropertyType CONTENT=%SFR<SP>Detached");
            }

            if (form.SubjectAttached)
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentPropertyType CONTENT=%SFR<SP>Attached");
            }
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentNumUnits CONTENT=1");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentSquareFeet CONTENT=" + form.SubjectAboveGLA.Replace(",",""));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentSubjectSite CONTENT=" + form.SubjectLotSize);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentAge CONTENT=" + form.SubjectYearBuilt);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentRooms CONTENT=" + form.SubjectRoomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentRoomsBedroom CONTENT=" + form.SubjectBedroomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentRoomsBath CONTENT=" + form.SubjectBathroomCount);
            
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentCondition CONTENT=%Average");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentSubjectParcelNumber CONTENT=" + form.SubjectPin);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentBrokerSubjectDistance CONTENT=" + form.SubjectProximityToOffice);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentSubjectLastSaleDate CONTENT=" + GlobalVar.theSubjectProperty.DateOfLastSale);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentSubjectLastSalePrice CONTENT=" + form.SubjectLastSalePrice);
            macro.AppendLine(@"TAG POS=0 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentPropertyTaxes CONTENT=" + GlobalVar.theSubjectProperty.PropertyTax);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentHOAFees CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentHOAPaymentCycle CONTENT=%Annually");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentHOAPaymentCurrent CONTENT=%Yes");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentHOAFeesDelinquent CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentProjectLegalAction CONTENT=%No");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentSubjectSchoolID CONTENT=%2");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentCurrentListing CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentPreviousListing CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentWhyUnsold CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentOccupancyStatus CONTENT=%Owner");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentOwnershipStatus CONTENT=%Main<SP>Residence");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentSubjectSecured CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentSubjectHeatingCooling CONTENT=GFA.C");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentSubjectEnergyEfficientItems CONTENT=NA");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentSubjectPorch CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentSubjectPatio CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentSubjectDeck CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentSubjectFirePlace CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentSubjectPool CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentSubjectFence CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentSubjectOtherItems CONTENT=NA");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentMelloRoos CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentSubjectDataSource CONTENT=%TaxRecords");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodBoundariesDefined CONTENT=1<SP>mile<SP>radius");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentPercentOwner CONTENT=90");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentPercentTenant CONTENT=10");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentREOCount CONTENT=5");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentBoardedCount CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodAverageAge CONTENT=25");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentAreaType CONTENT=%Suburban");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodNewResPriceLow CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodNewResPriceHigh CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodListingsInArea CONTENT=35");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodCompetitionInArea CONTENT=9");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodCompetitionPriceLow CONTENT=150000");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodCompetitionPriceHigh CONTENT=250000");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodSalesInArea CONTENT=3");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodSalesPriceLow CONTENT=180000");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodSalesPriceHigh CONTENT=225000");
            macro.AppendLine(@"TAG POS=3 TYPE=TD FORM=NAME:frmImportFormRender ATTR=TXT:Select<SP>a<SP>Value>>DecreasingIncreasingStable");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodListings CONTENT=%Stable");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodDOM CONTENT=90");
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodComments CONTENT=None");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentCompOneOwnership CONTENT=%Owner<SP>Occupant");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentCompOneProximity CONTENT=%1<SP>mile");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentCompTwoOwnership CONTENT=%Owner<SP>Occupant");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentCompTwoProximity CONTENT=%1<SP>mile");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentCompThreeOwnership CONTENT=%Owner<SP>Occupant");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentCompThreeProximity CONTENT=%1<SP>mile");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentCompOneSchoolID CONTENT=%2");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentCompTwoSchoolID CONTENT=%2");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentCompThreeSchoolID CONTENT=%2");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentCompMostComparable CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentCompOneCompWhyVariance CONTENT=Same");
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentCompTwoCompWhyVariance CONTENT=Same");
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentCompThreeCompWhyVariance CONTENT=Same");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentListingOneOwnership CONTENT=%Owner<SP>Occupant");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentListingOneProximity CONTENT=%1<SP>mile");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentListingTwoOwnership CONTENT=%Owner<SP>Occupant");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentListingTwoProximity CONTENT=%1<SP>mile");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentListingThreeOwnership CONTENT=%Investor");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentListingThreeOwnership CONTENT=%Owner<SP>Occupant");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentListingThreeProximity CONTENT=%1<SP>mile");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentListingOneSchoolID CONTENT=%2");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentListingTwoSchoolID CONTENT=%2");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentListingThreeSchoolID CONTENT=%2");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentListingMostComparable CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentListingOneInspected CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentListingTwoInspected CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentListingThreeInspected CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentMktOccupancyStatus CONTENT=%Owner");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentMostLikelyBuyer CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=NAME:frmImportFormRender ATTR=TXT:ORDMS<SP>has<SP>the<SP>real<SP>estate<SP>license<SP>#<SP>471009163<SP>on<SP>file<SP>for<SP>the<SP>broker<SP>Dawn<SP>Zurick");
            macro.AppendLine(@"TAG POS=1 TYPE=B FORM=NAME:frmImportFormRender ATTR=TXT:471009163");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentPreparersLicense");
            macro.AppendLine(@"TAG POS=1 TYPE=B FORM=NAME:frmImportFormRender ATTR=TXT:471009163");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentPreparersLicense CONTENT=471009163");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentPreparingBrokersName CONTENT=Dawn<SP>Zurick");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentBrokerLicensedDate CONTENT=9/1/2006");
            macro.AppendLine(@"TAG POS=3 TYPE=INPUT:RADIO FORM=NAME:frmImportFormRender ATTR=NAME:ifcContentBPOReason CONTENT=YES");

            //tbd
            //rent
            // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentEstimatedRent CONTENT=" & );

            WriteScript(form.SubjectFilePath, "subject-prefill.iim", macro);

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 120);


        }

        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {
            string sol;
            if (saleOrList == "sale")
            {
                sol = "Comp";
            }
            else
            {
                sol = "Listing";
            }
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");

            macro.AppendLine(@"SET !TIMEOUT_STEP 0");
            //macro.AppendLine(@"FRAME NAME=pageView");
            macro.AppendLine(@"SET !REPLAYSPEED FAST");

           

            foreach (string field in fieldList.Keys)
            {
                if (field.Contains("*"))
                {
                    //drop down box
                     //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContentListingOneLeaseHoldFeeSimple CONTENT=%Fee<SP>Simple");

                     macro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=ID:ifcContent{0}{1} CONTENT=%{2}\r\n", compNum, field.Replace("*", ""), fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                    //macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:*{0}{1}_{2}\r\n", sol, field.Replace("*", ""), Regex.Match(compNum, @"\d").Value);
                    //macro.AppendLine(@"DS CMD=KEY X={{!TAGX}} Y={{!TAGY}} CONTENT=" + fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                }
                else
                {
                      macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:*{0}{1} CONTENT={2}\r\n" ,compNum, field, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                    //macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:*{0}{1}{2} CONTENT={3}\r\n", sol, Regex.Match(compNum, @"\d").Value, field, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                }

                //  macro.AppendLine(@"DS CMD=CLICK X={{!TAGX}} Y={{!TAGY}}");
            }

            //radio buttons
            //this is turning on the "No" porch button
            macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=ID:ifcContentrb{0}PorchNo&&VALUE:ifcContentrb{0}PorchNo CONTENT={1}\r\n", compNum, "YES" );
            macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=ID:ifcContentrb{0}PatioNo&&VALUE:ifcContentrb{0}PatioNo CONTENT={1}\r\n", compNum, "YES");
            macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=ID:ifcContentrb{0}DeckNo&&VALUE:ifcContentrb{0}DeckNo CONTENT={1}\r\n", compNum, "YES");
            macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=ID:ifcContentrb{0}FirePlaceNo&&VALUE:ifcContentrb{0}FirePlaceNo CONTENT={1}\r\n", compNum, "YES");
            macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=ID:ifcContentrb{0}PoolNo&&VALUE:ifcContentrb{0}PoolNo CONTENT={1}\r\n", compNum, "YES");
            macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=ID:ifcContentrb{0}FenceNo&&VALUE:ifcContentrb{0}FenceNo CONTENT={1}\r\n", compNum, "YES");

            //
            //TBD
            //
            //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListOLDate_1");
            //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListCLPrice_1");
            //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListCLDate_1");

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 60);
        }
    }
}
