using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{

    

    class OldRep    
    {
        protected void WriteScript(string path, string filename, StringBuilder script)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(path + "\\" + filename))
            {
                file.Write(script);
            }


        }

         

        protected Form1 callingForm;
        public Dictionary<string, string> StyleMlsTypeMap = new Dictionary<string, string>()
        {
            {"1 Story",        "1<SP>Story"}, 
            {"1.5 Story",      "1.5<SP>Stories"},
            {"2 Stories",      "2<SP>Stories"},
            {"3 Stories",      "Other"},
            {"4+ Stories", "Other"},
            {"Coach House", "1<SP>Story"},
            {"Earth", "Other"},
            {"Hillside", "Split<SP>Level"},
            {"Raised Ranch", "Split<SP>Level"},
            {"Split Level", "Split<SP>Level"},
            {@"Split Level w/ Sub", "Split<SP>Level"},
            {"Other", "Other"},
            {"Tear Down", "Other"}
        };

        public string StyleString(string t)
        {

             return StyleMlsTypeMap[t.Split(',')[0]];
            //if (t.Contains(","))
            //{
            //    t = t.Split(',')[0];
            //}

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
            //    //case "Split Level":
            //    //case "Hillside":
            //    //case "Raised Ranch":
            //    //    numberOfAboveGradeLevels = 1;
            //    //    basementGLADivisionFactor = 3;
            //    //    break;
            //    //case @"Split Level w/Sub":
            //    //    numberOfAboveGradeLevels = 1;
            //    //    basementGLADivisionFactor = 2;
            //    //    break;

            //    //case "Other":
            //    //case "Tear Down":
            //    //case "3 Stories":
            //    //case "4+ Stories":
            //    default:
            //        r = "%Other";
            //        break;
            //}
            //return r;
        }

        protected  string BasementString()
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

        protected string GarageString()
        {
            try
            {
                string s = "%";
                string stalls = Regex.Match(callingForm.SubjectParkingType, @"\d").Value;
                string ad = Regex.Match(callingForm.SubjectParkingType, @"det|att", RegexOptions.IgnoreCase).Value;
                return s + stalls + "Car" + ad[0].ToString().ToUpper() + ad.Substring(1);
            }
            catch
            {
                return "%None";
            }

        }
            

        public void Prefill(iMacros.App iim, Form1 form)
        {
            callingForm = form;
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");
        

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
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentPropertyTaxes CONTENT=" + GlobalVar.theSubjectProperty.PropertyTax);

            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentHOAFees CONTENT=0");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentHOAPaymentCycle CONTENT=%Annually");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentHOAPaymentCurrent CONTENT=%Yes");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentHOAFeesDelinquent CONTENT=0");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentProjectLegalAction CONTENT=%No");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentSubjectSchoolID CONTENT=%2");
            //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentCurrentListing CONTENT=YES");
            //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentPreviousListing CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentWhyUnsold CONTENT=na");

            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentOccupancyStatus CONTENT=%Owner");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentOwnershipStatus CONTENT=%Main<SP>Residence");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentSubjectSecured CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentSubjectHeatingCooling CONTENT=GFA.C");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentSubjectEnergyEfficientItems CONTENT=NA");
            //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentSubjectPorch CONTENT=YES");
            //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentSubjectPatio CONTENT=YES");
            //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentSubjectDeck CONTENT=YES");
            //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentSubjectFirePlace CONTENT=YES");
            //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentSubjectPool CONTENT=YES");
            //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentSubjectFence CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentSubjectOtherItems CONTENT=NA");
           // macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentMelloRoos CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentSubjectDataSource CONTENT=%TaxRecords");

            string subjectComments = "The subject is a conforming home within the neighborhood. No adverse conditions were noted at the time of inspection based on exterior observations. Unable to determine interior condition due to exterior inspection only, so subject was assumed to be in average condition for this report. No known special concerns, encroachments, easements, water rights, environmental concerns, flood zones.";


            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:frmImportFormRender ATTR=NAME:ifcContentSubjectComments CONTENT=" + subjectComments.Replace(" ", "<SP>"));


            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodBoundariesDefined CONTENT=1<SP>mile<SP>radius");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentPercentOwner CONTENT=90");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentPercentTenant CONTENT=10");
            //reo comp count
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentREOCount CONTENT=" + form.setOfComps.numberREOSales);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentBoardedCount CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodAverageAge CONTENT=" + form.subjectNeighborhood.medianAge);
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentAreaType CONTENT=%Suburban");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodNewResPriceLow CONTENT=0");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodNewResPriceHigh CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodListingsInArea CONTENT=" + form.subjectNeighborhood.numberActiveListings);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodCompetitionInArea CONTENT=" + form.setOfComps.numberActiveListings);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodCompetitionPriceLow CONTENT=" + form.setOfComps.minListPrice);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodCompetitionPriceHigh CONTENT=" + form.setOfComps.maxListPrice);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodSalesInArea CONTENT=" + form.setOfComps.numberSoldListings);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodSalesPriceLow CONTENT=" + form.setOfComps.minSalePrice);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodSalesPriceHigh CONTENT=" + form.setOfComps.maxSalePrice);
           // macro.AppendLine(@"TAG POS=3 TYPE=TD FORM=ID:frmImportFormRender ATTR=TXT:Select<SP>a<SP>Value>>DecreasingIncreasingStable");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodListings CONTENT=%Stable");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodDOM CONTENT=" + form.setOfComps.avgDom);


            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodSixMonthAbsorptionRate CONTENT=" + (Math.Round(form.subjectNeighborhood.AbsorbtionRate * 6)).ToString());
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodCurrentInventory CONTENT=" + form.subjectNeighborhood.numberActiveListings);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodMonthsSupply CONTENT=" + form.subjectNeighborhood.MonthsSupply);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodSaletoListRatio CONTENT=" + (form.subjectNeighborhood.saleToListRatio * 100).ToString());
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhood3MonthAverageListPrice CONTENT=" + form.subjectNeighborhood.ThreeMonthListPrice.ToString());
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhood6MonthAverageListPrice CONTENT=" + form.subjectNeighborhood.SixMonthListPrice.ToString());
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhood3MonthAverageSalePrice CONTENT=" + form.subjectNeighborhood.ThreeMonthSalePrice.ToString());
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhood6MonthAverageSalePrice CONTENT=" + form.subjectNeighborhood.SixMonthSalePrice.ToString());



            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:frmImportFormRender ATTR=NAME:ifcContentNeighborhoodComments CONTENT=<TBD>");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentCompOneOwnership CONTENT=%Owner<SP>Occupant");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentCompOneProximity CONTENT=%1<SP>mile");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentCompTwoOwnership CONTENT=%Owner<SP>Occupant");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentCompTwoProximity CONTENT=%1<SP>mile");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentCompThreeOwnership CONTENT=%Owner<SP>Occupant");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentCompThreeProximity CONTENT=%1<SP>mile");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentCompOneSchoolID CONTENT=%2");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentCompTwoSchoolID CONTENT=%2");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentCompThreeSchoolID CONTENT=%2");

            string defaultCompComment = "Reasonable substitute for the subject property, similar in most areas.".Replace(" ", "<SP>");

            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentCompMostComparable CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:frmImportFormRender ATTR=NAME:ifcContentCompOneCompWhyVariance CONTENT=" + defaultCompComment);
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:frmImportFormRender ATTR=NAME:ifcContentCompTwoCompWhyVariance CONTENT=" + defaultCompComment);
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:frmImportFormRender ATTR=NAME:ifcContentCompThreeCompWhyVariance CONTENT=" + defaultCompComment);
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListingOneOwnership CONTENT=%Owner<SP>Occupant");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListingOneProximity CONTENT=%1<SP>mile");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListingTwoOwnership CONTENT=%Owner<SP>Occupant");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListingTwoProximity CONTENT=%1<SP>mile");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListingThreeOwnership CONTENT=%Investor");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListingThreeOwnership CONTENT=%Owner<SP>Occupant");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListingThreeProximity CONTENT=%1<SP>mile");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListingOneSchoolID CONTENT=%2");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListingTwoSchoolID CONTENT=%2");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListingThreeSchoolID CONTENT=%2");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListingMostComparable CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListingOneInspected CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListingTwoInspected CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListingThreeInspected CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentMktOccupancyStatus CONTENT=%Owner");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentMostLikelyBuyer CONTENT=YES");


            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListPriceRepaired90120 CONTENT=" + form.SubjectMarketValueList);
            //macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentValueQuickSale CONTENT=" + form.SubjectQuickSaleValue);
            //macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListPriceQuickSale CONTENT=" + form.SubjectQuickSaleValue);
            //macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");

            //    macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=ID:frmImportFormRender ATTR=TXT:ORDMS<SP>has<SP>the<SP>real<SP>estate<SP>license<SP>#<SP>471009163<SP>on<SP>file<SP>for<SP>the<SP>broker<SP>Dawn<SP>Zurick");
            //macro.AppendLine(@"TAG POS=1 TYPE=B FORM=ID:frmImportFormRender ATTR=TXT:471009163");
            //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentPreparersLicense");
            // macro.AppendLine(@"TAG POS=1 TYPE=B FORM=ID:frmImportFormRender ATTR=TXT:471009163");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentPreparersLicense CONTENT=471009163");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentPreparingBrokersName CONTENT=Dawn<SP>Zurick");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentBrokerLicensedDate CONTENT=9/1/2006");
            macro.AppendLine(@"TAG POS=3 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentBPOReason CONTENT=YES");

            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentEstimatedRent CONTENT=" + form.SubjectRent);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentValueAsIs90120 CONTENT=" + form.SubjectMarketValue);
            macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
            macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
            macro.AppendLine(@"ONDIALOG POS=2 BUTTON=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListPriceAsIs90120 CONTENT=" + form.SubjectMarketValueList);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentLandValueAsIs CONTENT=" + form.SubjectLandValue);
            macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");


            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentValueRepaired90120 CONTENT=" + form.SubjectMarketValue);
           // macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES")
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListPriceRepaired90120 CONTENT=" + form.SubjectMarketValueList);
           // macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentValueQuickSale CONTENT=" + form.SubjectQuickSaleValue);
           // macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentListPriceQuickSale CONTENT=" + form.SubjectQuickSaleValue);
          // macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");

        //    macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=ID:frmImportFormRender ATTR=TXT:ORDMS<SP>has<SP>the<SP>real<SP>estate<SP>license<SP>#<SP>471009163<SP>on<SP>file<SP>for<SP>the<SP>broker<SP>Dawn<SP>Zurick");
            //macro.AppendLine(@"TAG POS=1 TYPE=B FORM=ID:frmImportFormRender ATTR=TXT:471009163");
          //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentPreparersLicense");
           //// macro.AppendLine(@"TAG POS=1 TYPE=B FORM=ID:frmImportFormRender ATTR=TXT:471009163");
           // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentPreparersLicense CONTENT=471009163");
           // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentPreparingBrokersName CONTENT=Dawn<SP>Zurick");
           // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=NAME:ifcContentBrokerLicensedDate CONTENT=9/1/2006");
           // macro.AppendLine(@"TAG POS=3 TYPE=INPUT:RADIO FORM=ID:frmImportFormRender ATTR=NAME:ifcContentBPOReason CONTENT=YES");



            //tbd
            //rent
            // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:frmImportFormRender ATTR=ID:ifcContentEstimatedRent CONTENT=" & );

            WriteScript(form.SubjectFilePath, "subject-prefill.iim", macro);

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 120);


        }

        public virtual void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
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

            //macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=ACTION:ImportFormRender.aspx?ImportFormID=2188164 ATTR=ID:ifcContentListingOnegMapImage");
            //macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=ACTION:ImportFormRender.aspx?ImportFormID=2188164 ATTR=ID:ifcContentListingTwogMapImage");
            //macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=ACTION:ImportFormRender.aspx?ImportFormID=2188164 ATTR=ID:ifcContentListingThreegMapImage");
            //macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=ACTION:ImportFormRender.aspx?ImportFormID=2188164 ATTR=ID:ifcContentCompOnegMapImage");
            //macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=ACTION:ImportFormRender.aspx?ImportFormID=2188164 ATTR=ID:ifcContentCompTwogMapImage");
            macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=ACTION:ImportFormRender.aspx?ImportFormID=2188164 ATTR=ID:ifcContent" + compNum + "gMapImage");

            //
            //TBD
            //
            //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListOLDate_1");
            //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListCLPrice_1");
            //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListCLDate_1");

           WriteScript(    GlobalVar.mainWindow.SubjectFilePath, compNum + "-fill.iim", macro);
        
            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 60);
        }
    }

    class OldRepFormV8:OldRep
    {
        public override void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {

            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");
            macro.AppendLine(@"SET !REPLAYSPEED FAST");

            base.CompFill(iim, saleOrList, compNum, fieldList);

            macro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx* ATTR=NAME:ifcContent{0}PropertyType CONTENT={1}\r\n", compNum, "%SFR<SP>Detached");
            macro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx* ATTR=NAME:ifcContent{0}SiteUnits CONTENT={1}\r\n", compNum, "%Acres");


            WriteScript(GlobalVar.mainWindow.SubjectFilePath, compNum + "-V8-fill.iim", macro);

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 60);

        }


//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneProximityDirection CONTENT=$Select<SP>a<SP>Value>>");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneSaleType CONTENT=$Select<SP>a<SP>Value>>");
//macro.AppendLine(@"TAG POS=4 TYPE=TD FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=TXT:(days)");
//macro.AppendLine(@"TAG POS=2 TYPE=TD FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=TXT:Select<SP>a<SP>Value>>County<SP>TaxDataQuickMLSOtherSitex");
//macro.AppendLine(@"TAG POS=4 TYPE=TD FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=TXT:Select<SP>a<SP>Value>>1.<SP>REO<SP>Sale2.<SP>Short<SP>Sale3.<SP>Court<SP>Ordered4.<SP>Estate<SP>Sale5.<SP>Reloca*");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneLocationFactor CONTENT=$Select<SP>a<SP>Value>>");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneLocation CONTENT=%N");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneSiteUnits CONTENT=%Acres");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneView CONTENT=%N");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneViewFactor CONTENT=%Res");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneDesign CONTENT=%Other");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneAttachedStructure CONTENT=$Select<SP>a<SP>Value>>");
//macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneDesignOther");
//macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=TXT:Select<SP>a<SP>Value>>Attached<SP>StructureDetached<SP>StructureSemi-detached<SP>structure");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneAttachedStructure CONTENT=$Select<SP>a<SP>Value>>");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneQualityOfConstruction CONTENT=%Q3");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneCondition CONTENT=%C4");
//macro.AppendLine(@"TAG POS=13 TYPE=TD FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=TXT:Select<SP>a<SP>Value>>NoYes");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneMaterialWorkLast15 CONTENT=$Select<SP>a<SP>Value>>");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentListingOnePropertyType CONTENT=%SFR<SP>Detached");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentListingTwoPropertyType CONTENT=%SFR<SP>Detached");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentListingThreePropertyType CONTENT=%SFR<SP>Detached");
//string macroCode = macro.ToString();
//// status = iim.iimPlayCode(macroCode, timeout);

//        macro.AppendLine(@"URL GOTO=https://ort.quandis.com/Decision/ImportFormRender.aspx?ImportFormID=2308912");
//macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneStoriesLevels CONTENT=2");
//macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompTwoStoriesLevels CONTENT=2");
//macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompThreeStoriesLevels CONTENT=2");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneQualityOfConstruction CONTENT=%Q3");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompTwoQualityOfConstruction CONTENT=%Q3");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompThreeQualityOfConstruction CONTENT=%Q3");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneCondition CONTENT=%C3");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompTwoCondition CONTENT=%C3");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompThreeCondition CONTENT=%C3");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneMaterialWorkLast15 CONTENT=$Select<SP>a<SP>Value>>");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneHeating CONTENT=%FWA");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompTwoHeating CONTENT=%FWA");
//macro.AppendLine(@"TAG POS=4 TYPE=TD FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=TXT:Select<SP>a<SP>Value>>Forced<SP>Warm<SP>AirHot<SP>Water<SP>Base<SP>BoardRadiantOtherNone");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompThreeHeating CONTENT=%FWA");
//macro.AppendLine(@"TAG POS=2 TYPE=TD FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=TXT:Select<SP>a<SP>Value>>Central<SP>ACIndividual<SP>ACOtherNone");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneCooling CONTENT=%Central<SP>AC");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompTwoCooling CONTENT=%Central<SP>AC");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompThreeCooling CONTENT=%Central<SP>AC");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneAttachedGarageSpaces CONTENT=%0");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneDetachedGarageSpaces CONTENT=%1");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneBuiltInGarageSpaces CONTENT=$Select<SP>a<SP>Value>>");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneCommonPorch CONTENT=%No");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneCommonPatio CONTENT=%No");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneCommonPool CONTENT=%No");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompTwoCommonPool CONTENT=%No");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompThreeCommonPool CONTENT=%No");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneCommonFence CONTENT=%No");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompTwoCommonFence CONTENT=%No");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompThreeCommonFence CONTENT=%No");
//macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneCommonFireplaces CONTENT=3");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOnePorch CONTENT=%No");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOnePatio CONTENT=%No");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOnePool CONTENT=%No");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneBasementExists CONTENT=%Yes");
//macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneBasementAccess CONTENT=%in");
//macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneBasementRecRooms CONTENT=0");
//macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneBasementBedrooms CONTENT=0");
//macro.AppendLine(@"TAG POS=1 TYPE=TABLE FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=TXT:SALE<SP>#1SALE<SP>#2SALE<SP>#3Property<SP>Type:Select<SP>a<SP>Value>>""Commercial<SP>Land""""MH<SP>Fixed""2*");
//macro.AppendLine(@"TAG POS=4 TYPE=TD FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=ID:tdCompOneBasementDetail");
//macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneBasementBathrooms CONTENT=0");
//macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneBasementBathroomsHalf CONTENT=0");
//macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneBasementOtherRooms CONTENT=0");
//macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneBasmentSquareFeet CONTENT=864");
//macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneBasmentFinishedSquareFeet CONTENT=0");
//macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneBathrooms CONTENT=1");
//macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:ImportFormRender.aspx?ImportFormID=2308912 ATTR=NAME:ifcContentCompOneBathroomsHalf CONTENT=1");
//string macroCode = macro.ToString();
//// status = iim.iimPlayCode(macroCode, timeout);

    }
}
