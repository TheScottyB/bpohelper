using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{

    //namespace solutionstar
    //{

    //     class WebFormFieldNames
    //    {

    //         public enum Groups {OriginalListing, PropertyAddress }
    //        public enum PropertyAddressFields { Address, Ctiy, Zip }
    //        public enum OriginalListingFields { Date, Price }
    //        public enum CommonFields { Datasource, DOM}
    //    }
    //}


    class SolutionStar:BPOFulfillment
    {

         private Dictionary<string, MacroParameterValues> NeighborhoodFieldList = new Dictionary<string, MacroParameterValues>();
         public enum ImacroType { CHECKBOX, TEXT, SELECT}
         public Dictionary<ImacroType, string> imacroTypeToStringAdapter = new Dictionary<ImacroType, string> {
            {ImacroType.CHECKBOX, "INPUT:CHECKBOX"},
            {ImacroType.SELECT, "SELECT"},
            {ImacroType.TEXT, "INPUT:TEXT"}
         };

        struct MacroParameterValues
        {
            string imacroCommandType;


      
        }

        struct WebFormFieldNames
        {

            public enum Groups { PropertyAddress, OriginalListing, CurrentListing, LastSale, LandAndStructure, GeneralProperties, SalesProperties, LivingArea, Basement, SquareFeet, Amenities, Other, HOA, Fee, IsHOA, PreviousListings, Predominant, Construction, Listings }
            public enum PriceTrendFields { Direction }
            public enum ListingsFields { NeighborhoodListings }
            public enum ConstructionFields { LowPrice, HighPrice }
            public enum PredominantFields { PriceTrend, Occupancy, PricesFrom, PricesTo }
            public enum PreviousListingsFields { PreviouslyListed }
            public enum PropertyAddressFields { Address, City, Zip }
            public enum OriginalListingFields { Date, Price, Price_two }
            public enum CurrentListingFields { Date, Price, Price_two, Concessions }
            public enum LivingAreaFields { Gross }
            public enum LandAndStructureFields { Age, LotSize, Site, TotalRooms, BedRooms, FullRooms, PartialRooms, DesignStyle, View, NumberOfUnits, ParkingStalls, OverallComparability, SourceOfFunds, DOM_two }
            public enum BasementFields { TotalArea, BasementType }
            public enum SquareFeetFields { Finished }
            public enum AmenitiesFields { Pool }
            public enum OtherFields { type }
            public enum GeneralPropertiesFields { IsHOA, HOA }
            public enum HOAFields { Fee }
            public enum FeeFields { Amount_two, Amount }
            public enum IsHOAFields { feesPerMonth }
            public enum CommonFields { Datasource, DOM, AverageMarketTime, ComparisonToSubject, Garage, DistressedSale, MedianMarketRent, REOPercentage, MarketTimingTrend, SourceOfFunds, Condition, APN_TaxId, OwnerOfPublicRecords, PropertyType, Construction, ImpactedByDiaster, EvidenceForDiaster, NewConstruction, Occupancy, Inspection, InspectionDate, Vacancy, CurrentListing, Location, Industrydistance, ComparableListingSupply }
            public enum InputTypes { TEXT, SELECT}
            public enum ConditonSelections { Excellent = 1, Good = 2, Average = 3, Fair = 4, Poor = 5}
            public enum BasementTypes {NA, None, Crawl_Space, Partial, Full}
    
        }

        struct WebFormSelectionBoxes
        {
            public enum AverageMarketTimeSelections { Months_0to3 = 1, Months_3to6 = 2, Months_6to12 = 3, Over_6_months = 4, Over_12_months = 5 }

        }


        struct WebCheckBoxGroups
        {
            public enum PropertyTypes { SFR=1, Condo=2,Co_op=3, PUD=4, Manufactured=5, Other=6}
            public enum YesNo { Yes = 1, No = 2}
            public enum OccupancyTypes { Owner = 1, Tenant = 2, Vacant = 3, Unknown = 4 }
            public enum InspectionTypes { Exterior = 1, Interior = 2 }
            public enum CurrentListingTypes { Yes_Currently = 1, Yes_in_Last_12_Months = 2, No = 3}
            public enum LocationTypes { Urban = 1, Suburban = 2, Rural = 3 }
            public enum PropertyValues { Increasing = 1, Stable = 2, Declining = 3 }
            public enum PredominantOccupancy { Owner = 1, Tenant = 2 }
            public enum VacancyTypes { percent0_5 = 1, percent5_10 = 2, percent10_20 = 3, precent20_100 = 4 }
        }

       private string saleCompText = "Sales.Sale";
       private string listingCompText = "Listings.Listing";
      
      // private enum WebFormFieldNames { DataSource, ProximityToSubject, PropertyAddress };
 
        //{0} is now TEXT, TEXTAREA,FILE, RADIO, CHECKBOX
       //{1} is now either Comparables. or Subject
       //{2} is now either Sales.Sale, Listings.Listing or "" (emptystring)
       //{3} is compNum or  "" (emptystring)
        //{4} fieldname
        //{5} is the value
       private string theBaseImacroCommand = "TAG POS={0} TYPE={1} FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.*.{2}{3}{4}.{5} CONTENT={6}\r\n";

          public SolutionStar()
        {
         
              
            compFieldListTranslator["NumberOfUnits"] = "units";
            compFieldListTranslator["DataSource"] = "datasource";
            compFieldListTranslator["GLA"] = "sqft";
            compFieldListTranslator["TotalRoomCount"] = "rooms";
            compFieldListTranslator["BedroomCount"] = "bed";
            compFieldListTranslator["TotalRoomCount"] = "rooms";

            compFieldListTranslator["YearBuilt"] = "yr_built";
            compFieldListTranslator["LotSize"] = "lotsize";
            compFieldListTranslator["MainView"] = "view";

            
     
            compFieldListTranslator["DOM"] = "dom";
            compFieldListTranslator["SalePrice"] = "sale_price";
            compFieldListTranslator["SalesDate"] = "sale_dt";


              


            compFieldListTranslator["NumberOfParkingSpaces"] = "garage";

            compFieldListTranslator.Add("FullBathCount", "bath");
            compFieldListTranslator.Add("HalfBathCount", "partial_bath");
            compFieldListTranslator.Add("OverallPropertyCondition", "condition");
            compFieldListTranslator.Add("OverallLocationDensity", "location");
            compFieldListTranslator.Add("BuildingSkeleton", "construction");
            compFieldListTranslator.Add("WaterFront", "waterfront");
            compFieldListTranslator.Add("OriginalListPrice", "original_lp");
            compFieldListTranslator.Add("NumberOfPriceReductions", "pricereductions");
          

        }


          public SolutionStar(MLSListing m) 
            : this()
        {
            targetComp = m;
        }
        
        
        private Dictionary<string, string> propTypeTranslator = new Dictionary<string, string>()
         {
            {"Detached", "Single Family"},
            {"Attached", "Condo"}
           

         };

        private Dictionary<string, string> subjectFieldListTranslator = new Dictionary<string, string>()
        {
            {"ParcelID", "sub_apn"}, 
            {"County", "County"},
            {"PropertyType", "*property_type"},
            {"Rent", "fair_mkt_rent"}
            //{"DR", "Drive"},
            //{"HWY", "Highway"},
            //{"LN", "Lane"},
            //{"PKWY", "Parkway"},
            //{"PL", "Place"},
            //{"PLZ", "Plaza"},
            //{"PL", "Point"},
            //{"PT", "Place"},
            //{"RD", "Road"},
            //{"SQ", "Square"},
            //{"ST", "Road"},
            //{"TER", "Terrace"},
            //{"TRL", "Trail"},
            //{"WAY", "Way"}
           
        };

         private Dictionary<string, string> subjectFieldList = new Dictionary<string, string>();


        private void helper_SetCommonFields()
        {


           
            compSelectionBoxList.Add(compFieldListTranslator["TransactionType"], "%Arm");

            //
            //Selection Boxes
            //
            compSelectionBoxList.Add(compFieldListTranslator["DataSource"], "%M");
            compSelectionBoxList.Add(compFieldListTranslator["NumberOfUnits"], "%1");
            compSelectionBoxList.Add(compFieldListTranslator["OverallPropertyCondition"], "%A");
            compSelectionBoxList.Add(compFieldListTranslator["OverallLocationDensity"], "%S");
            compSelectionBoxList.Add(compFieldListTranslator["BuildingSkeleton"], "%F");
            compSelectionBoxList.Add(compFieldListTranslator["MainView"], "%R");
            compSelectionBoxList.Add(compFieldListTranslator["WaterFront"], "%N");
            if (targetComp.NumberOfPriceChanges >= 5)
            {
                compSelectionBoxList.Add(compFieldListTranslator["NumberOfPriceReductions"], "%5");
            }
            else
            {
                compSelectionBoxList.Add(compFieldListTranslator["NumberOfPriceReductions"], "%" + targetComp.NumberOfPriceChanges.ToString());
            }
        

            compSelectionBoxList.Add("lotsize_unit", "%A");
            compSelectionBoxList.Add("gated", "%N");
            compSelectionBoxList.Add("pool", "%N");
            compSelectionBoxList.Add("parking", "%D");
            compSelectionBoxList.Add("designAppeal", "%E");


            //
            //Text Fields
            //
            //main characteristics
            compTextFieldList.Add(compFieldListTranslator["GLA"], targetComp.GLA.ToString());
            compTextFieldList.Add(compFieldListTranslator["NumberOfUnits"], "1");
            compTextFieldList.Add(compFieldListTranslator["YearBuilt"], targetComp.YearBuilt.ToString());
            compTextFieldList.Add(compFieldListTranslator["LotSize"], targetComp.Lotsize.ToString());
            //rooms
            compTextFieldList.Add(compFieldListTranslator["TotalRoomCount"], targetComp.TotalRoomCount.ToString());
            compTextFieldList.Add(compFieldListTranslator["BedroomCount"], targetComp.BedroomCount);
            compTextFieldList.Add(compFieldListTranslator["FullBathCount"], targetComp.FullBathCount);
            compTextFieldList.Add(compFieldListTranslator["HalfBathCount"], targetComp.HalfBathCount);
            //garage
            compTextFieldList.Add(compFieldListTranslator["NumberOfParkingSpaces"], targetComp.NumberGarageStalls());
            //pricing and listing history
            compTextFieldList.Add(compFieldListTranslator["OriginalListPrice"], targetComp.OriginalListPrice.ToString());
            
            compTextFieldList.Add(compFieldListTranslator["DOM"], targetComp.DOM);
            //Misc
            compTextFieldList.Add("energyeff", "NA");
            compTextFieldList.Add("other", "NA");
            compTextFieldList.Add("heat_cool", @"FA/CA");
            compTextFieldList.Add("concessions", "NA");


        }


      
        protected override string  GenerateSubjectFillScript()
        {
            StringBuilder macro = new StringBuilder();
            //borrowers name (already filled)
            //#APN
            subjectFieldList.Add(subjectFieldListTranslator["ParcelID"], GlobalVar.theSubjectProperty.ParcelID);

            //County
            subjectFieldList.Add(subjectFieldListTranslator["County"], GlobalVar.theSubjectProperty.County);

            //PropertyType (selection box)
                //Single Family
                //Condo
                //*bunch of others* but we don't care

            //FairMArketRent
            //subjectFieldList.Add(subjectFieldListTranslator["Rent"], GlobalVar.theSubjectProperty);
            //Secure (Yes/No)
            //Occupancy (Yes/No)
            //ZoningCode 
            //ZoningDescription
            //ZoningCompliance
            //IllegalUnits
                //if yes, decription
            //currentUse
            //bestUse
                //if no, describe
            //currentlyListed (Yes/No)
            //listedLast36Months
                //listing status, if either of the above
                //
                //(current listing section)
                //
            //set of red flag checkboxes
                //damaged
                //contruction
                //environmental
                //zoning
                //market activity
                //boarded
                //stigma
                //other
                //none
                    //if anything checked, comments

          //  TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$sub_apn CONTENT=ttt
            foreach (string field in subjectFieldList.Keys)
            {
                if (field.Contains("*"))
                {

                }
                else
                {
                    //orignal way using * instead of C
                    //macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:*{0}{1}_{2} CONTENT={3}\r\n", sol, field, Regex.Match(compNum, @"\d").Value, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                    macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21${0} CONTENT={1}\r\n", field, subjectFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                }



             
            }


            return macro.ToString();


        }

        protected override void GenerateSaleCompFillScript()
        {
            helper_CompFillHeaderScript();
            helper_SetCommonFields();

            //fields and/or changes specific to SALE comps
            compFieldListTranslator["CurrentListPrice"] = "sale_lp";
            compTextFieldList.Add(compFieldListTranslator["CurrentListPrice"], targetComp.CurrentListPrice.ToString());
            compTextFieldList.Add(compFieldListTranslator["SalePrice"], targetComp.SalePrice.ToString());
            compTextFieldList.Add(compFieldListTranslator["SalesDate"], targetComp.SalesDate.ToShortDateString());

            foreach (string field in compTextFieldList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:*s{0}_{1} CONTENT={2}\r\n", targetCompNumber, field, compTextFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compSelectionBoxList.Keys)
            {                      //"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control* ATTR=NAME:*{0}{1} CONTENT={2}");
                theMacro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:*s{0}_{1} CONTENT={2}\r\n", targetCompNumber, field, compSelectionBoxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }


            foreach (string field in compRadioButtonList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ACTION:/Order/OrderEdit* ATTR=NAME:SalesComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compRadioButtonList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compCheckboxList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEdit* ATTR=NAME:SalesComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compCheckboxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:*S"+ targetCompNumber + "OrigListDt CONTENT=" + targetComp.ListDateString);


            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEditWizardFNMA/step3/11488927 ATTR=NAME:SalesComp1.HasForcedWarmAirForHeat CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEditWizardFNMA/step3/11488927 ATTR=NAME:SalesComp1.HasCentralAir CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=ACTION:/Order/OrderEditWizardFNMA/step3/11488927 ATTR=CLASS:mceIframeContainer<SP>mceFirst<SP>mceLast");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:/Order/OrderEditWizardFNMA/step3/11488927 ATTR=NAME:SaveButton");  
        }

        protected override void GenerateListCompFillScript()
        {
            helper_CompFillHeaderScript();
            helper_SetCommonFields();

            //fields and/or changes specific to LIST comps
            compFieldListTranslator["CurrentListPrice"] = "current_lp";
            compTextFieldList.Add(compFieldListTranslator["CurrentListPrice"], targetComp.CurrentListPrice.ToString());
            compTextFieldList.Add("list_dt", targetComp.ListDate.ToShortDateString());


            foreach (string field in compTextFieldList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXTFORM=NAME:aspnetForm ATTR=NAME:*l{0}_{1} CONTENT={2}\r\n", targetCompNumber, field, compTextFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compSelectionBoxList.Keys)
            {                      //"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control* ATTR=NAME:*{0}{1} CONTENT={2}");
                theMacro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:*l{0}_{1} CONTENT={2}\r\n", targetCompNumber, field, compSelectionBoxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }


            foreach (string field in compRadioButtonList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ACTION:/Order/OrderEdit* ATTR=NAME:SalesComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compRadioButtonList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compCheckboxList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEdit* ATTR=NAME:SalesComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compCheckboxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }


        }

        protected void GenerateCompFillScript(string webformCompName)
        {
            //{0} is now TEXT, TEXTAREA,FILE, RADIO, CHECKBOX
       //{1} is now either Comparables. or Subject
       //{2} is now either Sales.Sale, Listings.Listing or "" (emptystring)
       //{3} is compNum or  "" (emptystring)
        //{4} fieldname
        //{5} is the value
       //private string theBaseImacroCommand = "TAG POS=1 TYPE=INPUT:{0} FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.{1}{2}{3}.{4} CONTENT={5}\r\n";

            compTextFieldList.Add("PropertyAddress.Address",targetComp.StreetAddress);
            compTextFieldList.Add("PropertyAddress.City",  targetComp.City);
            compTextFieldList.Add("PropertyAddress.Zip", targetComp.Zipcode);
            compTextFieldList.Add("DataSource", "MLS");
            compTextFieldList.Add("ProximityToSubject", targetComp.ProximityToSubject.ToString());
            compTextFieldList.Add("SalesProperties.OriginalListing.Date", targetComp.ListDateString);
            compTextFieldList.Add("SalesProperties.OriginalListing.Price", targetComp.OriginalListPrice.ToString());
            compTextFieldList.Add("SalesProperties.CurrentListing.Price", targetComp.CurrentListPrice.ToString());
            compTextFieldList.Add("DOM", targetComp.DOM);
            compTextFieldList.Add("SourceOfFunds", targetComp.FinancingMlsString);
            compTextFieldList.Add("SalesProperties.CurrentListing.Concessions", targetComp.PointsMlsString);
            compTextFieldList.Add("DistressedSale", targetComp.DistressedSaleYesNo());
            compTextFieldList.Add("GeneralProperties.IsHOA.feesPerMonth", "0");
            compTextFieldList.Add("LandAndStructure.LivingArea.Gross", targetComp.GLA.ToString());
            compTextFieldList.Add("LandAndStructure.Age", targetComp.Age.ToString());
            compTextFieldList.Add("LandAndStructure.LotSize", targetComp.Lotsize.ToString());
            compTextFieldList.Add("LandAndStructure.Site", "1");
            compTextFieldList.Add("LandAndStructure.TotalRooms", targetComp.TotalRoomCount.ToString());
            compTextFieldList.Add("LandAndStructure.Bedrooms", targetComp.BedroomCount);
            compTextFieldList.Add("LandAndStructure.FullRooms", targetComp.FullBathCount.ToString());
            compTextFieldList.Add("LandAndStructure.PartialRooms", targetComp.HalfBathCount.ToString());
            compTextFieldList.Add("LandAndStructure.DesignStyle", "Typical");
            compTextFieldList.Add("LandAndStructure.View", "Residential");
            compTextFieldList.Add("ComparisonToSubject", "Same");
            compTextFieldList.Add("LandAndStructure.NumberOfUnits", "1");
            compTextFieldList.Add("LandAndStructure.Basement.TotalArea", targetComp.BasementGLA());
            compTextFieldList.Add("LandAndStructure.SquareFeet.Finished", targetComp.BasementFinishedGLA());
            compTextFieldList.Add("Garage", targetComp.GarageString());
            compTextFieldList.Add("LandAndStructure.ParkingStalls", targetComp.NumberGarageStalls());
            compTextFieldList.Add("Amenities.Pool", "NA");
            compTextFieldList.Add("Amenities.Other.type", "NA");
            compTextFieldList.Add("LandAndStructure.OverallComparability", "Equal");



            if (webformCompName == "Sales.Sale")
            {
                compTextFieldList.Add("SalesProperties.LastSale.Date", targetComp.SalesDate.ToShortDateString());
                compTextFieldList.Add("SalesProperties.LastSale.Price", targetComp.SalePrice.ToString());


            }



            foreach (string field in compTextFieldList.Keys)
            {
                //"TAG POS={0} TYPE={1} FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.*.{2}{3}{4}.{5} CONTENT={6}\r\n";
                theMacro.AppendFormat(theBaseImacroCommand, "1", "INPUT:TEXT", "Comparables.", webformCompName, targetCompNumber, field, compTextFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            compTextFieldList.Clear();
            int x = (int)WebFormFieldNames.ConditonSelections.Average;
            compTextFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.Groups.Basement, WebFormFieldNames.BasementFields.BasementType), helper_BasementTypeTranslate(targetComp.BasementType));
            compTextFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Condition), x.ToString());

            foreach (string field in compTextFieldList.Keys)
            {
                theMacro.AppendFormat(theBaseImacroCommand, "1", "SELECT", "Comparables.", webformCompName, targetCompNumber, field.Replace("_", "<SP>"), "%" + compTextFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }
           

        }

        protected string helper_MakeCommandString(object param1)
        {
            return param1.ToString();
        }
        protected string helper_MakeCommandString(object param1, object param2)
        {
            return param1.ToString() + "." + param2.ToString();
        }

        protected string helper_MakeCommandString(object param1, object param2, object param3)
        {
            return param1.ToString() + "." + param2.ToString() + "." + param3.ToString();
        }

        protected string helper_MakeCommandString(object param1, object param2, object param3, object param4)
        {
            return param1.ToString() + "." + param2.ToString() + "." + param3.ToString() + "." + param4.ToString();
        }

        private string helper_BasementTypeTranslate(string ourType)
        {
            string theirType = WebFormFieldNames.BasementTypes.NA.ToString();

            if (ourType.ToLower().Contains("full"))
                return WebFormFieldNames.BasementTypes.Full.ToString();
            if (ourType.ToLower().Contains("partail"))
                return WebFormFieldNames.BasementTypes.Partial.ToString();
            if (ourType.ToLower().Contains("crawl"))
                return WebFormFieldNames.BasementTypes.Crawl_Space.ToString();
            if (ourType.ToLower().Contains("none"))
                return WebFormFieldNames.BasementTypes.None.ToString();
            return theirType;
        }

       

        public void Prefill(iMacros.App iim, Form1 form)
        {
            
            subjectFieldList.Add(WebFormFieldNames.CommonFields.Datasource.ToString(), "Public Records");
            subjectFieldList.Add(WebFormFieldNames.Groups.PropertyAddress.ToString() + "." + WebFormFieldNames.PropertyAddressFields.Zip.ToString(), GlobalVar.theSubjectProperty.Zip);
            subjectFieldList.Add(WebFormFieldNames.Groups.PropertyAddress.ToString() + "." + WebFormFieldNames.PropertyAddressFields.Address.ToString(), GlobalVar.theSubjectProperty.AddressLine1);
            subjectFieldList.Add(WebFormFieldNames.Groups.PropertyAddress.ToString() + "." + WebFormFieldNames.PropertyAddressFields.City.ToString(), GlobalVar.theSubjectProperty.City);


            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.APN_TaxId), GlobalVar.theSubjectProperty.ParcelID);
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.OwnerOfPublicRecords), GlobalVar.theSubjectProperty.MainForm.SubjectOOR);
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.Groups.LivingArea, WebFormFieldNames.LivingAreaFields.Gross), GlobalVar.theSubjectProperty.MainForm.SubjectAboveGLA);
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.Age ), GlobalVar.theSubjectProperty.Age.ToString());
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.LotSize), GlobalVar.theSubjectProperty.MainForm.SubjectLotSize);
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.Site), "1");
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.TotalRooms), GlobalVar.theSubjectProperty.MainForm.SubjectRoomCount);
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.BedRooms), GlobalVar.theSubjectProperty.MainForm.SubjectBedroomCount);
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.FullRooms), GlobalVar.theSubjectProperty.FullBathCount);
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.PartialRooms), GlobalVar.theSubjectProperty.HalfBathCount);
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.DesignStyle), "Typical");
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.View), "Residential");
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.NumberOfUnits), "1");
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.Groups.Basement, WebFormFieldNames.BasementFields.TotalArea), GlobalVar.theSubjectProperty.MainForm.SubjectBasementGLA);
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.Groups.SquareFeet, WebFormFieldNames.SquareFeetFields.Finished), GlobalVar.theSubjectProperty.MainForm.SubjectBasementFinishedGLA);
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Garage), GlobalVar.theSubjectProperty.MainForm.SubjectParkingType);
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.ParkingStalls), GlobalVar.theSubjectProperty.GarageStallCount);
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.Amenities, WebFormFieldNames.AmenitiesFields.Pool), "Unk");
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.Amenities, WebFormFieldNames.Groups.Other, WebFormFieldNames.OtherFields.type), "Unk");
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.InspectionDate), GlobalVar.theSubjectProperty.InspectionDate());
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.SourceOfFunds), "NA");
            if (!String.IsNullOrWhiteSpace(GlobalVar.theSubjectProperty.MainForm.SubjectAssessmentInfo.amount))
            {
                subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.GeneralProperties, WebFormFieldNames.GeneralPropertiesFields.HOA, WebFormFieldNames.HOAFields.Fee, WebFormFieldNames.FeeFields.Amount ), GlobalVar.theSubjectProperty.MainForm.SubjectAssessmentInfo.amount);
            }


            foreach (string field in subjectFieldList.Keys)
            {
                theMacro.AppendFormat(theBaseImacroCommand, "1", "INPUT:TEXT", "Subject", "", "", field, subjectFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            subjectFieldList.Clear();
            int x = (int)WebFormFieldNames.ConditonSelections.Average;
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.Groups.Basement, WebFormFieldNames.BasementFields.BasementType), helper_BasementTypeTranslate(GlobalVar.theSubjectProperty.MainForm.SubjectBasementType));
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Condition), x.ToString());
            
            foreach (string field in subjectFieldList.Keys)
            {
                theMacro.AppendFormat(theBaseImacroCommand, "1", "SELECT", "Subject", "", "", field.Replace("_", "<SP>"), "%" + subjectFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            subjectFieldList.Clear();
            if (GlobalVar.theSubjectProperty.MainForm.SubjectAttached)
            {
                subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.PropertyType), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.PropertyTypes.Condo));
                theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.NormalSale.Repaired.SubjectLandValue CONTENT=0");
     

            }
            else
            {
                subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.PropertyType), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.PropertyTypes.SFR));
                theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.NormalSale.Repaired.SubjectLandValue CONTENT=10000");
            }
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Construction), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.YesNo.No));
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.ImpactedByDiaster), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.YesNo.No));
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Occupancy), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.OccupancyTypes.Unknown));
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Inspection), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.InspectionTypes.Exterior));
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.CurrentListing), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.CurrentListingTypes.No));
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.SalesProperties, WebFormFieldNames.Groups.PreviousListings, WebFormFieldNames.PreviousListingsFields.PreviouslyListed), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.YesNo.No));
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Location), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.LocationTypes.Suburban));
            foreach (string field in subjectFieldList.Keys)
            {
                theMacro.AppendFormat(theBaseImacroCommand, subjectFieldList[field], "INPUT:CHECKBOX", "Subject", "", "", field.Replace("_", "<SP>"), "YES");
            }

            //
            //Neighborhood Checkboxes
            //
            subjectFieldList.Clear();
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.Predominant, WebFormFieldNames.PredominantFields.PriceTrend, WebFormFieldNames.PriceTrendFields.Direction), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.PropertyValues.Stable));
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.Predominant, WebFormFieldNames.PredominantFields.Occupancy), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.PredominantOccupancy.Owner));
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Industrydistance), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.YesNo.No));
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Vacancy), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.VacancyTypes.percent5_10));
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.NewConstruction), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.YesNo.No));
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.EvidenceForDiaster), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.YesNo.No));
            foreach (string field in subjectFieldList.Keys)
            {
                theMacro.AppendFormat(theBaseImacroCommand, subjectFieldList[field], "INPUT:CHECKBOX", "Neighborhood", "", "", field.Replace("_", "<SP>"), "YES");
            }

            //
            //Neighborhood TEXT
            //
            subjectFieldList.Clear();
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.MedianMarketRent), GlobalVar.theSubjectProperty.MainForm.SubjectRent);
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.REOPercentage), "Stable");
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.MarketTimingTrend), "Stable");
            foreach (string field in subjectFieldList.Keys)
            {
                theMacro.AppendFormat(theBaseImacroCommand, "1", "INPUT:TEXT", "Neighborhood", "", "", field, subjectFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            //
            //Neighborhood SELECT
            //
            subjectFieldList.Clear();
            subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.AverageMarketTime), "0-3 Months");
            foreach (string field in subjectFieldList.Keys)
            {
                theMacro.AppendFormat(theBaseImacroCommand, "1", "SELECT", "Neighborhood", "", "", field.Replace("_", "<SP>"), "%" + subjectFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            //
            //Neighborhood stats
            //
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Neighborhood.Listings.NeighborhoodListings CONTENT=" + form.SubjectNeighborhood.numberActiveListings);
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Neighborhood.Predominant.PricesFrom CONTENT=" + form.SubjectNeighborhood.minListPrice);
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Neighborhood.Predominant.PricesTo CONTENT=" + form.SubjectNeighborhood.maxListPrice);
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Neighborhood.ComparableListingSupply CONTENT=" + form.SubjectNeighborhood.numberSoldListings);
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Neighborhood.Construction.LowPrice CONTENT=" + form.SubjectNeighborhood.minSalePrice);
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Neighborhood.Construction.HighPrice CONTENT=" + form.SubjectNeighborhood.maxSalePrice);


            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.NormalSale.ASIS.MarketValue CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectMarketValue);
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.NormalSale.ASIS.SuggestedListPrice CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectMarketValueList);
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.NormalSale.Repaired.MarketValue CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectMarketValue);
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.NormalSale.Repaired.SuggestedListPrice CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectMarketValueList);

            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.QuickSale.ASIS.MarketValue CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectQuickSaleValue);
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.QuickSale.ASIS.SuggestedListPrice CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectQuickSaleValue);
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.QuickSale.Repaired.MarketValue CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectQuickSaleValue);
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.QuickSale.Repaired.SuggestedListPrice CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectQuickSaleValue);
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.NormalSale.ASIS.FairMarketRent CONTENT=" + GlobalVar.theSubjectProperty.MainForm.RoundTo50(GlobalVar.theSubjectProperty.MainForm.SubjectRent)).ToString();

            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:bpomainform ATTR=NAME:signaturefile CONTENT=" + GlobalVar.theSubjectProperty.MainForm.DropBoxFolder +   "\\BPOs\\Dawn-sig.JPG");
      
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:bpomainform ATTR=NAME:signatureupload");

            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Vendor.FullName CONTENT=Dawn<SP>Zurick");
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Vendor.License.Date CONTENT=" + DateTime.Now.ToShortDateString());
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Vendor.License.Number CONTENT=471.0096163");
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Vendor.License.State CONTENT=IL");
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Vendor.Company.Name CONTENT=OKRP");
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Vendor.PhoneNo CONTENT=815-315-0203");
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Vendor.Address.Address CONTENT=10325<SP>Main");

            theMacro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Reconciliation.FinalComments CONTENT=This<SP>evaluation<SP>was<SP>prepared<SP>by<SP>a<SP>licensed<SP>real<SP>estate<SP>broker<SP>and<SP>is<SP>not<SP>an<SP>appraisal.<SP>This<SP>evaluation<SP>cannot<SP>be<SP>used<SP>for<SP>the<SP>purpose<SP>of<SP>obtaining<SP>financing.<BR><LF>Notwithstanding<SP>any<SP>preprinted<SP>language<SP>to<SP>the<SP>contrary,<SP>this<SP>is<SP>not<SP>an<SP>appraisal<SP>of<SP>the<SP>market<SP>value<SP>of<SP>the<SP>property.<SP>If<SP>an<SP>appraisal<SP>is<SP>desired,<SP>the<SP>services<SP>of<SP>a<SP>licensed<SP>or<SP>certified<SP>appraiser<SP>must<SP>be<SP>obtained.");


       
            WriteScript(GlobalVar.theSubjectProperty.MainForm.SubjectFilePath, "prefill.iim", theMacro);

            string macroCode = theMacro.ToString();
            iim.iimPlayCode(macroCode, 60);
        }

        private string helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.VacancyTypes vacancyTypes)
        {
            return ((int)vacancyTypes).ToString();
        }

        private string helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.PredominantOccupancy predominantOccupancy)
        {
            return ((int)predominantOccupancy).ToString();
        }

        private string helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.PropertyValues property_Values)
        {
            return ((int)property_Values).ToString();
        }

        private string helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.LocationTypes locationTypes)
        {
            return ((int)locationTypes).ToString();
        }

        private string helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.CurrentListingTypes currentListingTypes)
        {
            return ((int)currentListingTypes).ToString();
        }

        private string helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.InspectionTypes inspectionTypes)
        {
            return ((int)inspectionTypes).ToString();
        }

        private string helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.OccupancyTypes occupancyTypes)
        {
            return ((int)occupancyTypes).ToString();
        }

        private string helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.YesNo yesNo)
        {
            return ((int)yesNo).ToString();
        }

        private string helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.PropertyTypes propertyTypes)
        {
            return ((int)propertyTypes).ToString();
        }

        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {
           

            targetCompNumber = Regex.Match(compNum, @"\d").Value;

              if (saleOrList == "sale")
            {

                GenerateCompFillScript(saleCompText);

            }
            else
            {

                GenerateCompFillScript(listingCompText);

            }
   
           

            WriteScript(fieldList["filepath"], compNum + ".iim", theMacro);

            string macroCode = theMacro.ToString();
            iim.iimPlayCode(macroCode, 60);
        }
    }
}
