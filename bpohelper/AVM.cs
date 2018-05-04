using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{



    class AVM : BPOFulfillment
    {
/// <summary>
/// </summary>
/// <param name="iim"></param>
/// <param name="form"></param>
/// http://avm.assetval.com/AVM/Realtor/PendingAssignments.aspx

      
        
          public AVM()
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

            compFieldListTranslator["ProximityToSubject"] = "prox";
     
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


        public AVM(MLSListing m) 
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


        private void helper_SetAddressFields()
         {
             int positionNumber = 0;
            
            int.TryParse(targetCompNumber,  out positionNumber);

            if (saleOrListFlag == "list")
            {
                positionNumber = positionNumber + 3;
            }

             theMacro.AppendLine(@"TAG POS=" + positionNumber.ToString() + " TYPE=INPUT:BUTTON FORM=NAME:aspnetForm ATTR=CLASS:btnBlue");
             theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$txtCompAddress CONTENT=" + targetComp.StreetAddress.Replace(" ", "<SP>"));
             theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$txtCompCity CONTENT=" + targetComp.City.Replace(" ", "<SP>"));
             theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$txtCompZip CONTENT=" + targetComp.Zipcode.Replace(" ", "<SP>"));
             theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$btnGetProximity");
             theMacro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
             theMacro.AppendLine(@"WAIT SECONDS=1");
             theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$btnAcceptProximity");
      
         }

        private void helper_SetCommonFields()
        {

            Dictionary<string, string> financingTypeTranslation = new Dictionary<string, string>()
            {
                {"Conventional", @"%Conventional"}, {"FHA", "%FHA/VA"}, {"VA", "%FHA/VA"},  {"Unknown", "%UNKNOWN"}, {"Cash", "%CASH"}

            };

            Dictionary<string, string> saleTypeTranslation = new Dictionary<string, string>()
            {
                {"Arms Length", "%OWNER"}, {"REO", "%REO"}, {"ShortSale", "%SHORT"}, {"Unknown", "%UNKNOWN"}
            };

            Dictionary<string, string> propertyTypeTranslation = new Dictionary<string, string>()
            {
                {"Detached", "%SFD"}, {"Attached", "%CO"}
            };

            Dictionary<string, string> buildingTypeTranslation = new Dictionary<string, string>()
            {
                {"1 Story", "%R"},  {"1.5 Story", "%BG"}, {"2 Stories", "%CT"}, {"Hillside", "%SP"}, {"Raised Ranch", "%SP"}, {@"Split Level w/ Sub", "%SP"}, {"Split Level", "%SP"},
                    {"Condo", "%LR"}, {"Townhouse-2 Story", "%T"}, {"Townhouse", "2story"},  {"Townhouse 3+ Stories", "2story"}, {"Townhouse-TriLevel", "2story"}, {"Townhouse‐Ranch", "1story"}  //Townhouse 3+ Stories, Townhouse-TriLevel
            };

            helper_SetAddressFields();
            
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
            compSelectionBoxList.Add(compFieldListTranslator["WaterFront"], "%" + targetComp.WaterFrontYesNo()[0].ToString());
            if (targetComp.NumberOfPriceChanges >= 5)
            {
                compSelectionBoxList.Add(compFieldListTranslator["NumberOfPriceReductions"], "%5");
            }
            else
            {
                compSelectionBoxList.Add(compFieldListTranslator["NumberOfPriceReductions"], "%" + targetComp.NumberOfPriceChanges.ToString());
            }


            //
            //04/28/18 updates
            //
            //cs1_garage_carport
            if (targetComp.AttachedGarage())
            {
                compSelectionBoxList.Add("garage_carport", "%A");
            }
            else if (targetComp.DetachedGarage())
            {
                compSelectionBoxList.Add("garage_carport", "%D");
            }
            else
            {
                compSelectionBoxList.Add("garage_carport", "%N");
            }

            //ddlCS1PropertyType
            theMacro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:*" + saleOrListFlag.ToUpper()[0] + targetCompNumber + "PropertyType CONTENT=" + propertyTypeTranslation[targetComp.UniversalLandUse()]);

            //cs1_style
            compSelectionBoxList.Add("style", buildingTypeTranslation[targetComp.Type]);

            //ddlCS1ViewComparision
            theMacro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:*" + saleOrListFlag.ToUpper()[0] + targetCompNumber + "ViewComparision CONTENT=" + "%E");

            //cs1_wastedisp
            compSelectionBoxList.Add("wastedisp", "%P");

            //cs1_watersource
            compSelectionBoxList.Add("watersource", "%P");

            //cs1_fireplace
            compSelectionBoxList.Add("fireplace", "%" + targetComp.NumberOfFireplaces );

            //cs1_owner
            compSelectionBoxList.Add("owner", saleTypeTranslation[targetComp.TransactionType]);

            //cs1_fin_type
            compSelectionBoxList.Add("fin_type", financingTypeTranslation[targetComp.FinancingMlsString]);

            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:aspnetForm ATTR=NAME:*2" + saleOrListFlag[0] + "_most_comparable_yn CONTENT=YES");

            //txtCS1HOAMonthlyFee
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:*" + saleOrListFlag.ToUpper()[0] + targetCompNumber + "HOAMonthlyFee CONTENT=0");

            //_adjustments
            theMacro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:aspnetForm ATTR=NAME:*" + saleOrListFlag[0] + targetCompNumber + "_adjustments CONTENT=" + "Similar size, age, style, features, condition and neighborhood as subject.".Replace(" ", "<SP>"));

            //
            //
            //

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
            compTextFieldList.Add(compFieldListTranslator["ProximityToSubject"], targetComp.ProximityToSubject.ToString());
            //basement
            if (targetComp.BasementType.Contains("Full"))
                compSelectionBoxList.Add("bsmt_fin_yn", "%F");
            if (targetComp.BasementType.Contains("Partial"))
                compSelectionBoxList.Add("bsmt_fin_yn", "%PF");

            if (targetComp.BasementType.Contains("None"))
                compSelectionBoxList.Add("bsmt_fin_yn", "%N");
            compTextFieldList.Add("bsmt_fin_per", targetComp.BasementFinishedPercentage());

            compTextFieldList.Add("sqft_bg", targetComp.BasementGLA());
            compTextFieldList.Add("rooms_bg", targetComp.NumberOfBasementRooms());
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

          //  TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_apn CONTENT=ttt
            foreach (string field in subjectFieldList.Keys)
            {
                if (field.Contains("*"))
                {

                }
                else
                {
                    //orignal way using * instead of C
                    //macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:*{0}{1}_{2} CONTENT={3}\r\n", sol, field, Regex.Match(compNum, @"\d").Value, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                    macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*${0} CONTENT={1}\r\n", field, subjectFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
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



        public void Prefill(iMacros.App iim, Form1 form)
        {
            
            StringBuilder macro = new StringBuilder();
            StringBuilder gotoMarketTab = new StringBuilder();
            StringBuilder gotoMainFormTab = new StringBuilder();
            StringBuilder marketForm = new StringBuilder();
            StringBuilder mainForm = new StringBuilder();

            gotoMarketTab.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$btnMarketability");

            //
            //Market Tab
            //
            //SUBJECT MARKETABILITY 
            marketForm.AppendLine(@"SET !ERRORIGNORE YES");
            marketForm.AppendLine(@"SET !TIMEOUT_STEP 1");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns33_ConfirmNeighbor CONTENT=%Yes");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns34_SubjectIs CONTENT=%Appropriate<SP>Improvement");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns35_SubjectConsistant CONTENT=%Yes");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns12_CurrentOccupant CONTENT=%Homeowner");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns38_AllFinancing CONTENT=%Yes");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns51_SubjectUpdated CONTENT=%No");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns75_Disaster CONTENT=%No");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns52_EmergencyRepair CONTENT=%No");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns54_SubjectVendalism CONTENT=%No");
            marketForm.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns56_AdverseSafity CONTENT=No<SP>known<SP>adverse<SP>environmental/safety<SP>concerns.");
            marketForm.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns37_NegativeMarket CONTENT=No<SP>known<SP>Negative<SP>attributes<SP>to<SP>marketability.");
            marketForm.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns36_PositiveMarket CONTENT=<TBD>");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns57_RecomendedStrategy CONTENT=%As-Is");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns58_MostLikeBuyer CONTENT=%Owner<SP>Occupant");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns59_HOA CONTENT=%No");

            //TODO:
            //HOA fields

            //GENERAL MARKET CONDITIONS 
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns40_MarketConditions CONTENT=%Stable");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns70_EmploymentConditions CONTENT=%Stable");
            marketForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns77_ComparableSales CONTENT=" + form.setOfComps.numberSoldListings);
            marketForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns78_Low CONTENT=" + form.SetOfComps.minSalePrice);
            marketForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns78_High CONTENT=200,000");
            marketForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns79_CompetativeListing CONTENT=20");
            marketForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns80_Low CONTENT=100,000");
            marketForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns80_High CONTENT=200,000");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns05_CurrentInventory CONTENT=%Balanced");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns26_ValuesAppDep CONTENT=%Appreciated");
            marketForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns26_Percent CONTENT=5");
            marketForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns26_Months CONTENT=3");
            marketForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns81_MedianMarketRent CONTENT=1500");
            marketForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns82_TypicalMarketingTime CONTENT=100");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns83_MarketingTimeTrend CONTENT=%Stable");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns84_ReoTrend CONTENT=%Stable");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns09_REONeighborhood CONTENT=%Yes");
            marketForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns42_PercentDistressDisc CONTENT=20");
            marketForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns43_Owner CONTENT=90");
            marketForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns43_Tenant CONTENT=9");
            marketForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns43_Vacant CONTENT=1");
            marketForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns71_PercentDistressDisc CONTENT=0");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns85_NewConstruction CONTENT=%No");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns86_Industrial CONTENT=%No");
            marketForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$ddAns87_Disaster CONTENT=%No");
            marketForm.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns89_NeighborhoodDescription CONTENT=<TBD>");
            marketForm.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$txtAns32_GeneralComments CONTENT=Stable<SP>market<SP>with<SP>about<SP>20%<SP>REO<SP>sales<SP>mixed<SP>with<SP>short<SP>sales<SP>and<SP>traditional.<SP>High<SP>demand<SP>under<SP>150k.<SP>Seller<SP>concessions<SP>are<SP>not<SP>typical<SP>for<SP>the<SP>area");
            marketForm.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$MarketabilityV2$btnSaveData");



            gotoMainFormTab.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$btnForm");
            //
            //Main Tab
            //
            mainForm.AppendLine(@"SET !ERRORIGNORE YES");
            mainForm.AppendLine(@"SET !TIMEOUT_STEP 1");
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$owner_name CONTENT=" + form.SubjectOOR.Replace(" ", "<SP>"));
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_apn CONTENT=" + form.SubjectPin.Replace(" ", "<SP>"));

            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$fair_mkt_rent CONTENT=" + form.SubjectRent.Replace(" ", "<SP>"));

            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_zone_class CONTENT=SFR");
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_zone_desc CONTENT=Residential");
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_currentuse CONTENT=SFR");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_zone_high CONTENT=%Y");
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sale_dt");
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sale_price");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$currently_listed_yn CONTENT=%N");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$prev_listed_yn CONTENT=%N");
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$chkRFDamaged CONTENT=NO");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$property_type CONTENT=%SF");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_secure CONTENT=%Y");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$vacant_yn CONTENT=%N");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_zone_comp CONTENT=%L");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_illegal_units CONTENT=%N");
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$chkRFNone CONTENT=YES");
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_units CONTENT=1");
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_sqft CONTENT=" + form.SubjectAboveGLA);
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_rooms CONTENT=" + form.SubjectRoomCount);
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_bed CONTENT=" + form.SubjectBedroomCount);
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_bath CONTENT=" + form.SubjectBathroomCount.Split('.')[0]);
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_partial_bath CONTENT=" + form.SubjectBathroomCount.Split('.')[1]);
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_bsmt_fin_yn CONTENT=%F");
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_yr_built CONTENT=" + form.SubjectYearBuilt);
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_lotsize CONTENT=" + form.SubjectLotSize);
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_lotsize_unit CONTENT=%A");
            mainForm.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$pt_category CONTENT=YES");
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_sqft_bg CONTENT=" + form.SubjectBasementGLA);
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$condition CONTENT=%A");
            //mainForm.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:aspnetForm ATTR=TXT:Close");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$ddTypeLocation CONTENT=%S");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_garage_carport CONTENT=%A");
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_garage CONTENT=" + Regex.Match(form.SubjectParkingType, @"\d").Value);
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_parking CONTENT=%D");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_style CONTENT=%R");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_construction CONTENT=%F");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_view CONTENT=%R");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$ddTypeSource CONTENT=%M");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_waterfront CONTENT=%N");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_gated CONTENT=%N");
            mainForm.AppendLine(@"TAG POS=1 TYPE=TD FORM=NAME:aspnetForm ATTR=TXT:Public<SP>Sewer<SP>Septic");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_wastedisp CONTENT=%P");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_watersource CONTENT=%P");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_fireplace CONTENT=%0");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sub_pool CONTENT=%N");

            //04/28/2018 updates
            //fixed garage stall count
            //ctl00$Body$BPOV2B$sale_dt
            //ctl00$Body$BPOV2B$sale_price
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sale_dt CONTENT=" + form.subjectLastSaleDateTextbox.Text);
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV*$sale_price CONTENT=" + form.subjectLastSalePriceTextbox.Text);
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV2B$asis_value CONTENT=" + form.SubjectMarketValue);
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV2B$asis_list_price CONTENT=" + form.SubjectMarketValueList);
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV2B$repaired_value CONTENT=" + form.SubjectMarketValue);
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV2B$repaired_list_price CONTENT=" + form.SubjectMarketValueList);
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV2B$asis_value_quicksale CONTENT=" + form.SubjectQuickSaleValue);
;
            mainForm.AppendLine(@"TAB T=1");

            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV2B$lot_value_high CONTENT=15000");
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV2B$lot_value_low CONTENT=5000");
            mainForm.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV2B$txtInspectionDate CONTENT=" + form.dateTimePickerInspectionDate.Value.ToShortDateString());
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV2B$ddlInspectionTime CONTENT=%01:00");
            mainForm.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV2B$ddlMeridiem CONTENT=%PM");
            mainForm.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV2B$txtRehab CONTENT=No<SP>repairs<SP>noted.");




            macro.Append(gotoMarketTab);
            macro.Append(marketForm);
            macro.Append(gotoMainFormTab);
            macro.Append(mainForm);

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 60);

        }

        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {
            iim.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string currentUrl = iim.iimGetLastExtract();

            targetCompNumber = Regex.Match(compNum, @"\d").Value;
            saleOrListFlag = saleOrList;

            if (saleOrList == "sale")
            {

                GenerateSaleCompFillScript();

            }
            else
            {

                GenerateListCompFillScript();

            }

            WriteScript(fieldList["filepath"], compNum + ".iim", theMacro);

            string macroCode = theMacro.ToString();
            iim.iimPlayCode(macroCode, 60);
        }
    }
}
