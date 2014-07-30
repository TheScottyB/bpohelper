using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;


namespace bpohelper
{

    class BpoFormBaseClass
    {}



    class M2M
    {
        protected MLSListing targetComp;
        protected string targetCompNumber;
        protected StringBuilder theMacro = new StringBuilder();


        public M2M()
        {
          
        }


        public M2M(MLSListing m)
        {
            targetComp = m;
        }


        //FNMA example url
        //http://www.marktomarket.us/Order/OrderEditWizardFNMA/Step3/11487827

        protected Dictionary<string, string> propTypeTranslator = new Dictionary<string, string>()
         {
            {"bpohelper.DetachedListing", "%SFR<SP>Detached"},
            {"bpohelper.AttachedListing", "%SFR<SP>Attached"}
         };



        protected Dictionary<string, string> parkingTypeTranslator = new Dictionary<string, string>()
         {
            {"Attached", "%Attached<SP>Garage"},
            {"Detached", "%Detached<SP>Garage"},
            {"None", "%None"}
         };

        protected Dictionary<string, string> subjectFieldListTranslator = new Dictionary<string, string>()
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

        protected Dictionary<string, string> subjectFieldList = new Dictionary<string, string>();


        protected Dictionary<string, string> compFieldListTranslator = new Dictionary<string, string>()
        {  
            //method to call on MLSlisting, name of field on form
            {"StreetAddress", "Address1"}, 
            {"City", "City"},
            {"State", "State"},
            {"Zipcode", "ZipCode"},
            {"ProximityToSubject", "DistanceFromSubject"},
            {"CurrentListPrice", "CurrentFinalListPrice-text"},
            {"ListDate", "OriginalListDate"},
            {"DateOfLastPriceChange", "CurrentListDate"},
            {"SalePrice", "SalePrice-text"},
            {"SalesDate", "SaleDate"},
            {"DOM", "DaysOnMarket"},
            {"TransactionType", "TransactionType"},
            {"IsDateOfContractKnown", "IsDateOfContractKnown"},
            {"FinancingConcessions", "FinancingConcessions-text"},
            {"ContractDate", "DateOfContractForSale"},
            {"TypeOfFinancingConcessions", "TypeOfFinancingConcessions"},
            {"LocationRating", "OverallRating"},
            {"MainLocationFactor", "SelectedLocationFactors"},
            {"ViewRating", "SiteViewRating"},
            {"MainView", "SelectedSiteViewFactors"},
            {"Design", "Design"},
            {"DesignDescription", "DesignDescription"},
            {"FNMAConstructionQualityRating", "QualityConstructionDescription"},
            {"FNMAConditionRating", "Condition"},
            {"PropertyType", "PropertyType"},
            {"LeaseholdOrFeeSimple", "LeaseholdFreeSimpleDescription"},
            {"GLA", "Gla"},
            {"DataSource", "DataSource"},
            {"MlsSource", "MlsSource"},
            {"MlsNumber", "Mlsnumber"},
            {"ParkingType", "Parking"},
            {"NumberOfParkingSpaces", "NumberOfParkingSpaces"},
            {"PorchesPatioDeckFireplacesDescription", "PorchesPatioDeckFireplacesDescription"},
            {"PoolType", "Pool"},
            {"LotSize", "FNMASiteSize"},
            {"SiteSizeFormat", "SiteSizeFormat"},
            {"YearBuilt", "YearBuilt"},
            {"NumberOfUnits", "NumberOfUnits"},
            {"TotalRoomCount", "UnitBreakdownRmCountTot1"},
            {"BedroomCount", "UnitBreakdownRmCountBr1"},
            {"BathroomCount", "UnitBreakdownRmCountBaPlusHalfBaths1"},
            {"BasementSQFT", "BasementSQFT"},
            {"BasementFinishedSQFT", "BasementFinishedSQFT"},
            {"BasementAccessType", "BasementAccessType"},
            {"NumberOfBasementRecRooms", "NumberOfBasementRecRooms"},
            {"NumberOfBasementBedrooms", "NumberOfBasementBedrooms"},
            {"NumberOfBasementBathrooms", "NumberOfBasementBathrooms"},
            {"NumberOfOtherBasementRooms", "NumberOfOtherBasementRooms"},
            {"NumberOfFireplacesOrWoodStoves", "NumberOfFireplacesOrWoodStoves"},
            {"NumberOfPatiosOrDecks", "NumberOfPatiosOrDecks"},
            {"NumberOfPools", "NumberOfPools"},
            {"NumberOfFencedAreas", "NumberOfFencedAreas"},
            {"NumberOfPorches", "NumberOfPorches"},
            {"NumberOfOtherAmenities", "NumberOfOtherAmenities"},
            {"HasForcedWarmAirForHeat", "HasForcedWarmAirForHeat"},
            {"HasCentralAir", "HasCentralAir"},
            {"ListingStatus", "SaleStatusListingType"},
                //{"TRL", "Trail"},
            //{"TRL", "Trail"},
            //{"TRL", "Trail"},
            //{"WAY", "Way"}
           
        };

        protected Dictionary<string, string> compTextFieldList = new Dictionary<string, string>();
        protected Dictionary<string, string> compSelectionBoxList = new Dictionary<string, string>();
        protected Dictionary<string, string> compRadioButtonList = new Dictionary<string, string>();
        protected Dictionary<string, string> compCheckboxList = new Dictionary<string, string>();

        protected Dictionary<string, string> listingCompTextFieldList = new Dictionary<string, string>();
        protected Dictionary<string, string> listingCompSelectionBoxList = new Dictionary<string, string>();


        protected void WriteScript(string path, string filename, StringBuilder script)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(path + "\\" + filename))
            {
                file.Write(script);
            }


        }

        protected virtual  string GenerateSubjectFillScript()
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

        protected virtual void GenerateListCompFillScript()
        {
            helper_CompFillHeaderScript();
            compTextFieldList.Add(compFieldListTranslator["StreetAddress"], targetComp.StreetAddress);
            compTextFieldList.Add(compFieldListTranslator["City"], targetComp.City);
            compSelectionBoxList.Add(compFieldListTranslator["State"], "%IL");
            compTextFieldList.Add(compFieldListTranslator["Zipcode"], targetComp.Zipcode);
            compTextFieldList.Add(compFieldListTranslator["ProximityToSubject"], targetComp.ProximityToSubject.ToString());
            compTextFieldList.Add(compFieldListTranslator["CurrentListPrice"], targetComp.CurrentListPrice.ToString());
            compTextFieldList.Add(compFieldListTranslator["ListDate"], targetComp.ListDate.ToShortDateString());
            compTextFieldList.Add(compFieldListTranslator["DateOfLastPriceChange"], targetComp.DateOfLastPriceChange.ToShortDateString());
            compTextFieldList.Add(compFieldListTranslator["DOM"], targetComp.DOM);
            compSelectionBoxList.Add(compFieldListTranslator["TransactionType"], "%Arm");
            compRadioButtonList.Add(compFieldListTranslator["IsDateOfContractKnown"], "NO");
            compSelectionBoxList.Add(compFieldListTranslator["ListingStatus"], "%Active");
            compSelectionBoxList.Add(compFieldListTranslator["TypeOfFinancingConcessions"], "%None");
            compSelectionBoxList.Add(compFieldListTranslator["LocationRating"], "%N");
            compSelectionBoxList.Add(compFieldListTranslator["MainLocationFactor"], "%Res");
            compSelectionBoxList.Add(compFieldListTranslator["ViewRating"], "%N");
            compSelectionBoxList.Add(compFieldListTranslator["MainView"], "%Res");
            compSelectionBoxList.Add(compFieldListTranslator["Design"], "%Other");
            compTextFieldList.Add(compFieldListTranslator["DesignDescription"], "Contemporary");
            compSelectionBoxList.Add(compFieldListTranslator["FNMAConstructionQualityRating"], "%Q3");
            compSelectionBoxList.Add(compFieldListTranslator["FNMAConditionRating"], "%C3");
            compSelectionBoxList.Add(compFieldListTranslator["PropertyType"], propTypeTranslator[targetComp.ToString()]);
            compSelectionBoxList.Add(compFieldListTranslator["LeaseholdOrFeeSimple"], "%Fee<SP>Simple");
            compTextFieldList.Add(compFieldListTranslator["GLA"], targetComp.GLA.ToString());
            compSelectionBoxList.Add(compFieldListTranslator["DataSource"], "%MLS");
            compTextFieldList.Add(compFieldListTranslator["MlsSource"], "MRED");
            compTextFieldList.Add(compFieldListTranslator["MlsNumber"], targetComp.MlsNumber);
            compSelectionBoxList.Add(compFieldListTranslator["ParkingType"], parkingTypeTranslator[targetComp.GarageType()]);
            compTextFieldList.Add(compFieldListTranslator["NumberOfParkingSpaces"], targetComp.NumberGarageStalls());
            compTextFieldList.Add(compFieldListTranslator["PorchesPatioDeckFireplacesDescription"], "Unavailable");
            compTextFieldList.Add(compFieldListTranslator["LotSize"], targetComp.Lotsize.ToString());
            compSelectionBoxList.Add(compFieldListTranslator["SiteSizeFormat"], "%Ac");
            compTextFieldList.Add(compFieldListTranslator["YearBuilt"], targetComp.YearBuilt.ToString());
            compSelectionBoxList.Add(compFieldListTranslator["NumberOfUnits"], "%1");
            compTextFieldList.Add(compFieldListTranslator["TotalRoomCount"], targetComp.TotalRoomCount.ToString());
            compTextFieldList.Add(compFieldListTranslator["BedroomCount"], targetComp.BedroomCount);
            compTextFieldList.Add(compFieldListTranslator["BathroomCount"], targetComp.BathroomCount);
            compTextFieldList.Add(compFieldListTranslator["BasementSQFT"], targetComp.BasementGLA());
            compTextFieldList.Add(compFieldListTranslator["BasementFinishedSQFT"], targetComp.BasementGLA());
            compSelectionBoxList.Add(compFieldListTranslator["BasementAccessType"], "%in");
            compCheckboxList.Add(compFieldListTranslator["HasForcedWarmAirForHeat"], "YES");
            compCheckboxList.Add(compFieldListTranslator["HasCentralAir"], "YES");

            foreach (string field in compTextFieldList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEdit* ATTR=NAME:ListingComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compTextFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compSelectionBoxList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEdit* ATTR=NAME:ListingComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compSelectionBoxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compRadioButtonList.Keys)
            {
                string positionOfButton = "1";
                if (compRadioButtonList[field].ToLower() == "no")
                {
                    positionOfButton = "2";
                }

                theMacro.AppendFormat("TAG POS={3} TYPE=INPUT:RADIO FORM=ACTION:/Order/OrderEdit* ATTR=NAME:ListingComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compRadioButtonList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"), positionOfButton);

            }

            foreach (string field in compCheckboxList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEdit* ATTR=NAME:ListingComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compCheckboxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

        }

        protected virtual void GenerateSaleCompFillScript()
        {
            helper_CompFillHeaderScript();
            compTextFieldList.Add(compFieldListTranslator["StreetAddress"], targetComp.StreetAddress);
            compTextFieldList.Add(compFieldListTranslator["City"], targetComp.City);
            compSelectionBoxList.Add(compFieldListTranslator["State"], "%IL");
            compTextFieldList.Add(compFieldListTranslator["Zipcode"], targetComp.Zipcode);
            compTextFieldList.Add(compFieldListTranslator["ProximityToSubject"], targetComp.ProximityToSubject.ToString());
            compTextFieldList.Add(compFieldListTranslator["CurrentListPrice"], targetComp.CurrentListPrice.ToString());
            compTextFieldList.Add(compFieldListTranslator["ListDate"], targetComp.ListDate.ToShortDateString());
            compTextFieldList.Add(compFieldListTranslator["DateOfLastPriceChange"], targetComp.DateOfLastPriceChange.ToShortDateString());
            compTextFieldList.Add(compFieldListTranslator["SalePrice"], targetComp.SalePrice.ToString());
            compTextFieldList.Add(compFieldListTranslator["SalesDate"], targetComp.SalesDate.ToShortDateString());
            compTextFieldList.Add(compFieldListTranslator["DOM"], targetComp.DOM);
            compSelectionBoxList.Add(compFieldListTranslator["TransactionType"], "%Arm");
            compRadioButtonList.Add(compFieldListTranslator["IsDateOfContractKnown"], "YES");
            compTextFieldList.Add(compFieldListTranslator["ContractDate"], targetComp.ContractDate);
            compSelectionBoxList.Add(compFieldListTranslator["TypeOfFinancingConcessions"], "%None");
            compSelectionBoxList.Add(compFieldListTranslator["LocationRating"], "%N");
            compSelectionBoxList.Add(compFieldListTranslator["MainLocationFactor"], "%Res");
            compSelectionBoxList.Add(compFieldListTranslator["ViewRating"], "%N");
            compSelectionBoxList.Add(compFieldListTranslator["MainView"], "%Res");
            compSelectionBoxList.Add(compFieldListTranslator["Design"], "%Other");
            compTextFieldList.Add(compFieldListTranslator["DesignDescription"], "Contemporary");
            compSelectionBoxList.Add(compFieldListTranslator["FNMAConstructionQualityRating"], "%Q3");
            compSelectionBoxList.Add(compFieldListTranslator["FNMAConditionRating"], "%C3");
            compSelectionBoxList.Add(compFieldListTranslator["PropertyType"], propTypeTranslator[targetComp.ToString()]);
            compSelectionBoxList.Add(compFieldListTranslator["LeaseholdOrFeeSimple"], "%Fee<SP>Simple");
            compTextFieldList.Add(compFieldListTranslator["GLA"], targetComp.GLA.ToString());
            compSelectionBoxList.Add(compFieldListTranslator["DataSource"], "%MLS");
            compTextFieldList.Add(compFieldListTranslator["MlsSource"], "MRED");
            compTextFieldList.Add(compFieldListTranslator["MlsNumber"], targetComp.MlsNumber);
            compSelectionBoxList.Add(compFieldListTranslator["ParkingType"], parkingTypeTranslator[targetComp.GarageType()]);
            compTextFieldList.Add(compFieldListTranslator["NumberOfParkingSpaces"], targetComp.NumberGarageStalls());
            compTextFieldList.Add(compFieldListTranslator["PorchesPatioDeckFireplacesDescription"], "Unavailable");
            compTextFieldList.Add(compFieldListTranslator["LotSize"], targetComp.Lotsize.ToString());
            compSelectionBoxList.Add(compFieldListTranslator["SiteSizeFormat"], "%Ac");
            compTextFieldList.Add(compFieldListTranslator["YearBuilt"], targetComp.YearBuilt.ToString());
            compSelectionBoxList.Add(compFieldListTranslator["NumberOfUnits"], "%1");
            compTextFieldList.Add(compFieldListTranslator["TotalRoomCount"], targetComp.TotalRoomCount.ToString());
            compTextFieldList.Add(compFieldListTranslator["BedroomCount"], targetComp.BedroomCount);
            compTextFieldList.Add(compFieldListTranslator["BathroomCount"], targetComp.BathroomCount);
            compTextFieldList.Add(compFieldListTranslator["BasementSQFT"], targetComp.BasementGLA());
            compTextFieldList.Add(compFieldListTranslator["BasementFinishedSQFT"], targetComp.BasementGLA());
            compSelectionBoxList.Add(compFieldListTranslator["BasementAccessType"], "%in");
            compCheckboxList.Add(compFieldListTranslator["HasForcedWarmAirForHeat"], "YES");
            compCheckboxList.Add(compFieldListTranslator["HasCentralAir"], "YES");


            foreach (string field in compTextFieldList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEdit* ATTR=NAME:SalesComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compTextFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compSelectionBoxList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEdit* ATTR=NAME:SalesComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compSelectionBoxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }


            foreach (string field in compRadioButtonList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ACTION:/Order/OrderEdit* ATTR=NAME:SalesComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compRadioButtonList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compCheckboxList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEdit* ATTR=NAME:SalesComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compCheckboxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }
      
    
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEditWizardFNMA/step3/11488927 ATTR=NAME:SalesComp1.HasForcedWarmAirForHeat CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEditWizardFNMA/step3/11488927 ATTR=NAME:SalesComp1.HasCentralAir CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=ACTION:/Order/OrderEditWizardFNMA/step3/11488927 ATTR=CLASS:mceIframeContainer<SP>mceFirst<SP>mceLast");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:/Order/OrderEditWizardFNMA/step3/11488927 ATTR=NAME:SaveButton");  
        }

        protected void helper_CompFillHeaderScript()
        {
            theMacro.Clear();
            theMacro.AppendLine(@"SET !ERRORIGNORE YES");
            theMacro.AppendLine(@"SET !TIMEOUT_STEP 0");
            theMacro.AppendLine(@"SET !REPLAYSPEED FAST");
        }



        public void Prefill(iMacros.App iim, Form1 form)
        {
            // iim.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string currentUrl = iim.iimGetLastExtract();

            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 1");
            macro.Append(GenerateSubjectFillScript());

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 120);
        }



        public virtual void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {
            //http://www.marktomarket.us/Order/OrderEditWizardFNMA/Step3/11487827
            iim.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string currentUrl = iim.iimGetLastExtract();

            targetCompNumber = Regex.Match(compNum, @"\d").Value;



            if (currentUrl.ToLower().Contains("fnma"))
            {

                //we are on fannie mae form
                //step 3 = sales comps
                //step 4 = active listings comps
                if (saleOrList == "sale")
                {
                    iim.iimPlayCode(@"TAG POS=1 TYPE=INPUT:SUBMIT ATTR=ID:step3_submit");
                    GenerateSaleCompFillScript();

                }
                else
                {
                    iim.iimPlayCode(@"TAG POS=1 TYPE=INPUT:SUBMIT ATTR=ID:step4_submit");
                    GenerateListCompFillScript();

                }

                WriteScript(fieldList["filepath"], compNum + ".iim", theMacro);

                string macroCode = theMacro.ToString();
                iim.iimPlayCode(macroCode, 60);



            }



        }


    }

    class M2MStandard : M2M
    {
        //
        //Exterior with Comp Photos - M
        //

        public M2MStandard(MLSListing m) : base(m)
        {

        }

        protected override void GenerateListCompFillScript()
        {
            helper_CompFillHeaderScript();
            compTextFieldList.Add(compFieldListTranslator["MlsNumber"], targetComp.MlsNumber);

            compTextFieldList.Add(compFieldListTranslator["StreetAddress"], targetComp.StreetAddress);
            compTextFieldList.Add(compFieldListTranslator["City"], targetComp.City);
            compSelectionBoxList.Add(compFieldListTranslator["State"], "%IL");
            compTextFieldList.Add(compFieldListTranslator["Zipcode"], targetComp.Zipcode);
            compSelectionBoxList.Add(compFieldListTranslator["PropertyType"], propTypeTranslator[targetComp.ToString()]);
            compTextFieldList.Add(compFieldListTranslator["OriginalListPrice"], targetComp.OriginalListPrice.ToString());
            compTextFieldList.Add(compFieldListTranslator["CurrentListPrice"], targetComp.CurrentListPrice.ToString());
            compTextFieldList.Add(compFieldListTranslator["ListDate"], targetComp.ListDate.ToShortDateString());
          
            compTextFieldList.Add(compFieldListTranslator["Adjustment-text"], "0");
            compTextFieldList.Add(compFieldListTranslator["ProximityToSubject"], targetComp.ProximityToSubject.ToString());
            compTextFieldList.Add(compFieldListTranslator["DOM"], targetComp.DOM);
            compTextFieldList.Add(compFieldListTranslator["Type"], targetComp.PropertyType());
            compSelectionBoxList.Add(compFieldListTranslator["Exterior"], "");
            compTextFieldList.Add(compFieldListTranslator["YearBuilt"], targetComp.YearBuilt.ToString());
            compTextFieldList.Add(compFieldListTranslator["GLA"], targetComp.GLA.ToString());
            compTextFieldList.Add(compFieldListTranslator["FinishedSf"], targetComp.GLA.ToString());
            compSelectionBoxList.Add(compFieldListTranslator["Condition"], "%3-Average");
            compTextFieldList.Add(compFieldListTranslator["TotalRoomCount"], targetComp.TotalRoomCount.ToString());
            compTextFieldList.Add(compFieldListTranslator["BedroomCount"], targetComp.BedroomCount);
            compTextFieldList.Add(compFieldListTranslator["FullBathrooms"], targetComp.FullBathCount);
            compTextFieldList.Add(compFieldListTranslator["HalfBaths"], targetComp.HalfBathCount);
            compTextFieldList.Add(compFieldListTranslator["NumberOfBasementRooms"], targetComp.NumberOfBasementRooms());
            compTextFieldList.Add(compFieldListTranslator["BasementSQFT"], targetComp.BasementGLA());
            compTextFieldList.Add(compFieldListTranslator["BasementFinishedPercentage"], targetComp.BasementFinishedPercentage());
            compSelectionBoxList.Add(compFieldListTranslator["ParkingType"], parkingTypeTranslator[targetComp.GarageString()]);
            compSelectionBoxList.Add(compFieldListTranslator["TransactionType"], salesTypeTranslator[targetComp.TransactionType]);
            compTextFieldList.Add(compFieldListTranslator["LotSize"], targetComp.Lotsize.ToString());
            compSelectionBoxList.Add(compFieldListTranslator["PoolType"], "%None");

            foreach (string field in compTextFieldList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEdit* ATTR=NAME:Listing{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compTextFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compSelectionBoxList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEdit* ATTR=NAME:Listing{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compSelectionBoxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compRadioButtonList.Keys)
            {
                string positionOfButton = "1";
                if (compRadioButtonList[field].ToLower() == "no")
                {
                    positionOfButton = "2";
                }

                theMacro.AppendFormat("TAG POS={3} TYPE=INPUT:RADIO FORM=ACTION:/Order/OrderEdit* ATTR=NAME:Listing{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compRadioButtonList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"), positionOfButton);

            }

            foreach (string field in compCheckboxList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEdit* ATTR=NAME:Listing{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compCheckboxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

        }

        protected override void GenerateSaleCompFillScript()
        {
            helper_CompFillHeaderScript();
            compTextFieldList.Add(compFieldListTranslator["MlsNumber"], targetComp.MlsNumber);

            compTextFieldList.Add(compFieldListTranslator["StreetAddress"], targetComp.StreetAddress);
            compTextFieldList.Add(compFieldListTranslator["City"], targetComp.City);
            compSelectionBoxList.Add(compFieldListTranslator["State"], "%IL");
            compTextFieldList.Add(compFieldListTranslator["Zipcode"], targetComp.Zipcode);
            compSelectionBoxList.Add(compFieldListTranslator["PropertyType"], propTypeTranslator[targetComp.ToString()]);
            compTextFieldList.Add(compFieldListTranslator["OriginalListPrice"], targetComp.OriginalListPrice.ToString());
            compTextFieldList.Add(compFieldListTranslator["ListDate"], targetComp.ListDate.ToShortDateString());
            compTextFieldList.Add(compFieldListTranslator["SalePrice"], targetComp.SalePrice.ToString());
            compTextFieldList.Add(compFieldListTranslator["SalesDate"], targetComp.SalesDate.ToShortDateString());
            compTextFieldList.Add(compFieldListTranslator["Adjustment-text"], "0");
            compTextFieldList.Add(compFieldListTranslator["ProximityToSubject"], targetComp.ProximityToSubject.ToString());
            compTextFieldList.Add(compFieldListTranslator["DOM"], targetComp.DOM);
            compTextFieldList.Add(compFieldListTranslator["Type"], targetComp.PropertyType());
            compSelectionBoxList.Add(compFieldListTranslator["Exterior"], "");
            compTextFieldList.Add(compFieldListTranslator["YearBuilt"], targetComp.YearBuilt.ToString());
            compTextFieldList.Add(compFieldListTranslator["GLA"], targetComp.GLA.ToString());
            compTextFieldList.Add(compFieldListTranslator["FinishedSf"], targetComp.GLA.ToString());
            compSelectionBoxList.Add(compFieldListTranslator["Condition"], "%3-Average");
            compTextFieldList.Add(compFieldListTranslator["TotalRoomCount"], targetComp.TotalRoomCount.ToString());
            compTextFieldList.Add(compFieldListTranslator["BedroomCount"], targetComp.BedroomCount);
            compTextFieldList.Add(compFieldListTranslator["FullBathrooms"], targetComp.FullBathCount);
            compTextFieldList.Add(compFieldListTranslator["HalfBaths"], targetComp.HalfBathCount);
            compTextFieldList.Add(compFieldListTranslator["NumberOfBasementRooms"], targetComp.NumberOfBasementRooms());
            compTextFieldList.Add(compFieldListTranslator["BasementSQFT"], targetComp.BasementGLA());
            compTextFieldList.Add(compFieldListTranslator["BasementFinishedPercentage"], targetComp.BasementFinishedPercentage());
             compSelectionBoxList.Add(compFieldListTranslator["ParkingType"], parkingTypeTranslator[targetComp.GarageString()]);
             compSelectionBoxList.Add(compFieldListTranslator["TransactionType"], salesTypeTranslator[targetComp.TransactionType]);
             compTextFieldList.Add(compFieldListTranslator["LotSize"], targetComp.Lotsize.ToString());
             compSelectionBoxList.Add(compFieldListTranslator["PoolType"], "%None");

       
   

            foreach (string field in compTextFieldList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEdit* ATTR=NAME:Sale{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compTextFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compSelectionBoxList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEdit* ATTR=NAME:Sale{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compSelectionBoxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }


            foreach (string field in compRadioButtonList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ACTION:/Order/OrderEdit* ATTR=NAME:Sale{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compRadioButtonList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compCheckboxList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEdit* ATTR=NAME:Sale{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compCheckboxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }


        }

        protected override string GenerateSubjectFillScript()
        {
            StringBuilder macro = new StringBuilder();
     
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:Subject.PropertyType CONTENT=%Single<SP>Family<SP>Detached");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:Subject.Style CONTENT=%1<SP>1/2<SP>Story");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:Subject.Exterior CONTENT=%Metal/Vinyl");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:Subject.YearBuilt CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectYearBuilt);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:Subject.AboveGradeSf CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectAboveGLA);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:Subject.FinishedSf CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectAboveGLA);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:Subject.TotalRooms CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectRoomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:Subject.Bedrooms CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectBedroomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:Subject.Bathrooms CONTENT=" + GlobalVar.theSubjectProperty.FullBathCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:Subject.HalfBaths CONTENT=" + GlobalVar.theSubjectProperty.HalfBathCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:Subject.BasementRooms CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:Subject.BasementSQFT CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectBasementGLA);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:Subject.PercentageBasementFinished CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:Subject.Garage CONTENT=%1<SP>ATTACHED");

            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:Subject.SiteSize CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectLotSize);
            
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:Subject.Pool CONTENT=%None");
            
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:/Order/OrderEditWizard/Step3/11622619 ATTR=NAME:SaveButton");
            return macro.ToString();

        }


        protected  Dictionary<string, string> salesTypeTranslator = new Dictionary<string, string>()
         { 
               {"Arms Length", "%Fair<SP>Market"},
             {"ShortSale", "%Short<SP>Sale"},
              {"REO", "%REO"},
            {"None", "%ON_SITE"}
         };

        protected Dictionary<string, string> parkingTypeTranslator = new Dictionary<string, string>()
         {
            {"1 Attached", "%1<SP>ATTACHED"},
             {"2 Attached", "%2<SP>ATTACHED"},
              {"3 Attached", "%3<SP>ATTACHED"},
               {"1 Detached", "%1<SP>DETACHED"},
             {"2 Detached", "%2<SP>DETACHED"},
              {"3 Detached", "%3<SP>DETACHED"},
            {"None", "%ON_SITE"}
         };

        public override void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {
            //http://www.marktomarket.us/Order/OrderEditWizardFNMA/Step3/11487827
            iim.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string currentUrl = iim.iimGetLastExtract();

            targetCompNumber = Regex.Match(compNum, @"\d").Value;



            if (!currentUrl.ToLower().Contains("fnma"))
            {

                //we are on non fannie mae form
                //step 3 = sales comps
                //step 4 = active listings comps
                if (saleOrList == "sale")
                {
                    iim.iimPlayCode(@"TAG POS=1 TYPE=INPUT:SUBMIT ATTR=ID:step3_submit");
                    GenerateSaleCompFillScript();

                }
                else
                {
                    iim.iimPlayCode(@"TAG POS=1 TYPE=INPUT:SUBMIT ATTR=ID:step4_submit");
                    GenerateListCompFillScript();

                }

                WriteScript(fieldList["filepath"], compNum + ".iim", theMacro);

                string macroCode = theMacro.ToString();
                iim.iimPlayCode(macroCode, 60);



            }
        }

        protected Dictionary<string, string> propTypeTranslator = new Dictionary<string, string>()
         {
            {"bpohelper.DetachedListing", "%Single<SP>Family<SP>Detached"},
            {"bpohelper.AttachedListing", "%Condominium"}
         };


         protected   Dictionary<string, string> compFieldListTranslator = new Dictionary<string, string>()
        {  
            //method to call on MLSlisting, name of field on form
            {"MlsNumber", "Mlsnumber"},
            {"StreetAddress", "Address1"}, 
            {"City", "City"},
            {"State", "State"},
            {"Zipcode", "ZipCode"},
            {"PropertyType", "PropertyType"},
            {"OriginalListPrice", "OriginalListPrice-text"},
            {"CurrentListPrice", "CurrentListPrice-text"},
            {"ListDate", "OriginalListDate"},
            {"SalePrice", "SalePrice-text"},
            {"SalesDate", "SaleDate"},
            {"Adjustment-text", "Adjustment-text"},
            {"ProximityToSubject", "Distance"},
            {"DOM", "DaysOnMarket"},
            {"Type", "Style"},
            {"Exterior", "Exterior"},
            {"YearBuilt", "YearBuilt"},
            {"GLA", "AboveGradeSf"},
            {"FinishedSf", "FinishedSf"},
            {"Condition", "Condition"},
            {"TotalRoomCount", "TotalRooms"},
            {"BedroomCount", "Bedrooms"},
            {"FullBathrooms", "Bathrooms"},
            {"HalfBaths", "HalfBaths"},
            {"NumberOfBasementRooms", "BasementRooms"},
            {"BasementSQFT", "BasementSQFT"},
            {"BasementFinishedPercentage", "PercentageBasementFinished"},
            {"ParkingType", "Garage"},
            {"TransactionType", "TransactionType"},
            {"LotSize", "SiteSize"},
            {"PoolType", "Pool"},
            //{"TRL", "Trail"},
            //{"WAY", "Way"}
           
        };


        //public void M2MStandard(MLSListing m)
        //{
        //    compFieldListTranslator.Clear();

        //    compFieldListTranslator.Add()
        //}

        //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.Mlsnumber CONTENT=11");
        //macro.AppendLine(@"'text input activated");
        //macro.AppendLine(@"'text input activated");
        //macro.AppendLine(@"'text input activated");
        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.State CONTENT=%IL");
        //macro.AppendLine(@"'text input activated");
        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.PropertyType CONTENT=%Single<SP>Family<SP>Detached");
        //macro.AppendLine(@"'text input activated");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.OriginalListDate CONTENT=1");
        //macro.AppendLine(@"'text input activated");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.SaleDate CONTENT=1");
        //macro.AppendLine(@"'text input activated");
        //macro.AppendLine(@"'text input activated");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.Distance CONTENT=1");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.DaysOnMarket CONTENT=1");
        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.Style CONTENT=%2<SP>Story");
        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.Exterior CONTENT=%Brick");
        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.Exterior CONTENT=%Metal/Vinyl");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.YearBuilt CONTENT=1900");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.AboveGradeSf CONTENT=100");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.FinishedSf CONTENT=100");
        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.Condition CONTENT=%3-Average");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.TotalRooms CONTENT=10");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.Bedrooms CONTENT=4");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.Bathrooms CONTENT=2");
        //macro.AppendLine(@"'text input activated");
        //macro.AppendLine(@"'text input activated");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.HalfBaths CONTENT=1");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.BasementRooms CONTENT=1");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.BasementSQFT CONTENT=10");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.PercentageBasementFinished CONTENT=100");
        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.Garage CONTENT=%2<SP>ATTACHED");
        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.TransactionType CONTENT=%Fair<SP>Market");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.SiteSize CONTENT=1");
        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.Pool CONTENT=%None");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:Sale1.AboveGradePriceSf CONTENT=0.00");
        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:/Order/OrderEditWizard/Step3/11515504 ATTR=NAME:NextButton");
        //string macroCode = macro.ToString();
        //// status = iim.iimPlayCode(macroCode, timeout);

    }

    class M2MFNMA : M2M
    {

           protected MLSListing targetComp;
        protected string targetCompNumber;
        protected StringBuilder theMacro = new StringBuilder();


        public M2MFNMA()
        {
          
        }


        public M2MFNMA(MLSListing m)
        {
            targetComp = m;
        }


        //FNMA example url
        //http://www.marktomarket.us/Order/OrderEditWizardFNMA/Step3/11487827

        protected Dictionary<string, string> propTypeTranslator = new Dictionary<string, string>()
         {
            {"bpohelper.DetachedListing", "%SFR<SP>Detached"},
            {"bpohelper.AttachedListing", "%SFR<SP>Attached"}
         };

        protected Dictionary<string, string> parkingTypeTranslator = new Dictionary<string, string>()
         {
            {"Attached", "%Attached<SP>Garage"},
            {"Detached", "%Detached<SP>Garage"}
         };

        protected Dictionary<string, string> subjectFieldListTranslator = new Dictionary<string, string>()
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

        protected Dictionary<string, string> subjectFieldList = new Dictionary<string, string>();


        protected Dictionary<string, string> compFieldListTranslator = new Dictionary<string, string>()
        {  
            //method to call on MLSlisting, name of field on form
            {"StreetAddress", "Address1"}, 
            {"City", "City"},
            {"State", "State"},
            {"Zipcode", "ZipCode"},
            {"ProximityToSubject", "DistanceFromSubject"},
            {"CurrentListPrice", "CurrentFinalListPrice-text"},
            {"ListDate", "OriginalListDate"},
            {"DateOfLastPriceChange", "CurrentListDate"},
            {"SalePrice", "SalePrice-text"},
            {"SalesDate", "SaleDate"},
            {"DOM", "DaysOnMarket"},
            {"TransactionType", "TransactionType"},
            {"IsDateOfContractKnown", "IsDateOfContractKnown"},
            {"FinancingConcessions", "FinancingConcessions-text"},
            {"ContractDate", "DateOfContractForSale"},
            {"TypeOfFinancingConcessions", "TypeOfFinancingConcessions"},
            {"LocationRating", "OverallRating"},
            {"MainLocationFactor", "SelectedLocationFactors"},
            {"ViewRating", "SiteViewRating"},
            {"MainView", "SelectedSiteViewFactors"},
            {"Design", "Design"},
            {"DesignDescription", "DesignDescription"},
            {"FNMAConstructionQualityRating", "QualityConstructionDescription"},
            {"FNMAConditionRating", "Condition"},
            {"PropertyType", "PropertyType"},
            {"LeaseholdOrFeeSimple", "LeaseholdFreeSimpleDescription"},
            {"GLA", "Gla"},
            {"DataSource", "DataSource"},
            {"MlsSource", "MlsSource"},
            {"MlsNumber", "Mlsnumber"},
            {"ParkingType", "Parking"},
            {"NumberOfParkingSpaces", "NumberOfParkingSpaces"},
            {"PorchesPatioDeckFireplacesDescription", "PorchesPatioDeckFireplacesDescription"},
            {"PoolType", "Pool"},
            {"LotSize", "FNMASiteSize"},
            {"SiteSizeFormat", "SiteSizeFormat"},
            {"YearBuilt", "YearBuilt"},
            {"NumberOfUnits", "NumberOfUnits"},
            {"TotalRoomCount", "UnitBreakdownRmCountTot1"},
            {"BedroomCount", "UnitBreakdownRmCountBr1"},
            {"BathroomCount", "UnitBreakdownRmCountBaPlusHalfBaths1"},
            {"BasementSQFT", "BasementSQFT"},
            {"BasementFinishedSQFT", "BasementFinishedSQFT"},
            {"BasementAccessType", "BasementAccessType"},
            {"NumberOfBasementRecRooms", "NumberOfBasementRecRooms"},
            {"NumberOfBasementBedrooms", "NumberOfBasementBedrooms"},
            {"NumberOfBasementBathrooms", "NumberOfBasementBathrooms"},
            {"NumberOfOtherBasementRooms", "NumberOfOtherBasementRooms"},
            {"NumberOfFireplacesOrWoodStoves", "NumberOfFireplacesOrWoodStoves"},
            {"NumberOfPatiosOrDecks", "NumberOfPatiosOrDecks"},
            {"NumberOfPools", "NumberOfPools"},
            {"NumberOfFencedAreas", "NumberOfFencedAreas"},
            {"NumberOfPorches", "NumberOfPorches"},
            {"NumberOfOtherAmenities", "NumberOfOtherAmenities"},
            {"HasForcedWarmAirForHeat", "HasForcedWarmAirForHeat"},
            {"HasCentralAir", "HasCentralAir"},
            {"ListingStatus", "SaleStatusListingType"},
                //{"TRL", "Trail"},
            //{"TRL", "Trail"},
            //{"TRL", "Trail"},
            //{"WAY", "Way"}
           
        };

        protected Dictionary<string, string> compTextFieldList = new Dictionary<string, string>();
        protected Dictionary<string, string> compSelectionBoxList = new Dictionary<string, string>();
        protected Dictionary<string, string> compRadioButtonList = new Dictionary<string, string>();
        protected Dictionary<string, string> compCheckboxList = new Dictionary<string, string>();

        protected Dictionary<string, string> listingCompTextFieldList = new Dictionary<string, string>();
        protected Dictionary<string, string> listingCompSelectionBoxList = new Dictionary<string, string>();


        protected void WriteScript(string path, string filename, StringBuilder script)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(path + "\\" + filename))
            {
                file.Write(script);
            }


        }

        protected string GenerateSubjectFillScript()
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

        protected virtual void GenerateListCompFillScript()
        {
            helper_CompFillHeaderScript();
            compTextFieldList.Add(compFieldListTranslator["StreetAddress"], targetComp.StreetAddress);
            compTextFieldList.Add(compFieldListTranslator["City"], targetComp.City);
            compSelectionBoxList.Add(compFieldListTranslator["State"], "%IL");
            compTextFieldList.Add(compFieldListTranslator["Zipcode"], targetComp.Zipcode);
            compTextFieldList.Add(compFieldListTranslator["ProximityToSubject"], targetComp.ProximityToSubject.ToString());
            compTextFieldList.Add(compFieldListTranslator["CurrentListPrice"], targetComp.CurrentListPrice.ToString());
            compTextFieldList.Add(compFieldListTranslator["ListDate"], targetComp.ListDate.ToShortDateString());
            compTextFieldList.Add(compFieldListTranslator["DateOfLastPriceChange"], targetComp.DateOfLastPriceChange.ToShortDateString());
            compTextFieldList.Add(compFieldListTranslator["DOM"], targetComp.DOM);
            compSelectionBoxList.Add(compFieldListTranslator["TransactionType"], "%Arm");
            compRadioButtonList.Add(compFieldListTranslator["IsDateOfContractKnown"], "NO");
            compSelectionBoxList.Add(compFieldListTranslator["ListingStatus"], "%Active");
            compSelectionBoxList.Add(compFieldListTranslator["TypeOfFinancingConcessions"], "%None");
            compSelectionBoxList.Add(compFieldListTranslator["LocationRating"], "%N");
            compSelectionBoxList.Add(compFieldListTranslator["MainLocationFactor"], "%Res");
            compSelectionBoxList.Add(compFieldListTranslator["ViewRating"], "%N");
            compSelectionBoxList.Add(compFieldListTranslator["MainView"], "%Res");
            compSelectionBoxList.Add(compFieldListTranslator["Design"], "%Other");
            compTextFieldList.Add(compFieldListTranslator["DesignDescription"], "Contemporary");
            compSelectionBoxList.Add(compFieldListTranslator["FNMAConstructionQualityRating"], "%Q3");
            compSelectionBoxList.Add(compFieldListTranslator["FNMAConditionRating"], "%C3");
            compSelectionBoxList.Add(compFieldListTranslator["PropertyType"], propTypeTranslator[targetComp.ToString()]);
            compSelectionBoxList.Add(compFieldListTranslator["LeaseholdOrFeeSimple"], "%Fee<SP>Simple");
            compTextFieldList.Add(compFieldListTranslator["GLA"], targetComp.GLA.ToString());
            compSelectionBoxList.Add(compFieldListTranslator["DataSource"], "%MLS");
            compTextFieldList.Add(compFieldListTranslator["MlsSource"], "MRED");
            compTextFieldList.Add(compFieldListTranslator["MlsNumber"], targetComp.MlsNumber);
            compSelectionBoxList.Add(compFieldListTranslator["ParkingType"], parkingTypeTranslator[targetComp.GarageType()]);
            compTextFieldList.Add(compFieldListTranslator["NumberOfParkingSpaces"], targetComp.NumberGarageStalls());
            compTextFieldList.Add(compFieldListTranslator["PorchesPatioDeckFireplacesDescription"], "Unavailable");
            compTextFieldList.Add(compFieldListTranslator["LotSize"], targetComp.Lotsize.ToString());
            compSelectionBoxList.Add(compFieldListTranslator["SiteSizeFormat"], "%Ac");
            compTextFieldList.Add(compFieldListTranslator["YearBuilt"], targetComp.YearBuilt.ToString());
            compSelectionBoxList.Add(compFieldListTranslator["NumberOfUnits"], "%1");
            compTextFieldList.Add(compFieldListTranslator["TotalRoomCount"], targetComp.TotalRoomCount.ToString());
            compTextFieldList.Add(compFieldListTranslator["BedroomCount"], targetComp.BedroomCount);
            compTextFieldList.Add(compFieldListTranslator["BathroomCount"], targetComp.BathroomCount);
            compTextFieldList.Add(compFieldListTranslator["BasementSQFT"], targetComp.BasementGLA());
            compTextFieldList.Add(compFieldListTranslator["BasementFinishedSQFT"], targetComp.BasementGLA());
            compSelectionBoxList.Add(compFieldListTranslator["BasementAccessType"], "%in");
            compCheckboxList.Add(compFieldListTranslator["HasForcedWarmAirForHeat"], "YES");
            compCheckboxList.Add(compFieldListTranslator["HasCentralAir"], "YES");

            foreach (string field in compTextFieldList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEdit* ATTR=NAME:ListingComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compTextFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compSelectionBoxList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEdit* ATTR=NAME:ListingComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compSelectionBoxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compRadioButtonList.Keys)
            {
                string positionOfButton = "1";
                if (compRadioButtonList[field].ToLower() == "no")
                {
                    positionOfButton = "2";
                }

                theMacro.AppendFormat("TAG POS={3} TYPE=INPUT:RADIO FORM=ACTION:/Order/OrderEdit* ATTR=NAME:ListingComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compRadioButtonList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"), positionOfButton);

            }

            foreach (string field in compCheckboxList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEdit* ATTR=NAME:ListingComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compCheckboxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

        }

        protected virtual void GenerateSaleCompFillScript()
        {
            helper_CompFillHeaderScript();
            compTextFieldList.Add(compFieldListTranslator["StreetAddress"], targetComp.StreetAddress);
            compTextFieldList.Add(compFieldListTranslator["City"], targetComp.City);
            compSelectionBoxList.Add(compFieldListTranslator["State"], "%IL");
            compTextFieldList.Add(compFieldListTranslator["Zipcode"], targetComp.Zipcode);
            compTextFieldList.Add(compFieldListTranslator["ProximityToSubject"], targetComp.ProximityToSubject.ToString());
            compTextFieldList.Add(compFieldListTranslator["CurrentListPrice"], targetComp.CurrentListPrice.ToString());
            compTextFieldList.Add(compFieldListTranslator["ListDate"], targetComp.ListDate.ToShortDateString());
            compTextFieldList.Add(compFieldListTranslator["DateOfLastPriceChange"], targetComp.DateOfLastPriceChange.ToShortDateString());
            compTextFieldList.Add(compFieldListTranslator["SalePrice"], targetComp.SalePrice.ToString());
            compTextFieldList.Add(compFieldListTranslator["SalesDate"], targetComp.SalesDate.ToShortDateString());
            compTextFieldList.Add(compFieldListTranslator["DOM"], targetComp.DOM);
            compSelectionBoxList.Add(compFieldListTranslator["TransactionType"], "%Arm");
            compRadioButtonList.Add(compFieldListTranslator["IsDateOfContractKnown"], "YES");
            compTextFieldList.Add(compFieldListTranslator["ContractDate"], targetComp.ContractDate);
            compSelectionBoxList.Add(compFieldListTranslator["TypeOfFinancingConcessions"], "%None");
            compSelectionBoxList.Add(compFieldListTranslator["LocationRating"], "%N");
            compSelectionBoxList.Add(compFieldListTranslator["MainLocationFactor"], "%Res");
            compSelectionBoxList.Add(compFieldListTranslator["ViewRating"], "%N");
            compSelectionBoxList.Add(compFieldListTranslator["MainView"], "%Res");
            compSelectionBoxList.Add(compFieldListTranslator["Design"], "%Other");
            compTextFieldList.Add(compFieldListTranslator["DesignDescription"], "Contemporary");
            compSelectionBoxList.Add(compFieldListTranslator["FNMAConstructionQualityRating"], "%Q3");
            compSelectionBoxList.Add(compFieldListTranslator["FNMAConditionRating"], "%C3");
            compSelectionBoxList.Add(compFieldListTranslator["PropertyType"], propTypeTranslator[targetComp.ToString()]);
            compSelectionBoxList.Add(compFieldListTranslator["LeaseholdOrFeeSimple"], "%Fee<SP>Simple");
            compTextFieldList.Add(compFieldListTranslator["GLA"], targetComp.GLA.ToString());
            compSelectionBoxList.Add(compFieldListTranslator["DataSource"], "%MLS");
            compTextFieldList.Add(compFieldListTranslator["MlsSource"], "MRED");
            compTextFieldList.Add(compFieldListTranslator["MlsNumber"], targetComp.MlsNumber);
            compSelectionBoxList.Add(compFieldListTranslator["ParkingType"], parkingTypeTranslator[targetComp.GarageType()]);
            compTextFieldList.Add(compFieldListTranslator["NumberOfParkingSpaces"], targetComp.NumberGarageStalls());
            compTextFieldList.Add(compFieldListTranslator["PorchesPatioDeckFireplacesDescription"], "Unavailable");
            compTextFieldList.Add(compFieldListTranslator["LotSize"], targetComp.Lotsize.ToString());
            compSelectionBoxList.Add(compFieldListTranslator["SiteSizeFormat"], "%Ac");
            compTextFieldList.Add(compFieldListTranslator["YearBuilt"], targetComp.YearBuilt.ToString());
            compSelectionBoxList.Add(compFieldListTranslator["NumberOfUnits"], "%1");
            compTextFieldList.Add(compFieldListTranslator["TotalRoomCount"], targetComp.TotalRoomCount.ToString());
            compTextFieldList.Add(compFieldListTranslator["BedroomCount"], targetComp.BedroomCount);
            compTextFieldList.Add(compFieldListTranslator["BathroomCount"], targetComp.BathroomCount);
            compTextFieldList.Add(compFieldListTranslator["BasementSQFT"], targetComp.BasementGLA());
            compTextFieldList.Add(compFieldListTranslator["BasementFinishedSQFT"], targetComp.BasementGLA());
            compSelectionBoxList.Add(compFieldListTranslator["BasementAccessType"], "%in");
            compCheckboxList.Add(compFieldListTranslator["HasForcedWarmAirForHeat"], "YES");
            compCheckboxList.Add(compFieldListTranslator["HasCentralAir"], "YES");


            foreach (string field in compTextFieldList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:/Order/OrderEdit* ATTR=NAME:SalesComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compTextFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compSelectionBoxList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=ACTION:/Order/OrderEdit* ATTR=NAME:SalesComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compSelectionBoxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }


            foreach (string field in compRadioButtonList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ACTION:/Order/OrderEdit* ATTR=NAME:SalesComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compRadioButtonList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compCheckboxList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEdit* ATTR=NAME:SalesComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compCheckboxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }
      
    
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEditWizardFNMA/step3/11488927 ATTR=NAME:SalesComp1.HasForcedWarmAirForHeat CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEditWizardFNMA/step3/11488927 ATTR=NAME:SalesComp1.HasCentralAir CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=ACTION:/Order/OrderEditWizardFNMA/step3/11488927 ATTR=CLASS:mceIframeContainer<SP>mceFirst<SP>mceLast");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:/Order/OrderEditWizardFNMA/step3/11488927 ATTR=NAME:SaveButton");  
        }

        protected void helper_CompFillHeaderScript()
        {
            theMacro.Clear();
            theMacro.AppendLine(@"SET !ERRORIGNORE YES");
            theMacro.AppendLine(@"SET !TIMEOUT_STEP 0");
            theMacro.AppendLine(@"SET !REPLAYSPEED FAST");
        }



        public void Prefill(iMacros.App iim, Form1 form)
        {
            // iim.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string currentUrl = iim.iimGetLastExtract();

            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 1");
            macro.Append(GenerateSubjectFillScript());

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 120);
        }



        public virtual void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {
            //http://www.marktomarket.us/Order/OrderEditWizardFNMA/Step3/11487827
            iim.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string currentUrl = iim.iimGetLastExtract();

            targetCompNumber = Regex.Match(compNum, @"\d").Value;



            if (currentUrl.ToLower().Contains("fnma"))
            {

                //we are on fannie mae form
                //step 3 = sales comps
                //step 4 = active listings comps
                if (saleOrList == "sale")
                {
                    iim.iimPlayCode(@"TAG POS=1 TYPE=INPUT:SUBMIT ATTR=ID:step3_submit");
                    GenerateSaleCompFillScript();

                }
                else
                {
                    iim.iimPlayCode(@"TAG POS=1 TYPE=INPUT:SUBMIT ATTR=ID:step4_submit");
                    GenerateListCompFillScript();

                }

                WriteScript(fieldList["filepath"], compNum + ".iim", theMacro);

                string macroCode = theMacro.ToString();
                iim.iimPlayCode(macroCode, 60);



            }



        }

    

    }
}
