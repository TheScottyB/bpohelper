using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{

    class Resnet : BPOFulfillment
    {

        public Resnet()
        {
            subject = GlobalVar.theSubjectProperty;
            callingForm = GlobalVar.mainWindow;
        }

        public Resnet(MLSListing m)
            : this()
        {
            targetComp = m;
        }

        enum elem { TEXT, SELECT, TEXTAREA, RADIO, CHECKBOX };

        private Dictionary<elem, string> inputType = new Dictionary<elem, string>()
        {
            {elem.TEXT, "INPUT:TEXT"},{elem.SELECT, "SELECT"}, {elem.TEXTAREA, "TEXTAREA"}, {elem.RADIO, "INPUT:RADIO"}, {elem.CHECKBOX, "INPUT:CHECKBOX"}
        };

        Dictionary<string, string> fifthThirdForm_buildingType = new Dictionary<string, string>()
            {
                {"1 Story", "1story"},  {"1.5 Story", "1story"}, {"2 Stories", "2story"}, {"Hillside", "split"}, {"Raised Ranch", "split"}, {@"Split Level w/ Sub", "split"}, {"Split Level", "split"},
                    {"Condo", "1story"}, {"Townhouse-2 Story", "2story"}, {"Townhouse", "2story"},  {"Townhouse 3+ Stories", "2story"}, {"Townhouse-TriLevel", "2story"}, {"Townhouse‐Ranch", "1story"}  //Townhouse 3+ Stories, Townhouse-TriLevel
            };

        private string targetCompString;
        private Dictionary<int, MLSListing> stack;
        private Form1 callingForm;
        private SubjectProperty subject;

        public void Prefill(iMacros.App iim, Form1 form)
        {
            Dictionary<double, string[]> commands = new Dictionary<double, string[]>();
            Dictionary<string, string> translationTable = new Dictionary<string, string>();

            Dictionary<string, string> fifthThirdForm_propertyType = new Dictionary<string, string>()
            {
                {"Detached", "5"}, {"Attached", "10"}
            };

            Dictionary<string, string> fifthThirdForm_locationType = new Dictionary<string, string>()
            {
                {location.Urban.ToString(), "%1"}, {location.Suburban.ToString(), "%2"}, {location.Rural.ToString(), "%3"}
            };

          
            Dictionary<string, string> occupancyTypeTranslation = new Dictionary<string, string>()
            {
                {"Occupied", "%2"}, {"Occupied by Owner", "%2"}, {"Occupied by Tenant", "%2"},
                    {"Vacant", "%10000"}, {"Other", "%2"}, {"Unknown", "%1"}
            };

            StringBuilder macro = new StringBuilder();

            //
            //Pass 2 new items
            //
            //AsIsOrRepaired
            commands.Add(2.001, new string[] { "1", inputType[elem.SELECT], "AsIsOrRepaired", "%True" });

            //ReviewersComment
            commands.Add(2.002, new string[] { "1", inputType[elem.TEXTAREA], "ReviewersComment", "No repairs were noted at the time of inspection based on exterior observations." });

            //CheckListComments
            commands.Add(2.003, new string[] { "1", inputType[elem.TEXTAREA], "CheckListComments", "All the comps are Reasonable substitute for the subject property, similar in most areas. Price opinion was based off comparable statistics to determine market price for subject. " });



            //InspectedDate
            commands.Add(0.01, new string[] { "1", inputType[elem.TEXT], "InspectedDate", GlobalVar.theSubjectProperty.InspectionDate() });
             commands.Add(0.02, new string[] { "1", inputType[elem.TEXT], "InspectedDate", GlobalVar.theSubjectProperty.InspectionDate() });
            //PropertyType - 10:condo; 5:Single Family
            commands.Add(0.03, new string[] { "1", inputType[elem.SELECT], "PropertyType", "%" + fifthThirdForm_propertyType[GlobalVar.theSubjectProperty.TypeOfMlsListing] });
            //ParcelNumber
            commands.Add(0.04, new string[] { "1", inputType[elem.TEXT], "ParcelNumber", GlobalVar.theSubjectProperty.ParcelID });
            //BpoBrokerLicenseNumber
            commands.Add(0.05, new string[] { "1", inputType[elem.TEXT], "BpoBrokerLicenseNumber", "971.009163" });
            //MarketConditions - Stable
            commands.Add(0.06, new string[] { "1", inputType[elem.SELECT], "MarketConditions", "%Stable" });
            //EmploymentConditions - Stable
            commands.Add(0.07, new string[] { "1", inputType[elem.SELECT], "EmploymentConditions", "%Stable" });
            //SalesLow
            commands.Add(0.08, new string[] { "1", inputType[elem.TEXT], "SalesLow", callingForm.SubjectNeighborhood.minSalePrice.ToString() });
            //SalesHigh
            commands.Add(0.09, new string[] { "1", inputType[elem.TEXT], "SalesHigh", callingForm.SubjectNeighborhood.maxSalePrice.ToString() });
            //MarketingDays
            commands.Add(0.10, new string[] { "1", inputType[elem.TEXT], "MarketingDays", callingForm.SubjectNeighborhood.avgDom.ToString() });

            //IsSubjectListed
            //TODO: put in actual listing info instead of defaulting to no.  
            commands.Add(0.11, new string[] { "1", inputType[elem.SELECT], "IsSubjectListed", "%False" });
            commands.Add(0.12, new string[] { "1", inputType[elem.SELECT], "IsSubjectListedPast12Mo", "%False" });

            //MandatoryAssociation
            //TODO: put in actual HOA info instead of deaulting t no
            commands.Add(0.13, new string[] { "1", inputType[elem.SELECT], "MandatoryAssociation", "%False" });

            //SubjectOccupancyStatus
            //occupancyTypeTranslation
            commands.Add(0.14, new string[] { "1", inputType[elem.SELECT], "SubjectOccupancyStatus", occupancyTypeTranslation[form.comboBoxSubjectOccupancy.Text] });

            //Condition
            commands.Add(0.15, new string[] { "1", inputType[elem.SELECT], "Condition", "%" + form.comboBoxSubjectCondition.Text });

            //LocationType: 1,2,3
            commands.Add(0.16, new string[] { "1", inputType[elem.SELECT], "LocationType", fifthThirdForm_locationType[form.comboBoxLocationDescr.Text] });

            //MarketPriceStability
            //MarketPriceStabilityPercent
            commands.Add(0.17, new string[] { "1", inputType[elem.SELECT], "MarketPriceStability", "%Stable" });
            commands.Add(0.18, new string[] { "1", inputType[elem.TEXT], "MarketPriceStabilityPercent", "0" });

            //Taxes
            commands.Add(0.19, new string[] { "1", inputType[elem.TEXT], "Taxes", callingForm.subjectTaxAmountTextBox.Text });

            //TEXTAREA PropertyComments
            //ListingComments
            //NeighDetailsAddnlComments
            commands.Add(0.20, new string[] { "1", inputType[elem.TEXTAREA], "PropertyComments", " No adverse conditions were noted at the time of inspection based on exterior observations. Unable to determine exterior or interior condition due to unable to view property from road, so subject was assumed to be in average condition. " });
            commands.Add(0.21, new string[] { "1", inputType[elem.TEXTAREA], "ListingComments", "NA" });
            commands.Add(0.22, new string[] { "1", inputType[elem.TEXTAREA], "NeighDetailsAddnlComments", "Neighborhood defined as 1 mile radius, centered on subject." });

            //DataSource
            commands.Add(0.23, new string[] { "1", inputType[elem.SELECT], "DataSource", "%assessor" });

            //Style
            //fifthThirdForm_buildingType
            commands.Add(0.24, new string[] { "1", inputType[elem.SELECT], "Style", "%" + fifthThirdForm_buildingType[callingForm.SubjectMlsType] });

            //LotSizeUnitType
            commands.Add(0.25, new string[] { "1", inputType[elem.SELECT], "LotSizeUnitType", "%Acres" });

            //LotSize
            commands.Add(0.26, new string[] { "1", inputType[elem.TEXT], "LotSize", callingForm.SubjectLotSize });

            //SubjectPropertyUnits
            commands.Add(0.27, new string[] { "1", inputType[elem.TEXT], "SubjectPropertyUnits", "1" });

            //subject property parameters
            commands.Add(0.28, new string[] { "1", inputType[elem.TEXT], "YearBuilt", callingForm.SubjectYearBuilt });
            commands.Add(0.29, new string[] { "1", inputType[elem.TEXT], "Rooms", callingForm.SubjectRoomCount });
            commands.Add(0.30, new string[] { "1", inputType[elem.TEXT], "Beds", callingForm.SubjectBedroomCount });
            commands.Add(0.31, new string[] { "1", inputType[elem.TEXT], "Baths", callingForm.SubjectBathroomCount });
            commands.Add(0.32, new string[] { "1", inputType[elem.TEXT], "GrossLivingArea", callingForm.SubjectAboveGLA });
            commands.Add(0.33, new string[] { "1", inputType[elem.TEXT], "Bsmt", callingForm.SubjectBasementType });

            //ConstructionType
            //TODO: Pick correct siding instead of defaulting to Mix
            commands.Add(0.34, new string[] { "1", inputType[elem.SELECT], "ConstructionType", "%4" });

            //GarageType
            if (GlobalVar.theSubjectProperty.AttachedGarage)
            {
                commands.Add(35, new string[] { "1", inputType[elem.SELECT], "GarageType", "%2" });
            }
            else if (GlobalVar.theSubjectProperty.DetachedGarage)
            {
                commands.Add(35, new string[] { "1", inputType[elem.SELECT], "GarageType", "%3" });
            }
            else
            {
                commands.Add(35, new string[] { "1", inputType[elem.SELECT], "GarageType", "%4" });
            }
            commands.Add(36, new string[] { "1", inputType[elem.TEXT], "GarageCount", GlobalVar.theSubjectProperty.GarageStallCount });

            //PoolSpa
            //TODO: Add pool support to form
            commands.Add(0.37, new string[] { "1", inputType[elem.TEXT], "PoolSpa", "Unk" });

            //FunctionalUtility
            //TODO: figure out wtf they want in this field
            commands.Add(0.38, new string[] { "1", inputType[elem.TEXT], "FunctionalUtility", "na" });

            //CheckListSubjectComments
            commands.Add(0.39, new string[] { "1", inputType[elem.TEXTAREA], "CheckListSubjectComments", "The subject is a conforming home within the neighborhood. No adverse conditions were noted at the time of inspection based on exterior observations. Unable to determine interior condition due to exterior inspection only, so subject was assumed to be in average condition for this report." });

            //AsIsValue
            //AsisListPrice
            //RepairedValue
            //RepairedListPrice
            commands.Add(0.41, new string[] { "1", inputType[elem.TEXT], "AsIsValue", callingForm.SubjectMarketValue });
            commands.Add(0.411, new string[] { "1", inputType[elem.TEXT], "AsisListPrice", callingForm.SubjectMarketValueList });
            commands.Add(0.42, new string[] { "1", inputType[elem.TEXT], "RepairedValue", callingForm.SubjectMarketValue });
            commands.Add(0.43, new string[] { "1", inputType[elem.TEXT], "RepairedListPrice", callingForm.SubjectMarketValueList });

            //CheckList01
            //CheckList02
            //CheckList03
            //CheckList04
            //CheckList05
            //DecliningValues
            //TODO:  Find correct answers using the data from the search. Possibly move this part to comp fill.
            commands.Add(0.44, new string[] { "1", inputType[elem.SELECT], "CheckList01", "%True" });
            commands.Add(0.45, new string[] { "1", inputType[elem.SELECT], "CheckList02", "%False" });
            commands.Add(0.46, new string[] { "1", inputType[elem.SELECT], "CheckList03", "%False" });
            commands.Add(0.47, new string[] { "1", inputType[elem.SELECT], "CheckList04", "%False" });
            commands.Add(0.48, new string[] { "1", inputType[elem.SELECT], "CheckList05", "%False" });
            commands.Add(0.49, new string[] { "1", inputType[elem.SELECT], "DecliningValues", "%False" });

            //ProviderComments
            commands.Add(0.50, new string[] { "1", inputType[elem.TEXTAREA], "ProviderComments", "The subject is a conforming home within a neighborhood that has stable values over the prior twelve months. Demand remains strong in this area while short sales and REO listings have declined and had no significant impact on values in area." });



            commands.Add(0, new string[] { "1", inputType[elem.RADIO], "CMA_TYPE_EXTERIOR", "YES" });
            commands.Add(1, new string[] { "1", inputType[elem.TEXT], "CMA_DATE", DateTime.Now.ToShortDateString() });
            commands.Add(2, new string[] { "1", inputType[elem.TEXT], "SUBJ_STYLE", "SFR" });
            commands.Add(3, new string[] { "1", inputType[elem.TEXT], "SUBJ_BEDROOMS", callingForm.SubjectBedroomCount });
            commands.Add(4, new string[] { "1", inputType[elem.TEXT], "SUBJ_SQFT", callingForm.SubjectAboveGLA });
            commands.Add(5, new string[] { "1", inputType[elem.TEXT], "SUBJ_AGE", (DateTime.Now.Year - Convert.ToInt32(callingForm.SubjectYearBuilt)).ToString() });
            commands.Add(6, new string[] { "1", inputType[elem.TEXT], "SUBJ_BATHS", callingForm.SubjectBathroomCount });
            commands.Add(7, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_NBRD_CONFORM_AVERAGE", "YES" });
            commands.Add(8, new string[] { "1", inputType[elem.CHECKBOX], "RPR_RECOMMEND_REPAIRS_NO", "YES" });
            commands.Add(9, new string[] { "1", inputType[elem.TEXT], "RECN_APPR_SIGN_DATE", DateTime.Now.ToShortDateString() });
            commands.Add(10, new string[] { "1", inputType[elem.TEXT], "RECN_APPR_PHONE", "815.315.0203" });
            commands.Add(11, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_EXT_CONDITION_AVERAGE", "YES" });
            commands.Add(12, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_LANDSCAPE_AVERAGE", "YES" });
            commands.Add(13, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_AREA_AVERAGE", "YES" });
            commands.Add(14, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_ROOF_AVERAGE", "YES" });
            commands.Add(15, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_OWNER_PRIDE_AVERAGE", "YES" });
            commands.Add(16, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_GEN_CONDITION_AVERAGE", "YES" });
            commands.Add(17, new string[] { "1", inputType[elem.TEXT], "NBR_PROP_TYPE", "SFR" });
            commands.Add(18, new string[] { "1", inputType[elem.CHECKBOX], "NBR_TREND_STABLE", "YES" });
            commands.Add(19, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_PROP_OCCUPIED_OCCUPIED", "YES" });
            commands.Add(20, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_CURRENTLY_LISTED_NO", "YES" });


            // commands.Add(5, new string[] { "1", inputType[elem.TEXT], "LIST_COUNT", callingForm.SubjectNeighborhood.numberActiveListings.ToString() });


            //commands.Add(5, new string[] { "1", inputType[elem.TEXT], "LIST_COUNT", callingForm.SubjectNeighborhood.numberActiveListings.ToString() });
            //commands.Add(6, new string[] { "1", inputType[elem.TEXT], "PROP_FOR_SALE_L", callingForm.SubjectNeighborhood.numberOfCompListings.ToString()});
            //commands.Add(7, new string[] { "1", inputType[elem.TEXT], "REO_CORP_OWNED_NBR", callingForm.SetOfComps.numberREOListings.ToString() });
            //commands.Add(8, new string[] { "1", inputType[elem.TEXT], "LIST_LOW_S", callingForm.SubjectNeighborhood.minListPrice.ToString() });
            //commands.Add(9, new string[] { "1", inputType[elem.TEXT], "LIST_HIGH_S", callingForm.SubjectNeighborhood.maxListPrice.ToString() });
            //commands.Add(10, new string[] { "1", inputType[elem.TEXT], "SALE_COUNT", callingForm.SubjectNeighborhood.numberSoldListings.ToString() });

            //commands.Add(12, new string[] { "1", inputType[elem.TEXT], "SALE_HIGH_S", });
            //commands.Add(13, new string[] { "1", inputType[elem.TEXT], "BOARDED_UP_HOMES_NBR", "0" });
            //commands.Add(14, new string[] { "1", inputType[elem.TEXT], "DOM_AVG_L",  });
            //commands.Add(15, new string[] { "1", inputType[elem.TEXT], "ZONING", "A1-R1" });
            //commands.Add(16, new string[] { "1", inputType[elem.TEXT], "ZONING", "A1-R1" });
            //commands.Add(17, new string[] { "1", inputType[elem.TEXT], "ZONING", "A1-R1" });


            // if (callingForm.comboBoxSubjectOccupancy.Text == "Occupied")
            //{

            //}

            //Occupied by owner
            //     commands.Add(5, new string[] { "1", inputType[elem.RADIO], "OCCUPANCY", "YES" });


            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:apn_s CONTENT=11111");
            //macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=NAME:BPO_FORM ATTR=TXT:*<SP>Property<SP>Type:");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:PROPERTYTYPE CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:ZONING CONTENT=111");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:ZONING_DESC_1 CONTENT=11");
            //macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=NAME:BPO_FORM ATTR=TXT:Legal<SP>Legal<SP>Nonconforming<SP>(Grandfathered)<SP>No<SP>Zoning<SP>Illegal");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:BPO_FORM ATTR=NAME:SUBJECT_USE_CODE_1 CONTENT=%L");
            //targetComp = subject.GetCurrentMlsListing();
            //CompFill(iim, "subject", "1", translationTable);

            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");

            foreach (var c in commands)
            {
                macro.AppendFormat("TAG POS={0} TYPE={1} ATTR=ID:*{2} CONTENT={3}\r\n", c.Value[0], c.Value[1], c.Value[2], c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            }




            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 30);
        }

        public void CompFill(iMacros.App iim, string saleOrList, string compString, Dictionary<string, string> fieldList)
        {
            StringBuilder macro = new StringBuilder();
            targetCompNumber = Regex.Match(compString, @"\d").Value;
            targetCompString = compString;


            Dictionary<string, string> saleTypeTranslation = new Dictionary<string, string>()
            {
                {"Arms Length", "1"}, {"REO", "3"}, {"ShortSale", "2"},
            };

            Dictionary<string, string> fifthThirdForm_saleTypeTranslation = new Dictionary<string, string>()
            {
                {"Arms Length", "4"}, {"REO", "3"}, {"ShortSale", "2"},
            };

            Dictionary<string, string> fifthThirdForm_propertyType = new Dictionary<string, string>()
            {
                {"Detached", "5"}, {"Attached", "10"}
            };

       
            Dictionary<string, string> financingTypeTranslation = new Dictionary<string, string>()
            {
                {"Conventional", "%number:393"}, {"FHA", "%number:394"}, {"VA", "%number:396"},  {"Unknown", "%number:395"},

            };
            Dictionary<string, string> heatingTypeTranslation = new Dictionary<string, string>()
            {
                {"Gas", "6"}, {"Gas, Forced Air", "6"}, {@"Gas, Hot Water/Steam", "6"},{@"Electric", "2"},

            };

            Dictionary<string, string> CoolingTypeTranslation = new Dictionary<string, string>()
            {
                {"Central Air", "1"},  {"None", "0"}
            };

            Dictionary<string, string> ListingDatePosition = new Dictionary<string, string>()
            {
                {"Sold1", "3"},  {"Sold2", "4"}, {"Sold3", "5"},  {"List1", "6"},  {"List2", "7"},  {"List3", "8"}
            };
            Dictionary<string, string> SoldDatePosition = new Dictionary<string, string>()
            {
                {"Sold1", "10"},  {"Sold2", "11"}, {"Sold3", "12"},  {"List1", "0"},  {"List2", "0"},  {"List3", "0"}
            };




            Dictionary<double, string[]> commands = new Dictionary<double, string[]>();

            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");




            //
            //Pass 2
            //
            //ConstructionType
            //TODO: Pick correct siding instead of defaulting to Mix
            commands.Add(2.001, new string[] { "1", inputType[elem.SELECT], "ConstructionType", "%4" });

            //PoolSpa
            //TODO: Add pool support to form
            commands.Add(2.002, new string[] { "1", inputType[elem.TEXT], "PoolSpa", "na" });

            //FunctionalUtility
            //TODO: figure out wtf they want in this field
            commands.Add(2.003, new string[] { "1", inputType[elem.TEXT], "FunctionalUtility", "na" });

            //Comparability
            commands.Add(2.004, new string[] { "1", inputType[elem.SELECT], "Comparability", "%2" });

            //SaleType
            //fix command 5

            //CheckListListing1Comments
            //CheckListSales1Comments
            //FORM=ACTION:/ProviderResponse/FormStaticFifthThirdView/* ATTR=NAME
            string compComments = "Similar size, age, style, features, condition and neighborhood as subject.".Replace(" ", "<SP>");

            int tcn = Convert.ToInt32(targetCompNumber) + 1;
            if (saleOrList == "Sale")
            {
                
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ACTION:/ProviderResponse/FormStaticFifthThirdView/* ATTR=NAME:CheckListSales" + tcn.ToString() + "Comments CONTENT=" + compComments);

            }
            else
            {
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ACTION:/ProviderResponse/FormStaticFifthThirdView/* ATTR=NAME:CheckListListing" + tcn.ToString() + "Comments CONTENT=" + compComments);

            }

            //proximity fix 
            //command 1


            commands.Add(0, new string[] { "1", inputType[elem.TEXT], "StreetAddress", targetComp.StreetAddress.Substring(0, Math.Min(targetComp.StreetAddress.Length, 17)) });
            commands.Add(0.1, new string[] { "1", inputType[elem.TEXT], "City", targetComp.City.Substring(0, Math.Min(targetComp.City.Length, 17)).Replace(" ", "<SP>") });
            commands.Add(0.2, new string[] { "1", inputType[elem.SELECT], "State", "%IL" });
            commands.Add(0.3, new string[] { "1", inputType[elem.TEXT], "PostalCode", targetComp.Zipcode });

            double proximity = targetComp.proximityToSubject;

            if (proximity > 1) 
            {
                proximity = 1; 
            }

            commands.Add(1, new string[] { "1", inputType[elem.TEXT], "ProximityDistance", proximity.ToString() });

            commands.Add(2, new string[] { "1", inputType[elem.TEXT], "SalePrice", targetComp.SalePrice.ToString() });
            commands.Add(3, new string[] { "1", inputType[elem.TEXT], "ListPriceAtSale", targetComp.CurrentListPrice.ToString() });

            commands.Add(4, new string[] { "1", inputType[elem.TEXT], "PriceReductionCount", targetComp.NumberOfPriceChanges.ToString() });

            commands.Add(5, new string[] { "1", inputType[elem.SELECT], "SaleType", "%" + fifthThirdForm_saleTypeTranslation[targetComp.TransactionType] });
            commands.Add(6, new string[] { "1", inputType[elem.SELECT], "DataSource", "%mls" });
            commands.Add(7, new string[] { "1", inputType[elem.TEXT], "SaleDate", _getMmYyyyDdDateString(targetComp.SalesDate)});
            commands.Add(7.1, new string[] { "1", inputType[elem.TEXT], "SaleDate", _getMmYyyyDdDateString(targetComp.SalesDate) });
            commands.Add(8, new string[] { "1", inputType[elem.TEXT], "Dom", targetComp.DOM });
            //
            //PropertySoldDom
            commands.Add(8.1, new string[] { "1", inputType[elem.TEXT], "PropertySoldDom", targetComp.DOM });
            //CurrentListPriceDom
            commands.Add(8.2, new string[] { "1", inputType[elem.TEXT], "CurrentListPriceDom", targetComp.DOM });


            commands.Add(9, new string[] { "1", inputType[elem.TEXT], "Concessions", targetComp.PointsInDollars });
            commands.Add(10, new string[] { "1", inputType[elem.SELECT], "FeeSimple", "%2" });
            //commands.Add(11, new string[] { "1", inputType[elem.TEXT], "LotSize", targetComp.Lotsize.ToString().Substring(0,4)});

            string lotSize = targetComp.Lotsize.ToString();
            if (callingForm.SubjectAttached)
            {
                lotSize = "0";
            }
            commands.Add(11, new string[] { "1", inputType[elem.TEXT], "LotSize", lotSize});
            commands.Add(12, new string[] { "1", inputType[elem.TEXT], "Units", "1" });
            commands.Add(13, new string[] { "1", inputType[elem.TEXT], "ViewDescription", "Residential" });
            commands.Add(14, new string[] { "1", inputType[elem.SELECT], "Condition", "%Average" });
            commands.Add(15, new string[] { "1", inputType[elem.TEXT], "Rooms", targetComp.TotalRoomCount.ToString() });
            commands.Add(16, new string[] { "1", inputType[elem.TEXT], "Beds", targetComp.BedroomCount });
            commands.Add(17, new string[] { "1", inputType[elem.TEXT], "Baths", targetComp.BathroomCount });
            commands.Add(18, new string[] { "1", inputType[elem.TEXT], "GrossLivingArea", targetComp.ProperGla(GlobalVar.mainWindow.SubjectAboveGLA) });
            commands.Add(19, new string[] { "1", inputType[elem.TEXT], "Bsmt", targetComp.BasementType + "/" + targetComp.NumberOfBasementRooms() });
            commands.Add(20, new string[] { "1", inputType[elem.TEXT], "HeatingCooling", targetComp.Heating + "/" + targetComp.Cooling });
            if (targetComp.AttachedGarage())
            {
                commands.Add(21, new string[] { "1", inputType[elem.SELECT], "GarageType", "%2" });
            }
            else if (targetComp.DetachedGarage())
            {
                commands.Add(21, new string[] { "1", inputType[elem.SELECT], "GarageType", "%3" });
            } else
            {
                commands.Add(21, new string[] { "1", inputType[elem.SELECT], "GarageType", "%4" });
            }
            commands.Add(22, new string[] { "1", inputType[elem.TEXT], "GarageCount", targetComp.NumberGarageStalls() });
            commands.Add(23, new string[] { "1", inputType[elem.TEXT], "Extras", "NA" });
            commands.Add(24, new string[] { "1", inputType[elem.SELECT], "LocationType", "%2" });
            commands.Add(25, new string[] { "1", inputType[elem.SELECT], "Appeal", "%Average" });
            commands.Add(26, new string[] { "1", inputType[elem.TEXT], "YearBuilt", targetComp.YearBuiltString });
            commands.Add(27, new string[] { "1", inputType[elem.TEXT], "ListPrice", targetComp.CurrentListPrice.ToString() });
            commands.Add(28, new string[] { "1", inputType[elem.TEXT], "OriginalListPrice", targetComp.OriginalListPrice.ToString() });

            //second update pass adding more fields
            commands.Add(29, new string[] { "1", inputType[elem.SELECT], "SaleType", "%" + fifthThirdForm_saleTypeTranslation[targetComp.TransactionType] });
            //PropertyType
            commands.Add(30, new string[] { "1", inputType[elem.SELECT], "PropertyType", "%" + fifthThirdForm_propertyType[targetComp.UniversalLandUse()] });
            //Style
            commands.Add(31, new string[] { "1", inputType[elem.SELECT], "Style", "%" + fifthThirdForm_buildingType[targetComp.Type]});

            //commands.Add(14, new string[] { "1", inputType[elem.SELECT], "ConstructionTypeID", "%number:430" });



            //commands.Add(24, new string[] { "1", inputType[elem.SELECT], "HEATING", "%" + heatingTypeTranslation[targetComp.Heating] });
            //commands.Add(25, new string[] { "1", inputType[elem.SELECT], "COOLING", "%" + CoolingTypeTranslation[targetComp.Cooling] });
           
         
            //commands.Add(4, new string[] { "1", inputType[elem.TEXT], "TotalRents", "1700" });
            //commands.Add(5, new string[] { ListingDatePosition[targetCompString], inputType[elem.TEXT], "dt", targetComp.ListDateString });
   
           
            
            //commands.Add(8.1, new string[] { "1", inputType[elem.SELECT], "FinanceTypeID", financingTypeTranslation[targetComp.FinancingMlsString] });
      

            //commands.Add(13, new string[] { "1", inputType[elem.SELECT], "ConditionTypeID", "%number:432" });
            //commands.Add(14, new string[] { "1", inputType[elem.SELECT], "ConstructionTypeID", "%number:430" });
            //commands.Add(18, new string[] { "1", inputType[elem.TEXT], "HalfBaths", targetComp.HalfBathCount });
            ////commands.Add(19, new string[] { "1", inputType[elem.TEXT], "AboveGradeSqft", targetComp.GLA.ToString() });
            //commands.Add(20, new string[] { "1", inputType[elem.TEXT], "BelowGradeSqft", targetComp.BasementGLA() });
            //commands.Add(21, new string[] { "1", inputType[elem.TEXT], "BasementPercentFinished", targetComp.BasementFinishedPercentage() });
            //commands.Add(23, new string[] { "1", inputType[elem.TEXT], "MLSNumber", targetComp.MlsNumber });


            //commands.Add(6, new string[] { "1", inputType[elem.TEXT], "SELLER_CONCESSIONS", targetComp.PointsInDollars });
            //commands.Add(6, new string[] { "1", inputType[elem.TEXT], "AGE", targetComp.Age.ToString() });
            //commands.Add(7, new string[] { "1", inputType[elem.TEXT], "LOT_SIZE", (targetComp.Lotsize * 43560).ToString() });

            //if (saleOrList == "sale")
            //{

            //  

            //    commands.Add(10.13, new string[] { "2", inputType[elem.CHECKBOX], "COMMENTS", "YES" });
            //}
            //if (saleOrList == "list")
            //{

            //    commands.Add(10.21, new string[] { "1", inputType[elem.TEXT], "FINANCING", "UTD" });
            //    commands.Add(10.22, new string[] { "2", inputType[elem.CHECKBOX], "RATING", "YES" });
            //}



            //commands.Add(4, new string[] { "1", inputType[elem.TEXT], "PRICE_REDUCTION_COUNT", targetComp.NumberOfPriceChanges.ToString() });

            //commands.Add(6, new string[] { "1", inputType[elem.TEXT], "SELLER_CONCESSIONS", targetComp.PointsInDollars });
           



            //commands.Add(9, new string[] { "1", inputType[elem.SELECT], "LOCATION_DESCR", "%" + GlobalVar.mainWindow.comboBoxLocationDescr.Text });

            //commands.Add(12, new string[] { "1", inputType[elem.SELECT], "LOTSIZE_TYPE_S", "%A" });

            //if (targetComp.Waterfront)
            //{
            //    commands.Add(13, new string[] { "1", inputType[elem.TEXT], "PROPERTY_VIEW", "Water" });
            //}
            //else
            //{
            //    commands.Add(13, new string[] { "1", inputType[elem.TEXT], "PROPERTY_VIEW", "Residential" });
            //}
            //commands.Add(16, new string[] { "1", inputType[elem.SELECT], "CONDITION", "%Average" });



            //commands.Add(20, new string[] { "1", inputType[elem.TEXT], "BATHS_NUM", targetComp.BathroomCount });

            //commands.Add(21, new string[] { "1", inputType[elem.TEXT], "LIVING_SQFEET", targetComp.ProperGla(GlobalVar.mainWindow.SubjectAboveGLA) });
            //commands.Add(22, new string[] { "1", inputType[elem.TEXT], "BSMT_SQFEET_S", targetComp.BasementGLA() });
            //commands.Add(23, new string[] { "1", inputType[elem.TEXT], "ROOMS_BELOW_GRADE", targetComp.NumberOfBasementRooms() });


            //commands.Add(24, new string[] { "1", inputType[elem.SELECT], "HEATING", "%" + heatingTypeTranslation[targetComp.Heating] });
            //commands.Add(25, new string[] { "1", inputType[elem.SELECT], "COOLING", "%" + CoolingTypeTranslation[targetComp.Cooling] });

          

            //commands.Add(28, new string[] { "1", inputType[elem.SELECT], "COOLING", "%" + CoolingTypeTranslation[targetComp.Cooling] });

            //commands.Add(29, new string[] { "1", inputType[elem.TEXT], "OTHER", "NA" });
            //commands.Add(30, new string[] { "2", inputType[elem.RADIO], "POOL_B", "YES" });
            //if (saleOrList == "sale")
            //{
            //    commands.Add(31, new string[] { "1", inputType[elem.TEXT], "SALE_PRICE", targetComp.SalePrice.ToString() });
            //    commands.Add(32, new string[] { "1", inputType[elem.TEXT], "CURRENTLIST_PRICE", targetComp.CurrentListPrice.ToString() });

            //}
            //#ng-view > div > form > div:nth-child(1) > div > div.tab-pane.ng-scope.active > div:nth-child(9) > div:nth-child(2) > table > tbody > tr:nth-child(2) > td:nth-child(3) > input

            iim.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string currentUrl = iim.iimGetLastExtract();

            string tcs = targetCompString;

            if (!currentUrl.Contains("FormStaticFifthThirdView"))
            {
                foreach (var c in commands)
                {
                    //if (c.Value[2] == "dt")
                    //{
                    //    tcs = "";
                    //}
                    //else
                    //{
                    //    tcs = targetCompString;
                    //}
                    macro.AppendFormat("TAG POS={0} TYPE={1} FORM=ID:formBpo ATTR=NAME:{2}Comps[{3}].{4} CONTENT={5}\r\n", c.Value[0], c.Value[1], saleOrList, targetCompString, c.Value[2], c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
                }
            } else
            {
                foreach (var c in commands)
                {
                    //if (c.Value[2] == "dt")
                    //{
                    //    tcs = "";
                    //}
                    //else
                    //{
                    //    tcs = targetCompString;
                    //}

                    if (c.Value[2] == "SaleDate")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=BUTTON ATTR=TXT:Done");
                    }

                    macro.AppendFormat("TAG POS={0} TYPE={1} FORM=ACTION:/ProviderResponse/FormStaticFifthThirdView/* ATTR=NAME:{2}ComparableData[{3}].{4} CONTENT={5}\r\n", c.Value[0], c.Value[1], saleOrList, targetCompString, c.Value[2], c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));

                    if (c.Value[2] == "SaleDate")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=BUTTON ATTR=TXT:Done");
                    }

                  
                    

                }
            }
           

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 60);

            //commands.Add(1, new string[] { inputType, "tbCity", targetComp.City.Replace(" ", "<SP>") });
            //commands.Add(2, new string[] { "SELECT", "ddlState", "%IL" });
            //commands.Add(3, new string[] { inputType, "tbZip", targetComp.Zipcode });

            //commands.Add(4.5, new string[] { inputType, "tbProximity", targetComp.ProximityToSubject.ToString() });
            //commands.Add(5, new string[] { inputType, "tbNumUnits", "1" });

            //if (targetComp.TransactionType == "REO")
            //{
            //    commands.Add(8, new string[] { "SELECT", "ddlSalesType", "%REO" });
            //}
            //else if (targetComp.TransactionType.ToLower().Contains("short"))
            //{
            //    commands.Add(8, new string[] { "SELECT", "ddlSalesType", "%Short Sale" });
            //}
            //else
            //{
            //    commands.Add(8, new string[] { "SELECT", "ddlSalesType", "%Fair Market" });
            //}

            //if (targetComp.FinancingMlsString.ToLower().Contains("cash"))
            //{
            //    commands.Add(9, new string[] { "SELECT", "ddlFinanceType", "%Cash" });
            //}
            //else if (targetComp.FinancingMlsString.ToLower().Contains("conv"))
            //{
            //    commands.Add(9, new string[] { "SELECT", "ddlFinanceType", "%Conv" });
            //}
            //else if (targetComp.FinancingMlsString.ToLower().Contains("fha"))
            //{
            //    commands.Add(9, new string[] { "SELECT", "ddlFinanceType", "%FHA" });
            //}
            //else if (targetComp.FinancingMlsString.ToLower().Contains("va"))
            //{
            //    commands.Add(9, new string[] { "SELECT", "ddlFinanceType", "%VA" });
            //}
            //else
            //{
            //    commands.Add(9, new string[] { "SELECT", "ddlFinanceType", "%None" });
            //}

            //if (saleOrList == "sale")
            //{
            //    commands.Add(6, new string[] { inputType, "tbListPriceAtSale", targetComp.CurrentListPrice.ToString() });
            //    commands.Add(7, new string[] { inputType, "tbSalePrice", targetComp.SalePrice.ToString() });
            //    //tbSaleDate1
            //    commands.Add(31, new string[] { inputType, "tbSaleDate", targetComp.SalesDate.ToString() });
            //}

            //if (saleOrList == "list")
            //{
            //    //tbCurrentListPrice
            //    commands.Add(5.5, new string[] { inputType, "tbCurrentListPrice", targetComp.CurrentListPrice.ToString() });
            //    //tbListingDate
            //    commands.Add(32, new string[] { inputType, "tbListingDate", targetComp.ListDateString});
            //    //tbListingAgent
            //    commands.Add(33, new string[] { inputType, "tbListingAgent", targetComp.ListingAgentName });
            //    //tbListingAgentPhone
            //    commands.Add(34, new string[] { inputType, "tbListingAgentPhone", targetComp.ListingAgentNamePhone.Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "") });
            //}

            //commands.Add(10, new string[] { inputType, "tbOriginalListPrice", targetComp.OriginalListPrice.ToString() });
            //commands.Add(11, new string[] { inputType, "tbDom", targetComp.DOM });
            //commands.Add(12, new string[] { inputType, "tbTtlRoomCount", targetComp.TotalRoomCount.ToString() });
            //commands.Add(13, new string[] { inputType, "tbTtlBedrmCount", targetComp.BedroomCount });
            //commands.Add(14, new string[] { inputType, "tbTtlBathrmCount", targetComp.BathroomCount });
            //commands.Add(15, new string[] { inputType, "tbGla", targetComp.ProperGla(GlobalVar.mainWindow.SubjectAboveGLA) });
            //commands.Add(16, new string[] { inputType, "tbLotSize", targetComp.Lotsize.ToString()});
            //commands.Add(17, new string[] { inputType, "tbYearBuilt", targetComp.YearBuiltString });

            //commands.Add(19, new string[] { inputType, "tbBasementSize", targetComp.BasementGLA() });
            //commands.Add(20, new string[] { inputType, "tbSellerConcession", targetComp.PointsMlsString });
            //commands.Add(21, new string[] { "SELECT", "ddlStyle", "%Conv" });
            //commands.Add(22, new string[] { "SELECT", "ddlConstructionType", "%Vinyl" });
            //if (GlobalVar.mainWindow.SubjectAttached)
            //{
            //    commands.Add(23, new string[] { "SELECT", "ddlProperyType", "%Attached" });
            //}else
            //{
            //    commands.Add(23, new string[] { "SELECT", "ddlProperyType", "%Detached" });
            //}
            //commands.Add(25, new string[] { "SELECT", "ddlView", "%Residential" });
            //commands.Add(26, new string[] { "SELECT", "ddlCondition", "%Average" });
            //string lresGarageStr = "None";
            //string numSpaces = targetComp.NumberGarageStalls();
            //string att_det = "";

            //if (!string.IsNullOrEmpty(numSpaces))
            //{
            //    if (targetComp.AttachedGarage())
            //    {
            //        att_det = "Attached";
            //    }
            //    else if (targetComp.DetachedGarage())
            //    {
            //        att_det = "Detached";
            //    }

            //    switch (numSpaces)
            //    {
            //        case "1":
            //            lresGarageStr = "1<SP>Car<SP>" + att_det;
            //            break;
            //        case "2":
            //            lresGarageStr = "2<SP>Car<SP>" + att_det;
            //            break;
            //        default:
            //            lresGarageStr = "2+<SP>Car<SP>" + att_det;
            //            break;
            //    }

            //}
            //commands.Add(27, new string[] { "SELECT", "ddlParking", "%" + lresGarageStr });
            //commands.Add(28, new string[] { "SELECT", "ddlComparedToSubject", "%Equal" });


            //commands.Add(29, new string[] { "TEXTAREA", "tbCompComment", ".TBD." });




            //foreach (var c in commands)
            //{
            //    macro.AppendFormat("TAG POS=1 TYPE={0} FORM=ID:aspnetForm ATTR=ID:*_{2}{1} CONTENT={3}\r\n", c.Value[0], targetCompNumber.ToString(), c.Value[1], c.Value[2].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            //}
            ////default to the middle until we can suggest which one is the most similar
            //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$rbMostComparable CONTENT=YES");



            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$tbAsIsValue CONTENT=111");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$tbRepairedValue CONTENT=111");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$tbSuggestAsIsValue CONTENT=11");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$tbSuggestRepairedValue CONTENT=11");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$tbAsIsQuickSaleValue CONTENT=11");
            ////macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$ddlVendorLicense CONTENT=%64589");







        }

        public void MmrFill(iMacros.App iim, string saleOrList, string compString, Dictionary<string, string> fieldList)
        {

            Dictionary<double, string[]> commands = new Dictionary<double, string[]>();
            StringBuilder macro = new StringBuilder();

            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");

            commands.Add(1, new string[] { "1", inputType[elem.TEXT], "Address", targetComp.StreetAddress });
            commands.Add(2, new string[] { "1", inputType[elem.TEXT], "Prox", targetComp.proximityToSubject.ToString() });
            commands.Add(3, new string[] { "1", inputType[elem.TEXT], "Beds", targetComp.BedroomCount });
            commands.Add(4, new string[] { "1", inputType[elem.TEXT], "Baths", targetComp.BathroomCount });
            commands.Add(5, new string[] { "1", inputType[elem.TEXT], "SqFt", targetComp.GLA.ToString()});
            commands.Add(6, new string[] { "1", inputType[elem.TEXT], "Condition", "Avg" });
            if (saleOrList == "List")
            {
                commands.Add(7, new string[] { "1", inputType[elem.TEXT], "AskPrice", targetComp.CurrentListPrice.ToString() });
                commands.Add(8, new string[] { "1", inputType[elem.TEXT], "ListDt", targetComp.ListDateString });
                commands.Add(10, new string[] { "1", inputType[elem.TEXT], "Fin", "NA" });

            }
            else
            {
                commands.Add(7, new string[] { "1", inputType[elem.TEXT], "AskPrice", targetComp.SalePrice.ToString() });
                commands.Add(8, new string[] { "1", inputType[elem.TEXT], "ListDt", targetComp.SalesDate.ToShortDateString() });
                commands.Add(10, new string[] { "1", inputType[elem.TEXT], "Fin", targetComp.FinancingMlsString });

            }
            commands.Add(9, new string[] { "1", inputType[elem.TEXT], "DOM", targetComp.DOM });







            foreach (var c in commands)
            {
                macro.AppendFormat("TAG POS={0} TYPE={1} ATTR=NAME:*{2}{3}{4} CONTENT={5}\r\n", c.Value[0], c.Value[1], compString[0], c.Value[2], compString[1], c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            }





            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 30);
        }
    }
}
