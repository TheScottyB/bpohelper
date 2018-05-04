using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{

    class Sandcastle : BPOFulfillment
    {

        public Sandcastle()
        {
            subject = GlobalVar.theSubjectProperty;
            callingForm = GlobalVar.mainWindow;
        }

        public Sandcastle(MLSListing m)
            : this()
        {
            targetComp = m;
        }

        enum elem { TEXT, SELECT, TEXTAREA, RADIO, CHECKBOX };

        private Dictionary<elem, string> inputType = new Dictionary<elem, string>()
        {
            {elem.TEXT, "INPUT:TEXT"},{elem.SELECT, "SELECT"}, {elem.TEXTAREA, "TEXTAREA"}, {elem.RADIO, "INPUT:RADIO"}, {elem.CHECKBOX, "INPUT:CHECKBOX"}
        };

        Dictionary<string, string> propertyTypeTranslation = new Dictionary<string, string>()
            {
                {"Attached", "%2"}, {"Detached", "%1"}
            };
        
        Dictionary<string, string> homeStyleTranslation = new Dictionary<string, string>()
            {
                {"1 Story", "%24"}, {"1.5 Story", "%25"}, {"2 Stories", "%26"}, {"Raised Ranch", "%27"}, {"Split Level", "%28"}, {@"Split Level w/Sub", "%28"}
            };

        Dictionary<string, string> basementStyleTranslation = new Dictionary<string, string>()
            {
                {"None", "%61"}, {"Full", "%62"}, {"Full, Walkout", "%62"}, {"Partial", "%63"}, {"Partial, Walkout", "%63"}
            };

        Dictionary<string, string> garageTranslation = new Dictionary<string, string>()
            {
                {"Attached", "%37"}, {"None", "%36"}, 
                    {"Detached", "%38"}
            };

        Dictionary<string, string> saleTypeTranslation = new Dictionary<string, string>()
            {
                {"Arms Length", "%47"}, {"REO", "%49"}, {"ShortSale", "%48"},  {"Contract", "%50"}, {"Corp", "%50"}
            };

        Dictionary<string, string> locationTypeTranslation = new Dictionary<string, string>()
            {
                {"Rural", "%2"}, {"Suburban", "%3"}, {"Urban", "%1"}
            };

        Dictionary<string, string> occupancyTypeTranslation = new Dictionary<string, string>()
            {
                {"Occupied", "%number:436"}, {"Occupied by Owner", "%number:436"}, {"Occupied by Tenant", "%number:436"},
                    {"Vacant", "%number:437"}, {"Other", "%number:439"}, {"Unknown", "%number:439"}
            };
    
        Dictionary<string, string> typeOfOccupancyTranslation = new Dictionary<string, string>()
            {
                {"Occupied", "%42"}, {"Occupied by Owner", "%42"}, {"Occupied by Tenant", "%43"},
                    {"Vacant", "%44"}, {"Other", "%45"}, {"Unknown", "%46"}
            };


        private string targetCompString;
        private Dictionary<int, MLSListing> stack;
        private Form1 callingForm;
        private SubjectProperty subject;

        public void Prefill(iMacros.App iim, Form1 form)
        {
            Dictionary<double, string[]> commands = new Dictionary<double, string[]>();
            Dictionary<string, string> translationTable = new Dictionary<string, string>();
            StringBuilder macro = new StringBuilder();

            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 1");
           
          

            //
            //Page 1 - Property Info Section
            //
            commands.Add(0.001, new string[] { "1", inputType[elem.TEXT], "fld_Date_Visit1", form.dateTimePickerInspectionDate.Value.ToShortDateString() });
            commands.Add(0.002, new string[] { "1", inputType[elem.TEXT], "fld_Time_Visit1", "13:00" });
            commands.Add(0.003, new string[] { "1", inputType[elem.SELECT], "fld_Occupied_By", typeOfOccupancyTranslation[form.comboBoxSubjectOccupancy.Text] });
            //TODO: Select correct listed state
            //default not listed in last 12 months
            commands.Add(0.004, new string[] { "1", inputType[elem.SELECT], "fld_Property_For_Sale", "%0" });

            //
            //Area Information
            //
            commands.Add(0.005, new string[] { "1", inputType[elem.SELECT], "fld_Location", locationTypeTranslation[form.comboBoxLocationDescr.Text] });
            //TODO: Select correct supply/demand state 
            //default in-balance
            commands.Add(0.006, new string[] { "1", inputType[elem.SELECT], "fld_Supply", "%5" });
            //TODO: Select correct state 
            //default stable
            commands.Add(0.007, new string[] { "1", inputType[elem.SELECT], "fld_Property_Values", "%8" });
            commands.Add(0.008, new string[] { "1", inputType[elem.SELECT], "fld_Ownership", "%14" });
            commands.Add(0.009, new string[] { "1", inputType[elem.TEXT], "fld_Properties_for_Sale", callingForm.SubjectNeighborhood.numberActiveListings.ToString() });
            commands.Add(0.011, new string[] { "1", inputType[elem.TEXT], "fld_Neighborhood_Days_on_Market", callingForm.SubjectNeighborhood.avgDom.ToString() });
            commands.Add(0.012, new string[] { "1", inputType[elem.TEXT], "fld_Average_Sales_Price", Math.Round(callingForm.SubjectNeighborhood.medianSalePrice).ToString() });
            commands.Add(0.013, new string[] { "1", inputType[elem.TEXT], "fld_sale_to_list_ratio", "93" });
            commands.Add(0.014, new string[] { "1", inputType[elem.TEXT], "fld_price_range_low", callingForm.SubjectNeighborhood.minListPrice.ToString() });
            commands.Add(0.015, new string[] { "1", inputType[elem.TEXT], "fld_price_range_high", callingForm.SubjectNeighborhood.maxListPrice.ToString() });

            //
            //Neighborhood comments
            //
            //fld_Neighborhood_Positive_Features
            commands.Add(0.016, new string[] { "1", inputType[elem.TEXTAREA], "fld_Neighborhood_Positive_Features", "*Positive neighborhood Comments needed to be placed here*" });
            //fld_Neighborhood_Negative_Features
            commands.Add(0.017, new string[] { "1", inputType[elem.TEXTAREA], "fld_Neighborhood_Negative_Features", "*Negitive neighborhood Comments needed to be placed here*" });
            //fld_Neighborhood_Location
            commands.Add(0.018, new string[] { "1", inputType[elem.TEXTAREA], "fld_Neighborhood_Location", "*Location impact on desireability Comments needed to be placed here*" });
            commands.Add(0.019, new string[] { "1", inputType[elem.TEXTAREA], "fld_Market_Condition", "*Current market condition Comments needed to be placed here*" });
            //fld_Neighborhood_Comments
            commands.Add(0.020, new string[] { "1", inputType[elem.TEXTAREA], "fld_Neighborhood_Comments", "*Any other neighborhood Comments needed to be placed here*" });

            //
            //Property information
            //
            //propertyTypeTranslation
            //fld_Property_Type
            commands.Add(0.021, new string[] { "1", inputType[elem.SELECT], "fld_Property_Type", propertyTypeTranslation[GlobalVar.theSubjectProperty.TypeOfMlsListing] });

            //fld_Property_View
            //TODO: translate view from bpo form
            //default to average
            commands.Add(0.022, new string[] { "1", inputType[elem.SELECT], "fld_Property_View", "%58" });
            //TODO: translate quality from bpo form
            //default to Q3
            //fld_Property_Quality
            commands.Add(0.023, new string[] { "1", inputType[elem.SELECT], "fld_Property_Quality", "%32" });
            //fld_Comparison_to_Neighborhood
            //TODO: translate value from bpo form
            //default to average
            commands.Add(0.024, new string[] { "1", inputType[elem.SELECT], "fld_Comparison_to_Neighborhood", "%58" });
            //fld_Condition
            //TODO: translate value from bpo form
            //default to C3
            commands.Add(0.025, new string[] { "1", inputType[elem.SELECT], "fld_Condition", "%17" });
            commands.Add(0.026, new string[] { "1", inputType[elem.SELECT], "fld_Design", homeStyleTranslation[form.SubjectMlsType] });
            //fld_Property_Construction
            //TODO: translate value from bpo form; field does not currently exsist on form
            //default to Frame
            commands.Add(0.027, new string[] { "1", inputType[elem.SELECT], "fld_Property_Construction", "%51" });
            //fld_Frontage
            //aka water frontage (unsupported default to 0)
            commands.Add(0.028, new string[] { "1", inputType[elem.TEXT], "fld_Frontage", "0" });
            //fld_Property_Size
            //aka building size aka above grade GLA
            commands.Add(0.029, new string[] { "1", inputType[elem.TEXT], "fld_Property_Size", form.SubjectAboveGLA });
            //fld_Age
            commands.Add(0.030, new string[] { "1", inputType[elem.TEXT], "fld_Age", GlobalVar.theSubjectProperty.Age.ToString() });
            //fld_Bedrooms
            commands.Add(0.031, new string[] { "1", inputType[elem.TEXT], "fld_Bedrooms", form.SubjectBedroomCount });
            commands.Add(0.032, new string[] { "1", inputType[elem.TEXT], "fld_Bathrooms", form.SubjectBathroomCount });

            //TODO: translate value from bpo form into a standard format using a new GarageType method on theSubjectProperty
            //default to Attached
            commands.Add(0.033, new string[] { "1", inputType[elem.SELECT], "fld_Garage_Type", "%37" });
            commands.Add(0.034, new string[] { "1", inputType[elem.TEXT], "fld_Garage_Cars", GlobalVar.theSubjectProperty.GarageStallCount });

            //TODO: rework bpo form field for basements into dropdown menus for correct user manual selections when mls sheet is missing for subject
            //default to Full
            commands.Add(0.035, new string[] { "1", inputType[elem.SELECT], "ddl_Basement_Id", "%62" });

            //TODO: Add the mls amenites field here.  Needs new field on bpo form
            commands.Add(0.036, new string[] { "1", inputType[elem.TEXT], "fld_Amenities", "*Copy mls amenities field here*" });


            //
            //Property Comments
            //
            commands.Add(0.037, new string[] { "1", inputType[elem.TEXTAREA], "fld_Property_Positive_Features", "*Positive property Comments needed to be placed here*" });
            commands.Add(0.038, new string[] { "1", inputType[elem.TEXTAREA], "fld_Property_Negative_Features", "*Negitive property Comments needed to be placed here*" });
            commands.Add(0.039, new string[] { "1", inputType[elem.TEXTAREA], "fld_Property_Repairs_Required", "*Repair Comments needed to be placed here*" });
            commands.Add(0.040, new string[] { "1", inputType[elem.TEXTAREA], "fld_Property_Comments", "*Any other property Comments needed to be placed here*" });





        
           //// commands.Add(0.003, new string[] { "1", inputType[elem.TEXT], "CurrentOwner",form.SubjectOOR });
           // //TaxAssessedValue
           // commands.Add(0.004, new string[] { "1", inputType[elem.TEXT], "TaxAssessedValue", form.SubjectAssessmentValue });

           // //RealEstateTaxes
           // commands.Add(0.007, new string[] { "1", inputType[elem.TEXT], "RealEstateTaxes", form.subjectTaxAmountTextBox.Text });
           // //TaxYear
           // commands.Add(0.008, new string[] { "1", inputType[elem.TEXT], "TaxYear", "2016" });
          
         
           // commands.Add(0.0101, new string[] { "1", inputType[elem.TEXT], "AssessorParcelNum", form.SubjectLandValue });



           // commands.Add(0.4, new string[] { "1", inputType[elem.TEXT], "UNITS", "1" });
           // // commands.Add(1, new string[] { "1", inputType[elem.TEXT], "Proximity", targetComp.proximityToSubject.ToString() });

           // commands.Add(4, new string[] { "1", inputType[elem.TEXT], "TotalRents", form.SubjectRent });
           // // commands.Add(5, new string[] { ListingDatePosition[targetCompString], inputType[elem.TEXT], "dt", targetComp.ListDateString });
           // //  commands.Add(6, new string[] { "1", inputType[elem.TEXT], "OrigListPrice", targetComp.OriginalListPrice.ToString() });
           // // commands.Add(7, new string[] { "1", inputType[elem.TEXT], "FinalListPrice", targetComp.CurrentListPrice.ToString() });
           // //  commands.Add(8, new string[] { "1", inputType[elem.TEXT], "SalesPrice", targetComp.SalePrice.ToString() });
           // // commands.Add(8.1, new string[] { "1", inputType[elem.SELECT], "FinanceTypeID", financingTypeTranslation[targetComp.FinancingMlsString] });
           // //  commands.Add(9, new string[] { SoldDatePosition[targetCompString], inputType[elem.TEXT], "dt", targetComp.SalesDate.ToShortDateString() });
           // //  commands.Add(10, new string[] { "1", inputType[elem.TEXT], "DOM", targetComp.DOM });

           // commands.Add(11, new string[] { "1", inputType[elem.TEXT], "YearBuilt", form.SubjectYearBuilt });
           // commands.Add(12, new string[] { "1", inputType[elem.TEXT], "LotSize", form.SubjectLotSize });
 
           // commands.Add(15, new string[] { "1", inputType[elem.TEXT], "TotalRooms", form.SubjectRoomCount });

           // commands.Add(17, new string[] { "1", inputType[elem.TEXT], "FullBaths", form.SubjectBathroomCount.Substring(0, 1) });
           // commands.Add(18, new string[] { "1", inputType[elem.TEXT], "HalfBaths", form.SubjectBathroomCount.Substring(2, 1) });

      
           // commands.Add(20, new string[] { "1", inputType[elem.TEXT], "BelowGradeSqft", form.SubjectBasementGLA });
           // commands.Add(21, new string[] { "1", inputType[elem.TEXT], "BasementPercentFinished", form.SubjectBasementFinishedGLA });
















         
            foreach (var c in commands)
            {
                macro.AppendFormat("TAG POS={0} TYPE={1} FORM=NAME:form1 ATTR=NAME:*{2} CONTENT={3}\r\n", c.Value[0], c.Value[1], c.Value[2], c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            }

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 30);


            ////
            ////Page 2
            ////
            //macro.Clear();
            //commands.Clear(); 
            ////macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:bpoForm ATTR=TXT:Page1");
            //macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:bpoForm ATTR=TXT:Page2");

            ////
            ////Top
            ////

            
            ////OwnershipPrideID; 431:Good; 432:Average; 433:Poor; 434:Excellent; 435:Fair
            //commands.Add(2.002, new string[] { "1", inputType[elem.SELECT], "OwnershipPrideID", "%number:432" });
            ////NeighborhoodID; 431:Good; 432:Average; 433:Poor; 434:Excellent; 435:Fair
            //commands.Add(2.003, new string[] { "1", inputType[elem.SELECT], "NeighborhoodID", "%number:432" });
            ////VandalismRiskID; 523:Low; 524:Medium; 525:High
            //commands.Add(2.004, new string[] { "1", inputType[elem.SELECT], "VandalismRiskID", "%number:523" });
            ////UnemploymentRateID; 537:Decreasing; 538:Increasing; 539:Stable
            //commands.Add(2.005, new string[] { "1", inputType[elem.SELECT], "UnemploymentRateID", "%number:539" });
            ////ForeclosureRateID; 537:Decreasing; 538:Increasing; 539:Stable
            //commands.Add(2.006, new string[] { "1", inputType[elem.SELECT], "ForeclosureRateID", "%number:539" });
            ////PredominantSellerID; 526:Reo; 527:Investors; 528:Owner Occupants
            //commands.Add(2.007, new string[] { "1", inputType[elem.SELECT], "PredominantSellerID", "%number:528" });
            ////PredominantBuyerID; 529:First Time; 530:Move-up; 531:Investors
            //commands.Add(2.008, new string[] { "1", inputType[elem.SELECT], "PredominantBuyerID", "%number:530" });
            ////PredominantFinancingID; 532:Cash; 533:Conventional; 534:FHA; 535:VA; 536:Other
            //commands.Add(2.009, new string[] { "1", inputType[elem.SELECT], "PredominantFinancingID", "%number:533" });
            ////OwnerOccupiedPercent
            //commands.Add(2.010, new string[] { "1", inputType[elem.TEXT], "OwnerOccupiedPercent", "90" });
            ////TenantOccupiedPercent
            //commands.Add(2.011, new string[] { "1", inputType[elem.TEXT], "TenantOccupiedPercent", "10" });


            //commands.Add(2.014, new string[] { "1", inputType[elem.SELECT], "PropertyValueID", "%number:539" });
            
            ////
            ////Mid
            ////

            ////NumberOfListing; No. of REO Listings 
            //commands.Add(2.015, new string[] { "1", inputType[elem.TEXT], "NumberOfListing", callingForm.SubjectNeighborhood.numberREOListings.ToString() });
            ////AvgPrice
            ////   commands.Add(2.016, new string[] { "1", inputType[elem.TEXT], "AvgPrice", callingForm.SubjectNeighborhood.medianListPrice.ToString() });
            ////LowPrice
            ////   commands.Add(2.017, new string[] { "1", inputType[elem.TEXT], "LowPrice", callingForm.SubjectNeighborhood.medianListPrice.ToString() });
            ////HighPrice
            ////   commands.Add(2.018, new string[] { "1", inputType[elem.TEXT], "HighPrice", callingForm.SubjectNeighborhood.medianListPrice.ToString() });
            ////NumOfREOSales - 90 days
            //commands.Add(2.019, new string[] { "1", inputType[elem.TEXT], "NumOfREOSales", (callingForm.SubjectNeighborhood.numberREOSales / 4).ToString() });

           
            ////Retail_AvgPrice
            //commands.Add(2.021, new string[] { "1", inputType[elem.TEXT], "Retail_AvgPrice", callingForm.SubjectNeighborhood.medianListPrice.ToString() });
          
            ////NumOfRetailSales
            //commands.Add(2.024, new string[] { "1", inputType[elem.TEXT], "NumOfRetailSales", (callingForm.SubjectNeighborhood.numberSoldListings/4).ToString() });

            ////NumberInDirectCompetition
            //commands.Add(2.025, new string[] { "1", inputType[elem.TEXT], "NumberInDirectCompetition", callingForm.SubjectNeighborhood.numberOfCompListings.ToString() });
            ////DC_AvgPrice
            //commands.Add(2.026, new string[] { "1", inputType[elem.TEXT], "DC_AvgPrice", callingForm.SetOfComps.medianListPrice.ToString() });
            ////DC_LowPrice
            //commands.Add(2.027, new string[] { "1", inputType[elem.TEXT], "DC_LowPrice", callingForm.SetOfComps.minListPrice.ToString() });
            ////DC_HighPrice
            //commands.Add(2.028, new string[] { "1", inputType[elem.TEXT], "DC_HighPrice", callingForm.SetOfComps.maxListPrice.ToString() });
            ////NumOfRetailSales
            //commands.Add(2.029, new string[] { "1", inputType[elem.TEXT], "NumberOfDirectSales", (callingForm.SetOfComps.numberSoldListings / 4).ToString() });

            ////FairMarketRents
            //commands.Add(2.030, new string[] { "1", inputType[elem.TEXT], "FairMarketRents", callingForm.SubjectRent });


          
            //foreach (var c in commands)
            //{
            //    if (c.Value[2] == "dt")
            //    {
            //        tcs = "";
            //    }
            //    else
            //    {
            //        tcs = targetCompString;
            //    }
            //    macro.AppendFormat("TAG POS={0} TYPE={1} ATTR=ng-model:*{3}*{2} CONTENT={4}\r\n", c.Value[0], c.Value[1], c.Value[2], tcs, c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            //}

            // macroCode = macro.ToString();
            //iim.iimPlayCode(macroCode, 30);



            //commands.Add(6, new string[] { "1", inputType[elem.TEXT], "PROP_FOR_SALE_L", callingForm.SubjectNeighborhood.numberOfCompListings.ToString()});
            //commands.Add(8, new string[] { "1", inputType[elem.TEXT], "LIST_LOW_S", callingForm.SubjectNeighborhood.minListPrice.ToString() });
            //commands.Add(9, new string[] { "1", inputType[elem.TEXT], "LIST_HIGH_S", callingForm.SubjectNeighborhood.maxListPrice.ToString() });
            //commands.Add(10, new string[] { "1", inputType[elem.TEXT], "SALE_COUNT", callingForm.SubjectNeighborhood.numberSoldListings.ToString() });
            //commands.Add(11, new string[] { "1", inputType[elem.TEXT], "SALE_LOW_S", callingForm.SubjectNeighborhood.minSalePrice.ToString() });
            //commands.Add(12, new string[] { "1", inputType[elem.TEXT], "SALE_HIGH_S", callingForm.SubjectNeighborhood.maxSalePrice.ToString() });
            //commands.Add(13, new string[] { "1", inputType[elem.TEXT], "BOARDED_UP_HOMES_NBR", "0" });
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
            

        }




        public void CompFill(iMacros.App iim, string saleOrList, string compString, Dictionary<string, string> fieldList)
        {
            StringBuilder macro = new StringBuilder();
            targetCompNumber = Regex.Match(compString, @"\d").Value;
            targetCompString = compString;


            Dictionary<string, string> financingTypeTranslation = new Dictionary<string, string>()
            {
               {"Cash", "%number:392"}, {"Conventional", "%number:393"}, {"FHA", "%number:394"}, {"VA", "%number:395"},  {"Contract", "%number:396"}, {"Unknown", "%number:396"}
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
            Dictionary<string, string> compNumpTranslation = new Dictionary<string, string>()
            {
                {"Sold1", "2"},  {"Sold2", "3"}, {"Sold3", "4"},  {"List1", "5"},  {"List2", "6"},  {"List3", "7"}
            };


            Dictionary<double, string[]> commands = new Dictionary<double, string[]>();

            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 1");

            if (saleOrList == "sale")
            {
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:form1 ATTR=TXT:Comparable<SP>Sales");
            }
            else
            {
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:form1 ATTR=TXT:Comparable<SP>Listings");
            }


            //id=ddl_Data_Source_Type_1
            //commands.Add(1, new string[] { "1", inputType[elem.SELECT], "ddl_Data_Source_Type_1", "%MLS" });
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=ID:ddl_Data_Source_Type_" + targetCompNumber + " CONTENT=%MLS");

            //fld_MLS
            commands.Add(2, new string[] { "1", inputType[elem.TEXT], "fld_MLS", targetComp.MlsNumber });

            //fld_Address
            commands.Add(3, new string[] { "1", inputType[elem.TEXT], "fld_Address", targetComp.StreetAddress });

            //fld_City
            commands.Add(4, new string[] { "1", inputType[elem.TEXT], "fld_City", targetComp.City.Replace(" ", "<SP>") });

            //fld_State
            commands.Add(5, new string[] { "1", inputType[elem.TEXT], "fld_State", "IL" });

            //fld_Zip
            commands.Add(6, new string[] { "1", inputType[elem.TEXT], "fld_Zip", targetComp.Zipcode });

            //fld_Proximity
            commands.Add(7, new string[] { "1", inputType[elem.TEXT], "fld_Proximity", targetComp.proximityToSubject.ToString() });

            //fld_Sale_Date
            //fld_Sale_Type
            //fld_Sale_Price
            if (saleOrList == "sale")
            {
                commands.Add(0.001, new string[] { "1", inputType[elem.TEXT], "fld_Sale_Date", targetComp.SalesDate.ToShortDateString() });
                commands.Add(0.002, new string[] { "1", inputType[elem.SELECT], "fld_Sale_Type", saleTypeTranslation[targetComp.TransactionType] });
                commands.Add(0.003, new string[] { "1", inputType[elem.TEXT], "fld_Sale_Price", targetComp.SalePrice.ToString() });
            }
            else
            {
                //fld_Sale_Price
                commands.Add(0.004, new string[] { "1", inputType[elem.TEXT], "fld_Sale_Price", targetComp.CurrentListPrice.ToString() });

            }

            //fld_List_Price
            commands.Add(8, new string[] { "1", inputType[elem.TEXT], "fld_List_Price", targetComp.OriginalListPrice.ToString() });

            //fld_Days_on_Market
            commands.Add(9, new string[] { "1", inputType[elem.TEXT], "fld_Days_on_Market", targetComp.DOM });

            //fld_Location
            //use whatever subject is set to on bpo form
            commands.Add(10, new string[] { "1", inputType[elem.SELECT], "fld_Location", locationTypeTranslation[GlobalVar.mainWindow.comboBoxLocationDescr.Text] });

            //fld_View
            //TODO: use whatever subject is set to.
            //default to average
            commands.Add(11, new string[] { "1", inputType[elem.SELECT], "fld_View", "%58" });

            //fld_Lot_Size
            commands.Add(12, new string[] { "1", inputType[elem.TEXT], "fld_Lot_Size", targetComp.Lotsize.ToString() });

            //fld_Property_Design
            commands.Add(13, new string[] { "1", inputType[elem.SELECT], "fld_Property_Design", homeStyleTranslation[targetComp.Type] });

            //fld_Quality
            //default to Q3
            commands.Add(14, new string[] { "1", inputType[elem.SELECT], "fld_Quality", "%32" });

            //fld_Construction
            //default to frame
            commands.Add(15, new string[] { "1", inputType[elem.SELECT], "fld_Construction", "%51" });

            //fld_Condition
            //default to C3
            commands.Add(16, new string[] { "1", inputType[elem.SELECT], "fld_Condition", "%17" });

            //fld_Age
            commands.Add(17, new string[] { "1", inputType[elem.TEXT], "fld_Age", targetComp.Age.ToString() });

            //fld_Bedrooms
            commands.Add(18, new string[] { "1", inputType[elem.TEXT], "fld_Bedrooms", targetComp.BedroomCount });

            //fld_Bathrooms
            commands.Add(19, new string[] { "1", inputType[elem.TEXT], "fld_Bathrooms", targetComp.BathroomCount });

            //fld_Garage_Type
            commands.Add(20, new string[] { "1", inputType[elem.SELECT], "fld_Garage_Type", garageTranslation[targetComp.GarageType()] });

            //fld_Garage_Spaces
            commands.Add(21, new string[] { "1", inputType[elem.TEXT], "fld_Garage_Spaces", targetComp.NumberGarageStalls() });

            //ddl_Basement_Id
            commands.Add(22, new string[] { "1", inputType[elem.SELECT], "ddl_Basement_Id", basementStyleTranslation[targetComp.BasementType] });

            //fld_Basement_Pct
            commands.Add(23, new string[] { "1", inputType[elem.TEXT], "fld_Basement_Pct", targetComp.BasementFinishedPercentage() });

            //fld_Amenities
            //TODO: Create a standard string to include features.
            commands.Add(24, new string[] { "1", inputType[elem.TEXT], "fld_Amenities", "*add features here*" });

            //fld_Living_Area
            commands.Add(25, new string[] { "1", inputType[elem.TEXT], "fld_Living_Area", targetComp.ProperGla(GlobalVar.mainWindow.SubjectAboveGLA) });




            foreach (var c in commands)
            {
                macro.AppendFormat("TAG POS={0} TYPE={1} FORM=NAME:form1 ATTR=NAME:*{2}_{4} CONTENT={3}\r\n", c.Value[0], c.Value[1], c.Value[2], c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""), targetCompNumber);
            }
            

            
            //commands.Add(0.4, new string[] { "1", inputType[elem.TEXT], "UNITS", "1" });
          
            //commands.Add(2, new string[] { "1", inputType[elem.SELECT], "PropertyTypeID", "%number:397" });
            //commands.Add(3, new string[] { "1", inputType[elem.SELECT], "HomeStyleID", "%number:426" });
            //commands.Add(4, new string[] { "1", inputType[elem.TEXT], "TotalRents", "1000" });


            //commands.Add(8.1, new string[] { "1", inputType[elem.SELECT], "FinanceTypeID", financingTypeTranslation[targetComp.FinancingMlsString] });
            //commands.Add(11, new string[] { "1", inputType[elem.TEXT], "YearBuilt", targetComp.YearBuiltString });

            //commands.Add(15, new string[] { "1", inputType[elem.TEXT], "TotalRooms", targetComp.TotalRoomCount.ToString() });
            //commands.Add(17, new string[] { "1", inputType[elem.TEXT], "FullBaths", targetComp.FullBathCount });
            //commands.Add(18, new string[] { "1", inputType[elem.TEXT], "HalfBaths", targetComp.HalfBathCount });
            //commands.Add(19, new string[] { "1", inputType[elem.TEXT], "AboveGradeSqft", targetComp.ProperGla(GlobalVar.mainWindow.SubjectAboveGLA) });
            //commands.Add(20, new string[] { "1", inputType[elem.TEXT], "BelowGradeSqft", targetComp.BasementGLA() });
            //commands.Add(22, new string[] { "1", inputType[elem.TEXT], "ViewDescript", "Residential" });
      

            ////TODO: add hasCarport support and hasExtParking support
            ////GarageTypeID
            //if (targetComp.AttachedGarage())
            //{
            //    commands.Add(24, new string[] { "1", inputType[elem.SELECT], "GarageTypeID", garageTranslation[targetComp.NumberGarageStalls() + " Attached"] });
            //}
            //else if (targetComp.DetachedGarage())
            //{
            //    commands.Add(24, new string[] { "1", inputType[elem.SELECT], "GarageTypeID", garageTranslation[targetComp.NumberGarageStalls() + " Detached"] });
            //}
            //else
            //{
            //    commands.Add(24, new string[] { "1", inputType[elem.SELECT], "GarageTypeID", "%number:1976" });

            //}
            ////TODO:  add hasPool support
            ////PoolFireID
            //if (String.IsNullOrWhiteSpace(targetComp.NumberOfFireplaces) ||   targetComp.NumberOfFireplaces == "0")
            //{
            //    commands.Add(25, new string[] { "1", inputType[elem.SELECT], "PoolFireID", "%number:453" });
            //}
            //else
            //{
            //    commands.Add(25, new string[] { "1", inputType[elem.SELECT], "PoolFireID", "%number:451" });
            //}

            ////SaleType  saleTypeTranslation
            //commands.Add(26, new string[] { "1", inputType[elem.SELECT], "SaleType",  saleTypeTranslation[targetComp.TransactionType]});

            ////Comment
            //commands.Add(27, new string[] { "1", inputType[elem.TEXTAREA], "Comment", targetComp.mlsHtmlFields["remarks"].value });





        
            //string tcs = targetCompString;
            //foreach (var c in commands)
            //{
            //    if (c.Value[2] == "dt")
            //    {
            //        tcs = "";
            //    }
            //    else
            //    {
            //        tcs = targetCompString;
            //    }
            //    macro.AppendFormat("TAG POS={0} TYPE={1} ATTR=ng-model:*{3}*{2} CONTENT={4}\r\n", c.Value[0], c.Value[1], c.Value[2], tcs, c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            //}

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 60);

            //
            //MMR
            //


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
            //commands.Add(24, new string[] { "SELECT", "ddlLocation", "%Average" });
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
    }
}
