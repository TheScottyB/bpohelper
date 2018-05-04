using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{

    class Pyramid : BPOFulfillment
    {

        public Pyramid()
        {
            subject = GlobalVar.theSubjectProperty;
            callingForm = GlobalVar.mainWindow;
        }

        public Pyramid(MLSListing m)
            : this()
        {
            targetComp = m;
        }

        enum elem { TEXT, SELECT, TEXTAREA, RADIO, CHECKBOX };

        private Dictionary<elem, string> inputType = new Dictionary<elem, string>()
        {
            {elem.TEXT, "INPUT:TEXT"},{elem.SELECT, "SELECT"}, {elem.TEXTAREA, "TEXTAREA"}, {elem.RADIO, "INPUT:RADIO"}, {elem.CHECKBOX, "INPUT:CHECKBOX"}
        };

        Dictionary<string, string> homeStyleTranslation = new Dictionary<string, string>()
            {
                {"1 Story", "%number:409"}, {"2 Stories", "%number:426"}, {"Raised Ranch", "%number:420"}
            };

        Dictionary<string, string> garageTranslation = new Dictionary<string, string>()
            {
                {"1 Attached", "%number:1964"}, {"2 Attached", "%number:1965"}, {"3 Attached", "%number:1966"}, {"4 Attached", "%number:1967"}, 
                    {"1 Detached", "%number:1968"},{"2 Detached", "%number:1969"},{"3 Detached", "%number:1970"}
            };

        Dictionary<string, string> saleTypeTranslation = new Dictionary<string, string>()
            {
                {"Arms Length", "%number:2026"}, {"REO", "%number:2027"}, {"ShortSale", "%number:2028"},  {"Contract", "%number:2028"}
            };

        Dictionary<string, string> occupancyTypeTranslation = new Dictionary<string, string>()
            {
                {"Occupied", "%number:436"}, {"Occupied by Owner", "%number:436"}, {"Occupied by Tenant", "%number:436"},
                    {"Vacant", "%number:437"}, {"Other", "%number:439"}, {"Unknown", "%number:439"}
            };
        Dictionary<string, string> typeOfOccupancyTranslation = new Dictionary<string, string>()
            {
                {"Occupied", "%number:312"}, {"Occupied by Owner", "%number:309"}, {"Occupied by Tenant", "%number:310"},
                    {"Vacant", "%number:312"}, {"Other", "%number:312"}, {"Unknown", "%number:312"}
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
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");

            //
            //Page 1 - Top Section
            //
            commands.Add(0.001, new string[] { "1", inputType[elem.TEXT], "PreviousListPrice", "0" });
            commands.Add(0.002, new string[] { "1", inputType[elem.TEXT], "PreviousSalePrice", "0" });
            //CurrentOwner
            commands.Add(0.003, new string[] { "1", inputType[elem.TEXT], "CurrentOwner",form.SubjectOOR });
            //TaxAssessedValue
            commands.Add(0.004, new string[] { "1", inputType[elem.TEXT], "TaxAssessedValue", form.SubjectAssessmentValue });
            //OccupancyTypeID 436 Occupied, 437 vacant, 439 unknown
            commands.Add(0.005, new string[] { "1", inputType[elem.SELECT], "OccupancyTypeID", occupancyTypeTranslation[form.comboBoxSubjectOccupancy.Text] });
            //OccType 309:Mortagor, 310:Renter, 311:Squatter, 312:Unknown, 3829:Personal Property
            commands.Add(0.006, new string[] { "1", inputType[elem.SELECT], "OccType", typeOfOccupancyTranslation[form.comboBoxSubjectOccupancy.Text] });
            //RealEstateTaxes
            commands.Add(0.007, new string[] { "1", inputType[elem.TEXT], "RealEstateTaxes", form.subjectTaxAmountTextBox.Text });
            //TaxYear
            commands.Add(0.008, new string[] { "1", inputType[elem.TEXT], "TaxYear", "2016" });
            //AssessorParcelNum
            commands.Add(0.009, new string[] { "1", inputType[elem.TEXT], "AssessorParcelNum", form.SubjectPin });
            //DelinquentTaxAmt
            //CodeViolationAmt
            //SpecialAssessment
            //SubjectLotValue
            commands.Add(0.0101, new string[] { "1", inputType[elem.TEXT], "AssessorParcelNum", form.SubjectLandValue });

            //<option value="number:398" label="Condo">Condo</option>
            //<option value="number:397" label="Single Family">Single Family</option>
            commands.Add(0, new string[] { "1", inputType[elem.SELECT], "PropertyTypeID", "%number:397" });

            //<option value="number:426" label="2 Story Conventional">2 Story Conventional</option>
            //<option value="number:423" label="2 Story Modern">2 Story Modern</option>
            //<option value="number:414" label="Condo">Condo</option>
            //<option value="number:418" label="Multi Level">Multi Level</option>
            //<option value="number:418" label="Multi Level">Multi Level</option>
            //<option value="number:420" label="Split Entry">Split Entry</option>
            //<option value="number:409" label="Ranch/Rambler">Ranch/Rambler</option>
            commands.Add(1, new string[] { "1", inputType[elem.SELECT], "HomeStyleID", homeStyleTranslation[form.SubjectMlsType] });


            commands.Add(0.4, new string[] { "1", inputType[elem.TEXT], "UNITS", "1" });
            // commands.Add(1, new string[] { "1", inputType[elem.TEXT], "Proximity", targetComp.proximityToSubject.ToString() });
            //   commands.Add(2, new string[] { "1", inputType[elem.SELECT], "PropertyTypeID", "%number:397" });
            //  commands.Add(3, new string[] { "1", inputType[elem.SELECT], "HomeStyleID", "%number:426" });
            commands.Add(4, new string[] { "1", inputType[elem.TEXT], "TotalRents", form.SubjectRent });
            // commands.Add(5, new string[] { ListingDatePosition[targetCompString], inputType[elem.TEXT], "dt", targetComp.ListDateString });
            //  commands.Add(6, new string[] { "1", inputType[elem.TEXT], "OrigListPrice", targetComp.OriginalListPrice.ToString() });
            // commands.Add(7, new string[] { "1", inputType[elem.TEXT], "FinalListPrice", targetComp.CurrentListPrice.ToString() });
            //  commands.Add(8, new string[] { "1", inputType[elem.TEXT], "SalesPrice", targetComp.SalePrice.ToString() });
            // commands.Add(8.1, new string[] { "1", inputType[elem.SELECT], "FinanceTypeID", financingTypeTranslation[targetComp.FinancingMlsString] });
            //  commands.Add(9, new string[] { SoldDatePosition[targetCompString], inputType[elem.TEXT], "dt", targetComp.SalesDate.ToShortDateString() });
            //  commands.Add(10, new string[] { "1", inputType[elem.TEXT], "DOM", targetComp.DOM });

            commands.Add(11, new string[] { "1", inputType[elem.TEXT], "YearBuilt", form.SubjectYearBuilt });
            commands.Add(12, new string[] { "1", inputType[elem.TEXT], "LotSize", form.SubjectLotSize });
            commands.Add(13, new string[] { "1", inputType[elem.SELECT], "ConditionTypeID", "%number:432" });
            commands.Add(14, new string[] { "1", inputType[elem.SELECT], "ConstructionTypeID", "%number:430" });
            commands.Add(15, new string[] { "1", inputType[elem.TEXT], "TotalRooms", form.SubjectRoomCount });
            commands.Add(16, new string[] { "1", inputType[elem.TEXT], "BedRooms", form.SubjectBedroomCount });
            commands.Add(17, new string[] { "1", inputType[elem.TEXT], "FullBaths", form.SubjectBathroomCount.Substring(0, 1) });
            commands.Add(18, new string[] { "1", inputType[elem.TEXT], "HalfBaths", form.SubjectBathroomCount.Substring(2, 1) });

            commands.Add(19, new string[] { "1", inputType[elem.TEXT], "AboveGradeSqft", form.SubjectAboveGLA });
            commands.Add(20, new string[] { "1", inputType[elem.TEXT], "BelowGradeSqft", form.SubjectBasementGLA });
            commands.Add(21, new string[] { "1", inputType[elem.TEXT], "BasementPercentFinished", form.SubjectBasementFinishedGLA });
            commands.Add(22, new string[] { "1", inputType[elem.TEXT], "ViewDescript", "Residential" });


            ///
            //04-28-2018 Updates
            ///

            //AsIs90Day
            //AsIsSuggestedListPrice
            //Repaired90Day
            //RepairedSuggestedListPrice
            //AsIsValQuickSale
            //QuickSaleSuggestedListPrice
            commands.Add(3.001, new string[] { "1", inputType[elem.TEXT], "AsIs90Day", form.SubjectMarketValue });
            commands.Add(3.002, new string[] { "1", inputType[elem.TEXT], "AsIsSuggestedListPrice", form.SubjectMarketValueList });
            commands.Add(3.003, new string[] { "1", inputType[elem.TEXT], "Repaired90Day", form.SubjectMarketValue });
            commands.Add(3.004, new string[] { "1", inputType[elem.TEXT], "RepairedSuggestedListPrice", form.SubjectMarketValueList });
            commands.Add(3.005, new string[] { "1", inputType[elem.TEXT], "AsIsValQuickSale", form.SubjectQuickSaleValue });
            commands.Add(3.006, new string[] { "1", inputType[elem.TEXT], "QuickSaleSuggestedListPrice", form.SubjectQuickSaleValue });

            //SubjectLotValue
            commands.Add(3.007, new string[] { "1", inputType[elem.TEXT], "SubjectLotValue", form.SubjectLandValue });

            //WaterType
            //SewerType
            //Defaluted to unknown
            //TODO: Add water/sewer fields to bpohelper form and set these fields on web form
            commands.Add(3.008, new string[] { "1", inputType[elem.SELECT], "WaterType", "%number:2210" });
            commands.Add(3.009, new string[] { "1", inputType[elem.SELECT], "SewerType", "%number:2207" });

            //OtherComment
            //Default comment
            //TODO: Save bpo comments in subjectinfo text and load between bpos.
            commands.Add(3.010, new string[] { "1", inputType[elem.TEXTAREA], "OtherComment", "The subject is a conforming home within the neighborhood. No adverse conditions were noted at the time of inspection based on exterior observations. Unable to determine interior condition due to exterior inspection only, so subject was assumed to be in average condition for this report." });



            //GarageTypeID
            if (GlobalVar.theSubjectProperty.AttachedGarage)
            {
                commands.Add(3.011, new string[] { "1", inputType[elem.SELECT], "GarageTypeID", garageTranslation[GlobalVar.theSubjectProperty.GarageStallCount + " Attached"]});
            }
            else if (GlobalVar.theSubjectProperty.DetachedGarage)
            {
                commands.Add(3.011, new string[] { "1", inputType[elem.SELECT], "GarageTypeID", garageTranslation[GlobalVar.theSubjectProperty.GarageStallCount + " Detached"] });
 
            }
            else
            {
                commands.Add(3.011, new string[] { "1", inputType[elem.SELECT], "GarageTypeID", "%number:1976" }); 
            }

            //PoolFireID
            //TODO:  Add pool fields to bpo helper
            if (String.IsNullOrWhiteSpace(form.SubjectNumberFireplaces))
            {
                commands.Add(3.012, new string[] { "1", inputType[elem.SELECT], "PoolFireID", "%number:451" }); 
            }
            else 
            {
                commands.Add(3.012, new string[] { "1", inputType[elem.SELECT], "PoolFireID", "%number:453" }); 
            }

            //SaleType
            //default to other, not sure what they what in this field
            commands.Add(3.013, new string[] { "1", inputType[elem.SELECT], "SaleType", "%number:2029" });

            //Analysis
            //Default comment
            //TODO: Save bpo comments in subjectinfo text and load between bpos.
            commands.Add(3.014, new string[] { "1", inputType[elem.TEXTAREA], "Analysis", "All the comps are reasonable substitute for the subject property, similar in most areas. Price opinion was based off comparable statistics." });


            ///
            //end
            ///










            string tcs = "BPOSubjCompItem";
            foreach (var c in commands)
            {
                if (c.Value[2] == "dt")
                {
                    tcs = "";
                }
                else
                {
                    tcs = targetCompString;
                }
                macro.AppendFormat("TAG POS={0} TYPE={1} ATTR=ng-model:*{3}*{2} CONTENT={4}\r\n", c.Value[0], c.Value[1], c.Value[2], tcs, c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            }

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 30);


            //
            //Page 2
            //
            macro.Clear();
            commands.Clear(); 
            //macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:bpoForm ATTR=TXT:Page1");
            macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:bpoForm ATTR=TXT:Page2");

            //
            //Top
            //

            //MarketTypeID; 243:Urban; 244:Surburban; 245:Rural; 246:Seasonal
            commands.Add(2.001, new string[] { "1", inputType[elem.SELECT], "MarketTypeID", "%number:244" });
            //OwnershipPrideID; 431:Good; 432:Average; 433:Poor; 434:Excellent; 435:Fair
            commands.Add(2.002, new string[] { "1", inputType[elem.SELECT], "OwnershipPrideID", "%number:432" });
            //NeighborhoodID; 431:Good; 432:Average; 433:Poor; 434:Excellent; 435:Fair
            commands.Add(2.003, new string[] { "1", inputType[elem.SELECT], "NeighborhoodID", "%number:432" });
            //VandalismRiskID; 523:Low; 524:Medium; 525:High
            commands.Add(2.004, new string[] { "1", inputType[elem.SELECT], "VandalismRiskID", "%number:523" });
            //UnemploymentRateID; 537:Decreasing; 538:Increasing; 539:Stable
            commands.Add(2.005, new string[] { "1", inputType[elem.SELECT], "UnemploymentRateID", "%number:539" });
            //ForeclosureRateID; 537:Decreasing; 538:Increasing; 539:Stable
            commands.Add(2.006, new string[] { "1", inputType[elem.SELECT], "ForeclosureRateID", "%number:539" });
            //PredominantSellerID; 526:Reo; 527:Investors; 528:Owner Occupants
            commands.Add(2.007, new string[] { "1", inputType[elem.SELECT], "PredominantSellerID", "%number:528" });
            //PredominantBuyerID; 529:First Time; 530:Move-up; 531:Investors
            commands.Add(2.008, new string[] { "1", inputType[elem.SELECT], "PredominantBuyerID", "%number:530" });
            //PredominantFinancingID; 532:Cash; 533:Conventional; 534:FHA; 535:VA; 536:Other
            commands.Add(2.009, new string[] { "1", inputType[elem.SELECT], "PredominantFinancingID", "%number:533" });
            //OwnerOccupiedPercent
            commands.Add(2.010, new string[] { "1", inputType[elem.TEXT], "OwnerOccupiedPercent", "90" });
            //TenantOccupiedPercent
            commands.Add(2.011, new string[] { "1", inputType[elem.TEXT], "TenantOccupiedPercent", "10" });
            //InventoryID; 537:Decreasing; 538:Increasing; 539:Stable
            commands.Add(2.012, new string[] { "1", inputType[elem.SELECT], "InventoryID", "%number:539" });
            //DemandID; 537:Decreasing; 538:Increasing; 539:Stable
            commands.Add(2.013, new string[] { "1", inputType[elem.SELECT], "DemandID", "%number:539" });
            //PropertyValueID; 537:Decreasing; 538:Increasing; 539:Stable
            commands.Add(2.014, new string[] { "1", inputType[elem.SELECT], "PropertyValueID", "%number:539" });
            
            //
            //Mid
            //

            //NumberOfListing; No. of REO Listings 
            commands.Add(2.015, new string[] { "1", inputType[elem.TEXT], "NumberOfListing", callingForm.SubjectNeighborhood.numberREOListings.ToString() });
            //AvgPrice
            //   commands.Add(2.016, new string[] { "1", inputType[elem.TEXT], "AvgPrice", callingForm.SubjectNeighborhood.medianListPrice.ToString() });
            //LowPrice
            //   commands.Add(2.017, new string[] { "1", inputType[elem.TEXT], "LowPrice", callingForm.SubjectNeighborhood.medianListPrice.ToString() });
            //HighPrice
            //   commands.Add(2.018, new string[] { "1", inputType[elem.TEXT], "HighPrice", callingForm.SubjectNeighborhood.medianListPrice.ToString() });
            //NumOfREOSales - 90 days
            commands.Add(2.019, new string[] { "1", inputType[elem.TEXT], "NumOfREOSales", (callingForm.SubjectNeighborhood.numberREOSales / 4).ToString() });

            //NumberOfRetail
            commands.Add(2.020, new string[] { "1", inputType[elem.TEXT], "NumberOfRetail", callingForm.SubjectNeighborhood.numberActiveListings.ToString() });
            //Retail_AvgPrice
            commands.Add(2.021, new string[] { "1", inputType[elem.TEXT], "Retail_AvgPrice", callingForm.SubjectNeighborhood.medianListPrice.ToString() });
            //Retail_LowPrice
            commands.Add(2.022, new string[] { "1", inputType[elem.TEXT], "Retail_LowPrice", callingForm.SubjectNeighborhood.minListPrice.ToString() });
            //Retail_HighPrice
            commands.Add(2.023, new string[] { "1", inputType[elem.TEXT], "Retail_HighPrice", callingForm.SubjectNeighborhood.maxListPrice.ToString() });
            //NumOfRetailSales
            commands.Add(2.024, new string[] { "1", inputType[elem.TEXT], "NumOfRetailSales", (callingForm.SubjectNeighborhood.numberSoldListings/4).ToString() });

            //NumberInDirectCompetition
            commands.Add(2.025, new string[] { "1", inputType[elem.TEXT], "NumberInDirectCompetition", callingForm.SetOfComps.numberActiveListings.ToString()});
            //DC_AvgPrice
            commands.Add(2.026, new string[] { "1", inputType[elem.TEXT], "DC_AvgPrice", callingForm.SetOfComps.medianListPrice.ToString() });
            //DC_LowPrice
            commands.Add(2.027, new string[] { "1", inputType[elem.TEXT], "DC_LowPrice", callingForm.SetOfComps.minListPrice.ToString() });
            //DC_HighPrice
            commands.Add(2.028, new string[] { "1", inputType[elem.TEXT], "DC_HighPrice", callingForm.SetOfComps.maxListPrice.ToString() });
            //NumOfRetailSales
            commands.Add(2.029, new string[] { "1", inputType[elem.TEXT], "NumberOfDirectSales", (callingForm.SetOfComps.numberSoldListings / 4).ToString() });

            //FairMarketRents
            commands.Add(2.030, new string[] { "1", inputType[elem.TEXT], "FairMarketRents", callingForm.SubjectRent });

            ///
            //04/28/18 updates
            ///

            //AvgDom
            //AvgListingTime
            //AvgClosingTime
            commands.Add(2.0001, new string[] { "1", inputType[elem.TEXT], "AvgDom", callingForm.SubjectNeighborhood.avgDom.ToString() });
            commands.Add(2.0002, new string[] { "1", inputType[elem.TEXT], "AvgListingTime", (callingForm.SubjectNeighborhood.avgDom + 45).ToString() });
            commands.Add(2.0003, new string[] { "1", inputType[elem.TEXT], "AvgClosingTime", "45" });

            //Conformity_To_Neighborhood
            //Positive_Features
            //Negative_Features
            //Recommended_Inspections
            //Interior_Odors
            //Marketing_Action_Plan
            //Anticipated_Resale_Problems
            //Default comment
            //TODO: Save bpo comments in subjectinfo text and load between bpos.
            commands.Add(2.0004, new string[] { "1", inputType[elem.TEXTAREA], "Conformity_To_Neighborhood", "Good" });
            commands.Add(2.0005, new string[] { "1", inputType[elem.TEXTAREA], "Positive_Features", "No change." });
            commands.Add(2.0006, new string[] { "1", inputType[elem.TEXTAREA], "Negative_Features", "No change." });
            commands.Add(2.0007, new string[] { "1", inputType[elem.TEXTAREA], "Recommended_Inspections", "No change." });
            commands.Add(2.0008, new string[] { "1", inputType[elem.TEXTAREA], "Interior_Odors", "No change." });
            commands.Add(2.0009, new string[] { "1", inputType[elem.TEXTAREA], "Anticipated_Resale_Problems", "No change." });
            commands.Add(2.0011, new string[] { "1", inputType[elem.TEXTAREA], "Marketing_Action_Plan", "No change." });




         
            //
            //
            //
          
            foreach (var c in commands)
            {
                if (c.Value[2] == "dt")
                {
                    tcs = "";
                }
                else
                {
                    tcs = targetCompString;
                }
                macro.AppendFormat("TAG POS={0} TYPE={1} ATTR=ng-model:*{3}*{2} CONTENT={4}\r\n", c.Value[0], c.Value[1], c.Value[2], tcs, c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            }

            //page 3 script 
            macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:bpoForm ATTR=TXT:Page3");
            macro.AppendLine(@"WAIT SECONDS=3");

            macro.AppendLine(@"TAG POS=66 TYPE=SELECT FORM=NAME:bpoForm ATTR=* CONTENT=%boolean:false");
            macro.AppendLine(@"TAG POS=67 TYPE=SELECT FORM=NAME:bpoForm ATTR=* CONTENT=%boolean:false");
            macro.AppendLine(@"TAG POS=68 TYPE=SELECT FORM=NAME:bpoForm ATTR=* CONTENT=%boolean:false");
            macro.AppendLine(@"TAG POS=73 TYPE=SELECT FORM=NAME:bpoForm ATTR=* CONTENT=%boolean:false");
            macro.AppendLine(@"TAG POS=74 TYPE=SELECT FORM=NAME:bpoForm ATTR=* CONTENT=%number:2851");
            macro.AppendLine(@"TAG POS=69 TYPE=SELECT FORM=NAME:bpoForm ATTR=* CONTENT=%boolean:false");
            macro.AppendLine(@"TAG POS=70 TYPE=SELECT FORM=NAME:bpoForm ATTR=* CONTENT=%boolean:false");
            macro.AppendLine(@"TAG POS=71 TYPE=SELECT FORM=NAME:bpoForm ATTR=* CONTENT=%boolean:false");
            macro.AppendLine(@"TAG POS=72 TYPE=SELECT FORM=NAME:bpoForm ATTR=* CONTENT=%boolean:false");

            //commands.Add(30.001, new string[] { "1", inputType[elem.TEXTAREA], "AsIsRepairAnalysis", "No change." });
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:bpoForm ATTR=ng-model:*AsIsRepairAnalysis CONTENT=No<SP>Change.");


             macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 30);



            //commands.Add(6, new string[] { "1", inputType[elem.TEXT], "PROP_FOR_SALE_L", callingForm.SubjectNeighborhood.numberOfCompListings.ToString()});
            //commands.Add(8, new string[] { "1", inputType[elem.TEXT], "LIST_LOW_S", callingForm.SubjectNeighborhood.minListPrice.ToString() });
            //commands.Add(9, new string[] { "1", inputType[elem.TEXT], "LIST_HIGH_S", callingForm.SubjectNeighborhood.maxListPrice.ToString() });
            //commands.Add(10, new string[] { "1", inputType[elem.TEXT], "SALE_COUNT", callingForm.SubjectNeighborhood.numberSoldListings.ToString() });
            //commands.Add(11, new string[] { "1", inputType[elem.TEXT], "SALE_LOW_S", callingForm.SubjectNeighborhood.minSalePrice.ToString() });
            //commands.Add(12, new string[] { "1", inputType[elem.TEXT], "SALE_HIGH_S", callingForm.SubjectNeighborhood.maxSalePrice.ToString() });
            //commands.Add(13, new string[] { "1", inputType[elem.TEXT], "BOARDED_UP_HOMES_NBR", "0" });
            //commands.Add(14, new string[] { "1", inputType[elem.TEXT], "DOM_AVG_L", callingForm.SubjectNeighborhood.avgDomSold.ToString() });
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
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");

            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT ATTR=NAME:*Address" + compNumpTranslation[targetCompString] + " CONTENT=" + targetComp.StreetAddress.Replace(" ", "<SP>"));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT ATTR=NAME:*ListDt" + compNumpTranslation[targetCompString] + "* CONTENT=" + targetComp.ListDateString.Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT ATTR=NAME:*CurrentListPrice" + compNumpTranslation[targetCompString] + " CONTENT=" + targetComp.CurrentListPrice.ToString().Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            if (saleOrList == "sale")
            {
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT ATTR=NAME:*SaleDt" + compNumpTranslation[targetCompString] + "* CONTENT=" + targetComp.SalesDate.ToShortDateString().Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT ATTR=NAME:*SalesPrice" + compNumpTranslation[targetCompString] + " CONTENT=" + targetComp.SalePrice.ToString().Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
          
            }
      
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT ATTR=NAME:*DOM" + compNumpTranslation[targetCompString] + " CONTENT=" + targetComp.DOM.Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT ATTR=NAME:*Sqft" + compNumpTranslation[targetCompString] + " CONTENT=" + targetComp.GLA.ToString().Replace(" ", "<SP>").Replace("$", "").Replace(",", "").Replace("-1", "0"));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT ATTR=NAME:*Bedrooms" + compNumpTranslation[targetCompString] + " CONTENT=" + targetComp.BedroomCount.Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT ATTR=NAME:*Baths" + compNumpTranslation[targetCompString] + " CONTENT=" + targetComp.BathroomCount.Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            if (targetComp.AttachedGarage())
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:*GarageType" + compNumpTranslation[targetCompString] + " CONTENT=" + garageTranslation[targetComp.NumberGarageStalls() + " Attached"].Replace("number:", ""));
            }
            else if (targetComp.DetachedGarage())
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:*GarageType" + compNumpTranslation[targetCompString] + " CONTENT=" + garageTranslation[targetComp.NumberGarageStalls() + " Detached"].Replace("number:", ""));

               
            }
            else
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:*GarageType" + compNumpTranslation[targetCompString] + " CONTENT=%1976");

            }

            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT ATTR=NAME:*YearBuilt" + compNumpTranslation[targetCompString] + " CONTENT=" + targetComp.YearBuiltString.Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));

            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA ATTR=NAME:*Comment" + compNumpTranslation[targetCompString] + " CONTENT=" + targetComp.mlsHtmlFields["remarks"].value.Replace("$", "").Replace(",", "").Replace(" ", "<SP>"));
            //TAG POS=1 TYPE=TEXTAREA ATTR=NAME:*Comment2 CONTENT=fuck<SP>is<SP>a<SP>comment

            ///
            //04-28-18 updates
            ///

            //Concessions
            commands.Add(0.001, new string[] { "1", inputType[elem.TEXT], "Concessions", "0" });


            commands.Add(0, new string[] { "1", inputType[elem.TEXT], "Address", targetComp.StreetAddress });
            commands.Add(0.1, new string[] { "1", inputType[elem.TEXT], "City", targetComp.City.Replace(" ", "<SP>") });
            commands.Add(0.2, new string[] { "1", inputType[elem.SELECT], "State", "%string:IL" });
            commands.Add(0.3, new string[] { "1", inputType[elem.TEXT], "Zip", targetComp.Zipcode });
            commands.Add(0.4, new string[] { "1", inputType[elem.TEXT], "UNITS", "1" });
            commands.Add(1, new string[] { "1", inputType[elem.TEXT], "Proximity", targetComp.proximityToSubject.ToString() });
            commands.Add(2, new string[] { "1", inputType[elem.SELECT], "PropertyTypeID", "%number:397" });
            commands.Add(3, new string[] { "1", inputType[elem.SELECT], "HomeStyleID", "%number:426" });
            commands.Add(4, new string[] { "1", inputType[elem.TEXT], "TotalRents", "1000" });
            commands.Add(5, new string[] { ListingDatePosition[targetCompString], inputType[elem.TEXT], "dt", targetComp.ListDateString });
            commands.Add(6, new string[] { "1", inputType[elem.TEXT], "OrigListPrice", targetComp.OriginalListPrice.ToString() });
            commands.Add(7, new string[] { "1", inputType[elem.TEXT], "FinalListPrice", targetComp.CurrentListPrice.ToString() });
            commands.Add(8, new string[] { "1", inputType[elem.TEXT], "SalesPrice", targetComp.SalePrice.ToString() });
            commands.Add(8.1, new string[] { "1", inputType[elem.SELECT], "FinanceTypeID", financingTypeTranslation[targetComp.FinancingMlsString] });
            commands.Add(9, new string[] { SoldDatePosition[targetCompString], inputType[elem.TEXT], "dt", targetComp.SalesDate.ToShortDateString() });
            commands.Add(10, new string[] { "1", inputType[elem.TEXT], "DOM", targetComp.DOM });
            commands.Add(11, new string[] { "1", inputType[elem.TEXT], "YearBuilt", targetComp.YearBuiltString });
            commands.Add(12, new string[] { "1", inputType[elem.TEXT], "LotSize", targetComp.Lotsize.ToString() });
            commands.Add(13, new string[] { "1", inputType[elem.SELECT], "ConditionTypeID", "%number:432" });
            commands.Add(14, new string[] { "1", inputType[elem.SELECT], "ConstructionTypeID", "%number:430" });
            commands.Add(15, new string[] { "1", inputType[elem.TEXT], "TotalRooms", targetComp.TotalRoomCount.ToString() });
            commands.Add(16, new string[] { "1", inputType[elem.TEXT], "BedRooms", targetComp.BedroomCount });
            commands.Add(17, new string[] { "1", inputType[elem.TEXT], "FullBaths", targetComp.FullBathCount });
            commands.Add(18, new string[] { "1", inputType[elem.TEXT], "HalfBaths", targetComp.HalfBathCount });
            commands.Add(19, new string[] { "1", inputType[elem.TEXT], "AboveGradeSqft", targetComp.ProperGla(GlobalVar.mainWindow.SubjectAboveGLA) });
            commands.Add(20, new string[] { "1", inputType[elem.TEXT], "BelowGradeSqft", targetComp.BasementGLA() });
            commands.Add(21, new string[] { "1", inputType[elem.TEXT], "BasementPercentFinished", targetComp.BasementFinishedPercentage() });
            commands.Add(22, new string[] { "1", inputType[elem.TEXT], "ViewDescript", "Residential" });
            commands.Add(23, new string[] { "1", inputType[elem.TEXT], "MLSNumber", targetComp.MlsNumber });

            //TODO: add hasCarport support and hasExtParking support
            //GarageTypeID
            if (targetComp.AttachedGarage())
            {
                commands.Add(24, new string[] { "1", inputType[elem.SELECT], "GarageTypeID", garageTranslation[targetComp.NumberGarageStalls() + " Attached"] });
            }
            else if (targetComp.DetachedGarage())
            {
                commands.Add(24, new string[] { "1", inputType[elem.SELECT], "GarageTypeID", garageTranslation[targetComp.NumberGarageStalls() + " Detached"] });
            }
            else
            {
                commands.Add(24, new string[] { "1", inputType[elem.SELECT], "GarageTypeID", "%number:1976" });

            }
            //TODO:  add hasPool support
            //PoolFireID
            if (String.IsNullOrWhiteSpace(targetComp.NumberOfFireplaces) ||   targetComp.NumberOfFireplaces == "0")
            {
                commands.Add(25, new string[] { "1", inputType[elem.SELECT], "PoolFireID", "%number:453" });
            }
            else
            {
                commands.Add(25, new string[] { "1", inputType[elem.SELECT], "PoolFireID", "%number:451" });
            }

            //SaleType  saleTypeTranslation
            commands.Add(26, new string[] { "1", inputType[elem.SELECT], "SaleType",  saleTypeTranslation[targetComp.TransactionType]});

            //Comment
            commands.Add(27, new string[] { "1", inputType[elem.TEXTAREA], "Comment", targetComp.mlsHtmlFields["remarks"].value });





        
            string tcs = targetCompString;
            foreach (var c in commands)
            {
                if (c.Value[2] == "dt")
                {
                    tcs = "";
                }
                else
                {
                    tcs = targetCompString;
                }
                macro.AppendFormat("TAG POS={0} TYPE={1} ATTR=ng-model:*{3}*{2} CONTENT={4}\r\n", c.Value[0], c.Value[1], c.Value[2], tcs, c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            }

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

