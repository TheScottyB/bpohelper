using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{

    class ClearCap : BPOFulfillment
    {
       
        public ClearCap()
        {
            subject = GlobalVar.theSubjectProperty;
            callingForm = GlobalVar.mainWindow;
        }

        public ClearCap(MLSListing m) 
            : this()
        {
            targetComp = m;
        }

        enum elem { TEXT, SELECT, TEXTAREA, RADIO };

         private  Dictionary<elem, string> inputType = new Dictionary<elem, string>()
        {
            {elem.TEXT, "INPUT:TEXT"},{elem.SELECT, "SELECT"}, {elem.TEXTAREA, "TEXTAREA"}, {elem.RADIO, "INPUT:RADIO"}
        };

    
        private Dictionary<int, MLSListing> stack;

        private Form1 callingForm;
        private SubjectProperty subject;

        public void Prefill(iMacros.App iim, Form1 form)
        {
            Dictionary<int, string[]> commands = new Dictionary<int, string[]>();
            Dictionary<string, string> translationTable = new Dictionary<string, string>();
            StringBuilder macro = new StringBuilder();

 //           'Complete Page
 //TAG POS=1 TYPE=HTML ATTR=* EXTRACT=HTM 
 //'Complete Page TEXT only
 //TAG POS=1 TYPE=HTML ATTR=* EXTRACT=TXT

            macro.AppendLine(@"TAG POS=1 TYPE=HTML ATTR=* EXTRACT=TXT");
            iim.iimPlayCode(macro.ToString(), 30);
            string reportType = iim.iimGetLastExtract();

            if (reportType.Contains("WEB2.20160715.022508.FORMV135 "))
            {
                //WEB2.20160715.022508.FORMV135
                #region original
                commands.Add(0, new string[] { "1", inputType[elem.TEXT], "apn_s", GlobalVar.theSubjectProperty.ParcelID });
                //SFR
                commands.Add(1, new string[] { "1", inputType[elem.RADIO], "PROPERTYTYPE", "YES" });
                //R1
                commands.Add(2, new string[] { "1", inputType[elem.TEXT], "ZONING", "A1-R1" });
                //SFR - maybe get from realist
                commands.Add(3, new string[] { "1", inputType[elem.TEXT], "ZONING_DESC_1", @"SFR w/Horses" });
                //conforming
                commands.Add(4, new string[] { "1", inputType[elem.SELECT], "SUBJECT_USE_CODE_1", "%L" });

                commands.Add(5, new string[] { "1", inputType[elem.TEXT], "LIST_COUNT", callingForm.SubjectNeighborhood.numberActiveListings.ToString() });
                commands.Add(6, new string[] { "1", inputType[elem.TEXT], "PROP_FOR_SALE_L", callingForm.SubjectNeighborhood.numberOfCompListings.ToString() });
                commands.Add(7, new string[] { "1", inputType[elem.TEXT], "REO_CORP_OWNED_NBR", callingForm.SetOfComps.numberREOListings.ToString() });
                commands.Add(8, new string[] { "1", inputType[elem.TEXT], "LIST_LOW_S", callingForm.SubjectNeighborhood.minListPrice.ToString() });
                commands.Add(9, new string[] { "1", inputType[elem.TEXT], "LIST_HIGH_S", callingForm.SubjectNeighborhood.maxListPrice.ToString() });
                commands.Add(10, new string[] { "1", inputType[elem.TEXT], "SALE_COUNT", callingForm.SubjectNeighborhood.numberSoldListings.ToString() });
                commands.Add(11, new string[] { "1", inputType[elem.TEXT], "SALE_LOW_S", callingForm.SubjectNeighborhood.minSalePrice.ToString() });
                commands.Add(12, new string[] { "1", inputType[elem.TEXT], "SALE_HIGH_S", callingForm.SubjectNeighborhood.maxSalePrice.ToString() });
                commands.Add(13, new string[] { "1", inputType[elem.TEXT], "BOARDED_UP_HOMES_NBR", "0" });
                commands.Add(14, new string[] { "1", inputType[elem.TEXT], "DOM_AVG_L", callingForm.SubjectNeighborhood.avgDomSold.ToString() });
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


                foreach (var c in commands)
                {
                    macro.AppendFormat("TAG POS={0} TYPE={1} FORM=ID:BPO_FORM ATTR=NAME:{2} CONTENT={3}\r\n", c.Value[0], c.Value[1], c.Value[2], c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
                }

                #endregion

            } else {

                #region form159
                if (GlobalVar.theSubjectProperty.ToString().Contains("Attached"))
                {
                    commands.Add(0, new string[] { "1", inputType[elem.SELECT], "PROPERTY_TYPE", "%Condominium" });
                }
                else
                {
                    commands.Add(0, new string[] { "1", inputType[elem.SELECT], "PROPERTY_TYPE", "%SF<SP>Detached" });
                }
                commands.Add(1, new string[] { "1", inputType[elem.SELECT], "PROPERTY_OCCUPANCY", "%Owner" });
                commands.Add(2, new string[] { "1", inputType[elem.SELECT], "PROPERTY_TOTAL_ROOMS", "%" + callingForm.SubjectRoomCount });
                commands.Add(3, new string[] { "1", inputType[elem.SELECT], "PROPERTY_BEDROOMS", "%" + callingForm.SubjectBedroomCount });
                if (GlobalVar.theSubjectProperty.HalfBathCount == "0")
                {
                    commands.Add(4, new string[] { "1", inputType[elem.SELECT], "PROPERTY_BATHS", "%" + GlobalVar.theSubjectProperty.FullBathCount });
                }
                else
                {
                    commands.Add(4, new string[] { "1", inputType[elem.SELECT], "PROPERTY_BATHS", "%" + GlobalVar.theSubjectProperty.FullBathCount + ".5" });
                }
                commands.Add(5, new string[] { "1", inputType[elem.TEXT], "PROPERTY_SQFT", callingForm.SubjectAboveGLA });
                commands.Add(6, new string[] { "1", inputType[elem.SELECT], "PROPERTY_LOCATION", "%Good" });
                commands.Add(7, new string[] { "1", inputType[elem.TEXT], "PROPERTY_LOTSIZE", callingForm.SubjectLotSize });
                commands.Add(8, new string[] { "1", inputType[elem.SELECT], "PROPERTY_LOTSIZE_TYPE_S", "%A" });
                //
                //HOME_STYLE_S_1
                //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:BPO_FORM ATTR=NAME:HOME_STYLE_S_1 CONTENT=%A-Frame");
                commands.Add(9, new string[] { "1", inputType[elem.TEXT], "PROPERTY_AGE", GlobalVar.theSubjectProperty.Age.ToString() });
                if (callingForm.SubjectParkingType.Contains("Attached"))
                {
                    commands.Add(10, new string[] { "1", inputType[elem.SELECT], "PROPERTY_GARAGE", "%" + GlobalVar.theSubjectProperty.GarageStallCount + "<SP>Attached" });
                }
                else if (callingForm.SubjectParkingType.Contains("Detached"))
                {
                    commands.Add(10, new string[] { "1", inputType[elem.SELECT], "PROPERTY_GARAGE", "%" + GlobalVar.theSubjectProperty.GarageStallCount + "<SP>Detached" });
                }
                else
                {
                    commands.Add(10, new string[] { "1", inputType[elem.SELECT], "PROPERTY_GARAGE", "%None" });
                }
                commands.Add(11, new string[] { "1", inputType[elem.SELECT], "PROPERTY_LANDSCAPING", "%2" });
                commands.Add(12, new string[] { "2", inputType[elem.RADIO], "REPAIRS_RECOMMEND", "YES" });
                commands.Add(13, new string[] { "1", inputType[elem.SELECT], "REPAIRS_OVERALL_CONDITION", "%Good" });
                commands.Add(14, new string[] { "2", inputType[elem.RADIO], "REPAIRS_IMMEDIATE_ACTION", "YES" });
                commands.Add(15, new string[] { "2", inputType[elem.RADIO], "REPAIRS_LEGAL_ISSUES", "YES" });
                commands.Add(16, new string[] { "2", inputType[elem.RADIO], "REPAIRS_ENVIRONMENTAL_ISSUES", "YES" });
                //
                //REPAIRS_COMMENTS
                //TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPO_FORM ATTR=NAME:REPAIRS_COMMENTS CONTENT=tbd
                commands.Add(17, new string[] { "1", inputType[elem.SELECT], "neighborhood_property_values", "%Stable" });
                commands.Add(18, new string[] { "1", inputType[elem.SELECT], "NEIGHBORHOOD_PREDOMINANT_OCCUPANCY", "%Owner" });
                commands.Add(19, new string[] { "1", inputType[elem.TEXT], "normmarketday", callingForm.SubjectNeighborhood.avgDom.ToString() });
                commands.Add(20, new string[] { "1", inputType[elem.SELECT], "NEIGHBORHOOD_VACANCY_RATE", "%5-10%" });
                commands.Add(21, new string[] { "1", inputType[elem.TEXT], "NEIGHBORHOOD_PRICE_RANGE_LOW", callingForm.SubjectNeighborhood.minListPrice.ToString() });
                commands.Add(22, new string[] { "1", inputType[elem.TEXT], "NEIGHBORHOOD_PRICE_RANGE_HIGH", callingForm.SubjectNeighborhood.maxListPrice.ToString() });
                commands.Add(23, new string[] { "2", inputType[elem.RADIO], "DEMAND_SUPPLY_S", "YES" });
                commands.Add(24, new string[] { "1", inputType[elem.TEXT], "NEIGHBORHOOD_ACTIVE_LISTINGS", callingForm.SubjectNeighborhood.numberActiveListings.ToString() });
                commands.Add(25, new string[] { "1", inputType[elem.TEXT], "NEIGHBORHOOD_ACTIVE_LISTING_RADIUS", "1" });
                commands.Add(26, new string[] { "2", inputType[elem.RADIO], "LOCATION_S", "YES" });
                //
                //NEIGHBORHOOD_COMMENTS
                //TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPO_FORM ATTR=NAME:NEIGHBORHOOD_COMMENTS CONTENT=tbd
                commands.Add(27, new string[] { "1", inputType[elem.TEXT], "VALUATION_AS_IS_LIST_90", callingForm.SubjectMarketValueList });
                commands.Add(28, new string[] { "1", inputType[elem.TEXT], "VALUATION_AS_IS_LIST_120", callingForm.SubjectMarketValueList });
                commands.Add(29, new string[] { "1", inputType[elem.TEXT], "VALUATION_AS_IS_LIST_180", callingForm.SubjectMarketValueList });
                commands.Add(30, new string[] { "1", inputType[elem.TEXT], "VALUATION_AS_IS_90", callingForm.SubjectMarketValue });
                commands.Add(31, new string[] { "1", inputType[elem.TEXT], "VALUATION_AS_IS_120", callingForm.SubjectMarketValue });
                commands.Add(32, new string[] { "1", inputType[elem.TEXT], "VALUATION_AS_IS_180", callingForm.SubjectMarketValue });
                commands.Add(33, new string[] { "1", inputType[elem.TEXT], "VALUATION_REPAIRED_LIST_90", callingForm.SubjectMarketValueList });
                commands.Add(34, new string[] { "1", inputType[elem.TEXT], "VALUATION_REPAIRED_LIST_120", callingForm.SubjectMarketValueList });
                commands.Add(35, new string[] { "1", inputType[elem.TEXT], "VALUATION_REPAIRED_LIST_180", callingForm.SubjectMarketValueList });
                commands.Add(36, new string[] { "1", inputType[elem.TEXT], "VALUATION_REPAIRED_90", callingForm.SubjectMarketValue });
                commands.Add(37, new string[] { "1", inputType[elem.TEXT], "VALUATION_REPAIRED_120", callingForm.SubjectMarketValue });
                commands.Add(38, new string[] { "1", inputType[elem.TEXT], "VALUATION_REPAIRED_180", callingForm.SubjectMarketValue });
                commands.Add(40, new string[] { "1", inputType[elem.SELECT], "VALUATION_LIST_PROPERTY_AS", "%AS-IS" });
                commands.Add(41, new string[] { "1", inputType[elem.TEXT], "VALUATION_SELLER_PAID_FINANCING", "0" });
                commands.Add(42, new string[] { "1", inputType[elem.TEXT], "VALUE_RENT_S", callingForm.SubjectRent });
                //
                //VALUATION_COMMENTS
                //TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPO_FORM ATTR=NAME:VALUATION_COMMENTS

                if (callingForm.SubjectBasementType == "None")
                {
                    //commands.Add(14.11, new string[] { "2", inputType[elem.RADIO], "BMST_B", "YES" });
                    //commands.Add(14.12, new string[] { "1", inputType[elem.TEXT], "BSMT_SQFEET_S", "0" });
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:BSMT_B_1 CONTENT=YES");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:BSMT_SQFEET_S_1 CONTENT=0");
                }
                else
                {
                    //commands.Add(14.21, new string[] { "1", inputType[elem.RADIO], "BMST_B", "YES" });
                    //commands.Add(14.22, new string[] { "1", inputType[elem.TEXT], "BSMT_PERCENT_S", targetComp.BasementFinishedPercentage() });
                    //commands.Add(14.23, new string[] { "1", inputType[elem.TEXT], "BSMT_SQFEET_S", targetComp.BasementGLA() });
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:BSMT_B_1 CONTENT=YES");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:BSMT_PERCENT_S_1 CONTENT=" + GlobalVar.theSubjectProperty.BasementFinishedPercentage.ToString());
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:BSMT_SQFEET_S_1 CONTENT=" + callingForm.SubjectBasementGLA);
                }



                foreach (var c in commands)
                {
                    macro.AppendFormat("TAG POS={0} TYPE={1} FORM=ID:BPO_FORM ATTR=NAME:{2} CONTENT={3}\r\n", c.Value[0], c.Value[1], c.Value[2], c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
                }

                #endregion
            }

          

     


           string macroCode = macro.ToString();
           iim.iimPlayCode(macroCode, 30);
        }

  
        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {
            StringBuilder macro = new StringBuilder();
            targetCompNumber = Regex.Match(compNum, @"\d").Value;

               macro.AppendLine(@"TAG POS=1 TYPE=HTML ATTR=* EXTRACT=TXT");
            iim.iimPlayCode(macro.ToString(), 30);
            string reportType = iim.iimGetLastExtract();

            if (reportType.Contains("WEB2.20160715.022508.FORMV135 "))
            {
                #region oldForm
                Dictionary<string, string> saleTypeTranslation = new Dictionary<string, string>()
                {
                    {"Arms Length", "19"}, {"REO", "1"}, {"ShortSale", "17"},
                };
                Dictionary<string, string> heatingTypeTranslation = new Dictionary<string, string>()
                {
                    {"Gas", "6"}, {"Gas, Forced Air", "6"}, {@"Gas, Hot Water/Steam", "6"},{@"Electric", "2"}
                

                };

                Dictionary<string, string> CoolingTypeTranslation = new Dictionary<string, string>()
                {
                    {"Central Air", "1"},  {"None", "0"},     {@"1 (Window/Wall Unit)", "2"}
                };



                Dictionary<double, string[]> commands = new Dictionary<double, string[]>();


                commands.Add(0, new string[] { "1", inputType[elem.TEXT], "ADDRESS", targetComp.StreetAddress });
                commands.Add(1, new string[] { "1", inputType[elem.TEXT], "ZIP", targetComp.Zipcode });
                if (saleOrList == "list")
                {
                    commands.Add(2, new string[] { "1", inputType[elem.TEXT], "LIST_PRICE", targetComp.CurrentListPrice.ToString() });
                    commands.Add(3, new string[] { "1", inputType[elem.TEXT], "ORIGLIST_PRICE", targetComp.OriginalListPrice.ToString() });
                }

                commands.Add(4, new string[] { "1", inputType[elem.TEXT], "PRICE_REDUCTION_COUNT", targetComp.NumberOfPriceChanges.ToString() });
                commands.Add(5, new string[] { "1", inputType[elem.SELECT], "DATASOURCE_S", "%MLS" });
                commands.Add(6, new string[] { "1", inputType[elem.TEXT], "SELLER_CONCESSIONS", targetComp.PointsInDollars });



                commands.Add(7, new string[] { "1", inputType[elem.SELECT], "SALES_TYPE", "%" + saleTypeTranslation[targetComp.TransactionType] });
                commands.Add(8, new string[] { "1", inputType[elem.TEXT], "DAYS_ONMKT", targetComp.DOM });


                commands.Add(9, new string[] { "1", inputType[elem.SELECT], "LOCATION_DESCR", "%" + GlobalVar.mainWindow.comboBoxLocationDescr.Text });
                commands.Add(10, new string[] { "1", inputType[elem.SELECT], "OWNERSHIP_TYPE", "%Fee Simple" });
                commands.Add(11, new string[] { "1", inputType[elem.TEXT], "LOTSIZE", targetComp.Lotsize.ToString() });
                commands.Add(12, new string[] { "1", inputType[elem.SELECT], "LOTSIZE_TYPE_S", "%A" });

                if (targetComp.Waterfront)
                {
                    commands.Add(13, new string[] { "1", inputType[elem.TEXT], "PROPERTY_VIEW", "Water" });
                }
                else
                {
                    commands.Add(13, new string[] { "1", inputType[elem.TEXT], "PROPERTY_VIEW", "Residential" });
                }
                commands.Add(14, new string[] { "1", inputType[elem.SELECT], "CONSTRUCTION_QUALITY", "%Average" });
                commands.Add(15, new string[] { "1", inputType[elem.TEXT], "YEAR_BUILT", targetComp.YearBuiltString });
                commands.Add(16, new string[] { "1", inputType[elem.SELECT], "CONDITION", "%Average" });
                commands.Add(17, new string[] { "1", inputType[elem.TEXT], "UNIT_NUM_S", "1" });
                commands.Add(18, new string[] { "1", inputType[elem.TEXT], "TOTAL_ROOMS_NUM", targetComp.TotalRoomCount.ToString() });
                commands.Add(19, new string[] { "1", inputType[elem.TEXT], "BEDROOMS_NUM", targetComp.BedroomCount });
                commands.Add(20, new string[] { "1", inputType[elem.TEXT], "BATHS_NUM", targetComp.BathroomCount });

                commands.Add(21, new string[] { "1", inputType[elem.TEXT], "LIVING_SQFEET", targetComp.ProperGla(GlobalVar.mainWindow.SubjectAboveGLA) });
                commands.Add(22, new string[] { "1", inputType[elem.TEXT], "BSMT_SQFEET_S", targetComp.BasementGLA() });
                commands.Add(23, new string[] { "1", inputType[elem.TEXT], "ROOMS_BELOW_GRADE", targetComp.NumberOfBasementRooms() });


                commands.Add(24, new string[] { "1", inputType[elem.SELECT], "HEATING", "%" + heatingTypeTranslation[targetComp.Heating] });
                commands.Add(25, new string[] { "1", inputType[elem.SELECT], "COOLING", "%" + CoolingTypeTranslation[targetComp.Cooling] });

                if (targetComp.AttachedGarage())
                {
                    commands.Add(26, new string[] { "1", inputType[elem.SELECT], "GARAGE_TYPE_S", "%AG" });
                }
                else if (targetComp.DetachedGarage())
                {
                    commands.Add(26, new string[] { "1", inputType[elem.SELECT], "GARAGE_TYPE_S", "%DG" });
                }

                commands.Add(27, new string[] { "1", inputType[elem.SELECT], "GARAGE_CAR_NUM_L", "%" + targetComp.NumberGarageStalls() });
                commands.Add(28, new string[] { "1", inputType[elem.SELECT], "COOLING", "%" + CoolingTypeTranslation[targetComp.Cooling] });

                commands.Add(29, new string[] { "1", inputType[elem.TEXT], "OTHER", "NA" });
                commands.Add(30, new string[] { "2", inputType[elem.RADIO], "POOL_B", "YES" });
                if (saleOrList == "sale")
                {
                    commands.Add(31, new string[] { "1", inputType[elem.TEXT], "SALE_PRICE", targetComp.SalePrice.ToString() });
                    commands.Add(32, new string[] { "1", inputType[elem.TEXT], "CURRENTLIST_PRICE", targetComp.CurrentListPrice.ToString() });
                    commands.Add(33, new string[] { "1", inputType[elem.TEXT], "DATEOFSALE", targetComp.SalesDate.ToShortDateString() });
                }


                foreach (var c in commands)
                {
                    macro.AppendFormat("TAG POS={0} TYPE={1} FORM=ID:BPO_FORM ATTR=NAME:{2}_{3} CONTENT={4}\r\n", c.Value[0], c.Value[1], c.Value[2], targetCompNumber, c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
                }

                #endregion
            }
            else
            {
                #region newForm
                Dictionary<string, string> saleTypeTranslation = new Dictionary<string, string>()
            {
                {"Arms Length", "2"}, {"REO", "1"}, {"ShortSale", "17"},
            };
                Dictionary<string, string> heatingTypeTranslation = new Dictionary<string, string>()
            {
                {"Gas", "6"}, {"Gas, Forced Air", "6"}, {@"Gas, Hot Water/Steam", "6"},{@"Electric", "2"},

            };

                Dictionary<string, string> CoolingTypeTranslation = new Dictionary<string, string>()
            {
                {"Central Air", "1"},  {"None", "0"}
            };

                Dictionary<string, string> PropertyTypeTranslation = new Dictionary<string, string>()
            {
                {"bpohelper.AttachedListing", "Condominium"},  {"bpohelper.DetachedListing", "SF Detached"}
            };


                Dictionary<double, string[]> commands = new Dictionary<double, string[]>();
                // string[] inputType = { "INPUT:TEXT", "SELECT", "TEXTAREA", "INPUT:RADIO" };




                macro.AppendLine(@"SET !ERRORIGNORE YES");
                macro.AppendLine(@"SET !TIMEOUT_STEP 0");
                // -WEB1.20160715.065437.FORMV159


                commands.Add(0, new string[] { "1", inputType[elem.TEXT], "ADDRESS", targetComp.StreetAddress });
                commands.Add(1, new string[] { "1", inputType[elem.TEXT], "ZIP_CODE", targetComp.Zipcode });
                commands.Add(2, new string[] { "1", inputType[elem.TEXT], "MLS_NUMBER", targetComp.MlsNumber });
                if (saleOrList == "list")
                {
                    commands.Add(3.11, new string[] { "1", inputType[elem.TEXT], "CUR_LIST_PRICE", targetComp.CurrentListPrice.ToString() });
                    commands.Add(3.12, new string[] { "1", inputType[elem.TEXT], "CUR_LIST_DATE", targetComp.DateOfLastPriceChange.ToShortDateString() });
                }
                else
                {
                    commands.Add(3.21, new string[] { "1", inputType[elem.TEXT], "ORG_LIST_PRICE", targetComp.OriginalListPrice.ToString() });
                    commands.Add(3.22, new string[] { "1", inputType[elem.TEXT], "CUR_LIST_PRICE", targetComp.CurrentListPrice.ToString() });
                    commands.Add(3.23, new string[] { "1", inputType[elem.TEXT], "SALE_PRICE", targetComp.SalePrice.ToString() });
                    commands.Add(3.24, new string[] { "1", inputType[elem.TEXT], "SALE_DATE", targetComp.SalesDate.ToShortDateString() });
                    commands.Add(3.25, new string[] { "1", inputType[elem.TEXT], "DAYS_ON_MARKET", targetComp.DOM });
                }
                commands.Add(4, new string[] { "1", inputType[elem.SELECT], "SALES_TYPE", "%" + saleTypeTranslation[targetComp.TransactionType] });
                //commands.Add(5, new string[] { "1", inputType[elem.SELECT], "PROPERTY_TYPE_S", "%" + PropertyTypeTranslation[targetComp.ToString()] });
                // commands.Add(6, new string[] { "1", inputType[elem.SELECT], "HOME_STYLE_S", "%" + saleTypeTranslation[targetComp.TransactionType] });
                commands.Add(7, new string[] { "1", inputType[elem.TEXT], "AGE", targetComp.Age.ToString() });
                commands.Add(8, new string[] { "1", inputType[elem.SELECT], "CONDITION", "%Good" });
                commands.Add(9, new string[] { "1", inputType[elem.SELECT], "TOTAL_ROOMS", "%" + targetComp.TotalRoomCount.ToString() });
                commands.Add(10, new string[] { "1", inputType[elem.SELECT], "BEDROOMS", "%" + targetComp.BedroomCount.ToString() });
                if (targetComp.HalfBathCount == "0")
                {
                    commands.Add(11.1, new string[] { "1", inputType[elem.SELECT], "BATHS", "%" + targetComp.FullBathCount.ToString() });
                }
                else
                {
                    commands.Add(11.2, new string[] { "1", inputType[elem.SELECT], "BATHS", "%" + targetComp.FullBathCount.ToString() + ".5" });
                }
                commands.Add(12, new string[] { "1", inputType[elem.TEXT], "SQFT", targetComp.ProperGla(GlobalVar.mainWindow.SubjectAboveGLA) });
                // commands.Add(13, new string[] { "1", inputType[elem.SELECT], "GARAGE", "%" + saleTypeTranslation[targetComp.TransactionType] });
                //      if (string.IsNullOrWhiteSpace(targetComp.BasementGLA()) || targetComp.BasementGLA() == "-1" || targetComp.BasementGLA() == "0")
                if (targetComp.BasementType == "None")
                {
                    //commands.Add(14.11, new string[] { "2", inputType[elem.RADIO], "BMST_B", "YES" });
                    //commands.Add(14.12, new string[] { "1", inputType[elem.TEXT], "BSMT_SQFEET_S", "0" });
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:BSMT_B_" + targetCompNumber + " CONTENT=YES");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:BSMT_SQFEET_S_" + targetCompNumber + " CONTENT=0");
                }
                else
                {
                    //commands.Add(14.21, new string[] { "1", inputType[elem.RADIO], "BMST_B", "YES" });
                    //commands.Add(14.22, new string[] { "1", inputType[elem.TEXT], "BSMT_PERCENT_S", targetComp.BasementFinishedPercentage() });
                    //commands.Add(14.23, new string[] { "1", inputType[elem.TEXT], "BSMT_SQFEET_S", targetComp.BasementGLA() });
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:BSMT_B_" + targetCompNumber + " CONTENT=YES");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:BSMT_PERCENT_S_" + targetCompNumber + " CONTENT=" + targetComp.BasementFinishedPercentage());
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:BSMT_SQFEET_S_" + targetCompNumber + " CONTENT=" + targetComp.BasementGLA());
                }
                //EXT_SIDING_S
                // commands.Add(15, new string[] { "1", inputType[elem.SELECT], "EXT_SIDING_S", "%" + saleTypeTranslation[targetComp.TransactionType] });
                //PORCH
                //commands.Add(16, new string[] { "1", inputType[elem.TEXT], "PORCH", targetComp.BasementGLA() });
                commands.Add(17, new string[] { "1", inputType[elem.TEXT], "LOTSIZE", targetComp.Lotsize.ToString() });
                commands.Add(18, new string[] { "1", inputType[elem.SELECT], "LOTSIZE_TYPE_S", "%A" });
                if (targetComp.GarageExsists())
                {
                    commands.Add(19.1, new string[] { "1", inputType[elem.SELECT], "GARAGE", "%" + targetComp.NumberGarageStalls() + " " + targetComp.GarageType() });
                }
                else
                {
                    commands.Add(19.2, new string[] { "1", inputType[elem.SELECT], "GARAGE", "%None" });
                }


                //if (targetComp.AttachedGarage())
                //{
                //    commands.Add(26, new string[] { "1", inputType[elem.SELECT], "GARAGE_TYPE_S", "%AG" });
                //}
                //else if (targetComp.DetachedGarage())
                //{
                //    commands.Add(26, new string[] { "1", inputType[elem.SELECT], "GARAGE_TYPE_S", "%DG" });
                //}

                //commands.Add(27, new string[] { "1", inputType[elem.SELECT], "GARAGE_CAR_NUM_L", "%" + targetComp.NumberGarageStalls() });


                //commands.Add(4, new string[] { "1", inputType[elem.TEXT], "PRICE_REDUCTION_COUNT", targetComp.NumberOfPriceChanges.ToString() });
                //commands.Add(5, new string[] { "1", inputType[elem.SELECT], "DATASOURCE_S", "%MLS" });
                //commands.Add(6, new string[] { "1", inputType[elem.TEXT], "SELLER_CONCESSIONS", targetComp.PointsInDollars });




                //commands.Add(8, new string[] { "1", inputType[elem.TEXT], "DAYS_ONMKT", targetComp.DOM });


                //commands.Add(9, new string[] { "1", inputType[elem.SELECT], "LOCATION_DESCR", "%" + GlobalVar.mainWindow.comboBoxLocationDescr.Text });
                //commands.Add(10, new string[] { "1", inputType[elem.SELECT], "OWNERSHIP_TYPE", "%Fee Simple" });
                //commands.Add(11, new string[] { "1", inputType[elem.TEXT], "LOTSIZE", targetComp.Lotsize.ToString() });
                //commands.Add(12, new string[] { "1", inputType[elem.SELECT], "LOTSIZE_TYPE_S", "%A" });

                //if (targetComp.Waterfront)
                //{
                //    commands.Add(13, new string[] { "1", inputType[elem.TEXT], "PROPERTY_VIEW", "Water" });
                //}
                //else
                //{
                //    commands.Add(13, new string[] { "1", inputType[elem.TEXT], "PROPERTY_VIEW", "Residential" });
                //}
                //commands.Add(14, new string[] { "1", inputType[elem.SELECT], "CONSTRUCTION_QUALITY", "%Average" });
                //commands.Add(15, new string[] { "1", inputType[elem.TEXT], "YEAR_BUILT", targetComp.YearBuiltString });

                //commands.Add(17, new string[] { "1", inputType[elem.TEXT], "UNIT_NUM_S", "1" });
                //commands.Add(18, new string[] { "1", inputType[elem.TEXT], "TOTAL_ROOMS_NUM", targetComp.TotalRoomCount.ToString() });
                //commands.Add(19, new string[] { "1", inputType[elem.TEXT], "BEDROOMS_NUM", targetComp.BedroomCount });
                //commands.Add(20, new string[] { "1", inputType[elem.TEXT], "BATHS_NUM", targetComp.BathroomCount });


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
                //    commands.Add(33, new string[] { "1", inputType[elem.TEXT], "DATEOFSALE", targetComp.SalesDate.ToShortDateString() });
                //}

                //fix for their bug
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:BPO_FORM  ATTR=NAME:*PROPERTY_TYPE_S_" + targetCompNumber + " CONTENT=%" + PropertyTypeTranslation[targetComp.ToString()].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));

                foreach (var c in commands)
                {
                    macro.AppendFormat("TAG POS={0} TYPE={1} FORM=ID:BPO_FORM ATTR=NAME:*{3}_{2} CONTENT={4}\r\n", c.Value[0], c.Value[1], c.Value[2], targetCompNumber, c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
                }



                #endregion
            }
      

           



            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 60);




          




        }
    }
}
