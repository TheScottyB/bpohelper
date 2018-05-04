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
            Dictionary<double, string[]> commands = new Dictionary<double, string[]>();
            Dictionary<string, string> translationTable = new Dictionary<string, string>();
            StringBuilder macro = new StringBuilder();

 //           'Complete Page
 //TAG POS=1 TYPE=HTML ATTR=* EXTRACT=HTM 
 //'Complete Page TEXT only
 //TAG POS=1 TYPE=HTML ATTR=* EXTRACT=TXT

            macro.AppendLine(@"TAG POS=1 TYPE=HTML ATTR=* EXTRACT=TXT");
            iim.iimPlayCode(macro.ToString(), 30);
            string reportType = iim.iimGetLastExtract();


            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");


            if (reportType.Contains(".FORMV135") || reportType.Contains("FORMV28"))
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

            }
            else if (reportType.Contains("FORMV159"))
            {
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
            else if (reportType.Contains("FORMV152"))
            {
                #region form152
                //I. General Conditions
                if (GlobalVar.theSubjectProperty.ToString().Contains("Attached"))
                {
                    commands.Add(0, new string[] { "2", inputType[elem.RADIO], "PROPERTYTYPE", "YES" });
                    commands.Add(0.01, new string[] { "1", inputType[elem.TEXT], "ZONING_DESC_1", @"Condo" });
                    //OWNER_USE_1
                    commands.Add(0.02, new string[] { "1", inputType[elem.SELECT], "OWNER_USE_1", "%Condo" });
                    //PROJECTED_USE_1
                    commands.Add(0.03, new string[] { "1", inputType[elem.SELECT], "PROJECTED_USE_1", "%Condo" });


                }
                else
                {
                    commands.Add(0, new string[] { "1", inputType[elem.RADIO], "PROPERTYTYPE", "YES" });
                    commands.Add(0.01, new string[] { "1", inputType[elem.TEXT], "ZONING_DESC_1", "SFR" });
                    //OWNER_USE_1
                    commands.Add(0.02, new string[] { "1", inputType[elem.SELECT], "OWNER_USE_1", "%SFR Det" });
                    //PROJECTED_USE_1
                    commands.Add(0.03, new string[] { "1", inputType[elem.SELECT], "PROJECTED_USE_1", "%SFR Det" });


                }

                commands.Add(0.1, new string[] { "1", inputType[elem.TEXT], "apn_s", GlobalVar.theSubjectProperty.ParcelID });
                commands.Add(0.2, new string[] { "1", inputType[elem.TEXT], "OWNER_LNAME", callingForm.SubjectOOR });
                //ZONING_SOURCE_CODE_1
                commands.Add(0.3, new string[] { "1", inputType[elem.RADIO], "ZONING_SOURCE_CODE_1", "YES" });
                //IS_NEW_CONSTRUCTION_1
                commands.Add(0.4, new string[] { "2", inputType[elem.RADIO], "IS_NEW_CONSTRUCTION_1", "YES" });
                //HAS_DISASTER_DAMAGE_1
                commands.Add(0.5, new string[] { "2", inputType[elem.RADIO], "HAS_DISASTER_DAMAGE_1", "YES" });
                //occupancy
                commands.Add(0.6, new string[] { "1", inputType[elem.RADIO], "occupancy", "YES" });
                //secure_1
                commands.Add(0.7, new string[] { "1", inputType[elem.RADIO], "secure_1", "YES" });
                //CONDITION_1
                commands.Add(0.8, new string[] { "1", inputType[elem.SELECT], "CONDITION_1", "%C3" });
                //CONDITIONCOMM
                commands.Add(0.9, new string[] { "1", inputType[elem.TEXTAREA], "CONDITIONCOMM", @"No adverse conditions noted during drive-by." });


                //HOA
                //hoa_1_has_hoa
                commands.Add(1, new string[] { "2", inputType[elem.RADIO], "hoa_1_has_hoa", "YES" });

                //II. Subject Sales & Listing History

                // III. Neighborhood & Market Information
                // 
                //GlobalVar.mainWindow.comboBoxLocationDescr.Text
                //location
                int x = 0;
                switch (GlobalVar.mainWindow.comboBoxLocationDescr.Text)
                {
                    case "Urban" :
                        x = 1;
                        break;
                    case "Suburban" :
                        x = 2;
                        break;
                    case "Rual":
                        x = 3;
                        break;
                }

                commands.Add(3, new string[] { x.ToString(), inputType[elem.RADIO], "location", "YES" });
                //LOCATION_INFLUENCE_CD
                commands.Add(3.01, new string[] { "1", inputType[elem.RADIO], "LOCATION_INFLUENCE_CD", "YES" });
                //LOCATION_FACTOR
                commands.Add(3.1, new string[] { "1", inputType[elem.SELECT], "LOCATION_FACTOR", "%Res" });
                //markettype
                commands.Add(3.2, new string[] { "2", inputType[elem.RADIO], "markettype", "YES" });
                //DOM_AVG_L
                commands.Add(3.201, new string[] { "1", inputType[elem.TEXT], "DOM_AVG_L", "90" });
                //MARKETTYPE_S
                commands.Add(3.21, new string[] { "2", inputType[elem.RADIO], "MARKETTYPE_S", "YES" });
                //REO_INVENTORY_TREND
                commands.Add(3.22, new string[] { "2", inputType[elem.RADIO], "REO_INVENTORY_TREND", "YES" });
                //PREDOM_OCCUPANCY
                commands.Add(3.23, new string[] { "1", inputType[elem.RADIO], "PREDOM_OCCUPANCY", "YES" });
                //RATE_PER_UNIT
                commands.Add(3.3, new string[] { "1", inputType[elem.TEXT], "RATE_PER_UNIT", callingForm.SubjectRent });
                //COMMERCIAL_USES
                commands.Add(3.4, new string[] { "2", inputType[elem.RADIO], "COMMERCIAL_USES", "YES" });
                //VACANCY_RATE
                commands.Add(3.5, new string[] { "2", inputType[elem.RADIO], "VACANCY_RATE", "YES" });
                //NEW_CONSTR_S
                commands.Add(3.6, new string[] { "2", inputType[elem.RADIO], "NEW_CONSTR_S", "YES" });
                //NEIGHBORHOOD_QUALITY_S
                commands.Add(3.7, new string[] { "1", inputType[elem.RADIO], "NEIGHBORHOOD_QUALITY_S", "YES" });
                //HAS_DISASTER_DAMAGE
                commands.Add(3.8, new string[] { "2", inputType[elem.RADIO], "HAS_DISASTER_DAMAGE", "YES" });
                //NUM_ACTV_LSTNGS
                commands.Add(3.9, new string[] { "1", inputType[elem.TEXT], "NUM_ACTV_LSTNGS", callingForm.cdNumberOfActiveListingTextBox.Text });
                //LIST_LOW_S
                commands.Add(3.901, new string[] { "1", inputType[elem.TEXT], "LIST_LOW_S", callingForm.SubjectNeighborhood.minListPrice.ToString() });
                //LIST_HIGH_S
                commands.Add(3.902, new string[] { "1", inputType[elem.TEXT], "LIST_HIGH_S", callingForm.SubjectNeighborhood.maxListPrice.ToString() });
                //NUM_COMP_SALES - cdNumberOfSoldListingTextBox
                commands.Add(3.903, new string[] { "1", inputType[elem.TEXT], "NUM_COMP_SALES", callingForm.cdNumberOfSoldListingTextBox.Text });
                //COMP_SALES_LOW_PRICE
                commands.Add(3.904, new string[] { "1", inputType[elem.TEXT], "COMP_SALES_LOW_PRICE", callingForm.SubjectNeighborhood.minSalePrice.ToString() });
                //COMP_SALES_HIGH_PRICE
                commands.Add(3.905, new string[] { "1", inputType[elem.TEXT], "COMP_SALES_HIGH_PRICE", callingForm.SubjectNeighborhood.maxSalePrice.ToString() });
                //neighborhoodcomm
                commands.Add(3.906, new string[] { "1", inputType[elem.TEXTAREA], "neighborhoodcomm", @"Small stand-alone subdivsion of multi-acre lots situated in rural area north of the town of Woodstock. " });
                //CONDITIONCOMM_S
                commands.Add(3.907, new string[] { "1", inputType[elem.TEXTAREA], "CONDITIONCOMM_S", @"Strong market at or under 225k.  Normal between 250-300. Weak above 300k. " });
                //COMMENT_CONCERN
                commands.Add(3.908, new string[] { "1", inputType[elem.TEXTAREA], "COMMENT_CONCERN", @"No red flags noted during drive-by inspection." });

                // subject data on comp page
                //DATASOURCE_S_1
                commands.Add(4, new string[] { "1", inputType[elem.SELECT], "DATASOURCE_S_1", "%Tax Records" });
                //UNIT_NUM_S_1
                commands.Add(4.1, new string[] { "1", inputType[elem.TEXT], "UNIT_NUM_S_1", "1" });
                //HOME_STYLE_S_1
                commands.Add(4.2, new string[] { "1", inputType[elem.SELECT], "HOME_STYLE_S_1", "%Conv" });
                //ATTACHMENT_TYPE_CD_1
                commands.Add(4.3, new string[] { "1", inputType[elem.SELECT], "ATTACHMENT_TYPE_CD_1", "%DT" });
                //LEVEL_S_1
                commands.Add(4.4, new string[] { "1", inputType[elem.TEXT], "LEVEL_S_1", "2" });
                //HOA_1
                commands.Add(4.5, new string[] { "1", inputType[elem.TEXT], "HOA_1", "0" });
                //Age_1
                commands.Add(4.6, new string[] { "1", inputType[elem.TEXT], "Age_1", GlobalVar.theSubjectProperty.Age.ToString() });
                //EXT_MATERIAL_CD_1
                commands.Add(4.7, new string[] { "1", inputType[elem.SELECT], "EXT_MATERIAL_CD_1", "%Frame" });
                //TOTAL_ROOMS_NUM_1
                commands.Add(4.8, new string[] { "1", inputType[elem.TEXT], "TOTAL_ROOMS_NUM_1", callingForm.SubjectRoomCount});
                //BEDROOMS_NUM_1
                commands.Add(4.9, new string[] { "1", inputType[elem.TEXT], "BEDROOMS_NUM_1", callingForm.SubjectBedroomCount});
                //BATHS_NUM_1
                commands.Add(4.101, new string[] { "1", inputType[elem.TEXT], "BATHS_NUM_1", GlobalVar.theSubjectProperty.FullBathCount});
                //HALF_BATHS_NUM_1
                commands.Add(4.11, new string[] { "1", inputType[elem.TEXT], "HALF_BATHS_NUM_1", GlobalVar.theSubjectProperty.HalfBathCount});
                //living_sqfeet_1
                commands.Add(4.12, new string[] { "1", inputType[elem.TEXT], "living_sqfeet_1", callingForm.SubjectAboveGLA });
                //GARAGE_1
                if (callingForm.SubjectParkingType.Contains("Attached"))
                {
                    commands.Add(4.13, new string[] { "1", inputType[elem.SELECT], "GARAGE_1", "%GA" });
                }
                else if (callingForm.SubjectParkingType.Contains("Detached"))
                {
                    commands.Add(4.13, new string[] { "1", inputType[elem.SELECT], "GARAGE_1", "%GD" });
                }
                else
                {
                    commands.Add(4.13, new string[] { "1", inputType[elem.SELECT], "GARAGE_1", "%None" });
                }
                //GARAGE_CAR_NUM_L_1
                commands.Add(4.14, new string[] { "1", inputType[elem.TEXT], "GARAGE_CAR_NUM_L_1", GlobalVar.theSubjectProperty.GarageStallCount });
                //LOTSIZE_1
                commands.Add(4.15, new string[] { "1", inputType[elem.TEXT], "LOTSIZE_1", callingForm.SubjectLotSize });
                //LOTSIZE_TYPE_S_1
                commands.Add(4.16, new string[] { "1", inputType[elem.SELECT], "LOTSIZE_TYPE_S_1", "%A" });
                //LOCATION_DESCR_1
                commands.Add(4.17, new string[] { "1", inputType[elem.SELECT], "LOCATION_DESCR_1", "%Average" });
                //BASEMENT_TYPE_CD_1
                commands.Add(4.18, new string[] { "1", inputType[elem.SELECT], "BASEMENT_TYPE_CD_1", "%IN" });
                //BSMT_PERCENT_S_1
                commands.Add(4.19, new string[] { "1", inputType[elem.TEXT], "BSMT_PERCENT_S_1", GlobalVar.theSubjectProperty.BasementFinishedPercentage.ToString() });
                //BSMT_SQFEET_S_1
                commands.Add(4.21, new string[] { "1", inputType[elem.TEXT], "BSMT_SQFEET_S_1", callingForm.SubjectBasementGLA });
                //PROPERTY_VIEW_INFLUENCE_CD_1
                commands.Add(4.22, new string[] { "1", inputType[elem.SELECT], "PROPERTY_VIEW_INFLUENCE_CD_1", "%N" });
                //PROPERTY_VIEW_1
                commands.Add(4.23, new string[] { "1", inputType[elem.SELECT], "PROPERTY_VIEW_1", "%Res" });
                //QUICK_AS_IS_VALUE
                commands.Add(4.24, new string[] { "1", inputType[elem.TEXT], "QUICK_AS_IS_VALUE", callingForm.SubjectQuickSaleValue });
                //QUICK_SALE_VALUE
                commands.Add(4.25, new string[] { "1", inputType[elem.TEXT], "QUICK_SALE_VALUE", callingForm.SubjectQuickSaleValue });
                //AS_IS_90_LIST_VALUE
                commands.Add(4.26, new string[] { "1", inputType[elem.TEXT], "AS_IS_90_LIST_VALUE", callingForm.SubjectMarketValueList });
                //AS_IS_90_SALE_VALUE
                commands.Add(4.27, new string[] { "1", inputType[elem.TEXT], "AS_IS_90_SALE_VALUE", callingForm.SubjectMarketValue });
                //REPAIRED_90_LIST_VALUE
                commands.Add(4.28, new string[] { "1", inputType[elem.TEXT], "REPAIRED_90_LIST_VALUE", callingForm.SubjectMarketValueList });
                //REPAIRED_90_SALE_VALUE
                commands.Add(4.29, new string[] { "1", inputType[elem.TEXT], "REPAIRED_90_SALE_VALUE", callingForm.SubjectMarketValue });
                //AS_IS_VALUE
                commands.Add(4.31, new string[] { "1", inputType[elem.TEXT], "AS_IS_VALUE", callingForm.SubjectMarketValueList });
                //SALE_VALUE
                commands.Add(4.32, new string[] { "1", inputType[elem.TEXT], "SALE_VALUE", callingForm.SubjectMarketValueList });
                //REPAIRED_VALUE
                commands.Add(4.33, new string[] { "1", inputType[elem.TEXT], "REPAIRED_VALUE", callingForm.SubjectMarketValueList });
                //REPAIRED_SALE_VALUE
                commands.Add(4.34, new string[] { "1", inputType[elem.TEXT], "REPAIRED_SALE_VALUE", callingForm.SubjectMarketValueList });
                //VALUE_RENT_S
                commands.Add(4.35, new string[] { "1", inputType[elem.TEXT], "VALUE_RENT_S", callingForm.SubjectRent });
                //pricingstrategycomm
                commands.Add(4.36, new string[] { "1", inputType[elem.TEXTAREA], "pricingstrategycomm", @"Searched a distance of at least 3 miles, up to 12 months in time. The comps bracket the subject in age, SF, and lot size, as well as used comps in with similar features and from the subjects market area. All the comps are Reasonable substitute for the subject property, similar in most areas. Price opinion was based off the average and median statistics gathered from the comparables and immediate neighborhood." });

                foreach (var c in commands)
                {
                    macro.AppendFormat("TAG POS={0} TYPE={1} FORM=ID:BPO_FORM ATTR=NAME:{2} CONTENT={3}\r\n", c.Value[0], c.Value[1], c.Value[2], c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
                }

                #endregion
            }
            else if (reportType.Contains("FORMV332 "))
            {
                #region FORMV332
                //I. General Conditions
                if (GlobalVar.theSubjectProperty.ToString().Contains("Attached"))
                {
                    //commands.Add(0, new string[] { "2", inputType[elem.RADIO], "PROPERTYTYPE", "YES" });
                    //commands.Add(0.01, new string[] { "1", inputType[elem.TEXT], "ZONING_DESC_1", @"Condo" });
                    //OWNER_USE_1
                    //commands.Add(0.02, new string[] { "1", inputType[elem.SELECT], "OWNER_USE_1", "%Condo" });
                    //PROJECTED_USE_1
                    //commands.Add(0.03, new string[] { "1", inputType[elem.SELECT], "PROJECTED_USE_1", "%Condo" });
                }
                else
                {
                    commands.Add(0, new string[] { "1", inputType[elem.SELECT], "PROPERTYTYPE", "%18" });
                    commands.Add(0.01, new string[] { "1", inputType[elem.TEXT], "ZONING", "%1" });
                    commands.Add(0.02, new string[] { "1", inputType[elem.SELECT], "SUBJECT_USE_CODE_1", "%L" });
                }

                commands.Add(0.1, new string[] { "1", inputType[elem.TEXT], "apn_s", GlobalVar.theSubjectProperty.ParcelID });
                //CONDITION_1
                commands.Add(0.8, new string[] { "1", inputType[elem.SELECT], "CONDITION_1", "%C3" });
                //IS_KITCHEN_BATH_REMODELED_1 = no
                commands.Add(0.3, new string[] { "2", inputType[elem.RADIO], "IS_KITCHEN_BATH_REMODELED_1", "YES" });
                //visibility_1
                commands.Add(0.4, new string[] { "1", inputType[elem.RADIO], "visibility_1", "YES" });
                //occupancy
                commands.Add(0.6, new string[] { "1", inputType[elem.RADIO], "occupancy_2", "YES" });
                //SECURE_COMMENTS
                commands.Add(0.9, new string[] { "1", inputType[elem.TEXTAREA], "SECURE_COMMENTS", @"Occupancy cannot be determined from the street as subject sits quite far back on the lot.  Lot appears maintained, however, the mail box is broken." });
                //CONDITIONCOMM
                commands.Add(0.91, new string[] { "1", inputType[elem.TEXTAREA], "legal_desc_1", @"DOC 2011R0006716 LT 3 /EX PC N SIDE/ ROBERT BARTLETTS ALTENBURG ACRES" });               
                //CONDITIONCOMM
                commands.Add(0.92, new string[] { "1", inputType[elem.TEXTAREA], "CONDITIONCOMM", @"The subject is a conforming home within the neighborhood. No adverse conditions were noted at the time of inspection based on exterior observations. Unable to determine interior condition due to exterior inspection only." });





             //   commands.Add(0.2, new string[] { "1", inputType[elem.TEXT], "OWNER_LNAME", callingForm.SubjectOOR });

                //HAS_DISASTER_DAMAGE_1
               // commands.Add(0.5, new string[] { "2", inputType[elem.RADIO], "HAS_DISASTER_DAMAGE_1", "YES" });
               
                //secure_1
             //   commands.Add(0.7, new string[] { "1", inputType[elem.RADIO], "secure_1", "YES" });



                //HOA
                //hoa_1_has_hoa
                commands.Add(1, new string[] { "2", inputType[elem.RADIO], "hoa_1_has_hoa", "YES" });

                //II. Subject Sales & Listing History

                // III. Neighborhood & Market Information
                // 
                //GlobalVar.mainWindow.comboBoxLocationDescr.Text
                //location
                int x = 0;
                switch (GlobalVar.mainWindow.comboBoxLocationDescr.Text)
                {
                    case "Urban":
                        x = 1;
                        break;
                    case "Suburban":
                        x = 2;
                        break;
                    case "Rual":
                        x = 3;
                        break;
                }

                commands.Add(3, new string[] { x.ToString(), inputType[elem.RADIO], "location", "YES" });
                //LOCATION_INFLUENCE_CD
                commands.Add(3.01, new string[] { "1", inputType[elem.RADIO], "LOCATION_INFLUENCE_CD", "YES" });
                //LOCATION_FACTOR
                commands.Add(3.1, new string[] { "1", inputType[elem.SELECT], "LOCATION_FACTOR", "%Res" });
                //markettype
                commands.Add(3.2, new string[] { "2", inputType[elem.RADIO], "markettype", "YES" });
                //DOM_AVG_L
                commands.Add(3.201, new string[] { "1", inputType[elem.TEXT], "DOM_AVG_L", "90" });
                //MARKETTYPE_S
                commands.Add(3.21, new string[] { "2", inputType[elem.RADIO], "MARKETTYPE_S", "YES" });
                //REO_INVENTORY_TREND
                commands.Add(3.22, new string[] { "2", inputType[elem.RADIO], "REO_INVENTORY_TREND", "YES" });
                //PREDOM_OCCUPANCY
                commands.Add(3.23, new string[] { "1", inputType[elem.RADIO], "PREDOM_OCCUPANCY", "YES" });
                //RATE_PER_UNIT
                commands.Add(3.3, new string[] { "1", inputType[elem.TEXT], "RATE_PER_UNIT", callingForm.SubjectRent });
                //COMMERCIAL_USES
                commands.Add(3.4, new string[] { "2", inputType[elem.RADIO], "COMMERCIAL_USES", "YES" });
                //VACANCY_RATE
                commands.Add(3.5, new string[] { "2", inputType[elem.RADIO], "VACANCY_RATE", "YES" });
                //NEW_CONSTR_S
                commands.Add(3.6, new string[] { "2", inputType[elem.RADIO], "NEW_CONSTR_S", "YES" });
                //NEIGHBORHOOD_QUALITY_S
                commands.Add(3.7, new string[] { "1", inputType[elem.RADIO], "NEIGHBORHOOD_QUALITY_S", "YES" });
                //HAS_DISASTER_DAMAGE
                commands.Add(3.8, new string[] { "2", inputType[elem.RADIO], "HAS_DISASTER_DAMAGE", "YES" });
                //NUM_ACTV_LSTNGS
                commands.Add(3.9, new string[] { "1", inputType[elem.TEXT], "NUM_ACTV_LSTNGS", callingForm.cdNumberOfActiveListingTextBox.Text });
                //LIST_LOW_S
                commands.Add(3.901, new string[] { "1", inputType[elem.TEXT], "LIST_LOW_S", callingForm.SubjectNeighborhood.minListPrice.ToString() });
                //LIST_HIGH_S
                commands.Add(3.902, new string[] { "1", inputType[elem.TEXT], "LIST_HIGH_S", callingForm.SubjectNeighborhood.maxListPrice.ToString() });
                //NUM_COMP_SALES - cdNumberOfSoldListingTextBox
                commands.Add(3.903, new string[] { "1", inputType[elem.TEXT], "NUM_COMP_SALES", callingForm.cdNumberOfSoldListingTextBox.Text });
                //COMP_SALES_LOW_PRICE
                commands.Add(3.904, new string[] { "1", inputType[elem.TEXT], "COMP_SALES_LOW_PRICE", callingForm.SubjectNeighborhood.minSalePrice.ToString() });
                //COMP_SALES_HIGH_PRICE
                commands.Add(3.905, new string[] { "1", inputType[elem.TEXT], "COMP_SALES_HIGH_PRICE", callingForm.SubjectNeighborhood.maxSalePrice.ToString() });
                //neighborhoodcomm
                commands.Add(3.906, new string[] { "1", inputType[elem.TEXTAREA], "neighborhoodcomm", @"Small stand-alone subdivsion of multi-acre lots situated in rural area north of the town of Woodstock. " });
                //CONDITIONCOMM_S
                commands.Add(3.907, new string[] { "1", inputType[elem.TEXTAREA], "CONDITIONCOMM_S", @"Strong market at or under 225k.  Normal between 250-300. Weak above 300k. " });
                //COMMENT_CONCERN
                commands.Add(3.908, new string[] { "1", inputType[elem.TEXTAREA], "COMMENT_CONCERN", @"No red flags noted during drive-by inspection." });

                // subject data on comp page
                //DATASOURCE_S_1
                commands.Add(4, new string[] { "1", inputType[elem.SELECT], "DATASOURCE_S_1", "%Tax Records" });
                //UNIT_NUM_S_1
                commands.Add(4.1, new string[] { "1", inputType[elem.TEXT], "UNIT_NUM_S_1", "1" });
                //HOME_STYLE_S_1
                commands.Add(4.2, new string[] { "1", inputType[elem.SELECT], "HOME_STYLE_S_1", "%Conv" });
                //ATTACHMENT_TYPE_CD_1
                commands.Add(4.3, new string[] { "1", inputType[elem.SELECT], "ATTACHMENT_TYPE_CD_1", "%DT" });
                //LEVEL_S_1
                commands.Add(4.4, new string[] { "1", inputType[elem.TEXT], "LEVEL_S_1", "2" });
                //HOA_1
                commands.Add(4.5, new string[] { "1", inputType[elem.TEXT], "HOA_1", "0" });
                //Age_1
                commands.Add(4.6, new string[] { "1", inputType[elem.TEXT], "Age_1", GlobalVar.theSubjectProperty.Age.ToString() });
                //EXT_MATERIAL_CD_1
                commands.Add(4.7, new string[] { "1", inputType[elem.SELECT], "EXT_MATERIAL_CD_1", "%Frame" });
                //TOTAL_ROOMS_NUM_1
                commands.Add(4.8, new string[] { "1", inputType[elem.TEXT], "TOTAL_ROOMS_NUM_1", callingForm.SubjectRoomCount });
                //BEDROOMS_NUM_1
                commands.Add(4.9, new string[] { "1", inputType[elem.TEXT], "BEDROOMS_NUM_1", callingForm.SubjectBedroomCount });
                //BATHS_NUM_1
                commands.Add(4.101, new string[] { "1", inputType[elem.TEXT], "BATHS_NUM_1", GlobalVar.theSubjectProperty.FullBathCount });
                //HALF_BATHS_NUM_1
                commands.Add(4.11, new string[] { "1", inputType[elem.TEXT], "HALF_BATHS_NUM_1", GlobalVar.theSubjectProperty.HalfBathCount });
                //living_sqfeet_1
                commands.Add(4.12, new string[] { "1", inputType[elem.TEXT], "living_sqfeet_1", callingForm.SubjectAboveGLA });
                //GARAGE_1
                if (callingForm.SubjectParkingType.Contains("Attached"))
                {
                    commands.Add(4.13, new string[] { "1", inputType[elem.SELECT], "GARAGE_1", "%GA" });
                }
                else if (callingForm.SubjectParkingType.Contains("Detached"))
                {
                    commands.Add(4.13, new string[] { "1", inputType[elem.SELECT], "GARAGE_1", "%GD" });
                }
                else
                {
                    commands.Add(4.13, new string[] { "1", inputType[elem.SELECT], "GARAGE_1", "%None" });
                }
                //GARAGE_CAR_NUM_L_1
                commands.Add(4.14, new string[] { "1", inputType[elem.TEXT], "GARAGE_CAR_NUM_L_1", GlobalVar.theSubjectProperty.GarageStallCount });
                //LOTSIZE_1
                commands.Add(4.15, new string[] { "1", inputType[elem.TEXT], "LOTSIZE_1", callingForm.SubjectLotSize });
                //LOTSIZE_TYPE_S_1
                commands.Add(4.16, new string[] { "1", inputType[elem.SELECT], "LOTSIZE_TYPE_S_1", "%A" });
                //LOCATION_DESCR_1
                commands.Add(4.17, new string[] { "1", inputType[elem.SELECT], "LOCATION_DESCR_1", "%Average" });
                //BASEMENT_TYPE_CD_1
                commands.Add(4.18, new string[] { "1", inputType[elem.SELECT], "BASEMENT_TYPE_CD_1", "%IN" });
                //BSMT_PERCENT_S_1
                commands.Add(4.19, new string[] { "1", inputType[elem.TEXT], "BSMT_PERCENT_S_1", GlobalVar.theSubjectProperty.BasementFinishedPercentage.ToString() });
                //BSMT_SQFEET_S_1
                commands.Add(4.21, new string[] { "1", inputType[elem.TEXT], "BSMT_SQFEET_S_1", callingForm.SubjectBasementGLA });
                //PROPERTY_VIEW_INFLUENCE_CD_1
                commands.Add(4.22, new string[] { "1", inputType[elem.SELECT], "PROPERTY_VIEW_INFLUENCE_CD_1", "%N" });
                //PROPERTY_VIEW_1
                commands.Add(4.23, new string[] { "1", inputType[elem.SELECT], "PROPERTY_VIEW_1", "%Res" });
                //QUICK_AS_IS_VALUE
                commands.Add(4.24, new string[] { "1", inputType[elem.TEXT], "QUICK_AS_IS_VALUE", callingForm.SubjectQuickSaleValue });
                //QUICK_SALE_VALUE
                commands.Add(4.25, new string[] { "1", inputType[elem.TEXT], "QUICK_SALE_VALUE", callingForm.SubjectQuickSaleValue });
                //AS_IS_90_LIST_VALUE
                commands.Add(4.26, new string[] { "1", inputType[elem.TEXT], "AS_IS_90_LIST_VALUE", callingForm.SubjectMarketValueList });
                //AS_IS_90_SALE_VALUE
                commands.Add(4.27, new string[] { "1", inputType[elem.TEXT], "AS_IS_90_SALE_VALUE", callingForm.SubjectMarketValue });
                //REPAIRED_90_LIST_VALUE
                commands.Add(4.28, new string[] { "1", inputType[elem.TEXT], "REPAIRED_90_LIST_VALUE", callingForm.SubjectMarketValueList });
                //REPAIRED_90_SALE_VALUE
                commands.Add(4.29, new string[] { "1", inputType[elem.TEXT], "REPAIRED_90_SALE_VALUE", callingForm.SubjectMarketValue });
                //AS_IS_VALUE
                commands.Add(4.31, new string[] { "1", inputType[elem.TEXT], "AS_IS_VALUE", callingForm.SubjectMarketValueList });
                //SALE_VALUE
                commands.Add(4.32, new string[] { "1", inputType[elem.TEXT], "SALE_VALUE", callingForm.SubjectMarketValueList });
                //REPAIRED_VALUE
                commands.Add(4.33, new string[] { "1", inputType[elem.TEXT], "REPAIRED_VALUE", callingForm.SubjectMarketValueList });
                //REPAIRED_SALE_VALUE
                commands.Add(4.34, new string[] { "1", inputType[elem.TEXT], "REPAIRED_SALE_VALUE", callingForm.SubjectMarketValueList });
                //VALUE_RENT_S
                commands.Add(4.35, new string[] { "1", inputType[elem.TEXT], "VALUE_RENT_S", callingForm.SubjectRent });
                //pricingstrategycomm
                commands.Add(4.36, new string[] { "1", inputType[elem.TEXTAREA], "pricingstrategycomm", @"Searched a distance of at least 3 miles, up to 12 months in time. The comps bracket the subject in age, SF, and lot size, as well as used comps in with similar features and from the subjects market area. All the comps are Reasonable substitute for the subject property, similar in most areas. Price opinion was based off the average and median statistics gathered from the comparables and immediate neighborhood." });

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


            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");


            //updates 4/13/18
            //MOAB-WEB7.20180413.081716.FORMV337 
            if (reportType.Contains(".FORMV135") || reportType.Contains("FORMV28") || reportType.Contains("FORMV337"))
            {
                #region oldForm
                Dictionary<string, string> saleTypeTranslation = new Dictionary<string, string>()
                {
                    {"Arms Length", "19"}, {"REO", "1"}, {"ShortSale", "17"},
                };
                Dictionary<string, string> heatingTypeTranslation = new Dictionary<string, string>()
                {
                    {"Gas", "6"}, {"Gas, Forced Air", "4"}, {"Propane, Forced Air", "4"}, {"Gas, Forced Air, Zoned", "4"}, {"Gas, Forced Air, 2+ Sep Heating Systems", "4"}, {"Gas, Forced Air, Radiant, 2+ Sep Heating Systems", "4"}, {"Propane, Forced Air, Radiant, 2+ Sep Heating Systems, Indv Controls, Zoned", "4"}, {@"Hot Water/Steam", "9"}, {@"Gas, Hot Water/Steam", "9"},{@"Electric", "2"}, {@"Gas, Hot Water/Steam, Baseboard", "2"}, {@"Baseboard, Radiant", "2"}, {@"Baseboard", "2"}
                };

                Dictionary<string, string> CoolingTypeTranslation = new Dictionary<string, string>()
                {
                    {"Central Air", "1"},  {"Central Air, Zoned", "1"},  {"Central Air, Zoned, 2 Separate Systems", "1"}, {"None", "0"},  {@"1 (Window/Wall Unit)", "2"},  {"Partial", "0"}, {@"3+ (Window/Wall Unit)", "2"} //Partial
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
            else if (reportType.Contains("FORMV159"))
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
            else if (reportType.Contains("FORMV152") || reportType.Contains("FORMV332"))
            {
                #region form152

                Dictionary<double, string[]> commands = new Dictionary<double, string[]>();

                Dictionary<string, string> saleTypeTranslation = new Dictionary<string, string>()
            {
                {"Arms Length", "34"}, {"REO", "1"}, {"ShortSale", "17"},
            };
                Dictionary<string, string> heatingTypeTranslation = new Dictionary<string, string>()
            {
                {"Gas", "6"}, {"Gas, Forced Air", "6"}, {@"Gas, Hot Water/Steam", "6"},{@"Electric", "2"}, {"Forced Air", "4"}

            };

                Dictionary<string, string> CoolingTypeTranslation = new Dictionary<string, string>()
            {
                {"Central Air", "1"},  {"None", "0"}
            };

                Dictionary<string, string> PropertyTypeTranslation = new Dictionary<string, string>()
            {
                {"bpohelper.AttachedListing", "Condo"},  {"bpohelper.DetachedListing", "SFR"}
            };

                Dictionary<string, string> BuildingTypeTranslation = new Dictionary<string, string>()
            {
                {"1 Story", "Ranch"},  {"1.5 Story", "Conv"}, {"2 Stories", "Conv"}, {"Hillside", @"Split/Bi-level"}, {"Raised Ranch", @"Split/Bi-level"}, {@"Split Level w/Sub", @"Split/Bi-level"}, {"Split Level", @"Split/Bi-level"},
                    {"Condo", "Conv"}, {@"Townhouse-2 Story", "Townhouse"}, {@"Townhouse-Ranch", "Ranch"}  //Townhouse-Ranch
            };                                                             //Townhouse-Ranch


                Dictionary<string, string> aboveGradeLevelsTranslation = new Dictionary<string, string>()
            {
                {"1 Story", "1"},  {"1.5 Story", "2"}, {"2 Stories", "2"}, {"Hillside", @"1"}, {"Raised Ranch", @"1"}, {@"Split Level w/Sub", @"2"}, {"Split Level", @"2"},
                    {"Condo", "1"}, {"Townhouse-2 Story", "2"}, {"Townhouse-Ranch", "1"}
            };

                Dictionary<string, string> financingTypeTranslation = new Dictionary<string, string>()
            {
                {"Conventional", "%Conventional"}, {"FHA", "%Seller"}, {"VA", "%Seller"},  {"Unknown", "%Conventional"}, {"Cash", "%Cash"}

            };

                //
                //Updates 4/1/18
                //
                commands.Add(401.1, new string[] { "1", inputType[elem.TEXT], "MLS_BOARD", "Midwest Real Estate Data, Inc" });
                commands.Add(401.2, new string[] { "1", inputType[elem.TEXT], "LISTING_DT", targetComp.DateOfLastPriceChange.ToShortDateString() });
                commands.Add(401.3, new string[] { "1", inputType[elem.SELECT], "OWNERSHIP_TYPE", "%Fee Simple" });
                commands.Add(401.4, new string[] { "1", inputType[elem.SELECT], "LOCATION_INFLUENCE_CD", "%N" });
                commands.Add(401.5, new string[] { "1", inputType[elem.SELECT], "LOCATION_DESCR", "%Res" });
                commands.Add(401.6, new string[] { "1", inputType[elem.SELECT], "CONSTRUCTION_QUALITY", "%Q4" });
                commands.Add(401.7, new string[] { "1", inputType[elem.SELECT], "PROPERTY_TYPE_S", "%49" });
                commands.Add(401.8, new string[] { "1", inputType[elem.TEXT], "YEAR_BUILT", targetComp.YearBuiltString });
                commands.Add(401.9, new string[] { "2", inputType[elem.RADIO], "BMST_B", "YES" });
                if (targetComp.BasementType == "None")
                {
                    commands.Add(401.11, new string[] { "2", inputType[elem.RADIO], "BMST_B", "YES" });
                    //commands.Add(14.12, new string[] { "1", inputType[elem.TEXT], "BSMT_SQFEET_S", "0" });
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:BSMT_B_" + targetCompNumber + " CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:BSMT_SQFEET_S_" + targetCompNumber + " CONTENT=0");
                }
                else
                {
                    commands.Add(401.12, new string[] { "1", inputType[elem.RADIO], "BMST_B", "YES" });
                    //commands.Add(14.22, new string[] { "1", inputType[elem.TEXT], "BSMT_PERCENT_S", targetComp.BasementFinishedPercentage() });
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:BSMT_B_" + targetCompNumber + " CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:BSMT_SQFEET_S_" + targetCompNumber + " CONTENT=" + targetComp.BasementGLA());
                }
                commands.Add(401.13, new string[] { "1", inputType[elem.TEXT], "*UNIT_1_BEDS", targetComp.BedroomCount });
                commands.Add(401.14, new string[] { "1", inputType[elem.TEXT], "*UNIT_1_ROOMS", targetComp.TotalRoomCount.ToString() });
                commands.Add(401.15, new string[] { "1", inputType[elem.TEXT], "*UNIT_1_FULLBATHS", targetComp.FullBathCount });
                commands.Add(401.16, new string[] { "1", inputType[elem.TEXT], "*UNIT_1_HALFBATHS", targetComp.HalfBathCount });
                commands.Add(401.17, new string[] { "1", inputType[elem.SELECT], "*HEATING", "%" + heatingTypeTranslation[targetComp.Heating] });
                commands.Add(401.18, new string[] { "1", inputType[elem.SELECT], "*COOLING", "%" + CoolingTypeTranslation[targetComp.Cooling] });
                if (targetComp.GarageExsists())
                {
                    commands.Add(401.19, new string[] { "1", inputType[elem.SELECT], "GARAGE_CAR_NUM_L", "%" + targetComp.NumberGarageStalls()  });
                }
                else
                {
                    commands.Add(401.21, new string[] { "1", inputType[elem.SELECT], "GARAGE_CAR_NUM_L", "%0" });
                }


                //
                //
                //



                commands.Add(0, new string[] { "1", inputType[elem.TEXT], "ADDRESS", targetComp.StreetAddress });
                commands.Add(1, new string[] { "1", inputType[elem.TEXT], "ZIP", targetComp.Zipcode });
                commands.Add(2, new string[] { "1", inputType[elem.SELECT], "DATASOURCE_S", "%MLS" });              
                commands.Add(3, new string[] { "1", inputType[elem.TEXT], "MLS_NUMBER", targetComp.MlsNumber });
                commands.Add(4, new string[] { "1", inputType[elem.SELECT], "PROPERTY_TYPE_S", "%" + PropertyTypeTranslation[targetComp.ToString()] });
                commands.Add(5, new string[] { "1", inputType[elem.TEXT], "UNIT_NUM_S", "1" });
                commands.Add(6, new string[] { "1", inputType[elem.SELECT], "SALES_TYPE", "%" + saleTypeTranslation[targetComp.TransactionType] });
                commands.Add(7, new string[] { "1", inputType[elem.SELECT], "HOME_STYLE_S", "%" + BuildingTypeTranslation[targetComp.Type] });
                if (GlobalVar.theSubjectProperty.ToString().Contains("Attached"))
                {
                    commands.Add(8, new string[] { "1", inputType[elem.SELECT], "ATTACHMENT_TYPE_CD", "%AT" });
                }
                else
                {
                    commands.Add(8, new string[] { "1", inputType[elem.SELECT], "ATTACHMENT_TYPE_CD", "%DT"});
                }
                commands.Add(9, new string[] { "1", inputType[elem.TEXT], "LEVEL_S", aboveGradeLevelsTranslation[targetComp.Type] });
                commands.Add(10, new string[] { "1", inputType[elem.TEXT], "ORIGLIST_DT", targetComp.ListDateString });
                commands.Add(11, new string[] { "1", inputType[elem.TEXT], "ORIGLIST_PRICE", targetComp.OriginalListPrice.ToString() });
                commands.Add(12, new string[] { "1", inputType[elem.TEXT], "CURRENTLIST_DT", targetComp.DateOfLastPriceChange.ToShortDateString() });
                commands.Add(13, new string[] { "1", inputType[elem.TEXT], "currentlist_price", targetComp.CurrentListPrice.ToString() });
                commands.Add(14, new string[] { "1", inputType[elem.TEXT], "SELLER_CONCESSIONS", targetComp.PointsInDollars });
                //TODO:
                //Find the actual hoa in the listings
                commands.Add(15, new string[] { "1", inputType[elem.TEXT], "HOA", "0" });
                if (saleOrList == "sale")
                {
                    commands.Add(15.1, new string[] { "1", inputType[elem.TEXT], "SALE_PRICE", targetComp.SalePrice.ToString() });
                    commands.Add(15.2, new string[] { "1", inputType[elem.TEXT], "dateofsale", targetComp.SalesDate.ToShortDateString() });
                }
                commands.Add(16, new string[] { "1", inputType[elem.TEXT], "days_onmkt", targetComp.DOM });
                commands.Add(17, new string[] { "1", inputType[elem.SELECT], "TYPE_FINANCE", financingTypeTranslation[targetComp.FinancingMlsString] });
                commands.Add(18, new string[] { "1", inputType[elem.TEXT], "AGE", targetComp.Age.ToString() });
                //TODO:
                //Default is C3, find a way to auto select the proper condition.
                commands.Add(19, new string[] { "1", inputType[elem.SELECT], "CONDITION", "%C3" });
                //TODO:
                //Default is frame, find a way to auto select the proper type.
                commands.Add(20, new string[] { "1", inputType[elem.SELECT], "EXT_MATERIAL_CD", "%Frame" });
                commands.Add(21, new string[] { "1", inputType[elem.TEXT], "TOTAL_ROOMS_NUM", targetComp.TotalRoomCount.ToString() });
                commands.Add(22, new string[] { "1", inputType[elem.TEXT], "BEDROOMS_NUM", targetComp.BedroomCount.ToString() });
                commands.Add(23, new string[] { "1", inputType[elem.TEXT], "BATHS_NUM", targetComp.FullBathCount.ToString() });
                commands.Add(24, new string[] { "1", inputType[elem.TEXT], "HALF_BATHS_NUM", targetComp.HalfBathCount.ToString() });
                commands.Add(25, new string[] { "1", inputType[elem.TEXT], "living_sqfeet", targetComp.ProperGla(GlobalVar.mainWindow.SubjectAboveGLA) });
                if (targetComp.AttachedGarage())
                {
                    commands.Add(26, new string[] { "1", inputType[elem.SELECT], "GARAGE*", "%GA" });
                }
                else if (targetComp.DetachedGarage())
                {
                    commands.Add(26, new string[] { "1", inputType[elem.SELECT], "GARAGE*", "%GD" });
                } 
                else
                {
                    commands.Add(26, new string[] { "1", inputType[elem.SELECT], "GARAGE*", "%DW" });
                }
                commands.Add(27, new string[] { "1", inputType[elem.TEXT], "GARAGE_CAR_NUM_L", targetComp.NumberGarageStalls() });
                commands.Add(28, new string[] { "1", inputType[elem.TEXT], "LOTSIZE", targetComp.Lotsize.ToString() });
                commands.Add(29, new string[] { "1", inputType[elem.SELECT], "LOTSIZE_TYPE_S", "%A" });
                commands.Add(30, new string[] { "1", inputType[elem.SELECT], "LOCATION_DESCR", "%Average" });
                //TODO:  Figureout if Basement is walkout or not.
                commands.Add(31, new string[] { "1", inputType[elem.SELECT], "BASEMENT_TYPE_CD", "%IN" });
                commands.Add(32, new string[] { "1", inputType[elem.TEXT], "BSMT_PERCENT_S", targetComp.BasementFinishedPercentage()});
                commands.Add(33, new string[] { "1", inputType[elem.TEXT], "BSMT_SQFEET_S", targetComp.BasementGLA() });
                //TODO: View types, adjustment tables. and heirarchy 
                commands.Add(34, new string[] { "1", inputType[elem.SELECT], "PROPERTY_VIEW_INFLUENCE_CD", "%N" });
                //PROPERTY_VIEW
                commands.Add(34.1, new string[] { "1", inputType[elem.SELECT], "PROPERTY_VIEW", "%Res" });
                commands.Add(35, new string[] { "1", inputType[elem.SELECT], "PROPERTY_VIEW_COMPARE_CD", "%SM" });
                //SUP_INF_EQ_S_7
                //TODO: adjustment tables and equalities 
                commands.Add(36, new string[] { "1", inputType[elem.SELECT], "SUP_INF_EQ_S", "%eq" });
                //ListingComm_6
                commands.Add(37, new string[] { "1", inputType[elem.TEXTAREA], "ListingComm", targetComp.mlsHtmlFields["remarks"].value });











            








              
          
              
               
            
               
                ////      if (string.IsNullOrWhiteSpace(targetComp.BasementGLA()) || targetComp.BasementGLA() == "-1" || targetComp.BasementGLA() == "0")
                
                ////EXT_SIDING_S
                //// commands.Add(15, new string[] { "1", inputType[elem.SELECT], "EXT_SIDING_S", "%" + saleTypeTranslation[targetComp.TransactionType] });
                ////PORCH
                ////commands.Add(16, new string[] { "1", inputType[elem.TEXT], "PORCH", targetComp.BasementGLA() });
               
                //else
                //{
                //    commands.Add(19.2, new string[] { "1", inputType[elem.SELECT], "GARAGE", "%None" });
                //}






                ////commands.Add(4, new string[] { "1", inputType[elem.TEXT], "PRICE_REDUCTION_COUNT", targetComp.NumberOfPriceChanges.ToString() });




                ////commands.Add(8, new string[] { "1", inputType[elem.TEXT], "DAYS_ONMKT", targetComp.DOM });


                ////commands.Add(9, new string[] { "1", inputType[elem.SELECT], "LOCATION_DESCR", "%" + GlobalVar.mainWindow.comboBoxLocationDescr.Text });
                ////commands.Add(10, new string[] { "1", inputType[elem.SELECT], "OWNERSHIP_TYPE", "%Fee Simple" });
                ////commands.Add(11, new string[] { "1", inputType[elem.TEXT], "LOTSIZE", targetComp.Lotsize.ToString() });
                ////commands.Add(12, new string[] { "1", inputType[elem.SELECT], "LOTSIZE_TYPE_S", "%A" });

                ////if (targetComp.Waterfront)
                ////{
                ////    commands.Add(13, new string[] { "1", inputType[elem.TEXT], "PROPERTY_VIEW", "Water" });
                ////}
                ////else
                ////{
                ////    commands.Add(13, new string[] { "1", inputType[elem.TEXT], "PROPERTY_VIEW", "Residential" });
                ////}
                ////commands.Add(14, new string[] { "1", inputType[elem.SELECT], "CONSTRUCTION_QUALITY", "%Average" });
                ////commands.Add(15, new string[] { "1", inputType[elem.TEXT], "YEAR_BUILT", targetComp.YearBuiltString });

                ////commands.Add(18, new string[] { "1", inputType[elem.TEXT], "TOTAL_ROOMS_NUM", targetComp.TotalRoomCount.ToString() });
                ////commands.Add(19, new string[] { "1", inputType[elem.TEXT], "BEDROOMS_NUM", targetComp.BedroomCount });
                ////commands.Add(20, new string[] { "1", inputType[elem.TEXT], "BATHS_NUM", targetComp.BathroomCount });


                ////commands.Add(22, new string[] { "1", inputType[elem.TEXT], "BSMT_SQFEET_S", targetComp.BasementGLA() });
                ////commands.Add(23, new string[] { "1", inputType[elem.TEXT], "ROOMS_BELOW_GRADE", targetComp.NumberOfBasementRooms() });


            

       

                ////commands.Add(29, new string[] { "1", inputType[elem.TEXT], "OTHER", "NA" });
                ////commands.Add(30, new string[] { "2", inputType[elem.RADIO], "POOL_B", "YES" });
                ////if (saleOrList == "sale")
                ////{
                ////    commands.Add(31, new string[] { "1", inputType[elem.TEXT], "SALE_PRICE", targetComp.SalePrice.ToString() });
                ////    commands.Add(32, new string[] { "1", inputType[elem.TEXT], "CURRENTLIST_PRICE", targetComp.CurrentListPrice.ToString() });
                ////    commands.Add(33, new string[] { "1", inputType[elem.TEXT], "DATEOFSALE", targetComp.SalesDate.ToShortDateString() });
                ////}

                //fix for their bug
               // macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:BPO_FORM  ATTR=NAME:*PROPERTY_TYPE_S_" + targetCompNumber + " CONTENT=%" + PropertyTypeTranslation[targetComp.ToString()].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));

                foreach (var c in commands)
                {
                    macro.AppendFormat("TAG POS={0} TYPE={1} FORM=ID:BPO_FORM ATTR=NAME:{2}_{3} CONTENT={4}\r\n", c.Value[0], c.Value[1], c.Value[2], targetCompNumber, c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));         
                }


             
                ////PROPERTY_VIEW_INFLUENCE_CD_1
                ////PROPERTY_VIEW_1
                //commands.Add(4.23, new string[] { "1", inputType[elem.SELECT], "PROPERTY_VIEW_1", "%Res" });
                ////QUICK_AS_IS_VALUE
                //commands.Add(4.24, new string[] { "1", inputType[elem.TEXT], "QUICK_AS_IS_VALUE", callingForm.SubjectQuickSaleValue });
                ////QUICK_SALE_VALUE
                //commands.Add(4.25, new string[] { "1", inputType[elem.TEXT], "QUICK_SALE_VALUE", callingForm.SubjectQuickSaleValue });
                ////AS_IS_90_LIST_VALUE
                //commands.Add(4.26, new string[] { "1", inputType[elem.TEXT], "AS_IS_90_LIST_VALUE", callingForm.SubjectMarketValueList });
                ////AS_IS_90_SALE_VALUE
                //commands.Add(4.27, new string[] { "1", inputType[elem.TEXT], "AS_IS_90_SALE_VALUE", callingForm.SubjectMarketValue });
                ////REPAIRED_90_LIST_VALUE
                //commands.Add(4.28, new string[] { "1", inputType[elem.TEXT], "REPAIRED_90_LIST_VALUE", callingForm.SubjectMarketValueList });
                ////REPAIRED_90_SALE_VALUE
                //commands.Add(4.29, new string[] { "1", inputType[elem.TEXT], "REPAIRED_90_SALE_VALUE", callingForm.SubjectMarketValue });
                ////AS_IS_VALUE
                //commands.Add(4.31, new string[] { "1", inputType[elem.TEXT], "AS_IS_VALUE", callingForm.SubjectMarketValueList });
                ////SALE_VALUE
                //commands.Add(4.32, new string[] { "1", inputType[elem.TEXT], "SALE_VALUE", callingForm.SubjectMarketValueList });
                ////REPAIRED_VALUE
                //commands.Add(4.33, new string[] { "1", inputType[elem.TEXT], "REPAIRED_VALUE", callingForm.SubjectMarketValueList });
                ////REPAIRED_SALE_VALUE
                //commands.Add(4.34, new string[] { "1", inputType[elem.TEXT], "REPAIRED_SALE_VALUE", callingForm.SubjectMarketValueList });
                ////VALUE_RENT_S
                //commands.Add(4.35, new string[] { "1", inputType[elem.TEXT], "VALUE_RENT_S", callingForm.SubjectRent });
                ////pricingstrategycomm
                //commands.Add(4.36, new string[] { "1", inputType[elem.TEXTAREA], "pricingstrategycomm", @"Searched a distance of at least 3 miles, up to 12 months in time. The comps bracket the subject in age, SF, and lot size, as well as used comps in with similar features and from the subjects market area. All the comps are Reasonable substitute for the subject property, similar in most areas. Price opinion was based off the average and median statistics gathered from the comparables and immediate neighborhood." });

                //foreach (var c in commands)
                //{
                //    macro.AppendFormat("TAG POS={0} TYPE={1} FORM=ID:BPO_FORM ATTR=NAME:{2} CONTENT={3}\r\n", c.Value[0], c.Value[1], c.Value[2], c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
                //}

                #endregion
            }

           



            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 60);




          




        }
    }
}
