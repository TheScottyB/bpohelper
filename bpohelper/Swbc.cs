using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{

    class SWBC : AMO
    {
        protected Dictionary<string, string> typeTranslation = new Dictionary<string, string>()
        {
          {"1 Story", "1 Story"}, {"Raised Ranch", "1 Story"},{"Coach House", "1 Story"},{"Hillside", "1 Story"},{"2 Stories", "2 Story"},{"1.5 Story", "1.5 Story"},{"Split Level", "Multilevel"},{ @"Split Level w/Sub", "Multilevel"}

        };

        protected Dictionary<string, string> garageTypeTranslation = new Dictionary<string, string>()
            {
                {"0", "None"}, {"1", "1 Car"},  {"None", "None"},  {"2", "2 Car"}, {"3", "3 Car"}, {"4", "4 Car"}, {"", "None"}
            };

        protected Dictionary<string, string> bpoTypeTranslation = new Dictionary<string, string>()
            {
               {"Exterior", "Exterior BPO"}, {"Interior", "Interior BPO"}, {"", ""}
            };

        public SWBC()
            : base()    {   }

        public SWBC(MLSListing m)
            : base(m) { }

        private string _subjectCurrentlyListed(bool b)
        {
            if (_isSubjectListed() && b == true || !_isSubjectListed() && b == false)
            {
                return "YES";
            }
            return "NO";
        }

        private string _reoYesNoFill(string saleType)
        {
            if (saleType == "REO")
            {
                return "%Yes";
            }

            return "%No";
        }
     
        public void Prefill(iMacros.App iim, Form1 form)
        {
            Dictionary<double, string[]> commands = new Dictionary<double, string[]>();

            Dictionary<string, string> translationTable = new Dictionary<string, string>();

            

            StringBuilder macro = new StringBuilder();

            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");


           // commands.Add(0, new string[] { "1", inputType[elem.RADIO], "CMA_TYPE_EXTERIOR", "YES" });




            //1: order info
            commands.Add(1.0, new string[] { "1", inputType[elem.TEXT], "CMA_INSPECTION_DATE", DateTime.Now.ToShortDateString() });
            
            commands.Add(1.1, new string[] { "1", inputType[elem.SELECT], "VALUATION_TYPE", "%" + bpoTypeTranslation[callingForm.comboBoxBpoType.Text] });
            commands.Add(1.2, new string[] { "1", inputType[elem.TEXT], "BROKER_FULL_NAME", "Scott Beilfuss" });
            commands.Add(1.3, new string[] { "1", inputType[elem.TEXT], "BROKER_COMPANY", "OKRP" });

            //2: subj info
            commands.Add(2.0, new string[] { "1", inputType[elem.TEXT], "SUBJ_ASSESS_PARCEL", callingForm.SubjectPin});
            commands.Add(2.1, new string[] { "1", inputType[elem.SELECT], "MRKT_SUBJ_TYPE", "%Single<SP>Family" });
            commands.Add(2.2, new string[] { "1", inputType[elem.CHECKBOX], "DATA_SOURCE_ASSESSOR", "YES" });
            //MRKT_ASSOC_FEE
            commands.Add(2.21, new string[] { "1", inputType[elem.TEXT], "MRKT_ASSOC_FEE",callingForm.subjectHoaTextBox.Text });
            commands.Add(2.3, new string[] { "1", inputType[elem.CHECKBOX], "CUR_LISTED_YES", _subjectCurrentlyListed(true)});
            commands.Add(2.4, new string[] { "1", inputType[elem.CHECKBOX], "CUR_LISTED_NO", _subjectCurrentlyListed(false) });

            //if subj listed
            if (_isSubjectListed())
            {
                commands.Add(2.5, new string[] { "1", inputType[elem.TEXT], "LIST_DATE", callingForm.subjectListDatedateTimePicker.Text});
                commands.Add(2.6, new string[] { "1", inputType[elem.TEXT], "SUBJ_CURRENT_LIST", callingForm.SubjectCurrentListPrice });
                commands.Add(2.7, new string[] { "1", inputType[elem.TEXT], "SUBJ_ORIG_LIST", callingForm.textBoxSubjectOriginalListPRice.Text });
                commands.Add(2.8, new string[] { "1", inputType[elem.TEXT], "DOM", callingForm.SubjectDom });
                commands.Add(2.9, new string[] { "1", inputType[elem.TEXT], "SUBJ_MLS", callingForm.labelMlsNumber.Text  });
                commands.Add(2.11, new string[] { "1", inputType[elem.TEXT], "MRKT_LISTING_BROKER", callingForm.SubjectListingAgent });
                commands.Add(2.12, new string[] { "1", inputType[elem.TEXT], "MRKT_LISTING_COMPANY", callingForm.textBoxSubjextListingBrokerage.Text });
                commands.Add(2.13, new string[] { "1", inputType[elem.TEXT], "MRKT_LISTING_PHONE", callingForm.SubjectBrokerPhone });
            }
            commands.Add(2.14, new string[] { "1", inputType[elem.SELECT], "SUBJ_OCCUPANCY", "%" + callingForm.comboBoxSubjectOccupancy.Text });

            //3: Values
            //normal as-is sale MRKT_NORM_ASIS
            commands.Add(3.0, new string[] { "1", inputType[elem.TEXT], "MRKT_NORM_ASIS", callingForm.SubjectMarketValue });
            //normal as-is list
            commands.Add(3.1, new string[] { "1", inputType[elem.TEXT], "MRKT_NORM_ASIS_REC", callingForm.SubjectMarketValueList });
            //normal repaired sale
            commands.Add(3.2, new string[] { "1", inputType[elem.TEXT], "MRKT_NORM_ASIS_REPAIRED", callingForm.SubjectMarketValue });
            //normal repaired list
            commands.Add(3.3, new string[] { "1", inputType[elem.TEXT], "MRKT_NORM_ASIS_REPAIRED_REC", callingForm.SubjectMarketValueList });

            //90-day as-is sale
            commands.Add(3.4, new string[] { "1", inputType[elem.TEXT], "MRKT_90DAY_ASIS", callingForm.SubjectMarketValue });
            ///90-day as-is list
            commands.Add(3.5, new string[] { "1", inputType[elem.TEXT], "MRKT_90DAY_ASIS_REC", callingForm.SubjectMarketValueList });
            ///90-day repaired sale
            commands.Add(3.6, new string[] { "1", inputType[elem.TEXT], "MRKT_90DAY_ASIS_REPAIRED", callingForm.SubjectMarketValue });
            ///90-day repaired list
            commands.Add(3.7, new string[] { "1", inputType[elem.TEXT], "MRKT_90DAY_ASIS_REPAIRED_REC", callingForm.SubjectMarketValueList });
            //comments id="SWBC-BPO2-Ext:GMC_COMMENTS"
            //Normal marketing time is ~90 days with no expected value change within 30 days  at time of this report.
            commands.Add(3.8, new string[] { "1", inputType[elem.TEXTAREA], "GMC_COMMENTS", "Normal marketing time is ~90 days with no expected value change within 30 days  at time of this report." });

            //4: Neighborhood
            //Ext:NBR_LOC_SUBURBAN
            commands.Add(4.0, new string[] { "1", inputType[elem.CHECKBOX], "NBR_LOC_SUBURBAN", "YES" });
            //PROP_VALUES_STABLE
            commands.Add(4.1, new string[] { "1", inputType[elem.CHECKBOX], "PROP_VALUES_STABLE", "YES" });
            //DEMAND_SUPPLY_BALANCE
            commands.Add(4.2, new string[] { "1", inputType[elem.CHECKBOX], "DEMAND_SUPPLY_BALANCE", "YES" });
            //SUBJ_CONFORM_GOOD
            commands.Add(4.3, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_CONFORM_GOOD", "YES" });
            //CURB_APPEAL_GOOD
            commands.Add(4.4, new string[] { "1", inputType[elem.CHECKBOX], "CURB_APPEAL_GOOD", "YES" });
            //PERCENT_CHANGE
            commands.Add(4.41, new string[] { "1", inputType[elem.TEXT], "PERCENT_CHANGE", "0" });
            //SWBC-BPO2-Ext:NORM_MRKT_TIME
            commands.Add(4.5, new string[] { "1", inputType[elem.TEXT], "NORM_MRKT_TIME", callingForm.SubjectNeighborhood.avgDom.ToString() });
            //id="SWBC-BPO2-Ext:NEIGH_PRICE_LOW"
            commands.Add(4.6, new string[] { "1", inputType[elem.TEXT], "NEIGH_PRICE_LOW", callingForm.SubjectNeighborhood.minSalePrice.ToString() });
            //SWBC-BPO2-Ext:NEIGH_PRICE_HIGH
            commands.Add(4.7, new string[] { "1", inputType[elem.TEXT], "NEIGH_PRICE_HIGH", callingForm.SubjectNeighborhood.maxSalePrice.ToString() });

            //NEIGH_ZONING
            commands.Add(4.75, new string[] { "1", inputType[elem.TEXT], "NEIGH_ZONING", "No Known Environmental/Zoning/Title/Legal Issues." });
            
            //id="SWBC-BPO2-Ext:NEIGH_COMMENT"
           
            
            commands.Add(4.8, new string[] { "1", inputType[elem.TEXTAREA], "NEIGH_COMMENT", callingForm.richTextBoxNeighborhoodComments.Text });

            //5:subject details
            commands.Add(5.0, new string[] { "1", inputType[elem.SELECT], "SALES_SUBJ_LOCATION", "%Residential" });
            //SALES_SUBJ_SUBDIVISION'
            commands.Add(5.1, new string[] { "1", inputType[elem.TEXT], "SALES_SUBJ_SUBDIVISION", callingForm.subjectSubdivisionTextbox.Text });
            //SALES_SUBJ_SITE
            commands.Add(5.2, new string[] { "1", inputType[elem.TEXT], "SALES_SUBJ_SITE", callingForm.SubjectLotSize });
         //   commands.Add(5.3, new string[] { "1", inputType[elem.SELECT], "SALES_SUBJ_PROPERTY_TYPE", "%Single<SP>Family"  });
          //  commands.Add(5.35, new string[] { "1", inputType[elem.SELECT], "SALES_SUBJ_DESIGN", "%" + typeTranslation[callingForm.SubjectMlsType] });

            //SALES_SUBJ_EXTERIOR'
            commands.Add(5.4, new string[] { "1", inputType[elem.TEXT], "SALES_SUBJ_EXTERIOR", callingForm.SubjectExteriorFinish });
            commands.Add(5.5, new string[] { "1", inputType[elem.TEXT], "SALES_SUBJ_AGE", (DateTime.Now.Year - Convert.ToInt32(callingForm.SubjectYearBuilt)).ToString() });
            commands.Add(5.6, new string[] { "1", inputType[elem.SELECT], "SALES_SUBJ_QUALITY", "%Average" });
            commands.Add(5.7, new string[] { "1", inputType[elem.SELECT], "SALES_SUBJ_CONDITION", "%Average" });
            //SALES_SUBJ_TOT_ROOMS'
            commands.Add(5.8, new string[] { "1", inputType[elem.TEXT], "SALES_SUBJ_TOT_ROOMS", callingForm.SubjectRoomCount });
            commands.Add(5.9, new string[] { "1", inputType[elem.TEXT], "SALES_SUBJ_BEDROOMS", callingForm.SubjectBedroomCount });
            commands.Add(5.11, new string[] { "1", inputType[elem.TEXT], "SUBJ_BATHS", callingForm.SubjectBathroomCount });
            commands.Add(5.12, new string[] { "1", inputType[elem.TEXT], "SUBJ_GLASF", callingForm.SubjectAboveGLA });
            //:SALES_SUBJ_BSMT_SQFT'
            commands.Add(5.13, new string[] { "1", inputType[elem.TEXT], "SALES_SUBJ_BSMT_SQFT", callingForm.SubjectBasementGLA });
            //SALES_SUBJ_BSMT_FINISHED'
            commands.Add(5.14, new string[] { "1", inputType[elem.TEXT], "SALES_SUBJ_BSMT_FINISHED", callingForm.SubjectBasementFinishedGLA });
            //TODO:  Translation table for garage
            commands.Add(5.15, new string[] { "1", inputType[elem.SELECT], "SALES_SUBJ_GAR_CPT", "%" + garageTypeTranslation[GlobalVar.theSubjectProperty.GarageStallCount] });
            commands.Add(5.16, new string[] { "1", inputType[elem.SELECT], "SALES_SUBJ_FENCE_POOL", "%None" });

            //6: Broker info
            //BROKER_LIC 471.009162
            commands.Add(6.0, new string[] { "1", inputType[elem.TEXT], "BROKER_LIC", "471.009162" });
            commands.Add(6.1, new string[] { "1", inputType[elem.TEXT], "BROKER_SIGN_DATE", DateTime.Now.ToShortDateString() });

            //7: Repairs
            commands.Add(7.0, new string[] { "1", inputType[elem.TEXT], "REP_INT_PAINT", "0" });
            commands.Add(7.1, new string[] { "1", inputType[elem.TEXT], "REP_EXT_FOUNDATION", "0" });
            commands.Add(7.2, new string[] { "1", inputType[elem.TEXT], "REP_INT_WALLS", "0" });
            commands.Add(7.3, new string[] { "1", inputType[elem.TEXT], "REP_EXT_LSCAPE", "0" });
            commands.Add(7.4, new string[] { "1", inputType[elem.TEXT], "REP_INT_CARPET", "0" });
            commands.Add(7.5, new string[] { "1", inputType[elem.TEXT], "REP_EXT_ROOF", "0" });
            commands.Add(7.6, new string[] { "1", inputType[elem.TEXT], "REP_EXT_WINDOWS", "0" });
            commands.Add(7.7, new string[] { "1", inputType[elem.TEXT], "REP_INT_ELEC", "0" });
            commands.Add(7.8, new string[] { "1", inputType[elem.TEXT], "REP_EXT_GARAGE", "0" });
            commands.Add(7.9, new string[] { "1", inputType[elem.TEXT], "REP_INT_OTHER", "0" });
            commands.Add(7.11, new string[] { "1", inputType[elem.TEXT], "REP_EXT_TRIM", "0" });
            commands.Add(7.12, new string[] { "1", inputType[elem.TEXT], "REP_EXT_PAINT", "0" });
            commands.Add(7.13, new string[] { "1", inputType[elem.TEXT], "REP_EXT_OTHER", "0" });
            commands.Add(7.14, new string[] { "1", inputType[elem.TEXT], "REP_TOTAL", "0" });
            commands.Add(7.15, new string[] { "1", inputType[elem.TEXTAREA], "REP_COMMENT", "No repairs noted from inspection."});

            //8: Other/Comments

            commands.Add(8.0, new string[] { "1", inputType[elem.TEXTAREA], "SUP_COMMENTS", "All comparable sales and listings are within the same market as the subject and are in direct competition and share the same school district, transportation access and shopping access as the subject." });


            //commands.Add(10, new string[] { "1", inputType[elem.TEXT], "RECN_APPR_PHONE", "815.315.0203" });
            //commands.Add(11, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_EXT_CONDITION_AVERAGE",  "YES" });
            //commands.Add(12, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_LANDSCAPE_AVERAGE", "YES" });
            //commands.Add(13, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_AREA_AVERAGE", "YES" });
            //commands.Add(14, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_ROOF_AVERAGE", "YES" });
            //commands.Add(15, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_OWNER_PRIDE_AVERAGE", "YES" });
            //commands.Add(16, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_GEN_CONDITION_AVERAGE", "YES" });
            //commands.Add(17, new string[] { "1", inputType[elem.TEXT], "NBR_PROP_TYPE", "SFR" });
            //commands.Add(18, new string[] { "1", inputType[elem.CHECKBOX], "NBR_TREND_STABLE", "YES" });
            //commands.Add(19, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_PROP_OCCUPIED_OCCUPIED", "YES" });
            //commands.Add(20, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_CURRENTLY_LISTED_NO", "YES" });


           // commands.Add(5, new string[] { "1", inputType[elem.TEXT], "LIST_COUNT", callingForm.SubjectNeighborhood.numberActiveListings.ToString() });
            

           //commands.Add(5, new string[] { "1", inputType[elem.TEXT], "LIST_COUNT", callingForm.SubjectNeighborhood.numberActiveListings.ToString() });
           //commands.Add(6, new string[] { "1", inputType[elem.TEXT], "PROP_FOR_SALE_L", callingForm.SubjectNeighborhood.numberOfCompListings.ToString()});
           //commands.Add(7, new string[] { "1", inputType[elem.TEXT], "REO_CORP_OWNED_NBR", callingForm.SetOfComps.numberREOListings.ToString() });
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


           foreach (var c in commands)
           {
               macro.AppendFormat("TAG POS={0} TYPE={1} ATTR=ID:*{2} CONTENT={3}\r\n", c.Value[0], c.Value[1], c.Value[2], c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
           }





           string macroCode = macro.ToString();
           iim.iimPlayCode(macroCode, 30);
        }

      
  
        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {
            StringBuilder macro = new StringBuilder();
            //targetCompNumber = Regex.Match(compNum, @"\d").Value;
            targetCompNumber = compNum;

            Dictionary<string, string> saleTypeTranslation = new Dictionary<string, string>()
            {
                {"Arms Length", "19"}, {"REO", "1"}, {"ShortSale", "17"},
            };
            Dictionary<string, string> heatingTypeTranslation = new Dictionary<string, string>()
            {
                {"Gas", "6"}, {"Gas, Forced Air", "6"}, {@"Gas, Hot Water/Steam", "6"},{@"Electric", "2"},

            };

            Dictionary<string, string> CoolingTypeTranslation = new Dictionary<string, string>()
            {
                {"Central Air", "1"},  {"None", "0"}
            };

          

     
          
            Dictionary<double, string[]> commands = new Dictionary<double, string[]>();

            ////macro.AppendLine(@"SET !ERRORIGNORE YES");
            ////macro.AppendLine(@"SET !TIMEOUT_STEP 0");

            commands.Add(0, new string[] { "1", inputType[elem.TEXT], "ADDRESS_1", targetComp.StreetAddress });
            commands.Add(1, new string[] { "1", inputType[elem.TEXT], "ADDRESS_2", targetComp.City + " " + targetComp.Zipcode });
            commands.Add(2, new string[] { "1", inputType[elem.TEXT], "MLS", targetComp.MlsNumber });
            commands.Add(3, new string[] { "1", inputType[elem.TEXT], "PROXIMITY", targetComp.proximityToSubject.ToString() });
            commands.Add(4, new string[] { "1", inputType[elem.TEXT], "ORG_LIST_PRICE", targetComp.OriginalListPrice.ToString() });
            if (saleOrList == "sale")
            {
                //SALES_COMP1_LIST_PRICE_AT_SALE
                commands.Add(5, new string[] { "1", inputType[elem.TEXT], "LIST_PRICE_AT_SALE", targetComp.CurrentListPrice.ToString() });
                commands.Add(6, new string[] { "1", inputType[elem.TEXT], "SALE_PRICE", targetComp.SalePrice.ToString() });
                commands.Add(7, new string[] { "1", inputType[elem.TEXT], "DATE_OF_SALE", targetComp.SalesDate.ToShortDateString() });
            } else
            {
                commands.Add(7.5, new string[] { "1", inputType[elem.TEXT], "LIST_PRICE", targetComp.CurrentListPrice.ToString() });
            }
           
            commands.Add(8, new string[] { "1", inputType[elem.TEXT], "DAYS_ON_MARKET", targetComp.DOM });
            if (saleOrList == "sale")
            {
                commands.Add(9, new string[] { "1", inputType[elem.TEXT], "SALES_FIN1", targetComp.FinancingMlsString });
            }
            commands.Add(10, new string[] { "1", inputType[elem.SELECT], "REO", _reoYesNoFill(targetComp.TransactionType) });
            commands.Add(11, new string[] { "1", inputType[elem.SELECT], "LOCATION", "%Residential" });
            //_SUBDIVISION
            commands.Add(12, new string[] { "1", inputType[elem.TEXT], "SUBDIVISION", targetComp.Subdivision });
            commands.Add(13, new string[] { "1", inputType[elem.TEXT], "SITE", targetComp.Lotsize.ToString() });
            //LISTS_COMP1_PROPERTY_TYPE
       //     commands.Add(14, new string[] { "1", inputType[elem.SELECT], "PROPERTY_TYPE", "%Single Family" });
            //LISTS_COMP1_DESIGN
        //    commands.Add(15, new string[] { "1", inputType[elem.SELECT], "DESIGN", "%" + typeTranslation[targetComp.Type] });
            //LISTS_COMP1_EXTERIOR
            commands.Add(16, new string[] { "1", inputType[elem.TEXT], "EXTERIOR", targetComp.ExteriorMlsString });
            commands.Add(17, new string[] { "1", inputType[elem.TEXT], "AGE", targetComp.Age.ToString() });
            commands.Add(18, new string[] { "1", inputType[elem.SELECT], "QUALITY", "%Average" });
            commands.Add(19, new string[] { "1", inputType[elem.SELECT], "CONDITION", "%Average" });
            //LISTS_COMP1_TOT_ROOMS
            commands.Add(20, new string[] { "1", inputType[elem.TEXT], "TOT_ROOMS", targetComp.TotalRoomCount.ToString() });
            commands.Add(21, new string[] { "1", inputType[elem.TEXT], "BEDROOMS", targetComp.BedroomCount });
            commands.Add(22, new string[] { "1", inputType[elem.TEXT], "BATHS", targetComp.BathroomCount });
            commands.Add(23, new string[] { "1", inputType[elem.TEXT], "GLASF", targetComp.GLA.ToString() });
            commands.Add(24, new string[] { "1", inputType[elem.TEXT], "BSMT_SQFT", targetComp.BasementGLA() });
            commands.Add(25, new string[] { "1", inputType[elem.TEXT], "BSMT_FINISHED", targetComp.BasementGLA() });
            //GAR_CPT
            commands.Add(26, new string[] { "1", inputType[elem.SELECT], "GAR_CPT", "%" + garageTypeTranslation[targetComp.NumberGarageStalls()] });
            commands.Add(27, new string[] { "1", inputType[elem.SELECT], "FENCE_POOL", "%None" });

            //commands.Add(4, new string[] { "1", inputType[elem.TEXT], "PRICE_REDUCTION_COUNT", targetComp.NumberOfPriceChanges.ToString() });
            //commands.Add(5, new string[] { "1", inputType[elem.SELECT], "DATASOURCE_S", "%MLS" });
            //commands.Add(6, new string[] { "1", inputType[elem.TEXT], "SELLER_CONCESSIONS", targetComp.PointsInDollars });
            //commands.Add(7, new string[] { "1", inputType[elem.SELECT], "SALES_TYPE", "%" + saleTypeTranslation[targetComp.TransactionType] });



            //commands.Add(9, new string[] { "1", inputType[elem.SELECT], "LOCATION_DESCR", "%" + GlobalVar.mainWindow.comboBoxLocationDescr.Text });
            //commands.Add(10, new string[] { "1", inputType[elem.SELECT], "OWNERSHIP_TYPE", "%Fee Simple" });
            
            //commands.Add(12, new string[] { "1", inputType[elem.SELECT], "LOTSIZE_TYPE_S", "%A" });

            //if (targetComp.Waterfront)
            //{
            //    commands.Add(13, new string[] { "1", inputType[elem.TEXT], "PROPERTY_VIEW", "Water" });
            //}
            //else
            //{
            //    commands.Add(13, new string[] { "1", inputType[elem.TEXT], "PROPERTY_VIEW", "Residential" });
            //}
      
            //commands.Add(15, new string[] { "1", inputType[elem.TEXT], "YEAR_BUILT", targetComp.YearBuiltString });
   

            //commands.Add(21, new string[] { "1", inputType[elem.TEXT], "LIVING_SQFEET", targetComp.ProperGla(GlobalVar.mainWindow.SubjectAboveGLA) });

            //commands.Add(23, new string[] { "1", inputType[elem.TEXT], "ROOMS_BELOW_GRADE", targetComp.NumberOfBasementRooms() });


            //commands.Add(24, new string[] { "1", inputType[elem.SELECT], "HEATING", "%" + heatingTypeTranslation[targetComp.Heating] });
            //commands.Add(25, new string[] { "1", inputType[elem.SELECT], "COOLING", "%" + CoolingTypeTranslation[targetComp.Cooling] });

            //if (targetComp.AttachedGarage())
            //{
            //    commands.Add(26, new string[] { "1", inputType[elem.SELECT], "GARAGE_TYPE_S", "%AG" });
            //}
            //else if (targetComp.DetachedGarage())
            //{
            //    commands.Add(26, new string[] { "1", inputType[elem.SELECT], "GARAGE_TYPE_S", "%DG" });
            //}

            //commands.Add(27, new string[] { "1", inputType[elem.SELECT], "GARAGE_CAR_NUM_L", "%" + targetComp.NumberGarageStalls() });
            //commands.Add(28, new string[] { "1", inputType[elem.SELECT], "COOLING", "%" + CoolingTypeTranslation[targetComp.Cooling] });

            //commands.Add(29, new string[] { "1", inputType[elem.TEXT], "OTHER", "NA" });
            //commands.Add(30, new string[] { "2", inputType[elem.RADIO], "POOL_B", "YES" });
     

            foreach (var c in commands)
            {
                macro.AppendFormat("TAG POS={0} TYPE={1} ATTR=ID:*:{5}*_COMP{3}_{2} CONTENT={4}\r\n", c.Value[0], c.Value[1], c.Value[2], targetCompNumber, c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""), saleOrList.Substring(0,1));
            }

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 60);

            //commands.Add(1, new string[] { inputType, "tbCity", targetComp.City.Replace(" ", "<SP>") });
            //commands.Add(2, new string[] { "SELECT", "ddlState", "%IL" });
            //commands.Add(3, new string[] { inputType, "tbZip", targetComp.Zipcode });
            //commands.Add(4, new string[] { inputType, "tbMlsNumber", targetComp.MlsNumber });
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
            //commands.Add(18, new string[] { inputType, "tbBasementFinished", targetComp.BasementFinishedPercentage() });
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
