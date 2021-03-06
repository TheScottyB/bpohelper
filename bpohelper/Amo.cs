﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{

    class AMO : BPOFulfillment
    {
       
        public AMO()
        {
            subject = GlobalVar.theSubjectProperty;
            callingForm = GlobalVar.mainWindow;
        }

        public AMO(MLSListing m) 
            : this()
        {
            targetComp = m;
        }

        protected enum elem { TEXT, SELECT, TEXTAREA, RADIO, CHECKBOX };

         protected  Dictionary<elem, string> inputType = new Dictionary<elem, string>()
        {
            {elem.TEXT, "INPUT:TEXT"},{elem.SELECT, "SELECT"}, {elem.TEXTAREA, "TEXTAREA"}, {elem.RADIO, "INPUT:RADIO"}, {elem.CHECKBOX, "INPUT:CHECKBOX"}
        };


         protected Dictionary<int, MLSListing> stack;

         protected Form1 callingForm;
         protected SubjectProperty subject;

        public void Prefill(iMacros.App iim, Form1 form)
        {
            Dictionary<int, string[]> commands = new Dictionary<int, string[]>();

            Dictionary<string, string> translationTable = new Dictionary<string, string>();

            

            StringBuilder macro = new StringBuilder();

            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");


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
            commands.Add(11, new string[] { "1", inputType[elem.CHECKBOX], "SUBJ_EXT_CONDITION_AVERAGE",  "YES" });
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

            commands.Add(0, new string[] { "1", inputType[elem.TEXT], "ADDRESS_1", targetComp.StreetAddress + " ," + targetComp.City + " " + targetComp.Zipcode });
            commands.Add(1, new string[] { "1", inputType[elem.TEXT], "PROXIMITY", targetComp.proximityToSubject.ToString() });
            commands.Add(2, new string[] { "1", inputType[elem.TEXT], "UNITS", "1" });
            commands.Add(3, new string[] { "1", inputType[elem.TEXT], "SQFT", targetComp.GLA.ToString() });
            commands.Add(4, new string[] { "1", inputType[elem.TEXT], "BEDROOMS", targetComp.BedroomCount });
            commands.Add(5, new string[] { "1", inputType[elem.TEXT], "BATHS", targetComp.BathroomCount });
            commands.Add(6, new string[] { "1", inputType[elem.TEXT], "AGE", targetComp.Age.ToString() });
            commands.Add(7, new string[] { "1", inputType[elem.TEXT], "LOT_SIZE", (targetComp.Lotsize * 43560).ToString() });
            commands.Add(8, new string[] { "1", inputType[elem.TEXT], "DAYS_ON_MARKET", targetComp.DOM });
            commands.Add(9, new string[] { "1", inputType[elem.TEXT], "ORG_LIST_PRICE", targetComp.OriginalListPrice.ToString() });
            if (saleOrList == "sale")
            {
                commands.Add(10.10, new string[] { "1", inputType[elem.TEXT], "SALE_DATE", targetComp.SalesDate.ToShortDateString() });
                commands.Add(10.11, new string[] { "1", inputType[elem.TEXT], "FINANCING", targetComp.FinancingMlsString });
                commands.Add(10.12, new string[] { "1", inputType[elem.TEXT], "SALE_PRICE", targetComp.SalePrice.ToString() });
                commands.Add(10.13, new string[] { "2", inputType[elem.CHECKBOX], "COMMENTS", "YES" });
            }
            if (saleOrList == "list")
            {
                commands.Add(10.20, new string[] { "1", inputType[elem.TEXT], "CURR_LIST_PRICE", targetComp.CurrentListPrice.ToString() });
                commands.Add(10.21, new string[] { "1", inputType[elem.TEXT], "FINANCING", "UTD" });
                commands.Add(10.22, new string[] { "2", inputType[elem.CHECKBOX], "RATING", "YES" });
            }
           


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
            //commands.Add(14, new string[] { "1", inputType[elem.SELECT], "CONSTRUCTION_QUALITY", "%Average" });
            //commands.Add(15, new string[] { "1", inputType[elem.TEXT], "YEAR_BUILT", targetComp.YearBuiltString });
            //commands.Add(16, new string[] { "1", inputType[elem.SELECT], "CONDITION", "%Average" });
            //commands.Add(17, new string[] { "1", inputType[elem.TEXT], "UNIT_NUM_S", "1" });
            //commands.Add(18, new string[] { "1", inputType[elem.TEXT], "TOTAL_ROOMS_NUM", targetComp.TotalRoomCount.ToString() });
            //commands.Add(19, new string[] { "1", inputType[elem.TEXT], "BEDROOMS_NUM", targetComp.BedroomCount });
            //commands.Add(20, new string[] { "1", inputType[elem.TEXT], "BATHS_NUM", targetComp.BathroomCount });

            //commands.Add(21, new string[] { "1", inputType[elem.TEXT], "LIVING_SQFEET", targetComp.ProperGla(GlobalVar.mainWindow.SubjectAboveGLA) });
            //commands.Add(22, new string[] { "1", inputType[elem.TEXT], "BSMT_SQFEET_S", targetComp.BasementGLA() });
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
            //if (saleOrList == "sale")
            //{
            //    commands.Add(31, new string[] { "1", inputType[elem.TEXT], "SALE_PRICE", targetComp.SalePrice.ToString() });
            //    commands.Add(32, new string[] { "1", inputType[elem.TEXT], "CURRENTLIST_PRICE", targetComp.CurrentListPrice.ToString() });
        
            //}
         

            foreach (var c in commands)
            {
                macro.AppendFormat("TAG POS={0} TYPE={1} ATTR=ID:*{3}_{2} CONTENT={4}\r\n", c.Value[0], c.Value[1], c.Value[2], targetCompNumber, c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
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
