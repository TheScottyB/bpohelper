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
        { }

        public ClearCap(MLSListing m) 
            : this()
        {
            targetComp = m;
        }

       enum elem { TEXT, SELECT, TEXTAREA, RADIO };

        private Dictionary<int, MLSListing> stack;

        private Form1 callingForm;

        public void Prefill(iMacros.App iim, Form1 form)
        {
            Dictionary<int, string[]> commands = new Dictionary<int, string[]>();
            string [] inputType = {"INPUT:TEXT", "SELECT", "TEXTAREA", "INPUT:RADIO"};

            

            StringBuilder macro = new StringBuilder();
           

    //        commands.Add(0, new string[] { inputType[0], "apn_s", GlobalVar.theSubjectProperty.ParcelID });
    //        //SFR
    //        commands.Add(1, new string[] { inputType[3], "PROPERTYTYPE", "YES" });
    //        //R1
    //        commands.Add(2, new string[] { inputType[0], "ZONING", "R1" });
    //        //SFR - maybe get from realist
    //        commands.Add(3, new string[] { inputType[0], "ZONING_DESC_1", "SFR" });
    //        //conforming
    //        commands.Add(4, new string[] { inputType[1], "SUBJECT_USE_CODE_1", "%L" });
    //        //Occupied by owner
    //        commands.Add(1, new string[] { inputType[3], "OCCUPANCY", "YES" });





    //        foreach (var c in commands)
    //        {
    //            macro.AppendFormat("TAG POS=1 TYPE={0} FORM=ID:BPO_FORM ATTR=NAME:{1} CONTENT={2}\r\n", c.Value[0], c.Value[1], c.Value[2].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
    //        }


      


    //        string macroCode = macro.ToString();
    //        iim.iimPlayCode(macroCode, 30);


    //;
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:OWNER_USE_1 CONTENT=YES");
    //        macro.AppendLine(@"TAG POS=3 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:SUBJ_IMPROVEMENT CONTENT=YES");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:IS_FINANCING_AVAIL CONTENT=YES");
    //        macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:hoa_1_has_hoa CONTENT=YES");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:hoa_1_has_hoa CONTENT=NO");
    //        macro.AppendLine(@"'radio input activated");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:LISTED CONTENT=YES");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LISTING_FIRM_1");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LISTING_AGENT_S");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LISTING_AGENT_PHONE_S CONTENT=<SP>");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:ORIGLIST_DT_1");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:ORIGLIST_PRICE_1");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LISTED_DATE_3");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LIST_AMOUNT_L_3");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LISTED_DATE_2");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LIST_AMOUNT_L_2");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LISTED_DATE_1");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LIST_AMOUNT_L_1");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LIST_DATE_1");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LIST_PRICE_1");
    //        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPO_FORM ATTR=NAME:CONDITIONCOMM_S");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LISTED_DATE_4");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LIST_AMOUNT_L_4");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LISTED_DATE_5");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LIST_AMOUNT_L_5");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LISTED_DATE_6");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LIST_AMOUNT_L_6");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:SOLD_DATE_4");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:SALE_AMOUNT_L_4");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:SOLD_DATE_5");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:SALE_AMOUNT_L_5");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:SOLD_DATE_6");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:SALE_AMOUNT_L_6");
    //        macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:LOCATION_DESCR_1 CONTENT=YES");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:BOUNDARY_NORTH");
    //        macro.AppendLine(@"TAG POS=2 TYPE=TD FORM=NAME:BPO_FORM ATTR=TXT:North<SP>Boundary<SP>Street:South<SP>Boundary<SP>Street:East<SP>Boundary<SP>Street:West<SP>Boundary<SP>*");
    //        macro.AppendLine(@"TAG POS=1 TYPE=STRONG FORM=NAME:BPO_FORM ATTR=TXT:South<SP>Boundary<SP>Street:");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:BOUNDARY_SOUTH");
    //        macro.AppendLine(@"TAG POS=69 TYPE=DIV ATTR=*");
    //        macro.AppendLine(@"TAG POS=175 TYPE=TD FORM=NAME:BPO_FORM ATTR=*");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:BOUNDARY_EAST");
    //        macro.AppendLine(@"TAG POS=3 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:LOCATION_DESCR_1 CONTENT=YES");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:BOUNDARY_RADIUS CONTENT=5");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LIST_COUNT");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:PROP_FOR_SALE_L");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:REO_CORP_OWNED_NBR");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LIST_LOW_S");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:LIST_HIGH_S");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:SALE_COUNT");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:SALE_LOW_S");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:SALE_HIGH_S");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:BOARDED_UP_HOMES_NBR");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:MARKETTYPE_S CONTENT=NO");
    //        macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:MARKETTYPE_S CONTENT=YES");
    //        macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:EMPLOYMENT_CONDITIONS CONTENT=YES");
    //        macro.AppendLine(@"TAG POS=3 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:MARKETTYPE CONTENT=YES");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:DEMAND_SUPPLY_S CONTENT=YES");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:DOM_AVG_L CONTENT=111");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPO_FORM ATTR=NAME:OWNER_PCT CONTENT=90");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:REPAIRS CONTENT=YES");
    //        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:BPO_FORM ATTR=NAME:PROBABLE_PURCHASER_S CONTENT=NO");


    //        #region pageOne
            ////
            ////Page 01 - Subject Property
            ////
            //commands.Add(0, new string[] { inputType, "tbInspectionDate", form.dateTimePickerInspectionDate.Value.ToShortDateString() });
            //commands.Add(1, new string[] { inputType, "tbApnNumber", GlobalVar.theSubjectProperty.ParcelID });
            //commands.Add(2, new string[] { inputType, "tbNumberOfUnits", "1" });
            //commands.Add(3, new string[] { inputType, "tbBedrooms",  GlobalVar.theSubjectProperty.MainForm.SubjectBedroomCount});
            //commands.Add(4, new string[] { inputType, "tbBathrooms", GlobalVar.theSubjectProperty.MainForm.SubjectBathroomCount });
            //commands.Add(5, new string[] { inputType, "tbTotalRooms", GlobalVar.theSubjectProperty.MainForm.SubjectRoomCount });
            //commands.Add(6, new string[] { inputType, "tbYearBuilt", GlobalVar.theSubjectProperty.MainForm.SubjectYearBuilt});
            //commands.Add(7, new string[] { inputType, "tbSquareFeet", GlobalVar.theSubjectProperty.MainForm.SubjectAboveGLA });
            //commands.Add(8, new string[] { inputType, "tbLotSize", GlobalVar.theSubjectProperty.MainForm.SubjectLotSize});
            //commands.Add(9, new string[] { inputType, "tbBasementSquareFeet", GlobalVar.theSubjectProperty.MainForm.SubjectBasementGLA});
            //commands.Add(10, new string[] { inputType, "tbBasementPercentage", GlobalVar.theSubjectProperty.BasementFinishedPercentage.ToString() });
            //commands.Add(11, new string[] { inputType, "tbSellerConcessions", "NA" });
            //commands.Add(12, new string[] { inputType, "tbEstimatedLandValue",  GlobalVar.theSubjectProperty.MainForm.SubjectLandValue});
            
            //if (GlobalVar.theSubjectProperty.IsActiveListing)
            //{
            //    commands.Add(13, new string[] { inputType, "tbOriginalListPrice", GlobalVar.theSubjectProperty.origListingPrice});
            //      commands.Add(14, new string[] { inputType, "tbOriginalListDate", GlobalVar.theSubjectProperty.GetFieldValue(@"List Date:")});
            //      commands.Add(15, new string[] { inputType, "tbMlsNumber", GlobalVar.theSubjectProperty.GetFieldValue(@"MLS #:") });

            //}
            //else
            //{
            //      commands.Add(13, new string[] { inputType, "tbOriginalListPrice", "NOTLISTED"});
            //        commands.Add(14, new string[] { inputType, "tbOriginalListDate", "NOTLISTED"});
            //        commands.Add(15, new string[] { inputType, "tbMlsNumber", "NOTLISTED" });
            //}

            //commands.Add(16, new string[] { "SELECT", "ddlStyle", "%Conv" });
            //commands.Add(17, new string[] { "SELECT", "ddlConstructionType", "%Frame" });





            //string lresGarageStr = "None";
            //string numSpaces = Regex.Match(form.SubjectParkingType, @"\d").Value;
            //string att_det = "";

            //if (!string.IsNullOrEmpty(numSpaces))
            //{
            //    if (form.SubjectParkingType.ToLower().Contains("att"))
            //    {
            //        att_det = "Attached";
            //    }
            //    else if (form.SubjectParkingType.ToLower().Contains("det"))
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

            //commands.Add(18, new string[] { "SELECT", "ddlParking", "%" + lresGarageStr });

            //if (GlobalVar.mainWindow.SubjectAttached)
            //{
            //    commands.Add(19, new string[] { "SELECT", "ddlPropertyType", "%Attached" });
            //}
            //else
            //{
            //    commands.Add(19, new string[] { "SELECT", "ddlPropertyType", "%Detached" });
            //}


            //commands.Add(20, new string[] { "SELECT", "ddlOccupancy", "%" + GlobalVar.mainWindow.comboBoxSubjectOccupancy.Text });
            //commands.Add(21, new string[] { "SELECT", "ddlCondition", "%" + GlobalVar.mainWindow.comboBoxSubjectCondition.Text });
            //commands.Add(22, new string[] { "SELECT", "ddlLocation", "%" + GlobalVar.mainWindow.comboBoxSubjectLocationCondition.Text });
            //commands.Add(23, new string[] { "SELECT", "ddlView", "%" + GlobalVar.mainWindow.comboBoxSubjectView.Text });
            //commands.Add(24, new string[] { "SELECT", "ddlDataSource", "%Tax Record"});
            

            //if (GlobalVar.theSubjectProperty.IsActiveListing)
            //{
            //    commands.Add(25, new string[] { "INPUT:RADIO", "rbIsCurrentlyListed_0", "YES" });
            //    commands.Add(26, new string[] { inputType, "tbDom", GlobalVar.mainWindow.SubjectDom });
            //    commands.Add(27, new string[] { inputType, "tbCurrentListPrice", GlobalVar.mainWindow.SubjectCurrentListPrice });
            //    commands.Add(28, new string[] { inputType, "tbCurrentListDate", GlobalVar.mainWindow.dateTimePickerSubjectCurrentListDate.Value.ToShortDateString() });
            //    commands.Add(29, new string[] { inputType, "tbListingCompany", GlobalVar.theSubjectProperty.GetFieldValue(@"Broker:")});
            //    commands.Add(30, new string[] { inputType, "tbListingCompanyPhone", GlobalVar.theSubjectProperty.BrokerPhone.Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ","") });

            //}
            //else
            //{
            //    commands.Add(25, new string[] { "INPUT:RADIO", "rbIsCurrentlyListed_1", "YES" });
            //}

            //commands.Add(31, new string[] { "INPUT:RADIO", "rbIsThereHoa_1", "YES" });

            //commands.Add(32, new string[] { "TEXTAREA", "tbCommentSubjectCondition", ".TBD." });
      

            //     commands.Add(33, new string[] { "INPUT:RADIO", "rbRecommendRepairs_1", "YES" });



            //     ////  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button4&&VALUE:Save");


            //     //  macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbSubjectConditionComment CONTENT=No<SP>adverse<SP>conditions<SP>were<SP>noted<SP>at<SP>the<SP>time<SP>of<SP>inspection<SP>based<SP>on<SP>exterior<SP>observations.");

            //     //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button2&&VALUE:Save<SP>and<SP>Continue");


            //foreach (var c in commands)
            //{
            //    macro.AppendFormat("TAG POS=1 TYPE={0} FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_{1} CONTENT={2}\r\n", c.Value[0], c.Value[1], c.Value[2].Replace(" ", "<SP>").Replace("$", "").Replace(",",""));
            //}
            // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$btnSaveAndContinue");


            // string macroCode = macro.ToString();
            // iim.iimPlayCode(macroCode, 30);

            //#endregion

            //#region pageTwo 
            // //
            // //Page 2 Neighborhood Data 
            // //
            // commands.Clear();
            // macro.Clear();
            // //tbSalesPriceLow

            // commands.Add(34, new string[] { inputType, "tbSalesPriceLow", GlobalVar.mainWindow.SubjectNeighborhood.minSalePrice.ToString() });
            // commands.Add(35, new string[] { inputType, "tbSalesPriceHigh", GlobalVar.mainWindow.SubjectNeighborhood.maxSalePrice.ToString() });
            // commands.Add(36, new string[] { inputType, "tbSalesPriceAverage", GlobalVar.mainWindow.SubjectNeighborhood.medianSalePrice.ToString() });

            // commands.Add(37, new string[] { inputType, "tbPercentageOwned", "90" });
            // commands.Add(38, new string[] { inputType, "tbPercentageTenant", "10" });

            // commands.Add(39, new string[] { "SELECT", "ddlPropertyValue", "%Increasing" });
            // commands.Add(40, new string[] { inputType, "tbPropertyValueRateChange", "1" });
            // commands.Add(41, new string[] { "SELECT", "ddlLocation", "%Suburban" });
            // commands.Add(42, new string[] { "SELECT", "ddlDemandAndSupply", "%Stable" });
            // commands.Add(43, new string[] { "SELECT", "ddlGrowth", "%Stable" });
            // commands.Add(44, new string[] { "SELECT", "ddlDemandAndSupply", "%Stable" });
            // commands.Add(45, new string[] { "SELECT", "ddlEconomy", "%Stable" });
            // commands.Add(46, new string[] { "SELECT", "ddlMarketingTime", "%3-6 mo" });

            // commands.Add(47, new string[] { "INPUT:RADIO", "rblIsReoDriven_1", "YES" });
            // commands.Add(48, new string[] { "INPUT:RADIO", "rblIsVacant_1", "YES" });
            // commands.Add(49, new string[] { "SELECT", "ddlProbablePurchaser", "%2nd Time Buyer" });
            // commands.Add(50, new string[] { "SELECT", "ddlFinanceType", "%Conv" });

            // commands.Add(51, new string[] { inputType, "tbSalesInImmediateArea", GlobalVar.mainWindow.SubjectNeighborhood.numberSoldListings.ToString() });
            // commands.Add(52, new string[] { inputType, "tbListingsInImmediateArea", GlobalVar.mainWindow.SubjectNeighborhood.numberActiveListings.ToString() });
            // commands.Add(53, new string[] { inputType, "tbSalesReo", GlobalVar.mainWindow.SubjectNeighborhood.numberREOSales.ToString() });
            // commands.Add(54, new string[] { inputType, "tbListingsReo", GlobalVar.mainWindow.SubjectNeighborhood.numberREOListings.ToString() });

            // commands.Add(55, new string[] { inputType, "tbEstimatedMonthlyRental", GlobalVar.mainWindow.SubjectRent});
            // commands.Add(56, new string[] { inputType, "tbComparableRentalListing", (Math.Round(GlobalVar.mainWindow.SubjectNeighborhood.numberActiveListings * .1)).ToString() });

            // commands.Add(57, new string[] { "INPUT:RADIO", "rbIsNewConstruction_1", "YES" });

            // commands.Add(58, new string[] { "TEXTAREA", "tbExplainObsolescense", ".TBD." });
            // commands.Add(59, new string[] { "TEXTAREA", "tbExplainLocationInfluences", ".TBD." });


            // foreach (var c in commands)
            // {
            //     macro.AppendFormat("TAG POS=1 TYPE={0} FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_{1} CONTENT={2}\r\n", c.Value[0], c.Value[1], c.Value[2].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            // }

            // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$btnSaveAndContinue");
            // macroCode = macro.ToString();
            // iim.iimPlayCode(macroCode, 30);


            //#endregion

          


        }

  
        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {
            StringBuilder macro = new StringBuilder();
            targetCompNumber = Regex.Match(compNum, @"\d").Value;
         
            Dictionary<string, string> saleTypeTranslation = new Dictionary<string, string>()
            {
                {"Arms Length", "19"}, {"REO", "1"}, {"ShortSale", "17"},
            };
            Dictionary<string, string> heatingTypeTranslation = new Dictionary<string, string>()
            {
                {"Gas", "6"}, {"Gas, Forced Air", "6"}, {@"Gas, Hot Water/Steam", "6"},

            };

            Dictionary<string, string> CoolingTypeTranslation = new Dictionary<string, string>()
            {
                {"Central Air", "1"},  {"None", "0"}
            };

     
          
            Dictionary<double, string[]> commands = new Dictionary<double, string[]>();
           // string[] inputType = { "INPUT:TEXT", "SELECT", "TEXTAREA", "INPUT:RADIO" };

            Dictionary<elem, string> inputType = new Dictionary<elem, string>();
            inputType.Add(elem.TEXT, "INPUT:TEXT");
            inputType.Add(elem.SELECT, "SELECT");
            inputType.Add(elem.TEXTAREA, "TEXTAREA");
            inputType.Add(elem.RADIO, "INPUT:RADIO");


            ////macro.AppendLine(@"SET !ERRORIGNORE YES");
            ////macro.AppendLine(@"SET !TIMEOUT_STEP 0");


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
