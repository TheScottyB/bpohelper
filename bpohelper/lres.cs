using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{

    class Lres : BPOFulfillment
    {
       
        public Lres()
        { }

        public Lres(MLSListing m) 
            : this()
        {
            targetComp = m;
        }
        private Dictionary<int, MLSListing> stack;
        private enum bpoEntryFormType { FORM1, FORM2 };

        private Form1 callingForm;
        private iMacros.App iim;

        private void Prefill_Form2()
        {
            Dictionary<double, string[]> commands = new Dictionary<double, string[]>();
            string inputType = "INPUT:TEXT";

            string fiftyWordComment = @"COMMENTS ON SUBJECT'S CONDITION: Notice: Our client relies heavily on your comments. Please provide original exterior physical description,overall exterior condition (& interior, if available) and surrounding neighborhood characteristics of the subject property base on a drive-by. Cut and paste comments from previous work is not acceptable. COMMENTS ON SUBJECT'S CONDITION: Notice: Our client relies heavily on your comments. Please provide original exterior physical description,overall exterior condition (& interior, if available) and surrounding neighborhood characteristics of the subject property base on a drive-by. Cut and paste comments from previous work is not acceptable. " ;

            StringBuilder macro = new StringBuilder();
            int timeout = 60;
            iMacros.Status status = new iMacros.Status();

            string lresGarageStr = "None";
            string numSpaces = Regex.Match(callingForm.SubjectParkingType, @"\d").Value;
            string att_det = "";

            if (!string.IsNullOrEmpty(numSpaces))
            {
                if (callingForm.SubjectParkingType.ToLower().Contains("att"))
                {
                    att_det = "CA";
                }
                else if (callingForm.SubjectParkingType.ToLower().Contains("det"))
                {
                    att_det = "CD";
                }

                switch (numSpaces)
                {
                    case "1":
                        lresGarageStr = "1" + att_det;
                        break;
                    case "2":
                        lresGarageStr = "2" + att_det;
                        break;
                    default:
                        lresGarageStr = "2+" + att_det;
                        break;
                }

            }

            //
            //Page 1
            //

            if (GlobalVar.theSubjectProperty.IsActiveListing)
            {
                commands.Add(1, new string[] { "1", "SELECT", "currently_listed", "%1" });
                commands.Add(1.1, new string[] { "1", inputType, "previous_dom",  GlobalVar.mainWindow.SubjectDom });
                commands.Add(1.2, new string[] { "1", inputType, "previous_list_price", GlobalVar.mainWindow.textBoxSubjectOriginalListPRice.Text });
                commands.Add(1.3, new string[] { "1", inputType, "current_list_price", GlobalVar.mainWindow.SubjectCurrentListPrice });
                commands.Add(1.4, new string[] { "1", inputType, "listing_company", GlobalVar.mainWindow.textBoxSubjextListingBrokerage.Text });
            }
            else
            {
                commands.Add(1.5, new string[] { "1", "SELECT", "currently_listed", "%0" });
            }

            commands.Add(0, new string[] { "1", "SELECT", "property_type", "%SFR" });
            if (GlobalVar.theSubjectProperty.IsOccupied)
            {
                commands.Add(2, new string[] { "1", "SELECT", "vacant_or_occupied", "%O" });
            }
            else
            {
                commands.Add(2, new string[] { "1", "SELECT", "vacant_or_occupied", "%V" });
            }
            commands.Add(3, new string[] { "1", "SELECT", "ca_condition", "%Average" });
            commands.Add(4, new string[] { "1", "SELECT", "ca_location", "%Average" });
            commands.Add(5, new string[] { "1", "SELECT", "ca_theview", "%Residential" });
            commands.Add(6, new string[] { "1", inputType, "seller_concessions", "NA" });
            commands.Add(7, new string[] { "1", "SELECT", "prop_source", "%1" });
            commands.Add(8, new string[] { "1", inputType, "hoa_fees_per_month", "0" });
            commands.Add(9, new string[] { "1", inputType, "legal_description", GlobalVar.theSubjectProperty.ParcelID });
            commands.Add(10, new string[] { "1", "TEXTAREA", "condition_comments", "......TBD......." + fiftyWordComment  });

            commands.Add(11, new string[] { "1", inputType, "ca_sqft_gla", GlobalVar.mainWindow.SubjectAboveGLA });
            commands.Add(11.1, new string[] { "1", "SELECT", "ca_style", "%Conv" });
            commands.Add(11.2, new string[] { "1", "SELECT", "ca_ext_walls", "%Frame" });
            commands.Add(11.3, new string[] { "1", inputType, "ca_total_rooms", GlobalVar.theSubjectProperty.MainForm.SubjectRoomCount });
            commands.Add(11.4, new string[] { "1", inputType, "ca_bedrooms", GlobalVar.theSubjectProperty.MainForm.SubjectBedroomCount });
            commands.Add(11.5, new string[] { "1", inputType, "ca_bathrooms", GlobalVar.theSubjectProperty.MainForm.SubjectBathroomCount });
            commands.Add(11.6, new string[] { "1", inputType, "basement_pct_finished", GlobalVar.theSubjectProperty.BasementFinishedPercentage.ToString() });
            commands.Add(11.7, new string[] { "1", inputType, "basement_sqft", GlobalVar.theSubjectProperty.MainForm.SubjectBasementGLA });
            commands.Add(11.8, new string[] { "1", "SELECT", "ca_garage_carport", "%" + lresGarageStr });
            commands.Add(11.9, new string[] { "1", inputType, "ca_lot_size", GlobalVar.theSubjectProperty.MainForm.SubjectLotSize });
            commands.Add(11.11, new string[] { "1", inputType, "ca_year_built", GlobalVar.theSubjectProperty.MainForm.SubjectYearBuilt });

            commands.Add(12, new string[] { "1", "SELECT", "rec_repair", "%2" });

            commands.Add(13, new string[] { "1", inputType, "nbhd_sale_price_low", GlobalVar.mainWindow.SubjectNeighborhood.minSalePrice.ToString() });
            commands.Add(13.1, new string[] { "1", inputType, "nbhd_sale_price_high", GlobalVar.mainWindow.SubjectNeighborhood.maxSalePrice.ToString() });
            commands.Add(13.2, new string[] { "1", inputType, "nbhd_sale_price_avg", GlobalVar.mainWindow.SubjectNeighborhood.medianSalePrice.ToString() });
            commands.Add(13.3, new string[] { "1", inputType, "nbhd_pct_owners", "90" });
            commands.Add(13.4, new string[] { "1", inputType, "nbhd_pct_tenants", "10" });
            commands.Add(13.5, new string[] { "1", inputType, "nbhd_listings", GlobalVar.mainWindow.SubjectNeighborhood.numberActiveListings.ToString() });
            commands.Add(13.6, new string[] { "1", inputType, "nbhd_listings_reo", GlobalVar.mainWindow.SubjectNeighborhood.numberREOListings.ToString() });
            commands.Add(13.7, new string[] { "1", inputType, "nbhd_sales", GlobalVar.mainWindow.SubjectNeighborhood.numberSoldListings.ToString() });
            commands.Add(13.8, new string[] { "1", inputType, "nbhd_sales_reo", GlobalVar.mainWindow.SubjectNeighborhood.numberREOSales.ToString() });
            commands.Add(13.9, new string[] { "1", inputType, "nbhd_estimated_monthly_rental", GlobalVar.mainWindow.SubjectRent });
            commands.Add(13.11, new string[] { "1", inputType, "nbhd_comparable_rental_listings", (Math.Round(GlobalVar.mainWindow.SubjectNeighborhood.numberActiveListings * .1)).ToString() });
            commands.Add(13.12, new string[] { "1", inputType, "economic_obsolescence", "None" });
            commands.Add(13.13, new string[] { "1", inputType, "location_influences", ".....TBD....." });




            //commands.Add(0, new string[] { inputType, "tbInspectionDate", form.dateTimePickerInspectionDate.Value.ToShortDateString() });
            //commands.Add(12, new string[] { inputType, "tbEstimatedLandValue", GlobalVar.theSubjectProperty.MainForm.SubjectLandValue });

          
            //if (GlobalVar.mainWindow.SubjectAttached)
            //{
            //    commands.Add(19, new string[] { "SELECT", "ddlPropertyType", "%Attached" });
            //}
            //else
            //{
            //    commands.Add(19, new string[] { "SELECT", "ddlPropertyType", "%Detached" });
            //}

            //commands.Add(33, new string[] { "INPUT:RADIO", "rbRecommendRepairs_1", "YES" });



            //////  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button4&&VALUE:Save");


            ////  macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbSubjectConditionComment CONTENT=aaaa;

            ////  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button2&&VALUE:Save<SP>and<SP>Continue");
            //commands.Add(57, new string[] { "INPUT:RADIO", "rbIsNewConstruction_1", "YES" });

            //commands.Add(58, new string[] { "TEXTAREA", "tbExplainObsolescense", ".TBD." });
            //commands.Add(59, new string[] { "TEXTAREA", "tbExplainLocationInfluences", ".TBD." });

            foreach (var c in commands)
            {
                macro.AppendFormat("TAG POS={0} TYPE={1} FORM=NAME:bpoForm ATTR=NAME:{2} CONTENT={3}\r\n", c.Value[0], c.Value[1], c.Value[2], c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            }
         
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:bpoForm ATTR=NAME:nbhd_growth_rate CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:bpoForm ATTR=NAME:nbhd_location CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:bpoForm ATTR=NAME:property_values CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:bpoForm ATTR=NAME:nbhd_supply_demand CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:bpoForm ATTR=NAME:nbhd_economy CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:bpoForm ATTR=NAME:nbhd_market_time_range CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoForm ATTR=NAME:nbhd_reo_market CONTENT=%N");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoForm ATTR=NAME:nbhd_boarded_homes CONTENT=%N");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoForm ATTR=NAME:new_construction CONTENT=%N");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoForm ATTR=NAME:probable_purchaser CONTENT=%Move<SP>Up");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoForm ATTR=NAME:probable_financing CONTENT=%Conv");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON ATTR=NAME:btnSaveAndCont");

       
          
            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 30);

            //#endregion
        }



        private void Prefill_Form1()
        {
            Dictionary<int, string[]> commands = new Dictionary<int, string[]>();
            string inputType = "INPUT:TEXT";

            StringBuilder macro = new StringBuilder();

            #region pageOne
            //
            //Page 01 - Subject Property
            //
            commands.Add(0, new string[] { inputType, "tbInspectionDate", callingForm.dateTimePickerInspectionDate.Value.ToShortDateString() });
            commands.Add(1, new string[] { inputType, "tbApnNumber", GlobalVar.theSubjectProperty.ParcelID });
            commands.Add(2, new string[] { inputType, "tbNumberOfUnits", "1" });
            commands.Add(3, new string[] { inputType, "tbBedrooms",  GlobalVar.theSubjectProperty.MainForm.SubjectBedroomCount});
            commands.Add(4, new string[] { inputType, "tbBathrooms", GlobalVar.theSubjectProperty.MainForm.SubjectBathroomCount });
            commands.Add(5, new string[] { inputType, "tbTotalRooms", GlobalVar.theSubjectProperty.MainForm.SubjectRoomCount });
            commands.Add(6, new string[] { inputType, "tbYearBuilt", GlobalVar.theSubjectProperty.MainForm.SubjectYearBuilt});
            commands.Add(7, new string[] { inputType, "tbSquareFeet", GlobalVar.theSubjectProperty.MainForm.SubjectAboveGLA });
            commands.Add(8, new string[] { inputType, "tbLotSize", GlobalVar.theSubjectProperty.MainForm.SubjectLotSize});
            commands.Add(9, new string[] { inputType, "tbBasementSquareFeet", GlobalVar.theSubjectProperty.MainForm.SubjectBasementGLA});
            commands.Add(10, new string[] { inputType, "tbBasementPercentage", GlobalVar.theSubjectProperty.BasementFinishedPercentage.ToString() });
            commands.Add(11, new string[] { inputType, "tbSellerConcessions", "NA" });
            commands.Add(12, new string[] { inputType, "tbEstimatedLandValue",  GlobalVar.theSubjectProperty.MainForm.SubjectLandValue});
            
            if (GlobalVar.theSubjectProperty.IsActiveListing)
            {
                commands.Add(13, new string[] { inputType, "tbOriginalListPrice", GlobalVar.theSubjectProperty.origListingPrice});
                  commands.Add(14, new string[] { inputType, "tbOriginalListDate", GlobalVar.theSubjectProperty.GetFieldValue(@"List Date:")});
                  commands.Add(15, new string[] { inputType, "tbMlsNumber", GlobalVar.theSubjectProperty.GetFieldValue(@"MLS #:") });

            }
            else
            {
                  commands.Add(13, new string[] { inputType, "tbOriginalListPrice", "NOTLISTED"});
                    commands.Add(14, new string[] { inputType, "tbOriginalListDate", "NOTLISTED"});
                    commands.Add(15, new string[] { inputType, "tbMlsNumber", "NOTLISTED" });
            }

            commands.Add(16, new string[] { "SELECT", "ddlStyle", "%Conv" });
            commands.Add(17, new string[] { "SELECT", "ddlConstructionType", "%Frame" });





            string lresGarageStr = "None";
            string numSpaces = Regex.Match(callingForm.SubjectParkingType, @"\d").Value;
            string att_det = "";

            if (!string.IsNullOrEmpty(numSpaces))
            {
                if (callingForm.SubjectParkingType.ToLower().Contains("att"))
                {
                    att_det = "Attached";
                }
                else if (callingForm.SubjectParkingType.ToLower().Contains("det"))
                {
                    att_det = "Detached";
                }

                switch (numSpaces)
                {
                    case "1":
                        lresGarageStr = "1<SP>Car<SP>" + att_det;
                        break;
                    case "2":
                        lresGarageStr = "2<SP>Car<SP>" + att_det;
                        break;
                    default:
                        lresGarageStr = "2+<SP>Car<SP>" + att_det;
                        break;
                }

            }

            commands.Add(18, new string[] { "SELECT", "ddlParking", "%" + lresGarageStr });

            if (GlobalVar.mainWindow.SubjectAttached)
            {
                commands.Add(19, new string[] { "SELECT", "ddlPropertyType", "%Attached" });
            }
            else
            {
                commands.Add(19, new string[] { "SELECT", "ddlPropertyType", "%Detached" });
            }


            commands.Add(20, new string[] { "SELECT", "ddlOccupancy", "%" + GlobalVar.mainWindow.comboBoxSubjectOccupancy.Text });
            commands.Add(21, new string[] { "SELECT", "ddlCondition", "%" + GlobalVar.mainWindow.comboBoxSubjectCondition.Text });
            commands.Add(22, new string[] { "SELECT", "ddlLocation", "%" + GlobalVar.mainWindow.comboBoxSubjectLocationCondition.Text });
            commands.Add(23, new string[] { "SELECT", "ddlView", "%" + GlobalVar.mainWindow.comboBoxSubjectView.Text });
            commands.Add(24, new string[] { "SELECT", "ddlDataSource", "%Tax Record"});
            

            if (GlobalVar.theSubjectProperty.IsActiveListing)
            {
                commands.Add(25, new string[] { "INPUT:RADIO", "rbIsCurrentlyListed_0", "YES" });
                commands.Add(26, new string[] { inputType, "tbDom", GlobalVar.mainWindow.SubjectDom });
                commands.Add(27, new string[] { inputType, "tbCurrentListPrice", GlobalVar.mainWindow.SubjectCurrentListPrice });
                commands.Add(28, new string[] { inputType, "tbCurrentListDate", GlobalVar.mainWindow.dateTimePickerSubjectCurrentListDate.Value.ToShortDateString() });
                commands.Add(29, new string[] { inputType, "tbListingCompany", GlobalVar.theSubjectProperty.GetFieldValue(@"Broker:")});
                commands.Add(30, new string[] { inputType, "tbListingCompanyPhone", GlobalVar.theSubjectProperty.BrokerPhone.Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ","") });

            }
            else
            {
                commands.Add(25, new string[] { "INPUT:RADIO", "rbIsCurrentlyListed_1", "YES" });
            }

            commands.Add(31, new string[] { "INPUT:RADIO", "rbIsThereHoa_1", "YES" });

            commands.Add(32, new string[] { "TEXTAREA", "tbCommentSubjectCondition", ".TBD." });
      

                 commands.Add(33, new string[] { "INPUT:RADIO", "rbRecommendRepairs_1", "YES" });



                 ////  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button4&&VALUE:Save");


                 //  macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbSubjectConditionComment CONTENT=No<SP>adverse<SP>conditions<SP>were<SP>noted<SP>at<SP>the<SP>time<SP>of<SP>inspection<SP>based<SP>on<SP>exterior<SP>observations.");

                 //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button2&&VALUE:Save<SP>and<SP>Continue");


            foreach (var c in commands)
            {
                macro.AppendFormat("TAG POS=1 TYPE={0} FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_{1} CONTENT={2}\r\n", c.Value[0], c.Value[1], c.Value[2].Replace(" ", "<SP>").Replace("$", "").Replace(",",""));
            }
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$btnSaveAndContinue");


             string macroCode = macro.ToString();
             iim.iimPlayCode(macroCode, 30);

            #endregion

            #region pageTwo 
             //
             //Page 2 Neighborhood Data 
             //
             commands.Clear();
             macro.Clear();
             //tbSalesPriceLow

             commands.Add(34, new string[] { inputType, "tbSalesPriceLow", GlobalVar.mainWindow.SubjectNeighborhood.minSalePrice.ToString() });
             commands.Add(35, new string[] { inputType, "tbSalesPriceHigh", GlobalVar.mainWindow.SubjectNeighborhood.maxSalePrice.ToString() });
             commands.Add(36, new string[] { inputType, "tbSalesPriceAverage", GlobalVar.mainWindow.SubjectNeighborhood.medianSalePrice.ToString() });

             commands.Add(37, new string[] { inputType, "tbPercentageOwned", "90" });
             commands.Add(38, new string[] { inputType, "tbPercentageTenant", "10" });

             commands.Add(39, new string[] { "SELECT", "ddlPropertyValue", "%Increasing" });
             commands.Add(40, new string[] { inputType, "tbPropertyValueRateChange", "1" });
             commands.Add(41, new string[] { "SELECT", "ddlLocation", "%Suburban" });
             commands.Add(42, new string[] { "SELECT", "ddlDemandAndSupply", "%Stable" });
             commands.Add(43, new string[] { "SELECT", "ddlGrowth", "%Stable" });
             commands.Add(44, new string[] { "SELECT", "ddlDemandAndSupply", "%Stable" });
             commands.Add(45, new string[] { "SELECT", "ddlEconomy", "%Stable" });
             commands.Add(46, new string[] { "SELECT", "ddlMarketingTime", "%3-6 mo" });

             commands.Add(47, new string[] { "INPUT:RADIO", "rblIsReoDriven_1", "YES" });
             commands.Add(48, new string[] { "INPUT:RADIO", "rblIsVacant_1", "YES" });
             commands.Add(49, new string[] { "SELECT", "ddlProbablePurchaser", "%2nd Time Buyer" });
             commands.Add(50, new string[] { "SELECT", "ddlFinanceType", "%Conv" });

             commands.Add(51, new string[] { inputType, "tbSalesInImmediateArea", GlobalVar.mainWindow.SubjectNeighborhood.numberSoldListings.ToString() });
             commands.Add(52, new string[] { inputType, "tbListingsInImmediateArea", GlobalVar.mainWindow.SubjectNeighborhood.numberActiveListings.ToString() });
             commands.Add(53, new string[] { inputType, "tbSalesReo", GlobalVar.mainWindow.SubjectNeighborhood.numberREOSales.ToString() });
             commands.Add(54, new string[] { inputType, "tbListingsReo", GlobalVar.mainWindow.SubjectNeighborhood.numberREOListings.ToString() });

             commands.Add(55, new string[] { inputType, "tbEstimatedMonthlyRental", GlobalVar.mainWindow.SubjectRent});
             commands.Add(56, new string[] { inputType, "tbComparableRentalListing", (Math.Round(GlobalVar.mainWindow.SubjectNeighborhood.numberActiveListings * .1)).ToString() });

             commands.Add(57, new string[] { "INPUT:RADIO", "rbIsNewConstruction_1", "YES" });

             commands.Add(58, new string[] { "TEXTAREA", "tbExplainObsolescense", ".TBD." });
             commands.Add(59, new string[] { "TEXTAREA", "tbExplainLocationInfluences", ".TBD." });


             foreach (var c in commands)
             {
                 macro.AppendFormat("TAG POS=1 TYPE={0} FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_{1} CONTENT={2}\r\n", c.Value[0], c.Value[1], c.Value[2].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
             }

             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$btnSaveAndContinue");
             macroCode = macro.ToString();
             iim.iimPlayCode(macroCode, 30);


            #endregion

        }

        public void Prefill(iMacros.App iim, Form1 form)
        {
            this.iim = iim;
            this.callingForm = form;

            switch (helper_GetFormType())
            {
                case bpoEntryFormType.FORM1 :
                    Prefill_Form1();
                    break;
                case bpoEntryFormType.FORM2 :
                    Prefill_Form2();
                    break;
            }
        }

        private bpoEntryFormType helper_GetFormType()
        {
            StringBuilder macro = new StringBuilder();
            int timeout = 60;
            iMacros.Status status = new iMacros.Status();

            macro.AppendLine(@"TAG POS=1 TYPE=SPAN ATTR=CLASS:sectHead EXTRACT=TXT");
            string macroCode = macro.ToString();
            status = iim.iimPlayCode(macroCode, timeout);

            if (status == iMacros.Status.sOk)
            {
                return bpoEntryFormType.FORM2;
            }
            else
            {
                return bpoEntryFormType.FORM1;
            }

           
        }

        private void helper_FillaComp(int compNumber)
        {

        }


        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {
            this.iim = iim;
            this.callingForm = GlobalVar.mainWindow;

            switch (helper_GetFormType())
            {
                case bpoEntryFormType.FORM1:
                    CompFill_1(iim, saleOrList,compNum,fieldList);
                    break;
                case bpoEntryFormType.FORM2:
                    CompFill_2(iim, saleOrList, compNum, fieldList);
                    break;
            }
        }


        private void CompFill_1(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {
            StringBuilder macro = new StringBuilder();
            targetCompNumber = Regex.Match(compNum, @"\d").Value;
            string inputType = "INPUT:TEXT";
            Dictionary<double, string[]> commands = new Dictionary<double, string[]>();

            //macro.AppendLine(@"SET !ERRORIGNORE YES");
            //macro.AppendLine(@"SET !TIMEOUT_STEP 0");







            commands.Add(0, new string[] { inputType, "tbAddress", targetComp.StreetAddress.Replace(" ", "<SP>") });
            commands.Add(1, new string[] { inputType, "tbCity", targetComp.City.Replace(" ", "<SP>") });
            commands.Add(2, new string[] { "SELECT", "ddlState", "%IL" });
            commands.Add(3, new string[] { inputType, "tbZip", targetComp.Zipcode });
            commands.Add(4, new string[] { inputType, "tbMlsNumber", targetComp.MlsNumber });
            commands.Add(4.5, new string[] { inputType, "tbProximity", targetComp.ProximityToSubject.ToString() });
            commands.Add(5, new string[] { inputType, "tbNumUnits", "1" });
           
            if (targetComp.TransactionType == "REO")
            {
                commands.Add(8, new string[] { "SELECT", "ddlSalesType", "%REO" });
            }
            else if (targetComp.TransactionType.ToLower().Contains("short"))
            {
                commands.Add(8, new string[] { "SELECT", "ddlSalesType", "%Short Sale" });
            }
            else
            {
                commands.Add(8, new string[] { "SELECT", "ddlSalesType", "%Fair Market" });
            }

            if (targetComp.FinancingMlsString.ToLower().Contains("cash"))
            {
                commands.Add(9, new string[] { "SELECT", "ddlFinanceType", "%Cash" });
            }
            else if (targetComp.FinancingMlsString.ToLower().Contains("conv"))
            {
                commands.Add(9, new string[] { "SELECT", "ddlFinanceType", "%Conv" });
            }
            else if (targetComp.FinancingMlsString.ToLower().Contains("fha"))
            {
                commands.Add(9, new string[] { "SELECT", "ddlFinanceType", "%FHA" });
            }
            else if (targetComp.FinancingMlsString.ToLower().Contains("va"))
            {
                commands.Add(9, new string[] { "SELECT", "ddlFinanceType", "%VA" });
            }
            else
            {
                commands.Add(9, new string[] { "SELECT", "ddlFinanceType", "%None" });
            }

            if (saleOrList == "sale")
            {
                commands.Add(6, new string[] { inputType, "tbListPriceAtSale", targetComp.CurrentListPrice.ToString() });
                commands.Add(7, new string[] { inputType, "tbSalePrice", targetComp.SalePrice.ToString() });
                //tbSaleDate1
                commands.Add(31, new string[] { inputType, "tbSaleDate", targetComp.SalesDate.ToString() });
            }

            if (saleOrList == "list")
            {
                //tbCurrentListPrice
                commands.Add(5.5, new string[] { inputType, "tbCurrentListPrice", targetComp.CurrentListPrice.ToString() });
                //tbListingDate
                commands.Add(32, new string[] { inputType, "tbListingDate", targetComp.ListDateString});
                //tbListingAgent
                commands.Add(33, new string[] { inputType, "tbListingAgent", targetComp.ListingAgentName });
                //tbListingAgentPhone
                commands.Add(34, new string[] { inputType, "tbListingAgentPhone", targetComp.ListingAgentNamePhone.Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "") });
            }

            commands.Add(10, new string[] { inputType, "tbOriginalListPrice", targetComp.OriginalListPrice.ToString() });
            commands.Add(11, new string[] { inputType, "tbDom", targetComp.DOM });
            commands.Add(12, new string[] { inputType, "tbTtlRoomCount", targetComp.TotalRoomCount.ToString() });
            commands.Add(13, new string[] { inputType, "tbTtlBedrmCount", targetComp.BedroomCount });
            commands.Add(14, new string[] { inputType, "tbTtlBathrmCount", targetComp.BathroomCount });
            commands.Add(15, new string[] { inputType, "tbGla", targetComp.ProperGla(GlobalVar.mainWindow.SubjectAboveGLA) });
            commands.Add(16, new string[] { inputType, "tbLotSize", targetComp.Lotsize.ToString()});
            commands.Add(17, new string[] { inputType, "tbYearBuilt", targetComp.YearBuiltString });
            commands.Add(18, new string[] { inputType, "tbBasementFinished", targetComp.BasementFinishedPercentage() });
            commands.Add(19, new string[] { inputType, "tbBasementSize", targetComp.BasementGLA() });
            commands.Add(20, new string[] { inputType, "tbSellerConcession", targetComp.PointsMlsString });
            commands.Add(21, new string[] { "SELECT", "ddlStyle", "%Conv" });
            commands.Add(22, new string[] { "SELECT", "ddlConstructionType", "%Vinyl" });
            if (GlobalVar.mainWindow.SubjectAttached)
            {
                commands.Add(23, new string[] { "SELECT", "ddlProperyType", "%Attached" });
            }else
            {
                commands.Add(23, new string[] { "SELECT", "ddlProperyType", "%Detached" });
            }
            commands.Add(24, new string[] { "SELECT", "ddlLocation", "%Average" });
            commands.Add(25, new string[] { "SELECT", "ddlView", "%Residential" });
            commands.Add(26, new string[] { "SELECT", "ddlCondition", "%Average" });
            string lresGarageStr = "None";
            string numSpaces = targetComp.NumberGarageStalls();
            string att_det = "";

            if (!string.IsNullOrEmpty(numSpaces))
            {
                if (targetComp.AttachedGarage())
                {
                    att_det = "Attached";
                }
                else if (targetComp.DetachedGarage())
                {
                    att_det = "Detached";
                }

                switch (numSpaces)
                {
                    case "1":
                        lresGarageStr = "1<SP>Car<SP>" + att_det;
                        break;
                    case "2":
                        lresGarageStr = "2<SP>Car<SP>" + att_det;
                        break;
                    default:
                        lresGarageStr = "2+<SP>Car<SP>" + att_det;
                        break;
                }

            }
            commands.Add(27, new string[] { "SELECT", "ddlParking", "%" + lresGarageStr });
            commands.Add(28, new string[] { "SELECT", "ddlComparedToSubject", "%Equal" });

          
            commands.Add(29, new string[] { "TEXTAREA", "tbCompComment", ".TBD." });

       
         

            foreach (var c in commands)
            {
                macro.AppendFormat("TAG POS=1 TYPE={0} FORM=ID:aspnetForm ATTR=ID:*_{2}{1} CONTENT={3}\r\n", c.Value[0], targetCompNumber.ToString(), c.Value[1], c.Value[2].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            }
            //default to the middle until we can suggest which one is the most similar
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$rbMostComparable CONTENT=YES");
        

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 60);



            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$tbAsIsValue CONTENT=111");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$tbRepairedValue CONTENT=111");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$tbSuggestAsIsValue CONTENT=11");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$tbSuggestRepairedValue CONTENT=11");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$tbAsIsQuickSaleValue CONTENT=11");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$ddlVendorLicense CONTENT=%64589");
        






        }

        private void CompFill_2(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {
            StringBuilder macro = new StringBuilder();
            targetCompNumber = Regex.Match(compNum, @"\d").Value;
            string inputType = "INPUT:TEXT";
            Dictionary<double, string[]> commands = new Dictionary<double, string[]>();

            string lresGarageStr = "None";
            string numSpaces = Regex.Match(callingForm.SubjectParkingType, @"\d").Value;
            string att_det = "";

            if (!string.IsNullOrEmpty(numSpaces))
            {
                if (callingForm.SubjectParkingType.ToLower().Contains("att"))
                {
                    att_det = "CA";
                }
                else if (callingForm.SubjectParkingType.ToLower().Contains("det"))
                {
                    att_det = "CD";
                }

                switch (numSpaces)
                {
                    case "1":
                        lresGarageStr = "1" + att_det;
                        break;
                    case "2":
                        lresGarageStr = "2" + att_det;
                        break;
                    default:
                        lresGarageStr = "2+" + att_det;
                        break;
                }

            }



            commands.Add(0, new string[] { "1", inputType, "prop_address", targetComp.StreetAddress.Replace(" ", "<SP>") });
            commands.Add(1, new string[] { "1", inputType, "prop_city", targetComp.City.Replace(" ", "<SP>") });
            commands.Add(2, new string[] { "1", "SELECT", "prop_state", "%IL" });
            commands.Add(3, new string[] { "1", inputType, "prop_zip", targetComp.Zipcode });
            commands.Add(4, new string[] { "1", "SELECT", "property_type", "%SFR" });
            commands.Add(5, new string[] { "1", "SELECT", "style", "%Split-Level" });
            commands.Add(6, new string[] { "1", "SELECT", "ext_walls", "%Frame" });
            commands.Add(7, new string[] { "1", inputType, "sqft_gla", targetComp.ProperGla(GlobalVar.mainWindow.SubjectAboveGLA) });
            commands.Add(8, new string[] { "1", inputType, "total_rooms", targetComp.TotalRoomCount.ToString() });
            commands.Add(9, new string[] { "1", inputType, "bedrooms", targetComp.BedroomCount });
            commands.Add(10, new string[] { "1", inputType, "bathrooms", targetComp.BathroomCount });
            commands.Add(11, new string[] { "1", inputType, "basement_pct_finished", targetComp.BasementFinishedPercentage() });
            commands.Add(12, new string[] { "1", inputType, "basement_sqft", targetComp.BasementGLA()});
            commands.Add(13, new string[] { "1", "SELECT", "garage_carport", "%" + lresGarageStr });
            commands.Add(14, new string[] { "1", inputType, "proximity_to_subj", targetComp.ProximityToSubject.ToString() });
            commands.Add(15, new string[] { "1", inputType, "lot_size", targetComp.Lotsize.ToString() });
            commands.Add(16, new string[] { "1", inputType, "year_built", targetComp.YearBuiltString });
            if (targetComp.TransactionType == "REO")
            {
                commands.Add(17, new string[] { "1", "SELECT", "sale_type", "%REO" });
            }
            else if (targetComp.TransactionType.ToLower().Contains("short"))
            {
                commands.Add(17, new string[] { "1", "SELECT", "sale_type", "%Short Sale" });
            }
            else
            {
                commands.Add(17, new string[] { "1", "SELECT", "sale_type", "%Fair Market" });
            }
            if (targetComp.FinancingMlsString.ToLower().Contains("cash"))
            {
                commands.Add(18, new string[] { "1", "SELECT", "finance_type", "%Cash" });
            }
            else if (targetComp.FinancingMlsString.ToLower().Contains("conv"))
            {
                commands.Add(18, new string[] { "1", "SELECT", "finance_type", "%Conv" });
            }
            else if (targetComp.FinancingMlsString.ToLower().Contains("fha"))
            {
                commands.Add(18, new string[] { "1", "SELECT", "finance_type", "%FHA" });
            }
            else if (targetComp.FinancingMlsString.ToLower().Contains("va"))
            {
                commands.Add(18, new string[] { "1", "SELECT", "finance_type", "%VA" });
            }
            else
            {
                commands.Add(18, new string[] { "1", "SELECT", "finance_type", "%None" });
            }
            commands.Add(19, new string[] { "1", inputType, "seller_concessions", targetComp.PointsMlsString });
            commands.Add(20, new string[] { "1", "SELECT", "condition", "%Average" });
            commands.Add(21, new string[] { "1", "SELECT", "location", "%Average" });
            commands.Add(22, new string[] { "1", "SELECT", "theview", "%Residential" });
        
            commands.Add(24, new string[] { "1", inputType, "days_on_market", targetComp.DOM });
            commands.Add(25, new string[] { "1", inputType, "original_list_price", targetComp.OriginalListPrice.ToString() });
            if (saleOrList == "sale")
            {
                commands.Add(23, new string[] { "1", inputType, "sale_date", targetComp.SalesDate.ToShortDateString() });
                commands.Add(26, new string[] { "1", inputType, "list_price_when_sold", targetComp.CurrentListPrice.ToString() });
                commands.Add(27, new string[] { "1", inputType, "sale_price", targetComp.SalePrice.ToString() });
            }
            if (saleOrList == "list")
            {
                commands.Add(31, new string[] { "1", inputType, "current_list_date", targetComp.ListDate.ToShortDateString() });
                commands.Add(32, new string[] { "1", inputType, "current_list_price", targetComp.CurrentListPrice.ToString() });
                commands.Add(33, new string[] { "1", inputType, "listing_agent_name", targetComp.ListingAgentName });
                commands.Add(34, new string[] { "1", inputType, "listing_agent_phone", targetComp.ListingAgentNamePhone.Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "") });
            }
            commands.Add(28, new string[] { "1", inputType, "legal_description", targetComp.ParcelNumber });
            commands.Add(29, new string[] { "1", inputType, "mls_number", targetComp.MlsNumber });
            commands.Add(30, new string[] { "1", inputType, "adjustments", "...................................TBD..................................."});

            commands.Add(35, new string[] { "1", "SELECT", "compared_to_subject", "%E" });
          //  commands.Add(36, new string[] { "2", "INPUT:RADIO", "closest_comparable", "YES" });
   

      
            //if (saleOrList == "list")
            //{
            //    //tbCurrentListPrice
            //    commands.Add(5.5, new string[] { inputType, "tbCurrentListPrice", targetComp.CurrentListPrice.ToString() });
            //    //tbListingDate
            //    commands.Add(32, new string[] { inputType, "tbListingDate", targetComp.ListDateString });
            //    //tbListingAgent
            //    commands.Add(33, new string[] { inputType, "tbListingAgent", targetComp.ListingAgentName });
            //    //tbListingAgentPhone
            //    commands.Add(34, new string[] { inputType, "tbListingAgentPhone", targetComp.ListingAgentNamePhone.Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "") });
            //}
   
          
            //commands.Add(18, new string[] { inputType, "tbBasementFinished", targetComp.BasementFinishedPercentage() });
            //commands.Add(19, new string[] { inputType, "tbBasementSize", targetComp.BasementGLA() });
            //commands.Add(21, new string[] { "SELECT", "ddlStyle", "%Conv" });
            //commands.Add(22, new string[] { "SELECT", "ddlConstructionType", "%Vinyl" });
            //if (GlobalVar.mainWindow.SubjectAttached)
            //{
            //    commands.Add(23, new string[] { "SELECT", "ddlProperyType", "%Attached" });
            //}
            //else
            //{
            //    commands.Add(23, new string[] { "SELECT", "ddlProperyType", "%Detached" });
            //}
           
       
     
      
            //commands.Add(28, new string[] { "SELECT", "ddlComparedToSubject", "%Equal" });


            //commands.Add(29, new string[] { "TEXTAREA", "tbCompComment", ".TBD." });

            foreach (var c in commands)
            {
                macro.AppendFormat("TAG POS={0} TYPE={1} FORM=NAME:bpoForm ATTR=NAME:*{2}_{3} CONTENT={4}\r\n", c.Value[0], c.Value[1], targetCompNumber.ToString(), c.Value[2], c.Value[3].Replace(" ", "<SP>").Replace("$", ""));
            }
          

            string macroCode = macro.ToString();


            iim.iimPlayCode(macroCode, 60);

   

        }
    }
}
