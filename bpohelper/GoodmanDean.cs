using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{



    class GoodmanDean : BPOFulfillment    
    {

             public GoodmanDean()
        {
            subject = GlobalVar.theSubjectProperty;
            callingForm = GlobalVar.mainWindow;
        }

             public GoodmanDean(MLSListing m) 
            : this()
        {
            targetComp = m;
        }
             private SubjectProperty subject;


        enum elem { TEXT, SELECT, TEXTAREA, RADIO };

         private  Dictionary<elem, string> inputType = new Dictionary<elem, string>()
        {
            {elem.TEXT, "INPUT:TEXT"},{elem.SELECT, "SELECT"}, {elem.TEXTAREA, "TEXTAREA"}, {elem.RADIO, "INPUT:RADIO"}
        };

    
        private Dictionary<int, MLSListing> stack;

        private Form1 callingForm;

       

        public void Prefill(iMacros.App iim, Form1 form)
        {
            //supported form(s) 
            //evalform2
            //bpoform10

            iim.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string currentUrl = iim.iimGetLastExtract();

            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 1");
        


            Dictionary<string, string> fieldList = new Dictionary<string, string>();

            //
            //eval2form
            //
            if (form.SubjectDetached)
            {
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblPropType_0&&VALUE:SFR CONTENT=YES");
            }

            if (form.SubjectAttached)
            {
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblPropType_1&&VALUE:CONDO CONTENT=YES");
            }

            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$rblTransferHistSubjYN CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$txtPropComment CONTENT=No<SP>data.");

            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtZoningDesignation CONTENT=SFR");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblOccupancy_3&&VALUE:UNKNOWN CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlProjectedUse CONTENT=%Residential");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlCurrentUse CONTENT=%Residential");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblZoningCompliance_0&&VALUE:LEGAL CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblPropTransferred_1&&VALUE:N CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblCurrentlyListed_1&&VALUE:N CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblPreviouslyListed_1&&VALUE:N CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_mkNbDensityRbl_1&&VALUE:SUBURBAN CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_mkLocalEconomyRbl_0&&VALUE:OVER<SP>75% CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_mkNeighborHoodValuesRbl_1&&VALUE:STABLE CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblOneUnitTrendsValue_1&&VALUE:STABLE CONTENT=YES");

            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$mkHousingSupplyRbl CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:aspnetForm ATTR=NAME:ctl00$ContentPlaceHolder1$rblOneUnitTrendsMktTime CONTENT=YES");

            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblMarketingTimeTrends_1&&VALUE:STABLE CONTENT=YES");

            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPosFeaCom CONTENT=market<SP>conditions");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtNFCom CONTENT=No<SP>adverse<SP>conditions<SP>observered.");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_chlPropDataSource_1&&VALUE:on CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_chlPropDataSource_2&&VALUE:on CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropLocation CONTENT=Suburban");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropLotSize CONTENT=" + form.SubjectLotSize);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblLotMeasure_1&&VALUE:Acre CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropView CONTENT=Residential");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropYearBuilt CONTENT=" + form.SubjectYearBuilt);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropDesign CONTENT=" + form.SubjectMlsType.Replace(" ", "<SP>"));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropCondition CONTENT=Average");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtTotalRooms CONTENT=" + form.SubjectRoomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropNumBed CONTENT=" + form.SubjectBedroomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropNumBath CONTENT=" + form.SubjectBathroomCount[0]);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropHalfBath CONTENT=" + form.SubjectBathroomCount[2]);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropSqFt CONTENT=" + form.SubjectAboveGLA.Replace(",", ""));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropBasement CONTENT=" + form.SubjectBasementType.Replace(" ", "<SP>") + "-" + form.SubjectBasementDetails.Replace(" ", "<SP>"));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropGarage CONTENT=" + form.SubjectParkingType.Replace(" ","<SP>"));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlPropPool_1&&VALUE:0 CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblPropREO_2&&VALUE:NO CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropSubdivision CONTENT=" + form.SubjectSubdivision.Replace(" ", "<SP>"));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropertyOther");
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtSubjectConditionAffectingValue");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropertyOther CONTENT=NA");
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtSubjectConditionAffectingValue CONTENT=default<SP>additional<SP>comments");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtAgentPrepared");
            //macro.AppendLine(@"TAG POS=482 TYPE=TD FORM=ID:aspnetForm ATTR=*");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtBrokerLicenseState");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtAgentOffice");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtBrokerCompanyAddress");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblAgentInfoVerified_1&&VALUE:N CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblAgentInfoVerified_0&&VALUE:Y CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtBrokerLicenseDate");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_chkSignature&&VALUE:on CONTENT=NO");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtBrokerLicenseDate CONTENT=09/05/2013");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_chkSignature&&VALUE:on CONTENT=YES");

            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtAPN CONTENT=" + form.SubjectPin);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropertyOwner CONTENT=" + form.SubjectOOR.Replace(" ", "<SP>"));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtFairMarketRent CONTENT=" + form.SubjectRent.Replace(",",""));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblAdverseInfluences_1&&VALUE:N CONTENT=YES");
          

            //
            //bpoform10 defaults
            //
            if (currentUrl.ToLower().Contains("bpoform10"))
            {
                //RadioButtons
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblPropNewConstruction_1&&VALUE:N CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblImpactedByDisaster_1&&VALUE:N CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblListedMultiple_1&&VALUE:0 CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblCurrentlyListed_2&&VALUE:NO CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_mkCompetingNewConstructionRbl_1&&VALUE:N CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblPredominantOccupancy_0&&VALUE:OWNER CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_mkNbDensityRbl_1&&VALUE:SUBURBAN CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblOneUnitTrendsValue_1&&VALUE:STABLE CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblREOTrend_1&&VALUE:STABLE CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblVacancy_1&&VALUE:5TO10 CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblEvidenceOfDisaster_1&&VALUE:N CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblIndustrial25mi_1&&VALUE:N CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblPropTransferred_1&&VALUE:NO CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblMarketingTimeTrend_1&&VALUE:STABLE CONTENT=YES");

                //TextFields
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtMonthlyRent CONTENT=1200");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtFairMarketRent CONTENT=" + form.SubjectRent.Replace(",", ""));
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPrsMLS CONTENT=Tax<SP>Records");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropertyUnits CONTENT=1");
                if (form.SubjectBasementType.ToLower().Contains("full"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlPropertyBasementType CONTENT=%Full");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropBsmtSqFt CONTENT=" + form.SubjectBasementGLA);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropBsmtFin CONTENT=" + form.SubjectBasementFinishedGLA);
                }
                else if (form.SubjectBasementType.ToLower().Contains("partial"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlPropertyBasementType CONTENT=%Partial");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropBsmtSqFt CONTENT=" + form.SubjectBasementGLA);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropBsmtFin CONTENT=" + form.SubjectBasementFinishedGLA);
                }
                else
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlPropertyBasementType CONTENT=%None");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropBsmtSqFt CONTENT=0");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropBsmtFin CONTENT=0");
                }
                
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropertyParkingType CONTENT=" + form.SubjectParkingType);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropertyParkingStalls CONTENT=" + Regex.Match(form.SubjectParkingType, @"\d").Value);
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlPropPool CONTENT=%0");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlPropCondition CONTENT=%Average");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropAsIsSalePrice CONTENT=" + form.SubjectMarketValue);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtPropRepairedSalePrice CONTENT=" + form.SubjectMarketValue);


            }

            Dictionary<double, string[]> commands = new Dictionary<double, string[]>();

            commands.Add(0, new string[] { "1", inputType[elem.TEXT], "txtNeighborhoodValuesLow", callingForm.subjectNeighborhood.minSalePrice.ToString() });
            commands.Add(1, new string[] { "1", inputType[elem.TEXT], "txtNeighborhoodValuesHigh", callingForm.subjectNeighborhood.maxSalePrice.ToString() });
            commands.Add(2, new string[] { "1", inputType[elem.TEXT], "txtNeighborhoodValuesPred", callingForm.subjectNeighborhood.medianSalePrice.ToString() });
            commands.Add(3, new string[] { "1", inputType[elem.TEXT], "txtNeighborhoodValuesLowAge", callingForm.subjectNeighborhood.newestHome.ToString() });
            commands.Add(4, new string[] { "1", inputType[elem.TEXT], "txtNeighborhoodValuesHighAge", callingForm.subjectNeighborhood.oldestHome.ToString() });
            commands.Add(5, new string[] { "1", inputType[elem.TEXT], "txtNeighborhoodValuesPredAge", callingForm.subjectNeighborhood.medianAge.ToString() });
            //txtNumberOfListings
            commands.Add(6, new string[] { "1", inputType[elem.TEXT], "txtNumberOfListings", callingForm.cdNumberOfActiveListingTextBox.Text});
            //txtCompValuesLow
            commands.Add(7, new string[] { "1", inputType[elem.TEXT], "txtCompValuesLow", callingForm.cdMinListPriceTextBox.Text });
            //txtCompValuesHigh
            commands.Add(8, new string[] { "1", inputType[elem.TEXT], "txtCompValuesHigh", callingForm.cdMaxListPriceTextBox.Text });
            //txtCompSalesNumber
            commands.Add(9, new string[] { "1", inputType[elem.TEXT], "txtCompSalesNumber", callingForm.cdNumberOfSoldListingTextBox.Text });
            //txtCompSalesValuesLow
            commands.Add(10, new string[] { "1", inputType[elem.TEXT], "txtCompSalesValuesLow", callingForm.cdMinSalePriceTextBox.Text });
            //txtCompSalesValuesHigh
            commands.Add(11, new string[] { "1", inputType[elem.TEXT], "txtCompValuesHigh", callingForm.cdMaxSalePriceTextBox.Text });
            //txtOVCNCom
            commands.Add(12, new string[] { "1", inputType[elem.TEXTAREA], "txtOVCNCom", @"Neighbor stats come from 1 mile radius search.  Comp matches and stats come from up to 5 mile radius." });
            //txtPosFeaCom
            commands.Add(13, new string[] { "1", inputType[elem.TEXTAREA], "txtPosFeaCom", @"Stable market with declining REO and short sales mostly traditional sales. High demand under 250k, weak above 300k, normal between. Seller concessions are not typical for the area." });
            //txtPropertyConcessions
            commands.Add(14, new string[] { "1", inputType[elem.TEXT], "txtPropertyConcessions", "NA" });
            //txtGeneralCom
            commands.Add(15, new string[] { "1", inputType[elem.TEXTAREA], "txtGeneralCom", @"Searched a distance of at least 1 miles, up to 12 months in time. The comps bracket the subject in age, SF, and lot size, as well as use comps in same condition with similar features, and from the subjects market area. All the comps are Reasonable substitute for the subject property, similar in most areas. Price opinion was based off comparable statistics." });
            //txtSubjectConditionAffectingValue
            commands.Add(16, new string[] { "1", inputType[elem.TEXTAREA], "txtSubjectConditionAffectingValue", @"The subject is a conforming home within the neighborhood. No adverse conditions were noted at the time of inspection based on exterior observations. Unable to determine interior condition due to exterior inspection only, so subject was assumed to be in average condition for this report." });
            //txtPropAsIsSalePrice
            commands.Add(17, new string[] { "1", inputType[elem.TEXT], "txtPropAsIsSalePrice", callingForm.SubjectMarketValue });
            //txtBrokerLicenseDate
            commands.Add(18, new string[] { "1", inputType[elem.TEXT], "txtBrokerLicenseDate", callingForm.dateTimePickerInspectionDate.Text });







            foreach (var c in commands)
            {
                macro.AppendFormat("TAG POS={0} TYPE={1} FORM=ID:aspnetForm ATTR=NAME:*{2} CONTENT={4}\r\n", c.Value[0], c.Value[1], c.Value[2], targetCompNumber, c.Value[3].Replace(" ", "<SP>").Replace("$", "").Replace(",", ""));
            }



            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 120);


        }

        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {
            iim.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string currentUrl = iim.iimGetLastExtract();

            string sol;
            if (saleOrList == "sale")
            {
                sol = "Comp";
            }
            else
            {
                sol = "Listing";
            }
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");

            macro.AppendLine(@"SET !TIMEOUT_STEP 0");
            //macro.AppendLine(@"FRAME NAME=pageView");
            macro.AppendLine(@"SET !REPLAYSPEED FAST");

           

            foreach (string field in fieldList.Keys)
            {
                if (field.Contains("*"))
                {
                    //drop down box
                    macro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:*{0}{1} CONTENT=%{2}\r\n", compNum, field.Replace("*", ""), fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                }
                else
                {
                    macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:*{0}{1} CONTENT={2}\r\n", compNum, field, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                }

                //  macro.AppendLine(@"DS CMD=CLICK X={{!TAGX}} Y={{!TAGY}}");
            }

            macro.AppendLine(@"TAG POS=1 TYPE=SPAN ATTR=TXT:close&&CLASS:ui-icon<SP>ui-icon-closethick");
            macro.AppendLine(@"WAIT SECONDS=1");
           

            //text fields with wrong format
           // macro.AppendFormat(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txt{0}Comparison CONTENT={1}", compNum,"Similar");

            //radio buttons
            if (saleOrList == "list")
            {
                macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:*{0}ActivePending_0&&VALUE:L CONTENT={1}\r\n", compNum, "YES");
            }

            if (currentUrl.ToLower().Contains("evalform2"))
            {
                macro.AppendLine(@"TAG POS=2 TYPE=LABEL FORM=ID:aspnetForm ATTR=TXT:MLS");

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_chlCS1DataSource_1&&VALUE:on CONTENT=YES");
                macro.AppendFormat("TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_chl{0}DataSource_0&&VALUE:on CONTENT={1}\r\n", compNum, "YES");
                macro.AppendFormat("TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_chl{0}DataSource_1&&VALUE:on CONTENT={1}\r\n", compNum, "YES");
                macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblLocation{0}_1&&VALUE:SIMILAR CONTENT={1}\r\n", compNum, "YES");
                macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblView{0}_1&&VALUE:SIMILAR CONTENT={1}\r\n", compNum, "YES");
                macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblCond{0}_1&&VALUE:SIMILAR CONTENT={1}\r\n", compNum, "YES");
                macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rbl{0}BasementComp_1&&VALUE:SIMILAR CONTENT={1}\r\n", compNum, "YES");
                macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rbl{0}Garage_1&&VALUE:SIMILAR CONTENT={1}\r\n", compNum, "YES");
                macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblREO{0}_1&&VALUE:SIMILAR CONTENT={1}\r\n", compNum, "YES");
                macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblSubDiv{0}_1&&VALUE:SIMILAR CONTENT={1}\r\n", compNum, "YES");
                macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblOther{0}_1&&VALUE:SIMILAR CONTENT={1}\r\n", compNum, "YES");
                macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblOverall{0}_1&&VALUE:SIMILAR CONTENT={1}\r\n", compNum, "YES");

                macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddl{0}PropPool_1&&VALUE:0 CONTENT={1}\r\n", compNum, "YES");

                //macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblREO{0}_2&&VALUE:{1} CONTENT=YES", compNum, fieldList["SaleType"]);
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblREOCS2_1&&VALUE:SHORTSALE CONTENT=YES");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblREOCS3_0&&VALUE:REO CONTENT=YES");


                //macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblLocation{0}_1&&VALUE:SIMILAR CONTENT={1}\r\n", compNum, "YES");
            }
            //
            //TBD
            //
            //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListOLDate_1");
            //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListCLPrice_1");
            //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListCLDate_1");

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 60);
        }
    }
}
