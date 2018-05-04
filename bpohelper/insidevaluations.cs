using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{

    class InsideValuation:BPOFulfillment
    {

         private Dictionary<string, MacroParameterValues> NeighborhoodFieldList = new Dictionary<string, MacroParameterValues>();
         public enum ImacroType { CHECKBOX, TEXT, SELECT}
         public Dictionary<ImacroType, string> imacroTypeToStringAdapter = new Dictionary<ImacroType, string> {
            {ImacroType.CHECKBOX, "INPUT:CHECKBOX"},
            {ImacroType.SELECT, "SELECT"},
            {ImacroType.TEXT, "INPUT:TEXT"}
         };

        struct MacroParameterValues
        {
            string imacroCommandType;
        }

      

        struct WebFormSelectionBoxes
        {
            public enum DataSourceSelections { MLS = 1, County = 2, Owner = 3, Broker = 4, CoStar = 5 }

        }

 
       private string theBaseImacroCommand = "TAG POS={0} TYPE={1} FORM=ACTION:PWSEdit* ATTR=NAME:*{2}{3}{4}{5} CONTENT={6}\r\n";

          public InsideValuation()
        {

        }


          public InsideValuation(MLSListing m) 
            : this()
        {
            targetComp = m;
        }
        
        
     

        private Dictionary<string, string> subjectFieldListTranslator = new Dictionary<string, string>()
        {
            {"ParcelID", "Apn"}, 
            {"County", "County"},
            {"PropertyType", "*property_type"},
            {"Rent", "fair_mkt_rent"},
            {"AboveGLA", "GBASqFt"},
            {"DataSource", "SqFtDatasource"}
            //{"DR", "Drive"},
            //{"HWY", "Highway"},
            //{"LN", "Lane"},
            //{"PKWY", "Parkway"},
            //{"PL", "Place"},
            //{"PLZ", "Plaza"},
            //{"PL", "Point"},
            //{"PT", "Place"},
            //{"RD", "Road"},
            //{"SQ", "Square"},
            //{"ST", "Road"},
            //{"TER", "Terrace"},
            //{"TRL", "Trail"},
            //{"WAY", "Way"}
           
        };

         private Dictionary<string, string> subjectFieldList = new Dictionary<string, string>();


    

      
        protected  new void   GenerateSubjectFillScript()
        {
            ///
            //5/1/18 updates
            ///

            //TextBoxOutBuildingsDesc
            subjectTextFieldList.Add("TextBoxOutBuildingsDesc", "NA");

            //TextBoxNegativeExternailities
            subjectTextFieldList.Add("TextBoxNegativeExternailities", "No negative externalities noted by drive by.");

            //TextBoxPositiveExternailities
            subjectTextFieldList.Add("TextBoxPositiveExternailities", "No positive externalities noted by drive by.");

            //TextBoxSubjectNeighborhoodComments
            subjectTextFieldList.Add("TextBoxSubjectNeighborhoodComments", "Typical suburban subdivision.");

            //DropDownListCompFinancingType
            subjectSelectionBoxList.Add("DropDownListCompFinancingType", "24");

            //TextBoxUnitVar
            subjectTextFieldList.Add("TextBoxUnitVar", "1");

            //TextBoxYearVar
            subjectTextFieldList.Add("TextBoxYearVar", "100");

            //TextBoxLocationVar
            subjectTextFieldList.Add("TextBoxLocationVar", "1");

            //TextBoxConditionVar
            subjectTextFieldList.Add("TextBoxConditionVar", "1");

            //TextBoxBathVar
            subjectTextFieldList.Add("TextBoxBathVar", "1000");

            //TextBoxBedVar
            subjectTextFieldList.Add("TextBoxBedVar", "1500");

            //TextBoxHalfBathVar
            subjectTextFieldList.Add("TextBoxHalfBathVar", "500");

            //TextBoxGarageAttachVar
            subjectTextFieldList.Add("TextBoxGarageAttachVar", "1000");

            //TextBoxGarageDetachVar
            subjectTextFieldList.Add("TextBoxGarageDetachVar", "500");

            //TextBoxGarageCarportVar
            subjectTextFieldList.Add("TextBoxGarageCarportVar", "1");

            //TextBoxPorchPatioDeck
            subjectTextFieldList.Add("TextBoxPorchPatioDeck", "NA");

            //TextBoxPoolVar
            subjectTextFieldList.Add("TextBoxPoolVar", "1");

            //TextBoxSpaVar
            subjectTextFieldList.Add("TextBoxSpaVar", "1");


            subjectTextFieldList.Add("TextBoxDateOfInspection", GlobalVar.theSubjectProperty.InspectionDate().ToString());
            subjectTextFieldList.Add("TextBoxApn", GlobalVar.theSubjectProperty.ParcelID);
            subjectTextFieldList.Add("TextBoxGBASqFt", GlobalVar.mainWindow.SubjectAboveGLA);
            subjectTextFieldList.Add("TextBoxYearBuilt", GlobalVar.mainWindow.SubjectYearBuilt);
            subjectSelectionBoxList.Add("DropDownListSqFtDatasource", "2");
            subjectSelectionBoxList.Add("DropDownListResPropertyType", "1");
            subjectSelectionBoxList.Add("DropDownListUnitsNo", "1");
            subjectSelectionBoxList.Add("DropDownListVacantUnits", "0");
            subjectTextFieldList.Add("TextBoxSiteAcres", GlobalVar.mainWindow.SubjectLotSize);
            subjectSelectionBoxList.Add("DropDownListPropertyCondition", "9");


            //TextBoxAssessorsValue
            subjectTextFieldList.Add("TextBoxAssessorsValue", GlobalVar.mainWindow.SubjectAssessmentValue);
            //TextBoxMonthlyRentValue
            subjectTextFieldList.Add("TextBoxMonthlyRentValue", GlobalVar.mainWindow.SubjectRent);
            //TextBoxLandValue
            subjectTextFieldList.Add("TextBoxLandValue", GlobalVar.mainWindow.SubjectLandValue);
            //TextBoxTaxes
            subjectTextFieldList.Add("TextBoxTaxes", GlobalVar.theSubjectProperty.PropertyTax);
            //DropDownListResPropertyZoneType
            subjectSelectionBoxList.Add("DropDownListResPropertyZoneType", "1");
            //TextBoxSchoolDistrict
            subjectTextFieldList.Add("TextBoxSchoolDistrict", GlobalVar.mainWindow.SubjectSchoolDistrict);
            //TextBoxSalesPriceLow
            subjectTextFieldList.Add("TextBoxSalesPriceLow", GlobalVar.mainWindow.SubjectNeighborhood.minSalePrice.ToString());
            //TextBoxSalesPriceHigh
            subjectTextFieldList.Add("TextBoxSalesPriceHigh", GlobalVar.mainWindow.SubjectNeighborhood.maxSalePrice.ToString());
            //TextBoxSalesPriceAvg
            subjectTextFieldList.Add("TextBoxSalesPriceAvg", GlobalVar.mainWindow.SubjectNeighborhood.medianSalePrice.ToString());
            //TextBoxListingPriceLow
            subjectTextFieldList.Add("TextBoxListingPriceLow", GlobalVar.mainWindow.SubjectNeighborhood.minListPrice.ToString());
            //TextBoxListingPriceHigh
            subjectTextFieldList.Add("TextBoxListingPriceHigh", GlobalVar.mainWindow.SubjectNeighborhood.maxListPrice.ToString());
            //TextBoxListingPriceAvg
            subjectTextFieldList.Add("TextBoxListingPriceAvg", GlobalVar.mainWindow.SubjectNeighborhood.medianListPrice.ToString());
            //DropDownListNeighborhoodPropertyCondition
            subjectSelectionBoxList.Add("DropDownListNeighborhoodPropertyCondition", "9");
            //DropDownListNeighborhoodPropertyConformity
            subjectSelectionBoxList.Add("DropDownListNeighborhoodPropertyConformity", "1");
            //TextBoxOwnerOccupantPct
            subjectTextFieldList.Add("TextBoxOwnerOccupantPct", "80");
            //TextBoxTenantOccupantPct
            subjectTextFieldList.Add("TextBoxTenantOccupantPct", "10");
            //TextBoxVacancyPct
            subjectTextFieldList.Add("TextBoxVacancyPct", "10");
            //TextBoxNumberOfCompetingListings
            subjectTextFieldList.Add("TextBoxNumberOfCompetingListings", GlobalVar.mainWindow.SubjectNeighborhood.numberOfCompListings.ToString());
            //TextBoxNumOfReoListings
            subjectTextFieldList.Add("TextBoxNumOfReoListings", GlobalVar.mainWindow.SubjectNeighborhood.numberREOListings.ToString());
            //TextBoxNumberOfCompetingSales
            subjectTextFieldList.Add("TextBoxNumberOfCompetingSales", GlobalVar.mainWindow.cdNumberOfSoldListingTextBox.Text);
            //TextBoxNumOfReoSales
            subjectTextFieldList.Add("TextBoxNumOfReoSales", GlobalVar.mainWindow.SubjectNeighborhood.numberREOSales.ToString());
            //TextBoxSubjectStyleDesign
            subjectTextFieldList.Add("TextBoxSubjectStyleDesign", GlobalVar.mainWindow.SubjectMlsType);           
            //DropDownListLocationSubject
            subjectSelectionBoxList.Add("DropDownListLocationSubject", "9");           
            subjectTextFieldList.Add("StyleDesignDesc", GlobalVar.mainWindow.SubjectMlsType);
            subjectTextFieldList.Add("TextBoxStoryNo", "1");                 
            subjectSelectionBoxList.Add("ListLocation", "9");
            subjectTextFieldList.Add("TextBoxViewDesc", "Residential");
            subjectTextFieldList.Add("TextBoxTotalRoom", GlobalVar.mainWindow.SubjectRoomCount);
            subjectTextFieldList.Add("TextBoxBed", GlobalVar.mainWindow.SubjectBedroomCount);
            subjectTextFieldList.Add("TextBoxBath", GlobalVar.mainWindow.SubjectBathroomCount[0].ToString());
            subjectTextFieldList.Add("TextBoxHalfBath", GlobalVar.mainWindow.SubjectBathroomCount[2].ToString());
            subjectTextFieldList.Add("TextBoxGBA", GlobalVar.mainWindow.SubjectAboveGLA);
            subjectTextFieldList.Add("TextBoxBasementSqFt", GlobalVar.mainWindow.SubjectBasementGLA);
            subjectTextFieldList.Add("TextBoxBasementPercentFinished", GlobalVar.mainWindow.SubjectBasementFinishedGLA);
            subjectSelectionBoxList.Add("DropDownListGarageStallNo", helper_SubjectGarageStallsTranslate());

            //TextBoxEstimatedMarketValue
            subjectTextFieldList.Add("TextBoxEstimatedMarketValue", GlobalVar.mainWindow.SubjectMarketValue);
            //TextBoxAsRepairedValue
            subjectTextFieldList.Add("TextBoxAsRepairedValue", GlobalVar.mainWindow.SubjectMarketValue);
            //TextBoxAsIs30DayValue
            subjectTextFieldList.Add("TextBoxAsIs30DayValue", GlobalVar.mainWindow.SubjectQuickSaleValue);
            //TextBox90DayValue
            subjectTextFieldList.Add("TextBox90DayValue", GlobalVar.mainWindow.SubjectMarketValue);
            //TextBoxSuggestedListAsIsValue
            subjectTextFieldList.Add("TextBoxSuggestedListAsIsValue", GlobalVar.mainWindow.SubjectMarketValueList);
       
            if (GlobalVar.mainWindow.SubjectParkingType.Contains("det"))
            {
                subjectSelectionBoxList.Add("DropDownListGarageType", "3");
            }
            else if (GlobalVar.mainWindow.SubjectParkingType.Contains("att"))
            {
                subjectSelectionBoxList.Add("DropDownListGarageType", "2");
            }
            
            foreach (string field in subjectTextFieldList.Keys)
            {
                theMacro.AppendFormat(theBaseImacroCommand, "1", "INPUT:TEXT", "", "", field, "", subjectTextFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }



            foreach (string field in subjectSelectionBoxList.Keys)
            {
                theMacro.AppendFormat(theBaseImacroCommand, "1", "SELECT", "", "", field, "", "%" + subjectSelectionBoxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

 








        }

        protected void GenerateCompFillScript(string webformCompName)
        {
            compTextFieldList.Add("Address",targetComp.StreetAddress);
            compTextFieldList.Add("City",  targetComp.City);
            compSelectionBoxList.Add("State", "IL");
            compTextFieldList.Add("Zip", targetComp.Zipcode);
            compTextFieldList.Add("TextBoxPrice", targetComp.CurrentListPrice.ToString());
       
            compTextFieldList.Add("OriginalListPrice", targetComp.OriginalListPrice.ToString());
            compTextFieldList.Add("OriginalListDate", targetComp.ListDateString);
            compTextFieldList.Add("LastReductionDate", targetComp.DateOfLastPriceChange.ToShortDateString());
            compSelectionBoxList.Add("Datasource", "1");
            compTextFieldList.Add("TextBoxMLSNum", targetComp.MlsNumber);
            compSelectionBoxList.Add("CompFinancingType", helper_TransactionTypeTranslate());
            compTextFieldList.Add("TextBoxSchoolDistrict", targetComp.SchoolDistrict);
            compSelectionBoxList.Add("PropertyTypeId", helper_PropertyTypeTranslate());
            compTextFieldList.Add("StyleDesignDesc", targetComp.Type);
            compTextFieldList.Add("TextBoxUnitNo", "1");
            compTextFieldList.Add("TextBoxStoryNo", targetComp.Levels());
            compTextFieldList.Add("TextBoxNumOutbuildings", "0");
            compTextFieldList.Add("TextBoxYearBuilt", targetComp.YearBuiltString);
            compTextFieldList.Add("TextBoxSiteAcres", Math.Round(targetComp.Lotsize, 2).ToString());
            compSelectionBoxList.Add("ListLocation", "9");
            compTextFieldList.Add("TextBoxViewDesc", "Residential");
            compSelectionBoxList.Add("CompCondition", "9");
            compTextFieldList.Add("TextBoxTotalRoom", targetComp.TotalRoomCount.ToString());
            compTextFieldList.Add("TextBoxBed", targetComp.BedroomCount);
            compTextFieldList.Add("TextBoxBath", targetComp.FullBathCount.ToString());
            compTextFieldList.Add("TextBoxHalfBath", targetComp.HalfBathCount.ToString());
            compTextFieldList.Add("TextBoxGBA", targetComp.GLA.ToString());
            compTextFieldList.Add("TextBoxBasementSqFt", targetComp.BasementGLA());
            compTextFieldList.Add("TextBoxBasementPercentFinished", targetComp.BasementFinishedPercentage());
            compSelectionBoxList.Add("DropDownListGarageStallNo",helper_GarageStallsTranslate());
    
            compSelectionBoxList.Add("DropDownListGarageType", helper_GarageTypeTranslate());

            //TextBoxPorchPatioDeck
            compTextFieldList.Add("TextBoxPorchPatioDeck", "NA");

            //TextBoxListingComments
            //TextBoxRecentSaleComments
            compTextAreaList.Add("TextBoxListingComments", "Similar size, age, style, features, condition and neighborhood as subject.");
            compTextAreaList.Add("TextBoxRecentSaleComments", "Similar size, age, style, features, condition and neighborhood as subject.");


            if (saleOrListFlag == "sale")
            {
                //TextBoxSalePrice
                compTextFieldList.Add("TextBoxSalePrice", targetComp.SalePrice.ToString());

                //TextBoxSaleSoldDate5
                compTextFieldList.Add("TextBoxSaleSoldDate", targetComp.SalesDate.ToShortDateString());
                compTextFieldList.Add("TextBoxSaleDom", targetComp.DOM);
            }
           



            //compTextFieldList.Add("ProximityToSubject", targetComp.ProximityToSubject.ToString());
            //compTextFieldList.Add("SourceOfFunds", targetComp.FinancingMlsString);
            //compTextFieldList.Add("SalesProperties.CurrentListing.Concessions", targetComp.PointsMlsString);

            //compTextFieldList.Add("GeneralProperties.IsHOA.feesPerMonth", "0");

            //compTextFieldList.Add("LandAndStructure.Age", targetComp.Age.ToString());
            //compTextFieldList.Add("LandAndStructure.Site", "1");
       
           
            //compTextFieldList.Add("LandAndStructure.DesignStyle", "Typical");
            
            //compTextFieldList.Add("ComparisonToSubject", "Same");


            //compTextFieldList.Add("LandAndStructure.SquareFeet.Finished", targetComp.BasementFinishedGLA());
            //compTextFieldList.Add("Garage", targetComp.GarageString());
            
            //compTextFieldList.Add("Amenities.Pool", "NA");
            //compTextFieldList.Add("Amenities.Other.type", "NA");
            //compTextFieldList.Add("LandAndStructure.OverallComparability", "Equal");




            theMacro.Clear();
            theMacro.AppendLine(@"SET !ERRORIGNORE YES");
            theMacro.AppendLine(@"SET !TIMEOUT_STEP 0");
            theMacro.AppendLine(@"SET !REPLAYSPEED FAST");


            foreach (string field in compTextFieldList.Keys)
            {
                theMacro.AppendFormat(theBaseImacroCommand, "1", "INPUT:TEXT", "", "", field, webformCompName, compTextFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

     

            foreach (string field in compSelectionBoxList.Keys)
            {
                theMacro.AppendFormat(theBaseImacroCommand, "1", "SELECT", "", "", field, webformCompName, "%" + compSelectionBoxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compTextAreaList.Keys)
            {
                theMacro.AppendFormat(theBaseImacroCommand, "1", "TEXTAREA", "", "", field, webformCompName, "" + compTextAreaList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

        }

        private string helper_SubjectGarageStallsTranslate()
        {
            //default if unknown
            string theirType = "";
            if (GlobalVar.theSubjectProperty.GarageStallCount.Contains("1"))
            {
                theirType = "2";
            }
            else if (GlobalVar.theSubjectProperty.GarageStallCount.Contains("2"))
            {
                theirType = "3";
            }
            else if (GlobalVar.theSubjectProperty.GarageStallCount.Contains("3"))
            {
                theirType = "4";
            }
            else if (GlobalVar.theSubjectProperty.GarageStallCount.Contains("0"))
            {
                theirType = "1";
            }

           

            return theirType;
        }

        private string helper_GarageTypeTranslate()
        {
            //default if unknown
            string theirType = "1";
            if (targetComp.AttachedGarage())
            {
                theirType = "2";
            }
            else if (targetComp.DetachedGarage())
            {
                theirType = "3";
            }
           
            return theirType;
        }

        private string helper_GarageStallsTranslate()
        {
            //default if unknown
            string theirType = "";
            if (targetComp.NumberGarageStalls().Contains("1"))
            {
                theirType = "2";
            }
            else if (targetComp.NumberGarageStalls().Contains("2"))
            {
                theirType = "3";
            }
            else if (targetComp.NumberGarageStalls().Contains("3"))
            {
                theirType = "4";
            }
            else if (targetComp.NumberGarageStalls().Contains("0"))
            {
                theirType = "1";
            }



            return theirType;
        }

        private string helper_PropertyTypeTranslate()
        {
            //default if unknown
            string theirType = "";
            if (targetComp.ToString().Contains("Detached"))
            {
                theirType = "1";
            }
            else if (targetComp.ToString().Contains("Attached"))
            {
                theirType = "2";
            }
            return theirType;
        }

        private string helper_TransactionTypeTranslate()
        {
            //default if unknown
            string theirType = "24";
           if (targetComp.TransactionType.Contains("REO"))
           {
               theirType = "13";
           } else if (targetComp.TransactionType.Contains("Short"))
           {
               theirType = "23";
           }
           return theirType;
        }

        public void Prefill(iMacros.App iim, Form1 form)
        {
            theMacro.Clear();
            theMacro.AppendLine(@"SET !ERRORIGNORE YES");
            theMacro.AppendLine(@"SET !TIMEOUT_STEP 0");
            theMacro.AppendLine(@"SET !REPLAYSPEED FAST");
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$RadioButtonListOccupancy CONTENT=YES");
            theMacro.AppendLine(@"TAG POS=3 TYPE=INPUT:RADIO ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$RadioButtonListImprovementComparedToNeighId CONTENT=YES");
            theMacro.AppendLine(@"TAG POS=3 TYPE=INPUT:RADIO ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$RadioButtonListlocalEconomyId CONTENT=YES");
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$RadioButtonListEmploymentRateChangeId CONTENT=YES");
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$RadioButtonListsalesPriceChangeId CONTENT=YES");
            theMacro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$RadioButtonListNeighborhoodMarketingDays CONTENT=YES");
            theMacro.AppendLine(@"TAG POS=3 TYPE=INPUT:RADIO ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$RadioButtonListNeighborhoodListingSupplyId CONTENT=YES");
            theMacro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$RadioButtonListREODrivenMarketFlag CONTENT=YES");
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$RadioButtonListFinancingAvailFlag CONTENT=YES");
            theMacro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$RadioButtonListHOA CONTENT=YES");
            theMacro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$RadioButtonListListedLast12Months CONTENT=YES");
            theMacro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$RadioButtonListCurrentlyListed CONTENT=YES");
            theMacro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$RadioButtonListSoldLast5Years CONTENT=YES");
            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$RadioButtonListlistingPctChangeId CONTENT=YES");
            string subComments = "The subject is a conforming home within the neighborhood. No adverse conditions were noted at the time of inspection based on exterior observations. Unable to determine interior condition due to exterior inspection only, so subject was assumed to be in average condition for this report.";

            theMacro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$TextBoxSubjectComments CONTENT=" + subComments.Replace(" ", "<SP>"));

            GenerateSubjectFillScript();

            WriteScript(GlobalVar.theSubjectProperty.MainForm.SubjectFilePath, "prefill.iim", theMacro);

            string macroCode = theMacro.ToString();
            iim.iimPlayCode(macroCode, 60);
        }

        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {
            saleOrListFlag = saleOrList;
            GenerateCompFillScript(compNum);
            WriteScript(fieldList["filepath"], compNum + ".iim", theMacro);

            string macroCode = theMacro.ToString();
            iim.iimPlayCode(macroCode, 60);
        }

        public void UploadPics(iMacros.App iim)
        {
            StringBuilder macro = new StringBuilder();



            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:PropertyWorkShopImages.aspx?SubjectProjectId=455624 ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$DataListNewImages$ctl01$DropDownListImageLabel CONTENT=%1");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:PropertyWorkShopImages.aspx?SubjectProjectId=455624 ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$DataListNewImages$ctl01$DropDownListImageView CONTENT=%1");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ACTION:PropertyWorkShopImages.aspx?SubjectProjectId=455624 ATTR=NAME:NewImageProminent CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:PropertyWorkShopImages.aspx?SubjectProjectId=455624 ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$DataListNewImages$ctl01$FileUpload1 CONTENT=C:\fakepath\Photo<SP>May<SP>03,<SP>9<SP>18<SP>15<SP>AM.jpg");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:PropertyWorkShopImages.aspx?SubjectProjectId=455624 ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$DataListNewImages$ctl02$DropDownListImageLabel CONTENT=%1");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:PropertyWorkShopImages.aspx?SubjectProjectId=455624 ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$DataListNewImages$ctl02$DropDownListImageView CONTENT=%2");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:PropertyWorkShopImages.aspx?SubjectProjectId=455624 ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$DataListNewImages$ctl02$FileUpload1 CONTENT=C:\fakepath\Photo<SP>May<SP>03,<SP>9<SP>18<SP>21<SP>AM.jpg");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:PropertyWorkShopImages.aspx?SubjectProjectId=455624 ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$DataListNewImages$ctl03$DropDownListImageLabel CONTENT=%1");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:PropertyWorkShopImages.aspx?SubjectProjectId=455624 ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$DataListNewImages$ctl03$DropDownListImageView CONTENT=%7");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:PropertyWorkShopImages.aspx?SubjectProjectId=455624 ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$DataListNewImages$ctl03$FileUpload1 CONTENT=C:\fakepath\Photo<SP>May<SP>03,<SP>9<SP>18<SP>04<SP>AM.jpg");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:PropertyWorkShopImages.aspx?SubjectProjectId=455624 ATTR=NAME:ctl00$ContentPlaceHolderWorkShopPage$ButtonUploadImages");
            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 60);
        }









            
    }
}
