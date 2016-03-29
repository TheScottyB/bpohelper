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
       

            if (GlobalVar.mainWindow.SubjectParkingType.Contains("det"))
            {
                subjectSelectionBoxList.Add("DropDownListGarageType", "3");
            }
            else if (GlobalVar.mainWindow.SubjectParkingType.Contains("att"))
            {
                subjectSelectionBoxList.Add("DropDownListGarageType", "2");
            }
            




            //subjectTextFieldList.Add("ProximityToSubject", targetComp.ProximityToSubject.ToString());
            //subjectTextFieldList.Add("SourceOfFunds", targetComp.FinancingMlsString);
            //subjectTextFieldList.Add("SalesProperties.CurrentListing.Concessions", targetComp.PointsMlsString);

            //subjectTextFieldList.Add("GeneralProperties.IsHOA.feesPerMonth", "0");

            //subjectTextFieldList.Add("LandAndStructure.Age", targetComp.Age.ToString());
            //subjectTextFieldList.Add("LandAndStructure.Site", "1");


            //subjectTextFieldList.Add("LandAndStructure.DesignStyle", "Typical");

            //subjectTextFieldList.Add("ComparisonToSubject", "Same");


            //subjectTextFieldList.Add("LandAndStructure.SquareFeet.Finished", targetComp.BasementFinishedGLA());
            //subjectTextFieldList.Add("Garage", targetComp.GarageString());

            //subjectTextFieldList.Add("Amenities.Pool", "NA");
            //subjectTextFieldList.Add("Amenities.Other.type", "NA");
            //subjectTextFieldList.Add("LandAndStructure.OverallComparability", "Equal");






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
            compTextFieldList.Add("Dom", targetComp.DOM);
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
            compTextFieldList.Add("TextBoxSiteAcres", targetComp.Lotsize.ToString());
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

            GenerateSubjectFillScript();




            //subjectFieldList.Add(subjectFieldListTranslator["ParcelID"], GlobalVar.theSubjectProperty.ParcelID);
            //subjectFieldList.Add(subjectFieldListTranslator["County"], GlobalVar.theSubjectProperty.County);
            //subjectFieldList.Add(subjectFieldListTranslator["AboveGLA"], GlobalVar.theSubjectProperty.MainForm.SubjectAboveGLA);


            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.OwnerOfPublicRecords), GlobalVar.theSubjectProperty.MainForm.SubjectOOR);
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.Groups.LivingArea, WebFormFieldNames.LivingAreaFields.Gross), GlobalVar.theSubjectProperty.MainForm.SubjectAboveGLA);
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.Age ), GlobalVar.theSubjectProperty.Age.ToString());
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.LotSize), GlobalVar.theSubjectProperty.MainForm.SubjectLotSize);
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.Site), "1");
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.TotalRooms), GlobalVar.theSubjectProperty.MainForm.SubjectRoomCount);
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.BedRooms), GlobalVar.theSubjectProperty.MainForm.SubjectBedroomCount);
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.FullRooms), GlobalVar.theSubjectProperty.FullBathCount);
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.PartialRooms), GlobalVar.theSubjectProperty.HalfBathCount);
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.DesignStyle), "Typical");
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.View), "Residential");
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.NumberOfUnits), "1");
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.Groups.Basement, WebFormFieldNames.BasementFields.TotalArea), GlobalVar.theSubjectProperty.MainForm.SubjectBasementGLA);
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.Groups.SquareFeet, WebFormFieldNames.SquareFeetFields.Finished), GlobalVar.theSubjectProperty.MainForm.SubjectBasementFinishedGLA);
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Garage), GlobalVar.theSubjectProperty.MainForm.SubjectParkingType);
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.ParkingStalls), GlobalVar.theSubjectProperty.GarageStallCount);
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.Amenities, WebFormFieldNames.AmenitiesFields.Pool), "Unk");
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.Amenities, WebFormFieldNames.Groups.Other, WebFormFieldNames.OtherFields.type), "Unk");
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.InspectionDate), GlobalVar.theSubjectProperty.InspectionDate());
            ////subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.LandAndStructure, WebFormFieldNames.LandAndStructureFields.SourceOfFunds), "NA");
            ////if (!String.IsNullOrWhiteSpace(GlobalVar.theSubjectProperty.MainForm.SubjectAssessmentInfo.amount))
            ////{
            ////    subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.GeneralProperties, WebFormFieldNames.GeneralPropertiesFields.HOA, WebFormFieldNames.HOAFields.Fee, WebFormFieldNames.FeeFields.Amount ), GlobalVar.theSubjectProperty.MainForm.SubjectAssessmentInfo.amount);
            ////}

            ////      private string theBaseImacroCommand = "TAG POS={0} TYPE={1} FORM=ACTION:PWSEdit* ATTR=NAME:*{2}{3}{4}{5} CONTENT={6}\r\n";

            //foreach (string field in subjectFieldList.Keys)
            //{
            //    theMacro.AppendFormat(theBaseImacroCommand, "1", "INPUT:TEXT", "", "", "", field, subjectFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            //}

           // subjectFieldList.Clear();
           // int x = (int)WebFormSelectionBoxes.DataSourceSelections.MLS;
           // subjectFieldList.Add(subjectFieldListTranslator["DataSource"], x.ToString());
           //// subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Condition), x.ToString());

           // foreach (string field in subjectFieldList.Keys)
           // {
           //     theMacro.AppendFormat(theBaseImacroCommand, "1", "SELECT", "", "", "", field.Replace("_", "<SP>"), "%" + subjectFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
           // }

            //subjectFieldList.Clear();
            //if (GlobalVar.theSubjectProperty.MainForm.SubjectAttached)
            //{
            //    subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.PropertyType), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.PropertyTypes.Condo));
            //    theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.NormalSale.Repaired.SubjectLandValue CONTENT=0");
     

            //}
            //else
            //{
            //    subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.PropertyType), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.PropertyTypes.SFR));
            //    theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.NormalSale.Repaired.SubjectLandValue CONTENT=10000");
            //}
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Construction), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.YesNo.No));
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.ImpactedByDiaster), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.YesNo.No));
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Occupancy), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.OccupancyTypes.Unknown));
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Inspection), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.InspectionTypes.Exterior));
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.CurrentListing), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.CurrentListingTypes.No));
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.SalesProperties, WebFormFieldNames.Groups.PreviousListings, WebFormFieldNames.PreviousListingsFields.PreviouslyListed), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.YesNo.No));
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Location), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.LocationTypes.Suburban));
            //foreach (string field in subjectFieldList.Keys)
            //{
            //    theMacro.AppendFormat(theBaseImacroCommand, subjectFieldList[field], "INPUT:CHECKBOX", "Subject", "", "", field.Replace("_", "<SP>"), "YES");
            //}

            ////
            ////Neighborhood Checkboxes
            ////
            //subjectFieldList.Clear();
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.Predominant, WebFormFieldNames.PredominantFields.PriceTrend, WebFormFieldNames.PriceTrendFields.Direction), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.PropertyValues.Stable));
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.Groups.Predominant, WebFormFieldNames.PredominantFields.Occupancy), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.PredominantOccupancy.Owner));
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Industrydistance), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.YesNo.No));
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.Vacancy), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.VacancyTypes.percent5_10));
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.NewConstruction), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.YesNo.No));
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.EvidenceForDiaster), helper_CorrectCheckBoxTagNumber(WebCheckBoxGroups.YesNo.No));
            //foreach (string field in subjectFieldList.Keys)
            //{
            //    theMacro.AppendFormat(theBaseImacroCommand, subjectFieldList[field], "INPUT:CHECKBOX", "Neighborhood", "", "", field.Replace("_", "<SP>"), "YES");
            //}

            ////
            ////Neighborhood TEXT
            ////
            //subjectFieldList.Clear();
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.MedianMarketRent), GlobalVar.theSubjectProperty.MainForm.SubjectRent);
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.REOPercentage), "Stable");
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.MarketTimingTrend), "Stable");
            //foreach (string field in subjectFieldList.Keys)
            //{
            //    theMacro.AppendFormat(theBaseImacroCommand, "1", "INPUT:TEXT", "Neighborhood", "", "", field, subjectFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            //}

            ////
            ////Neighborhood SELECT
            ////
            //subjectFieldList.Clear();
            //subjectFieldList.Add(helper_MakeCommandString(WebFormFieldNames.CommonFields.AverageMarketTime), "0-3 Months");
            //foreach (string field in subjectFieldList.Keys)
            //{
            //    theMacro.AppendFormat(theBaseImacroCommand, "1", "SELECT", "Neighborhood", "", "", field.Replace("_", "<SP>"), "%" + subjectFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            //}

            ////
            ////Neighborhood stats
            ////
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Neighborhood.Listings.NeighborhoodListings CONTENT=" + form.SubjectNeighborhood.numberActiveListings);
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Neighborhood.Predominant.PricesFrom CONTENT=" + form.SubjectNeighborhood.minListPrice);
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Neighborhood.Predominant.PricesTo CONTENT=" + form.SubjectNeighborhood.highListPrice);
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Neighborhood.ComparableListingSupply CONTENT=" + form.SubjectNeighborhood.numberOfSales);
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Neighborhood.Construction.LowPrice CONTENT=" + form.SubjectNeighborhood.minSalePrice);
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Neighborhood.Construction.HighPrice CONTENT=" + form.SubjectNeighborhood.maxSalePrice);


            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.NormalSale.ASIS.MarketValue CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectMarketValue);
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.NormalSale.ASIS.SuggestedListPrice CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectMarketValueList);
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.NormalSale.Repaired.MarketValue CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectMarketValue);
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.NormalSale.Repaired.SuggestedListPrice CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectMarketValueList);

            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.QuickSale.ASIS.MarketValue CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectQuickSaleValue);
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.QuickSale.ASIS.SuggestedListPrice CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectQuickSaleValue);
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.QuickSale.Repaired.MarketValue CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectQuickSaleValue);
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.QuickSale.Repaired.SuggestedListPrice CONTENT=" + GlobalVar.theSubjectProperty.MainForm.SubjectQuickSaleValue);
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.PriceOpinion.Matrix.NormalSale.ASIS.FairMarketRent CONTENT=" + GlobalVar.theSubjectProperty.MainForm.RoundTo50(GlobalVar.theSubjectProperty.MainForm.SubjectRent)).ToString();

            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:bpomainform ATTR=NAME:signaturefile CONTENT=" + GlobalVar.theSubjectProperty.MainForm.DropBoxFolder +   "\\BPOs\\Dawn-sig.JPG");
      
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:bpomainform ATTR=NAME:signatureupload");

            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Vendor.FullName CONTENT=Dawn<SP>Zurick");
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Vendor.License.Date CONTENT=" + DateTime.Now.ToShortDateString());
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Vendor.License.Number CONTENT=471.0096163");
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Vendor.License.State CONTENT=IL");
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Vendor.Company.Name CONTENT=OKRP");
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Vendor.PhoneNo CONTENT=815-315-0203");
            //theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Vendor.Address.Address CONTENT=10325<SP>Main");

            //theMacro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:bpomainform ATTR=NAME:Order_Data.Form.Valuation.Reconciliation.FinalComments CONTENT=This<SP>evaluation<SP>was<SP>prepared<SP>by<SP>a<SP>licensed<SP>real<SP>estate<SP>broker<SP>and<SP>is<SP>not<SP>an<SP>appraisal.<SP>This<SP>evaluation<SP>cannot<SP>be<SP>used<SP>for<SP>the<SP>purpose<SP>of<SP>obtaining<SP>financing.<BR><LF>Notwithstanding<SP>any<SP>preprinted<SP>language<SP>to<SP>the<SP>contrary,<SP>this<SP>is<SP>not<SP>an<SP>appraisal<SP>of<SP>the<SP>market<SP>value<SP>of<SP>the<SP>property.<SP>If<SP>an<SP>appraisal<SP>is<SP>desired,<SP>the<SP>services<SP>of<SP>a<SP>licensed<SP>or<SP>certified<SP>appraiser<SP>must<SP>be<SP>obtained.");


       
            WriteScript(GlobalVar.theSubjectProperty.MainForm.SubjectFilePath, "prefill.iim", theMacro);

            string macroCode = theMacro.ToString();
            iim.iimPlayCode(macroCode, 60);
        }



        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {

            GenerateCompFillScript(compNum);
            WriteScript(fieldList["filepath"], compNum + ".iim", theMacro);

            string macroCode = theMacro.ToString();
            iim.iimPlayCode(macroCode, 60);
        }
    }
}
