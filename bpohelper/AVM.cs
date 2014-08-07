using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{



    class AVM : BPOFulfillment
    {
/// <summary>
/// </summary>
/// <param name="iim"></param>
/// <param name="form"></param>
/// http://avm.assetval.com/AVM/Realtor/PendingAssignments.aspx

      
        
          public AVM()
        {
            compFieldListTranslator["NumberOfUnits"] = "units";
            compFieldListTranslator["DataSource"] = "datasource";
            compFieldListTranslator["GLA"] = "sqft";
            compFieldListTranslator["TotalRoomCount"] = "rooms";
            compFieldListTranslator["BedroomCount"] = "bed";
            compFieldListTranslator["TotalRoomCount"] = "rooms";

            compFieldListTranslator["YearBuilt"] = "yr_built";
            compFieldListTranslator["LotSize"] = "lotsize";
            compFieldListTranslator["MainView"] = "view";

            compFieldListTranslator["ProximityToSubject"] = "prox";
     
            compFieldListTranslator["DOM"] = "dom";
            compFieldListTranslator["SalePrice"] = "sale_price";
            compFieldListTranslator["SalesDate"] = "sale_dt";





            compFieldListTranslator["NumberOfParkingSpaces"] = "garage";

            compFieldListTranslator.Add("FullBathCount", "bath");
            compFieldListTranslator.Add("HalfBathCount", "partial_bath");
            compFieldListTranslator.Add("OverallPropertyCondition", "condition");
            compFieldListTranslator.Add("OverallLocationDensity", "location");
            compFieldListTranslator.Add("BuildingSkeleton", "construction");
            compFieldListTranslator.Add("WaterFront", "waterfront");
            compFieldListTranslator.Add("OriginalListPrice", "original_lp");
            compFieldListTranslator.Add("NumberOfPriceReductions", "pricereductions");
          

        }


        public AVM(MLSListing m) 
            : this()
        {
            targetComp = m;
        }
        
        
        private Dictionary<string, string> propTypeTranslator = new Dictionary<string, string>()
         {
            {"Detached", "Single Family"},
            {"Attached", "Condo"}
           

         };

        private Dictionary<string, string> subjectFieldListTranslator = new Dictionary<string, string>()
        {
            {"ParcelID", "sub_apn"}, 
            {"County", "County"},
            {"PropertyType", "*property_type"},
            {"Rent", "fair_mkt_rent"}
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


        private void helper_SetAddressFields()
         {
             int positionNumber = 0;
            
            int.TryParse(targetCompNumber,  out positionNumber);

            if (saleOrListFlag == "list")
            {
                positionNumber = positionNumber + 3;
            }

             theMacro.AppendLine(@"TAG POS=" + positionNumber.ToString() + " TYPE=INPUT:BUTTON FORM=NAME:aspnetForm ATTR=CLASS:btnBlue");
             theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$txtCompAddress CONTENT=" + targetComp.StreetAddress.Replace(" ", "<SP>"));
             theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$txtCompCity CONTENT=" + targetComp.City.Replace(" ", "<SP>"));
             theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$txtCompZip CONTENT=" + targetComp.Zipcode.Replace(" ", "<SP>"));
             theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$btnGetProximity");
             theMacro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
             theMacro.AppendLine(@"WAIT SECONDS=1");
             theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$btnAcceptProximity");
      
         }

        private void helper_SetCommonFields()
        {


            helper_SetAddressFields();
            
           
            compSelectionBoxList.Add(compFieldListTranslator["TransactionType"], "%Arm");

            //
            //Selection Boxes
            //
            compSelectionBoxList.Add(compFieldListTranslator["DataSource"], "%M");
            compSelectionBoxList.Add(compFieldListTranslator["NumberOfUnits"], "%1");
            compSelectionBoxList.Add(compFieldListTranslator["OverallPropertyCondition"], "%A");
            compSelectionBoxList.Add(compFieldListTranslator["OverallLocationDensity"], "%S");
            compSelectionBoxList.Add(compFieldListTranslator["BuildingSkeleton"], "%F");
            compSelectionBoxList.Add(compFieldListTranslator["MainView"], "%R");
            compSelectionBoxList.Add(compFieldListTranslator["WaterFront"], "%" + targetComp.WaterFrontYesNo()[0].ToString());
            if (targetComp.NumberOfPriceChanges >= 5)
            {
                compSelectionBoxList.Add(compFieldListTranslator["NumberOfPriceReductions"], "%5");
            }
            else
            {
                compSelectionBoxList.Add(compFieldListTranslator["NumberOfPriceReductions"], "%" + targetComp.NumberOfPriceChanges.ToString());
            }
        

            compSelectionBoxList.Add("lotsize_unit", "%A");
            compSelectionBoxList.Add("gated", "%N");
            compSelectionBoxList.Add("pool", "%N");
            compSelectionBoxList.Add("parking", "%D");
            compSelectionBoxList.Add("designAppeal", "%E");


            //
            //Text Fields
            //
            //main characteristics
            compTextFieldList.Add(compFieldListTranslator["GLA"], targetComp.GLA.ToString());
            compTextFieldList.Add(compFieldListTranslator["NumberOfUnits"], "1");
            compTextFieldList.Add(compFieldListTranslator["YearBuilt"], targetComp.YearBuilt.ToString());
            compTextFieldList.Add(compFieldListTranslator["LotSize"], targetComp.Lotsize.ToString());
            //rooms
            compTextFieldList.Add(compFieldListTranslator["TotalRoomCount"], targetComp.TotalRoomCount.ToString());
            compTextFieldList.Add(compFieldListTranslator["BedroomCount"], targetComp.BedroomCount);
            compTextFieldList.Add(compFieldListTranslator["FullBathCount"], targetComp.FullBathCount);
            compTextFieldList.Add(compFieldListTranslator["HalfBathCount"], targetComp.HalfBathCount);
            //garage
            compTextFieldList.Add(compFieldListTranslator["NumberOfParkingSpaces"], targetComp.NumberGarageStalls());
            //pricing and listing history
            compTextFieldList.Add(compFieldListTranslator["OriginalListPrice"], targetComp.OriginalListPrice.ToString());
            
            compTextFieldList.Add(compFieldListTranslator["DOM"], targetComp.DOM);
            //Misc
            compTextFieldList.Add("energyeff", "NA");
            compTextFieldList.Add("other", "NA");
            compTextFieldList.Add("heat_cool", @"FA/CA");
            compTextFieldList.Add("concessions", "NA");
            compTextFieldList.Add(compFieldListTranslator["ProximityToSubject"], targetComp.ProximityToSubject.ToString());
            //basement
            if (targetComp.BasementType.Contains("Full"))
                compSelectionBoxList.Add("bsmt_fin_yn", "%F");
            if (targetComp.BasementType.Contains("Partial"))
                compSelectionBoxList.Add("bsmt_fin_yn", "%PF");

            if (targetComp.BasementType.Contains("None"))
                compSelectionBoxList.Add("bsmt_fin_yn", "%N");
            compTextFieldList.Add("bsmt_fin_per", targetComp.BasementFinishedPercentage());

            compTextFieldList.Add("sqft_bg", targetComp.BasementGLA());
            compTextFieldList.Add("rooms_bg", targetComp.NumberOfBasementRooms());
        }


      
        protected override string  GenerateSubjectFillScript()
        {
            StringBuilder macro = new StringBuilder();
            //borrowers name (already filled)
            //#APN
            subjectFieldList.Add(subjectFieldListTranslator["ParcelID"], GlobalVar.theSubjectProperty.ParcelID);

            //County
            subjectFieldList.Add(subjectFieldListTranslator["County"], GlobalVar.theSubjectProperty.County);

            //PropertyType (selection box)
                //Single Family
                //Condo
                //*bunch of others* but we don't care

            //FairMArketRent
            //subjectFieldList.Add(subjectFieldListTranslator["Rent"], GlobalVar.theSubjectProperty);
            //Secure (Yes/No)
            //Occupancy (Yes/No)
            //ZoningCode 
            //ZoningDescription
            //ZoningCompliance
            //IllegalUnits
                //if yes, decription
            //currentUse
            //bestUse
                //if no, describe
            //currentlyListed (Yes/No)
            //listedLast36Months
                //listing status, if either of the above
                //
                //(current listing section)
                //
            //set of red flag checkboxes
                //damaged
                //contruction
                //environmental
                //zoning
                //market activity
                //boarded
                //stigma
                //other
                //none
                    //if anything checked, comments

          //  TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$sub_apn CONTENT=ttt
            foreach (string field in subjectFieldList.Keys)
            {
                if (field.Contains("*"))
                {

                }
                else
                {
                    //orignal way using * instead of C
                    //macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:*{0}{1}_{2} CONTENT={3}\r\n", sol, field, Regex.Match(compNum, @"\d").Value, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                    macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21${0} CONTENT={1}\r\n", field, subjectFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                }



             
            }


            return macro.ToString();


        }


        protected override void GenerateSaleCompFillScript()
        {
            helper_CompFillHeaderScript();
            helper_SetCommonFields();

            //fields and/or changes specific to SALE comps
            compFieldListTranslator["CurrentListPrice"] = "sale_lp";
            compTextFieldList.Add(compFieldListTranslator["CurrentListPrice"], targetComp.CurrentListPrice.ToString());
            compTextFieldList.Add(compFieldListTranslator["SalePrice"], targetComp.SalePrice.ToString());
            compTextFieldList.Add(compFieldListTranslator["SalesDate"], targetComp.SalesDate.ToShortDateString());

            foreach (string field in compTextFieldList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:*s{0}_{1} CONTENT={2}\r\n", targetCompNumber, field, compTextFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compSelectionBoxList.Keys)
            {                      //"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control* ATTR=NAME:*{0}{1} CONTENT={2}");
                theMacro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:*s{0}_{1} CONTENT={2}\r\n", targetCompNumber, field, compSelectionBoxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }


            foreach (string field in compRadioButtonList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ACTION:/Order/OrderEdit* ATTR=NAME:SalesComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compRadioButtonList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compCheckboxList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEdit* ATTR=NAME:SalesComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compCheckboxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            theMacro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:*S"+ targetCompNumber + "OrigListDt CONTENT=" + targetComp.ListDateString);


            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEditWizardFNMA/step3/11488927 ATTR=NAME:SalesComp1.HasForcedWarmAirForHeat CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEditWizardFNMA/step3/11488927 ATTR=NAME:SalesComp1.HasCentralAir CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=ACTION:/Order/OrderEditWizardFNMA/step3/11488927 ATTR=CLASS:mceIframeContainer<SP>mceFirst<SP>mceLast");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:/Order/OrderEditWizardFNMA/step3/11488927 ATTR=NAME:SaveButton");  
        }

        protected override void GenerateListCompFillScript()
        {
            helper_CompFillHeaderScript();
            helper_SetCommonFields();

            //fields and/or changes specific to LIST comps
            compFieldListTranslator["CurrentListPrice"] = "current_lp";
            compTextFieldList.Add(compFieldListTranslator["CurrentListPrice"], targetComp.CurrentListPrice.ToString());
            compTextFieldList.Add("list_dt", targetComp.ListDate.ToShortDateString());


            foreach (string field in compTextFieldList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXTFORM=NAME:aspnetForm ATTR=NAME:*l{0}_{1} CONTENT={2}\r\n", targetCompNumber, field, compTextFieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compSelectionBoxList.Keys)
            {                      //"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control* ATTR=NAME:*{0}{1} CONTENT={2}");
                theMacro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:*l{0}_{1} CONTENT={2}\r\n", targetCompNumber, field, compSelectionBoxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }


            foreach (string field in compRadioButtonList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ACTION:/Order/OrderEdit* ATTR=NAME:SalesComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compRadioButtonList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }

            foreach (string field in compCheckboxList.Keys)
            {
                theMacro.AppendFormat("TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ACTION:/Order/OrderEdit* ATTR=NAME:SalesComp{0}.{1} CONTENT={2}\r\n", targetCompNumber, field, compCheckboxList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
            }


        }



        public void Prefill(iMacros.App iim, Form1 form)
        {
            
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$owner_name CONTENT=" + form.SubjectOOR.Replace(" ", "<SP>"));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$sub_apn CONTENT=" + form.SubjectOOR.Replace(" ", "<SP>"));

            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$fair_mkt_rent CONTENT=" + form.SubjectRent.Replace(" ", "<SP>"));

            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$sub_zone_class CONTENT=SFR");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$sub_zone_desc CONTENT=Residential");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$sub_currentuse CONTENT=SFR");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$sub_zone_high CONTENT=%Y");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$sale_dt");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$sale_price");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$currently_listed_yn CONTENT=%N");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$prev_listed_yn CONTENT=%N");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:aspnetForm ATTR=NAME:ctl00$Body$BPOV21$chkRFDamaged CONTENT=NO");
            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 60);

        }

        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {
            iim.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string currentUrl = iim.iimGetLastExtract();

            targetCompNumber = Regex.Match(compNum, @"\d").Value;
            saleOrListFlag = saleOrList;

            if (saleOrList == "sale")
            {

                GenerateSaleCompFillScript();

            }
            else
            {

                GenerateListCompFillScript();

            }

            WriteScript(fieldList["filepath"], compNum + ".iim", theMacro);

            string macroCode = theMacro.ToString();
            iim.iimPlayCode(macroCode, 60);
        }
    }
}
