using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{



    class AVM
    {
/// <summary>
/// </summary>
/// <param name="iim"></param>
/// <param name="form"></param>
/// http://avm.assetval.com/AVM/Realtor/PendingAssignments.aspx

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

      
        private string  GenerateSubjectFillScript()
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


        public void Prefill(iMacros.App iim, Form1 form)
        {
            
           // iim.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string currentUrl = iim.iimGetLastExtract();

            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 1");
            macro.Append(GenerateSubjectFillScript());


            
         


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

                macro.AppendFormat("TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rblREO{0}_2&&VALUE:{1} CONTENT=YES", compNum, fieldList["SaleType"]);
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
