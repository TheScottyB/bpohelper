using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{

    class Lres
    {
        private readonly IYourForm form;
        public Lres(IYourForm form)
        {
            this.form = form;
        }

        public void Prefill(iMacros.App iim)
        {

            //
            //Page 01
            //
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbApnNumber CONTENT=" + form.SubjectPin);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbTotalRoomCount CONTENT=" + form.SubjectRoomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbTotalBedroom CONTENT=" + form.SubjectBedroomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbTotalBathroom CONTENT=" + form.SubjectBathroomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbLivingArea CONTENT=" + form.SubjectAboveGLA.Replace(",",""));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbLotSize CONTENT=" + form.SubjectLotSize);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbYearBuilt CONTENT=" + form.SubjectYearBuilt);
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlCondition CONTENT=%Average");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlStyle CONTENT=%Cntmp");
            if (form.SubjectBasementType.ToLower().Contains("full"))
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlBasementFeature CONTENT=%Full");
            }
            else if (form.SubjectBasementType.ToLower().Contains("partial"))
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlBasementFeature CONTENT=%Partial");
            }
            else
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlBasementFeature CONTENT=%Crawl");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlBasement CONTENT=%None");
            }

            if (form.SubjectBasementDetails.ToLower().Contains("finished") && !form.SubjectBasementDetails.ToLower().Contains("unfinished"))
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlBasement CONTENT=%Fully<SP>Finished");
            }
            else
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlBasement CONTENT=%Unfinished");
            }
            
           

            //[OPTION]1 Car Attached
            //[OPTION]1 Car Detached
            //[OPTION]1 Carport
            //[OPTION]1 on Street
            //[OPTION]1 Uncovered
            //[OPTION]2 Car Attached
            //[OPTION]2 Car Detached
            //[OPTION]2 Carport
            //[OPTION]2 on Street
            //[OPTION]2 Uncovered
            //[OPTION]2+ Car Attached
            //[OPTION]2+ Car Detached
            //[OPTION]2+ Carport
            //[OPTION]2+ on Street
            //[OPTION]2+ Uncovered
            //[OPTION]None
            string lresGarageStr = "None";
            string numSpaces = Regex.Match(form.SubjectParkingType, @"\d").Value;
            string att_det = "";

            if (!string.IsNullOrEmpty(numSpaces))
            {
                if (form.SubjectParkingType.ToLower().Contains("att"))
                {
                    att_det = "Attached";
                }
                else if (form.SubjectParkingType.ToLower().Contains("det"))
                {
                    att_det = "Detached";
                }

                switch (numSpaces)
                {
                    case "1" :
                        lresGarageStr = "1<SP>Car<SP>" + att_det;
                        break;
                    case "2" :
                        lresGarageStr = "2<SP>Car<SP>" + att_det;
                        break;
                    default :
                        lresGarageStr = "2+<SP>Car<SP>" + att_det;
                        break;
                }

            }

            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlGarage CONTENT=%" + lresGarageStr );
            //macro.AppendLine(@"WAIT SECONDS = 5");

            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlInfoSource CONTENT=%Tax<SP>Record");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlPropertyType CONTENT=%SFR");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlOccupancy CONTENT=%Occupied");
            //[OPTION]Aluminum[OPTION]Brick[OPTION]Log[OPTION]Mix[OPTION]Stone[OPTION]Vinyl[OPTION]Wood[OPTION]Other
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlConstructionType CONTENT=%Other");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbPropertyTaxes CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbInspectionDate CONTENT=" + DateTime.Now.ToShortDateString());
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbRangeHigh CONTENT=400000");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbRangeLow CONTENT=50000");
            //           macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button3&&VALUE:Return<SP>to<SP>Main<SP>Menu");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbRangeAverage CONTENT=150000");

          //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button4&&VALUE:Save");


            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbSubjectConditionComment CONTENT=No<SP>adverse<SP>conditions<SP>were<SP>noted<SP>at<SP>the<SP>time<SP>of<SP>inspection<SP>based<SP>on<SP>exterior<SP>observations.");
          
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button2&&VALUE:Save<SP>and<SP>Continue");

            
         
            string macroCode = macro.ToString();
             iim.iimPlayCode(macroCode, 30);

            //
            //Page 2
            //
             macro.Clear();
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rbRecommendSale_0&&VALUE:As<SP>Is CONTENT=YES");
             macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbRepairConditionComment CONTENT=No<SP>repairs<SP>noted<SP>from<SP>drive-by.");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button2&&VALUE:Save<SP>and<SP>Continue");
             iim.iimPlayCode(macro.ToString(), 30);
            //
            //Page 3
            //
             macro.Clear();
             macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlLocation CONTENT=%Average");
             macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlMarketConditions CONTENT=%Stable");
             macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlLocalEconomy CONTENT=%Stable");
             macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlAreaType CONTENT=%Surburban");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbAverageMarketTime CONTENT=120");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbAppDep CONTENT=0");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbLocationFactor CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbWires CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbBoardUps CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbCommercialUses CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbFreeway CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbRailroad CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbAirport CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbWaste CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbFloodPlain CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbOtherComments CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button2&&VALUE:Save<SP>and<SP>Continue");
             iim.iimPlayCode(macro.ToString(), 30);

            //
            //Page 4
            //
             macro.Clear();
             macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlLocation CONTENT=%Average");
             macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlMarketConditions CONTENT=%Stable");
             macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlLocalEconomy CONTENT=%Stable");
             macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlAreaType CONTENT=%Surburban");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbAverageMarketTime CONTENT=120");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbAppDep CONTENT=0");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbLocationFactor CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbWires CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbBoardUps CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbCommercialUses CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbFreeway CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbRailroad CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbAirport CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbWaste CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbFloodPlain CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbOtherComments CONTENT=NA");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button2&&VALUE:Save<SP>and<SP>Continue");
             iim.iimPlayCode(macro.ToString(), 30);

            //
            //Page 5
            //
             macro.Clear();
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbListAsIsDaysThirty CONTENT=30900");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbSaleAsIsDaysThirty CONTENT=30000");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbListAsIsDaysNinety CONTENT=40900");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbSaleAsIsDaysNinety CONTENT=40000");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbListAsRepairedDaysNinety CONTENT=40900");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbSaleAsRepairedDaysNinety CONTENT=40000");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbListAsRepairedMktPrice CONTENT=50900");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbSaleAsRepairedMktPrice CONTENT=50000");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button2&&VALUE:Save<SP>and<SP>Continue");
             iim.iimPlayCode(macro.ToString(), 30);

            //
            //Page 6
            //
             macro.Clear();

             //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbHoaFees CONTENT=0");
             //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbHoaCompany CONTENT=na");
             //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbHoaPhone CONTENT=na");
             //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button2&&VALUE:Save<SP>and<SP>Continue");
             //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbHoaPhone CONTENT=na");
             //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button2&&VALUE:Save<SP>and<SP>Continue");
             //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbHoaFees");
             //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbHoaCompany");
             //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbHoaPhone");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button2&&VALUE:Save<SP>and<SP>Continue");


             iim.iimPlayCode(macro.ToString(), 30);

             //
             //Page 7
             //
             macro.Clear();

             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_rbCurrentlyListed_1&&VALUE:0 CONTENT=YES");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbListingAgent");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbCurrentListPrice");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbDom");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbCurrentListDate");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbOrigListPrice");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbDomAtOrigLP");
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button2&&VALUE:Save<SP>and<SP>Continue");


             iim.iimPlayCode(macro.ToString(), 30);

             //
             //Page 8
             //
             macro.Clear();
             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button2&&VALUE:Save<SP>and<SP>Continue");

             iim.iimPlayCode(macro.ToString(), 30);

             //
             //Page 9
             //
             macro.Clear();



             macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button2&&VALUE:Save<SP>and<SP>Continue");

             iim.iimPlayCode(macro.ToString(), 30);



        }

        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {
           
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");

            macro.AppendLine(@"SET !TIMEOUT_STEP 0");



            foreach (string field in fieldList.Keys)
            {
                if (fieldList[field].Contains("%"))
                {
                    macro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:*{0}{1} CONTENT={2}\r\n", field, Regex.Match(compNum, @"\d").Value, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));

                }
                else
                {
                    macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:*{0}{1} CONTENT={2}\r\n", field, Regex.Match(compNum, @"\d").Value, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                }
            }
           

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 60);

          
           
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbProximity1");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbProximity2");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_tbProximity3");
          
          
         
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlStyle1 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlStyle2 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlStyle3 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlConstructionType1 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlConstructionType2 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlConstructionType3 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlBasementFeature1 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlBasementFeature2 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlBasementFeature3 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlBasementFinish1 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlBasementFinish2 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlBasementFinish3 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlLocation1 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlLocation2 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlLocation3 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlCondition1 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlCondition2 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlCondition3 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlGarage1 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlGarage2 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlGarage3 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlCompared1 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlCompared2 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlCompared3 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlInfoSource1 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlInfoSource2 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlInfoSource3 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlTransType1 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlTransType2 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlTransType3 CONTENT=#1");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button1&&VALUE:Previous<SP>Page");






        }
    }
}
