using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{

    //
    //based on form "i"
    //

    class LandSafe
    {

        protected void WriteScript(string path, string filename, StringBuilder script)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(path + "\\" + filename))
            {
                file.Write(script);
            }


        }

        public void Prefill(iMacros.App iim, Form1 form)
        {
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 1");
            macro.AppendLine(@"FRAME NAME=pageView");

            Dictionary<string, string> fieldList = new Dictionary<string, string>();

            //page 1
            macro.AppendLine(@"FRAME NAME=pageView");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=VALUE:SFR CONTENT=NO");
            macro.AppendLine(@"FRAME NAME=mainFrame");
            macro.AppendLine(@"TAG POS=1 TYPE=IFRAME FORM=ID:formviewform ATTR=ID:form_view");
            macro.AppendLine(@"FRAME NAME=pageView");
            if (form.SubjectAttached)
            {
                macro.AppendLine(@"TAG POS=2 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=NAME:PROPTYPE CONTENT=YES");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=NAME:PROJTYPE CONTENT=YES");

            }else
            {
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=VALUE:SFR CONTENT=YES");
            }
           // macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:IEFORM ATTR=NAME:COMMENTCONDITIONIMPROVE CONTENT=No<SP>adverse<SP>conditions<SP>noted.");
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:IEFORM ATTR=NAME:REVIEWINTERIOR CONTENT=No<SP>adverse<SP>conditions<SP>noted.");
            
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=VALUE:OWNER CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=ID:rbtnSalesResponseNO&&VALUE:NO CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=ID:rbtnSoldListedNo&&VALUE:NO CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=ID:rbtnEvidRepNo&&VALUE:NO CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=VALUE:AVERAGE CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=VALUE:SUBURBAN CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=VALUE:SOMEWHATSIMILAR CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=VALUE:OVER<SP>75% CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=ID:rbtnStable&&VALUE:STABLE CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=NAME:BUILTUP CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SINGLEFAMPRICELOW CONTENT=" + form.SubjectNeighborhood.minSalePrice);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SINGLEFAMPRICEHIGH CONTENT=" + form.SubjectNeighborhood.maxSalePrice);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:VALUETRENDMKTAREASOURCE CONTENT=MLS");
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:IEFORM ATTR=NAME:REVIEWNBRHOODCOMMENTS CONTENT=Approx.<SP>20%<SP>active<SP>REO<SP>mixed<SP>with<SP>short<SP>sales<SP>and<SP>traditional<SP>sales.<SP><SP>High<SP>demand<SP>at<SP>and<SP>below<SP>150k.");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO ATTR=NAME:SALESHISTORYRESPONSE CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO ATTR=NAME:SUBJECTSOLDLISTED CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO ATTR=NAME:PROPVALUES CONTENT=YES");
            macro.AppendLine(@"TAG POS=3 TYPE=INPUT:RADIO ATTR=NAME:BUILTUP CONTENT=YES");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO ATTR=NAME:EVIDENCEREPAIRS CONTENT=YES");

            ////change page
            //macro.AppendLine(@"FRAME NAME=pageSelect");
            //macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:myForm ATTR=ID:saveWIP&&VALUE:Save<SP>WIP");

          

            //macro.AppendLine(@"WAIT SECONDS=3");
            //macro.AppendLine(@"FRAME NAME=pageSelect");
         
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:myForm ATTR=ID:pageOptions CONTENT=%*_page_2.htm");

            //page 2, default 
            macro.AppendLine(@"FRAME NAME=pageView");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTDATASRC CONTENT=Tax");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTDATA CONTENT=Tax");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP1DATA CONTENT=MLS");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP2DATA CONTENT=MLS");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP3DATA CONTENT=MLS");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTPROJSIZETYPE CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTPROJSIZE CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP1PROJSIZETYPE CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP2PROJSIZETYPE CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP3PROJSIZETYPE CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTCONDITION CONTENT=Average");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTCONDITION CONTENT=Average");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP1CONDITION CONTENT=Average");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP2CONDITION CONTENT=Average");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP3CONDITION CONTENT=Average");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP1PROJSIZE CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP2PROJSIZE CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP3PROJSIZE CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP1DATA CONTENT=MLS");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP2DATA CONTENT=MLS");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP3DATA CONTENT=MLS");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP1CONDITION CONTENT=Average");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP2CONDITION CONTENT=Average");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP3CONDITION CONTENT=Average");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTPROPTYPE CONTENT=" + form.SubjectMlsType.Replace(" ","<SP>"));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTPROPTYPE CONTENT=" + form.SubjectMlsType.Replace(" ", "<SP>"));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTLOTSIZE CONTENT=" + form.SubjectLotSize);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTLOTSIZE CONTENT=" + form.SubjectLotSize);
            //Exterior Construction / Finish field
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTDESIGNSTYLE CONTENT=" + form.SubjectExteriorFinish);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT ATTR=NAME:LISTSUBJECTDESIGNSTYLE CONTENT=" + form.SubjectExteriorFinish);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTYRBUILT CONTENT=" + form.SubjectYearBuilt);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTYRBUILT CONTENT=" + form.SubjectYearBuilt);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTGBASQFT CONTENT=" + form.SubjectAboveGLA);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTLIVINGSQFT CONTENT=" + form.SubjectAboveGLA);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTBEDROOMS CONTENT=" + form.SubjectBedroomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTBEDROOMS CONTENT=" + form.SubjectBedroomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTBATH CONTENT=" + form.SubjectBathroomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTBATH CONTENT=" + form.SubjectBathroomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTBASEMENT CONTENT=" + form.SubjectBasementType.Replace(" ","<SP>"));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTBASEMENT CONTENT=" + form.SubjectBasementType.Replace(" ","<SP>"));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTGARAGE CONTENT=" + form.SubjectParkingType.Replace(" ","<SP>"));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTGARAGE CONTENT=" + form.SubjectParkingType.Replace(" ", "<SP>"));    
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTAMENITIES CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SALESCOMPUNLABELED CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTUNLABELEDDESC CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTMARKETCHANGE CONTENT=no");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP1AMENITIES CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP2AMENITIES CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP3AMENITIES CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP1UNLABELED CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP2UNLABELED CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP3UNLABELED CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTOTHERAMENITIES CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP1OTHERAMENITIES CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP2OTHERAMENITIES CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP3OTHERAMENITIES CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMPUNLABELEDNAME CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMPUNLABELEDDESC CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP1UNLABELED CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP2UNLABELED CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP3UNLABELED CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:SUBJECTLEASEFEE CONTENT=%FEE<SP>SIMPLE");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:SUBJECTMARKETCHANGE CONTENT=%NO");

            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:COMP1LEASEFEE CONTENT=%FEE<SP>SIMPLE");

            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:COMP2LEASEFEE CONTENT=%FEE<SP>SIMPLE");
            macro.AppendLine(@"WAIT SECONDS=1");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:COMP3LEASEFEE CONTENT=%FEE<SP>SIMPLE");


            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:LISTCOMP1LEASEFEE CONTENT=%FEE<SP>SIMPLE");
     
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:LISTCOMP2LEASEFEE CONTENT=%FEE<SP>SIMPLE");

            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:LISTCOMP3LEASEFEE CONTENT=%FEE<SP>SIMPLE");

            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:COMP1MARKETCHANGE CONTENT=%NO");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:COMP2MARKETCHANGE CONTENT=%NO");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:COMP3MARKETCHANGE CONTENT=%NO");

            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:LISTCOMP1MARKETCHANGE CONTENT=%NO");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:LISTCOMP2MARKETCHANGE CONTENT=%NO");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:LISTCOMP3MARKETCHANGE CONTENT=%NO");

            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT ATTR=NAME:SUBJECTSALESPRICE CONTENT=N/A");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:SUBJECTSALETYPE CONTENT=%NO");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT ATTR=NAME:SUBJECTSALEDT CONTENT=N/A");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT ATTR=NAME:SUBJECTDAYSONMARKET CONTENT=N/A");

            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:LISTSUBJECTSALETYPE CONTENT=%NO");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT ATTR=NAME:LISTSUBJECTPRICEREVDT CONTENT=n/a");
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=NAME:LISTSUBJECTLEASEFEE CONTENT=%FEE<SP>SIMPLE");
         



            ////
            ////change to page 3
            ////

            //macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
            //macro.AppendLine(@"FRAME NAME=pageSelect");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:myForm ATTR=NAME:SaveWIP");
            //macro.AppendLine(@"WAIT SECONDS=3");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:myForm ATTR=ID:pageOptions CONTENT=%7001BOADB/7001BOADB_page_3.htm");

     

            //
            //Page 3
            //
            macro.AppendLine(@"FRAME NAME=pageView");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:ASISVALUE CONTENT=" + form.SubjectMarketValue);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:ESTMARKETVALUENORMALREPAIRED CONTENT=" + form.SubjectMarketValue);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=NAME:ASISDATERESPONSE CONTENT=NO");
            macro.AppendLine(@"TAG POS=3 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=NAME:ASISDATERESPONSE CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:ASISDATEAMT3 CONTENT=" + form.SubjectQuickSaleValue);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:APPRAISERDTSIGNED CONTENT=" + DateTime.Now.ToShortDateString());
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:APPRAISER CONTENT=Scott<SP>Beilfuss");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:APPRAISERCONAME CONTENT=OK<SP>Assoc,RealtyPlus");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:APPRAISERSTREET CONTENT=202<SP>Highbridge");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:APPRAISERCITY CONTENT=McHenry");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:APPRAISERSTATEPROV CONTENT=IL");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:APPRAISERZIP CONTENT=60050");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:APPRAISERPHONE CONTENT=8153150203");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:APPRAISEREMAIL CONTENT=beilsco@gmail.com");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:APPRAISERSTATELICNUM CONTENT=471.009162");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:APPRAISERLICSTATEPROV CONTENT=IL");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=NAME:REVIEWSUBJECTPHOTORESPONSE CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=NAME:REVIEWSALESCOMPPHOTORESPONSE CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=NAME:REVIEWLISTCOMPPHOTORESPONSE CONTENT=YES");

           


            string macroCode = macro.ToString();
            WriteScript(form.SubjectFilePath, "prefill.iim", macro);
            iim.iimPlayCode(macroCode, 120);


        }

        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList, Form1 form)
        {
            string sol;
            if (saleOrList == "sale")
            {
                sol = "COMP";
            }
            else
            {
                sol = "LISTCOMP";
            }
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");

            macro.AppendLine(@"SET !TIMEOUT_STEP 0");
            macro.AppendLine(@"FRAME NAME=pageView");
            macro.AppendLine(@"SET !REPLAYSPEED FAST");


            foreach (string field in fieldList.Keys)
            {
                if (field.Contains("*"))
                {
                    //drop down box
                    macro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=NAME:IEFORM ATTR=NAME:{0}{2}{1} CONTENT=%{3}\r\n", sol, field.Replace("*", ""), Regex.Match(compNum, @"\d").Value, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                    //macro.AppendLine(@"DS CMD=KEY X={{!TAGX}} Y={{!TAGY}} CONTENT=" + fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                }
                else
                {
                    macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:*{0}{1}{2} CONTENT={3}\r\n", sol, Regex.Match(compNum, @"\d").Value, field, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                    if (field.Contains("STREET"))
                    {
                        macro.AppendFormat("TAG POS=1 TYPE=TEXTAREA FORM=NAME:IEFORM ATTR=NAME:{0}{1}*STREET CONTENT={3}\r\n", sol, Regex.Match(compNum, @"\d").Value, field, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                    }
                }

                //  macro.AppendLine(@"DS CMD=CLICK X={{!TAGX}} Y={{!TAGY}}");
            }
            //
            //TBD
            //
            //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListOLDate_1");
            //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListCLPrice_1");
            //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListCLDate_1");

            WriteScript(form.SubjectFilePath, compNum, macro);
            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 60);
        }
    }
}
