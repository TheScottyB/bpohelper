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
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=VALUE:SFR CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=VALUE:OWNER CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=ID:rbtnSalesResponseNO&&VALUE:NO CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=ID:rbtnSoldListedNo&&VALUE:NO CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=ID:rbtnEvidRepNo&&VALUE:NO CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=VALUE:AVERAGE CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=VALUE:SUBURBAN CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=VALUE:SOMEWHATSIMILAR CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=VALUE:OVER<SP>75% CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:IEFORM ATTR=ID:rbtnStable&&VALUE:STABLE CONTENT=YES");

            //change page
            macro.AppendLine(@"FRAME NAME=pageSelect");
            macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:myForm ATTR=ID:saveWIP&&VALUE:Save<SP>WIP");

          

            macro.AppendLine(@"WAIT SECONDS=3");
            macro.AppendLine(@"FRAME NAME=pageSelect");
         
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:myForm ATTR=ID:pageOptions CONTENT=%*_page_2.htm");

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
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTLEASEFEE CONTENT=Fee<SP>Simple");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTLEASEFEE CONTENT=Fee<SP>Simple");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP1LEASEFEE CONTENT=Fee<SP>Simple");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP2LEASEFEE CONTENT=Fee<SP>Simple");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:COMP3LEASEFEE CONTENT=Fee<SP>Simple");
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
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP1LEASEFEE CONTENT=Fee<SP>Simple");
           // macro.AppendLine(@"TAG POS=1 TYPE=DIV FORM=NAME:IEFORM ATTR=ID:Lyr158");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP2LEASEFEE CONTENT=Fee<SP>Simple");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP3LEASEFEE CONTENT=Fee<SP>Simple");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP1CONDITION CONTENT=Average");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP2CONDITION CONTENT=Average");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTCOMP3CONDITION CONTENT=Average");

            //subject info, top
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTSALESPRICE");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTSALETYPE");
            
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTSALEDT");
            //macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTDAYSONMARKET");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTPROPTYPE CONTENT=" + form.SubjectMlsType.Replace(" ","<SP>"));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTPROPTYPE CONTENT=" + form.SubjectMlsType.Replace(" ", "<SP>"));
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTLOTSIZE CONTENT=" + form.SubjectLotSize);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTLOTSIZE CONTENT=" + form.SubjectLotSize);
            //Exterior Construction / Finish field
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTDESIGNSTYLE");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTYRBUILT CONTENT=" + form.SubjectYearBuilt);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTYRBUILT CONTENT=" + form.SubjectYearBuilt);

            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTGBASQFT CONTENT=" + form.SubjectAboveGLA);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTLIVINGSQFT CONTENT=" + form.SubjectAboveGLA);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTBEDROOMS CONTENT=" + form.SubjectBedroomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTBEDROOMS CONTENT=" + form.SubjectBedroomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTBATH CONTENT=" + form.SubjectBathroomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTBATH CONTENT=" + form.SubjectBathroomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTBASEMENT CONTENT=" + form.SubjectBasementType);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTBASEMENT CONTENT=" + form.SubjectBasementType);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTGARAGE CONTENT=" + form.SubjectParkingType);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTGARAGE CONTENT=" + form.SubjectParkingType);
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTAMENITIES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SALESCOMPUNLABELED");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTUNLABELEDDESC");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SUBJECTMARKETCHANGE");
            //macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
            //
            //bottom
            //
          //  macro.AppendLine(@"FRAME NAME=pageView");
         
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTSALESPRICE");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTSALETYPE");

            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTSALEDT");
            //macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTDAYSONMARKET");
            

            
           
            //Exterior Construction / Finish field
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTDESIGNSTYLE");
           

            
   
          
          
       
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTAMENITIES");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:SALESCOMPUNLABELED");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTUNLABELEDDESC");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTMARKETCHANGE");
            //macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTDATA CONTENT=Tax");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:LISTSUBJECTPROJSIZE CONTENT=na");
            
           
        
          


           


            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 120);


        }

        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
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
                    macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:*{0}{1}_{2}\r\n", sol, field.Replace("*", ""), Regex.Match(compNum, @"\d").Value);
                    macro.AppendLine(@"DS CMD=KEY X={{!TAGX}} Y={{!TAGY}} CONTENT=" + fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                }
                else
                {
                    macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:IEFORM ATTR=NAME:*{0}{1}{2} CONTENT={3}\r\n", sol, Regex.Match(compNum, @"\d").Value, field, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                }

                //  macro.AppendLine(@"DS CMD=CLICK X={{!TAGX}} Y={{!TAGY}}");
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
