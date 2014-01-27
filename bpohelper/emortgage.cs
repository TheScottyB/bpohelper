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

    class Emortgage
    {
       
        public void Prefill(iMacros.App iim, Form1 form)
        {
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");

            macro.AppendLine(@"SET !TIMEOUT_STEP 0");
            macro.AppendLine(@"FRAME NAME=main");

            //
            //property info section
            //
            if (form.SubjectAttached)
            {
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropType CONTENT=Condo");
            }
            else
            {
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropType CONTENT=SF<SP>Detach");
            }
           
         
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropSqFt CONTENT=" + form.SubjectAboveGLA.Replace(",", ""));

            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropBR CONTENT=" + form.SubjectBedroomCount);

            string tBath = form.SubjectBathroomCount;
            try
            {
                if (Convert.ToInt16(tBath[2].ToString()) > 0)
                {
                    tBath = tBath[0] + "." + "5"; 
                 }
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropBA CONTENT=" + tBath);
            }
            catch
            {

            }

            
          
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropTR CONTENT=" + form.SubjectRoomCount);
          
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropBsmtPerFin");
            //macro.AppendLine(@"TAG POS=7 TYPE=A FORM=NAME:BPOForm ATTR=CLASS:sbutt");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropGar");
            //macro.AppendLine(@"TAG POS=4 TYPE=A ATTR=CLASS:sitem");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropYrBlt CONTENT=" + form.SubjectYearBuilt);
          
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sIsListed");
     
     //       macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sIsListed12");


            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropSDistrict CONTENT=" + form.SubjectSchoolDistrict.Replace(" ", "<SP>"));



            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropAMC CONTENT=Good");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropLoc CONTENT=Good");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropPool CONTENT=None");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropExtAdd CONTENT=na");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropOccOwn CONTENT=Owner");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropSource CONTENT=County<SP>Tax");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropDSourceDoc CONTENT=Yes");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropIntPaint CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropIntStruct CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropIntAppl CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropIntUtils CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropIntFloor CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropIntOther CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropIntClean CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropExtPaint CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropExtStruct CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropExtLand CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropExtRoof CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropExtWindow CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropExtOther CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropRecRep CONTENT=No");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sNDataCN CONTENT=No");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropLL CONTENT=No");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sNDataEP CONTENT=No");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sNDataHS CONTENT=Stable");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sNDataCV CONTENT=5-10%");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sNDataNT CONTENT=Owner");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sNDataMktST CONTENT=Mixed");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropOcc CONTENT=AsIs");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:FinanceCost CONTENT=0");

            //
            //Brokers Market Analysis:
            //
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:ProbSaleAsIs CONTENT=" + form.SubjectMarketValue);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:ProbSaleMkt CONTENT=" + form.SubjectMarketValue);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:SuggValAsIs CONTENT=" + form.SubjectMarketValue);
            


            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 50);


        }

        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string>fieldList)
        {
            string sol;
            if (saleOrList == "sale")
            {
                sol = "Sales";
            }
            else
            {
                sol = "List";
            }
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");

            macro.AppendLine(@"SET !TIMEOUT_STEP 0");
            macro.AppendLine(@"FRAME NAME=main");


           foreach (string field in fieldList.Keys)
           {
               macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:*{0}{1}_{2} CONTENT={3}\r\n", sol, field, Regex.Match(compNum,@"\d").Value,  fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>") );
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
