using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;



namespace bpohelper
{



    class FirstAmerican : BPOFulfillment 
    {
        public FirstAmerican()
        { }


        public FirstAmerican(MLSListing m) 
            : this()
        {
            targetComp = m;
        }

        private Form1 callingForm;
       
        public void Prefill(iMacros.App iim, Form1 form)
        {
           


        }

        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string> fieldList)
        {

            targetCompNumber = Regex.Match(compNum, @"\d").Value;

            StringBuilder macro = new StringBuilder();

          
            


            string inputType = "INPUT:TEXT";

             //Dictionary<string, string> inputType = new Dictionary<string, string>();
             //inputType.Add("text", "INPUT:TEXT");
             //inputType.Add("selection", "SELECT");


            int compNumber = 0;

            if (saleOrList == "list")
            {
               
                compNumber = Convert.ToInt32(targetCompNumber) - 1;
            }
            else
            {

                compNumber = Convert.ToInt32(targetCompNumber) + 2;
            }

            Dictionary<int, string[]> commands =  new Dictionary<int, string[]>();


            commands.Add(0, new string[] { inputType, "Address_Line1", targetComp.StreetAddress.Replace(" ", "<SP>") });
            commands.Add(1, new string[] { inputType, "Address_City", targetComp.City.Replace(" ", "<SP>") });
            commands.Add(2, new string[] { "SELECT", "Address_State", "%IL"});
            commands.Add(3, new string[] { inputType, "Address_ZipCode", targetComp.Zipcode });
            commands.Add(4, new string[] { inputType, "OrigListPrice", targetComp.OriginalListPrice.ToString() });
            commands.Add(5, new string[] { inputType, "CurrListPrice", targetComp.CurrentListPrice.ToString() });

            if (saleOrList == "sale")
            {
                commands.Add(6, new string[] { inputType, "SaleListDate", targetComp.SalesDate.ToString("MMddyyyy") });
            }
            else
            {
                commands.Add(6, new string[] { inputType, "SaleListDate", targetComp.ListDate.ToString("MMddyyyy") });
            }
            commands.Add(7, new string[] { inputType, "DaysOnMarket", targetComp.DOM});

            if (targetComp.TransactionType == "REO")
            {
                commands.Add(8, new string[] { "SELECT", "BankReoID", "%480" });
            }
            else
            {
                commands.Add(8, new string[] { "SELECT", "BankReoID", "%481" });
            }

            if (targetComp.Waterfront)
            {
                commands.Add(9, new string[] { "SELECT", "ViewID", "%436" });
            }
            else
            {
                commands.Add(9, new string[] { "SELECT", "ViewID", "%437" });
            }
            commands.Add(10, new string[] { inputType, "YearBuilt", targetComp.YearBuiltString });
            commands.Add(11, new string[] { "SELECT", "ConditionID", "%444" });
            commands.Add(12, new string[] { inputType, "TotalRooms", targetComp.TotalRoomCount.ToString() });
            commands.Add(13, new string[] { inputType, "Bedrooms", targetComp.BedroomCount });
            commands.Add(14, new string[] { inputType, "Bathrooms", targetComp.BathroomCount });
            commands.Add(15, new string[] { inputType, "LivingArea", targetComp.ProperGla(GlobalVar.mainWindow.SubjectAboveGLA) });


            if (targetComp.FinishedBasement)
            {
                if(targetComp.BasementType.ToLower() == "partial")
                {
                    commands.Add(16, new string[] { "SELECT", "BasementID", "%549" });
                }
                else
                {
                    commands.Add(16, new string[] { "SELECT", "BasementID", "%482" });
                }
               
            }
            else if (!(targetComp.BasementType.ToLower() == "none"))
            {
                if (targetComp.BasementType.ToLower() == "partial")
                {
                    commands.Add(16, new string[] { "SELECT", "BasementID", "%559" });
                }
                else
                {
                    commands.Add(16, new string[] { "SELECT", "BasementID", "%484" });

                }
            }
            else
            {
                commands.Add(16, new string[] { "SELECT", "BasementID", "%489" });
            }

            if (targetComp.GarageExsists())
            {
                if (targetComp.AttachedGarage())
                {
                    switch (targetComp.intNumberGarageStalls())
                    {
                        case 0:
                            break;
                        case 1:
                            commands.Add(17, new string[] { "SELECT", "GarageID", "%428" });
                            break;
                        case 2:
                             commands.Add(17, new string[] { "SELECT", "GarageID", "%429" });
                            break;
                        case 3:
                            commands.Add(17, new string[] { "SELECT", "GarageID", "%430" });
                            break;
                        case 4:
                            commands.Add(17, new string[] { "SELECT", "GarageID", "%431" });
                            break;
                        default:
                            commands.Add(17, new string[] { "SELECT", "GarageID", "%431" });
                            break;
                    }
                } 
                else if (targetComp.DetachedGarage())
                {
                    switch (targetComp.intNumberGarageStalls())
                    {
                        case 0:
                            break;
                        case 1:
                            commands.Add(17, new string[] { "SELECT", "GarageID", "%432" });
                            break;
                        case 2:
                            commands.Add(17, new string[] { "SELECT", "GarageID", "%433" });
                            break;
                        case 3:
                            commands.Add(17, new string[] { "SELECT", "GarageID", "%434" });
                            break;
                        case 4:
                            commands.Add(17, new string[] { "SELECT", "GarageID", "%435" });
                            break;
                        default:
                            commands.Add(17, new string[] { "SELECT", "GarageID", "%436" });
                            break;
                    }
                }

            }
            else
            {
                commands.Add(17, new string[] { "SELECT", "GarageID", "%550" });
            }

         
                
         
            

             foreach (var c in commands)
             {
                 macro.AppendFormat("TAG POS=1 TYPE={0} FORM=ID:aspnetForm ATTR=NAME:ctl00$cpAppContent$tcWorkQueueDetails$tpForm$ctl01$rpHeader$ctl00$rpComps$ctl0{1}*_{2} CONTENT={3}\r\n", c.Value[0], compNumber.ToString(), c.Value[1], c.Value[2]);
             }


          


            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:./Detail.aspx ATTR=NAME:ctl00$cpAppContent$tcWorkQueueDetails$tpForm$ctl01$rpHeader$ctl00$rpComps$ctl00$txtComp_Address_Line1 CONTENT=2");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:./Detail.aspx ATTR=NAME:ctl00$cpAppContent$tcWorkQueueDetails$tpForm$ctl01$rpHeader$ctl00$rpComps$ctl00$txtComp_Address_Line2 CONTENT=2");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:./Detail.aspx ATTR=NAME:ctl00$cpAppContent$tcWorkQueueDetails$tpForm$ctl01$rpHeader$ctl00$rpComps$ctl00$txtComp_Address_City CONTENT=11");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:./Detail.aspx ATTR=NAME:ctl00$cpAppContent$tcWorkQueueDetails$tpForm$ctl01$rpHeader$ctl00$rpComps$ctl00$ddlComp_Address_State CONTENT=%AK");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:./Detail.aspx ATTR=NAME:ctl00$cpAppContent$tcWorkQueueDetails$tpForm$ctl01$rpHeader$ctl00$rpComps$ctl00$txtComp_Address_ZipCode CONTENT=2222");

            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 30);
            
        } 
    }
}
