using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{



    class Trinity : BPOFulfillment 
    {
        public Trinity()
        { }


        public Trinity(MLSListing m) 
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
            //saleOrListFlag = saleOrList;

            StringBuilder macro = new StringBuilder();
            //x1,x2,x3,x4,x5,x6
            //s1,s2,s3,l1,l2,l3 - 23 field differ
            //902 - address
            //903 - city st zip
            //904 - unknown price field
            //905, 928, 951, 974, 997, 1020  - original list price
            //

            int startingFieldNum = 902;
            int compStep = 23;
            int selectionBoxStep = 54;
            int compNumber = 0;

            if (saleOrList == "sale")
            {
                startingFieldNum = startingFieldNum + compStep * (Convert.ToInt32(targetCompNumber) - 1);
                selectionBoxStep = selectionBoxStep * (Convert.ToInt32(targetCompNumber) - 1);
                compNumber = Convert.ToInt32(targetCompNumber) - 1;
            }
            else
            {
                startingFieldNum = startingFieldNum + compStep * (Convert.ToInt32(targetCompNumber) + 2);
                selectionBoxStep = selectionBoxStep * (Convert.ToInt32(targetCompNumber) + 2);
                compNumber = Convert.ToInt32(targetCompNumber) + 2;
            }

            Dictionary<int, string[]> commands = new Dictionary<int, string[]>();

            //default input type
            string inputType = "INPUT:TEXT";

            commands.Add(0, new string[] { inputType, targetComp.StreetAddress.Replace(" ", "<SP>") });
            commands.Add(1, new string[] { inputType, (targetComp.City + ",IL " + targetComp.Zipcode).Replace(" ", "<SP>") });
            commands.Add(2, new string[] { inputType, targetComp.OriginalListPrice.ToString() });
            commands.Add(3, new string[] { inputType, targetComp.OriginalListPrice.ToString() });
            //905 to 923 - Seller Concessions 
            commands.Add(21, new string[] { inputType, "0" });
            //923 to 906
            commands.Add(4, new string[] { inputType, targetComp.MlsNumber });
            //907 - either sale or list date
            if (saleOrList == "sale")
            {
                commands.Add(5, new string[] { inputType, targetComp.SalesDate.ToShortDateString()} );
            }
            else
            {
                commands.Add(5, new string[] { inputType, targetComp.ListDate.ToShortDateString()} );
            }

            //908
            commands.Add(6, new string[] { inputType, targetComp.DOM });

            //909 - selection box - exsisting structure = 1197, 1251, 1305
            int x = 1197 + selectionBoxStep;
            commands.Add(7, new string[] { "SELECT", "%" + x.ToString() });

            //910 - type ie SFR=1198, 1252 or Condo=1253
            x = 1198 + selectionBoxStep;
            commands.Add(8, new string[] { "SELECT", "%" + x.ToString() });

            //911
            commands.Add(9, new string[] { inputType, targetComp.GLA.ToString() });

            //1286, 1287 - 384 step gap 911 to 1286 - basement sf 
            commands.Add(1286 + compNumber - startingFieldNum , new string[] { inputType, targetComp.BasementGLA() });

            //1295 - 1286 to 1295 - basement finished sf 
            commands.Add(1295 + compNumber - startingFieldNum, new string[] { inputType, targetComp.BasementFinishedGLA() });

            //912 - 1295 to 912 - yearbuilt
            commands.Add(10, new string[] { inputType, targetComp.YearBuiltString });

            //1304 - effective age - skipping for now

            //913 bedrooms - 1=1203, 2=1204, 3=1205
            Dictionary<string, int> bedroomTranslation = new Dictionary<string,int>();
            bedroomTranslation.Add("1", 1203);
            bedroomTranslation.Add("2", 1204);
            bedroomTranslation.Add("3", 1205);
            bedroomTranslation.Add("4", 1206);
            int selectNumber = bedroomTranslation[targetComp.BedroomCount] + selectionBoxStep;
            string valueString = "%" + selectNumber.ToString();

            commands.Add(11, new string[] { "SELECT", valueString });

            //914 full baths - 1=1212, 2=1213...
            Dictionary<string, int> fullBathTranslation = new Dictionary<string, int>();
            fullBathTranslation.Add("1", 1212);
            fullBathTranslation.Add("2", 1213);
            fullBathTranslation.Add("3", 1214);
            fullBathTranslation.Add("4", 1215);
            selectNumber = fullBathTranslation[targetComp.FullBathCount] + selectionBoxStep;
            valueString = "%" + selectNumber.ToString();
            commands.Add(12, new string[] { "SELECT", valueString });

            //915 half baths - 0=1864, 1=1221, 2=1222...
            Dictionary<string, int> halfBathTranslation = new Dictionary<string, int>();
            halfBathTranslation.Add("0", 1864 + compNumber);
            halfBathTranslation.Add("1", 1221 + selectionBoxStep);
            halfBathTranslation.Add("2", 1222 + selectionBoxStep);
            halfBathTranslation.Add("3", 1223 + selectionBoxStep);
            selectNumber = halfBathTranslation[targetComp.HalfBathCount];
            valueString = "%" + selectNumber.ToString();
            commands.Add(13, new string[] { "SELECT", valueString });

            //916 attached stalls - 0=1864, 1=1221, 2=1222...
            Dictionary<string, int> attGarageStallsTranslation = new Dictionary<string, int>();
            attGarageStallsTranslation.Add("0", 1888 + compNumber);
            attGarageStallsTranslation.Add("1", 1226 + selectionBoxStep);
            attGarageStallsTranslation.Add("2", 1227 + selectionBoxStep);
            attGarageStallsTranslation.Add("3", 1228 + selectionBoxStep);
            if (targetComp.AttachedGarage())
            {
                selectNumber = attGarageStallsTranslation[targetComp.NumberGarageStalls()];
            } else
            {
                selectNumber = 1888 + compNumber;
            }
            valueString = "%" + selectNumber.ToString();
            commands.Add(14, new string[] { "SELECT", valueString });

            //917 dettached stalls - 0=1864, 1=1221, 2=1222...
            Dictionary<string, int> detGarageStallsTranslation = new Dictionary<string, int>();
            detGarageStallsTranslation.Add("0", 1896 + compNumber);
            detGarageStallsTranslation.Add("1", 1233 + selectionBoxStep);
            detGarageStallsTranslation.Add("2", 1234 + selectionBoxStep);
            detGarageStallsTranslation.Add("3", 1235 + selectionBoxStep);
            if (targetComp.DetachedGarage())
            {
                selectNumber = detGarageStallsTranslation[targetComp.NumberGarageStalls()];
            }
            else
            {
                selectNumber = 1896 + compNumber;
            }
            valueString = "%" + selectNumber.ToString();
            commands.Add(15, new string[] { "SELECT", valueString });
                
            //1313 carports set to 0
            x = 1934 + compNumber * 8;
            commands.Add(1313 + compNumber - startingFieldNum, new string[] { "SELECT", "%" + x.ToString() });

            //920 lot size
            commands.Add(18, new string[] { inputType, targetComp.Lotsize.ToString() });

            //921 distance to subject\
            commands.Add(19, new string[] { inputType, targetComp.ProximityToSubject.ToString() });

            //922 view comparison
            x = 1241 + selectionBoxStep;
            commands.Add(20, new string[] { "SELECT", "%" + x.ToString() });

            //1330 const quality
            x = 1827 + compNumber * 3;
            commands.Add(1330 + compNumber - startingFieldNum, new string[] { "SELECT", "%" + x.ToString() });

            //1338 current condition 
            x = 1845 + compNumber * 3;
            commands.Add(1338 + compNumber - startingFieldNum, new string[] { "SELECT", "%" + x.ToString() });

            //1347 pool 
            x = 1988 + compNumber * 3;
            commands.Add(1347 + compNumber - startingFieldNum, new string[] { "SELECT", "%" + x.ToString() });

            //1356 fireplace 
            x = 2016 + compNumber * 5;
            commands.Add(1356 + compNumber - startingFieldNum, new string[] { "SELECT", "%" + x.ToString() });

            //1365  other amenities
            commands.Add(1365 + compNumber - startingFieldNum, new string[] { inputType, "NA" });


            //924 REO?
            Dictionary<string, int> distressed = new Dictionary<string, int>();
            distressed.Add("Yes", 1247);
            distressed.Add("No", 1248);
     
            if (targetComp.DistressedSale())
            {
                selectNumber = 1247 + selectionBoxStep;
            }
            else
            {
                selectNumber = 1248 + selectionBoxStep;
            }
            valueString = "%" + selectNumber.ToString();
            commands.Add(22, new string[] { "SELECT", valueString });


            foreach (var c in commands)
            {
                macro.AppendLine(@"TAG POS=1 TYPE=" + c.Value[0] + " FORM=NAME:frmProcessDraw ATTR=NAME:*CustForm_67_" + (startingFieldNum + c.Key).ToString() + " CONTENT=" + c.Value[1]);
            }



            //for (int i = 0; i < 23; i++)
            //{
            //      macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_" + (startingFieldNum + i).ToString() + " CONTENT=" + commands[i]);


            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_903");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_904");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_905");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_923");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_906");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:txtCustForm_67_907");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_908");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_911");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1286");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1295");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_912");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1304");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_920");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_921");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1365");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_925");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_926");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_927");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_928");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_946");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_929");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:txtCustForm_67_930");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_931");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_934");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1287");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1296");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_935");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1305");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_943");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_944");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1366");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_948");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_949");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_950");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_951");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_969");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_952");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:txtCustForm_67_953");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_954");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_957");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1288");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1297");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_958");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1306");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_966");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_967");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1367");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_971");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_972");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_973");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_974");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_992");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_975");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:txtCustForm_67_976");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_977");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_980");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1289");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1298");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_981");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1307");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_989");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_990");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1368");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_994");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_995");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_996");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_997");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1015");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_998");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:txtCustForm_67_999");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1000");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1003");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1290");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1299");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1004");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1308");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1012");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1013");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1369");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1017");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1018");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1019");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1020");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1038");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1021");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:txtCustForm_67_1022");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1023");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1026");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1291");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1300");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1027");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1309");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1035");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1036");
            ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmProcessDraw ATTR=NAME:CustForm_67_1370");
            //}

          
            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 30);
            
        }

           
    }
}
