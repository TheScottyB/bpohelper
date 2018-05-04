using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Net;
using System.IO;
//using //BitMiracle.Docotic.Pdf;
using System.Diagnostics;
using System.Collections;
using System.Diagnostics;
using System.Threading;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Linq;
using System.Reflection;
using ImageResizer;
using iMacros;



namespace bpohelper
{
    public class MRED_Driver
    {
        private readonly IYourForm form;
        public MRED_Driver(IYourForm form)
        {
            this.form = form;
        }
        //private iMacros.App driver;
        private int timeout = 30;

        public void NewMapSearch(iMacros.App d)
        {

            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !TIMEOUT_STEP 30");
            macro.AppendLine(@"SET  !REPLAYSPEED MEDIUM");
            macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=TXT:Search");
            macro.AppendLine(@"FRAME NAME=workspace");
            macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:Click<SP>here<SP>to<SP>select<SP>boundaries");
            // macro.AppendLine(@"TAB T=2");
            macro.AppendLine(@"TAG POS=1 TYPE=SPAN FORM=NAME:dc ATTR=TXT:Center<SP>On...");
            //macro.AppendLine(@"wait seconds=1");

            //hardcoded         
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:dc ATTR=ID:address CONTENT=172<SP>Morningside<SP>Ln<SP>W,<SP>Buffalo<SP>Grove,<SP>IL<SP>60089");

            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:dc ATTR=ID:address CONTENT=" + form.SubjectFullAddress.Replace(" ", "<SP>"));

            //macro.AppendLine(@"wait seconds=1");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=VALUE:Go");
            //macro.AppendLine(@"wait seconds=1");
            macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:dc ATTR=SRC:*/circle.png");
            //macro.AppendLine(@"wait seconds=1");


            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:dc ATTR=ID:distanceInput CONTENT=" + form.SearchMapRadius);
            // macro.AppendLine(@"wait seconds=1");

            macro.AppendLine(@"SAVEAS TYPE=PNG FOLDER=" + form.SubjectFilePath.Replace(" ", "<SP>") + " FILE=search_map_" + DateTime.Now.Ticks.ToString() + ".png");
            // macro.AppendLine(@"wait seconds=1");
            //macro.AppendLine(@"FRAME name=dc");
            //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=NAME:<SP>OK<SP>");
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");
            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET  !REPLAYSPEED FAST");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=ID:<SP>OK<SP>");
            // macro.AppendLine(@"TAB CLOSE");
            //  macro.AppendLine(@"TAB T=1");

            string macroCode = macro.ToString();
            // d.iimOpen("-ie", false, 30);
            iMacros.Status s = d.iimPlayCode(macroCode, 30);

            if (s == iMacros.Status.sOk)
            {
                GlobalVar.mainWindow.buttonSetupSearch.ForeColor = Color.Green;
            }

        }

        public void Save1MiSearch(iMacros.App d)
        {
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"FRAME NAME=workspace");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:dc ATTR=ID:searchButtonTop&&VALUE:View<SP>Results");
            macro.AppendLine(@"FRAME NAME=subheader");
            macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:header ATTR=ID:STATS_BLOCK");
            macro.AppendLine(@"FRAME F=5");
            macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=ID:printLink");
            macro.AppendLine(@"FRAME NAME=workspace");
            macro.AppendLine(@"ONPRINT P=8 BUTTON=PRINT");
            macro.AppendLine(@"TAG POS=1 TYPE=IFRAME FORM=NAME:dc ATTR=ID:statisticsFrame");
            macro.AppendLine(@"FRAME NAME=subheader");
            macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:header ATTR=ID:refineSearchLink");
            string macroCode = macro.ToString();
            d.iimPlayCode(macroCode, timeout);
        }

        public void SaveStats(iMacros.App d, MarketStats m)
        {
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"FRAME NAME=workspace");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:dc ATTR=ID:searchButtonTop&&VALUE:View<SP>Results");
            macro.AppendLine(@"FRAME NAME=subheader");
            macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:header ATTR=ID:STATS_BLOCK");
            macro.AppendLine(@"WAIT SECONDS=1");

            macro.AppendLine(@"FRAME F=5");
            macro.AppendLine(@"TAG POS=3 TYPE=TABLE FORM=NAME:dc ATTR=TXT:* EXTRACT=TXT");
            // macro.AppendLine(@"SAVEAS TYPE=EXTRACT FOLDER=" + form.SubjectFilePath.Replace(" ", "<SP>") +" FILE=mytable_{{!NOW:yymmdd_hhnnss}}.csv");
            macro.AppendLine(@"TAG POS=4 TYPE=TABLE FORM=NAME:dc ATTR=TXT:* EXTRACT=TXT");
            // macro.AppendLine(@"SAVEAS TYPE=EXTRACT FOLDER=" + form.SubjectFilePath.Replace(" ", "<SP>") + " FILE=mytable_{{!NOW:yymmdd_hhnnss}}.csv");
            macro.AppendLine(@"");
            macro.AppendLine(@"WAIT SECONDS=3");

            string macroCode = macro.ToString();
            d.iimPlayCode(macroCode, timeout);

            string t1 = d.iimGetExtract(1);
            string t2 = d.iimGetExtract(2);

            string pattern = "Sold#NEXT##NEXT#(\\d+)#NEXT#";
            Match match = Regex.Match(t1, pattern);
            m.totalSold = match.Groups[1].Value;

            pattern = "Under Contract#NEXT##NEXT#(\\d+)#NEXT#";
            match = Regex.Match(t1, pattern);
            m.totalPending = match.Groups[1].Value;

            pattern = "Active#NEXT##NEXT#(\\d+)#NEXT#";
            match = Regex.Match(t1, pattern);
            m.totalActive = match.Groups[1].Value;

            m.totalOnMarket = (Convert.ToInt16(m.totalActive) + Convert.ToInt16(m.totalPending)).ToString();


            //listprice max, avg,med,min
            pattern = "#NEWLINE#LP\\s*[^#]+#NEXT##NEXT#([^#]+)#NEXT#([^#]+)#NEXT#([^#]+)#NEXT#([^#]+)#NEXT#";
            match = Regex.Match(t2, pattern);
            m.listPrice[0] = match.Groups[1].Value;
            m.listPrice[1] = match.Groups[2].Value;
            m.listPrice[2] = match.Groups[3].Value;
            m.listPrice[3] = match.Groups[4].Value;

            pattern = "#NEWLINE#SP\\s*[^#]+#NEXT##NEXT#([^#]+)#NEXT#([^#]+)#NEXT#([^#]+)#NEXT#([^#]+)";
            match = Regex.Match(t2, pattern);
            m.soldPrice[0] = match.Groups[1].Value;
            m.soldPrice[1] = match.Groups[2].Value;
            m.soldPrice[2] = match.Groups[3].Value;
            m.soldPrice[3] = match.Groups[4].Value;



            //MessageBox.Show(t1);

        }

        private void SetCompSearch(string a)
        {

        }

        public void FindComps(iMacros.App d, bool newSearch)
        {
            //are we on the search screen
            #region setup search screen
            int timeout = 90;
            StringBuilder macro = new StringBuilder();
            string test = "";


            macro.AppendLine(@"SET !TIMEOUT_STEP 30");
            macro.AppendLine(@"FRAME NAME=workspace");
            macro.AppendLine(@"TAG POS=8 TYPE=TD FORM=NAME:dc ATTR=TXT:* EXTRACT=TXT");
            d.iimPlayCode(macro.ToString(), timeout);
            test = d.iimGetLastExtract();

            if (test.Contains("* Status"))
            {
                if (newSearch)
                {
                    macro.AppendLine(@"SET !ERRORIGNORE YES");
                    macro.AppendLine(@"FRAME NAME=workspace");

                    //
                    //First search is the 1 mile radius for neighborhood stats 12 months
                    //

                    if (!GlobalVar.isNeighborhoodSearchComplete)
                    {
                        macro.AppendLine(@"FRAME NAME=workspace");
                        macro.AppendLine(@"TAG POS=6 TYPE=IMG FORM=NAME:dc ATTR=ID:picklisticon");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBMONTHS_BACK7 CONTENT=YES");
                        macro.AppendLine(@"TAG POS=6 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=NAME:<SP>OK<SP>");
                    }





                    //macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:dc ATTR=ID:picklisticon");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST0&&VALUE:ACTV CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST1&&VALUE:AUCT CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST2&&VALUE:BOMK CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST3&&VALUE:CTG CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST4&&VALUE:NEW CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST5&&VALUE:PCHG CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST6&&VALUE:RACT CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST7&&VALUE:TEMP CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST1&&VALUE:CLSD CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBST3&&VALUE:PEND CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=VALUE:OK");



                    //  macro.AppendLine(@"FRAME NAME=workspace");
                    // macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=NAME:dc ATTR=TXT:1<SP>Month2<SP>Months3<SP>Months4<SP>Months5<SP>Months6<SP>Months9<SP>Months12<SP>Months24<SP>MonthsAll<SP>Mo*");
                    //

                    ////six months
                    ////
                    //macro.AppendLine(@"TAG POS=6 TYPE=IMG FORM=NAME:dc ATTR=ID:picklisticon");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBMONTHS_BACK5&&VALUE:6<SP>Months CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=6 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=VALUE:OK");

                    //macro.AppendLine(@"TAG POS=6 TYPE=IMG FORM=NAME:dc ATTR=ID:picklisticon");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBMONTHS_BACK7&&VALUE:12<SP>Months CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=1 TYPE=SPAN FORM=NAME:dc ATTR=ID:1DISP");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBMONTHS_BACK0&&VALUE:6<SP>Months CONTENT=YES");
                    //macro.AppendLine(@"TAG POS=6 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=VALUE:OK");


                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:dc ATTR=ID:MONTHS_BACKID CONTENT=12<SP>Months");
                    //macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:ui-active-menuitem");
                    // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=ID:countButtonTop&&VALUE:Count");
                    // macro.AppendLine(@"TAG POS=0 TYPE=INPUT:TEXT FORM=NAME:dc ATTR=ID:MONTHS_BACKID CONTENT=1<SP>Month");

                    //
                    //attached search
                    //
                    macro.AppendLine(@"TAG POS=5 TYPE=IMG FORM=NAME:dc ATTR=ID:picklisticon");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:CBMONTHS_BACK5&&VALUE:6<SP>Months CONTENT=YES");
                    macro.AppendLine(@"TAG POS=5 TYPE=INPUT:BUTTON FORM=NAME:dc ATTR=VALUE:OK");
                    macro.AppendLine(@"SAVEAS TYPE=PNG FOLDER=" + form.SubjectFilePath.Replace(" ", "<SP>") + " FILE=search_page_" + DateTime.Now.Ticks.ToString() + ".png");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:dc ATTR=ID:searchButtonTop&&VALUE:View<SP>Results");
                    macro.AppendLine(@"TAG POS=11 TYPE=SPAN FORM=NAME:dc ATTR=CLASS:columnHeader");
                    macro.AppendLine(@"FRAME NAME=subheader");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:header ATTR=ID:report_type CONTENT=%agentfull");
                }
                else
                {
                    macro.AppendLine(@"SAVEAS TYPE=PNG FOLDER=" + form.SubjectFilePath.Replace(" ", "<SP>") + " FILE=search_page_" + DateTime.Now.Ticks.ToString() + ".png");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:dc ATTR=ID:searchButtonTop&&VALUE:View<SP>Results");
                    macro.AppendLine(@"TAG POS=11 TYPE=SPAN FORM=NAME:dc ATTR=CLASS:columnHeader");
                    macro.AppendLine(@"FRAME NAME=subheader");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:header ATTR=ID:report_type CONTENT=%agentfull");
                    //macro.AppendLine(@"TAG POS=2 TYPE=SPAN FORM=NAME:dc ATTR=CLASS:Label EXTRACT=TXT");
                }

                string macroCode = macro.ToString();
                d.iimPlayCode(macroCode, timeout);
            }
            else
            {
                macro.Clear();
                macro.AppendLine(@"FRAME NAME=subheader");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:header ATTR=ID:report_type EXTRACT=TXT");
                d.iimPlayCode(macro.ToString(), timeout);

                test = d.iimGetLastExtract();
                //is the list of properties already open
                if (!test.Contains("Full - Agent"))
                {
                    //if not then break
                    MessageBox.Show("Search Not Setup");
                    return;
                }
            }
            #endregion

            //run search
            MlsDrivers r = new MlsDrivers(form);
            r.ReadMlsSheets(d);
        }

    }

    public class MlsDrivers
    {
        public MlsDrivers()
        {

        }
        private readonly IYourForm form;
        public MlsDrivers(IYourForm form)
        {
            this.form = form;

        }
        string listPrice;
        public string mlsStatus;
        public string assessmentIncludes;
        public string assessmentContactPhone;
        public string assessmentContact;
        public string assessmentFrequncy;
        public string assessmentAmount;
        public string roomCount;
        public string bedroomCount;
        public string bathCount;
        public string fullBathCount;
        public string halfBathCount = "0";
        public string basementType;
        public string basementDetails;
        public string numFireplaces = "0";
        public string parking;
        public string style;
        public string type;
        public bool attached;
        public bool detached;
        public string listAgent;
        public string brokerPhone;
        public string dom;
        public string mlsTaxAmount;

        private string rawText;

        private string GetFieldValue(string fn)
        {
            //:x*(.*?)xxx


            string pattern1 = "x*(.*?)xxx";
            string pattern2 = "x*(.*?)\n";
            string returnValue = "NotFound";
            string result = "";

            //result = Regex.Match(rawText, string.Format(@"{0}{1}", fn, pattern1)).Groups[1].Value;
            result = Regex.Match(rawText, string.Format(@"{0}{1}|{0}{2}", fn, pattern1, pattern2)).Groups[1].Value;

            if (String.IsNullOrWhiteSpace(result))
            {
                returnValue = Regex.Match(rawText, string.Format(@"{0}{1}", fn, pattern2)).Groups[1].Value;
            }
            else
            {
                returnValue = result;
            }

            return returnValue.Replace("$", "").Replace(",", "");

            //return Regex.Match(rawText, string.Format(@"{0}{1}|{0}{2}", fn, pattern1, pattern2)).Groups[1].Value;
        }



        public string Get_Distance(string o, string d)
        {


            //
            //new way to reduce google map webservice calls
            //

            //Position pos1 = new Position();
            //pos1.Latitude = Convert.ToDouble(hotelx.latitude);
            //pos1.Longitude = Convert.ToDouble(hotelx.longitude);
            //Position pos2 = new Position();
            //pos2.Latitude = Convert.ToDouble(lat);
            //pos2.Longitude = Convert.ToDouble(lng);
            //Haversine calc = new Haversine();

            //double result = calc.Distance(pos1, pos2, DistanceType.Miles);


            //
            //end new way
            //\\

            //Google distance matric webservice
            //http://maps.googleapis.com/maps/api/distancematrix/xml?origins=60050&destinations=60002&sensor=false&units=imperial
            //
            //string zip = "60085";





            string distance = "0";
            //Lot#104 --> Lot 104
            string googlestr = "http://maps.googleapis.com/maps/api/distancematrix/xml?origins=" + o.Replace("#", " ") + "&destinations=" + d + "&sensor=false&units=imperial";

            // Create a request for the URL. 		
            WebRequest request = WebRequest.Create(googlestr);

            // Get the response.
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();


            // Get the stream containing content returned by the server.
            Stream dataStream = response.GetResponseStream();
            // Open the stream using a StreamReader for easy access.
            StreamReader reader = new StreamReader(dataStream);
            // Read the content.
            string responseFromServer = reader.ReadToEnd();
            // Display the content.

            //MessageBox.Show(responseFromServer);

            //<distance>
            //<value>42315</value>
            //<text>26.3 mi</text>
            //</distance>

            string pattern = "(\\d+.\\d*) [mf]";
            Match match = Regex.Match(responseFromServer, pattern);
            distance = match.Groups[1].Value;
            if (responseFromServer.Contains(" ft</text>"))
            {
                distance = "0.05";
            }
            //MessageBox.Show(distance);
            //Console.WriteLine(responseFromServer);
            // Cleanup the streams and the response.
            reader.Close();
            dataStream.Close();
            response.Close();

            return distance;
        }
        public double Get_Distance(GeoPoint o, GeoPoint d)
        {
            //
            //new way to reduce google map webservice calls
            //and return "crow flys" distance
            //

            //Position pos1 = new Position();
            //pos1.Latitude = Convert.ToDouble(hotelx.latitude);
            //pos1.Longitude = Convert.ToDouble(hotelx.longitude);
            //Position pos2 = new Position();
            //pos2.Latitude = Convert.ToDouble(lat);
            //pos2.Longitude = Convert.ToDouble(lng);
            Haversine calc = new Haversine();

            double result = calc.Distance(o, d, DistanceType.Miles);

            return result;
        }
        private int Age(string yearBuilt)
        {
            DateTime subject_age = new DateTime((Convert.ToInt32(yearBuilt)), 1, 1);

            TimeSpan ts = DateTime.Now - subject_age;

            return ts.Days / 365;
        }
        public void GetSubjectInfo(string s)
        {
            MLSListing m = new MLSListing();

            //160 vs 32, non-breaking space issue in pdf translations
            s = s.Replace(" ", " ");
            
            rawText = s;


            string masterPattern = @"xxx(.*?)xxx|xxx(.*?)\n";
            MatchCollection x = Regex.Matches(s, masterPattern);

            GlobalVar.theSubjectProperty.PrintedMlsSheetNameValuePairs = x;
            string listDate = "";
            string p = @"List\s*Date:\s*(\d+.\d+.\d\d\d\d)";
            MatchCollection myMc = Regex.Matches(s, p);
            if (myMc.Count > 0)
            {
                 listDate = myMc[0].Groups[1].Value;

            }
           


            string pattern = @"Ax*mount:\$(\d+,*\d*\.*\d*)";
            MatchCollection mc = Regex.Matches(s, pattern);
            try
            {
                assessmentAmount = mc[0].Groups[1].Value;
            }
            catch
            {
                assessmentAmount = "-1";
            }

            try
            {
                mlsTaxAmount = mc[1].Groups[1].Value;
            }
            catch
            {
                mlsTaxAmount = "-1";
            }

            string oorPhone = "";
            brokerPhone = "";
            string coListerPhone = "";
            //pattern = "Ph #:x*([^\\nx]+)|Ph #:x*([^\\n]+)";
            pattern = @"xxxPh #:(.*?)xxx|xxxPh #:(.*?)\n";
            mc = Regex.Matches(s, pattern);
            if (mc.Count > 0)
            {
                 oorPhone = mc[0].Groups[1].Value;
                if (mc[1].Groups[1].Value.Length < 10)
                {
                    brokerPhone = mc[1].Groups[2].Value;
                } else
                {
                    brokerPhone = mc[1].Groups[1].Value;
                }
               
                 coListerPhone = mc[2].Groups[1].Value;
            }
        
            pattern = "Lst. Mkt. Time:x*([^\\nx]+)|Lst. Mkt. Time:x*([^\\n]+)";
            Match match = Regex.Match(s, pattern);
            dom = match.Groups[1].Value;

            pattern = "List Agent:x*([^\\nx]+)|List Agent:x*([^\\n]+)";
            match = Regex.Match(s, pattern);
            listAgent = match.Groups[1].Value;



            pattern = "List Price:x*([^\\nx]+)|List Price:x*([^\\n]+)";
            match = Regex.Match(s, pattern);
            listPrice = match.Groups[1].Value;

            pattern = @"(ACTV|CTG|CLSD|TEMP|CANC|PCHG|EXP)\n";
            match = Regex.Match(s, pattern);
            mlsStatus = match.Groups[1].Value;



            ////pattern = "Status:x*([^\\nx]+)|Status:x*([^\\n]+)";
            ////match = Regex.Match(s, pattern);
            ////mlsStatus = match.Groups[1].Value;



            pattern = "Asmt Incl:x*([^\\nx]+)|Asmt Incl:x*([^\\n]+)";
            match = Regex.Match(s, pattern);
            assessmentIncludes = match.Groups[1].Value;

            pattern = "Phone:x*([^\\nx]+)|Phone:x*([^\\n]+)";
            match = Regex.Match(s, pattern);
            assessmentContactPhone = match.Groups[1].Value;

            pattern = "Contact Name:([^x]+)";
            match = Regex.Match(s, pattern);
            assessmentContact = match.Groups[1].Value;

            pattern = "Frequency:([^ ]+)";
            match = Regex.Match(s, pattern);
            assessmentFrequncy = match.Groups[1].Value;

            pattern = "Rooms:([^x]+)";
            match = Regex.Match(s, pattern);
            roomCount = match.Groups[1].Value;


            pattern = @"Detached\s+Single";
            match = Regex.Match(s, pattern);
            if (match.Success)
            {
                m = new DetachedListing();
                detached = true;
                attached = false;
            }

            pattern = @"Ax*ttached\s+Single";
            match = Regex.Match(s, pattern);
            if (match.Success)
            {
                m = new AttachedListing();
                detached = false;
                attached = true;
            }


            //Bathrooms2 /xxx
            //Bathrooms 1 / 1
            pattern = "Bathrooms\\s*(\\d+)";
            match = Regex.Match(s, pattern);
            fullBathCount = match.Groups[1].Value;

            pattern = "Bathrooms\\s*\\d\\s.\\s(\\d+)";
            match = Regex.Match(s, pattern);
            if (match.Success)
            {
                halfBathCount = match.Groups[1].Value;
            }

            bathCount = fullBathCount + "." + halfBathCount;

            //Bedrooms:3xxx
            pattern = "Bedrooms:(\\d+)";
            match = Regex.Match(s, pattern);
            bedroomCount = match.Groups[1].Value;

            pattern = "Basement:([^x]+)";
            match = Regex.Match(s, pattern);
            basementType = match.Groups[1].Value;
            //Basement Details: Partially Finished
            //Basement Details:Partially Finished

            //this works, but one saved match will be ""
            //pattern = "Basement Details:x*([^\\n]+)xxx|Basement Details:x*([^\\n]+";

            pattern = "Basement Details:(.*?)xxx|Basement Details:(.*?)\\n";


            match = Regex.Match(s, pattern);
            if (match.Groups[1].Value != "")
            {
                basementDetails = match.Groups[1].Value;
            }
            else
            {
                basementDetails = match.Groups[2].Value;
            }


            //Fireplaces:xxx1
            //Fireplaces:0
            pattern = "Fireplaces:\\w*(\\d+)";
            match = Regex.Match(s, pattern);
            numFireplaces = match.Groups[1].Value;
            if (string.IsNullOrWhiteSpace(numFireplaces))
            {
                numFireplaces = "0";
            }



            //Parking:Garage
            //Parking:Exterior Space(s)
            pattern = "Parking:x*([^\\n]+)";
            match = Regex.Match(s, pattern);
            parking = match.Groups[1].Value;

            if (match.Success)
            {


                //Spaces:Gar:2
                pattern = "Spaces:Gar:(\\d+)";
                match = Regex.Match(s, pattern);
                parking = parking + " " + match.Groups[1].Value;

                //pattern = "Garage Type:x*([^x\\n]+)xxx";
                pattern = @"Garage\s+Type:x*(Attached|Detached)";
                match = Regex.Match(s, pattern);
                parking = parking + " " + match.Groups[1].Value;



            }
            else
            {
                pattern = "Spaces:Gar:(\\d+)";
                match = Regex.Match(s, pattern);
                parking = parking + " " + match.Groups[1].Value;


                //Parking Type:xxxParking Avail
                //Parking Type:xxxDetached Garage
                pattern = "Parking Type:x*([^x\\n]+)";
                match = Regex.Match(s, pattern);
                parking = parking = parking + " " + match.Groups[1].Value;
            }

            //pattern = @"xxxStyle:(.*?)xxx|xxxStyle:(.*?)\n";
            //pattern = @"Style:x*([^x\nG]+)x*";
            pattern = @"Style:(.*?)xxx|Style:(.*?)\n";
            match = Regex.Match(s, pattern);
            style = match.Groups[1].Value;

            //required field, except in archive listings, ie
            //Type:xxxGarage Ownership:xxxSewer:
            //pattern = "Type:x*([^x\\n]+)xxx";
            //above patern fails
            pattern = @"Type:x*([^x\nG]+)x*";

            match = Regex.Match(s, pattern);
            string type = match.Groups[1].Value;

            if (!match.Success)
            {
                //bad print from pdf convertor
                pattern = @"Tx*yx*px*e:x*(Tx*ownhouse)";
                match = Regex.Match(s, pattern);
                type = match.Groups[1].Value.Replace("x", "");
            }


            form.SubjectBathroomCount = bathCount;
            form.SubjectBedroomCount = bedroomCount;
            form.SubjectRoomCount = roomCount;
            form.SubjectBasementType = basementType;
            form.SubjectBasementDetails = basementDetails;
            form.SubjectNumberFireplaces = numFireplaces;
            form.SubjectParkingType = parking;
            form.SubjectStyle = style;
            form.SubjectMlsType = type;
            form.SubjectAttached = attached;
            form.SubjectDetached = detached;
            form.SubjectAssessmentInfo.amount = assessmentAmount;

            form.SubjectAssessmentInfo.frequency = assessmentFrequncy;
            form.SubjectAssessmentInfo.includes = assessmentIncludes;
            form.SubjectAssessmentInfo.managementContact = assessmentContact;
            form.SubjectAssessmentInfo.managementPhone = assessmentContactPhone;
            form.SubjectMlsStatus = mlsStatus;
            form.SubjectCurrentListPrice = listPrice;
            form.SubjectListingAgent = listAgent;
            form.SubjectBrokerPhone = brokerPhone;
            form.SubjectDom = dom;

            GlobalVar.theSubjectProperty.PropertyTax = mlsTaxAmount;
            GlobalVar.theSubjectProperty.mlsStatus = mlsStatus;
            GlobalVar.theSubjectProperty.origListingPrice = GetFieldValue(@"Orig List Price:");
            GlobalVar.theSubjectProperty.rawtextFromPdfActiveListing = s;

            string mlsNum = GetFieldValue(@"MLS #:");


            if (mlsNum != "NotFound")
            {
                m.Status = mlsStatus;
                m.MlsNumber = GetFieldValue(@"MLS #:");
                GlobalVar.mainWindow.labelMlsNumber.Text = m.MlsNumber;
                m.ListingBrokerageName = GetFieldValue(@"Broker:");
                GlobalVar.mainWindow.textBoxSubjextListingBrokerage.Text = m.ListingBrokerageName;
                m.ListDateString = GetFieldValue(@"List Date:");
                GlobalVar.mainWindow.subjectListDatedateTimePicker.Value = Convert.ToDateTime(m.ListDateString);
                m.OffMarketDateString = GetFieldValue(@"Off Market:");
                try
                {
                    m.OriginalListPrice = Convert.ToDouble(GetFieldValue(@"Orig List Price:"));
                    GlobalVar.mainWindow.textBoxSubjectOriginalListPRice.Text = m.OriginalListPrice.ToString();
                }
                catch
                {

                }
                try
                {
                    m.CurrentListPrice = Convert.ToDouble(GetFieldValue(@"List Price:"));
                    GlobalVar.mainWindow.SubjectCurrentListPrice = m.CurrentListPrice.ToString();
                }
                catch
                {

                }

                GlobalVar.theSubjectProperty.ParcelID = Regex.Match(GetFieldValue("PIN:"), @"\d+").Value;
                GlobalVar.theSubjectProperty.AddMlsListing(m);

                GlobalVar.mainWindow.labelMlsNumber.Text = m.MlsNumber;
                form.SubjectPin = GlobalVar.theSubjectProperty.ParcelID;
            }
        }


        public async void ReadMlsSheets(iMacros.App d)
        {
            form.StatusUpdate = this.ToString() + "-->ReadMlsSheets";

           // iMacros.App d = browser1;

           // d = browser1;
           // uint browerPid = d.iimGetBrowserPid();
            
            GoogleFusionTable realist_bpohelper = new GoogleFusionTable("1UKrOVmhPWrgLP5d5bDCsiW9whMIK8aLxKhcyOaI");

            await realist_bpohelper.helper_OAuthFusion();

            GoogleCloudDatastore gcds = new GoogleCloudDatastore();
            await gcds.DataStoreTester();


            

            List<string> comps = new List<String>();
            List<MLSListing> listings = new List<MLSListing>();
            Status status = iMacros.Status.sOk;

            RealProperty currentProperty = new RealProperty();
            Neighborhood currentNeighborhood = new Neighborhood();
            Neighborhood setOfComps = new Neighborhood();
            List<RealProperty> rpl = new List<RealProperty>();
            List<RealProperty> rplComps = new List<RealProperty>();



            //extracts  all the html
            form.SetStatusBar = "Reading MLS sheet...";

            StringBuilder extractHtml = new StringBuilder();
            extractHtml.AppendLine(@"FRAME NAME=workspace");
            extractHtml.AppendLine(@"TAG POS=1 TYPE=DIV FORM=NAME:dc ATTR=CLASS:report EXTRACT=HTM");
            var myStatus = d.iimPlayCode(extractHtml.ToString());

            string htmlCode;// = d.iimGetLastExtract();

            MLSListing currentListing; // = new MLSListing(htmlCode);




            List<decimal> activePriceList = new List<decimal>();
            List<decimal> soldPriceList = new List<decimal>();
            List<decimal> picDiffList = new List<decimal>();
            List<int> domList = new List<int>();

            List<decimal> compActvPriceList = new List<decimal>();
            List<decimal> compSoldPriceList = new List<decimal>();

            List<int> compDomList = new List<int>();
            List<int> compActvDomList = new List<int>();
            List<int> compSoldDomList = new List<int>();
            List<int> compAgeList = new List<int>();
            List<int> compActvAgeList = new List<int>();
            List<int> compSoldAgeList = new List<int>();

            int shortSales = 0;
            int reoSales = 0;
            int shortActive = 0;
            int reoActive = 0;

            int compShortClosed = 0;
            int compReoClosed = 0;
            int compShortActive = 0;
            int compReoActive = 0;

            int numberSameTypeAsSubject = 0;
            int numberComparableGla = 0;
            int numberComparableAge = 0;
            int numberComparableLot = 0;

            int pageNumber = 1;


            string min_sale;
            string max_sale;
            string min_list;
            string max_list;

            int numListingsActive = 0;
            int numListingsClosed = 0;
            int numListingsPending = 0;

            int numCompsActive = 0;
            int numCompsClosed = 0;
            int numCompsPending = 0;

            ArrayList list = new ArrayList();
            //IEnumerable<decimal> myEnumerable;
            List<decimal> pricePerSfList = new List<decimal>();
            List<int> ageList = new List<int>();
            Dictionary<string, int> subdivisionList = new Dictionary<string, int>();
            Dictionary<string, int> mlsTypeList = new Dictionary<string, int>();
            Dictionary<string, int> listingAgentList = new Dictionary<string, int>();

            string attOrDet = "Unknown";

            bool stillRecords = true;
            bool searchCache = false;
            int count = 0;

            GlobalVar.searchCacheMlsListings.Clear();
            GlobalVar.searchCacheMlsListings = GlobalVar.listingsFromLastSearch;

            //if (GlobalVar.searchCacheMlsListings.Count > 0)
            //{
            //    searchCache = true;
            //}

            if (form.CacheSearch)
            {
                searchCache = true;
            }

            iMacros.Status s = status;

            //first record
            //s = d.iimPlayCode(macro12.ToString());

            //while (status == iMacros.Status.sOk)
            //for (int i = 0; i < 6; i++)

            #region paging
            while (stillRecords)
            {

               // iMacros.App d = new App();

             //   d = browser1;
              
                //
                //this gets the all html code for current listing page
                //
                if (searchCache)
                {
                    htmlCode = GlobalVar.searchCacheMlsListings[count].rawData;

                }
                else
                {
                    //s = d.iimPlayCode(extractHtml.ToString());
                    htmlCode = d.iimGetLastExtract();
                    if (htmlCode.Contains("#EANF#"))
                    {
                        break;
                    }
                }

                //

                //
                //this creates correct listing object and fills/sets known field/value pairs based on above html
                //
                if (form.SubjectAttached)
                {
                    currentListing = new AttachedListing(htmlCode);
                    //currentListing.proximityToSubject = Convert.ToDouble(Get_Distance(currentListing.mlsHtmlFields["address"].value, form.SubjectFullAddress));
                    listings.Add(currentListing);
                }
                else if (form.SubjectDetached)
                {
                    //constructor parses and sets all field/value pairs
                    currentListing = new DetachedListing(htmlCode);
                    listings.Add(currentListing);
                    ////do we have the realist record for the parcel number, if so, do we have lng lat pair stored?
                    //if (realist_bpohelper.RecordExists(currentListing.mlsHtmlFields["Parcel ID:"].value))
                    //{

                    //}
                    //else
                    //{

                    //    try
                    //    {
                    //        currentListing.proximityToSubject = Convert.ToDouble(Get_Distance(currentListing.mlsHtmlFields["address"].value, form.SubjectFullAddress));
                    //    }
                    //    catch
                    //    {
                    //        currentListing.proximityToSubject = -1;
                    //    }
                    //    listings.Add(currentListing);
                    //}
                }
                else
                {
                    currentListing = new MLSListing(htmlCode);
                    currentListing.proximityToSubject = Convert.ToDouble(Get_Distance(currentListing.mlsHtmlFields["address"].value, form.SubjectFullAddress));
                    listings.Add(currentListing);
                }

                string mlsStatusCheck = currentListing.mlsHtmlFields["status"].value;

                if (mlsStatusCheck == "AUCT")
                {
                    if (searchCache)
                    {
                        count++;
                        if (count >= GlobalVar.searchCacheMlsListings.Count)
                        {
                            stillRecords = false;
                        }
                    }
                    else
                    {
                        StringBuilder move_through_comps_macro = new StringBuilder();
                        //GlobalVar.searchCacheMlsListings.Add(currentListing);
                        string tagPosition = "2";
                        if (pageNumber == 1)
                        {
                            tagPosition = "1";
                        }
                        move_through_comps_macro.AppendLine(@"FRAME NAME=navpanel");

                        move_through_comps_macro.AppendLine(@"TAG POS=" + tagPosition + " TYPE=DIV ATTR=onclick:show*");
                    //    move_through_comps_macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=TXT:Next");
                        move_through_comps_macro.AppendLine(@"FRAME NAME=workspace");

                        move_through_comps_macro.AppendLine(@"TAG POS=1 TYPE=DIV FORM=NAME:dc ATTR=CLASS:report EXTRACT=HTM");
                        status = d.iimPlayCode(move_through_comps_macro.ToString(), 60);
                        if (status != Status.sOk)
                        {
                            stillRecords = false;
                        }
                        pageNumber++;
                    }
                    continue;
                }

                //htmlcode is too big
                //mlsListingCache.AddMlsRecord(currentListing.MlsNumber, htmlCode);

                currentProperty.AddMlsListing(currentListing);


                //  string tempdata = currentListing.rawData;

                // currentListing.rawData = "";

                form.StatusUpdate = "Reading--> " + currentListing.MlsNumber;
                #region read_mls_sheet
                #region Table1
                //
                //Table 1
                //
                string pattern = "";
                Match match;
                //// Detached Single#NEXT#MLS #:#NEXT#07933784#NEXT#List Price:#NEXT#$150,000#NEXT##NEWLINE#Status:#NEXT#CLSD#NEXT#List Date:#NEXT#10/27/2011#NEXT#Orig List Price:#NEXT#$150,000#NEXT##NEWLINE#Area:#NEXT#50  #NEXT#List Dt Rec:#NEXT#10/28/2011#NEXT#Sold Price:#NEXT#$140,000#NEXT##NEWLINE#Address:#NEXT#2316 W Fairview Ln , McHenry, Illinois 60051#NEXT##NEWLINE#Directions:#NEXT#Rt. 120 to River Rd N. to Lincoln E to Hillside to Fairview E#NEXT##NEWLINE#Sold by:#NEXT#Ed Kanabay (16107) / Results Realty USA (1886) #NEXT#Lst. Mkt. Time:#NEXT#18#NEXT##NEWLINE#Closed:#NEXT#02/15/2012#NEXT#Contract:#NEXT#11/13/2011#NEXT#Points:#NEXT##NEXT##NEWLINE#Off Market:#NEXT#11/13/2011#NEXT#Financing:#NEXT#Conventional#NEXT#Contingency:#NEXT##NEXT##NEWLINE#Year Built:#NEXT#1977#NEXT#Blt Before 78:#NEXT#Yes#NEXT#Curr. Leased:#NEXT#No#NEXT##NEWLINE#Dimensions:#NEXT#80X150#NEXT##NEWLINE#Ownership:#NEXT#Fee Simple#NEXT#Subdivision:#NEXT#Eastwood Manor#NEXT#Model:#NEXT##NEXT##NEWLINE#Corp Limits:#NEXT#Unincorporated#NEXT#Township:#NEXT#Mchenry#NEXT#County:#NEXT#Mc Henry#NEXT##NEWLINE#Coordinates:#NEXT#N:33 W:31 #NEXT## Fireplaces:#NEXT##NEXT##NEWLINE#Rooms:#NEXT#8#NEXT#Bathrooms (full/half):#NEXT#3 / 0#NEXT#Parking:#NEXT#Garage#NEXT##NEWLINE#Bedrooms:#NEXT#3#NEXT#Master Bath:#NEXT#Full#NEXT## Spaces:#NEXT#Gar:2 #NEXT##NEWLINE#Basement:#NEXT#None#NEXT#Bsmnt. Bath:#NEXT#No#NEXT#Parking Incl. In Price:#NEXT#Yes
                //string pattern = "MLS #:#NEXT#(\\d+)#NEXT#";
                //Match match = Regex.Match(table1, pattern);
                //string mlsnum = match.Groups[1].Value;
                //
                //pattern = "Subdivision:#NEXT#([^#]+)";
                //match = Regex.Match(table1, pattern);
                //string mls_subdivision = match.Groups[1].Value;
                //doesnt match $1,000,000
                //pattern = "List Price:#NEXT#(.[0-9]+.[0-9]+)";
                //pattern = "List Price:#NEXT#([^#]+)";
                //match = Regex.Match(table1, pattern);
                //string current_list_price = match.Groups[1].Value;
                //
                //pattern = "List Date:#NEXT#(\\d+.\\d+.\\d+)#NEXT#";
                //match = Regex.Match(table1, pattern);
                //string list_date = match.Groups[1].Value;
                //

                //pattern = "Orig List Price:#NEXT#([^#]+)";
                //match = Regex.Match(table1, pattern);
                //string orig_list_price = match.Groups[1].Value;

                //pattern = "Sold Price:#NEXT#([^#]+)";
                //match = Regex.Match(table1, pattern);
                //string sold_price = match.Groups[1].Value;
                //Old way above
                //
                //New Way below
                //

                string mlsnum = currentListing.MlsNumber;
                string mls_subdivison = currentListing.mlsHtmlFields["subdivision"].value;
                string current_list_price = currentListing.mlsHtmlFields["listPrice"].value;
                string list_date = currentListing.mlsHtmlFields["listDate"].value;
                string orig_list_price = currentListing.mlsHtmlFields["origListPrice"].value;


                string sold_price = currentListing.mlsHtmlFields["soldPrice"].value;
                if (String.IsNullOrWhiteSpace(sold_price))
                    sold_price = "0";

                string sale_type = "Normal";
                if (sold_price.Contains("(S)"))
                {
                    sale_type = "Short";
                    sold_price = sold_price.Replace("(S)", "");
                    shortSales++;
                }

                if (sold_price.Contains("(F)"))
                {
                    sale_type = "REO";
                    sold_price = sold_price.Replace("(F)", "");
                    reoSales++;
                }

                if (sold_price.Contains("(C)"))
                {
                    sale_type = "Court";
                    sold_price = sold_price.Replace("(C)", "");
                    reoSales++;
                }

                #region address

                //fuill line address
                //Address:#NEXT#2627 Sycamore Dr , Waukegan, Illinois 60085#NEXT#

                // pattern = "Address:#NEXT#([^#]+)#NEXT#";
                // match = Regex.Match(table1, pattern);
                //string address = match.Groups[1].Value;
                string address = currentListing.mlsHtmlFields["address"].value;
                string[] tempstrarry = address.Split(',');
                string street_number = "";
                string street_name = "";
                string street_postfix = "";
                string city = tempstrarry[1];
                string zip = tempstrarry[2].Split(' ')[2];
                string full_street_address = tempstrarry[0];
                string[] street_combo = tempstrarry[0].Split(' ');


                ////
                ////geocode address
                ////ex:
                ////http://maps.googleapis.com/maps/api/geocode/json?address=1600+Amphitheatre+Parkway,+Mountain+View,+CA&sensor=true_or_false


                //string googlestr = @" http://maps.googleapis.com/maps/api/geocode/json?address=" + address.Replace(" ", "+") + "&sensor=false";


                ////// Create a request for the URL. 		
                //WebRequest request = WebRequest.Create(googlestr);

                ////// Get the response.
                //HttpWebResponse response = (HttpWebResponse)request.GetResponse();


                //// Get the stream containing content returned by the server.
                //Stream dataStream = response.GetResponseStream();
                //// Open the stream using a StreamReader for easy access.
                //StreamReader reader = new StreamReader(dataStream);
                //// Read the content.
                //string responseFromServer = reader.ReadToEnd();
                //// Display the content.
                ////MessageBox.Show(response.ContentLength.ToString());
                ////MessageBox.Show(responseFromServer);
                //////Console.WriteLine(responseFromServer);
                ////// Cleanup the streams and the response.

                //reader.Close();
                //dataStream.Close();
                //response.Close();


                ////ex response:
                ////  "formatted_address" : "107 Glenwood Dr, Round Lake Beach, IL 60073, USA",
                ////  "geometry" : {
                ////   "location" : {
                ////   "lat" : 42.3653120,
                ////   "lng" : -88.08547999999999

                //pattern = @"lat. : (\d+.\d+)";
                //match = Regex.Match(responseFromServer, pattern);
                //string propertyLatitude =  match.Groups[1].Value;

                // pattern = @"lng. : (-\d+.\d+)";
                // match = Regex.Match(responseFromServer, pattern);
                //string propertyLongitude =  match.Groups[1].Value;

                // string propertyLatitude = realist_bpohelper.curRec["Latitude"];
                //  string propertyLongitude = realist_bpohelper.curRec["Longitude"];





                //need to remove unit in attached listing addresses.
                if (tempstrarry[0].Contains("Unit"))
                {
                    tempstrarry[0] = tempstrarry[0].Remove(tempstrarry[0].IndexOf("Unit"));
                }




                //2316 W Fairview Ln , McHenry, Illinois 60051
                pattern = "^(\\d+)\\s+\\w\\s+(\\w+)\\s+(\\w+)";
                match = Regex.Match(tempstrarry[0], pattern);

                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;

                }

                //2316 Fairview Ln , McHenry, Illinois 60051
                pattern = "^(\\d+)\\s+(\\w\\w+)\\s+(\\w+)";
                match = Regex.Match(tempstrarry[0], pattern);
                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;
                    if (street_postfix.Contains("Bay"))
                    {
                        street_postfix = "";
                    }

                }

                //1309 N Chapel Hill Rd , McHenry, Illinois 60051
                pattern = "^(\\d+)\\s+\\w\\s+(\\w+\\s+\\w+)\\s+(\\w+)";
                match = Regex.Match(tempstrarry[0], pattern);
                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;
                }


                //769 White Pine Cir , Lake In The Hills, Illinois 60156
                pattern = "^(\\d+)\\s+(\\w\\w+\\s+\\w+)\\s+(\\w+)[^\\w+]";
                match = Regex.Match(tempstrarry[0], pattern);
                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;
                }

                // 21727 NORTH Hickory Hill Dr 
                pattern = "^(\\d+)\\s+NORTH\\s+(\\w\\w+\\s+\\w+)\\s+(\\w+)";
                match = Regex.Match(tempstrarry[0], pattern);
                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;
                }



                //if (street_combo.Length == 4)
                //{
                //    //2316 W Fairview Ln , McHenry, Illinois 60051
                //    pattern = "^(\\d+)\\s+\\w+\\s+(\\w+)\\s+(\\w+)\\s+,\\s+(\\w+),\\s+\\w+\\s+(\\d+)";
                //    match = Regex.Match(address, pattern);
                //    street_number = match.Groups[1].Value;
                //     street_name = match.Groups[2].Value;
                //     street_postfix = match.Groups[3].Value;
                //     city = match.Groups[4].Value;
                //     zip = match.Groups[5].Value;
                //}
                //else if (street_combo.Length == 5)
                //{
                //    //1309 N Chapel Hill Rd , McHenry, Illinois 60051
                //    pattern = "^(\\d+)\\s+\\w+\\s+(\\w+\\s+\\w+)\\s+(\\w+)\\s+,\\s+(\\w+),\\s+\\w+\\s+(\\d+)";
                //    match = Regex.Match(address, pattern);
                //    street_number = match.Groups[1].Value;
                //    street_name = match.Groups[2].Value;
                //    street_postfix = match.Groups[3].Value;
                //    city = match.Groups[4].Value;
                //    zip = match.Groups[5].Value;

                //}
                //else
                //{

                //    //2316 Fairview Ln , McHenry, Illinois 60051
                //    pattern = "^(\\d+)\\s+(\\w+)\\s+(\\w+)\\s+,\\s+(\\w+),\\s+\\w+\\s+(\\d+)";
                //    match = Regex.Match(address, pattern);
                //    street_number = match.Groups[1].Value;
                //    street_name = match.Groups[2].Value;
                //    street_postfix = match.Groups[3].Value;
                //    city = match.Groups[4].Value;
                //    zip = match.Groups[5].Value;
                //}

                #endregion

                string dom = currentListing.mlsHtmlFields["daysOnMarket"].value;
                string closed_date = currentListing.mlsHtmlFields["closedDate"].value;
                string year_built = currentListing.mlsHtmlFields["yearBulit"].value;
                if (year_built == "NEW" || year_built == "0000" || year_built == "NC")
                {
                    try
                    {
                        year_built = currentListing.SalesDate.Year.ToString();
                    }
                    catch
                    {
                        year_built = DateTime.Now.Year.ToString();
                    }

                    currentListing.mlsHtmlFields["yearBulit"].value = year_built;
                }
                else if (year_built == "UNK" || year_built == "UNKN" || String.IsNullOrWhiteSpace(year_built) || !Regex.IsMatch(year_built, @"\d\d\d\d"))
                {
                    //blank it, for realist extract logic to work below
                    currentListing.mlsHtmlFields["yearBulit"].value = "";
                    year_built = "";
                }
                string room_count = currentListing.mlsHtmlFields["roomCount"].value;

                pattern = @"(\d+)\s*.\s*(\d+)";
                match = Regex.Match(currentListing.mlsHtmlFields["bathrooms"].value, pattern);
                string full_bath = match.Groups[1].Value;
                string half_bath = match.Groups[2].Value;

                string bedrooms = currentListing.mlsHtmlFields["bedrooms"].value;
                pattern = "\\s*(\\d+)";
                match = Regex.Match(bedrooms, pattern);


                if (!match.Success)
                {
                    //3 + 1 below grade
                    pattern = "\\s*(\\d+)\\s+\\d[^#]+";
                    match = Regex.Match(bedrooms, pattern);
                    bedrooms = match.Groups[1].Value;
                }

                string number_of_firplaces = currentListing.mlsHtmlFields["numFireplaces"].value;
                string basement = "None";
                basement = currentListing.mlsHtmlFields["basement"].value;
                string parking = currentListing.mlsHtmlFields["parking"].value;

                string mls_number_of_spaces = currentListing.mlsHtmlFields["numSpaces"].value;

                //pattern = "Lst. Mkt. Time:#NEXT#(\\d+)#NEXT#";
                // match = Regex.Match(table1, pattern);
                // string dom = match.Groups[1].Value;

                // pattern = "Closed:#NEXT#(\\d+.\\d+.\\d+)#NEXT#";
                // match = Regex.Match(table1, pattern);
                // string closed_date = match.Groups[1].Value;


                //pattern = "Year Built:#NEXT#(\\d+)#NEXT#";
                // match = Regex.Match(table1, pattern);
                // string year_built = match.Groups[1].Value;



                //pattern = "Rooms:#NEXT#(\\d+)#NEXT#";
                //match = Regex.Match(table1, pattern);
                //string room_count = match.Groups[1].Value;

                ////attached 0 space
                ////Bathrooms (Full/Half):#NEXT#1/1#NEXT#
                ////Bathrooms (Full/Half):#NEXT#1/1#NEXT#
                //pattern = "Bathrooms \\([Ff]ull\\/[Hh]alf\\):#NEXT#(\\d+)\\s*.\\s*(\\d+)#NEXT#";
                //match = Regex.Match(table1, pattern);
                //string full_bath = match.Groups[1].Value;
                //string half_bath = match.Groups[2].Value;

                //pattern = "Bedrooms:#NEXT#(\\d+)";
                //match = Regex.Match(table1, pattern);
                //string bedrooms = match.Groups[1].Value;

                //if (!match.Success)
                //{
                //    //3 + 1 below grade
                //    pattern = "Bedrooms:#NEXT#(\\d+)\\s+\\d[^#]+";
                //    match = Regex.Match(table1, pattern);
                //    bedrooms = match.Groups[1].Value;
                //}
                //pattern = "# Fireplaces:#NEXT#(\\d+)#NEXT#";
                //match = Regex.Match(table1, pattern);
                //string number_of_firplaces = match.Groups[1].Value;

                ////Full, English
                //pattern = "Basement:#NEXT#([^#]+)";
                //match = Regex.Match(table1, pattern);
                //string basement = "None";
                //basement = match.Groups[1].Value;

                //pattern = "Parking:#NEXT#(\\w+)#NEXT#";
                //match = Regex.Match(table1, pattern);
                //string parking = match.Groups[1].Value;


                //pattern = "# Spaces:#NEXT#\\w+:(\\d)";  //# Spaces:#NEXT#Gar:2 #NEXT#
                //match = Regex.Match(table1, pattern);
                //string mls_number_of_spaces = match.Groups[1].Value;



                #endregion


                //
                //Table4:  Details
                //

                // pattern = "#Type:(\\d*\\s*\\w+)#NEXT#";
                string type = currentListing.mlsHtmlFields["mlsType"].value;
                string mls_garage_type = currentListing.mlsHtmlFields["garageType"].value;

                pattern = "Gar:(\\d)";
                match = Regex.Match(mls_number_of_spaces, pattern);
                string mls_garage_spaces = match.Groups[1].Value;
                string basementDetails = currentListing.mlsHtmlFields["basementDetails"].value;

                //pattern = "#Type:([^#]+)";
                //match = Regex.Match(table4, pattern);
                //string type = match.Groups[1].Value;

                //pattern = "Garage Type:(([^#]+))";  //pattern = "Lot Acres:#NEXT#(([^#]+))";
                //match = Regex.Match(table4, pattern);
                //string mls_garage_type = match.Groups[1].Value;



                pattern = "Finished";  //pattern = "Lot Acres:#NEXT#(([^#]+))";
                match = Regex.Match(basementDetails, pattern);
                bool finished_basement = false;
                if (match.Success)
                {
                    finished_basement = true;
                }


                //
                //end table 4
                //


                pattern = "\\s*(\\d+)\\s*";
                match = Regex.Match(currentListing.mlsHtmlFields["pin"].value, pattern);
                //string pin = table2.Substring(table2.IndexOf("PIN:") + 4, 10);
                string pin = match.Groups[1].Value;

             

                string mls_gla = currentListing.mlsHtmlFields["mlsGla"].value;

                mls_gla = mls_gla.Replace("*", "");

                if (string.IsNullOrEmpty(mls_gla))
                {
                    mls_gla = "0";
                }


                //pattern = "Appx SF:#NEXT#(\\d+)#NEXT#";
                //match = Regex.Match(table2, pattern);
                ////string pin = table2.Substring(table2.IndexOf("PIN:") + 4, 10);
                //string mls_gla = match.Groups[1].Value;

                ////Attached MLS sheet in table 1
                //if (!match.Success)
                //{
                //    match = Regex.Match(table1, pattern);
                //    if (match.Success)
                //    {
                //        mls_gla = match.Groups[1].Value;
                //    }
                //    else
                //    {
                //        mls_gla = "0";
                //    }

                //}{

                string mls_lot_size = null;

                try
                {
                    mls_lot_size = currentListing.mlsHtmlFields["acerage"].value;
                }
                catch
                {

                }

                if (string.IsNullOrEmpty(mls_lot_size))
                {
                    mls_lot_size = "0";
                }


                //pattern = "Acreage:#NEXT#(\\d+.\\d+)";
                //match = Regex.Match(table2, pattern);
                ////string pin = table2.Substring(table2.IndexOf("PIN:") + 4, 10);
                //string mls_lot_size = "0";
                //mls_lot_size = match.Groups[1].Value;

                //if (!match.Success)
                //{
                //    pattern = "Acreage:#NEXT#(\\d+)";
                //    match = Regex.Match(table2, pattern);
                //    mls_lot_size = match.Groups[1].Value;

                //}

                //
                //listing agent info
                //


                string listing_Agent = currentListing.mlsHtmlFields["listAgent"].value;

                //pattern = "List Agent:#NEXT#([^\\(]+)";
                //match = Regex.Match(table5, pattern);
                //string listing_Agent = "";
                //listing_Agent = match.Groups[1].Value;

                if (listingAgentList.ContainsKey(listing_Agent))
                {
                    listingAgentList[listing_Agent]++;
                }
                else
                {
                    try
                    {
                        listingAgentList.Add(listing_Agent, 1);
                    }
                    catch
                    {
                        MessageBox.Show("Problem adding to listing agent list: " + listing_Agent);
                    }
                }

                string broker = currentListing.mlsHtmlFields["broker"].value;

                //pattern = "Broker:#NEXT#([^\\(]+)";
                //match = Regex.Match(table5, pattern);
                //string broker = "";
                //broker = match.Groups[1].Value;
                string phone = currentListing.mlsHtmlFields["phone"].value;

                //pattern = "Ph #:#NEXT#([^#]+)";
                //match = Regex.Match(table5, pattern);
                //string phone = "";
                //phone = match.Groups[1].Value;
                string additionalSalesInfo = currentListing.mlsHtmlFields["additionalSalesInfo"].value;

                ////pattern = "Additional Sales Information:#NEXT#([^#]+)|Addl. Sales Info.:#NEXT#([^#]+)";
                ////match = Regex.Match(table5, pattern);
                ////string additionalSalesInfo = "";
                ////additionalSalesInfo = match.Groups[1].Value;

                //if (additionalSalesInfo == "")
                //{
                //    additionalSalesInfo = match.Groups[2].Value;
                //}


                string mls_status = currentListing.mlsHtmlFields["status"].value;

                //pattern = @"Status:#NEXT#(\w+)#NEXT#";
                //match = Regex.Match(table1, pattern);
                //string mls_status = match.Groups[1].Value;

                if (mls_status != "CLSD" && additionalSalesInfo.Contains("Short Sale"))
                {
                    shortActive++;
                }

                if (mls_status != "CLSD" && additionalSalesInfo.Contains("REO"))
                {
                    reoActive++;
                }

                ////Google distance matric webservice
                ////http://maps.googleapis.com/maps/api/distancematrix/xml?origins=60050&destinations=60002&sensor=false&units=imperial
                ////

                //string googlestr = "http://maps.googleapis.com/maps/api/distancematrix/xml?origins=60050&destinations=" + zip + "&sensor=false&units=imperial";

                //// Create a request for the URL. 		
                //WebRequest request = WebRequest.Create(googlestr);

                //// Get the response.
                //HttpWebResponse response = (HttpWebResponse)request.GetResponse();


                //// Get the stream containing content returned by the server.
                //Stream dataStream = response.GetResponseStream();
                //// Open the stream using a StreamReader for easy access.
                //StreamReader reader = new StreamReader(dataStream);
                //// Read the content.
                //string responseFromServer = reader.ReadToEnd();
                //// Display the content.
                ////MessageBox.Show(response.ContentLength.ToString());
                ////MessageBox.Show(responseFromServer);
                ////Console.WriteLine(responseFromServer);
                //// Cleanup the streams and the response.
                //reader.Close();
                //dataStream.Close();
                //response.Close();







                #endregion
                form.StatusUpdate = "Done.";

                #region Price Change extraction
                //
                //Price Change extraction
                //
                //StringBuilder macro_get_price_changes = new StringBuilder();

                //macro_get_price_changes.AppendLine(@"FRAME NAME=workspace");
                //macro_get_price_changes.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:Additional<SP>Information");
                //macro_get_price_changes.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:Listing<SP>&<SP>Property<SP>History");
                //macro_get_price_changes.AppendLine(@"'New tab opened");
                //macro_get_price_changes.AppendLine(@"TAB T=2");
                //macro_get_price_changes.AppendLine(@"TAG POS=1 TYPE=TABLE FORM=NAME:NoFormName ATTR=CLASS:innergridview EXTRACT=TXT");
                //macro_get_price_changes.AppendLine(@"FRAME NAME=subheader");
                //macro_get_price_changes.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:header ATTR=ID:closeWinLink");
                //// macro_get_price_changes.AppendLine(@"TAB CLOSE");
                //string macroCode = macro_get_price_changes.ToString();

                //d.iimPlayCode(macroCode, 60);

                //string price_change_table = d.iimGetLastExtract(1);


                //int count = new Regex("-> PCHG").Matches(price_change_table).Count;



                ////string searchTerm = "PCHG";
                //////Convert the string into an array of words
                ////string[] source = price_change_table.Split(new char[] { '.', '?', '!', ' ', ';', ':', ',' }, StringSplitOptions.RemoveEmptyEntries);

                ////// Create and execute the query. It executes immediately 
                ////// because a singleton value is produced.
                ////// Use ToLowerInvariant to match "data" and "Data" 
                ////var matchQuery = from word in source
                ////                 where word.ToLowerInvariant() == searchTerm.ToLowerInvariant()
                ////                 select word;

                ////// Count the matches.
                ////int wordCount = matchQuery.Count();


                #endregion


                #region image_compare

                //StringBuilder save_pics_macro = new StringBuilder();

                ////string filepath = "C:\\Users\\Scott\\Documents\\My<SP>Dropbox\\BPOs\\" + search_address_textbox.Text;
                //string filepath = form.SubjectFilePath;
                //string filename = "comp.jpg";


                //    save_pics_macro.AppendLine(@"ONDOWNLOAD FOLDER=" + filepath.Replace(" ", "<SP>") + " FILE=comp.jpg");

                //save_pics_macro.AppendLine(@"FRAME NAME=workspace");
                //save_pics_macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:dc ATTR=SRC:*/PICS/*");
                //save_pics_macro.AppendLine(@"'New tab opened");
                ////save_pics_macro.AppendLine(@"WAIT SECONDS=5");
                //save_pics_macro.AppendLine(@"TAB T=2");
                //save_pics_macro.AppendLine(@"FRAME F=0");
                //save_pics_macro.AppendLine(@"TAG POS=2 TYPE=IMG FORM=NAME:dc ATTR=HREF:""*.JPEG"" CONTENT=EVENT:SAVEITEM");
                ////save_pics_macro.AppendLine(@"WAIT SECONDS=5");
                //save_pics_macro.AppendLine(@"TAB CLOSE");

                //string save_pics_macroCode = save_pics_macro.ToString();
                //status = d.iimPlayCode(save_pics_macroCode, 30);

                ////get the full path of the images
                //string image1Path = Path.Combine(filepath, "front.JPG");

                //string image2Path = Path.Combine(filepath, "comp.jpg");

                ////compare the two
                ////Console.Write("Comparing: " + bmp1 + " and " + bmp2 + ", with a threshold of " + threshold);
                //Bitmap firstBmp = (Bitmap)System.Drawing.Image.FromFile(image1Path);
                //Bitmap secondBmp = (Bitmap)System.Drawing.Image.FromFile(image2Path);
                //form.SubjectPic = firstBmp;
                //form.CompPic = new Bitmap(secondBmp);
                //form.PicDiffLabel = (firstBmp.PercentageDifference(secondBmp, 3) * 100).ToString();
                //picDiffList.Add(Convert.ToDecimal(firstBmp.PercentageDifference(secondBmp, 3) * 100));


                //using (secondBmp)
                //{
                //    //save the diffgram
                //    firstBmp.GetDifferenceImage(secondBmp, true).Save(image1Path + "_diff.png");
                //    //MessageBox.Show((string.Format("Difference: {0:0.0} %", firstBmp.PercentageDifference(secondBmp, 3) * 100)));

                //}
                #endregion

                form.StatusUpdate = "Reading Realist--> " + currentListing.MlsNumber;
                #region realist extraction

                string realist_rawHtml = "";
                string realist_gla = "0";
                string realist_subdivision = "";
                string censusTract = "";
                string realist_schoolDistrict = "";
                string realist_lotAcres = "0";
                string realist_address = "";

                if (form.UpdateRealist || !realist_bpohelper.RecordExists(pin))
                {
                    #region realist extraction update/add

                    #region open realist and extract data script

                    StringBuilder realist_extraction_macro = new StringBuilder();

                    //open realist
                    //realist_extraction_macro.AppendLine(@"SET !ERRORIGNORE YES");
                    realist_extraction_macro.AppendLine(@"SET !TIMEOUT_STEP 0");
                    realist_extraction_macro.AppendLine(@"FRAME NAME=workspace");
                    realist_extraction_macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:Realist<SP>Tax<SP>Report");
                    realist_extraction_macro.AppendLine(@" TAB T=2");
                    //  realist_extraction_macro.AppendLine(@"SET !TIMEOUT_STEP 0");

                    string realist_extraction_macro_code = realist_extraction_macro.ToString();
                    status = d.iimPlayCode(realist_extraction_macro_code, 60);
                    realist_extraction_macro.Clear();

                    if (status == Status.sOk)
                    {
                        //see what we have opened
                        s = d.iimPlayCode("SET !TIMEOUT_STEP 12 \r\n TAG POS=1 TYPE=TABLE ATTR=* EXTRACT=TXT");
                        realist_rawHtml = d.iimGetLastExtract();
                        //realist window is open but no properties found so skip the steps below
                        if (!realist_rawHtml.Contains("No properties were found."))
                        {
                            //gets all html
                            s = d.iimPlayCode("SET !TIMEOUT_STEP 12 \r\n TAG POS=1 TYPE=DIV ATTR=ID:htmlApplication EXTRACT=HTM");
                            realist_rawHtml = d.iimGetLastExtract();
                            //gets the address
                            s = d.iimPlayCode("SET !TIMEOUT_STEP 0 \r\n TAG POS=1 TYPE=DIV ATTR=ID:headerText EXTRACT=TXT");
                            realist_address = d.iimGetLastExtract();
                        }

                    }


                    if (realist_rawHtml.Contains("Parcel"))
                    {

                        #region goodRecord
                        //Owner Information
                        realist_extraction_macro.AppendLine(@"TAG POS=1 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");
                        realist_extraction_macro.AppendLine(@"SET !TIMEOUT_STEP 1");
                        //Location Information
                        realist_extraction_macro.AppendLine(@"TAG POS=2 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");

                        //Tax Information Table
                        realist_extraction_macro.AppendLine(@"TAG POS=3 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");

                        //Assessment & Tax
                        //Assessment
                        realist_extraction_macro.AppendLine(@"TAG POS=1 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");
                        //Tax
                        realist_extraction_macro.AppendLine(@"TAG POS=2 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");


                        //Characteristics
                        realist_extraction_macro.AppendLine(@"TAG POS=4 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");


                        //Estimated Value
                        realist_extraction_macro.AppendLine(@"TAG POS=5 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");

                        //Listing Information
                        //realist_extraction_macro.AppendLine(@"TAG POS=6 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");
                        //most recent mls listing
                        //realist_extraction_macro.AppendLine(@"TAG POS=3 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");

                        //Last Market Sale & Sales History
                        //first row
                        //realist_extraction_macro.AppendLine(@"TAG POS=4 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");
                        //second row
                        //realist_extraction_macro.AppendLine(@"TAG POS=5 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");

                        //Mortgage History
                        //first row
                        //realist_extraction_macro.AppendLine(@"TAG POS=6 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");
                        //second row
                        //  realist_extraction_macro.AppendLine(@"TAG POS=7 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");

                        //Foreclosure History
                        //   realist_extraction_macro.AppendLine(@"TAG POS=8 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");



                        //realist_extraction_macro.AppendLine(@"TAB CLOSE");
                        realist_extraction_macro_code = realist_extraction_macro.ToString();

                        status = d.iimPlayCode(realist_extraction_macro_code, 60);



                        string realist_owner_table = d.iimGetLastExtract(1);
                        string realist_location_table = d.iimGetLastExtract(2);
                        string realist_taxinfo_table = d.iimGetLastExtract(3);
                        string realist_assessment_table = d.iimGetLastExtract(4);
                        string realist_tax_table = d.iimGetLastExtract(5);
                        string realist_char_table = d.iimGetLastExtract(6);
                        string realist_estimatedvalue_table = d.iimGetLastExtract(7);
                        string realist_listinginfo1_table = d.iimGetLastExtract(8);
                        string realist_listinginfo2_table = d.iimGetLastExtract(9);
                        string realist_saleshisotry1_table = d.iimGetLastExtract(10);
                        string realist_saleshisotry2_table = d.iimGetLastExtract(11);
                        string realist_morthisotry1_table = d.iimGetLastExtract(12);
                        string realist_morthisotry2_table = d.iimGetLastExtract(13);
                        string realist_foreclosure_table = d.iimGetLastExtract(14);


                        //  s = d.iimPlayCode(@"TAG POS=1 TYPE=DIV ATTR=ID:headerText EXTRACT=TXT");
                        //  realist_address = d.iimGetLastExtract();

                        //  s = d.iimPlayCode(@"TAG POS=1 TYPE=DIV ATTR=ID:htmlApplication EXTRACT=HTM");
                        //  realist_rawHtml = d.iimGetLastExtract();

                        s = d.iimPlayCode(@"TAB CLOSE");

                        //          MessageBox.Show(realist_address + realist_rawHtml);



                        string allTables = realist_owner_table + "#NEXT#" + realist_location_table + "#NEXT#" + realist_taxinfo_table + "#NEXT#" + realist_char_table + "#NEXT#" + realist_estimatedvalue_table;



                        allTables.Replace("# #", "##");


                        //Township:#NEXT#Grafton Twp#NEXT#Census Tract:#NEXT#8711.04#NEXT##NEWLINE#Township Range Sect:#NEXT#43N-7E-27#NEXT#Carrier Route:#NEXT#R010#NEXT##NEWLINE#Subdivision:#NEXT#Pasquinellis Huntley Mdws#NEXT##NEXT# 

                        string[] sep = { "#NEWLINE#" };


                        string[] splitonnewline = allTables.Split(sep, StringSplitOptions.RemoveEmptyEntries);


                        string valuePattern = @"#NEXT#([^#]+)#NEXT#";
                        string namePattern = @"^([^#]+)#NEXT#|#NEXT#([^#]+)#NEXT#";
                        string fieldName;
                        string fieldValue;
                        string parid = Regex.Match(allTables, @"Parcel ID:#NEXT#([^#]+)").Groups[1].Value;

                        Dictionary<string, string> realistReportNameValuePairs = new Dictionary<string, string>();
                        #endregion

                        #region parse data and add/update fusion table
                        //if we found a parcel id, which is the key for the record
                        if (!String.IsNullOrEmpty(parid))
                        {

                            Stopwatch stopWatch = new Stopwatch();
                            stopWatch.Start();
                            string[] gFormatedLocation = new string[3];


                            gFormatedLocation = realist_address.Split(',');




                            if (!realist_bpohelper.RecordExists(parid))
                            {
                                //if realist is missing addess, split will only have 2 fields, not 3
                                try
                                {
                                    realist_bpohelper.AddRecord(parid, gFormatedLocation[0] + " " + gFormatedLocation[1] + " " + gFormatedLocation[2]);
                                }
                                catch
                                {
                                    realist_bpohelper.AddRecord(parid, address.Replace(",", ""));
                                }
                            }
                            else
                            {
                                try
                                {
                                    //add address to pending updates, incase it's missing or needs updating in fusion table
                                    realistReportNameValuePairs.Add("Location", gFormatedLocation[0] + " " + gFormatedLocation[1] + " " + gFormatedLocation[2]);
                                }
                                catch
                                {

                                }

                            }



                            stopWatch.Stop();
                            // Get the elapsed time as a TimeSpan value.
                            TimeSpan sts = stopWatch.Elapsed;

                            // Format and display the TimeSpan value. 
                            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                                sts.Hours, sts.Minutes, sts.Seconds,
                                sts.Milliseconds / 10);

                            form.SetStatusBar = elapsedTime;



                            //
                            //Check and add new colums
                            //
                            foreach (string valuePairLine in splitonnewline)
                            {


                                string modStr = Regex.Replace(valuePairLine, @"# ", "Number").Replace("Garage #", "Garage Number");


                                MatchCollection mc = Regex.Matches(modStr, namePattern);
                                foreach (Match m in mc)
                                {


                                    string v = m.Value.Replace("#NEXT#", "").Trim().Replace("(", @"\(").Replace(")", @"\)") + "#NEXT#([^#]+)";
                                    if (!v.Contains("NODATA"))
                                    {
                                        if (realist_bpohelper.ColumnExsists(m.Value.Replace("#NEXT#", "").Trim()))
                                        {


                                            realistReportNameValuePairs.Add(m.Value.Replace("#NEXT#", "").Trim(), Regex.Match(modStr, v).Groups[1].Value);
                                            //realist_bpohelper.UpdateRecord(parid, m.Value.Replace("#NEXT#", "").Trim(), Regex.Match(modStr, v).Groups[1].Value);
                                            //MessageBox.Show(m.Value + "exsists.");
                                        }
                                        else
                                        {
                                            if (m.Value != "NEXT")
                                            {
                                                realist_bpohelper.AddColumn(m.Value.Replace("#NEXT#", "").Trim());

                                                try
                                                {
                                                    realistReportNameValuePairs.Add(m.Value.Replace("#NEXT#", "").Trim(), Regex.Match(modStr, v).Groups[1].Value);
                                                }
                                                catch
                                                {
                                                    MessageBox.Show("problem adding realist nvp.");
                                                }


                                            }
                                        }
                                    }

                                }

                            }


                            //
                            //Update record
                            //

                            realistReportNameValuePairs.Concat(realist_bpohelper.Geocode(address));

                            realist_bpohelper.UpdateRecord(parid, realistReportNameValuePairs);

                            realist_bpohelper.m_rowid = "";


                            // MessageBox.Show(splitonnewline[0]);

                        }



                        //
                        //Location Information
                        //
                        pattern = "Subdivision:#NEXT#([^#]+)";
                        match = Regex.Match(realist_location_table, pattern);

                        realist_subdivision = match.Groups[1].Value;

                        //School District:
                        pattern = "School District:#NEXT#([^#]+)";
                        match = Regex.Match(realist_location_table, pattern);

                        realist_schoolDistrict = match.Groups[1].Value;

                        //Lot Acres:#NEXT#0.1257#NEXT#
                        pattern = "Lot Acres:#NEXT#([^#]+)";
                        match = Regex.Match(realist_char_table, pattern);

                        string realist_lot_size = "0";
                        realist_lot_size = match.Groups[1].Value;
                        realist_lotAcres = realist_lot_size;

                        //MessageBox.Show("MLS Lot: " + mls_lot_size + " " + "Realist Lot Size:" + realist_lot_size);

                        if (mls_lot_size == "0" | mls_lot_size == "")
                        {
                            mls_lot_size = realist_lot_size;
                        }

                        string realist_mls_gla;
                        string realist_tax_gla;

                        //Building Sq Ft:#NEXT#MLS: 1,381#NEXT#
                        pattern = "Building(| Above Grade) Sq Ft:#NEXT#([^#]+)";
                        match = Regex.Match(realist_char_table, pattern);

                        if (match.Success)
                        {
                            if (match.Captures[0].Value.Contains("MLS:"))
                            {
                                realist_mls_gla = Regex.Match(match.Groups[2].Value, @"MLS: (\S+)").Groups[1].Value;
                                realist_gla = realist_mls_gla;
                                if (match.Captures[0].Value.Contains("Tax:"))
                                {
                                    //both mls and tax
                                    realist_tax_gla = Regex.Match(match.Groups[2].Value, @"Tax: (\S+)").Groups[1].Value;
                                    realist_gla = realist_tax_gla;
                                }
                            }
                            else if (match.Captures[0].Value.Contains("Tax:"))
                            {
                                realist_tax_gla = Regex.Match(match.Groups[2].Value, @"Tax: (\S+)").Groups[1].Value;
                                realist_gla = realist_tax_gla;
                            }
                            else
                            {

                                    realist_gla = match.Groups[2].Value;
                                    if (mls_gla == "0" | mls_gla == "")
                                    {
                                        mls_gla = realist_gla;
                                    }
                                
                               
                            }
                        } else
                        {
                            realist_gla = "-1";
                        }


                    

                       
                        pattern = "Year Built:#NEXT#([^#]+)";
                        match = Regex.Match(realist_char_table, pattern);

                        string realist_yearbuilt = "";
                        realist_yearbuilt = Regex.Match(match.Groups[1].Value, @"(\d\d\d\d)").Value;

                        //MessageBox.Show("MLS GLA: " + mls_gla + " " + "Realist GLA:" + realist_gla);

                        //MessageBox.Show("MLS YB: " + year_built + " " + "Realist YB:" + realist_yearbuilt);
                        if (string.IsNullOrEmpty(year_built) || year_built.ToLower() == "unk" || year_built == "0")
                        {
                            if (!string.IsNullOrEmpty(realist_yearbuilt))
                            {
                                year_built = realist_yearbuilt;
                            }
                            else
                            {
                                year_built = "1950";
                            }
                        }



                        DateTime subject_age = new DateTime((Convert.ToInt32(year_built)), 1, 1);

                        TimeSpan ts = DateTime.Now - subject_age;

                        int age = ts.Days / 365;


                        pattern = @"Census Tract:#NEXT#([^#]+)";
                        match = Regex.Match(realist_char_table, pattern);

                        if (match.Success)
                        {
                            censusTract = match.Groups[1].Value;
                        }
                        else
                        {
                            match = Regex.Match(realist_location_table, pattern);
                            censusTract = match.Groups[1].Value;
                        }


                    }
                        #endregion //cccc

                    #endregion  //ccccdsddd

                    else
                    {
                        //unusable record OR
                        //No realist pop-up ie Wisconson address

                        s = d.iimPlayCode("SET !TIMEOUT_STEP 0 \r\n TAB T=2 \r\n TAB CLOSE");
                        if (string.IsNullOrEmpty(year_built) || year_built.ToLower() == "unk" || year_built == "0")
                        {
                            year_built = "1950";
                        }



                        DateTime subject_age = new DateTime((Convert.ToInt32(year_built)), 1, 1);

                        TimeSpan ts = DateTime.Now - subject_age;

                        int age = ts.Days / 365;

                    }




                }
                    #endregion

                else
                {
                    #region realist read from fusion table
                    try
                    {
                        realist_lotAcres = realist_bpohelper.curRec["Lot Acres:"];
                    }
                    catch
                    {

                    }

                    if (mls_lot_size == "0" | mls_lot_size == "")
                    {
                        mls_lot_size = realist_lotAcres;
                        double x;
                        Double.TryParse(realist_lotAcres, out x);
                        currentListing.Lotsize = x;
                    }
                    try
                    {
                        realist_gla = realist_bpohelper.curRec["Building Above Grade Sq Ft:"];
                        if (String.IsNullOrWhiteSpace(realist_gla))
                        {
                            realist_gla = realist_bpohelper.curRec["Building Sq Ft:"];
                             if (realist_gla.Contains("MLS:"))
                            {
                                 realist_gla = Regex.Match(realist_gla, @"MLS: (\S+)").Groups[1].Value;
                                
                                if (realist_gla.Contains("Tax:"))
                                {
                                    //both mls and tax
                                    realist_gla  = Regex.Match(realist_gla, @"Tax: (\S+)").Groups[1].Value;
                                 
                                }
                            }
                            else if (realist_gla.Contains("Tax:"))
                            {
                                realist_gla = Regex.Match(match.Groups[2].Value, @"Tax: (\S+)").Groups[1].Value;
                                
                            }
                        }
                    }
                    catch
                    {
                        try
                        {
                            realist_gla = realist_bpohelper.curRec["Building Sq Ft:"];
                        }

                        catch
                        {
                            realist_gla = "-1";
                        }
                      
                    }

                    
                    if (String.IsNullOrEmpty(realist_gla))
                    {
                        realist_gla = "-1";
                    }
                    else
                    {
                        if (mls_gla == "0" | mls_gla == "")
                        {
                            mls_gla = realist_gla;
                            int x;
                            Int32.TryParse("mls_gla", out x);
                            currentListing.GLA = x;
                        }
                    }
                    try
                    {
                        realist_subdivision = realist_bpohelper.curRec["Subdivision:"];
                        censusTract = realist_bpohelper.curRec["Census Tract:"];
                        realist_schoolDistrict = realist_bpohelper.curRec["School District:"];
                    }
                    catch
                    {
                        //optionsl fields, move on
                    }
                   
                    if (string.IsNullOrEmpty(year_built) || year_built.ToLower() == "unk" || year_built == "0")
                    {
                        if (realist_bpohelper.curRec["Year Built:"] != "")
                        {
                            if (realist_bpohelper.curRec["Year Built:"].Contains("Tax:"))
                            {
                                year_built = Regex.Match(realist_bpohelper.curRec["Year Built:"], @"Tax:\s*(\d\d\d\d)").Groups[1].Value;

                            }
                            else if (realist_bpohelper.curRec["Year Built:"].Contains("MLS:"))
                            {
                                year_built = Regex.Match(realist_bpohelper.curRec["Year Built:"], @"MLS:\s*(\d\d\d\d)").Groups[1].Value;
                            }
                            else
                            {
                                year_built = realist_bpohelper.curRec["Year Built:"];
                            }

                        }
                        else
                        {
                            year_built = "1950";
                        }
                    }
                    realist_bpohelper.Geocode(address);

                    #endregion
                }

                string propertyLatitude = "";
                string propertyLongitude = "";

                int xx = 0;
                Int32.TryParse(realist_gla.Replace(",", ""), out xx);
                currentListing.RealistGLA = xx;

                double dd = 0;
                Double.TryParse(realist_lotAcres, out dd);
                currentListing.RealistLotSize = dd;

                try
                {

                    propertyLatitude = realist_bpohelper.curRec["Latitude"];
                    propertyLongitude = realist_bpohelper.curRec["Longitude"];

                    GeoPoint destPoint;
                    try
                    {
                        destPoint.Latitude = Convert.ToDouble(propertyLatitude);
                        destPoint.Longitude = Convert.ToDouble(propertyLongitude);
                    }
                    catch
                    {
                        destPoint = GlobalVar.subjectPoint;
                    }

                    currentListing.GeoPointGd = propertyLatitude + ", " + propertyLongitude;
                    currentListing.proximityToSubject = Math.Round(Get_Distance(GlobalVar.subjectPoint, destPoint), 2);

                }
                catch
                {

                }

                #endregion
                form.StatusUpdate = "Done.";

                RealistReport currentRealistReport = new RealistReport();

                currentRealistReport.subdivision = realist_subdivision;

                currentProperty.myRealistReport = currentRealistReport;

                rpl.Add(currentProperty);

                //
                //Collect stats
                //
                #region reporting
                domList.Add(Convert.ToInt16(dom));

                if (Convert.ToDecimal(realist_gla) > 0 && Convert.ToDecimal(mls_gla) > Convert.ToDecimal(realist_gla))
                {
                   // mls_gla = realist_gla;
                }

                if (sold_price == " ")
                {
                    sold_price = current_list_price;
                }

                decimal[,] multiDimensionalArray1 = new decimal[2, 2];

                multiDimensionalArray1[0, 0] = Convert.ToDecimal(sold_price.Replace("$", "").Replace(",", ""));

                multiDimensionalArray1[0, 1] = Convert.ToDecimal(mls_gla.Replace(",", ""));

                ageList.Add(this.Age(year_built));
                int c = 0;

                if (realist_subdivision != "")
                {
                    try
                    {
                        c = subdivisionList[realist_subdivision];
                        subdivisionList[realist_subdivision] = c + 1;

                    }
                    catch
                    {
                        try
                        {
                            subdivisionList.Add(realist_subdivision, 1);
                        }
                        catch
                        {
                            MessageBox.Show("Failed adding to subdivision list: " + realist_subdivision);
                        }

                    }
                }

                if (type != "")
                {
                    try
                    {
                        c = mlsTypeList[type];
                        mlsTypeList[type] = c + 1;

                    }
                    catch
                    {
                        try
                        {

                            mlsTypeList.Add(type, 1);
                        }
                        catch
                        {
                            MessageBox.Show("Failed adding to mls type list: " + mlsTypeList);
                        }
                    }
                }


                // subdivisionList.Add(mls_subdivision + ", " + realist_subdivision, 1);


                list.Add(mls_status + ", " + sold_price.Replace("$", "").Replace(",", "") + ", " + mls_gla.Replace(",", ""));

                switch (mls_status)
                {
                    case "CLSD":
                        numListingsClosed++;
                        soldPriceList.Add(Convert.ToDecimal(sold_price.Replace("$", "").Replace(",", "")));

                        if (mls_gla != "0")
                        {
                            pricePerSfList.Add(Convert.ToDecimal(sold_price.Replace("$", "").Replace(",", "")) / Convert.ToDecimal(mls_gla));
                            try
                            {



                                //   int x = form.RawSFDatatable.Insert(Convert.ToDouble(propertyLongitude), Convert.ToDouble(propertyLatitude), Convert.ToDecimal(sold_price.Replace("$", "").Replace(",", "")), Convert.ToInt16(mls_gla.Replace(",", "")), zip, city, censusTract, mlsnum);
                            }
                            catch (Exception ex)
                            {
                                try
                                {
                                    //     int x = form.RawSFDatatable.Update(Convert.ToDouble(propertyLongitude), Convert.ToDouble(propertyLatitude), Convert.ToDecimal(sold_price.Replace("$", "").Replace(",", "")), Convert.ToInt16(mls_gla.Replace(",", "")), zip, city, censusTract, mlsnum);
                                }
                                catch
                                {
                                }

                                // MessageBox.Show(ex.Message);
                            }
                        }
                        break;
                    case "CTG":
                    case "ACTV":
                        numListingsActive++;
                        activePriceList.Add(Convert.ToDecimal(current_list_price.Replace("$", "").Replace(",", "")));
                        break;
                    case "PEND":
                        numListingsPending++;
                        activePriceList.Add(Convert.ToDecimal(current_list_price.Replace("$", "").Replace(",", "")));
                        break;

                    default:
                        //Console.WriteLine("Invalid selection. Please select 1, 2, or 3.");
                        break;
                }


                #endregion

                //
                //compare to subject
                //
                #region compare

                CompCriteria cc = new CompCriteria();
                List<string> compNotes = new List<string>();
                int ageDiff = Math.Abs(Convert.ToInt16(form.SubjectYearBuilt) - Convert.ToInt16(year_built));
                int glaDiff = Math.Abs(Convert.ToInt32(form.SubjectAboveGLA.Replace(",", "")) - Convert.ToInt32(mls_gla.Replace(",", ""))); // yes there are GLA over 32000 - ie 
                bool typeMatch = false;
                bool glaMatch = false;
                bool ageMatch = false;
                bool lotMatch = false;
                bool basementMatch = false;
                bool mredBasement = false;
                bool subjectBasement = false;



                // if anything other than none = it has a basement of some type
                if (currentListing.mlsHtmlFields["basement"].value != "None" || type.ToLower().Contains("split") || type.ToLower().Contains("raised"))
                {
                    mredBasement = true;
                }

                //same with subject
                if (form.SubjectBasementType.ToLower() != "none" || form.SubjectMlsType.ToLower().Contains("split") || form.SubjectMlsType.ToLower().Contains("raised"))
                {
                    subjectBasement = true;
                }


                if (currentListing.mlsHtmlFields["basement"].value.Contains(form.SubjectBasementType))
                {
                    basementMatch = true;
                }
                else if (subjectBasement && mredBasement)
                {
                    //for now, any type of basement will match
                    basementMatch = true;
                }


                if (form.SubjectMlsType.ToLower().Contains("duplex") && type.ToLower().Contains("duplex"))
                {
                    typeMatch = true;
                }

                //for townhome searches
                if (form.SubjectMlsType.ToLower().Contains("townhouse") && type.ToLower().Contains("townhouse"))
                {
                    typeMatch = true;
                }

                //for condo searches (multi-type selected on listing, as long as on is condo, we have to consider it)
                if (form.SubjectMlsType.ToLower().Contains("condo") && type.ToLower().Contains("condo"))
                {
                    typeMatch = true;
                }

                //for 1.5 story 

                if (form.SubjectMlsType.ToLower().Contains("1.5 story") && type.ToLower().Contains("2 stories"))
                {
                    typeMatch = true;
                }

                if (form.SubjectMlsType.ToLower().Contains("townhouse") && type.ToLower().Contains("condo"))
                {
                    typeMatch = true;
                }

                //for splits
                if (form.SubjectMlsType.ToLower().Contains("split") && type.ToLower().Contains("split"))
                {
                    typeMatch = true;
                }
                if (form.SubjectMlsType.ToLower().Contains("raised") && type.ToLower().Contains("split"))
                {
                    typeMatch = true;
                }

                

                if (type.Contains(form.SubjectMlsType.Replace(" ", " ")))
                {
                    typeMatch = true;
                }

                if (form.IgnoreType)
                {
                    typeMatch = true;
                }


                if (GlobalVar.ccc == GlobalVar.CompCompareSystem.NABPOP)
                {
                    if (glaDiff < Convert.ToInt16(form.SubjectAboveGLA.Replace(",", "")) * .2)
                        glaMatch = true;
                }
                else if (GlobalVar.ccc == GlobalVar.CompCompareSystem.USER)
                {
                    if (GlobalVar.lowerGLA <= Convert.ToDecimal(mls_gla.Replace(",", "")) && Convert.ToDecimal(mls_gla.Replace(",", "")) <= GlobalVar.upperGLA)
                    {
                        glaMatch = true;
                    }
                }


                if (form.IgnoreAge || cc.Age(form.SubjectYearBuilt, year_built))
                {
                    ageMatch = true;
                }

                if (form.SubjectAttached)
                {
                    lotMatch = true;
                }
                else if (cc.LotSize(form.SubjectLotSize, mls_lot_size))
                {
                    lotMatch = true;
                }


                if (typeMatch)
                {
                    numberSameTypeAsSubject++;
                }

                if (glaMatch)
                {
                    numberComparableGla++;
                }

                if (ageMatch)
                {
                    numberComparableAge++;
                }

                if (lotMatch)
                {
                    numberComparableLot++;
                }



                //form.SetStatusBar = "SubjectLotSize--> " + form.SubjectLotSize + "; CompLotSize--> " + mls_lot_size + "; Comparable = " + lotMatch.ToString();
                //string str = string.Format("MLS: {0} is {1} from subject.", currentListing.mlsHtmlFields["mlsNumber"].value, currentListing.proximityToSubject);
                //form.SetStatusBar = str;

             



                //
                //Format and display comp info to user
                //
                Color txtColor = new Color();


                form.AddInfoColor(currentListing.mlsHtmlFields["mlsNumber"].value, Color.Black);

                form.AddInfo = "  ";

                if (currentListing.proximityToSubject <= 1)
                {
                    form.AddInfoColor(currentListing.proximityToSubject.ToString(), Color.Green);
                }
                else
                {
                    form.AddInfoColor(currentListing.proximityToSubject.ToString(), Color.Yellow);
                }

                form.AddInfo = "\t";

                if (typeMatch)
                {
                    form.AddInfoColor(currentListing.mlsHtmlFields["mlsType"].value, Color.Green);
                }
                else
                {
                    form.AddInfoColor(currentListing.mlsHtmlFields["mlsType"].value, Color.Red);
                }


                //Type:2 Stories, Hillside = 19 chars, needs 2 tabs
                if (currentListing.mlsHtmlFields["mlsType"].value.Length >= 20)
                {
                    form.AddInfo = "\t";
                }
                else if (currentListing.mlsHtmlFields["mlsType"].value.Length >= 11)
                {
                    form.AddInfo = "\t\t";
                }
                else
                {
                    form.AddInfo = "\t\t\t";
                }



                if (glaMatch)
                {
                    //something we don't use the gla from mls listing or its blank
                    //form.AddInfoColor(currentListing.mlsHtmlFields["mlsGla"].value, Color.Green);
                    txtColor = Color.Green;
                }
                else
                {
                    txtColor = Color.Red;
                }

                form.AddInfoColor(mls_gla + "/" + realist_gla, txtColor);

                form.AddInfo = "\t\t";

                if (ageMatch)
                {
                    //form.AddInfoColor(currentListing.mlsHtmlFields["yearBulit"].value, Color.Green);
                    form.AddInfoColor(year_built, Color.Green);

                }
                else
                {
                    //form.AddInfoColor(currentListing.mlsHtmlFields["yearBulit"].value, Color.Red);
                    form.AddInfoColor(year_built, Color.Red);
                }

                form.AddInfo = "\t";

                if (form.SubjectDetached)
                {
                    if (lotMatch)
                    {
                        //something we don't use the lot size from mls listing or its blank
                        //form.AddInfoColor(currentListing.mlsHtmlFields["mlsGla"].value, Color.Green);
                        txtColor = Color.Green;
                    }
                    else
                    {
                        txtColor = Color.Red;
                    }

                    form.AddInfoColor(mls_lot_size, txtColor);
                    //if (lotMatch)
                    //{
                    //    form.AddInfoColor(currentListing.mlsHtmlFields["acerage"].value, Color.Green);
                    //}
                    //else
                    //{
                    //    form.AddInfoColor(currentListing.mlsHtmlFields["acerage"].value, Color.Red);
                    //}
                }

                form.AddInfo = "\t";

                if (basementMatch)
                {
                    txtColor = Color.Green;
                }
                else
                {
                    txtColor = Color.Red;
                }
                form.AddInfoColor(currentListing.mlsHtmlFields["basement"].value, txtColor);

                //Type:2 Stories, Hillside = 19 chars, needs 2 tabs
                if (currentListing.mlsHtmlFields["basement"].value.Length >= 20)
                {
                    form.AddInfo = "\t";
                }
                else if (currentListing.mlsHtmlFields["basement"].value.Length >= 11)
                {
                    form.AddInfo = "\t\t";
                }
                else
                {
                    form.AddInfo = "\t\t\t";
                }



                if (typeMatch && glaMatch && ageMatch && lotMatch && basementMatch)
                {
                    form.AddInfo = "Yes";
                }
                else if (typeMatch && glaMatch || glaMatch && ageMatch)
                {
                    form.AddInfo = "Maybe";
                }
                else
                {
                    form.AddInfo = "No";
                }
                form.AddInfo = "\t";
                 //
                //Scoring
                //
                #region Scoring 
                double compScore = 0;
                double maxCompScore = 0;

                maxCompScore++;
                if (currentListing.proximityToSubject < 0.5)
                {
                    compScore++;
                }
                else if (currentListing.proximityToSubject < 1)
                {
                    compScore = compScore + 0.5;
                }

                maxCompScore++;
                if (typeMatch)
                {
                    compScore++;
                }

                maxCompScore++;
                double percentChange = Math.Abs(currentListing.GLA - Convert.ToInt32(form.SubjectAboveGLA.Replace(",", ""))) / Convert.ToInt32(form.SubjectAboveGLA.Replace(",", ""));
                if (percentChange <= .1)
                {
                    compScore++;
                } else   if (percentChange <= .3)
                {
                    compScore = compScore +.5; 
                }

                      maxCompScore++;
                if (ageDiff < 10)
                {
                    compScore++;
                } else   if (ageDiff < 15)
                {
                    compScore = compScore +.5; 
                }

                 maxCompScore++;
                if (basementMatch)
                {
                    compScore++;
                }

                  maxCompScore++;
                 percentChange = Math.Abs(currentListing.Lotsize - Convert.ToDouble(form.SubjectLotSize.Replace(",", ""))) / Convert.ToDouble(form.SubjectLotSize.Replace(",", ""));
                if (percentChange <= .1)
                {
                    compScore++;
                } else   if (percentChange <= .3)
                {
                    compScore = compScore +.5; 
                }
                #endregion
                   form.AddInfo = compScore.ToString();
          //        form.AddInfo = "\r\n";


                form.AddInfo = "\r\n";

                GlobalVar.mainWindow.textBoxMaxCompScore.Text = maxCompScore.ToString();  






                // string rtfDisplay = string.Format("{0} {1} {2} {3} {4} {5} {6}", currentListing.mlsHtmlFields["mlsNumber"].value, currentListing.proximityToSubject.ToString()., );


                //MessageBox.Show("Type: " + type + " ->" + typeMatch.ToString() + "GLA: " + mls_gla + " ->" + glaMatch.ToString() + " YearBuilt: " + year_built + " ->" + ageDiff.ToString());
                //MessageBox.Show("YearBuilt: " + year_built + " ->" + ageDiff.ToString() + " " + ageMatch.ToString());
                bool isComp = false;
                #region ProcessComp
                if (typeMatch && glaMatch && ageMatch && lotMatch && basementMatch)
                {
                    isComp = true;
                    //List<int> compDomList = new List<int>();
                    //List<int> compActvDomList = new List<int>();
                    //List<int> compSoldDomList = new List<int>();
                    //List<int> compAgeList = new List<int>();
                    //List<int> compActvAgeList = new List<int>();
                    //List<int> compSoldAgeList = new List<int>();

                    //int numCompsActive = 0;
                    //int numCompsClosed = 0;
                    //int numCompsPending = 0;

                    //int compShortClosed = 0;
                    //int compReoClosed = 0;
                    //int compShortActive = 0;
                    //int compReoActive = 0;


                    compDomList.Add(Convert.ToInt16(dom));
                    compAgeList.Add(this.Age(year_built));

                    if (mls_status == "CLSD")
                    {
                        numCompsClosed++;
                        compSoldPriceList.Add(Convert.ToDecimal(sold_price.Replace("$", "").Replace(",", "")));
                        compSoldDomList.Add(Convert.ToInt16(dom));
                        compSoldAgeList.Add(this.Age(year_built));
                        if (sale_type == "REO")
                        {
                            compReoClosed++;
                        }
                        if (sale_type == "Short")
                        {
                            compShortClosed++;
                        }
                    }
                    else
                    {
                        if (additionalSalesInfo.Contains("Short Sale"))
                        {
                            compShortActive++;
                        }

                        if (additionalSalesInfo.Contains("REO"))
                        {
                            compReoActive++;
                        }
                        numCompsActive++;
                        compActvDomList.Add(Convert.ToInt16(dom));
                        compActvPriceList.Add(Convert.ToDecimal(current_list_price.Replace("$", "").Replace(",", "")));
                        compActvAgeList.Add(this.Age(year_built));
                    }


                    if (mls_status == "PEND")
                        numCompsPending++;





                    rplComps.Add(currentProperty);



                    //compActivePriceList


                    StringBuilder macro = new StringBuilder();
                    if (!searchCache)
                    {
                        macro.AppendLine(@"FRAME NAME=workspace");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:dc ATTR=ID:dcid* CONTENT=YES");

                        status = d.iimPlayCode(macro.ToString(), 30);

                    }

                    comps.Add(mlsnum);
                    if (realist_subdivision == form.SubjectSubdivision)
                    {
                        compNotes.Add("Same subdivision as subject");
                    }
                    if (realist_schoolDistrict == form.SubjectSchoolDistrict)
                    {
                        compNotes.Add("Same school district as subject");
                    }

                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(form.SubjectFilePath + "\\" + mlsnum + "notes.txt"))
                    {
                        foreach (object o in compNotes)
                        {
                            file.WriteLine(o.ToString());
                        }
                    }
                    //form = comps.Count.ToString();

                    form.NumberOfCompsFound = comps.Count.ToString();
                }
                #endregion

               


               
                #endregion

                #region savingToDatabases
                //save order info to DB
                //ataTable BPOtable = this.form.bpoTableAdapter1.GetData();

                // string tempstore = currentListing.rawData;

                // currentListing.rawData = "";


                ////sending mred records to datastore
                //string json = JsonConvert.SerializeObject(currentListing);
                //string url = "https://active-century-477.appspot.com/api/mred/v1/mredlistings/";

                //HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                //request.Method = "POST";

                //System.Text.UTF8Encoding encoding = new System.Text.UTF8Encoding();
                //Byte[] byteArray = encoding.GetBytes(json);

                //request.ContentLength = byteArray.Length;
                //request.ContentType = @"application/json";

                //using (Stream dataStream = request.GetRequestStream())
                //{
                //    dataStream.Write(byteArray, 0, byteArray.Length);
                //}
                //long length = 0;
                //try
                //{
                //    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                //    {
                //        length = response.ContentLength;
                //       // form.AddInfo = response.StatusDescription;
                //    }
                //}
                //catch (WebException ex)
                //{
                //    // Log exception and throw as for GET example above
                //}



                ////currentListing.rawData = tempstore;



                //// gcds.StoreMredListing(currentListing);

                #endregion

                if (searchCache)
                {
                    count++;
                    if (count >= GlobalVar.searchCacheMlsListings.Count)
                    {
                        stillRecords = false;
                    }
                }

                else
                {
                    StringBuilder move_through_comps_macro = new StringBuilder();
                    //GlobalVar.searchCacheMlsListings.Add(currentListing);
                    string tagPosition = "2";
                    if (pageNumber == 1)
                    {
                        tagPosition = "1";
                    }

                    if (pageNumber % 50 == 0)
                    {
                        move_through_comps_macro.AppendLine(@"FRAME NAME=subheader");
                        move_through_comps_macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:header ATTR=TXT:Advanced");
                        move_through_comps_macro.AppendLine(@"FRAME NAME=workspace");
                        move_through_comps_macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:dc ATTR=NAME:<SP><SP>Cancel<SP><SP>");
                    }
                    move_through_comps_macro.AppendLine(@"FRAME NAME=navpanel");
                    //move_through_comps_macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=TXT:Next");
                    move_through_comps_macro.AppendLine(@"TAG POS=" + tagPosition + " TYPE=DIV ATTR=onclick:show*");
                    move_through_comps_macro.AppendLine(@"FRAME NAME=workspace");
                    move_through_comps_macro.AppendLine(@"TAG POS=1 TYPE=DIV FORM=NAME:dc ATTR=CLASS:report EXTRACT=HTM");

               

                    status = d.iimPlayCode(move_through_comps_macro.ToString(), 60);
                    if (status != Status.sOk)
                    {
                        stillRecords = false;
                    }

                }
                pageNumber++;
               

                if (!form.PerserveCompSetData & compActvPriceList.Count > 0 & compSoldPriceList.Count > 0)
                {
                    setOfComps.maxSalePrice = compSoldPriceList.Max();
                    setOfComps.minSalePrice = compSoldPriceList.Min();
                    setOfComps.medianListPrice = Convert.ToDouble(compActvPriceList.Median());
                    setOfComps.medianAge = compAgeList.Median();
                    setOfComps.medianSalePrice = Convert.ToDouble(compSoldPriceList.Median());
                    setOfComps.numberSoldListings = compSoldPriceList.Count;
                    setOfComps.numberActiveListings = compActvPriceList.Count;
                    setOfComps.maxListPrice = compActvPriceList.Max();
                    setOfComps.minListPrice = compActvPriceList.Min();
                    setOfComps.numberREOSales = compReoClosed;
                    setOfComps.numberREOListings = compReoActive;
                    setOfComps.numberShortSales = compShortClosed;
                    setOfComps.numberOfShortSaleListings = compShortActive;
                    setOfComps.oldestHome = compAgeList.Max();
                    setOfComps.newestHome = compAgeList.Min();
                    setOfComps.avgDomActv = Convert.ToInt32(compActvDomList.Average());
                    setOfComps.avgDom = Convert.ToInt32(compDomList.Average());
                    form.SetOfComps = setOfComps;
                    

                }


           

            } //end while loop

            #endregion

            try
            {


                currentNeighborhood.minListPrice = activePriceList.Min();
                currentNeighborhood.medianListPrice = Convert.ToDouble(activePriceList.Median());
                currentNeighborhood.maxListPrice = activePriceList.Max();

                currentNeighborhood.minSalePrice = soldPriceList.Min();
                currentNeighborhood.medianSalePrice = Convert.ToDouble(soldPriceList.Median());
                currentNeighborhood.maxSalePrice = soldPriceList.Max();
                currentNeighborhood.medianAge = ageList.Median();



                currentNeighborhood.newestHome = ageList.Min();
                currentNeighborhood.numberOfCompListings = comps.Count;
                currentNeighborhood.numberSoldListings = soldPriceList.Count;
                currentNeighborhood.numberOfShortSaleListings = shortActive;
                currentNeighborhood.oldestHome = ageList.Max();
                currentNeighborhood.avgDom = Convert.ToInt32(domList.Average());
                currentNeighborhood.numberActiveListings = activePriceList.Count;
                currentNeighborhood.numberREOListings = reoActive;
                currentNeighborhood.numberREOSales = reoSales;
                currentNeighborhood.numberShortSales = shortSales;

                if (!form.PerserveNeighorhoodData)
                {
                    form.SubjectNeighborhood = currentNeighborhood;
                }


                //List<int> compDomList = new List<int>();
                //List<int> compActvDomList = new List<int>();
                //List<int> compSoldDomList = new List<int>();
                //List<int> compAgeList = new List<int>();
                //List<int> compActvAgeList = new List<int>();
                //List<int> compSoldAgeList = new List<int>();
                if (compActvPriceList.Count > 0 & compSoldPriceList.Count > 0)
                {
                    setOfComps.maxSalePrice = compSoldPriceList.Max();
                    setOfComps.minSalePrice = compSoldPriceList.Min();
                    setOfComps.medianListPrice = Convert.ToDouble(compActvPriceList.Median());
                    setOfComps.medianAge = compAgeList.Median();
                    setOfComps.medianSalePrice = Convert.ToDouble(compSoldPriceList.Median());
                    setOfComps.numberSoldListings = compSoldPriceList.Count;
                    setOfComps.numberActiveListings = compActvPriceList.Count;
                    setOfComps.maxListPrice = compActvPriceList.Max();
                    setOfComps.minListPrice = compActvPriceList.Min();
                    setOfComps.numberREOSales = compReoClosed;
                    setOfComps.numberREOListings = compReoActive;
                    setOfComps.numberShortSales = compShortClosed;
                    setOfComps.numberOfShortSaleListings = compShortActive;
                    setOfComps.oldestHome = compAgeList.Max();
                    setOfComps.newestHome = compAgeList.Min();
                    setOfComps.avgDomActv = Convert.ToInt32(compActvDomList.Average());
                    setOfComps.avgDom = Convert.ToInt32(compDomList.Average());
                }


                if (!form.PerserveCompSetData)
                {
                    form.SetOfComps = setOfComps;
                }




            }
            catch (Exception e)
            {
                //tbd
            }


            using (System.IO.StreamWriter file = new System.IO.StreamWriter(form.SubjectFilePath + "\\" + DateTime.Now.Ticks.ToString() + "_clsd_stats.txt"))
            {


                foreach (string s1 in list)
                {
                    if (s1.Contains("CLSD"))
                    {
                        file.WriteLine(s1);
                    }
                }
            }

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(form.SubjectFilePath + "\\" + DateTime.Now.Ticks.ToString() + "_active_stats.txt"))
            {


                foreach (string s1 in list)
                {
                    if (!s1.Contains("CLSD"))
                    {
                        file.WriteLine(s1);
                    }
                }
            }


            string searchName = "";

            if (string.IsNullOrEmpty(form.CurrentSearchName))
            {
                searchName = form.SubjectFilePath + "\\" + DateTime.Now.Ticks.ToString() + "_summary.txt";
            }
            else
            {
                searchName = form.SubjectFilePath + "\\" + form.CurrentSearchName + "_" + DateTime.Now.Ticks.ToString() + "_summary.txt";
            }

            try
            {


                using (System.IO.StreamWriter file = new System.IO.StreamWriter(form.SubjectFilePath + "\\searchComments.txt"))
                {
                    file.WriteLine("#closed: " + numListingsClosed.ToString() + " #active: " + numListingsActive.ToString() + " #pending: " + numListingsPending.ToString());
                    file.WriteLine("Average/Mean $/Above GLA, sale: " + Decimal.Round(pricePerSfList.Average(), 2).ToString() + " Median $/Above GLA, sale: " + Decimal.Round(pricePerSfList.Median(), 2).ToString());
                    file.WriteLine("Min Sale: {0}, Max Sale: {1}, Average Sale: {2}, Median Sale: {3}", soldPriceList.Min(), soldPriceList.Max(), Convert.ToInt32(soldPriceList.Average()), soldPriceList.Median());
                    file.WriteLine("Min List: {0}, Max List: {1}, Average List: {2}, Median List: {3}", activePriceList.Min(), activePriceList.Max(), Convert.ToInt32(activePriceList.Average()), activePriceList.Median());
                    file.WriteLine("Min dom: {0}, Max dom: {1}, Average dom: {2}, Median dom: {3}", domList.Min(), domList.Max(), Convert.ToInt32(domList.Average()), domList.Median());
                    file.WriteLine("Min Age: {0}, Max Age: {1}, Average Age: {2}, Median Age: {3}", ageList.Min(), ageList.Max(), Convert.ToInt32(ageList.Average()), ageList.Median());
                    file.WriteLine("REO Sold: {0}, REO Active: {1}, Short Sold: {2}, Short Active: {3}", reoSales, reoActive, shortSales, shortActive);
                }
            }
            catch
            {
                //no closed comps
            }

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(searchName))
            {


                int numSameSubdivision = 0;
                if (subdivisionList.ContainsKey("form.SubjectSubdivision"))
                {
                    numSameSubdivision = subdivisionList[form.SubjectSubdivision];
                }
                try
                {
                    file.WriteLine("#closed: " + numListingsClosed.ToString() + " #active: " + numListingsActive.ToString() + " #pending: " + numListingsPending.ToString());
                    file.WriteLine("Average/Mean $/Above GLA, sale: " + Decimal.Round(pricePerSfList.Average(), 2).ToString() + " Median $/Above GLA, sale: " + Decimal.Round(pricePerSfList.Median(), 2).ToString());
                    file.WriteLine("Min Sale: {0}, Max Sale: {1}, Average Sale: {2}, Median Sale: {3}", soldPriceList.Min(), soldPriceList.Max(), Convert.ToInt32(soldPriceList.Average()), soldPriceList.Median());
                    file.WriteLine("Min List: {0}, Max List: {1}, Average List: {2}, Median List: {3}", activePriceList.Min(), activePriceList.Max(), Convert.ToInt32(activePriceList.Average()), activePriceList.Median());
                    file.WriteLine("Min dom: {0}, Max dom: {1}, Average dom: {2}, Median dom: {3}", domList.Min(), domList.Max(), Convert.ToInt32(domList.Average()), domList.Median());
                    file.WriteLine("Min Age: {0}, Max Age: {1}, Average Age: {2}, Median Age: {3}", ageList.Min(), ageList.Max(), Convert.ToInt32(ageList.Average()), ageList.Median());
                    file.WriteLine("REO Sold: {0}, REO Active: {1}, Short Sold: {2}, Short Active: {3}", reoSales, reoActive, shortSales, shortActive);
                    file.WriteLine("# Same Type: {0}, # Same Subdivision {1}, # Comparable Age: {2},# Comparable aGLA: {3}, Comparable LotSize: {4}, ", numberSameTypeAsSubject.ToString(), numSameSubdivision.ToString(), numberComparableAge.ToString(), numberComparableGla.ToString(), numberComparableLot.ToString());


                }
                catch
                {

                }
                //picDiffList
                // file.WriteLine("Min {0}, Max {1}, Average {2}, Median {3}", picDiffList.Min(), picDiffList.Max(), picDiffList.Average(), picDiffList.Median());
                foreach (MLSListing m in listings)
                {
                    file.WriteLine("MLS: {0} is {1} from subject.", m.mlsHtmlFields["mlsNumber"].value, m.proximityToSubject);
                }


                foreach (object o in mlsTypeList)
                {
                    file.WriteLine(o.ToString());
                }

                foreach (object o in subdivisionList)
                {
                    file.WriteLine(o.ToString());
                }



                foreach (object o in listingAgentList)
                {
                    file.WriteLine(o.ToString());
                }
                foreach (MLSListing m in listings)
                {
                    file.WriteLine(m.mlsHtmlFields["remarks"].value);
                }


            }





            var queryRealProperties = from prop in rpl
                                      where !string.IsNullOrWhiteSpace(prop.Subdivision)
                                      select prop;



            System.Xml.Serialization.XmlSerializer writer =
      new System.Xml.Serialization.XmlSerializer(typeof(RealProperty));

            System.IO.StreamWriter xmlTestfile = new System.IO.StreamWriter(form.SubjectFilePath +
                @"\SerializationOverview.xml");

            foreach (RealProperty rp in queryRealProperties)
            {

                writer.Serialize(xmlTestfile, rp);
            }

            xmlTestfile.Close();

            var queryListings = from cust in listings
                                where cust.AttachedGarage()
                                select cust;

            var queryListingsRemarks = from cust in listings
                                       where cust.mlsHtmlFields["remarks"].value.ToLower().Contains("charming")
                                       select cust;

            var queryListingsProximity = from l in listings
                                         where l.proximityToSubject < .25
                                         select l;


            // MessageBox.Show(queryListingsRemarks.Count().ToString() + " listing used the word *charming*");


            System.Xml.Serialization.XmlSerializer writer2 =
            new System.Xml.Serialization.XmlSerializer(typeof(List<MLSListing>));


            string filename = @"\listingsFromSearchRunAt-" + DateTime.Now.ToString("dd-MM-yy-HH-mm-ss-ffff") + ".xml";
            System.IO.StreamWriter xmlTestfile2 = new System.IO.StreamWriter(form.SubjectFilePath + filename);
            System.IO.StreamWriter xmlTestfile3 = new System.IO.StreamWriter(form.SubjectFilePath + @"\listingFromLastSearch.xml");

            writer2.Serialize(xmlTestfile2, listings);
            writer2.Serialize(xmlTestfile3, listings);
            //foreach (MLSListing l in listings)
            //{
            //    writer2.Serialize(xmlTestfile2, l);
            //}

            xmlTestfile2.Close();

            GlobalVar.listingsFromLastSearch.Clear();
            GlobalVar.listingsFromLastSearch = listings;

            // MessageBox.Show(queryRealProperties.Count().ToString());

            //MessageBox.Show(GlobalVar.searchCacheMlsListings.Count.ToString());

            // MessageBox.Show("Comps found: " + comps.Count.ToString());
            if (compSoldPriceList.Count > 1)
            {
                form.SubjectMarketValue = compSoldPriceList.Median().ToString();
            }
        }
    }
}