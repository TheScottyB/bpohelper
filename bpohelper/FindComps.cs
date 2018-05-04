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
using System.Globalization;



namespace bpohelper
{
    public partial class Form1
    {
        private bool helper_comps_OnCompOne()
        {
            StringBuilder macro12 = new StringBuilder();
            macro12.AppendLine(@"FRAME NAME=navpanel");
            macro12.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=CLASS:showingtext EXTRACT=TXT");
            //   macro12.AppendLine(@"FRAME NAME=workspace");
            //  macro12.AppendLine(@"TAG POS=1 TYPE=DIV FORM=NAME:dc ATTR=CLASS:report EXTRACT=HTM");
            iMacros.Status s = iim.iimPlayCode(macro12.ToString());
            //  string htmlCode = iim.iimGetLastExtract();
            string header = iim.iimGetExtract(0);

            return Regex.IsMatch(header, @"showing\s*1\s*of\s*6");

        }

        private iMacros.Status helper_comps_GotoFirstComp()
        {
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");
            macro.AppendLine(@"FRAME NAME=subheader");
            macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:header ATTR=TXT:List<SP>View");
            macro.AppendLine(@"FRAME NAME=workspace");
            macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=HREF:javascript:clickedDCID('1');");
            string macroCode = macro.ToString();


            return iim.iimPlayCode(macroCode, 60);
        }

        async private void run_script_Click(object sender, EventArgs e)
        {
            MLSListing m;

            //todo: verify connection to mred 
            //verify on favs page
            //verify there are 6 comps
            //


            if (!helper_comps_OnCompOne())
            {
                var dr = MessageBox.Show("Do you want to start at comp 1?", "Error", MessageBoxButtons.YesNoCancel);
                if (dr == DialogResult.Yes)
                {
                    helper_comps_GotoFirstComp();
                }
                else if (dr == DialogResult.Cancel)
                {
                    return;
                }
            }

            #region statusAlertsInit
            Logger.log("Starting Comp Filling...");
            // Display the ProgressBar control.
            pBar2.Visible = true;
            // Set Minimum to 1 to represent the first file being copied.
            pBar2.Minimum = 1;
            // Set Maximum to the total number of files to copy.
            pBar2.Maximum = 6;
            // Set the initial value of the ProgressBar.
            pBar2.Value = 1;
            // Set the Step property to a value of 1 to represent each file being copied.
            pBar2.Step = 1;

            // Display the ProgressBar control.
            pBar1.Visible = true;
            // Set Minimum to 1 to represent the first file being copied.
            pBar1.Minimum = 1;
            // Set Maximum to the total number of files to copy.
            pBar1.Maximum = 24;
            // Set the initial value of the ProgressBar.
            pBar1.Value = 1;
            // Set the Step property to a value of 1 to represent each file being copied.
            pBar1.Step = 1;
            #endregion

            StringBuilder macro12 = new StringBuilder();
            //   macro12.AppendLine(@"FRAME NAME=navpanel");
            //    macro12.AppendLine(@"TAG POS=1 TYPE=TD ATTR=ID:navtd EXTRACT=TXT");
            macro12.AppendLine(@"FRAME NAME=workspace");
            macro12.AppendLine(@"TAG POS=1 TYPE=DIV FORM=NAME:dc ATTR=CLASS:report EXTRACT=HTM");
            iMacros.Status s = iim.iimPlayCode(macro12.ToString());
            string htmlCode = iim.iimGetLastExtract();
            // string header = iim.iimGetExtract(0);


            if (subjectAttachedRadioButton.Checked)
            {
                m = new AttachedListing(htmlCode);
            }
            else if (subjectDetachedradioButton.Checked)
            {
                m = new DetachedListing(htmlCode);
            }
            else
            {
                m = new MLSListing(htmlCode);
            }




            #region check for mred and start read
            iim2.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            string currentUrl = iim2.iimGetLastExtract();
            string input_comp_name = streetnameTextBox.Text;

            if (currentUrl.ToLower().Contains("usres"))
            {
                input_comp_name = "usres";
            }

            streetnameTextBox.Text = input_comp_name;
            streetnameTextBox.Update();


            StringBuilder move_through_comps_macro = new StringBuilder();
           // move_through_comps_macro.AppendLine(@"SET !TIMEOUT_STEP 0");
            move_through_comps_macro.AppendLine(@"FRAME NAME=navpanel");
            // move_through_comps_macro.AppendLine(@"TAG POS=2 TYPE=DIV ATTR=onclick:show*");
            move_through_comps_macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=TXT:Next");
            #endregion


            #region setup comp names

            int active_comps = 0;
            int closed_comps = 0;

            string[] comp_name_list = { "RECENT_SALE1", "RECENT_SALE2", "RECENT_SALE3" };
            string[] active_name_list = { "COMPARABLE1", "COMPARABLE2", "COMPARABLE3" };






            if (currentUrl.ToLower().Contains("pyramidplatform"))
            {

                comp_name_list[0] = "Sold1";
                comp_name_list[1] = "Sold2";
                comp_name_list[2] = "Sold3";
                active_name_list[0] = "List1";
                active_name_list[1] = "List2";
                active_name_list[2] = "List3";

            }

            if (currentUrl.ToLower().Contains("insidevaluation"))
            {

                comp_name_list[0] = "5";
                comp_name_list[1] = "6";
                comp_name_list[2] = "7";
                active_name_list[0] = "2";
                active_name_list[1] = "3";
                active_name_list[2] = "4";

            }

            if (currentUrl.ToLower().Contains("clearcapital"))
            {

                comp_name_list[0] = "6";
                comp_name_list[1] = "7";
                comp_name_list[2] = "8";
                active_name_list[0] = "2";
                active_name_list[1] = "3";
                active_name_list[2] = "4";

            }

            if (currentUrl.ToLower().Contains("solutionstar") || currentUrl.ToLower().Contains("amoservices") || currentUrl.ToLower().Contains("swbcls-forms"))
            {

                comp_name_list[0] = "1";
                comp_name_list[1] = "2";
                comp_name_list[2] = "3";
                active_name_list[0] = "1";
                active_name_list[1] = "2";
                active_name_list[2] = "3";

            }

            ////https://swbcls-forms.clearvalueconsulting.com/home/FormIndex?orderID=1288284&orderItemID=1&token=es76OHr59M%2fpYeYvEZXa33WFlZ4oP0UlZ3KvaC9qVHJETgVXIe9h%2ba96rOe43Wl7n%2balzbJMSpmV6gOB6Jd%2b6kqgiVgMLlThvywUG3fRMVfxoedtfuuASSvggcRCavW2
            //// if (currentUrl.ToLower().Contains("amoservices"))
            //if (currentUrl.ToLower().Contains("swbcls-forms"))
            //{

            //    comp_name_list[0] = "SALES_COMP1";
            //    comp_name_list[1] = "SALES_COMP2";
            //    comp_name_list[2] = "SALES_COMP3";
            //    active_name_list[0] = "LISTS_COMP1";
            //    active_name_list[1] = "LISTS_COMP2";
            //    active_name_list[2] = "LISTS_COMP3";

            //}

            if (streetnumTextBox.Text == "emort")
            {

                comp_name_list[0] = "_1";
                comp_name_list[1] = "_2";
                comp_name_list[2] = "_3";
                active_name_list[0] = "_1";
                active_name_list[1] = "_2";
                active_name_list[2] = "_3";

            }

            if (currentUrl.ToLower().Contains("bpofulfillment.com"))
            {

                comp_name_list[0] = "2";
                comp_name_list[1] = "3";
                comp_name_list[2] = "4";
                active_name_list[0] = "5";
                active_name_list[1] = "6";
                active_name_list[2] = "7";

            }

            if (streetnumTextBox.Text == "sls")
            {

                comp_name_list[0] = "1";
                comp_name_list[1] = "2";
                comp_name_list[2] = "3";
                active_name_list[0] = "1_2";
                active_name_list[1] = "2_2";
                active_name_list[2] = "3_2";

            }


            if (input_comp_name == "usres" || currentUrl.ToLower().Contains("res.net"))
            {


                if (currentUrl.ToLower().Contains(@"valuation.res.net/providerresponse"))
                {
                    //           .click('select#SaleComps_0__SaleType')
                    //.click('select#SaleComps_1__SaleType')
                    //.click('select#SaleComps_2__SaleType')
                    //.click('select#ListComps_0__State')
                    //.click('select#ListComps_1__State')
                    //.click('select#ListComps_2__State'

                    comp_name_list[0] = "0";
                    comp_name_list[1] = "1";
                    comp_name_list[2] = "2";
                    active_name_list[0] = "0";
                    active_name_list[1] = "1";
                    active_name_list[2] = "2";





                }
                else
                {
                    comp_name_list[0] = "S1";
                    comp_name_list[1] = "S2";
                    comp_name_list[2] = "S3";
                    active_name_list[0] = "L1";
                    active_name_list[1] = "L2";
                    active_name_list[2] = "L3";
                }



            }

            if (currentUrl.ToLower().Contains("valuationpartners"))
            {

                comp_name_list[0] = "COMP1";
                comp_name_list[1] = "COMP2";
                comp_name_list[2] = "COMP3";
                active_name_list[0] = "LISTCOMP1";
                active_name_list[1] = "LISTCOMP2";
                active_name_list[2] = "LISTCOMP3";

            }

            if (currentUrl.ToLower().Contains("equi-trax"))
            {

                comp_name_list[0] = "Sales1";
                comp_name_list[1] = "Sales2";
                comp_name_list[2] = "Sales3";
                active_name_list[0] = "List1";
                active_name_list[1] = "List2";
                active_name_list[2] = "List3";
            }

            if (currentUrl.ToLower().Contains("marktomarket"))
            {

                comp_name_list[0] = "Sale1";
                comp_name_list[1] = "Sale2";
                comp_name_list[2] = "Sale3";
                active_name_list[0] = "Listing1";
                active_name_list[1] = "Listing2";
                active_name_list[2] = "Listing3";
            }

            if (currentUrl.ToLower().Contains("ort.quandis"))
            {

                comp_name_list[0] = "CompOne";
                comp_name_list[1] = "CompTwo";
                comp_name_list[2] = "CompThree";
                active_name_list[0] = "ListingOne";
                active_name_list[1] = "ListingTwo";
                active_name_list[2] = "ListingThree";
            }


            if (currentUrl.ToLower().Contains("goodmandean"))
            {
                comp_name_list[0] = "CS1";
                comp_name_list[1] = "CS2";
                comp_name_list[2] = "CS3";
                active_name_list[0] = "CL1";
                active_name_list[1] = "CL2";
                active_name_list[2] = "CL3";

            }

            if (currentUrl.ToLower().Contains("exceleras"))
            {
                comp_name_list[0] = "0";
                comp_name_list[1] = "1";
                comp_name_list[2] = "2";
                active_name_list[0] = "0";
                active_name_list[1] = "1";
                active_name_list[2] = "2";

            }

            if (streetnumTextBox.Text == "nvs")
            {

                comp_name_list[0] = "CSale1";
                comp_name_list[1] = "CSale2";
                comp_name_list[2] = "CSale3";
                active_name_list[0] = "CList1";
                active_name_list[1] = "CList2";
                active_name_list[2] = "CList3";

            }

            #endregion

            string[] compPrices = { "", "", "" };

            Dictionary<string, MLSListing> stack = new Dictionary<string, MLSListing>();


            for (int i = 0; i < 6; i++)
            {
                #region loop through the 6 comps

                //
                //Read MRED listing sheet
                //
                #region read tables


                string sale_or_list_flag = "";



                if (m.Status.Contains("CLSD"))
                {
                    input_comp_name = comp_name_list[closed_comps];
                    closed_comps = closed_comps + 1;
                    sale_or_list_flag = "sale";

                }
                else
                {
                    input_comp_name = active_name_list[active_comps];
                    active_comps = active_comps + 1;
                    sale_or_list_flag = "list";
                }
                #endregion
                //
                //
                //

                //
                //Save pics
                //
                #region save comp pic
                //open pic tab
                StringBuilder openTab = new StringBuilder();
                openTab.AppendLine(@"FRAME NAME=workspace");
                openTab.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:dc ATTR=CLASS:*navImage*");
                //openTab.AppendLine(@"FRAME F=0");
                //openTab.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=HREF:javascript:openWindow* EXTRACT=TXT");

                string openTab_macroCode = openTab.ToString();
                status = iim.iimPlayCode(openTab_macroCode, 30);

                openTab.Clear();

                StringBuilder save_pics_macro = new StringBuilder();

                //string filepath = "C:\\Users\\Scott\\Documents\\My<SP>Dropbox\\BPOs\\" + search_address_textbox.Text;
                string filepath = search_address_textbox.Text;
                string filename = "comp.jpg";

                if (sale_or_list_flag == "sale")
                {
                    save_pics_macro.AppendLine(@"ONDOWNLOAD FOLDER=" + filepath.Replace(" ", "<SP>") + " FILE=S" + closed_comps.ToString() + ".jpg WAIT=YES");
                }
                else
                {
                    save_pics_macro.AppendLine(@"ONDOWNLOAD FOLDER=" + filepath.Replace(" ", "<SP>") + " FILE=L" + active_comps.ToString() + ".jpg WAIT=YES");
                }

                //line changed for new version of connectmls

                //save_pics_macro.AppendLine(@"'New tab opened");
                // save_pics_macro.AppendLine(@"WAIT SECONDS=2");
                save_pics_macro.AppendLine(@"TAB T=2");

                //save_pics_macro.AppendLine(@"TAG POS=2 TYPE=IMG FORM=NAME:dc ATTR=HREF:""*.JPEG"" CONTENT=EVENT:SAVEITEM");
                //save_pics_macro.AppendLine(@"WAIT SECONDS=5");
                //also changed for new connectmls
                //save_pics_macro.AppendLine(@"TAG POS=2 TYPE=A FORM=NAME:dc ATTR=TXT:Download<SP>This<SP>Photo CONTENT=EVENT:SAVETARGETAS");
                // save_pics_macro.AppendLine(@"TAG POS=2 TYPE=A FORM=NAME:dc ATTR=TXT:Download<SP>This<SP>Photo CONTENT=EVENT:SAVEITEM");

                // iim.iimPlayCode(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=HREF:javascript:openWindow('http%3A%2F%2Ftours.databasedads.com%2F2797441%2F1314-S-W* EXTRACT=TXT");




                //#save_pics_macro.AppendLine(@"FRAME F=0");

                if (iim.iimGetLastExtract().Contains("Virtual Tour"))
                {
                    // save_pics_macro.AppendLine(@"TAG POS=3 TYPE=IMG FORM=NAME:dc ATTR=HREF:""*.JPEG"" CONTENT=EVENT:SAVEITEM");
                    save_pics_macro.AppendLine(@"TAG POS=2 TYPE=A FORM=NAME:dc ATTR=TXT:Download<SP>This<SP>Photo");

                }
                else
                {
                    //  save_pics_macro.AppendLine(@"TAG POS=2 TYPE=IMG FORM=NAME:dc ATTR=HREF:""*.JPEG"" CONTENT=EVENT:SAVEITEM");
                    save_pics_macro.AppendLine(@"TAG POS=2 TYPE=A FORM=NAME:dc ATTR=TXT:Download<SP>This<SP>Photo");
                }


                save_pics_macro.AppendLine(@"WAIT SECONDS=2");
                save_pics_macro.AppendLine(@"TAB CLOSE");

                string save_pics_macroCode = save_pics_macro.ToString();
                status = iim.iimPlayCode(save_pics_macroCode, 30);
                #endregion
                //
                //
                //

                pBar1.PerformStep();
                # region read_mls_sheet

                string mlsnum = m.MlsNumber;
                string current_list_price = m.CurrentListPrice.ToString();
                string list_date = m.ListDateString;
                string orig_list_price = m.OriginalListPrice.ToString();
                string sold_price = m.SalePrice.ToString();
                string contract_date = m.ContractDate;
                string financingType = m.FinancingMlsString;

                if (sale_or_list_flag == "sale")
                {
                    compPrices[closed_comps - 1] = sold_price;
                }

                string sale_type = m.TransactionType;
                string address = m.StreetAddress;

                string street_number = "";
                string street_name = "";
                string street_postfix = "";
                string city = m.City;
                string zip = m.Zipcode;
                string full_street_address = m.StreetAddress;

                if (zip.Contains("-"))
                {
                    zip = zip.Remove(zip.IndexOf("-"));
                }

                //need to remove unit in attached listing addresses.
                if (full_street_address.Contains("Unit"))
                {
                    full_street_address = full_street_address.Remove(full_street_address.IndexOf("Unit"));
                }




                //2316 W Fairview Ln , McHenry, Illinois 60051
                string pattern = "^(\\d+)\\s+\\w\\s+(\\w+)\\s+(\\w+)";
                Match match = Regex.Match(full_street_address, pattern);

                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;

                }

                //2316 Fairview Ln , McHenry, Illinois 60051
                pattern = "^(\\d+)\\s+(\\w\\w+)\\s+(\\w+)";
                match = Regex.Match(full_street_address, pattern);
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
                match = Regex.Match(full_street_address, pattern);
                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;
                }


                //769 White Pine Cir , Lake In The Hills, Illinois 60156
                pattern = "^(\\d+)\\s+(\\w\\w+\\s+\\w+)\\s+(\\w+)[^\\w+]";
                match = Regex.Match(full_street_address, pattern);
                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;
                }

                // 21727 NORTH Hickory Hill Dr 
                pattern = "^(\\d+)\\s+NORTH\\s+(\\w\\w+\\s+\\w+)\\s+(\\w+)";
                match = Regex.Match(full_street_address, pattern);
                if (match.Success)
                {
                    street_number = match.Groups[1].Value;
                    street_name = match.Groups[2].Value;
                    street_postfix = match.Groups[3].Value;
                }
                string dom = m.DOM;
                string closed_date = m.SalesDate.ToShortDateString();
                string year_built = m.YearBuilt.ToString();
                string county = m.County;
                string room_count = m.TotalRoomCount.ToString();
                string full_bath = m.FullBathCount;
                string half_bath = m.HalfBathCount;
                string bedrooms = m.BedroomCount;
                string number_of_firplaces = m.NumberOfFireplaces;
                bool hasFireplace = false;

                if (number_of_firplaces != "")
                    hasFireplace = true;

                string basement = m.BasementType;
                bool fullBasement = false;
                bool partialBasement = false;

                if (basement.Contains("Full") || basement.Contains("English") || basement.Contains("Walkout"))
                    fullBasement = true;
                if (basement.Contains("Partial"))
                {
                    partialBasement = true;
                }

                string parking = m.MredParkingString;
                string mls_number_of_spaces = m.NumberGarageStalls();
                string mls_subdivision = m.Subdivision;

                if (string.IsNullOrWhiteSpace(mls_subdivision))
                {
                    mls_subdivision = "Unk/NA";
                }

                string financing = m.FinancingMlsString;
                string type = m.Type; //mls type field
                string mls_garage_type = m.GarageType();
                string mls_garage_spaces = m.NumberGarageStalls().ToString();

                bool finished_basement = false;
                if (m.FinishedBasement)
                {
                    finished_basement = true;
                }

                string exteriorDetails = m.ExteriorMlsString;
                string pin = m.ParcelNumber;
                string mls_gla = m.GLA.ToString();

                mls_gla = mls_gla.Replace(",", "");
                string mls_lot_size = "0";
                mls_lot_size = m.Lotsize.ToString();
                string listing_Agent = m.ListingAgentName;
                string broker = m.ListingBrokerageName;
                string phone = m.ListingAgentPhone();
                string additionalSalesInfo = m.AdditionalSalesInfo();

                #endregion

                pBar1.PerformStep();
                #region Price Change extraction
                //
                //Price Change extraction
                //
                StringBuilder macro_get_price_changes = new StringBuilder();

                macro_get_price_changes.AppendLine(@"FRAME NAME=workspace");
                macro_get_price_changes.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:Additional<SP>Information");
                macro_get_price_changes.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:Listing<SP>&<SP>Property<SP>History");
                macro_get_price_changes.AppendLine(@"'New tab opened");
                //macro_get_price_changes.AppendLine(@"TAB T=2");
                macro_get_price_changes.AppendLine(@"TAG POS=1 TYPE=TABLE FORM=NAME:NoFormName ATTR=CLASS:innergridview EXTRACT=TXT");
                macro_get_price_changes.AppendLine(@"FRAME NAME=subheader");
                // macro_get_price_changes.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:header ATTR=ID:closeWinLink");
                macro_get_price_changes.AppendLine(@"TAB CLOSE");
                string macroCode = macro_get_price_changes.ToString();

                status = iim.iimPlayCode(macroCode, 60);

                string price_change_table = iim.iimGetLastExtract(1);

                //STATUS: PCHG -> ACTV#NEXT#ACTV#NEXT#$449,900 #NEXT#05/21/2012#NEXT#
                DateTime lastPriceChangeDate = new DateTime();
                DateTime firstPriceChangeDate = new DateTime();

                match = Regex.Match(price_change_table, @"STATUS: PCHG -> ACTV#NEXT#ACTV#NEXT#\$[\d,]+ #NEXT#(\d+.\d+.\d+)#NEXT#");
                MatchCollection mc = Regex.Matches(price_change_table, @"STATUS: PCHG -> ACTV#NEXT#ACTV#NEXT#\$[\d,]+ #NEXT#(\d+.\d+.\d+)#NEXT#");


                string domAtCurrentListPrice;
                string domAtOriginalListPrice;
                try
                {
                    lastPriceChangeDate = DateTime.Parse(match.Groups[1].Value);
                    firstPriceChangeDate = DateTime.Parse(mc[mc.Count - 1].Groups[1].Value);

                    domAtCurrentListPrice = (DateTime.Now - lastPriceChangeDate).Days.ToString();
                    domAtOriginalListPrice = (firstPriceChangeDate - DateTime.Parse(list_date)).Days.ToString();
                }
                catch
                {
                    lastPriceChangeDate = DateTime.Parse(list_date);
                    firstPriceChangeDate = lastPriceChangeDate;
                    domAtCurrentListPrice = dom;
                    domAtOriginalListPrice = dom;
                }




                // domAtCurrentListPrice = (DateTime.Now - lastPriceChangeDate).Days.ToString();
                //  domAtOriginalListPrice = (DateTime.Parse(list_date) - firstPriceChangeDate).Days.ToString();

                // MessageBox.Show(domAtCurrentListPrice + ", " + domAtOriginalListPrice);


                int count = new Regex("-> PCHG").Matches(price_change_table).Count;



                //string searchTerm = "PCHG";
                ////Convert the string into an array of words
                //string[] source = price_change_table.Split(new char[] { '.', '?', '!', ' ', ';', ':', ',' }, StringSplitOptions.RemoveEmptyEntries);

                //// Create and execute the query. It executes immediately 
                //// because a singleton value is produced.
                //// Use ToLowerInvariant to match "data" and "Data" 
                //var matchQuery = from word in source
                //                 where word.ToLowerInvariant() == searchTerm.ToLowerInvariant()
                //                 select word;

                //// Count the matches.
                //int wordCount = matchQuery.Count();


                #endregion
                pBar1.PerformStep();

                #region realist extraction
                GoogleFusionTable realist_bpohelper = new GoogleFusionTable("1UKrOVmhPWrgLP5d5bDCsiW9whMIK8aLxKhcyOaI");
                await realist_bpohelper.helper_OAuthFusion();

                string realist_rawHtml = "";
                string realist_gla = "0";
                string realist_subdivision = "";
                string censusTract = "";
                string realist_schoolDistrict = "";
                string realist_lotAcres = "0";
                string realist_address = "";


                //StringBuilder realist_extraction_macro = new StringBuilder();
                //realist_extraction_macro.AppendLine(@"SET !TIMEOUT_STEP 10");
                //realist_extraction_macro.AppendLine(@"FRAME NAME=workspace");
                //realist_extraction_macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:Realist<SP>Tax<SP>Report");
                //realist_extraction_macro.AppendLine(@"'New tab opened");
                ////realist_extraction_macro.AppendLine(@"TAB T=2");
                //realist_extraction_macro.AppendLine(@"FRAME F=0");
                //realist_extraction_macro.AppendLine(@"SET !TIMEOUT_STEP 10");
                //realist_extraction_macro.AppendLine(@"TAG POS=4 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");
                //realist_extraction_macro.AppendLine(@"TAG POS=2 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");
                //realist_extraction_macro.AppendLine(@"TAB CLOSE");
                //string realist_extraction_macro_code = realist_extraction_macro.ToString();

                //status = iim.iimPlayCode(realist_extraction_macro_code, 60);

                //string realist_char_table = iim.iimGetLastExtract(1);
                //string realist_location_table = iim.iimGetLastExtract(2);

                ////
                ////Location Information
                ////
                //pattern = "Subdivision:#NEXT#([^#]+)";
                //match = Regex.Match(realist_location_table, pattern);
                // realist_subdivision = "";
                //realist_subdivision = match.Groups[1].Value;

                ////
                ////Characteristics
                ////
                ////Lot Acres:#NEXT#0.1257#NEXT#
                //pattern = "Lot Acres:#NEXT#([^#]+)";
                //match = Regex.Match(realist_char_table, pattern);

                //string realist_lot_size = "0";
                //realist_lot_size = match.Groups[1].Value;

                ////MessageBox.Show("MLS Lot: " + mls_lot_size + " " + "Realist Lot Size:" + realist_lot_size);

                //if (mls_lot_size == "0" | mls_lot_size == "")
                //{
                //    mls_lot_size = realist_lot_size;
                //}

                //pattern = "Building Above Grade Sq Ft:#NEXT#([^#]+)";
                //match = Regex.Match(realist_char_table, pattern);

                // realist_gla = "0";
                //realist_gla = match.Groups[1].Value;

                ////MessageBox.Show("MLS GLA: " + mls_gla + " " + "Realist GLA:" + realist_gla);

                //if (mls_gla == "0" | mls_gla == "")
                //{
                //    mls_gla = realist_gla.Replace(",","");
                //}

                //pattern = "Year Built:#NEXT#([^#]+)";
                //match = Regex.Match(realist_char_table, pattern);

                //string realist_yearbuilt = "";
                //realist_yearbuilt = match.Groups[1].Value;

                ////MessageBox.Show("MLS GLA: " + mls_gla + " " + "Realist GLA:" + realist_gla);

                ////MessageBox.Show("MLS YB: " + year_built + " " + "Realist YB:" + realist_yearbuilt);




                DateTime subject_age = new DateTime((Convert.ToInt32(SubjectYearBuilt)), 1, 1);

                TimeSpan ts = DateTime.Now - subject_age;

                int age = ts.Days / 365;


                if (!realist_bpohelper.RecordExists(pin))
                {
                    #region realist extraction update/add

                    #region open realist and extract data script

                    StringBuilder realist_extraction_macro = new StringBuilder();

                    //open realist
                    //realist_extraction_macro.AppendLine(@"SET !ERRORIGNORE YES");
                    realist_extraction_macro.AppendLine(@"SET !TIMEOUT_STEP 0");
                    realist_extraction_macro.AppendLine(@"FRAME NAME=workspace");
                    realist_extraction_macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:dc ATTR=TXT:Realist<SP>Tax<SP>Report");

                    string realist_extraction_macro_code = realist_extraction_macro.ToString();
                    status = iim.iimPlayCode(realist_extraction_macro_code, 60);
                    realist_extraction_macro.Clear();

                    if (status == Status.sOk)
                    {
                        //see what we have opened
                        s = iim.iimPlayCode("SET !TIMEOUT_STEP 12 \r\n TAG POS=1 TYPE=TABLE ATTR=* EXTRACT=TXT");
                        realist_rawHtml = iim.iimGetLastExtract();
                        //realist window is open but no properties found so skip the steps below
                        if (!realist_rawHtml.Contains("No properties were found."))
                        {
                            //gets all html
                            s = iim.iimPlayCode("SET !TIMEOUT_STEP 12 \r\n TAG POS=1 TYPE=DIV ATTR=ID:htmlApplication EXTRACT=HTM");
                            realist_rawHtml = iim.iimGetLastExtract();
                            //gets the address
                            s = iim.iimPlayCode("SET !TIMEOUT_STEP 0 \r\n TAG POS=1 TYPE=DIV ATTR=ID:headerText EXTRACT=TXT");
                            realist_address = iim.iimGetLastExtract();
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

                        status = iim.iimPlayCode(realist_extraction_macro_code, 60);



                        string realist_owner_table = iim.iimGetLastExtract(1);
                        string realist_location_table = iim.iimGetLastExtract(2);
                        string realist_taxinfo_table = iim.iimGetLastExtract(3);
                        string realist_assessment_table = iim.iimGetLastExtract(4);
                        string realist_tax_table = iim.iimGetLastExtract(5);
                        string realist_char_table = iim.iimGetLastExtract(6);
                        string realist_estimatedvalue_table = iim.iimGetLastExtract(7);
                        string realist_listinginfo1_table = iim.iimGetLastExtract(8);
                        string realist_listinginfo2_table = iim.iimGetLastExtract(9);
                        string realist_saleshisotry1_table = iim.iimGetLastExtract(10);
                        string realist_saleshisotry2_table = iim.iimGetLastExtract(11);
                        string realist_morthisotry1_table = iim.iimGetLastExtract(12);
                        string realist_morthisotry2_table = iim.iimGetLastExtract(13);
                        string realist_foreclosure_table = iim.iimGetLastExtract(14);


                        //  s = iim.iimPlayCode(@"TAG POS=1 TYPE=DIV ATTR=ID:headerText EXTRACT=TXT");
                        //  realist_address = iim.iimGetLastExtract();

                        //  s = iim.iimPlayCode(@"TAG POS=1 TYPE=DIV ATTR=ID:htmlApplication EXTRACT=HTM");
                        //  realist_rawHtml = iim.iimGetLastExtract();

                        s = iim.iimPlayCode(@"TAB CLOSE");

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
                                //add address to pending updates, incase it's missing or needs updating in fusion table
                                realistReportNameValuePairs.Add("Location", gFormatedLocation[0] + " " + gFormatedLocation[1] + " " + gFormatedLocation[2]);

                            }





                            //
                            //Check and add new colums
                            //
                            foreach (string valuePairLine in splitonnewline)
                            {


                                string modStr = Regex.Replace(valuePairLine, @"# ", "Number").Replace("Garage #", "Garage Number");


                                MatchCollection mc1 = Regex.Matches(modStr, namePattern);
                                foreach (Match m1 in mc1)
                                {


                                    string v = m1.Value.Replace("#NEXT#", "").Trim().Replace("(", @"\(").Replace(")", @"\)") + "#NEXT#([^#]+)";
                                    if (!v.Contains("NODATA"))
                                    {
                                        if (realist_bpohelper.ColumnExsists(m1.Value.Replace("#NEXT#", "").Trim()))
                                        {


                                            realistReportNameValuePairs.Add(m1.Value.Replace("#NEXT#", "").Trim(), Regex.Match(modStr, v).Groups[1].Value);
                                            //realist_bpohelper.UpdateRecord(parid, m1.Value.Replace("#NEXT#", "").Trim(), Regex.Match(modStr, v).Groups[1].Value);
                                            //MessageBox.Show(m1.Value + "exsists.");
                                        }
                                        else
                                        {
                                            if (m1.Value != "NEXT")
                                            {
                                                realist_bpohelper.AddColumn(m1.Value.Replace("#NEXT#", "").Trim());

                                                try
                                                {
                                                    realistReportNameValuePairs.Add(m1.Value.Replace("#NEXT#", "").Trim(), Regex.Match(modStr, v).Groups[1].Value);
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
                        double x = -1;

                        Double.TryParse(realist_lot_size, out x);
                        m.RealistLotSize = x;
                        //MessageBox.Show("MLS Lot: " + mls_lot_size + " " + "Realist Lot Size:" + realist_lot_size);

                        if (mls_lot_size == "0" | mls_lot_size == "")
                        {
                            mls_lot_size = realist_lot_size;

                        }

                        pattern = "Building Above Grade Sq Ft:#NEXT#([^#]+)";
                        match = Regex.Match(realist_char_table, pattern);


                        realist_gla = match.Groups[1].Value;

                        int xi = -1;

                        Int32.TryParse(realist_gla, out xi);
                        m.RealistGLA = xi;


                        //MessageBox.Show("MLS GLA: " + mls_gla + " " + "Realist GLA:" + realist_gla);
                        if (match.Success)
                        {
                            if (mls_gla == "0" | mls_gla == "")
                            {
                                mls_gla = realist_gla;
                            }
                        }
                        else
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



                        subject_age = new DateTime((Convert.ToInt32(year_built)), 1, 1);

                        ts = DateTime.Now - subject_age;

                        age = ts.Days / 365;


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

                        s = iim.iimPlayCode("SET !TIMEOUT_STEP 0 \r\n TAB T=2 \r\n TAB CLOSE");
                        if (string.IsNullOrEmpty(year_built) || year_built.ToLower() == "unk" || year_built == "0")
                        {
                            year_built = "1950";
                        }



                        subject_age = new DateTime((Convert.ToInt32(year_built)), 1, 1);

                        ts = DateTime.Now - subject_age;

                        age = ts.Days / 365;

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

                    if (mls_lot_size == "0" | mls_lot_size == "" | mls_lot_size == "-1")
                    {
                        mls_lot_size = realist_lotAcres;
                        double x;
                        Double.TryParse(realist_lotAcres, out x);
                        m.Lotsize = x;
                    }

                    realist_gla = realist_bpohelper.curRec["Building Above Grade Sq Ft:"];
                    if (String.IsNullOrEmpty(realist_gla))
                    {
                        realist_gla = "-1";
                    }
                    else
                    {
                        if (mls_gla == "0" | mls_gla == "" | mls_gla == "-1")
                        {
                            mls_gla = realist_gla;
                            int x;
                            Int32.TryParse(mls_gla, out x);
                            m.GLA = x;
                        }
                    }
                    realist_subdivision = realist_bpohelper.curRec["Subdivision:"];
                    censusTract = realist_bpohelper.curRec["Census Tract:"];
                    realist_schoolDistrict = realist_bpohelper.curRec["School District:"];
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
                m.RealistGLA = xx;

                double dd = 0;
                Double.TryParse(realist_lotAcres, out dd);
                m.RealistLotSize = dd;

                if (string.IsNullOrEmpty(m.YearBuiltString) || m.YearBuiltString.ToLower() == "unk" || m.YearBuiltString == "0")
                {
                    if (!string.IsNullOrEmpty(year_built))
                    {
                        int yxx = 0;
                      Int32.TryParse(year_built, out yxx);
                      m.YearBuilt = yxx;
                    }
                    else
                    {
                        year_built = "1950";
                    }
                }

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

                    MlsDrivers x = new MlsDrivers();

                    m.proximityToSubject = Math.Round(x.Get_Distance(GlobalVar.subjectPoint, destPoint), 2);

                }
                catch
                {

                }

                #endregion
                pBar1.PerformStep();


                // Perform the increment on the ProgressBar.

                #region sandcastle
                if (currentUrl.ToLower().Contains("sandcastlefs"))
                {
                    #region code

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    Sandcastle bpoform = new Sandcastle(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion


                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                }

                #endregion

                #region pyramid
                if (currentUrl.ToLower().Contains("pyramid"))
                {
                    #region code

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    Pyramid bpoform = new Pyramid(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion


                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                }

                #endregion

                //clearvalue platform

                #region SWBC
                if (currentUrl.ToLower().Contains("swbc"))
                {
                    #region code

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    SWBC bpoform = new SWBC(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion


                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                }

                #endregion


                #region AMO-swbc
                if (currentUrl.ToLower().Contains("amoservices"))
                {
                    #region code

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    AMO bpoform = new AMO(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion


                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                }

                #endregion

                #region ClearCap
                if (currentUrl.ToLower().Contains("clearcapital"))
                {
                    #region code

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    ClearCap bpoform = new ClearCap(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion


                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                }

                #endregion

                //
                //First American - First Pass
                //
                #region FirstAmerican
                if (currentUrl.ToLower().Contains("firstam"))
                {
                    #region code

                    //s = iim.iimPlayCode(macro12.ToString());
                    //htmlCode = iim.iimGetLastExtract();
                    //if (subjectAttachedRadioButton.Checked)
                    //{
                    //    m = new AttachedListing(htmlCode);
                    //}
                    //else if (subjectDetachedradioButton.Checked)
                    //{
                    //    m = new DetachedListing(htmlCode);
                    //}
                    //else
                    //{
                    //    m = new MLSListing(htmlCode);
                    //}

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    FirstAmerican bpoform = new FirstAmerican(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion

                    #region basementlogic




                    #endregion

                    #region garagelogic

                    #endregion

                    #region fireplacelogic

                    #endregion

                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                }

                #endregion


                //
                //Trinity - First Pass
                //
                #region Trinity
                if (currentUrl.ToLower().Contains("trinityonline"))
                {
                    #region code

                    //s = iim.iimPlayCode(macro12.ToString());
                    //htmlCode = iim.iimGetLastExtract();
                    //if (subjectAttachedRadioButton.Checked)
                    //{
                    //    m = new AttachedListing(htmlCode);
                    //}
                    //else if (subjectDetachedradioButton.Checked)
                    //{
                    //    m = new DetachedListing(htmlCode);
                    //}
                    //else
                    //{
                    //    m = new MLSListing(htmlCode);
                    //}

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    Trinity bpoform = new Trinity(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion

                    #region basementlogic




                    #endregion

                    #region garagelogic

                    #endregion

                    #region fireplacelogic

                    #endregion

                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                }

                #endregion

                //
                //inside valuation - First Pass
                //
                #region insideval
                if (currentUrl.ToLower().Contains("insidevaluation"))
                {
                    #region code

                    //s = iim.iimPlayCode(macro12.ToString());
                    //htmlCode = iim.iimGetLastExtract();
                    //if (subjectAttachedRadioButton.Checked)
                    //{
                    //    m = new AttachedListing(htmlCode);
                    //}
                    //else if (subjectDetachedradioButton.Checked)
                    //{
                    //    m = new DetachedListing(htmlCode);
                    //}
                    //else
                    //{
                    //    m = new MLSListing(htmlCode);
                    //}

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    InsideValuation bpoform = new InsideValuation(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion

                    #region basementlogic




                    #endregion

                    #region garagelogic

                    #endregion

                    #region fireplacelogic

                    #endregion


                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                }

                #endregion

                //
                //REOcBPO - First Pass
                //
                #region REOcBPO
                if (currentUrl.ToLower().Contains("reo-central"))
                {
                    #region code

                    s = iim.iimPlayCode(macro12.ToString());
                    htmlCode = iim.iimGetLastExtract();
                    if (subjectAttachedRadioButton.Checked)
                    {
                        m = new AttachedListing(htmlCode);
                    }
                    else if (subjectDetachedradioButton.Checked)
                    {
                        m = new DetachedListing(htmlCode);
                    }
                    else
                    {
                        m = new MLSListing(htmlCode);
                    }

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    REOcBPO bpoform = new REOcBPO(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion

                    #region basementlogic




                    #endregion

                    #region garagelogic

                    #endregion

                    #region fireplacelogic

                    #endregion

                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                }

                #endregion

                //
                //solutionstar - First Pass
                //
                #region solutionstar
                if (currentUrl.ToLower().Contains("solutionstar"))
                {
                    #region code

                    s = iim.iimPlayCode(macro12.ToString());
                    htmlCode = iim.iimGetLastExtract();
                    if (subjectAttachedRadioButton.Checked)
                    {
                        m = new AttachedListing(htmlCode);
                    }
                    else if (subjectDetachedradioButton.Checked)
                    {
                        m = new DetachedListing(htmlCode);
                    }
                    else
                    {
                        m = new MLSListing(htmlCode);
                    }

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    SolutionStar bpoform = new SolutionStar(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion

                    #region basementlogic




                    #endregion

                    #region garagelogic

                    #endregion

                    #region fireplacelogic

                    #endregion

                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                }

                #endregion

                //
                //AVM - First Pass
                //
                #region AVM
                if (currentUrl.ToLower().Contains("avm.assetval.com"))
                {
                    #region code

                    s = iim.iimPlayCode(macro12.ToString());
                    htmlCode = iim.iimGetLastExtract();
                    if (subjectAttachedRadioButton.Checked)
                    {
                        m = new AttachedListing(htmlCode);
                    }
                    else if (subjectDetachedradioButton.Checked)
                    {
                        m = new DetachedListing(htmlCode);
                    }
                    else
                    {
                        m = new MLSListing(htmlCode);
                    }

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    AVM bpoform = new AVM(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion

                    #region basementlogic




                    #endregion

                    #region garagelogic

                    #endregion

                    #region fireplacelogic

                    #endregion

                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                }

                #endregion

                //
                //M2M Fannie Mae Form
                //
                #region M2M Fannie Mae
                if (currentUrl.ToLower().Contains("ordereditwizardfnma"))
                {
                    #region code

                    s = iim.iimPlayCode(macro12.ToString());
                    htmlCode = iim.iimGetLastExtract();
                    if (subjectAttachedRadioButton.Checked)
                    {
                        m = new AttachedListing(htmlCode);
                    }
                    else if (subjectDetachedradioButton.Checked)
                    {
                        m = new DetachedListing(htmlCode);
                    }
                    else
                    {
                        m = new MLSListing(htmlCode);
                    }

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    M2M bpoform = new M2M(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion

                    #region basementlogic




                    #endregion

                    #region garagelogic

                    #endregion

                    #region fireplacelogic

                    #endregion

                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                }

                #endregion

                //
                //Dispo
                //
                #region Dispo
                if (currentUrl.ToLower().Contains("exceleras"))
                {
                    #region code
                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    Exceleras bpoform = new Exceleras();

                    fieldList.Add("MlsNumber", mlsnum);

                    fieldList.Add("Address_StreetAddress", full_street_address);
                    fieldList.Add("Address_City", city);
                    fieldList.Add("*Address_State", "IL");
                    fieldList.Add("Address_ZipCode", zip);
                    fieldList.Add("Address_County", county);
                    fieldList.Add("ProximityToSubject_EstimatedDistance", Get_Distance(SubjectFullAddress.Replace(",", " "), full_street_address + " " + city + " IL " + zip));
                    fieldList.Add("SubdivisionName", realist_subdivision);

                    fieldList.Add("ListingHistory_OriginalListPrice", orig_list_price);
                    fieldList.Add("ListingHistory_CurrentListPrice", current_list_price);
                    fieldList.Add("ListingHistory_CurrentListPriceDate", lastPriceChangeDate.ToShortDateString());
                    fieldList.Add("ListingHistory_DaysOnMarket", dom);

                    fieldList.Add("SaleHistory_OriginalListPrice", orig_list_price);
                    fieldList.Add("SaleHistory_FinalListPrice", current_list_price);
                    fieldList.Add("SaleHistory_SalePrice", sold_price);
                    fieldList.Add("ClosingDetails_UnderContractDate", contract_date);
                    fieldList.Add("ClosingDetails_ClosingDate", closed_date);
                    fieldList.Add("SaleHistory_DaysOnMarket", dom);

                    if (m.TransactionType == "REO")
                    {
                        fieldList.Add("*TransactionType", "1");
                    }
                    else if (m.TransactionType == "Short Sale")
                    {
                        fieldList.Add("*TransactionType", "2");
                    }
                    else
                    {
                        fieldList.Add("*TransactionType", "3");
                    }





                    if (SubjectMlsType.ToLower().Contains("1 story"))
                    {
                        fieldList.Add("*Structure_StyleType", "1<SP>Story");
                    }
                    else if (SubjectMlsType.ToLower().Contains("2 stories"))
                    {
                        fieldList.Add("*Structure_StyleType", "2<SP>Story");
                    }
                    else if (SubjectMlsType.ToLower().Contains("raised ranch"))
                    {
                        fieldList.Add("*Structure_StyleType", "Multilevel");
                    }
                    else if (SubjectMlsType.ToLower().Contains("split") || SubjectMlsType.ToLower().Contains("other"))
                    {
                        fieldList.Add("*Structure_StyleType", "Multilevel");
                    }
                    else if (SubjectMlsType.ToLower().Contains("1.5"))
                    {
                        fieldList.Add("*StyleType", "1.5<SP>Story");
                    }

                    fieldList.Add("*PropertyType", "1");
                    fieldList.Add("Site_LotSize", mls_lot_size);

                    fieldList.Add("Structure_YearBuilt", year_built);
                    fieldList.Add("Structure_LivingAreaSquareFeet", mls_gla);
                    fieldList.Add("Structure_GradedRooms_TotalRoomCount", room_count);
                    fieldList.Add("Structure_GradedRooms_BedroomCount", bedrooms);
                    fieldList.Add("Structure_GradedRooms_FullBathCount", full_bath);
                    fieldList.Add("Structure_GradedRooms_PowderRoomCount", half_bath);



                    #endregion

                    #region basementlogic

                    //Basement 
                    if (basement.ToLower().Contains("none"))
                    {
                        fieldList.Add("*SquareFeet", "0");
                        fieldList.Add("Structure*Basement*PercentFinished", "0");
                        fieldList.Add("BasementType", "crawl");
                    }
                    else
                    {
                        if (fullBasement)
                        {
                            fieldList.Add("BasementType", "Full");
                            if (type == "1 Story")
                            {
                                fieldList.Add("Structure*Basement*SquareFeet", mls_gla);
                            }
                            else if (type == "2 Stories")
                            {
                                fieldList.Add("Structure*Basement*SquareFeet", (Convert.ToInt64(mls_gla) / 2).ToString());
                            }
                            else
                            {
                                fieldList.Add("Structure*Basement*SquareFeet", (Convert.ToInt64(mls_gla) / 3).ToString());
                            }
                        }

                        if (partialBasement)
                        {
                            fieldList.Add("BasementType", "Partial");
                            if (type == "1 Story")
                            {
                                fieldList.Add("Structure*Basement*SquareFeet", (Convert.ToInt64(mls_gla) / 2).ToString());
                            }
                            else if (type == "2 Stories")
                            {
                                fieldList.Add("Structure*Basement*SquareFeet", (Convert.ToInt64(mls_gla) / 4).ToString());
                            }
                            else
                            {
                                fieldList.Add("Structure*Basement*SquareFeet", (Convert.ToInt64(mls_gla) / 6).ToString());
                            }
                        }

                        if (finished_basement)
                        {
                            fieldList.Add("Structure*Basement*PercentFinished", "100");
                        }
                        else
                        {
                            fieldList.Add("Structure*Basement*PercentFinished", "0");
                        }
                    }


                    #endregion

                    #region garagelogic
                    string garStr = "";
                    if (mls_garage_spaces == "1")
                    {
                        garStr = "1<SP>Car";
                    }
                    else
                    {
                        garStr = mls_garage_spaces + "<SP>Cars";
                    }


                    if (mls_garage_type.ToLower().Contains("att"))
                    {
                        fieldList.Add("*Structure_Garages_AttachedCarCountType", garStr);
                    }
                    if (mls_garage_type.ToLower().Contains("det"))
                    {
                        fieldList.Add("*Structure_Garages_DetachedCarCountType", garStr);
                    }
                    #endregion

                    #region fireplacelogic
                    fieldList.Add("Fireplace", hasFireplace.ToString());
                    #endregion

                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                }

                #endregion
                //
                //Goodman Dean
                //
                #region goodmandean code

                if (currentUrl.ToLower().Contains("goodmandean"))
                {

                    GoodmanDean bpoform = new GoodmanDean();
                    Dictionary<string, string> fieldList = new Dictionary<string, string>();

                    fieldList.Add("Address", full_street_address);
                    fieldList.Add("City", city);
                    fieldList.Add("Zip", zip);
                    fieldList.Add("Distance", Get_Distance(full_street_address + ", " + zip, SubjectFullAddress));

                    //fieldList.Add("PRICEREVDT", lastPriceChangeDate.ToString("d", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                    //fieldList.Add("ReductionCount", count.ToString());
                    fieldList.Add("Design", type);
                    fieldList.Add("TotalRooms", room_count);
                    fieldList.Add("Bed", bedrooms);
                    //fieldList.Add("Bath", full_bath + "." + half_bath);
                    fieldList.Add("Bath", full_bath);
                    fieldList.Add("HalfBath", half_bath);
                    fieldList.Add("MLS", mlsnum);

                    //
                    //bpoform10
                    //
                    if (currentUrl.ToLower().Contains("bpoform10"))
                    {

                        //property characteristics
                        fieldList.Add("HOA", "0");
                        fieldList.Add("Location", "Suburban");
                        fieldList.Add("Other", realist_subdivision);
                        fieldList.Add("*Condition", "Average");
                        fieldList.Add("*Overall", "Similar");
                        fieldList.Add("Comparison", "Similar");
                        fieldList.Add("*PropPool", "0");

                        //sale characteristics
                        fieldList.Add("OriginalList", orig_list_price);

                        fieldList.Add("View", "Residential");
                        fieldList.Add("Units", "1");

                        //fields that need logic
                        if (string.IsNullOrEmpty(financingType))
                        {
                            fieldList.Add("LoanType", "NA");
                        }
                        else
                        {
                            fieldList.Add("LoanType", financingType);
                        }
                        if (fullBasement && finished_basement)
                        {
                            fieldList.Add("*BasementType", "Full");
                            fieldList.Add("BsmtSqFt", mls_gla);
                            fieldList.Add("BsftFin", "100");
                        }
                        else if (fullBasement && !finished_basement)
                        {
                            fieldList.Add("*BasementType", "Partial");
                            fieldList.Add("BsmtSqFt", mls_gla);
                            fieldList.Add("BsftFin", "0");
                        }
                        else if (partialBasement && !finished_basement)
                        {
                            fieldList.Add("*BasementType", "Partial");
                            fieldList.Add("BsmtSqFt", (Convert.ToInt64(mls_gla) / 2).ToString());
                            fieldList.Add("BsftFin", "0");
                        }
                        else if (partialBasement && finished_basement)
                        {
                            fieldList.Add("*BasementType", "Partial");
                            fieldList.Add("BsmtSqFt", (Convert.ToInt64(mls_gla) / 2).ToString());
                            fieldList.Add("BsftFin", "100");
                        }
                        else
                        {
                            fieldList.Add("*BasementType", "None");
                            fieldList.Add("BsmtSqFt", "0");
                            fieldList.Add("BsftFin", "0");
                        }

                        if (sale_type == "Normal")
                        {
                            fieldList.Add("*DistressedSale", "No");
                        }
                        else if (sale_type == "Short")
                        {
                            fieldList.Add("*DistressedSale", "Short<SP>Sale");
                        }
                        else if (sale_type == "REO")
                        {
                            fieldList.Add("*DistressedSale", "REO");
                        }
                    }

                    if (parking.ToLower().Contains("gar"))
                    {
                        if (mls_garage_type.ToLower().Contains("att"))
                        {
                            //attached gar
                            fieldList.Add("ParkingType", "Gar<SP>Attached");
                            fieldList.Add("ParkingStalls", mls_garage_spaces);
                        }
                        else
                        {
                            //we dont care, well call it det (if we have type gar and spaces and its not attached)
                            fieldList.Add("ParkingType", "Gar<SP>Dettached");
                            fieldList.Add("ParkingStalls", mls_garage_spaces);
                        }

                    }
                    else
                    {
                        fieldList.Add("ParkingType", "None");
                        fieldList.Add("ParkingStalls", "0");
                    }


                    if (sale_type == "Normal")
                    {
                        fieldList.Add("SaleType", "NO");
                    }
                    else if (sale_type == "Short")
                    {
                        fieldList.Add("SaleType", "SHORTSALE");
                    }
                    else if (sale_type == "REO")
                    {
                        fieldList.Add("SaleType", "REO");
                    }
                    //sale_type


                    //fieldList.Add("PROXIMITY", this.Get_Distance(SubjectFullAddress, address));
                    //fieldList.Add("PROPTYPE", type);

                    if (sale_or_list_flag == "sale")
                    {
                        fieldList.Add("List", current_list_price);
                        fieldList.Add("Sale", sold_price);
                        //
                        //dont know why i did the below, but it doesnt work on evalform2
                        //
                        // fieldList.Add("ListDate", closed_date);
                        // fieldList.Add("SaleDate", contract_date);
                        fieldList.Add("ListDate", contract_date);
                        fieldList.Add("SaleDate", closed_date);
                    }
                    else
                    {
                        fieldList.Add("List", orig_list_price);
                        fieldList.Add("SalePrice", current_list_price);
                        fieldList.Add("ListDate", list_date);
                        fieldList.Add("SaleDate", closed_date);
                    }

                    fieldList.Add("DOM", dom);


                    fieldList.Add("Lot", mls_lot_size);
                    fieldList.Add("YearBuilt", year_built);
                    fieldList.Add("SqFt", mls_gla);

                    // fieldList.Add("LIVINGSQFT", mls_gla);

                    //fieldList.Add("*LeaseHoldFeeSimple", "Fee Simple");
                    //fieldList.Add("*State", "IL");
                    // fieldList.Add("*NumUnits", "1 SFR");


                    // fieldList.Add("*DesignAppeal", "Average");
                    // fieldList.Add("*QualityOfConstruction", "Average");

                    //  fieldList.Add("*FunctionalUtility", "Average");


                    fieldList.Add("Concessions", "NA");
                    //  fieldList.Add("HeatingCooling", "GFA.C");
                    // fieldList.Add("EnergyEfficientItems", "na");
                    //  fieldList.Add("OtherItems", "na");






                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                }

                #endregion
                //
                //Old Republic
                //
                #region oldrepublic code

                if (currentUrl.ToLower().Contains("ort.quandis"))
                {

                    //mportFormID=2308914

                    OldRep bpoform = new OldRep();
                    Dictionary<string, string> fieldList = new Dictionary<string, string>();

                    iim2.iimPlayCode("TAG POS=1 TYPE=TABLE ATTR=ID:Table1 EXTRACT=TXT");
                    if (iim2.iimGetLastExtract().Contains("F-V8"))
                    {
                        bpoform = new OldRepFormV8();
                        fieldList.Add("Bathrooms", full_bath);
                        fieldList.Add("BathroomsHalf", half_bath);
                        fieldList.Add("MLSNumber", mlsnum);
                        fieldList.Add("*View", "Neutral");
                        fieldList.Add("*LocationFactor", "Residential");


                    }
                    else
                    {
                        fieldList.Add("Bathrooms", full_bath + "." + half_bath);
                        fieldList.Add("*View", "Average");
                    }



                    fieldList.Add("Address", full_street_address);
                    fieldList.Add("City", city);
                    fieldList.Add("PostalCode", zip);


                    //fieldList.Add("PRICEREVDT", lastPriceChangeDate.ToString("d", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                    fieldList.Add("ReductionCount", count.ToString());

                    fieldList.Add("Rooms", room_count);


                    fieldList.Add("Bedrooms", bedrooms);


                    //needs logic for selection list
                    //needs logic for selection list

                    if (fullBasement && finished_basement)
                    {
                        fieldList.Add("*Basement", "90% Finished");
                    }
                    else if (fullBasement || partialBasement && !finished_basement)
                    {
                        fieldList.Add("*Basement", "Unfinished");
                    }
                    else if (partialBasement && finished_basement)
                    {
                        fieldList.Add("*Basement", "50% Finished");
                    }
                    else
                    {
                        fieldList.Add("*Basement", "None");
                    }


                    if (parking.ToLower().Contains("gar"))
                    {
                        if (mls_garage_type.ToLower().Contains("att"))
                        {
                            //attached gar
                            fieldList.Add("*Parking", mls_garage_spaces + "CarAtt");
                        }
                        else
                        {
                            //we dont care, well call it det (if we have type gar and spaces and its not attached)
                            fieldList.Add("*Parking", mls_garage_spaces + "CarDet");
                        }

                    }
                    else
                    {
                        fieldList.Add("*Parking", "None");
                    }



                    if (m.TransactionType == "ShortSale")
                    {
                        fieldList.Add("*SaleType", "Short Sale");
                    }
                    else if (m.TransactionType == "REO")
                    {
                        fieldList.Add("*SaleType", "REO");
                    }
                    else
                    {
                        fieldList.Add("*SaleType", "Standard");
                    }


                    //fieldList.Add("PROXIMITY", this.Get_Distance(SubjectFullAddress, address));
                    //fieldList.Add("PROPTYPE", type);

                    if (sale_or_list_flag == "sale")
                    {
                        fieldList.Add("ListPrice", current_list_price);
                        fieldList.Add("SalePrice", sold_price);
                    }
                    else
                    {
                        fieldList.Add("ListPrice", orig_list_price);
                        fieldList.Add("SalePrice", current_list_price);
                    }

                    fieldList.Add("SaleDate", closed_date);
                    fieldList.Add("DOM", dom);


                    fieldList.Add("Site", mls_lot_size);
                    fieldList.Add("Age", year_built);
                    fieldList.Add("SquareFeet", mls_gla);
                    // fieldList.Add("LIVINGSQFT", mls_gla);

                    fieldList.Add("*LeaseHoldFeeSimple", "Fee Simple");
                    fieldList.Add("*State", "IL");
                    fieldList.Add("*NumUnits", "1 SFR");
                    fieldList.Add("*LocationCondition", "Suburban");

                    fieldList.Add("*DesignAppeal", "Average");
                    fieldList.Add("*QualityOfConstruction", "Average");
                    fieldList.Add("*Condition", "Average");
                    fieldList.Add("*FunctionalUtility", "Average");
                    fieldList.Add("*DataSource", "MLS");

                    fieldList.Add("Incentive", "na");
                    fieldList.Add("HeatingCooling", "GFA.C");
                    fieldList.Add("EnergyEfficientItems", "na");
                    fieldList.Add("OtherItems", "na");
                    fieldList.Add("*Style", bpoform.StyleString(type));

                    //ifcContentCompOneSaleType
                    //SchoolID
                    //Style




                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                }

                #endregion
                //
                //Landsafe
                //
                #region landsafe
                if (currentUrl.ToLower().Contains(@"collateraldna.com/forms"))
                {
                    LandSafe bpoform = new LandSafe();
                    Dictionary<string, string> fieldList = new Dictionary<string, string>();

                    s = iim.iimPlayCode(macro12.ToString());
                    htmlCode = iim.iimGetLastExtract();
                    if (subjectAttachedRadioButton.Checked)
                    {
                        m = new AttachedListing(htmlCode);
                    }
                    else if (subjectDetachedradioButton.Checked)
                    {
                        m = new DetachedListing(htmlCode);
                    }
                    else
                    {
                        m = new MLSListing(htmlCode);
                    }

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    fieldList.Add("ADDRSTREET", full_street_address);
                    fieldList.Add("STREET", full_street_address);
                    fieldList.Add("ADDRCITY", city);
                    fieldList.Add("CITY", city);
                    fieldList.Add("ADDRSTATEPROV", "IL");
                    fieldList.Add("STATEPROV", "IL");
                    fieldList.Add("ADDRZIP", zip);
                    fieldList.Add("ZIP", zip);
                    fieldList.Add("LISTINGPRICE", current_list_price);
                    fieldList.Add("PRICEREVDT", lastPriceChangeDate.ToString("d", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                    fieldList.Add("BASEMENT", basement);
                    fieldList.Add("BEDROOMS", bedrooms);
                    fieldList.Add("BATH", full_bath + "." + half_bath);
                    fieldList.Add("GARAGE", mls_garage_type);
                    fieldList.Add("PROXIMITY", this.Get_Distance(SubjectFullAddress, address));
                    fieldList.Add("SALESPRICE", sold_price);
                    fieldList.Add("SALEDT", m.SalesDate.ToString("d", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                    fieldList.Add("MARKETDAYS", dom);
                    fieldList.Add("PROPTYPE", type);
                    fieldList.Add("LOTSIZE", mls_lot_size);
                    fieldList.Add("YRBUILT", year_built);
                    fieldList.Add("GBASQFT", mls_gla);
                    fieldList.Add("LIVINGSQFT", mls_gla);
                    fieldList.Add("DESIGNSTYLE", m.ExteriorMlsString);

                    if (m.TransactionType == "REO" || m.TransactionType == "ShortSale")
                    {
                        fieldList.Add("*SALETYPE", "YES");
                    }
                    else
                    {
                        fieldList.Add("*SALETYPE", "NO");
                    }




                    //try
                    //{
                    //    if (Convert.ToInt32(half_bath) != 0)
                    //    {
                    //        fieldList.Add("*BA", full_bath + full_bath);
                    //    }
                    //    else
                    //    {
                    //        fieldList.Add("*BA", full_bath);
                    //    }
                    //}

                    //catch
                    //{
                    //    fieldList.Add("*BA", full_bath);
                    //}



                    //fieldList.Add("*Type", "s");



                    ////fieldList.Add("*Style", "2 Story Conv");


                    //fieldList.Add("*Loc", "r");

                    //fieldList.Add("GLASqFt", mls_gla);
                    //fieldList.Add("EstValSqFt", "0");
                    //if (finished_basement)
                    //{
                    //    fieldList.Add("BsmtPerFin", "100");
                    //}
                    //else
                    //{
                    //    fieldList.Add("BsmtPerFin", "0");
                    //}

                    //fieldList.Add("EstValBsmtPerFin", "0");
                    //fieldList.Add("YrBlt", year_built);
                    //fieldList.Add("EstValYrBlt", "0");
                    //fieldList.Add("LotSize", mls_lot_size + "ac");
                    //fieldList.Add("EstValLotSize", "0");
                    //fieldList.Add("DOM", dom);
                    //fieldList.Add("EstValDOM", "0");
                    //fieldList.Add("Other", "n/a");
                    //fieldList.Add("EstValOther", "0");
                    //if (parking.ToLower().Contains("gar"))
                    //{
                    //    fieldList.Add("*Gar", mls_garage_spaces);
                    //}
                    //else
                    //{
                    //    fieldList.Add("*Gar", "n");
                    //}
                    //fieldList.Add("EstValGar", "0");
                    //fieldList.Add("*Pool", "n");


                    //fieldList.Add("ExtAdd", "0");
                    //fieldList.Add("EstValExt", "0");
                    //fieldList.Add("*Land", "r");
                    //fieldList.Add("Fin", "n/a");
                    //fieldList.Add("EstValFin", "0");
                    //fieldList.Add("*Cond", "a");
                    //fieldList.Add("EstValCond", "0");
                    //fieldList.Add("Comm", "Equal to subject in age, size, type and location");
                    //fieldList.Add("OLPrice", orig_list_price);
                    //fieldList.Add("OLDate", list_date);
                    //fieldList.Add("CLPrice", current_list_price);
                    //fieldList.Add("CLDate", lastPriceChangeDate.ToShortDateString());
                    //fieldList.Add("SPrice", sold_price);
                    //fieldList.Add("SDate", closed_date);
                    //fieldList.Add("SLPrice", current_list_price);
                    //fieldList.Add("SaleType", sale_type);
                    //fieldList.Add("Source", "MLS");
                    //fieldList.Add("SourceID", mlsnum);
                    //fieldList.Add("SDistrict", "Unknown");
                    //fieldList.Add("SDivision", "n/a");
                    //fieldList.Add("MLSArea", "n/a");
                    //fieldList.Add("*Construct", "a");




                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList, this);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                }
                #endregion
                //
                //Equitrax
                //
                #region equitrax
                //form Z or form X?, completely different systems
                //form z and n, very similar
                if (currentUrl.ToLower().Contains("equi-trax"))
                {
                    bool formX = true;

                    if (formX)
                    {
                        #region formx
                        {
                            Equitrax bpoform = new Equitrax();
                            Dictionary<string, string> fieldList = new Dictionary<string, string>();
                            fieldList.Add("BsmtSqFt", m.BasementGLA());
                            fieldList.Add("BsmtFinSqFt", m.BasementFinishedGLA());
                            fieldList.Add("MLS", mlsnum);

                            fieldList.Add("filepath", SubjectFilePath);
                            fieldList.Add("Address", full_street_address);
                            fieldList.Add("City", city);
                            fieldList.Add("*State", "ii");
                            fieldList.Add("Zip", zip);
                            fieldList.Add("*TR", room_count);
                            fieldList.Add("*BR", bedrooms);
                            fieldList.Add("*REO", m.DistressedSaleYesNo());

                            try
                            {
                                if (Convert.ToInt32(half_bath) != 0)
                                {
                                    fieldList.Add("*BA", full_bath + full_bath);
                                }
                                else
                                {
                                    fieldList.Add("*BA", full_bath);
                                }
                            }

                            catch
                            {
                                fieldList.Add("*BA", full_bath);
                            }



                            fieldList.Add("*Type", "s");



                            fieldList.Add("*Style", m.Type.Replace(" ", "<SP>"));



                            fieldList.Add("*Loc", "Good");

                            fieldList.Add("GLASqFt", mls_gla);
                            fieldList.Add("EstValSqFt", "0");
                            if (finished_basement)
                            {
                                fieldList.Add("BsmtPerFin", "100");
                            }
                            else
                            {
                                fieldList.Add("BsmtPerFin", "0");
                            }

                            fieldList.Add("EstValBsmtPerFin", "0");
                            fieldList.Add("YrBlt", year_built);
                            fieldList.Add("EstValYrBlt", "0");
                            fieldList.Add("LotSize", mls_lot_size + "ac");
                            fieldList.Add("EstValLotSize", "0");
                            fieldList.Add("DOM", dom);
                            fieldList.Add("EstValDOM", "0");
                            fieldList.Add("Other", "n/a");
                            fieldList.Add("EstValOther", "0");
                            if (parking.ToLower().Contains("gar"))
                            {
                                fieldList.Add("*Gar", mls_garage_spaces);
                            }
                            else
                            {
                                fieldList.Add("*Gar", "n");
                            }
                            fieldList.Add("EstValGar", "0");
                            fieldList.Add("*Pool", "n");


                            fieldList.Add("ExtAdd", "0");
                            fieldList.Add("EstValExt", "0");
                            fieldList.Add("*Land", "r");
                            fieldList.Add("Fin", m.FinancingMlsString);
                            fieldList.Add("EstValFin", "0");
                            fieldList.Add("*Cond", "a");
                            fieldList.Add("EstValCond", "0");
                            fieldList.Add("Comm", "Equal to subject in age, size, type and location");
                            fieldList.Add("OLPrice", orig_list_price);
                            fieldList.Add("OLDate", list_date);
                            fieldList.Add("CLPrice", current_list_price);
                            fieldList.Add("CLDate", lastPriceChangeDate.ToShortDateString());
                            fieldList.Add("SPrice", sold_price);
                            fieldList.Add("SDate", closed_date);
                            fieldList.Add("SLPrice", current_list_price);
                            fieldList.Add("SaleType", sale_type);
                            fieldList.Add("Source", "MLS");
                            fieldList.Add("SourceID", mlsnum);
                            fieldList.Add("SDistrict", "Unknown");
                            fieldList.Add("Subdivision", mls_subdivision.Replace(" ", "<SP>"));
                            fieldList.Add("MLSArea", "n/a");
                            fieldList.Add("*Construct", "a");
                            fieldList.Add("Amenities1", "n/a");
                            fieldList.Add("Amenities2", "n/a");

                        #endregion
                            //
                            //Form Z additions
                            //
                            if (bpoform.ReportType(iim2) == "z" || bpoform.ReportType(iim2) == "n")
                            {
                                #region form Z and N additions
                                fieldList.Add("LDate", list_date);
                                fieldList.Add("LaSPrice", current_list_price);
                                fieldList.Add("SqFt", mls_gla);
                                fieldList.Add("EnergyEff", "N/A");
                                fieldList.Add("Ext", "N/A");
                                fieldList.Add("*REOCorp", "No");

                                fieldList.Add("FencePool", "N/A");
                                if (SubjectAttached)
                                {
                                    fieldList.Add("Site", "0ac");
                                }
                                else
                                {
                                    fieldList.Add("Site", mls_lot_size + "ac");
                                }

                                fieldList.Add("AC", "GFA.C");

                                //Bath
                                #region bathroom
                                fieldList.Remove("*BA");
                                fieldList.Add("*BA", full_bath);
                                fieldList.Add("*BAH", half_bath);

                                //string tBath = fieldList["*BA"];
                                //try
                                //{
                                //    fieldList.Add("*BAH", half_bath);
                                //    //if there is a half bath, assuming /d./d pattern, hence the try block, incase /d format 
                                //    if (Convert.ToInt16(tBath[2].ToString()) > 0)
                                //    {
                                //        tBath = tBath[0].ToString() + ".5";
                                //    }
                                //    else
                                //    {
                                //        tBath = tBath[0].ToString();
                                //    }
                                //    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropBA CONTENT=" + tBath);
                                //    fieldList.Remove("*BA");
                                //    fieldList.Add("*BA", tBath);

                                //}

                                //catch
                                //{

                                //}

                                #endregion

                                //Design Appeal
                                #region type selection
                                fieldList.Remove("*Type");

                                fieldList.Add("*Type", bpoform.TypeString(m));


                                if (type.ToLower().Contains("1 story") || type.ToLower().Contains("ranch"))
                                {
                                    fieldList.Add("*Appeal", "Single<SP>Story");
                                }
                                else if (type.ToLower().Contains("2 stories") || type.ToLower().Contains("townhome"))
                                {
                                    fieldList.Add("*Appeal", "2-Story<SP>Conv");
                                }
                                else if (type.ToLower().Contains("raised ranch"))
                                {
                                    fieldList.Add("*Appeal", "Split/Bi-Level");
                                }
                                else if (type.ToLower().Contains("split") || type.ToLower().Contains("other"))
                                {
                                    fieldList.Add("*Appeal", "Tri/Muilti-Level");
                                }
                                else if (type.ToLower().Contains("1.5"))
                                {
                                    fieldList.Add("*Appeal", "Cape");
                                }
                                else
                                {
                                    fieldList.Add("*Appeal", "2-Story<SP>Conv");
                                }
                                #endregion
                                //Garage
                                #region garage

                                fieldList.Remove("*Gar");
                                //fieldList.Add("*Gar", bpoform.GarageString(m));
                                string contentString = "";
                                if (mls_garage_type.ToLower().Contains("att"))
                                {
                                    contentString = (mls_garage_spaces + "<SP>Attached");
                                }
                                else if (mls_garage_type.ToLower().Contains("det"))
                                {
                                    contentString = (mls_garage_spaces + "<SP>Detached");
                                }
                                else
                                {
                                    contentString = "None";
                                }
                                fieldList.Add("*Gar", contentString);
                                #endregion


                                fieldList.Add("BasementSF", m.BasementGLA());
                                fieldList.Add("Basement", m.BasementFinishedPercentage());

                                //    //Basement 
                                //    if (basement.ToLower().Contains("none"))
                                //    {
                                //        fieldList.Add("BasementSF", m.BasementGLA());
                                //        fieldList.Add("Basement", "0");
                                //    }
                                //    else
                                //    {
                                //        try
                                //        {
                                //            if (fullBasement)
                                //            {
                                //                if (type == "1 Story")
                                //                {
                                //                    fieldList.Add("BasementSF", mls_gla);
                                //                }
                                //                else if (type == "2 Stories")
                                //                {
                                //                    fieldList.Add("BasementSF", (Convert.ToInt64(mls_gla) / 2).ToString());
                                //                }
                                //                else
                                //                {
                                //                    fieldList.Add("BasementSF", (Convert.ToInt64(mls_gla) / 3).ToString());
                                //                }
                                //            }

                                //            if (partialBasement)
                                //            {
                                //                if (type == "1 Story")
                                //                {
                                //                    fieldList.Add("BasementSF", (Convert.ToInt64(mls_gla) / 2).ToString());
                                //                }
                                //                else if (type == "2 Stories")
                                //                {
                                //                    fieldList.Add("BasementSF", (Convert.ToInt64(mls_gla) / 4).ToString());
                                //                }
                                //                else
                                //                {
                                //                    fieldList.Add("BasementSF", (Convert.ToInt64(mls_gla) / 6).ToString());
                                //                }
                                //            }

                                //            if (finished_basement)
                                //            {
                                //                fieldList.Add("Basement", "100");
                                //            }
                                //            else
                                //            {
                                //                fieldList.Add("Basement", "0");
                                //            }
                                //        }
                                //        catch { }
                                //        }
                                #endregion
                            }

                            bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                            status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                            iim.iimPlayCode(macro12.ToString());
                            htmlCode = iim.iimGetLastExtract();

                            // Perform the increment on the ProgressBar.
                            pBar2.PerformStep();

                            if (subjectAttachedRadioButton.Checked)
                            {
                                m = new AttachedListing(htmlCode);
                            }
                            else if (subjectDetachedradioButton.Checked)
                            {
                                m = new DetachedListing(htmlCode);
                            }
                            else
                            {
                                m = new MLSListing(htmlCode);
                            }
                        }
                    }
                    else
                    {
                        #region not currently used
                        Equitrax bpoform = new Equitrax();
                        Dictionary<string, string> fieldList = new Dictionary<string, string>();

                        fieldList.Add("Address", full_street_address);
                        fieldList.Add("City", city);
                        //
                        //proximity button
                        //
                        //
                        //Reo Y/N dropdown
                        //
                        fieldList.Add("SPrice", sold_price);
                        fieldList.Add("SDate", closed_date);
                        fieldList.Add("OLPrice", orig_list_price);
                        fieldList.Add("LDate", list_date);
                        fieldList.Add("LaSPrice", current_list_price);
                        fieldList.Add("*BR", bedrooms);
                        try
                        {
                            if (Convert.ToInt32(half_bath) != 0)
                            {
                                fieldList.Add("*BA", full_bath + full_bath);
                            }
                            else
                            {
                                fieldList.Add("*BA", full_bath);
                            }
                        }

                        catch
                        {
                            fieldList.Add("*BA", full_bath);
                        }
                        fieldList.Add("*TR", room_count);
                        fieldList.Add("*Type", "s");
                        fieldList.Add("SqFt", mls_gla);
                        fieldList.Add("*Loc", "g");
                        fieldList.Add("YrBlt", year_built);
                        if (parking.ToLower().Contains("gar"))
                        {
                            fieldList.Add("*Gar", mls_garage_spaces);
                        }
                        else
                        {
                            fieldList.Add("*Gar", "n");
                        }
                        //
                        //Patio/Porches text field
                        //
                        //
                        //fences/pools txt field
                        //
                        fieldList.Add("Fin", "n/a");
                        fieldList.Add("*Leasehold", "f");
                        fieldList.Add("LotSize", mls_lot_size + "ac");
                        fieldList.Add("*View", "Residential");
                        //
                        //Design / Appeal dropdown
                        //
                        //
                        //bsmt gla txt box
                        //
                        //
                        //bsmt fin %
                        //
                        fieldList.Add("*QualConst", "g");
                        fieldList.Add("*FcnUtil", "y");
                        //
                        //hvac text box
                        //
                        fieldList.Add("EnergyEff", "N/A");
                        fieldList.Add("Other", "N/A");
                        fieldList.Add("Source", "MLS");
                        fieldList.Add("*Cond", "g");

                        //if (finished_basement)
                        //{
                        //    fieldList.Add("BsmtPerFin", "100");
                        //}
                        //else
                        //{
                        //    fieldList.Add("BsmtPerFin", "0");
                        //}


                        fieldList.Add("DOM", dom);








                        fieldList.Add("CLDate", lastPriceChangeDate.ToShortDateString());


                        fieldList.Add("CLPrice", current_list_price);
                        fieldList.Add("SaleType", sale_type);

                        fieldList.Add("SourceID", mlsnum);
                        fieldList.Add("SDistrict", "Unknown");
                        fieldList.Add("SDivision", "n/a");
                        fieldList.Add("MLSArea", "n/a");
                        fieldList.Add("*Construct", "a");




                        bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                        status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                        #endregion
                    }
                }

                #endregion
                //
                //lres
                //
                #region lres
                if (currentUrl.ToLower().Contains("lres"))
                {
                    #region code


                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    stack.Add(input_comp_name, m);

                    if (stack.Count == 6)
                    {
                        Dictionary<string, string> fieldList = new Dictionary<string, string>();
                        fieldList.Add("filepath", SubjectFilePath);
                        //for (int j = 0; j++; j<6)
                        //{

                        //}
                        StringBuilder macro = new StringBuilder();

                        macro.AppendLine(@"TAG POS=2 TYPE=INPUT:RADIO FORM=NAME:bpoForm ATTR=NAME:S1_closest_comparable");

                        macroCode = macro.ToString();
                        iim2.iimPlayCode(macroCode, 60);

                        Lres bpoforms1 = new Lres(stack["RECENT_SALE1"]);
                        bpoforms1.CompFill(iim2, "sale", "RECENT_SALE1", fieldList);

                        Lres bpoforms2 = new Lres(stack["RECENT_SALE2"]);
                        bpoforms2.CompFill(iim2, "sale", "RECENT_SALE2", fieldList);

                        Lres bpoforms3 = new Lres(stack["RECENT_SALE3"]);
                        bpoforms3.CompFill(iim2, "sale", "RECENT_SALE3", fieldList);





                        //macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");

                        macro.Clear();
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:bpoForm ATTR=NAME:btnSaveAndCont");
                        macro.AppendLine(@"WAIT SECONDS=3");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:bpoForm ATTR=NAME:btnSaveForm");

                        macroCode = macro.ToString();
                        iim2.iimPlayCode(macroCode, 60);

                        Lres bpoforms4 = new Lres(stack["COMPARABLE1"]);
                        bpoforms4.CompFill(iim2, "list", "COMPARABLE1", fieldList);

                        Lres bpoforms5 = new Lres(stack["COMPARABLE2"]);
                        bpoforms5.CompFill(iim2, "list", "COMPARABLE2", fieldList);

                        Lres bpoforms6 = new Lres(stack["COMPARABLE3"]);
                        bpoforms6.CompFill(iim2, "list", "COMPARABLE3", fieldList);

                    }


                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                    #endregion
                }
                #endregion
                //
                //Emort
                //
                #region emort
                if (currentUrl.ToLower().Contains("emortgage"))
                {
                    Emortgage bpoform = new Emortgage();
                    Dictionary<string, string> fieldList = new Dictionary<string, string>();

                    //newest form
                    bool formE = true;


                    //common fields
                    fieldList.Add("Address", full_street_address);
                    fieldList.Add("City", city);
                    fieldList.Add("Zip", zip);


                    if (formE)
                    {

                        fieldList.Add("filepath", SubjectFilePath);


                        //fieldList.Add("*State", "ii");
                        fieldList.Add("*TR", room_count);
                        fieldList.Add("*BR", bedrooms);

                        //Baths
                        #region bathroom
                        string tBath = full_bath + "." + half_bath;
                        try
                        {
                            //if there is a half bath, assuming /d./d pattern, hence the try block, incase /d format 
                            if (Convert.ToInt16(tBath[2].ToString()) > 0)
                            {
                                tBath = tBath[0].ToString() + ".5";
                            }
                            else
                            {
                                tBath = tBath[0].ToString();
                            }
                            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropBA CONTENT=" + tBath);
                            fieldList.Add("*BA", tBath);
                        }
                        catch
                        {

                        }

                        #endregion

                        if (hasFireplace)
                        {
                            fieldList.Add("*Fireplace", "Yes");
                        }
                        else
                        {
                            fieldList.Add("*Fireplace", "No");

                        }


                        // fieldList.Add("*Type", "s");



                        //fieldList.Add("*Style", "2 Story Conv");


                        // fieldList.Add("*Loc", "r");

                        fieldList.Add("SqFt", mls_gla);

                        if (fullBasement || partialBasement)
                        {
                            fieldList.Add("*Bsmt", "Yes");
                            fieldList.Add("*BsmtSqFt", (Convert.ToInt16(mls_gla) / 2).ToString());

                            if (finished_basement)
                            {
                                fieldList.Add("BsmtPerFin", "100");
                            }
                            else
                            {
                                fieldList.Add("BsmtPerFin", "0");
                            }
                        }
                        else
                        {
                            fieldList.Add("*Bsmt", "No");
                        }



                        fieldList.Add("Appeal", "Average");
                        fieldList.Add("YrBlt", year_built);
                        //      fieldList.Add("EstValYrBlt", "0");
                        fieldList.Add("LotSize", mls_lot_size + "ac");
                        //    fieldList.Add("EstValLotSize", "0");
                        fieldList.Add("DOM", dom);
                        //    fieldList.Add("EstValDOM", "0");
                        fieldList.Add("Other", "n/a");
                        //    fieldList.Add("EstValOther", "0");
                        //if (parking.ToLower().Contains("gar"))
                        //{
                        //    fieldList.Add("*Gar", mls_garage_spaces);
                        //}
                        //else
                        //{
                        //    fieldList.Add("*Gar", "n");
                        //}
                        //   fieldList.Add("EstValGar", "0");
                        //   fieldList.Add("*Pool", "n");


                        // fieldList.Add("ExtAdd", "0");
                        //  fieldList.Add("EstValExt", "0");
                        //  fieldList.Add("*Land", "r");
                        //  fieldList.Add("Fin", "n/a");
                        //  fieldList.Add("EstValFin", "0");
                        //  fieldList.Add("*Cond", "a");
                        //  fieldList.Add("EstValCond", "0");
                        fieldList.Add("Comm", "Equal to subject in age, size, type and location");
                        fieldList.Add("OLPrice", orig_list_price);
                        fieldList.Add("OLDate", list_date);
                        fieldList.Add("CLPrice", current_list_price);
                        fieldList.Add("LRDate", lastPriceChangeDate.ToShortDateString());
                        fieldList.Add("SPrice", sold_price);
                        fieldList.Add("SDate", closed_date);
                        fieldList.Add("SLPrice", current_list_price);
                        fieldList.Add("SaleType", sale_type);
                        fieldList.Add("Source", "MLS");
                        fieldList.Add("SourceID", mlsnum);
                        fieldList.Add("SDistrict", "Unknown");
                        fieldList.Add("SDivision", "n/a");
                        fieldList.Add("MLSArea", "n/a");
                        // fieldList.Add("*Construct", "a");

                        //Bath
                        #region bathroom
                        // fieldList.Remove("*BA");
                        //fieldList.Add("*BA", full_bath);
                        //        fieldList.Add("*BAH", half_bath);

                        //string tBath = fieldList["*BA"];
                        //try
                        //{
                        //    fieldList.Add("*BAH", half_bath);
                        //    //if there is a half bath, assuming /d./d pattern, hence the try block, incase /d format 
                        //    if (Convert.ToInt16(tBath[2].ToString()) > 0)
                        //    {
                        //        tBath = tBath[0].ToString() + ".5";
                        //    }
                        //    else
                        //    {
                        //        tBath = tBath[0].ToString();
                        //    }
                        //    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropBA CONTENT=" + tBath);
                        //    fieldList.Remove("*BA");
                        //    fieldList.Add("*BA", tBath);

                        //}

                        //catch
                        //{

                        //}

                        #endregion

                        //Design Appeal
                        #region type selection
                        // fieldList.Remove("*Type");
                        fieldList.Add("*Type", bpoform.TypeString(m));


                        if (type.ToLower().Contains("1 story") || type.ToLower().Contains("ranch"))
                        {
                            fieldList.Add("*Style", "Single<SP>Story");
                        }
                        else if (type.ToLower().Contains("2 stories") || type.ToLower().Contains("townhome"))
                        {
                            fieldList.Add("*Style", "2-Story<SP>Conv");
                        }
                        else if (type.ToLower().Contains("raised ranch"))
                        {
                            fieldList.Add("*Style", "Split/Bi-Level");
                        }
                        else if (type.ToLower().Contains("split") || type.ToLower().Contains("other"))
                        {
                            fieldList.Add("*Style", "Tri/Multi-Level");
                        }
                        else if (type.ToLower().Contains("1.5"))
                        {
                            fieldList.Add("*Style", "Cape");
                        }
                        else
                        {
                            fieldList.Add("*Style", "2-Story<SP>Conv");
                        }
                        #endregion
                        //Garage
                        #region garage

                        //fieldList.Remove("*Gar");
                        fieldList.Add("*Gar", bpoform.GarageString(m));
                        //if (SubjectParkingType.ToLower().Contains("gar"))
                        //{

                        //    string contentString = "";

                        //    if (SubjectParkingType.ToLower().Contains("att"))
                        //    {
                        //        contentString = (Regex.Match(SubjectParkingType, @"\d").Value + "<SP>Attached");
                        //    }
                        //    else if (SubjectParkingType.ToLower().Contains("det"))
                        //    {
                        //        contentString = (Regex.Match(SubjectParkingType, @"\d").Value + "<SP>Detached");
                        //    }
                        //    fieldList.Add("*Gar", contentString);

                        //}
                        //else
                        //{
                        //    fieldList.Add("*Gar", "None");
                        //}
                        #endregion



                        //fieldList.Add("BasementSF", m.BasementGLA());
                        //fieldList.Add("Basement", m.BasementFinishedPercentage());






                    }
                    else  // original code/form
                    {

                        fieldList.Add("TR", room_count);
                        fieldList.Add("BR", bedrooms);
                        fieldList.Add("BA", full_bath + "." + half_bath.Replace("1", "5"));
                        fieldList.Add("EstValBABR", "0");
                        fieldList.Add("Type", "SF Detached");
                        fieldList.Add("EstValType", "0");


                        fieldList.Add("Style", "2 Story Conv");
                        fieldList.Add("EstValStyle", "0");

                        fieldList.Add("Loc", "Good");
                        fieldList.Add("EstValLoc", "0");
                        fieldList.Add("SqFt", mls_gla);
                        fieldList.Add("EstValSqFt", "0");
                        if (finished_basement)
                        {
                            fieldList.Add("BsmtPerFin", "100");
                        }
                        else
                        {
                            fieldList.Add("BsmtPerFin", "0");
                        }

                        fieldList.Add("EstValBsmtPerFin", "0");
                        fieldList.Add("YrBlt", year_built);
                        fieldList.Add("EstValYrBlt", "0");
                        fieldList.Add("LotSize", mls_lot_size + "ac");
                        fieldList.Add("EstValLotSize", "0");
                        fieldList.Add("DOM", dom);
                        fieldList.Add("EstValDOM", "0");
                        fieldList.Add("Other", "n/a");
                        fieldList.Add("EstValOther", "0");
                        if (parking.ToLower().Contains("gar"))
                        {
                            fieldList.Add("Gar", mls_garage_spaces + " " + mls_garage_type);
                        }
                        else
                        {
                            fieldList.Add("Gar", "None");
                        }
                        fieldList.Add("EstValGar", "0");
                        fieldList.Add("Pool", "No");


                        fieldList.Add("ExtAdd", "0");
                        fieldList.Add("EstValExt", "0");
                        fieldList.Add("Land", "Good");
                        fieldList.Add("Fin", "n/a");
                        fieldList.Add("EstValFin", "0");
                        fieldList.Add("Cond", "Good");
                        fieldList.Add("EstValCond", "0");
                        fieldList.Add("Comm", "Equal to subject in age, size, type and location");
                        fieldList.Add("OLPrice", orig_list_price);
                        fieldList.Add("SPrice", sold_price);
                        fieldList.Add("SDate", closed_date);
                        fieldList.Add("SLPrice", current_list_price);
                        fieldList.Add("SaleType", sale_type);
                        fieldList.Add("Source", "MLS");
                        fieldList.Add("SourceID", mlsnum);
                        fieldList.Add("SDistrict", "Unknown");
                        fieldList.Add("SDivision", "n/a");
                        fieldList.Add("MLSArea", "n/a");

                    }


                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                }
                #endregion
                //
                //Mainstreet aka BPOfulfillment aka Redbell
                //
                #region mainstreet
                if (currentUrl.ToLower().Contains("bpofulfillment.com"))
                {
                    StringBuilder macro = new StringBuilder();
                    macro.AppendLine(@"SET !ERRORIGNORE YES");
                    macro.AppendLine(@"SET !TIMEOUT_STEP 0");

                    //TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=findbpo&assetid=3394712&orderid=2067826 ATTR=NAME:BPO$Comparables$txtAddress22 CONTENT=rer
                    //TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtAddress22 CONTENT=thenewvalue2
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtAddress" + input_comp_name + " CONTENT=" + full_street_address.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtCity" + input_comp_name + " CONTENT=" + city.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtCounty" + input_comp_name + " CONTENT=" + county.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtZip" + input_comp_name + " CONTENT=" + zip);
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtProximity" + input_comp_name + " CONTENT=0.00");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtSubdivision" + input_comp_name + " CONTENT=" + mls_subdivision.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtAreaName" + input_comp_name + " CONTENT=" + mls_subdivision.Replace(" ", "<SP>"));


                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboDataSource" + input_comp_name + " CONTENT=%305");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtCurrentListDate" + input_comp_name + "$txtMonth CONTENT=" + list_date.Substring(0, 2));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtCurrentListDate" + input_comp_name + "$txtDay CONTENT=" + list_date.Substring(3, 2));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtCurrentListDate" + input_comp_name + "$txtYear CONTENT=" + list_date.Substring(6, 4));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtCurrentListPrice" + input_comp_name + " CONTENT=" + current_list_price.Replace("$", "").Replace(",", ""));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtListDate" + input_comp_name + "$txtMonth CONTENT=" + list_date.Substring(0, 2));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtListDate" + input_comp_name + "$txtDay CONTENT=" + list_date.Substring(3, 2));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtListDate" + input_comp_name + "$txtYear CONTENT=" + list_date.Substring(6, 4));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtFinalListPrice" + input_comp_name + " CONTENT=" + current_list_price.Replace("$", "").Replace(",", ""));

                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboConstruction" + input_comp_name + " CONTENT=%104");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboLandscaping" + input_comp_name + " CONTENT=%112");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboPropertyType" + input_comp_name + " CONTENT=%764");

                    if (closed_date != "")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtSalePrice" + input_comp_name + " CONTENT=" + sold_price.Replace("$", "").Replace(",", ""));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtDateOfSale" + input_comp_name + "$txtMonth CONTENT=" + closed_date.Substring(0, 2));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtDateOfSale" + input_comp_name + "$txtDay CONTENT=" + closed_date.Substring(3, 2));

                        //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtConcessions" + input_comp_name + " CONTENT=0.00");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtDateOfSale" + input_comp_name + "$txtYear CONTENT=" + closed_date.Substring(6, 4));
                    }
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtConcessions" + input_comp_name + " CONTENT=0.00");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtNumUnits" + input_comp_name + " CONTENT=1");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboPropertyType" + input_comp_name + " CONTENT=%142");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboPropertyStyle" + input_comp_name + " CONTENT=%0");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboPropertyStyle" + input_comp_name + " CONTENT=%0");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboCondition" + input_comp_name + " CONTENT=%348");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboView" + input_comp_name + " CONTENT=%485");
                    macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                    double x = -1;
                    string st = mls_lot_size;
                    Double.TryParse(st, out x);
                    x = Math.Round(x, 2);

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtYearBuilt" + input_comp_name + " CONTENT=" + year_built);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtLotSize" + input_comp_name + " CONTENT=" + x.ToString());

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtSite" + input_comp_name + " CONTENT=Level/100%");

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtTotalSQ" + input_comp_name + " CONTENT=" + mls_gla);
                    macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtTotalRooms" + input_comp_name + " CONTENT=" + room_count);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtBedrooms" + input_comp_name + " CONTENT=" + bedrooms);
                    macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                    //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtBedrooms2 CONTENT=3");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtFullBath" + input_comp_name + " CONTENT=" + full_bath);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtHalfBath" + input_comp_name + " CONTENT=" + half_bath);
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboBasement" + input_comp_name + " CONTENT=%330");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtBasementFinished" + input_comp_name + " CONTENT=0.00");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboGarageCarport" + input_comp_name + " CONTENT=%156");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtPoolSpaFirplace" + input_comp_name + " CONTENT=na");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtAmenities" + input_comp_name + " CONTENT=na");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboSuperior" + input_comp_name + " CONTENT=%496");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboSaleType" + input_comp_name + " CONTENT=%445");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtMLS" + input_comp_name + "");
                    macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*TxtMulSales1");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*txtMLS" + input_comp_name + " CONTENT=" + mlsnum);
                    //   macro.AppendLine(@"'TAG POS=0 TYPE=TEXTAREA FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*TxtMulSales1 CONTENT=EASY<SP>LIVING<SP>ALL<SP>ON<SP>ONE<SP>LEVEL!<SP>LAWN<SP>CARE<SP>&<SP>SNOW<SP>REMOVAL<SP>INCLUDED!<SP>LIVING<SP>ROOM<SP>W/<SP>VAULTED<SP>CEILING,<SP>CHANDELIER<SP>FOR<SP>DINING<SP>TABLE,<SP>CORNER<SP>GAS<SP>FIREPLACE<SP>W/CERAMIC<SP>TILE<SP>SURROUND<SP>&<SP>NICHE<SP>ABOVE<SP>FOR<SP>FLAT-SCREEN<SP>TV!<SP>9'<SP>CEILINGS!<SP>MASTER<SP>SUITE<SP>W/PRIVATE<SP>BATH<SP>&<SP>WALK-IN<SP>CLOSET.<SP>LARGE<SP>EAT-IN<SP>KITCHEN<SP>W/PANTRY<SP>&<SP>WINDOW<SP>SEAT!<SP>MUDROOM<SP>OFF<SP>GARAGE<SP>WITH<SP>LAUNDRY.<SP>2.5<SP>CAR<SP>GARAGE<SP>W/PLENTY<SP>OF<SP>STORAGE<SP>SPACE.<SP>CUSTOM<SP>BLINDS<SP>STAY!<SP>HUGE<SP>BACKYARD!<SP>");

                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*cboViewFactor1" + input_comp_name + "  CONTENT=%485");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*PropertyStyle" + input_comp_name + "  CONTENT=%93");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*PropertyType" + input_comp_name + "  CONTENT=%764");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:Bpo.aspx?control=* ATTR=NAME:*LeaseHold" + input_comp_name + "  CONTENT=%415");

                    //string macroCode = macro.ToString();
                    status = iim2.iimPlayCode(macro.ToString(), 60);

                    s = iim.iimPlayCode(macro12.ToString());
                    htmlCode = iim.iimGetLastExtract();
                    if (subjectAttachedRadioButton.Checked)
                    {
                        m = new AttachedListing(htmlCode);
                    }
                    else if (subjectDetachedradioButton.Checked)
                    {
                        m = new DetachedListing(htmlCode);
                    }
                    else
                    {
                        m = new MLSListing(htmlCode);
                    }

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    BPOFulfillment bpoform = new BPOFulfillment(m);

                    fieldList.Add("filepath", SubjectFilePath);



                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);


                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);


                }

                #endregion


                #region valuationPartners
                if (currentUrl.ToLower().Contains("valuationpartners"))
                {

                    StringBuilder macro = new StringBuilder();


                    if (!input_comp_name.Contains("LIST"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "ADDRSTREET CONTENT=" + full_street_address.Replace(" ", "<SP>"));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "ADDRCITY CONTENT=" + city.Replace(" ", "<SP>"));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "ADDRZIP CONTENT=" + zip);

                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "SALEDT CONTENT=" + closed_date);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "SALESPRICE CONTENT=" + sold_price);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "SRCFUNDS CONTENT=" + m.FinancingMlsString.Replace(" ", "<SP>"));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "ORIGPRICE CONTENT=" + orig_list_price);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "LISTPRICE CONTENT=" + current_list_price);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=NAME:txt" + input_comp_name + "ORIGLISTINGDT CONTENT=" + m.ListDateString);


                    }
                    else
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "STREET CONTENT=" + full_street_address.Replace(" ", "<SP>"));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "CITY CONTENT=" + city.Replace(" ", "<SP>"));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "ZIP CONTENT=" + zip);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "ORGPRICE CONTENT=" + orig_list_price);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "LISTINGPRICE CONTENT=" + current_list_price);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=NAME:txt" + input_comp_name + "ORGLISTINGDT CONTENT=" + m.ListDateString);
                    }



                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "DATA CONTENT=MLS");

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "PROXIMITY CONTENT=" + m.ProximityToSubject);




                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "MARKETDAYS CONTENT=" + dom);

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "SCONCESSION CONTENT=" + m.PointsMlsString);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "DISTRESSEDSALE CONTENT=" + m.DistressedSaleYesNo());
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "HOAASSESSMENT CONTENT=" + m.);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "LIVINGSQFT CONTENT=" + mls_gla);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "AGEYRS CONTENT=" + age);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "LOTSIZE CONTENT=" + m.Lotsize);
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "SITE");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "ROOMTOT CONTENT=" + room_count);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "BEDROOMS CONTENT=" + bedrooms);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "BATHFULL CONTENT=" + full_bath);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "BATHHALF CONTENT=" + half_bath);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "DESIGNSTYLE CONTENT=" + type.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "VIEW CONTENT=Residential");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "VIEWCOMPARISON CONTENT=Equal");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "NUMUNITS CONTENT=1");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "BASEMENTTYPE CONTENT=" + basement.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "BASEMENT CONTENT=" + m.BasementGLA());
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "BASEMENTFIN CONTENT=" + m.BasementFinishedPercentage());
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "GARAGE CONTENT=" + m.GarageType());
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "GARAGENUMCARS CONTENT=" + m.NumberGarageStalls());
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "POOL CONTENT=N/A");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:txt" + input_comp_name + "CONDITION CONTENT=AVG");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:ddl" + input_comp_name + "MOSTCOMPARABLE CONTENT=%Equal");


                    string b = "Similar size, age, style, features, condition and neighborhood as subject. ".Replace(" ", "<SP>");
                    macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:form1 ATTR=NAME:txtCOMMENTSALESCOMPARE CONTENT=" + b);
                    macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:form1 ATTR=NAME:txtCOMMENTSALESCOMPARE2 CONTENT=" + b);
                    macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:form1 ATTR=NAME:txtCOMMENTSALESCOMPARE3 CONTENT=" + b);
                    macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:form1 ATTR=NAME:txtANALYSISLISTINGDATA CONTENT=" + b);
                    macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:form1 ATTR=NAME:txtANALYSISLISTINGDATA2 CONTENT=" + b);
                    macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:form1 ATTR=NAME:txtANALYSISLISTINGDATA3 CONTENT=" + b);

                    macroCode = macro.ToString();

                    status = iim2.iimPlayCode(macroCode, 30);



                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);

                }


                #endregion

                #region m2m

                if (currentUrl.ToLower().Contains(@"ordereditwizard/"))
                {
                    #region code

                    s = iim.iimPlayCode(macro12.ToString());
                    htmlCode = iim.iimGetLastExtract();
                    if (subjectAttachedRadioButton.Checked)
                    {
                        m = new AttachedListing(htmlCode);
                    }
                    else if (subjectDetachedradioButton.Checked)
                    {
                        m = new DetachedListing(htmlCode);
                    }
                    else
                    {
                        m = new MLSListing(htmlCode);
                    }

                    m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    m.DateOfLastPriceChange = lastPriceChangeDate;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    M2MStandard bpoform = new M2MStandard(m);

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion

                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                }

                #endregion

                #region eval
                if (currentUrl.ToLower().Contains("evalonline"))
                {




                    StringBuilder macro = new StringBuilder();

                    if (input_comp_name == "COMPARABLE1")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Competitive<SP>Listing<SP>1");
                    }
                    if (input_comp_name == "COMPARABLE2")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Competitive<SP>Listing<SP>2");
                    }
                    if (input_comp_name == "COMPARABLE3")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Competitive<SP>Listing<SP>3");
                    }

                    if (input_comp_name == "RECENT_SALE1")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Comparable<SP>Sales<SP>1");
                    }

                    if (input_comp_name == "RECENT_SALE2")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Comparable<SP>Sales<SP>2");
                    }

                    if (input_comp_name == "RECENT_SALE3")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Comparable<SP>Sales<SP>3");
                    }

                    macro.AppendLine(@"WAIT SECONDS=2");



                    //macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Competitive<SP>Listing<SP>1");
                    //macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Competitive<SP>Listing<SP>2");
                    //macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:Competitive<SP>Listing<SP>3");

                    // C# snippet generated by iMacros Editor.
                    // See http://wiki.imacros.net/Web_Scripting for details on how to use the iMacros Scripting Interface.

                    // iMacros.AppClass iim = new iMacros.AppClass();
                    // iMacros.Status status = iim.iimOpen("", true, timeout);
                    //StringBuilder macro = new StringBuilder();


                    //
                    //Common Fields w/ matching var names
                    //

                    //lot size sf
                    if (mls_lot_size == "")
                        mls_lot_size = "0";

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10634 CONTENT=" + (Convert.ToDecimal(mls_lot_size) * 43560).ToString());
                    //macro.AppendLine(@"WAIT SECONDS=3");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10636 CONTENT=" + mls_lot_size);

                    //garage type and spaces
                    //garage type 1 = attached, 4 = detached
                    if (mls_garage_type.ToLower().Contains("att"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10630 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10632 CONTENT=" + mls_garage_spaces);
                    }
                    else if (mls_garage_type.ToLower().Contains("det"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10630 CONTENT=%4");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10632 CONTENT=" + mls_garage_spaces);
                    }
                    else
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10630 CONTENT=%3");
                    }


                    //att=1 or det=3
                    if (subjectAttachedRadioButton.Checked)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10604 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10610 CONTENT=%7");
                    }
                    else
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10604 CONTENT=%3");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10610 CONTENT=%1");

                    }
                    //mlstype
                    switch (type)
                    {
                        case "1 Story":
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%1");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Ranch");
                            if (fullBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + mls_gla);
                            }
                            else if (partialBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + (Convert.ToInt16(mls_gla.Replace(",", "")) / 2));
                            }
                            else
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=0");
                            }
                            break;
                        case "1.5 Story":
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%3");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Conventional");
                            if (fullBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + (Convert.ToInt16(mls_gla.Replace(",", "")) / 2));
                            }
                            else if (partialBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + (Convert.ToInt16(mls_gla.Replace(",", "")) / 4));
                            }
                            else
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=0");
                            }

                            break;
                        case "2 Stories":
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%4");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Contemporary");
                            if (fullBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + (Convert.ToInt16(mls_gla.Replace(",", "")) / 2));
                            }
                            else if (partialBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + (Convert.ToInt16(mls_gla.Replace(",", "")) / 4));
                            }
                            else
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=0");
                            }
                            break;
                        case "Raised Ranch":
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%9");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Raised<SP>Ranch");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%1");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Ranch");
                            if (fullBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + mls_gla);
                            }
                            else if (partialBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + (Convert.ToInt16(mls_gla.Replace(",", "")) / 2));
                            }
                            else
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=0");
                            }
                            break;

                        case "Split Level":
                        case @"Split Level w/Sub":
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%25");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Split<SP>Level");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%4");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Contemporary");
                            if (fullBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + (Convert.ToInt16(mls_gla.Replace(",", "")) / 2));
                            }
                            else if (partialBasement)
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=" + (Convert.ToInt16(mls_gla.Replace(",", "")) / 4));
                            }
                            else
                            {
                                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=0");
                            }
                            break;
                    }

                    //basement
                    if (basement.ToLower().Contains("full"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10624 CONTENT=%3");
                        if (finished_basement)
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10626 CONTENT=%5");
                        }
                        else
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10626 CONTENT=%1");
                        }
                    }
                    else
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10624 CONTENT=%2");
                    }
                    //patio, porch, deck, 3=no
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10640 CONTENT=%3");
                    //fireplace, 3=no
                    if (hasFireplace)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10642 CONTENT=%1");
                    }
                    else
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10642 CONTENT=%3");
                    }

                    string content = "%Other<SP>(Explain)";
                    //exterior walls
                    if (exteriorDetails.Contains("Vinyl Siding"))
                    {
                        content = "%Vinyl<SP>Siding";
                    }
                    else if (exteriorDetails.Contains("Brick"))
                    {
                        content = "%Brick";
                    }
                    else if (exteriorDetails.Contains("Cedar"))
                    {
                        content = "%Wood";
                    }

                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10648 CONTENT=" + content);
                    if (content == "Other<SP>(Explain)")
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10822 CONTENT=" + exteriorDetails);
                    }

                    //
                    //Distance logic
                    //
                    string proximity = this.Get_Distance(subjectFullAddressTextbox.Text, address);


                    if (Convert.ToDouble(proximity) > 4 && Convert.ToDouble(proximity) <= 7)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%11");

                    }
                    else if (Convert.ToDouble(proximity) > 0 && Convert.ToDouble(proximity) <= .125)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%3");
                    }
                    else if (Convert.ToDouble(proximity) > .125 && Convert.ToDouble(proximity) <= .3)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%4");
                    }
                    else if (Convert.ToDouble(proximity) > .3 && Convert.ToDouble(proximity) <= .425)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%5");
                    }
                    else if (Convert.ToDouble(proximity) > .425 && Convert.ToDouble(proximity) <= .55)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%6");
                    }
                    else if (Convert.ToDouble(proximity) > .55 && Convert.ToDouble(proximity) <= .675)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%7");
                    }
                    else if (Convert.ToDouble(proximity) > .675 && Convert.ToDouble(proximity) <= .8)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%7");
                    }
                    else if (Convert.ToDouble(proximity) > .8 && Convert.ToDouble(proximity) <= 2)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%7");
                    }
                    else if (Convert.ToDouble(proximity) > 2 && Convert.ToDouble(proximity) <= 4)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%11");
                    }
                    //


                    if (input_comp_name.Contains("SALE"))
                    {
                        #region salecomps



                        if (input_comp_name == "RECENT_SALE2")
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10916 CONTENT=%3");
                        }
                        else
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10916 CONTENT=%1");
                        }
                        //financing
                        switch (financing)
                        {
                            case "FHA":
                                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10906 CONTENT=%1");
                                break;
                            case "VA":
                                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10906 CONTENT=%1");
                                break;
                            case "Conventional":
                                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10906 CONTENT=%3");
                                break;
                            case "Cash":
                                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10906 CONTENT=%5");
                                break;

                        }
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10906 CONTENT=%3");
                        //condition
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10900 CONTENT=%3");

                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10918 CONTENT=Equal");

                        if (sale_type == "Normal")
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10905 CONTENT=%Arms<SP>Length");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10908 CONTENT=%11");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10910 CONTENT=%3");
                        }
                        else if (sale_type == "Short")
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10905 CONTENT=%Short<SP>Sale");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10908 CONTENT=%11");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10910 CONTENT=%1");
                        }
                        else if (sale_type == "REO")
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10905 CONTENT=%REO");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10908 CONTENT=%13");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10910 CONTENT=%1");
                        }






                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10500 CONTENT=" + full_street_address.Replace(" ", "<SP>"));

                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10504 CONTENT=" + city.Replace(" ", "<SP>"));

                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10506 CONTENT=%IL");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10508 CONTENT=" + zip);
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10510 CONTENT=%4");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10512 CONTENT=%3");

                        //prox to subject, 6=within 1 mi
                        // macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%6");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10516 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10518 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10518 CONTENT=%3");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10520 CONTENT=na");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10522 CONTENT=na");


                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10600 CONTENT=" + age.ToString());


                        //number of units = 1
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10602 CONTENT=1");



                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10604 CONTENT=%2");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%25");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Split<SP>Level");
                        //legal land use = residential
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10608 CONTENT=%1");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10610 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10611 CONTENT=%Framed");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10612 CONTENT=" + mls_gla);
                        //GLA source 1=tax records
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10614 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10616 CONTENT=" + room_count);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10618 CONTENT=" + bedrooms);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10620 CONTENT=" + full_bath);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10622 CONTENT=" + half_bath);
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10624 CONTENT=%3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10626 CONTENT=%5");



                        // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10634");
                        // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10636 CONTENT=.2");
                        // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:formInput ATTR=ID:submit&&VALUE:SAVE<SP>&<SP>VALIDATE");
                        macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10636 CONTENT=.2");

                        //pool = no
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10638 CONTENT=%3");

                        // macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10640 CONTENT=%3");


                        //guest house = no
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10644 CONTENT=%3");


                        //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10646");

                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10700");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10702");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10704");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10706");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10708 CONTENT=" + orig_list_price);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10710 CONTENT=" + sold_price);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10712 CONTENT=" + closed_date);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10714 CONTENT=" + dom);
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10716");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10717 CONTENT=%1");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10718");
                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10720 CONTENT=See<SP>subject<SP>comments.");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10722 CONTENT=" + subjectRentTextbox.Text);

                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10902 CONTENT=$--Choose<SP>Option--");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10904");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10905 CONTENT=%Arms<SP>Length");
                        //macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=ID:formInput ATTR=TXT:--Choose<SP>Option--FHA<SP>VA<SP>Conventional<SP>Owner-financed<SP>Cash<SP>Other<SP>-Explain<SP>in<SP>comm*");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10906 CONTENT=%1");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10908 CONTENT=%11");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10910 CONTENT=%3");

                        //compare to subject = equal
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10914 CONTENT=%2");


                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10916 CONTENT=%1");


                        //view = residential
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10921 CONTENT=%10");

                        //view effect = equal
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10922 CONTENT=%3");
                        //macro.AppendLine(@"TAG POS=19 TYPE=TD FORM=ID:formInput ATTR=CLASS:formLabelAlignRight");
                        //personally inspect property; 3 = no
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10926 CONTENT=%3");
                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10918 CONTENT=Equal");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:formInput ATTR=ID:submit&&VALUE:SAVE<SP>&<SP>VALIDATE");
                        macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:InputForm");
                        macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:output");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:formInput ATTR=ID:submit&&VALUE:SAVE<SP>&<SP>VALIDATE");
                        ////macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:InputForm");
                        ////macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:output");
                        ////string macroCode = macro.ToString();
                        ////status = iim.iimPlayCode(macroCode, timeout);
                        #endregion
                    }
                    else
                    {

                        #region listcomps
                        if (input_comp_name == "COMPARABLE2")
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10816 CONTENT=%3");
                        }
                        else
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10816 CONTENT=%1");
                        }
                        if (additionalSalesInfo.Contains("Short"))
                        {
                            //  macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10805 CONTENT=%Short<SP>Sale");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10808 CONTENT=%11");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10810 CONTENT=%1");

                        }
                        else if (additionalSalesInfo.Contains("REO"))
                        {
                            // macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10805 CONTENT=%REO");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10808 CONTENT=%13");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10810 CONTENT=%1");
                        }
                        else
                        {
                            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10805 CONTENT=%Arms<SP>Length");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10808 CONTENT=%11");
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10810 CONTENT=%3");

                        }

                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10500 CONTENT=" + full_street_address.Replace(" ", "<SP>"));

                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10504 CONTENT=" + city.Replace(" ", "<SP>"));

                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10506 CONTENT=%IL");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10508 CONTENT=" + zip);

                        //pud = no
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10510 CONTENT=%4");

                        //corner = no
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10512 CONTENT=%3");


                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%1");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%5");

                        //location = same 
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10516 CONTENT=%1");


                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10518 CONTENT=%3");


                        //dev name  - sub name
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10520 CONTENT=na");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10522 CONTENT=na");


                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10600 CONTENT=" + age.ToString());
                        //units = 1
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10602 CONTENT=1");

                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10604 CONTENT=%3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%27");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%29");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%26");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10606 CONTENT=%25");
                        ////macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10607 CONTENT=%Split<SP>Level");

                        //legal land use = res
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10608 CONTENT=%1");

                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10610 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10611 CONTENT=%Framed");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10612 CONTENT=" + mls_gla);
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10614 CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10616 CONTENT=" + room_count);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10618 CONTENT=" + bedrooms);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10620 CONTENT=" + full_bath);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10622 CONTENT=" + half_bath);
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10624 CONTENT=%3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10626 CONTENT=%5");

                        //basement sf
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10628 CONTENT=");


                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10630 CONTENT=%1");
                        //macro.AppendLine(@"WAIT SECONDS=3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10632 CONTENT=2");
                        //macro.AppendLine(@"WAIT SECONDS=3");



                        //macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                        //macro.AppendLine(@"WAIT SECONDS=3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=B FORM=ID:formInput ATTR=TXT:Convert<SP>to<SP>Sqft");
                        //macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                        //macro.AppendLine(@"ONDIALOG POS=2 BUTTON=YES");
                        //macro.AppendLine(@"ONDIALOG POS=3 BUTTON=YES");
                        //macro.AppendLine(@"ONDIALOG POS=4 BUTTON=YES");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10636 CONTENT=.2");
                        //macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                        //macro.AppendLine(@"ONDIALOG POS=2 BUTTON=YES");
                        //macro.AppendLine(@"ONDIALOG POS=3 BUTTON=YES");
                        //macro.AppendLine(@"ONDIALOG POS=4 BUTTON=YES");

                        //pool = no
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10638 CONTENT=%3");
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10614 CONTENT=%8");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10642 CONTENT=%1");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10642 CONTENT=%3");
                        //guest house = no
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10644 CONTENT=%3");

                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10518 CONTENT=%3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10514 CONTENT=%7");

                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10646");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10648 CONTENT=%Brick");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10648 CONTENT=%Vinyl<SP>Siding");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10750");
                        //agent name

                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10752 CONTENT=" + listing_Agent.Replace(" ", "<SP>"));

                        //agency address
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10754 CONTENT=" + broker.Replace(" ", "<SP>"));


                        //agency phone
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10756 CONTENT=" + phone.Replace(" ", "<SP>"));

                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10758 CONTENT=" + list_date);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10760 CONTENT=" + current_list_price);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10762 CONTENT=" + list_date);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10764 CONTENT=" + current_list_price);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10766 CONTENT=" + dom);
                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10768 CONTENT=" + mlsnum);
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10769 CONTENT=%0");
                        //offered conssessions
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10770 CONTENT=0");
                        //rent different field number for sale comps
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10772 CONTENT=" + subjectRentTextbox.Text);

                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10767 CONTENT=See<SP>subject<SP>comments.");

                        //condition = avg
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10800 CONTENT=%3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10802 CONTENT=$--Choose<SP>Option--");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10804");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10805 CONTENT=$--Choose<SP>Option--");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10806 CONTENT=$--Choose<SP>Option--");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10808 CONTENT=%11");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10810 CONTENT=%3");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10812");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10810 CONTENT=%1");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10812");
                        //compare to subject = equal
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10814 CONTENT=%2");

                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10816 CONTENT=$--Choose<SP>Option--");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10808 CONTENT=%13");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10810 CONTENT=%1");
                        //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:formInput ATTR=ID:fieldid_10812");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10814 CONTENT=%2");
                        //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10816 CONTENT=%3");

                        //personally inspect = no
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10818 CONTENT=%3");

                        //lsiting view = res
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10823 CONTENT=%10");

                        //view affect, same
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:formInput ATTR=NAME:fieldid_10824 CONTENT=%3");

                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:formInput ATTR=NAME:fieldid_10820 CONTENT=Equal");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:formInput ATTR=ID:submit&&VALUE:SAVE<SP>&<SP>VALIDATE");
                        macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:InputForm");
                        macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:output");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:formInput ATTR=ID:submit&&VALUE:SAVE<SP>&<SP>VALIDATE");
                        macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:InputForm");
                        macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:output");

                        #endregion

                    }


                    macroCode = macro.ToString();
                    status = iim2.iimPlayCode(macroCode, 30);



                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                    //MessageBox.Show("Comp:" + input_comp_name + " finished.");

                }
                #endregion

                #region fill usres

                if (currentUrl.ToLower().Contains(@"mp_formmsr"))
                {

                    //stop the other res.net url match
                    currentUrl = "mp_formmsr";


                    #region code
                    try
                    {
                        m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    }
                    catch
                    {
                        m.proximityToSubject = 0;
                    }

                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    Resnet bpoform = new Resnet(m);


                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion

                    sale_or_list_flag = new CultureInfo("en-US").TextInfo.ToTitleCase(sale_or_list_flag);
                    bpoform.MmrFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);


                }

                if (currentUrl.ToLower().Contains(@"valuation.res.net/providerresponse/"))
                {


                    #region code
                    try
                    {
                        m.proximityToSubject = Convert.ToDouble(Get_Distance(m.mlsHtmlFields["address"].value, this.SubjectFullAddress));
                    }
                    catch
                    {
                        m.proximityToSubject = 0;
                    }
                 
                    m.DateOfLastPriceChange = lastPriceChangeDate;
                    m.NumberOfPriceChanges = count;

                    Dictionary<string, string> fieldList = new Dictionary<string, string>();
                    Resnet bpoform = new Resnet(m);
                   

                    fieldList.Add("filepath", SubjectFilePath);

                    #endregion

                    sale_or_list_flag = new CultureInfo("en-US").TextInfo.ToTitleCase(sale_or_list_flag);
                    bpoform.CompFill(iim2, sale_or_list_flag, input_comp_name, fieldList);
                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);


                }
                else if (currentUrl.ToLower().Contains("usres") || currentUrl.ToLower().Contains("res.net"))
                {

                    StringBuilder macro = new StringBuilder();
                    macro.AppendLine(@"SET !ERRORIGNORE YES");
                    macro.AppendLine(@"SET !TIMEOUT_STEP 0");

                    if (input_comp_name.Contains("L"))
                    {

                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:IMAGE FORM=NAME:InputForm ATTR=NAME:btnSave");
                        //macro.AppendLine(@"WAIT SECONDS=5");
                        macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:InputForm ATTR=SRC:http://valuations.usres.com/BpoImages/bpo_pg2_btn.gif");
                        macro.AppendLine(@"TAG POS=2 TYPE=IMG FORM=NAME:InputForm ATTR=SRC:https://agents.res.net/ResImages/bpo_pg2_btn.gif");
                        //macro.AppendLine(@"WAIT SECONDS=5");



                    }


                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abAddr" + input_comp_name + " CONTENT=" + full_street_address.Replace(" ", "<SP>"));

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abCity" + input_comp_name + " CONTENT=" + city.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abState" + input_comp_name + " CONTENT=%IL");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abZip" + input_comp_name + " CONTENT=" + zip);

                    //
                    //Distance logic
                    //
                    string proximity = this.Get_Distance(subjectFullAddressTextbox.Text, address);

                    if (Convert.ToDouble(proximity) > 4 && Convert.ToDouble(proximity) <= 7)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%2Mile");
                    }
                    else if (Convert.ToDouble(proximity) > 0 && Convert.ToDouble(proximity) <= .125)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%1Blk");
                    }
                    else if (Convert.ToDouble(proximity) > .125 && Convert.ToDouble(proximity) <= .3)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%2Blk");
                    }
                    else if (Convert.ToDouble(proximity) > .3 && Convert.ToDouble(proximity) <= .425)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%3Blk");
                    }
                    else if (Convert.ToDouble(proximity) > .425 && Convert.ToDouble(proximity) <= .55)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%4Blk");
                    }
                    else if (Convert.ToDouble(proximity) > .55 && Convert.ToDouble(proximity) <= .675)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%5Blk");
                    }
                    else if (Convert.ToDouble(proximity) > .675 && Convert.ToDouble(proximity) <= .8)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%6Blk");
                    }
                    else if (Convert.ToDouble(proximity) > .8 && Convert.ToDouble(proximity) <= .9125)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%7Blk");
                    }
                    else if (Convert.ToDouble(proximity) > .9125 && Convert.ToDouble(proximity) <= 4)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abProx" + input_comp_name + " CONTENT=%1Mile");
                    }
                    //

                    if (input_comp_name.Contains("S"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSp" + input_comp_name + " CONTENT=" + sold_price.Replace("$", "").Replace(",", ""));
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abStype1 CONTENT=%t");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSaleDt" + input_comp_name + " CONTENT=" + closed_date);
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abLp" + input_comp_name + " CONTENT=" + current_list_price.Replace("$", "").Replace(",", ""));
                    }

                    if (input_comp_name.Contains("L"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abLpOrig" + input_comp_name + " CONTENT=" + orig_list_price.Replace("$", "").Replace(",", ""));
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSp" + input_comp_name + " CONTENT=" + current_list_price.Replace("$", "").Replace(",", ""));
                    }



                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abLpRed" + input_comp_name + " CONTENT=" + count);
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abSrce" + input_comp_name + " CONTENT=%MLS");

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abDom" + input_comp_name + " CONTENT=" + dom);
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abLoc" + input_comp_name + " CONTENT=%Suburban");

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abRight" + input_comp_name + " CONTENT=Fee<SP>Simple");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSite" + input_comp_name + " CONTENT=" + mls_lot_size);

                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abUnits" + input_comp_name + " CONTENT=%1");

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abView" + input_comp_name + " CONTENT=Neighborhood");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abDesign" + input_comp_name + " CONTENT=%Avg");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abAge" + input_comp_name + " CONTENT=" + year_built);
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abCond" + input_comp_name + " CONTENT=%Avg");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abRooms" + input_comp_name + " CONTENT=" + room_count);

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBeds" + input_comp_name + " CONTENT=" + bedrooms);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBaths" + input_comp_name + " CONTENT=" + full_bath + "." + half_bath);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abGla" + input_comp_name + " CONTENT=" + mls_gla);

                    //Basement and finished rooms below grade
                    if (finished_basement)
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBsmt" + input_comp_name + " CONTENT=" + basement.Replace(" ", "<SP>") + "<SP>/<SP>Yes");
                    }
                    else
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abBsmt" + input_comp_name + " CONTENT=" + basement.Replace(" ", "<SP>") + "<SP>/<SP>No");
                    }

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abHeat" + input_comp_name + " CONTENT=Gas/Central");


                    if (mls_garage_type.ToLower().Contains("att"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abGarage" + input_comp_name + " CONTENT=%" + mls_garage_spaces + "CA");
                    }
                    else if (mls_garage_type.ToLower().Contains("det"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abGarage" + input_comp_name + " CONTENT=%" + mls_garage_spaces + "CD");
                    }
                    else
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:InputForm ATTR=NAME:abGarage CONTENT=%N");
                    }


                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:IMAGE FORM=NAME:InputForm ATTR=NAME:btnSave");
                    macro.AppendLine(@"WAIT SECONDS=5");

                    if (input_comp_name.Contains("L"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:IMAGE FORM=NAME:InputForm ATTR=NAME:btnSave");
                        macro.AppendLine(@"TAG POS=2 TYPE=IMG FORM=NAME:InputForm ATTR=SRC:http://valuations.usres.com/BpoImages/bpo_pg1_btn.gif");
                        macro.AppendLine(@"TAG POS=2 TYPE=IMG FORM=NAME:InputForm ATTR=SRC:https://agents.res.net/ResImages/bpo_pg1_btn.gif");



                    }

                    macroCode = macro.ToString();
                    status = iim2.iimPlayCode(macroCode, 60);


                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                    //MessageBox.Show(input_comp_name);

                    //break;

                }

                #endregion

                StringBuilder macro3 = new StringBuilder();
                string macroCode3 = "";
                #region fill NVS

                if (streetnumTextBox.Text == "nvs")
                {

                    macro3.AppendLine(@"SET !ERRORIGNORE YES");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "DistToSubj CONTENT=" + Get_Distance(subjectFullAddressTextbox.Text, address));
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    //
                    //TAG POS=1 TYPE=TD FORM=NAME:bpoform ATTR=TXT:?<SP>Urban<SP>Suburban<SP>Rural
                    // 3=rural
                    //

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Location_id CONTENT=%2");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Condition CONTENT=%Average");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Rooms CONTENT=" + room_count);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Beds CONTENT=" + bedrooms);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "FullBaths CONTENT=" + full_bath);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "HalfBaths CONTENT=" + half_bath);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    if (mls_gla == "")
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "SqFeet CONTENT=" + subjectAboveGlaTextbox.Text.Replace(",", ""));
                    }
                    else
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "SqFeet CONTENT=" + mls_gla.Replace(",", ""));
                    }

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "LotSize CONTENT=" + mls_lot_size);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    //
                    //specific logic to form
                    //
                    if (year_built == "")
                    {
                        year_built = "1900";
                    }
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "YearBuilt CONTENT=" + year_built);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "ListPrice CONTENT=" + current_list_price.Replace("$", "").Replace(",", ""));
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "DOM CONTENT=" + dom);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    //
                    //need select list logic
                    if (type.ToLower().Contains("1 story"))
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Style CONTENT=%Ranch");
                    }
                    else
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Style CONTENT=%Contemporary");
                    }
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    //
                    //

                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Construction_type CONTENT=%1");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

                    if (number_of_firplaces == "" | number_of_firplaces == "0")
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Fireplace CONTENT=%0");
                    }
                    else
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Fireplace CONTENT=%1");
                    }


                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Pool CONTENT=%N");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Carport CONTENT=%0");

                    if (parking.Contains("Gar"))
                    {
                        if (mls_garage_type.Contains("Att"))
                        {
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Garage CONTENT=%A" + mls_garage_spaces);
                        }
                        else
                        {
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Garage CONTENT=%D" + mls_garage_spaces);
                        }
                    }
                    else
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Garage CONTENT=%N");
                    }

                    //macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Garage CONTENT=%A2");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Porch CONTENT=%N");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Patio CONTENT=%N");


                    //
                    //Addition data for sales
                    //
                    if (input_comp_name.Contains("Sale"))
                    {

                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Landscaping CONTENT=%Average");
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "SaleDate CONTENT=" + closed_date);
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "SalePrice CONTENT=" + sold_price.Replace("$", "").Replace(",", ""));
                        //macro3.AppendLine(@"");

                    }


                    //crawl
                    if (subjectBasementDetailsTextbox.Text.Contains("Crawl"))
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%C");
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Basement CONTENT=%C");

                    }

                    //none
                    if (subjectBasementTypeTextbox.Text == "None" | subjectBasementTypeTextbox.Text == "")   //.Text.Contains("None"))
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%N");
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Basement CONTENT=%N");
                    }


                    if (subjectBasementTypeTextbox.Text.Contains("Full"))
                    {
                        if (subjectBasementDetailsTextbox.Text.Contains("Finished"))
                        {
                            //Full finished                    
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%F");
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Basement CONTENT=%F");
                        }
                        else
                        {
                            //full unfinished  
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%Y");
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Basement CONTENT=%Y");
                        }
                    }



                    if (subjectBasementTypeTextbox.Text.Contains("Partial"))
                    {
                        if (subjectBasementDetailsTextbox.Text.Contains("Finished"))
                        {
                            //partial finish
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%P");
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Basement CONTENT=%P");
                        }
                        else
                        {
                            //partial unfinished
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:Basement CONTENT=%X");
                            macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Basement CONTENT=%X");
                        }
                    }



                    macro3.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:bpoform ATTR=NAME:" + input_comp_name + "Explain CONTENT=Same<SP>size<SP>GLA<SP>and<SP>lot,<SP>same<SP>area<SP>type.");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:bpoform ATTR=VALUE:Go<SP>to<SP>Next<SP>Section<SP>-<SP>Save");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:bpoform ATTR=VALUE:Go<SP>to<SP>Next<SP>Section<SP>-<SP>Save");
                    macroCode3 = macro3.ToString();
                    status = iim2.iimPlayCode(macroCode3, 60);

                    //MessageBox.Show(status.ToString());







                    //
                    //TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:bpoform ATTR=VALUE:Go<SP>to<SP>Next<SP>Section<SP>-<SP>Save
                    //TAG POS=1 TYPE=TEXTAREA FORM=NAME:bpoform ATTR=NAME:CList1Explain CONTENT=Same<SP>size<SP>GLA<SP>and<SP>lot,<SP>same<SP>area<SP>type.
                    //ONDIALOG POS=1 BUTTON=NO

                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                    //return;

                }

                #endregion

                #region fill sls
                if (streetnumTextBox.Text == "sls")
                {


                    StringBuilder macro = new StringBuilder();
                    //macro.AppendLine(@"VERSION BUILD=8021952");
                    //macro.AppendLine(@"TAB T=1");
                    //macro.AppendLine(@"TAB CLOSEALLOTHERS");
                    //macro.AppendLine(@"URL GOTO=https://portal.nreis.com/VendorWorkflow/VendorNewForm.aspx?OrderID=1684158&SelectionID=1873443&MCID=17&ContactID=74134&SkipPipeline=Y&VendorNewForm=Y");

                    macro.AppendLine(@"SET !ERRORIGNORE YES");
                    macro.AppendLine(@"TAG POS=3 TYPE=TD FORM=ID:form1 ATTR=CLASS:BpoGrid");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompAdr" + input_comp_name + " CONTENT=" + full_street_address.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompCty" + input_comp_name + " CONTENT=" + city.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlCompSta" + input_comp_name + " CONTENT=%IL");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompZip" + input_comp_name + " CONTENT=" + zip);

                    //wait for proximity to load
                    macro.AppendLine(@"WAIT SECONDS=3");


                    // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompProximity1");
                    macro.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompZip1 CONTENT=60098");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompProximity1");
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompSalePrice2 CONTENT=565656");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompOrigPrice" + input_comp_name + " CONTENT=" + orig_list_price.Replace("$", "").Replace(",", ""));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompSalePrice" + input_comp_name + " CONTENT=" + sold_price.Replace("$", "").Replace(",", ""));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompListPrice" + input_comp_name + " CONTENT=" + current_list_price.Replace("$", "").Replace(",", ""));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompDataSrc" + input_comp_name + " CONTENT=MLS");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompSaleDate" + input_comp_name + " CONTENT=" + closed_date);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompListDate" + input_comp_name + " CONTENT=" + list_date);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompConcessions" + input_comp_name + " CONTENT=NA");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlCompLoc" + input_comp_name + " CONTENT=%2");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlCompLeaseFee" + input_comp_name + " CONTENT=%1");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompSize" + input_comp_name + " CONTENT=" + mls_lot_size.Replace(",", ""));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompView" + input_comp_name + " CONTENT=Residential");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompDesign" + input_comp_name + " CONTENT=" + type.Replace(" ", "<SP>") + "/Avg");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlCompQuality" + input_comp_name + " CONTENT=%3");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompAge" + input_comp_name + " CONTENT=" + age.ToString());
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlCompCond" + input_comp_name + " CONTENT=%3");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompRmCountTotal" + input_comp_name + " CONTENT=" + room_count);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompRmCountBdms" + input_comp_name + " CONTENT=" + bedrooms);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompRmCountBaths" + input_comp_name + " CONTENT=" + full_bath);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompRmCountHBath" + input_comp_name + " CONTENT=" + half_bath);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompLivArea" + input_comp_name + " CONTENT=" + mls_gla.Replace(",", ""));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompBaseRms" + input_comp_name + "");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompPercFin" + input_comp_name + "");
                    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form1 ATTR=ID:BPOForm1_ddlCompUtility" + input_comp_name + " CONTENT=%3");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompHeatAC" + input_comp_name + " CONTENT=GASFA/Central");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompEnergy" + input_comp_name + " CONTENT=" + "NA");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompGarage" + input_comp_name + " CONTENT=" + mls_garage_spaces + mls_garage_type);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompDeckEtc" + input_comp_name + " CONTENT=" + "NA");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompPoolEtc" + input_comp_name + " CONTENT=" + "NA");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompOther" + input_comp_name + " CONTENT=" + "NA");


                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form1 ATTR=ID:BPOForm1_txtCompConcessionsAdj1");
                    macroCode = macro.ToString();
                    status = iim2.iimPlayCode(macroCode, 60);

                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);
                    //return;

                }
                #endregion

                #region imort dialog control
                if (currentUrl.ToLower().Contains("propertysmart"))
                {
                    macro3.AppendLine(@"SET !ERRORIGNORE YES");
                    macro3.AppendLine(@"SET !TIMEOUT_STEP 0");
                    macro3.AppendLine(@"SET !ERRORIGNORE YES");
                    macro3.AppendLine(@"FRAME NAME=_MAIN");
                    macro3.AppendLine(@"ONWEBPAGEDIALOG KEYS={CLOSE}");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                #endregion

                    #region comp input
                    //macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/APN CONTENT=" + subjectpin_textbox.Text);
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/SUBJECT_PROPERTY/School_District CONTENT=school");

                    // macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:Form1 ATTR=NAME:PS_FORM/RECENT_SALE1/Datasource CONTENT=%MLS");


                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Street_Address1_Number CONTENT=" + street_number);
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Street_Address1_Text CONTENT=" + street_name.Replace(" ", "<SP>"));
                    //
                    //postfix fixing
                    //
                    if (street_postfix.ToUpper() == "TER")
                    {
                        street_postfix = "TERRACE";
                    }

                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Street_Address1_Ext CONTENT=%" + street_postfix);
                 
                
                 


                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/City CONTENT=" + city.Replace(" ", "<SP>"));
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/State CONTENT=IL");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/State CONTENT=%IL");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Zip_Code CONTENT=" + zip);

                    macro3.AppendLine(@"WAIT SECONDS=1");
                    macro3.AppendLine(@"DS cmd=KEY X={{!TAGX}} Y={{!TAGY}}  CONTENT=1{BACKSPACE}");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Unit_Number");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");


                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");

                    //fnma
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Miles CONTENT=" + this.Get_Distance(subjectFullAddressTextbox.Text, m.mlsHtmlFields["address"].value));
                    //
                    //fnma
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Location CONTENT=Suburban");
                    //
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Data_Source CONTENT=MLS");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Location CONTENT=Suburban");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Concessions  CONTENT=None");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Financing_Concessions CONTENT=None");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Utility CONTENT=Typical");

                    //single source updates Q1 2018
                    //singlesource update
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Street_Address1 CONTENT=" + full_street_address.Replace(" ", "<SP>"));
                    //Original_List_Date_Input
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Original_List_Date_Input CONTENT=" + list_date);
                    //Original_List_Price
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Original_List_Price CONTENT=" +  orig_list_price.Replace("$", "").Replace(",", ""));
                    //Last_Listing_Price
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Last_Listing_Price CONTENT=" + current_list_price.Replace("$", "").Replace(",", ""));
                    //DOM
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/DOM CONTENT=" + dom);
                    //GLA
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/GLA CONTENT=" + mls_gla.Replace(",", ""));
                    //Age
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Age CONTENT=" + age);
                    //Full_Baths
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Full_Baths CONTENT=" + full_bath);
                    //Half_Baths
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Half_Baths CONTENT=" + half_bath);
                    //View_Rating_Select
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/View_Rating_Select CONTENT=%N");
                    //View_Factor
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/View_Factor CONTENT=%Res");
                    //View_Comparison
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/View_Comparison CONTENT=Equal");
                    //Units
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Units CONTENT=1");
                    //Design_Attachments
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Design_Attachments CONTENT=%DT");
                    //Location_Rating
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Location_Rating CONTENT=%N");
                    //Location_Factor
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Location_Factor CONTENT=%REs");





                    if (SubjectDetached)
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Property_Type CONTENT=%SFR<SP>Detached");
                    }


                    if (string.IsNullOrWhiteSpace(mls_subdivision))
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Sub_Neighborhood CONTENT=Unk");
                    }
                    else
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Sub_Neighborhood CONTENT=" + mls_subdivision.Replace(" ", "<SP>"));
                    }





                    if (input_comp_name.Contains("SALE"))
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Num_Reductions CONTENT=" + count);
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/List_At_Sale CONTENT=" + current_list_price.Replace("$", "").Replace(",", ""));
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");






                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Sale_Date CONTENT=" + closed_date);
                        //fnma
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Sales_Date CONTENT=" + closed_date);
                        //
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");






                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        //fnma
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Sales_Price CONTENT=" + sold_price.Replace("$", "").Replace(",", ""));

                        //
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Sale_Price CONTENT=" + sold_price.Replace("$", "").Replace(",", ""));

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");






                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    }
                    else
                    {
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Datasource CONTENT=MLS");
                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Price_Reductions CONTENT=" + count);
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

                        macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/List_Price CONTENT=" + current_list_price.Replace("$", "").Replace(",", ""));



                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");


                        macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

                    }

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Days_On_Market CONTENT=" + dom);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");


                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");

                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Location CONTENT=%Suburban");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");


                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    //fnma
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Leasehold CONTENT=Fee<SP>Simple");
                    //
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Fee_Simple CONTENT=Fee<SP>Simple");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Fee_Simple CONTENT=%Fee<SP>Simple");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");


                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    //fnam
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Site CONTENT=" + mls_lot_size);
                    //
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Lot_Size CONTENT=" + mls_lot_size);
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");


                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/View CONTENT=Residential");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/View CONTENT=%Residential");
                    //fnma
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Design CONTENT=" + type.Replace(" ", "<SP>") + "/Average");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Construction CONTENT=Average");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Condition CONTENT=Average");
                    //
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Style_Design CONTENT=%Average");

                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Year_Built CONTENT=" + year_built);
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Condition CONTENT=%Average");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");


                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=4 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");


                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");






                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=5 BUTTON=NO");

                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Total_Rooms CONTENT=" + room_count);
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Bedrooms CONTENT=" + bedrooms);
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Bathrooms CONTENT=" + full_bath);
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Half_Bathrooms CONTENT=" + half_bath);

                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=3 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");

                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=OK CONTENT=");


                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=1 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Living_Square_Feet CONTENT=" + mls_gla.Replace(",", ""));

                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Heating_Cooling CONTENT=Gas<SP>FA/Central");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Heating_Cooling CONTENT=%FWA");

                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Basement CONTENT=%Yes");

                    if (basement.Contains("None"))
                    {


                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Basement CONTENT=%No");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Basement_Finished CONTENT=%No");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    }

                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Fireplace CONTENT=%Yes");
                    if (number_of_firplaces == "")
                    {
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Fireplace CONTENT=%No");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                    }

                    if (parking != "Garage")
                    {

                        //TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/RECENT_SALE1/Garage_Type CONTENT=%Attached
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                        macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Garage CONTENT=%NO");
                        macro3.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");

                    }

                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Garage_Carport CONTENT=" + subjectParkingTypeTextbox.Text.Replace(" ", "<SP>"));
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Design_Style CONTENT=" + type.Replace(" ", "<SP>"));

                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Fireplace_Adj EXTRACT=TXT");
                    macro3.AppendLine(@"DS cmd=CLICK X={{!TAGX}} Y={{!TAGY}}  ");


                    //ONDIALOG POS=1 BUTTON=NO
                    //TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/RECENT_SALE1/Design_Style CONTENT=2<SP>Story
                    //TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/RECENT_SALE1/Design_Style CONTENT=2<SP>Story<SP>Split
                    //ONDIALOG POS=1 BUTTON=NO


                    //TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/RECENT_SALE1/Basement_Adj
                    //TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/RECENT_SALE1/Basement_Finished_Adj














                    //TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/RECENT_SALE1/Street_Address1_Number CONTENT=1508

                    //macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:InputForm ATTR=NAME:abSpS1 CONTENT=4321");


                    macro3.AppendLine(@"FRAME NAME=_MAIN");
                    // macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:Form1 ATTR=NAME:PS_FORM/RECENT_SALE1/Datasource CONTENT=%MLS");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Orig_List_Date CONTENT=" + list_date);

                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Orig_List_Price CONTENT=" + orig_list_price.Replace("$", "").Replace(",", ""));
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Concessions_Type CONTENT=%None");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Land_Value CONTENT=5000");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Datasource CONTENT=%MLS");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Fee_Simple CONTENT=%Fee<SP>Simple");
                    macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Num_Units CONTENT=1");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/View CONTENT=%Residential");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/View_Comparison CONTENT=%Equal");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Other CONTENT=%Average");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Overall_Comp CONTENT=%Equal");
                    macro3.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:PS_FORM/" + input_comp_name + "/Heating_Cooling CONTENT=%FWA");


                    macroCode3 = macro3.ToString();
                    status = iim2.iimPlayCode(macroCode3, 60);

                    #endregion

                    status = iim.iimPlayCode(move_through_comps_macro.ToString(), 30);


                }

                s = iim.iimPlayCode(macro12.ToString());
                htmlCode = iim.iimGetLastExtract();
                //  header = iim.iimGetExtract(0);

                if (subjectAttachedRadioButton.Checked)
                {
                    m = new AttachedListing(htmlCode);
                }
                else if (subjectDetachedradioButton.Checked)
                {
                    m = new DetachedListing(htmlCode);
                }
                else
                {
                    m = new MLSListing(htmlCode);
                }

                #endregion
            }

            //
            //post processing
            //
            pBar1.Visible = false;
            pBar2.Visible = false;




            //#region equi-trax post processing
            //if (streetnumTextBox.Text == "equi-trax")
            //{

            ////    StringBuilder macro = new StringBuilder();
            ////    macro.AppendLine(@"");
            ////    macro.AppendLine(@"FRAME NAME=main");
            ////    macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:b_sv");
            ////    macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:dmsg_close_btn");

            ////    macro.AppendLine(@"");
            ////    macro.AppendLine(@"FRAME NAME=main");
            ////    macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:b_ap");
            ////    macro.AppendLine(@"FRAME NAME=iFileMan");
            ////    macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=CLASS:s_button");
            ////    macro.AppendLine(@"FRAME NAME=main");
            ////    macro.AppendLine(@"TAG POS=1 TYPE=IFRAME ATTR=ID:iFileMan");
            ////    macro.AppendLine(@"FRAME NAME=iFileMan");
            ////    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f6 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S1.jpg");
            ////    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f7 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S2.jpg");
            ////    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f8 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S3.jpg");
            ////    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f9 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L1.jpg");
            ////    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f10 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L2.jpg");
            ////    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f11 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L3.jpg");
            ////    macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:HMSG");
            ////    macro.AppendLine(@"TAG POS=18 TYPE=A ATTR=CLASS:s_button");
            ////    macro.AppendLine(@"TAG POS=1 TYPE=NOBR ATTR=TXT:Save<SP>Changes");
            ////    macro.AppendLine(@"FRAME NAME=main");
            ////    macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:b_ManageFiles_close");
            ////    macro.AppendLine(@"TAG POS=11 TYPE=A ATTR=CLASS:s_button");



            ////    string macroCode = macro.ToString();
            ////    status = iim2.iimPlayCode(macroCode, 30);

            //}
            //#endregion

            // if (streetnumTextBox.Text == "usres")
            //    {
            //
            //load comp pics
            //
            //    StringBuilder macro = new StringBuilder();
            //    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:InputForm ATTR=CLASS:auditlink");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:/BPOAGENTUPLOADPROC:* ATTR=NAME:bpo_*_65 CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") +  "\\L1.jpg");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:/BPOAGENTUPLOADPROC:* ATTR=NAME:bpo_*_70 CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") +  "\\L2.jpg");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:/BPOAGENTUPLOADPROC:* ATTR=NAME:bpo_*_75 CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") +  "\\L3.jpg");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:/BPOAGENTUPLOADPROC:* ATTR=NAME:bpo_*_85 CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") +  "\\S1.jpg");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:/BPOAGENTUPLOADPROC:* ATTR=NAME:bpo_*_90 CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") +  "\\S2.jpg");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:/BPOAGENTUPLOADPROC:* ATTR=NAME:bpo_*_95 CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") +  "\\S3.jpg");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:/BPOAGENTUPLOADPROC:* ATTR=VALUE:Upload<SP>Pictures");
            //    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=ACTION:/BPOAGENTUPLOADPROC:* ATTR=TXT:Assigned");
            //    string macroCode = macro.ToString();
            //     status = iim2.iimPlayCode(macroCode, 60);

            //}


            //if (streetnumTextBox.Text == "")
            //{
            //    StringBuilder macro = new StringBuilder();

            //    macro.AppendLine(@"FRAME NAME=_MAIN");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:Form1 ATTR=ID:fname_Listing1_Front CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") + "\\L1.jpg");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:Form1 ATTR=ID:fname_Listing2_Front CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") + "\\L2.jpg");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:Form1 ATTR=ID:fname_Listing3_Front CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") + "\\L3.jpg");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:Form1 ATTR=ID:fname_Sale1_Front CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") + "\\S1.jpg");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:Form1 ATTR=ID:fname_Sale2_Front CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") + "\\S2.jpg");
            //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:Form1 ATTR=ID:fname_Sale3_Front CONTENT=" + search_address_textbox.Text.Replace(" ", "<SP>") + "\\S3.jpg");
            //    macro.AppendLine(@"FRAME NAME=_TOP_MENU");
            //    macro.AppendLine(@"TAG POS=1 TYPE=B FORM=ID:Form1 ATTR=TXT:SAVE");

            //    string macroCode = macro.ToString();
            //    status = iim2.iimPlayCode(macroCode, 30);
            //}


            //this.button_imort_pics_Click(sender, e);

        }
    }
}