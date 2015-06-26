using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using WatiN.Core;
using System.Text.RegularExpressions;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using WatiN;
using WatiN.Core.Native;
using iMacros;
using System.Net;
using System.IO;
using Google.GData.Client;
using Google.GData.Spreadsheets;
//using //BitMiracle.Docotic.Pdf;
using System.Diagnostics;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Collections;
using HtmlAgilityPack;
using XnaFan.ImageComparison;
using Google.Apis.Fusiontables.v1;
using DotNetOpenAuth.OAuth2;
using Google.Apis.Authentication;
using Google.Apis.Authentication.OAuth2;
using Google.Apis.Authentication.OAuth2.DotNetOpenAuth;
using Google.Apis.Samples.Helper;
using System.Diagnostics;
using System.Threading;
using System.Xml;
using System.Xml.Schema;

namespace bpohelper
{
    public partial class Form1 : System.Windows.Forms.Form, IYourForm
    {
        private void button1_Click(object sender, EventArgs e)
        {
            m2m = new IE("https://m2m.aspengrove.net/Library/Security/Login.aspx?ReturnUrl=/index.aspx");



            m2m.TextField(Find.ByName("tbxPassword")).TypeText("P-192993");

            m2m.Button("btnLogin").Click();
            m2m.Link(Find.ByTitle("View Accepted Worklist")).Click();

            m2m.Form("formWorklist").Link(Find.ByTitle("Input BPO Data")).Click();


            mybpo.subjectaddress = m2m.Form("formOrderEditWizard").Table(Find.ByIndex(0)).TableCell(Find.ByIndex(2)).Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {

            // Attach Photos
            m2m = IE.AttachTo<IE>(Find.ByUrl(new Regex("https://www.emortgagelogic.com/main.html")));
            WatiN.Core.Frame mainframe = m2m.Frame(Find.ByName("main"));
            WatiN.Core.Form bpoform = mainframe.Form(Find.ByName("BPOForm"));

            mainframe.Link(Find.ByText("Attach Photos")).Click();


            //foreach (WatiN.Core.Frame f in mainframe.Frames)
            //{

            //    MessageBox.Show(f.Name);
            //    foreach (WatiN.Core.Form t in f.Forms)
            //    {

            //        MessageBox.Show(t.Name);
            //        //foreach (WatiN.Core.TextField tf in t.TextFields)
            //        //{

            //        //    MessageBox.Show(tf.Name);

            //        //}
            //    }
            //}

            foreach (WatiN.Core.TableCell tc in mainframe.Frame(Find.ByIndex(0)).Form(Find.ByIndex(1)).TableCells)
            {
                MessageBox.Show(tc.InnerHtml);
                // tc.Flash();
            }

            WatiN.Core.TableCell targetcell = mainframe.Frame(Find.ByIndex(0)).Form(Find.ByIndex(1)).TableCell(Find.ByText(new Regex("FileContent_f5")));

            targetcell.Flash();

            // WatiN.Core.DivCollection dc = mainframe.Divs;

            // WatiN.Core.Div d = dc.First(Find.ById("d_ManageFiles"));


            //WatiN.Core.Frame f = d.Div(Find.ByIndex(1)).Div(Find.ByIndex(1));

            WatiN.Core.TextFieldCollection tfc = mainframe.Frame(Find.ByIndex(0)).Form(Find.ByIndex(1)).TextFields;



            //targetcell.TextField(Find.ByName("FileContent_f5")).TypeText("test");





        }

        private void button3_Click(object sender, EventArgs e)
        {
            realist = new IE("http://realist2.firstamres.com/searchaddress.jsp?firsttime=yes&tabAction=SEARCHES");

            try
            {
                realist.TextField(Find.ByName("OWNER_PROP_HOUSE_NO_entered")).TypeText(mybpo.subjectstreetnumber());
                realist.TextField(Find.ByName("OWNER_PROP_STREET")).TypeText(mybpo.subjectstreetname());
                realist.Button(Find.ByName("Submit")).Click();
                //realist.Button(Find.ByValue("View Listing")).Click();
                //Regex urlpattern = new Regex("ExternalSearch");
                //listingswindow = IE.AttachToIE(Find.ByUrl(urlpattern));
            }
            catch
            {
                //do something
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Fill in step one button
            // Please provide the 2 most recent sales for the subject property from your MLS, or other online service.
            string ownername = realist.Table("Owner Info").TableRow(Find.ByTextInColumn("Owner Name:", 1)).TableCell(Find.ByIndex(realist.Table("Owner Info").TableCell(Find.ByText("Owner Name:")).Index + 2)).Text;

            m2m.Form("formOrderEditWizard").TextField(Find.ByName("tbxOwnerPublicRecord")).TypeText(mybpo.CompleteStepOne(ownername));




            //most recent sale
            // string lastsaledate = popup.Table("All_Sales").TableCell(Find.ByIndex(4)).Text;
            // string lastsaleamount = popup.Table("All_Sales").TableCell(Find.ByIndex(10)).Text; 



        }

        private void button5_Click(object sender, EventArgs e)
        {

            m2m = IE.AttachTo<IE>(Find.ByUrl(new Regex("https://www.emortgagelogic.com/main.html")));
            //based on form G
            WatiN.Core.Frame mainframe = m2m.Frame(Find.ByName("main"));
            //WatiN.Core.Form bpoform = mainframe.Form(Find.ByName("BPOForm"));
            //WatiN.Core.TextFieldCollection tfc = bpoform.TextFields;
            //WatiN.Core.TextField tf;
            mainframe.Span(Find.ById("themenu")).Link(Find.ByTitle("Find available work within your territory")).Click();

            //            LinkCollection lc = mainframe.Links;

            //lc.First(Find.ByTitle("Find available work within your territory")).Click();

        }

        private void button6_Click(object sender, EventArgs e)
        {
            mybpo.mlsurl = streetnameTextBox.Text;
            mybpo.bpocompany = streetnumTextBox.Text;

            if (radioButton1.Checked)
            {
                mybpo.Fill_Sale1();
                mybpo.Fill_Sale2();
                mybpo.Fill_Sale3();
            }
            else if (radioButton2.Checked)
            {
                mybpo.Fill_Sale2();
            }
            else if (radioButton3.Checked)
            {
                mybpo.Fill_Sale3();
            }
            else if (radioButton4.Checked)
            {
                mybpo.Fill_List1();
                mybpo.Fill_List2();
                mybpo.Fill_List3();
            }
            else if (radioButton5.Checked)
            {
                mybpo.Fill_List2();
            }
            else if (radioButton6.Checked)
            {
                mybpo.Fill_List3();
            }


        }

        private void button7_Click(object sender, EventArgs e)
        {

            if (m2m == null)
            {
                Regex url = new Regex("aspengrove.net");

                m2m = Browser.AttachTo<IE>(Find.ByUrl(url));

            }

            //m2mf = FireFox.AttachTo<FireFox>(Find.ByUrl("https://m2m.aspengrove.net/Image/ImageListAHM.aspx?OrderID=7386124&CancelUrl=https%3a%2f%2fm2m.aspengrove.net%2fOrder%2fOrderSelect.aspx%3fOrderID%3d7386124"));

            // m2mf = new FireFox();


            m2m.Link(Find.ByTitle("Add Image")).Click();


            m2m.SelectList("ddlCheckList").Select("Address Verification (Subject)");
            // m2m.Form("formImageEdit").Table(Find.ByIndex(1)).TableRow("trFile").TableCell(Find.ByIndex(1)).TextField(Find.ByName("UploadFile")).TypeText("\\\\DLINK-DA7C61\\HDD_a\\OK Realty Plus\\bpos\\4413 Dennis\\list1.jpg");



            FileUpload fup = m2m.Form("formImageEdit").Table(Find.ByIndex(1)).TableRow("trFile").TableCell(Find.ByIndex(1)).FileUpload("UploadFile");

            // m2m.DialogWatcher.CloseUnhandledDialogs = true;

            fup.Set("\\\\DLINK-DA7C61\\HDD_a\\OK Realty Plus\\bpos\\4413 Dennis\\list1.jpg");

            //fup.Change();


            string temp = fup.FileName;

            fup.KeyPress('d');


            // m2m.Form("formImageEdit").Button("btnSave").Click();








        }

        private void imarco_test_Click(object sender, EventArgs e)
        {
            iim.iimOpen("-ie", true, 5);
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"VERSION BUILD=8001865");
            macro.AppendLine(@"TAB T=1");
            macro.AppendLine(@"TAB CLOSEALLOTHERS");
            macro.AppendLine(@"URL GOTO=http://connectmls2.mredllc.com/");
            string macroCode = macro.ToString();
            status = iim.iimPlayCode(macroCode, 5);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            iim.iimExit(5);

            iim.iimOpen("-ie", false, 60);


            //  StringBuilder macro3 = new StringBuilder();

            //  iim.iimOpen("-ie", false);


            //  macro3.AppendLine(@"TAB T=2");
            // // macro3.AppendLine(@"TAB CLOSEALLOTHERS");

            //  macro3.AppendLine(@"FRAME NAME=_MAIN");
            //  macro3.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Cur_Market_Cond&&VALUE:STABLE CONTENT=YES");
            // // macro3.AppendLine(@"TAG POS=11 TYPE=TABLE FORM=NAME:dc ATTR=TXT:* EXTRACT=TXT");
            ////  macro3.AppendLine(@"TAG POS=16 TYPE=TABLE FORM=NAME:dc ATTR=TXT:* EXTRACT=TXT");
            ////  macro3.AppendLine(@"TAG POS=17 TYPE=TABLE FORM=NAME:dc ATTR=TXT:* EXTRACT=TXT");



            //  string macroCode3 = macro3.ToString();
            //  status = iim.iimPlayCode(macroCode3, 5);
        }

        private void button9_Click(object sender, EventArgs e)
        {

            //
            //usres BPO header info with bpo form open
            //
            //BPO order number
            //macro.AppendLine(@"TAG POS=3 TYPE=TABLE FORM=NAME:InputForm ATTR=CLASS:form_txt_blk EXTRACT=TXT");\
            //Subject address and loan number
            //macro.AppendLine(@"TAG POS=10 TYPE=TR FORM=NAME:InputForm ATTR=TXT:* EXTRACT=TXT");

            //save order info to DB
            StringBuilder macro = new StringBuilder();

            macro.AppendLine(@"FRAME NAME=_MAIN");
            macro.AppendLine(@"TAG POS=7 TYPE=TABLE FORM=ID:Form1 ATTR=TXT:* EXTRACT=TXT");
            macro.AppendLine(@"TAG POS=11 TYPE=TR FORM=ID:Form1 ATTR=TXT:* EXTRACT=TXT");
            //macro.AppendLine(@"TAG POS=6 TYPE=TABLE FORM=ID:Form1 ATTR=TXT:* EXTRACT=TXT");

            //macro.AppendLine(@"TAG POS=9 TYPE=TABLE FORM=ID:Form1 ATTR=TXT:* EXTRACT=TXT");
            //macro.AppendLine(@"TAG POS=8 TYPE=TABLE FORM=ID:Form1 ATTR=TXT:* EXTRACT=TXT");
            string macroCode = macro.ToString();
            iim.iimOpen("-ie", false);
            status = iim.iimPlayCode(macroCode, 60);

            DataTable BPOtable = bpoTableAdapter1.GetData();


            //
            //imortgage Table1
            //320698006b#NEXT##NEWLINE#Product:BPO Exterior#NEXT##NEWLINE#Due Date:6/4/2012 1:47:56 PM#NEXT##NEWLINE#You have exceeeded the 24 hour reassign limit. You can not cancel this order.
            //320600564b#NEXT##NEWLINE#Product:BPO Exterior#NEXT##NEWLINE#Due Date:6/8/2012#NEXT##NEWLINE#You have exceeeded the 24 hour reassign limit. You can not cancel this order.
            string table1 = iim.iimGetLastExtract(1);
            string table2 = iim.iimGetLastExtract(2);

            string pattern = "^(\\d+\\w)#NEXT#";
            Match match = Regex.Match(table1, pattern);
            string ordernumber = match.Groups[1].Value;

            pattern = "Product:\\w+ (\\w+)#NEXT#";
            match = Regex.Match(table1, pattern);
            string type = match.Groups[1].Value;

            //pattern = "Due Date:(\\d+\\/\\d+\\/\\d+\\s+\\d+\\:\\d+\\:\\d\\s+\\ww)#NEXT#";

            //with time or without time

            DateTime duedate;
            try
            {
                pattern = "Due Date:(\\d+\\/\\d+\\/\\d+\\s\\d+:\\d+:\\d+\\s+\\w+)";
                match = Regex.Match(table1, pattern);
                duedate = Convert.ToDateTime(match.Groups[1].Value);
            }
            catch
            {
                pattern = "Due Date:(\\d+\\/\\d+\\/\\d+)";
                match = Regex.Match(table1, pattern);
                duedate = Convert.ToDateTime(match.Groups[1].Value);
            }
            pattern = "^\\$(\\d+.\\d+)";
            match = Regex.Match(table2, pattern);
            decimal payment_amount = Convert.ToDecimal(match.Groups[1].Value);

            string msg_string = ordernumber + " " + type + " " + duedate.ToShortDateString() + " " + payment_amount.ToString();

            MessageBox.Show(msg_string);

            try
            {
                int x = this.bpoTableAdapter1.Insert("Imortgage", "", subjectpin_textbox.Text, "", "", "", "", "", "", "", 0, 0, 0, 0, DateTime.Now, duedate, type, ordernumber, payment_amount, 0, DateTime.Now, duedate, new DateTime(), "", "", "", false);

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }


            //            try
            //            {
            //                int x = this.mailing_listTableAdapter1.Insert(mlrow.Pin, mlrow.Address, mlrow.City, mlrow.Zip, firstname,lastname , mlrow.MaxOfEventDate, mlrow.State, mlrow.Dateofmailing, mlrow.DoNotMail);
            //            }
            //            catch (Exception ex)
            //            {

            //                MessageBox.Show(ex.Message);
            //            }

            //        }



            //     DataTable mailings_datatable = mailingsTableAdapter1.GetData();
            //    DataTable property_table = propertyTableAdapter1.GetData();
            //    phoenix_projectDataSet.mailing_listDataTable mailingslist_dt = mailing_listTableAdapter1.GetData();
            //    DataTable people_prop_link_dt = people_property_linkTableAdapter1.GetData();
            //    DataTable people_dt = peopleTableAdapter1.GetData();
            //    DataTable court_events_dt = court_eventsTableAdapter1.GetData();
            //    DataTable cases_dt = casesTableAdapter1.GetData();


            //    mailingslist_dt.Clear();


            //    string exp  = "Pin=813404011";

            //    DataRowCollection ptrc = property_table.Rows;


            //    DataRow[] drc = mailings_datatable.Select(exp);





            //    //DataRowCollection drc = mailings_datatable.Rows;

            //    foreach (DataRow r in ptrc)
            //    {
            //        object [] property = r.ItemArray;


            //        //string expression = "Pin = '813404011'";
            //        string expression = "Pin = '" + property[0].ToString() + "'";

            //        drc = mailings_datatable.Select(expression);

            //        if (drc.Length == 0)
            //        {



            //            //object[] temprec = { 1234, "address", "city", "zip", "firstname", "lastname", DateTime.Now, "state", DateTime.Now, false };


            //            foreach (DataRow ppl in people_prop_link_dt.Rows)
            //            {

            //                if (ppl.ItemArray[0].Equals(property[0]))
            //                {


            //                    DataRow [] owner = people_dt.Select("ID='" + ppl.ItemArray[1] + "'");

            //                    if (owner[0].ItemArray[4].Equals(false))
            //                    {
            //                        string exp2 = "caseid='" + property[5].ToString()+ "'";

            //                        DataRow [] ce = court_events_dt.Select(exp2, "EventDate");



            //                        object[] temprec = { property[0], property[1], property[2], property[3], owner[0].ItemArray[1], owner[0].ItemArray[2], ce[ce.Length-1].ItemArray[2], "IL", DateTime.Now, false };

            //                        mailingslist_dt.LoadDataRow(temprec, true);
            //                    }
            //                }
            //            }



            //        }


            //    }



            //    mailing_listTableAdapter1.ClearBeforeFill = true;




            //        phoenix_projectDataSet.mailing_listRow mlrow;
            //        for (int i = 0; i < 24; i++)
            //        {

            //            mlrow = (phoenix_projectDataSet.mailing_listRow)mailingslist_dt.Rows[i];

            //           string firstname = "";
            //           string lastname = "";
            //          //  if (mlrow.FirstName != DBNull.Value)

            //            if (!mlrow.IsFirstNameNull())
            //                firstname = mlrow.FirstName;
            //            if (!mlrow.IsLastNameNull() )
            //                 lastname = mlrow.LastName;

            //            try
            //            {
            //                int x = this.mailing_listTableAdapter1.Insert(mlrow.Pin, mlrow.Address, mlrow.City, mlrow.Zip, firstname,lastname , mlrow.MaxOfEventDate, mlrow.State, mlrow.Dateofmailing, mlrow.DoNotMail);
            //            }
            //            catch (Exception ex)
            //            {

            //                MessageBox.Show(ex.Message);
            //            }

            //        }

            //        MessageBox.Show("Print cards, click to continue");


            //        phoenix_projectDataSet.mailing_listDataTable current_table = mailing_listTableAdapter1.GetData();
            //        //phoenix_projectDataSet.mailingsDataTable allmailings_dt;





            //        foreach (phoenix_projectDataSet.mailing_listRow cmlrow in current_table.Rows)
            //        {


            //            int x = this.mailingsTableAdapter1.Insert(cmlrow.Pin, cmlrow.Address, cmlrow.City, cmlrow.Zip, cmlrow.FirstName, cmlrow.LastName, cmlrow.MaxOfEventDate, cmlrow.State, cmlrow.Dateofmailing, "helpforhomeowners v3");


            //        }
            //        int y = mailing_listTableAdapter1.DeleteQuery();








            //    //int x = mailing_listTableAdapter1.DeleteQuery();
            //    //int y = mailing_listTableAdapter1.FillME(mailingslist_dt);


            //   // mailing_listTableAdapter1.Dispose();






            //    //mailing_listTableAdapter1.



            //}



        }

        private void import_realist_button_Click(object sender, EventArgs e)
        {

            DataTable subject_table = subjectTableAdapter.GetData();


            iim.iimOpen("-ie", false, 30);

            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"VERSION BUILD=8021952");
            macro.AppendLine(@"TAB T=2");
            macro.AppendLine(@"TAG POS=13 TYPE=TD ATTR=CLASS:multiColumnDataText");
            macro.AppendLine(@"TAG POS=3 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");
            macro.AppendLine(@"TAG POS=2 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");
            macro.AppendLine(@"TAG POS=4 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");
            macro.AppendLine(@"TAG POS=5 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");
            macro.AppendLine(@"TAG POS=4 TYPE=TABLE ATTR=CLASS:dataGridTable EXTRACT=TXT");
            macro.AppendLine(@"TAG POS=1 TYPE=TABLE ATTR=CLASS:multiColumnTable EXTRACT=TXT");
            macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:headerText EXTRACT=TXT");


            string macroCode = macro.ToString();
            status = iim.iimPlayCode(macroCode, 60);


            //Table1: Tax information
            //Parcel ID:#NEXT#0307178017#NEXT#% Improved:#NEXT#73%#NEXT##NEWLINE#Tax Area:#NEXT#DU056#NEXT#Exemption(s):#NEXT#Homestead,#NEXT##NEWLINE#Lot # :#NEXT#160#NEXT##NEXT# #NEXT##NEWLINE#Legal Description:#NEXT#THE PRAIRIES AND MEADOWS OF WINCHESTER GLEN LT 160
            //
            string table1 = iim.iimGetLastExtract(1);
            string table2 = iim.iimGetLastExtract(2);
            string table3 = iim.iimGetLastExtract(3);
            string table4 = iim.iimGetLastExtract(4);
            string table5 = iim.iimGetLastExtract(5);
            string table6 = iim.iimGetLastExtract(6);
            string realist_address = iim.iimGetLastExtract(7);


            string pattern = "^Parcel\\s+ID:#NEXT#(\\d+)#NEXT#";
            Match match = Regex.Match(table1, pattern);
            string pin = match.Groups[1].Value;

            pattern = ". Improved:#NEXT#(\\d+)";
            match = Regex.Match(table1, pattern);
            string realist_percent_improved = match.Groups[1].Value;


            //Table2: Location information
            //
            //
            pattern = "School District:#NEXT#([^#]+)";
            match = Regex.Match(table2, pattern);
            string realist_school_district = match.Groups[1].Value;

            pattern = "Census Tract:#NEXT#(\\d+.\\d+)";
            match = Regex.Match(table2, pattern);
            string realist_census_tract = match.Groups[1].Value;

            pattern = "Carrier Route:#NEXT#(\\w\\d+)";
            match = Regex.Match(table2, pattern);
            string realist_carrier_route = match.Groups[1].Value;

            pattern = "Subdivision:#NEXT#(([^#]+))";
            match = Regex.Match(table2, pattern);
            string realist_subdivision = match.Groups[1].Value;

            pattern = "Township:#NEXT#(([^#]+))";
            match = Regex.Match(table2, pattern);
            string realist_township = match.Groups[1].Value;



            //Table3: Characteristics
            //
            //
            pattern = "Lot Acres:#NEXT#(([^#]+))";
            match = Regex.Match(table3, pattern);
            string realist_lot_acres = match.Groups[1].Value;


            //Table4: Estimated Value
            //
            //


            //Table5: Market and Sales History
            //
            //

            //
            //process address as
            //3122 Erika Ln, Carpentersville, IL 60110-3462, Kane County
            //to match fields in bpo database subject table
            pattern = "Lot Acres:#NEXT#(([^#]+))";
            match = Regex.Match(table3, pattern);
            string[] temparray = realist_address.Split(',');

            string address = temparray[0];
            string city = temparray[1];

            pattern = "(\\w.+) County";
            match = Regex.Match(temparray[3], pattern);
            string county = match.Groups[1].Value;
            //
            //end process address
            //



            MessageBox.Show(realist_lot_acres + " " + pin + " " + realist_percent_improved + " " + realist_school_district + " " + realist_census_tract + " " + realist_carrier_route + " " + realist_subdivision);

            try
            {
                //int x = this.subjectTableAdapter.InsertQuery(pin,"", realist_census_tract, address,city,county,realist_township, null,null,null,null);


            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }


        }

        private void import_market_stats_button_Click(object sender, EventArgs e)
        {
            iim.iimOpen("-ie", false, 30);
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"VERSION BUILD=8021952");
            macro.AppendLine(@"FRAME F=5");
            macro.AppendLine(@"TAG POS=2 TYPE=TABLE FORM=NAME:dc ATTR=TXT:* EXTRACT=TXT");
            string macroCode = macro.ToString();
            status = iim.iimPlayCode(macroCode, 60);

            string market_stats_table = iim.iimGetLastExtract(1);

            market_stats_string_textbox.Text = market_stats_table;

            //MessageBox.Show(market_stats_table);

            //iim2.iimOpen("", false, 30);
            //StringBuilder macro2 = new StringBuilder();
            //macro2.AppendLine(@"FRAME NAME=_MAIN");
            //macro2.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/NeighborhoodSales1 CONTENT=63000");
            //macro2.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/NeighborhoodSales2 CONTENT=126000");
            //macro2.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/NeighborhoodSalesNum CONTENT=29");
            //macro2.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/NeighborhoodListings1 CONTENT=65000");
            //macro2.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/NeighborhoodListings2 CONTENT=130000");
            //macro2.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/NeighborhoodListingsNum CONTENT=16");
            //macro2.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Number_Competing CONTENT=10");
            //macro2.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:PS_FORM/MARKET_DATA/Number_Boarded CONTENT=0");
            // macroCode = macro.ToString();
            //status = iim2.iimPlayCode(macroCode, 30);


        }

        private void upload_eval_interior_button_Click(object sender, EventArgs e)
        {


            iim.iimOpen("", false, 30);
            StringBuilder macro = new StringBuilder();

            macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:content");
            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"ONSCRIPTERROR CONTINUE=YES");

            //string path = "C:\\Users\\Scott\\Documents\\My<SP>Dropbox\\BPOs\\324<SP>STEWART<SP>AVE\\";
            string path = search_address_textbox.Text + "\\";


            string[,] picturenames = { { "front", "1" }, { "leftside", "2" }, { "rightside", "3" }, { "rear", "4" }, { "street", "5" },
                                            { "s1", "6" },{ "s2", "7" },{ "s3", "8" },{ "l1", "9" },{ "l2", "10" },{ "l3", "11" },{ "address", "12" },
                                                { "angle", "13" },{ "garage-exterior", "14" },  { "living", "16" }, { "dining", "17" },{ "kitchen", "18" },
                                                { "family", "19" }, { "bed-1", "20" }, { "bath-1", "21" }, 
                                                    };



            // "", };

            for (int i = 0; i < picturenames.GetLength(0); i++)
            {
                macro.AppendLine(@"TAG POS=" + picturenames[i, 1] + " TYPE=INPUT:FILE FORM=NAME:userupload ATTR=NAME:userfile[] CONTENT=" + path.Replace(" ", "<SP>") + picturenames[i, 0] + ".jpg");

            }



            //FRONT of subject
            //LEFT of subject side
            //RIGHT of subject side
            //REAR of subject
            //STREET view of neighborhood
            //COMP SALE 1
            //COMP SALE 2
            //COMP SALE 3
            //COMP LISTING 1
            //COMP LISTING 2
            //COMP LISTING 3
            //Address Photo
            //Subject Front Angled View
            //GARAGE EXTERIOR VIEW
            //NEGATIVE NEIGHBORHOOD INFLUENCE(S)
            //LIVING ROOM
            //DINING ROOM
            //KITCHEN
            //FAMILY ROOM/GREAT ROOM/DEN
            //BEDROOM(S)
            //CLOSET(S)
            //BATHROOM(S)
            //HALF BATH(S)
            //BASEMENT
            //ATTIC
            //HALLWAY(S)
            //MECHANICALS
            //GARAGE INTERIOR VIEW
            //EXTERIOR DAMAGE
            //INTERIOR DAMAGE
            //Exterior Damage - STRUCTURAL
            //Exterior Damage - ROOF
            //Exterior Damage - WINDOWS/DOORS
            //Exterior Damage - PAINTING
            //Exterior Damage - SIDING/TRIM
            //Exterior Damage - LANDSCAPE
            //Exterior Damage - GARAGE
            //Exterior Damage - POOL/SPA
            //Exterior Damage - TRASH REMOVAL
            //Exterior Damage - TERMITE/PEST DAMAGE
            //Exterior Damage - OUTBUILDINGS
            //Interior Damage - STRUCTURAL
            //Interior Damage - PLUMBING/FIXTURES
            //Interior Damage - FLOOR/CARPET
            //Interior Damage - WALLS/CEILING
            //Interior Damage - PAINTING
            //Interior Damage - UTILITIES
            //Interior Damage - APPLIANCES
            //Interior Damage - KITCHEN CABINETS/CARPENTRY
            //Interior Damage - TRASH REMOVAL
            //Interior Damage - TERMITE/PEST DAMAGE
            //Misc. Photo 


            //bedroom upload
            //    macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:userupload ATTR=ID:FileTypID CONTENT=%10");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:userupload ATTR=NAME:filedesc[] CONTENT=bed-1");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:userupload ATTR=NAME:userfile[] CONTENT=C:\fakepath\bed-1.jpg");
            //macro.AppendLine(@"TAG POS=2 TYPE=SELECT FORM=NAME:userupload ATTR=ID:FileTypID CONTENT=%10");
            //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:TEXT FORM=NAME:userupload ATTR=NAME:filedesc[] CONTENT=bed-2");
            //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:FILE FORM=NAME:userupload ATTR=NAME:userfile[] CONTENT=C:\fakepath\bed-2.jpg");
            //macro.AppendLine(@"TAG POS=3 TYPE=SELECT FORM=NAME:userupload ATTR=ID:FileTypID CONTENT=%10");
            //macro.AppendLine(@"TAG POS=3 TYPE=INPUT:TEXT FORM=NAME:userupload ATTR=NAME:filedesc[] CONTENT=bed-3");
            //macro.AppendLine(@"TAG POS=3 TYPE=INPUT:FILE FORM=NAME:userupload ATTR=NAME:userfile[] CONTENT=C:\fakepath\bed-3.jpg");
            //macro.AppendLine(@"TAG POS=4 TYPE=SELECT FORM=NAME:userupload ATTR=ID:FileTypID CONTENT=%10");
            //macro.AppendLine(@"TAG POS=4 TYPE=INPUT:TEXT FORM=NAME:userupload ATTR=NAME:filedesc[] CONTENT=bed-4");
            //macro.AppendLine(@"TAG POS=4 TYPE=INPUT:FILE FORM=NAME:userupload ATTR=NAME:userfile[] CONTENT=C:\fakepath\bed-4.jpg");
            //macro.AppendLine(@"TAG POS=5 TYPE=SELECT FORM=NAME:userupload ATTR=ID:FileTypID CONTENT=%10");
            //macro.AppendLine(@"TAG POS=5 TYPE=INPUT:TEXT FORM=NAME:userupload ATTR=NAME:filedesc[] CONTENT=bed-5");
            //macro.AppendLine(@"TAG POS=5 TYPE=INPUT:FILE FORM=NAME:userupload ATTR=NAME:userfile[] CONTENT=C:\fakepath\bed-5.jpg");

            //bath upload
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:userupload ATTR=ID:FileTypID CONTENT=%12");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:userupload ATTR=NAME:filedesc[] CONTENT=bath-1");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:userupload ATTR=NAME:userfile[] CONTENT=C:\fakepath\bath-1.jpg");
            //macro.AppendLine(@"TAG POS=2 TYPE=SELECT FORM=NAME:userupload ATTR=ID:FileTypID CONTENT=%12");
            //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:TEXT FORM=NAME:userupload ATTR=NAME:filedesc[] CONTENT=bath-2");
            //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:FILE FORM=NAME:userupload ATTR=NAME:userfile[] CONTENT=C:\fakepath\bath-2.jpg");
            //macro.AppendLine(@"TAG POS=3 TYPE=SELECT FORM=NAME:userupload ATTR=ID:FileTypID CONTENT=%12");
            //macro.AppendLine(@"TAG POS=3 TYPE=INPUT:TEXT FORM=NAME:userupload ATTR=NAME:filedesc[] CONTENT=bath-3");
            //macro.AppendLine(@"TAG POS=3 TYPE=INPUT:FILE FORM=NAME:userupload ATTR=NAME:userfile[] CONTENT=C:\fakepath\bath-3.jpg");
            //macro.AppendLine(@"TAG POS=4 TYPE=SELECT FORM=NAME:userupload ATTR=ID:FileTypID CONTENT=%12");
            //macro.AppendLine(@"TAG POS=4 TYPE=INPUT:TEXT FORM=NAME:userupload ATTR=NAME:filedesc[] CONTENT=bath-4");
            //macro.AppendLine(@"TAG POS=4 TYPE=INPUT:FILE FORM=NAME:userupload ATTR=NAME:userfile[] CONTENT=C:\fakepath\bath-4.jpg");


            //basement upload
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=NAME:userupload ATTR=ID:FileTypID CONTENT=%14");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:userupload ATTR=NAME:filedesc[] CONTENT=basement-1");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:userupload ATTR=NAME:userfile[] CONTENT=C:\fakepath\basement-1.jpg");
            //macro.AppendLine(@"TAG POS=2 TYPE=SELECT FORM=NAME:userupload ATTR=ID:FileTypID CONTENT=%14");
            //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:TEXT FORM=NAME:userupload ATTR=NAME:filedesc[] CONTENT=basement-2");
            //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:FILE FORM=NAME:userupload ATTR=NAME:userfile[] CONTENT=C:\fakepath\basement-2.jpg");
            //macro.AppendLine(@"TAG POS=3 TYPE=SELECT FORM=NAME:userupload ATTR=ID:FileTypID CONTENT=%14");
            //macro.AppendLine(@"TAG POS=3 TYPE=INPUT:TEXT FORM=NAME:userupload ATTR=NAME:filedesc[] CONTENT=basement-3");
            //macro.AppendLine(@"TAG POS=3 TYPE=INPUT:FILE FORM=NAME:userupload ATTR=NAME:userfile[] CONTENT=C:\fakepath\basement-3.jpg");
            //macro.AppendLine(@"TAG POS=4 TYPE=SELECT FORM=NAME:userupload ATTR=ID:FileTypID CONTENT=%14");
            //macro.AppendLine(@"TAG POS=4 TYPE=INPUT:TEXT FORM=NAME:userupload ATTR=NAME:filedesc[] CONTENT=basement-4");
            //macro.AppendLine(@"TAG POS=4 TYPE=INPUT:FILE FORM=NAME:userupload ATTR=NAME:userfile[] CONTENT=C:\fakepath\basement-4.jpg");

            string macroCode = macro.ToString();
            status = iim.iimPlayCode(macroCode, 60);
        }

        private void import_mls_sheet_button_Click(object sender, EventArgs e)
        {
            iim.iimOpen("-ie", false, 30);
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"FRAME NAME=workspace");
            //macro.AppendLine(@"TAG POS=1 TYPE=TABLE FORM=NAME:dc ATTR=CLASS:gridview EXTRACT=TXT");
            // macro.AppendLine(@"TAG POS=3 TYPE=TD FORM=NAME:dc ATTR=TXT:* EXTRACT=TXT");


            for (int i = 1; i < 22; i++)
            {
                macro.AppendLine(@"TAG POS=" + i.ToString() + " TYPE=TABLE FORM=NAME:dc ATTR=TXT:* EXTRACT=TXT");

            }

            string macroCode = macro.ToString();
            status = iim.iimPlayCode(macroCode, 30);
            string extraction = iim.iimGetLastExtract(1);





            textBox1.Text = extraction;


        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button_eval_ext_pics_Click(object sender, EventArgs e)
        {
            //iim.iimOpen("", false, 30);
            StringBuilder macro = new StringBuilder();

            //macro.AppendLine(@"TAG POS=1 TYPE=DIV ATTR=ID:content");
            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"ONSCRIPTERROR CONTINUE=YES");

            //string path = "C:\\Users\\Scott\\Documents\\My<SP>Dropbox\\BPOs\\324<SP>STEWART<SP>AVE\\";
            string path = search_address_textbox.Text + "\\";


            string[,] picturenames = { { "front", "1" }, { "street", "2" }, { "s1", "3" }, { "s2", "4" }, { "s3", "5" },
                                            { "l1", "6" },{ "l2", "7" },{ "l3", "8" },{ "address", "9" },{ "angle", "10" }
                                                    };



            // "", };

            for (int i = 0; i < picturenames.GetLength(0); i++)
            {
                macro.AppendLine(@"TAG POS=" + picturenames[i, 1] + " TYPE=INPUT:FILE FORM=NAME:userupload ATTR=NAME:userfile[] CONTENT=" + path.Replace(" ", "<SP>") + picturenames[i, 0] + ".jpg");

            }



            string macroCode = macro.ToString();
            status = iim2.iimPlayCode(macroCode, 60);
        }

        private void button_imort_pics_Click(object sender, EventArgs e)
        {

            StringBuilder macro = new StringBuilder();
            iim.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");
            iim2.iimPlayCode(@"ADD !EXTRACT {{!URLCURRENT}}");

            string currentUrl = iim2.iimGetLastExtract();

            macro.Clear();
            macro.AppendLine(@"SET !TIMEOUT_STEP 30");
            macro.AppendLine(@"SET !ERRORIGNORE YES");


            if (currentUrl.ToLower().Contains("dnaforms"))
            {
                macro.AppendLine(@"FRAME NAME=pageSelect");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=NAME:myForm ATTR=NAME:ARImages");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagetype CONTENT=%oitCompPhoto");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:salescompnumber CONTENT=%1");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagefile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S1.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:addremoveimageswork.aspx ATTR=CLASS:buttons");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RESET FORM=ACTION:addremoveimageswork.aspx ATTR=ID:reset");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagetype CONTENT=%oitCompPhoto");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:salescompnumber CONTENT=%2");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagefile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S2.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:addremoveimageswork.aspx ATTR=CLASS:buttons");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RESET FORM=ACTION:addremoveimageswork.aspx ATTR=ID:reset");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagetype CONTENT=%oitCompPhoto");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:salescompnumber CONTENT=%3");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagefile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S3.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:addremoveimageswork.aspx ATTR=CLASS:buttons");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RESET FORM=ACTION:addremoveimageswork.aspx ATTR=ID:reset");

                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagetype CONTENT=%oitListingCompPhoto");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:listingcompnumber CONTENT=%1");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagefile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L1.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:addremoveimageswork.aspx ATTR=CLASS:buttons");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RESET FORM=ACTION:addremoveimageswork.aspx ATTR=ID:reset");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagetype CONTENT=%oitListingCompPhoto");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:listingcompnumber CONTENT=%2");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagefile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L2.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:addremoveimageswork.aspx ATTR=CLASS:buttons");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RESET FORM=ACTION:addremoveimageswork.aspx ATTR=ID:reset");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagetype CONTENT=%oitListingCompPhoto");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:listingcompnumber CONTENT=%3");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagefile CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L3.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:addremoveimageswork.aspx ATTR=CLASS:buttons");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RESET FORM=ACTION:addremoveimageswork.aspx ATTR=ID:reset");

                string[] fileEntries = Directory.GetFiles(search_address_textbox.Text);

                foreach (string fileName in fileEntries)
                {
                    if (fileName.ToLower().Contains(".jpg"))
                    {
                        if (fileName.ToLower().Contains("address"))
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagetype CONTENT=%oitSubjectAdditional"); 
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:desc3 CONTENT=Address");
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagefile CONTENT=" + fileName.Replace(" ", "<SP>"));
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:addremoveimageswork.aspx ATTR=CLASS:buttons");
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RESET FORM=ACTION:addremoveimageswork.aspx ATTR=ID:reset");

                        }
                        else if (fileName.ToLower().Contains("front"))
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagetype CONTENT=%oitSubjectFrontView");
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:desc1 CONTENT=Front");
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagefile CONTENT=" + fileName.Replace(" ", "<SP>"));
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:addremoveimageswork.aspx ATTR=CLASS:buttons");
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RESET FORM=ACTION:addremoveimageswork.aspx ATTR=ID:reset");
                        }
                        else if (fileName.ToLower().Contains("street"))
                        {

                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagetype CONTENT=%oitSubjectStreetView");
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:desc2 CONTENT=Street");
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagefile CONTENT=" + fileName.Replace(" ", "<SP>"));
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:addremoveimageswork.aspx ATTR=CLASS:buttons");
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RESET FORM=ACTION:addremoveimageswork.aspx ATTR=ID:reset");
                        }
                        else if (fileName.ToLower().Contains("side"))
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagetype CONTENT=%oitSubjectRearView");
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:desc1 CONTENT=Side");
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:addremoveimageswork.aspx ATTR=NAME:imagefile CONTENT=" + fileName.Replace(" ", "<SP>"));
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:addremoveimageswork.aspx ATTR=CLASS:buttons");
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RESET FORM=ACTION:addremoveimageswork.aspx ATTR=ID:reset");
                        }
                    }
                }

           


                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                //macro.AppendLine(@"TAG POS=2 TYPE=INPUT:BUTTON FORM=ACTION:addremoveimageswork.aspx ATTR=CLASS:buttons");



                string macroCode = macro.ToString();
                status = iim2.iimPlayCode(macroCode, 120);

            }

            if (currentUrl.ToLower().Contains("reo-central"))
            {
                macro.AppendLine("FRAME NAME=\"AddBpoWindow\"");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:form1 ATTR=ID:Wizard1_uploadComp1file0 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S1.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:form1 ATTR=ID:Wizard1_uploadComp2file0 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S2.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:form1 ATTR=ID:Wizard1_uploadComp3file0 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S3.jpg");

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:form1 ATTR=ID:Wizard1_uploadList1file0 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L1.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:form1 ATTR=ID:Wizard1_uploadList2file0 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L2.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:form1 ATTR=ID:Wizard1_uploadList3file0 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L3.jpg");

                string macroCode = macro.ToString();
                status = iim2.iimPlayCode(macroCode, 120);

            }

            if (currentUrl.ToLower().Contains("solutionstar"))
            {

              
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:# ATTR=NAME:user_0_Image CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S1.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:# ATTR=NAME:Image_0_Description CONTENT=Sale1");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=ACTION:# ATTR=NAME:toggler*");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:*# ATTR=NAME:user_1_Image CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S2.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:*# ATTR=NAME:Image_1_Description CONTENT=Sale2");
                macro.AppendLine(@"TAG POS=2 TYPE=INPUT:BUTTON FORM=ACTION:*# ATTR=NAME:toggler*");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:*# ATTR=NAME:user_2_Image CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S3.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:*# ATTR=NAME:Image_2_Description CONTENT=Sale3");
                macro.AppendLine(@"TAG POS=3 TYPE=INPUT:BUTTON FORM=ACTION:*# ATTR=NAME:toggler*");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:*# ATTR=NAME:user_3_Image CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L1.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:*# ATTR=NAME:Image_3_Description CONTENT=List1");
                macro.AppendLine(@"TAG POS=4 TYPE=INPUT:BUTTON FORM=ACTION:*# ATTR=NAME:toggler*");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:*# ATTR=NAME:user_4_Image CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L2.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:*# ATTR=NAME:Image_4_Description CONTENT=List2");
                macro.AppendLine(@"TAG POS=5 TYPE=INPUT:BUTTON FORM=ACTION:*# ATTR=NAME:toggler*");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:*# ATTR=NAME:user_5_Image CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L3.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:*# ATTR=NAME:Image_5_Description CONTENT=List3");
                macro.AppendLine(@"TAG POS=6 TYPE=INPUT:BUTTON FORM=ACTION:*# ATTR=NAME:toggler*");

                string[] fileEntries = Directory.GetFiles(search_address_textbox.Text);

                foreach (string fileName in fileEntries)
                {
                    if (fileName.ToLower().Contains(".jpg") && !Regex.IsMatch(fileName, @"_large|_medium|_thumb"))
                    {
                        if (fileName.ToLower().Contains("address"))
                        {
                            GenerateVersions(fileName);
                            string fileToUpload = fileName.Replace(".jpg", "") + "_large.jpg";


                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:*# ATTR=NAME:user_6_Image CONTENT=" + fileToUpload.Replace(" ", "<SP>"));
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:*# ATTR=NAME:Image_6_Description CONTENT=Address");
                            macro.AppendLine(@"TAG POS=7 TYPE=INPUT:BUTTON FORM=ACTION:*# ATTR=NAME:toggler*");
                        }
                        else if (fileName.ToLower().Contains("front"))
                        {
                            GenerateVersions(fileName);
                            string fileToUpload = fileName.Replace(".jpg", "") + "_large.jpg";
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:*# ATTR=NAME:user_7_Image CONTENT=" + fileToUpload.Replace(" ", "<SP>"));
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:*# ATTR=NAME:Image_7_Description CONTENT=Front");
                            macro.AppendLine(@"TAG POS=8 TYPE=INPUT:BUTTON FORM=ACTION:*# ATTR=NAME:toggler*");
                        }
                        else if (fileName.ToLower().Contains("street1"))
                        {
                            GenerateVersions(fileName);
                            string fileToUpload = fileName.Replace(".jpg", "") + "_large.jpg";
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:*# ATTR=NAME:user_8_Image CONTENT=" + fileToUpload.Replace(" ", "<SP>"));
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:*# ATTR=NAME:Image_8_Description CONTENT=Street1");
                            macro.AppendLine(@"TAG POS=9 TYPE=INPUT:BUTTON FORM=ACTION:*# ATTR=NAME:toggler*");
                        }
                        else if (fileName.ToLower().Contains("street2"))
                        {
                            GenerateVersions(fileName);
                            string fileToUpload = fileName.Replace(".jpg", "") + "_large.jpg";
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ACTION:*# ATTR=NAME:user_9_Image CONTENT=" + fileToUpload.Replace(" ", "<SP>"));
                            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ACTION:*# ATTR=NAME:Image_9_Description CONTENT=Street2");
                            macro.AppendLine(@"TAG POS=10 TYPE=INPUT:BUTTON FORM=ACTION:*# ATTR=NAME:toggler*");
                        }
                    }
                }



                string macroCode = macro.ToString();
                status = iim2.iimPlayCode(macroCode, 120);
            }

            if (currentUrl.ToLower().Contains("equi-trax"))
            {

                macro.AppendLine(@"");
                macro.AppendLine(@"FRAME NAME=main");
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:b_sv");
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:dmsg_close_btn");

                macro.AppendLine(@"");
                macro.AppendLine(@"FRAME NAME=main");
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:b_ap");
                macro.AppendLine(@"FRAME NAME=iFileMan");
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=CLASS:s_button");
                macro.AppendLine(@"FRAME NAME=main");
                macro.AppendLine(@"TAG POS=1 TYPE=IFRAME ATTR=ID:iFileMan");
                macro.AppendLine(@"FRAME NAME=iFileMan");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f18 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S1.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f17 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S2.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f16 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\S3.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f15 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L1.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f14 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L2.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=NAME:FileUploadForm ATTR=NAME:FileContent_f14 CONTENT=" + SubjectFilePath.Replace(" ", "<SP>") + @"\L3.jpg");
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:HMSG");
                macro.AppendLine(@"TAG POS=18 TYPE=A ATTR=CLASS:s_button");
                macro.AppendLine(@"TAG POS=1 TYPE=NOBR ATTR=TXT:Save<SP>Changes");
                macro.AppendLine(@"FRAME NAME=main");
                macro.AppendLine(@"TAG POS=1 TYPE=A ATTR=ID:b_ManageFiles_close");
                macro.AppendLine(@"TAG POS=11 TYPE=A ATTR=CLASS:s_button");



                string macroCode = macro.ToString();
                status = iim2.iimPlayCode(macroCode, 30);
            }

            #region dispo

            if (currentUrl.ToLower().Contains("dispo"))
            {
                Dispo bpoform = new Dispo();
                bpoform.UploadPics(this, iim2);
            }

            #endregion


            //if (currentUrl.ToLower().Contains("rapidclose"))

            if (iim2.iimGetLastExtract().ToLower().Contains("rapidclose") || iim.iimGetLastExtract().ToLower().Contains("rapidclose"))
            {

                
                string[] fileEntries = Directory.GetFiles(search_address_textbox.Text);

                foreach (string fileName in fileEntries)
                {
                    
                    if (Regex.IsMatch(fileName, @"\\.*address.*\.jpg"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=ID:updateOrderForm ATTR=VALUE:Upload<SP>Image");
                        macro.AppendLine(@"TAB T=2");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:uploadImageForm ATTR=ID:uploadedFile CONTENT=" + fileName.Replace(" ", "<SP>"));
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:uploadImageForm ATTR=ID:SLIType CONTENT=%2");
                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:uploadImageForm ATTR=NAME:description CONTENT=" + Regex.Match(fileName, @"(address.*.jpg)"));
                        macro.AppendLine(@"TAG POS=1 TYPE=BUTTON:SUBMIT FORM=ID:uploadImageForm ATTR=CLASS:button");
                        //macro.AppendLine(@"TAB T=1");
                        //macro.AppendLine(@"TAB CLOSE");

                    }
                    else if (Regex.IsMatch(fileName, @"\\.*front.*\.jpg"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=ID:updateOrderForm ATTR=VALUE:Upload<SP>Image");
                        macro.AppendLine(@"TAB T=2");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:uploadImageForm ATTR=ID:uploadedFile CONTENT=" + fileName.Replace(" ", "<SP>"));
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:uploadImageForm ATTR=ID:SLIType CONTENT=%1");
                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:uploadImageForm ATTR=NAME:description CONTENT=" + Regex.Match(fileName, @"(front.*.jpg)"));
                        macro.AppendLine(@"TAG POS=1 TYPE=BUTTON:SUBMIT FORM=ID:uploadImageForm ATTR=CLASS:button");
                        // macro.AppendLine(@"TAB T=1");
                        // macro.AppendLine(@"TAB CLOSE");
                    }
                    else if (Regex.IsMatch(fileName, @"\\.*left.*\.jpg"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=ID:updateOrderForm ATTR=VALUE:Upload<SP>Image");
                        macro.AppendLine(@"TAB T=2");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:uploadImageForm ATTR=ID:uploadedFile CONTENT=" + fileName.Replace(" ", "<SP>"));
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:uploadImageForm ATTR=ID:SLIType CONTENT=%4");
                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:uploadImageForm ATTR=NAME:description CONTENT=" + Regex.Match(fileName, @"(left.*.jpg)"));
                        macro.AppendLine(@"TAG POS=1 TYPE=BUTTON:SUBMIT FORM=ID:uploadImageForm ATTR=CLASS:button");
                        //macro.AppendLine(@"TAB T=1");
                        //macro.AppendLine(@"TAB CLOSE");
                    }
                    else if (Regex.IsMatch(fileName, @"\\.*right.*\.jpg"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=ID:updateOrderForm ATTR=VALUE:Upload<SP>Image");
                        macro.AppendLine(@"TAB T=2");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:uploadImageForm ATTR=ID:uploadedFile CONTENT=" + fileName.Replace(" ", "<SP>"));
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:uploadImageForm ATTR=ID:SLIType CONTENT=%5");
                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:uploadImageForm ATTR=NAME:description CONTENT=" + Regex.Match(fileName, @"(right.*.jpg)"));
                        macro.AppendLine(@"TAG POS=1 TYPE=BUTTON:SUBMIT FORM=ID:uploadImageForm ATTR=CLASS:button");
                        //macro.AppendLine(@"TAB T=1");
                        //macro.AppendLine(@"TAB CLOSE");
                    }
                    else if (Regex.IsMatch(fileName, @"\\.*street.*\.jpg"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=ID:updateOrderForm ATTR=VALUE:Upload<SP>Image");
                        macro.AppendLine(@"TAB T=2");
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:uploadImageForm ATTR=ID:uploadedFile CONTENT=" + fileName.Replace(" ", "<SP>"));
                        macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:uploadImageForm ATTR=ID:SLIType CONTENT=%8");
                        macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:uploadImageForm ATTR=NAME:description CONTENT=" + Regex.Match(fileName, @"(street.*.jpg)"));
                        macro.AppendLine(@"TAG POS=1 TYPE=BUTTON:SUBMIT FORM=ID:uploadImageForm ATTR=CLASS:button");
                        //macro.AppendLine(@"TAB T=1");
                        //macro.AppendLine(@"TAB CLOSE");
                    }
                }
                //  ProcessFile(fileName);
                // status = iim.iimPlayCode(macroCode, timeout);


                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=ID:updateOrderForm ATTR=VALUE:Upload<SP>Image");
                //macro.AppendLine(@"'New tab opened");
                //macro.AppendLine(@"TAB T=2");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:uploadImageForm ATTR=ID:uploadedFile CONTENT=C:\fakepath\angle.27.jpg");
                //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:uploadImageForm ATTR=ID:SLIType CONTENT=%4");
                //macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:uploadImageForm ATTR=ID:description CONTENT=Angle");
                //macro.AppendLine(@"TAG POS=1 TYPE=BUTTON:SUBMIT FORM=ID:uploadImageForm ATTR=CLASS:button");
                //macro.AppendLine(@"TAB T=1");
                //macro.AppendLine(@"TAB CLOSE");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=ID:updateOrderForm ATTR=VALUE:Upload<SP>Image");
                //macro.AppendLine(@"'New tab opened");
                //macro.AppendLine(@"TAB T=2");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:uploadImageForm ATTR=ID:uploadedFile CONTENT=C:\fakepath\front.02.jpg");
                //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:uploadImageForm ATTR=ID:SLIType CONTENT=%1");
                //macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:uploadImageForm ATTR=ID:description CONTENT=Front");
                //macro.AppendLine(@"TAG POS=1 TYPE=BUTTON:SUBMIT FORM=ID:uploadImageForm ATTR=CLASS:button");
                //macro.AppendLine(@"TAB T=1");
                //macro.AppendLine(@"TAB CLOSE");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=ID:updateOrderForm ATTR=VALUE:Upload<SP>Image");
                //macro.AppendLine(@"'New tab opened");
                //macro.AppendLine(@"TAB T=2");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:uploadImageForm ATTR=ID:uploadedFile CONTENT=C:\fakepath\left-side.56.jpg");
                //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:uploadImageForm ATTR=ID:SLIType CONTENT=%4");
                //macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:uploadImageForm ATTR=ID:description CONTENT=Left<SP>side<SP>of<SP>building");
                //macro.AppendLine(@"TAG POS=1 TYPE=BUTTON:SUBMIT FORM=ID:uploadImageForm ATTR=CLASS:button");
                //macro.AppendLine(@"TAB T=1");
                //macro.AppendLine(@"TAB CLOSE");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=ID:updateOrderForm ATTR=VALUE:Upload<SP>Image");
                //macro.AppendLine(@"'New tab opened");
                //macro.AppendLine(@"TAB T=2");
                //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:uploadImageForm ATTR=ID:SLIType CONTENT=$--select--");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:uploadImageForm ATTR=ID:uploadedFile CONTENT=C:\fakepath\right-back.07.jpg");
                //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:uploadImageForm ATTR=ID:SLIType CONTENT=%7");
                //macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:uploadImageForm ATTR=ID:description CONTENT=Rear-right<SP>half<SP>building");
                //macro.AppendLine(@"TAG POS=1 TYPE=BUTTON:SUBMIT FORM=ID:uploadImageForm ATTR=CLASS:button");
                //macro.AppendLine(@"TAB T=1");
                //macro.AppendLine(@"TAB CLOSE");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:BUTTON FORM=ID:updateOrderForm ATTR=VALUE:Upload<SP>Image");
                //macro.AppendLine(@"'New tab opened");
                //macro.AppendLine(@"TAB T=2");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:uploadImageForm ATTR=ID:uploadedFile CONTENT=C:\fakepath\street.49.jpg");
                //macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:uploadImageForm ATTR=ID:description CONTENT=Street");
                //macro.AppendLine(@"TAG POS=1 TYPE=TD FORM=ID:uploadImageForm ATTR=TXT:AddCancel");
                //macro.AppendLine(@"TAG POS=1 TYPE=BUTTON:SUBMIT FORM=ID:uploadImageForm ATTR=CLASS:button");
                //macro.AppendLine(@"TAG POS=1 TYPE=BUTTON:SUBMIT ATTR=CLASS:dijitButtonNode");
                //macro.AppendLine(@"'TAG POS=0 TYPE=BUTTON:SUBMIT ATTR=CLASS:dijitButtonNode");
                //macro.AppendLine(@"TAG POS=1 TYPE=BUTTON:SUBMIT FORM=ID:uploadImageForm ATTR=CLASS:button");
                //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:uploadImageForm ATTR=ID:SLIType CONTENT=%8");
                //macro.AppendLine(@"TAG POS=1 TYPE=BUTTON:SUBMIT FORM=ID:uploadImageForm ATTR=CLASS:button");
                //macro.AppendLine(@"TAB T=1");
                //macro.AppendLine(@"TAB CLOSE");
                string macroCode = macro.ToString();
                if (iim2.iimGetLastExtract().ToLower().Contains("rapidclose") )
                {
                    status = iim2.iimPlayCode(macroCode, 30);
                }else
                {
                    status = iim.iimPlayCode(macroCode, 30);
                }

               
            }
            else if (currentUrl.ToLower().Contains("propertysmart"))
            {

                macro.AppendLine(@"FRAME NAME=_MAIN");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:Subject1 CONTENT=%Subject");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:Description1 CONTENT=%Other");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:other1 CONTENT=Ceiling");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:Form1 ATTR=ID:File_Uploader1_File_Upload CONTENT=" + search_address_textbox + "\\ceiling.jpg");
                macro.AppendLine(@"ONDIALOG POS=1 BUTTON=YES");
                macro.AppendLine(@"ONDIALOG POS=2 BUTTON=NO");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:Subject2 CONTENT=%Subject");
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:Form1 ATTR=ID:Description2 CONTENT=%Other");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:other2 CONTENT=dec2");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:Form1 ATTR=ID:other1 CONTENT=desc1");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:Form1 ATTR=ID:File_Uploader1_File_Upload1 CONTENT=C:\fakepath\bath-1.jpg");
                string macroCode = macro.ToString();


                string[] fileEntries = Directory.GetFiles(search_address_textbox.Text);

                foreach (string fileName in fileEntries)
                {

                }
                //  ProcessFile(fileName);
                // status = iim.iimPlayCode(macroCode, timeout);
            }
        }
        private void button11_Click(object sender, EventArgs e)
        {
            //get the full path of the images
            string image1Path = Path.Combine(@"C:\Users\Scott\Documents\My Dropbox\BPOs\69 GLEN EAGLES CT\", "SAM_0093.jpg");
            string image2Path = Path.Combine(@"C:\Users\Scott\Documents\My Dropbox\BPOs\69 GLEN EAGLES CT\", "SAM_0094.jpg");

            //compare the two
            //Console.Write("Comparing: " + bmp1 + " and " + bmp2 + ", with a threshold of " + threshold);
            Bitmap firstBmp = (Bitmap)System.Drawing.Image.FromFile(image1Path);
            Bitmap secondBmp = (Bitmap)System.Drawing.Image.FromFile(image2Path);
            //save the diffgram
            firstBmp.GetDifferenceImage(secondBmp, true).Save(image1Path + "_diff.png");
            MessageBox.Show((string.Format("Difference: {0:0.0} %", firstBmp.PercentageDifference(secondBmp, 3) * 100)));
            //Console.WriteLine("ENTER see histogram for ");
            //Console.ReadLine();
            //Console.WriteLine();
            //Console.WriteLine("Creating histogram for ");
            Histogram hist = new Histogram(firstBmp);
            hist.Visualize().Save(image1Path + "_hist.png");
            //Console.WriteLine(hist.ToString());
            //Console.WriteLine("ENTER to continue...");
            //Console.ReadLine();
        }

        private void button_save_comp_pics_Click(object sender, EventArgs e)
        {

            ////
            ////Save pics
            ////
            //StringBuilder save_pics_macro = new StringBuilder();

            ////string filepath = "C:\\Users\\Scott\\Documents\\My<SP>Dropbox\\BPOs\\" + search_address_textbox.Text;
            //string filepath = search_address_textbox.Text;
            //string filename = "comp.jpg";

            //if (sale_or_list_flag == "sale")
            //{
            //    save_pics_macro.AppendLine(@"ONDOWNLOAD FOLDER=" + filepath.Replace(" ", "<SP>") + " FILE=S" + closed_comps.ToString() + ".jpg");
            //}
            //else
            //{
            //    save_pics_macro.AppendLine(@"ONDOWNLOAD FOLDER=" + filepath.Replace(" ", "<SP>") + " FILE=L" + active_comps.ToString() + ".jpg");
            //}
            //save_pics_macro.AppendLine(@"FRAME NAME=workspace");
            //save_pics_macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:dc ATTR=SRC:*/PICS/*");
            ////save_pics_macro.AppendLine(@"'New tab opened");
            //save_pics_macro.AppendLine(@"TAB T=2");
            //save_pics_macro.AppendLine(@"FRAME F=0");
            //save_pics_macro.AppendLine(@"TAG POS=2 TYPE=IMG FORM=NAME:dc ATTR=HREF:""*.JPEG"" CONTENT=EVENT:SAVEITEM");
            //save_pics_macro.AppendLine(@"TAB CLOSE");

            //string save_pics_macroCode = save_pics_macro.ToString();
            //status = iim.iimPlayCode(save_pics_macroCode, 30);
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void search_address_textbox_TextChanged(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"VERSION BUILD=8021952");
            macro.AppendLine(@"TAB T=1");
            macro.AppendLine(@"FRAME NAME=subheader");
            macro.AppendLine(@"TAG POS=1 TYPE=IMG FORM=NAME:header ATTR=ID:favorites");
            macro.AppendLine(@"'New tab opened");
            macro.AppendLine(@"TAB T=2");
            macro.AppendLine(@"FRAME F=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=NAME:dc ATTR=ID:selHeadingnewHeading&&VALUE:newHeading CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:dc ATTR=ID:enterNewHeading CONTENT=521<SP>Flossmoor<SP>Ave,<SP>Waukegan");
            macro.AppendLine(@"'TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:dc ATTR=ID:Add&&VALUE:Add");
            macro.AppendLine(@"");
            macro.AppendLine(@"'TAB CLOSE");
            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 30);
        }

        private void button14_Click(object sender, EventArgs e)
        {

            //subject page
            ServiceLinkSubjectPage sp = new ServiceLinkSubjectPage(iim2);

            Dictionary<string, string> commonFields = new Dictionary<string, string>() {
                 {"location", "s"}, 
                 {"propType", "cc"}, 
                 {"attDet", "a"} 
             };


            sp.FillPage(commonFields);


            // ServiceLinkInterface sd = new ServiceLinkInterface(iim2);



            //sd.Location(ServiceLinkInterface.locationType.Suburban);
            //sd.PropertyType(ServiceLinkInterface.propertyType.Condo);

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }
    }
}
