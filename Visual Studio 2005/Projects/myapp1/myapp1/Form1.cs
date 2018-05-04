using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using WatiN.Core;

namespace myapp1
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public Form1()
        {
            InitializeComponent();
            Datetoruntextbox.Text = "04/11/2009";
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ArrayList list = new ArrayList();


            app_testDataSetTableAdapters.expiredsTableAdapter ta = new app_testDataSetTableAdapters.expiredsTableAdapter();
            //app_testDataSet2TableAdapters.expiredsTableAdap ta = new app_testDataSetTableAdapters.expiredsTableAdapter();
            //okDataSetTableAdapters.expiredsTableAdapter ta = new okDataSetTableAdapters.expiredsTableAdapter();
            
            // Open a new Internet Explorer window and
            // goto the google website.
            IE ie = new IE("http://mredllc.connectmls.com/");

            



            ie.TextField(Find.ByName("userid")).TypeText("55426");
            ie.TextField(Find.ByName("password")).TypeText("shade123");


            // Click the Google search button.
            ie.Button(Find.By("type", "submit")).Click();




            Frame targetFrame = ie.Frame("header");
            LinkCollection lc = targetFrame.Links;


            lc[5].Click(); //search


            System.Threading.Thread.Sleep(2000);


            targetFrame = ie.Frame("main");


            FrameCollection fc = targetFrame.Frames;




            Frame savedsearchesframe = fc[2].Frame("savedsearches");




            lc = savedsearchesframe.Links;









            foreach (Link l in lc)
            {
                if (l.InnerHtml == "daily exp")
                {
                    l.Click();
                    break;
                }
            }


            ie.Frame(Find.ByName("main")).Frame(Find.ByName("workspace")).TextField(Find.ByName("minXD")).TypeText(Datetoruntextbox.Text);
            ie.Frame(Find.ByName("main")).Frame(Find.ByName("workspace")).TextField(Find.ByName("maxXD")).TypeText(Datetoruntextbox.Text);






            ie.Frame(Find.ByName("main")).Frame(Find.ByName("workspace")).Button(Find.ById("searchButtonTop")).Click();








            Table t = ie.Frame(Find.ByName("main")).Frame(Find.ByName("workspace")).Table(Find.ByIndex(0));






            //string[] propinfo = new string[65];


            bool header = true;
            foreach (TableRow trow in t.OwnTableRows)
            {
                //skip header
                if (!header)
                {

                    // int i = 0;
                    TableCellCollection tcc = trow.TableCells;


                    //trow.Link(Find.ById("addlInfoImg")).Click();
                    LinkCollection tlc = trow.Links;
                    string s;
                    foreach (Link l in tlc)
                    {
                       //if (l.InnerHtml.Contains("Click here to view the Listing History for this property"))
                        if (l.InnerHtml.Contains("Click here for additional information about this listing, including Tax"))
                            l.Click();
                    }

                    // row number, select, mls#, status, Street @, Prefix, Street Name, Unit#, City, Area, LP, RP, Property Type, Rooms, Photo, info


                    string zipcode = "";
                    if (tcc[10].Text.Length == 3)
                    {
                        zipcode = "60" + tcc[10].Text;
                    }
                    else if (tcc[10].Text.Length == 2)
                    {
                        zipcode = "600" + tcc[10].Text;
                    }
                    else if (tcc[10].Text.Length == 1)
                    {
                        zipcode = "6000" + tcc[10].Text;
                    }




                    string propinfo = tcc[2].Text + ", " + tcc[4].Text + ", " + tcc[5].Text + ", " + tcc[6].Text + ", " +
                         tcc[7].Text + ", " + tcc[8].Text + ", " + tcc[9].Text + ", " + tcc[10].Text + ", " + tcc[13].Text;
                      //      0         1      2      3       4       5          6          7        8     9     10   11  12            13        
                    //// row number, select, mls#, status, Street @, Prefix, Street Name, postfix, Unit#, City, Area, LP, RP, Property Type, Rooms, Photo, info

                    try
                    {
                        ta.Insert(tcc[2].Text, DateTime.Today, tcc[4].Text, tcc[5].Text, tcc[6].Text,
                             tcc[7].Text, tcc[8].Text, tcc[9].Text, zipcode, false);

                    }
                    catch (System.Data.OleDb.OleDbException ex)
                    {
                        if (ex.Message.Contains("primary key"))
                        {
                            //ignore
                        }
                    }

                    

                    
                    // foreach (TableCell c in tcc)
                    // {

                    //propinfo[i] = c.Text;


                    //i++;


                    //}
                    list.Add(propinfo);


                }

                string temp = trow.Text;
                header = false;


            }


            System.Threading.Thread.Sleep(1000);


            //System.IO.StreamWriter file = new System.IO.StreamWriter(@"c:\dailyexplist.txt");


            //foreach (string s in list)
            //    file.WriteLine(s);
            //file.Close();

        }
    }
}