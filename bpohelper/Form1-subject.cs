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
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;



namespace bpohelper
{
    public partial class Form1
    {
        private void Form1_Load(object sender, EventArgs e)
        {
            string s = Properties.Settings.Default.BPO_SandboxConnectionString1;
            //MessageBox.Show(s);

            AppDomain.CurrentDomain.SetData("DataDirectory", DropBoxFolder + @"\BPOs\");

            //MessageBox.Show(s);

            //Properties.Settings.Default.BPO_SandboxConnectionString1 = DropBoxFolder + @"\BPOs\";
            // TODO: This line of code loads data into the '_BPO_SandboxDataSet1.RawSFData' table. You can move, or remove it, as needed.
            this.rawSFDataTableAdapter.Fill(this._BPO_SandboxDataSet.RawSFData);

            // TODO: This line of code loads data into the '_BPO_SandboxDataSet1.subject' table. You can move, or remove it, as needed.
            this.subjectTableAdapter.Fill(this._BPO_SandboxDataSet.subject);
            if (!string.IsNullOrEmpty(search_address_textbox.Text))
            {
                //  this.importSubjectInfoButton(this, new EventArgs());
            }
            else
            {
                // search_address_textbox
                //everyones dropbox folder is different, hence this helper function
                //DropBoxFolder

                string baseDir = DropBoxFolder + @"\BPOs\";
                string[] dirs = Directory.GetDirectories(baseDir);
                string mostRecentAccess;
                TimeSpan ts = new TimeSpan();

                foreach (string d in dirs)
                {
                    ts = Directory.GetLastAccessTime(d) - DateTime.Now;
                    if (ts.Days > -5)
                    {
                        comboBox3.Items.Add(d);
                    }

                }
                TypeDetachedList tdl = new TypeDetachedList();

                foreach (string key in tdl.mlsTypeDetached.Keys)
                {
                    comboBox4.Items.Add(tdl.mlsTypeDetached[key]);
                }

                TypeAttachedList tal = new TypeAttachedList();
                foreach (string key in tal.mlsTypeAttached.Keys)
                {
                    comboBox4.Items.Add(tal.mlsTypeAttached[key]);
                }
            }
        }

        private void importSubjectInfoButton(object sender, EventArgs e)
        {

            RealistReport realistSubject = new RealistReport(this);
            MlsDrivers mlsSubject = new MlsDrivers(this);
            // TownshipReport subjectTownshipRecord = new TownshipReport(this);

            //string[] pages = { "", "", "", "", "", "", "", "" };

            //to do:
            //get a list of pdfs
            //
            string[] filePaths = Directory.GetFiles(search_address_textbox.Text, "*.pdf");
            string filename = "";
            string currentText = "";

            IEnumerable<System.Windows.Forms.TextBox> query1 = this.groupBox1.Controls.OfType<System.Windows.Forms.TextBox>();
            IEnumerable<System.Windows.Forms.TextBox> query2 = this.neighborhoodDataGroupBox.Controls.OfType<System.Windows.Forms.TextBox>();
            IEnumerable<System.Windows.Forms.TextBox> query3 = this.neighborhoodDataGroupBox.Controls.OfType<System.Windows.Forms.TextBox>();

            foreach (System.Windows.Forms.TextBox t in query1)
            {
                t.Clear();
            }

            foreach (System.Windows.Forms.TextBox t in query2)
            {
                t.Clear();
            }

            foreach (System.Windows.Forms.TextBox t in query3)
            {
                t.Clear();
            }

            try
            {
                bpoCommentsTextBox.Clear();
                bpoCommentsTextBox.LoadFile(SubjectFilePath + "\\bpocomments.rtf");
            }
            catch
            {

            }

            try
            {
                richTextBoxNeighborhoodComments.Clear();
                richTextBoxNeighborhoodComments.LoadFile(SubjectFilePath + "\\neighborhood-comments.rtf");
            }
            catch
            {

            }

            try
            {
                richTextBoxNeighborhoodComments.Clear();
                richTextBoxNeighborhoodComments.LoadFile(SubjectFilePath + "\\neighborhood-comments.rtf");
            }
            catch
            {

            }

    

         



            if (File.Exists(SubjectFilePath + "\\subjectinfo.txt"))
            {
                #region load data from subjectinfo file

                //load other fields not displayed
                foreach (string f in filePaths)
                {
                    if (f.ToLower().Contains("realist") || f.ToLower().Contains("property detail report"))
                    {
                        StringBuilder pages = new StringBuilder();
                        PdfReader pdfReader = new PdfReader(f);
                        ITextExtractionStrategy strat = new PdfHelper.LocationTextExtractionStrategyEx();
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            pages.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, strat));
                        }

                        //string fullReport = "";
                        //foreach (string p in pages)
                        //{
                        //    fullReport = fullReport + p;
                        //}

                        realistSubject.GetSubjectInfo(pages.ToString());
                        GlobalVar.theSubjectProperty.myRealistReport = realistSubject;
                        this.statusTextBox.AppendText("Realist Loaded...");
                    }

                    if (f.ToLower().Contains("report-") || f.ToLower().Contains("connectmls"))
                    {
                        StringBuilder pages = new StringBuilder();
                        PdfReader pdfReader = new PdfReader(f);
                        ITextExtractionStrategy strat = new PdfHelper.LocationTextExtractionStrategyEx();
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            pages.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, strat));
                        }

                        mlsSubject.GetSubjectInfo(pages.ToString());

                        this.statusTextBox.AppendText("MLS listing Loaded...");
                    }
                }

                //overwrite with what is saved; leave the rest alone
                string line = "";
                string[] splitLine;
                using (System.IO.StreamReader file = new System.IO.StreamReader(this.SubjectFilePath + "\\" + "subjectinfo.txt"))
                {
                    while (!file.EndOfStream)
                    {
                        line = file.ReadLine();
                        splitLine = line.Split(';');


                        System.Windows.Forms.Control[] c = this.Controls.Find(splitLine[0], true);

                        if (c[0] is TextBox || c[0] is ComboBox)
                        {
                            c[0].Text = splitLine[1];
                        }

                        if (c[0] is System.Windows.Forms.RadioButton)
                        {
                            System.Windows.Forms.RadioButton t = (System.Windows.Forms.RadioButton)c[0];

                            t.Checked = Convert.ToBoolean(splitLine[1]);
                        }

                        if (c[0] is System.Windows.Forms.DateTimePicker)
                        {
                            System.Windows.Forms.DateTimePicker t = (System.Windows.Forms.DateTimePicker)c[0];

                            t.Value = Convert.ToDateTime(splitLine[1]);
                        }


                    }

                }

                #endregion
            }
            else
            {
                #region Load Date from Pdf files
                foreach (string f in filePaths)
                {
                    if (f.ToLower().Contains("realist") || f.ToLower().Contains("property detail report"))
                    {
                        StringBuilder pages = new StringBuilder();
                        PdfReader pdfReader = new PdfReader(f);
                        ITextExtractionStrategy strat = new PdfHelper.LocationTextExtractionStrategyEx();
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            pages.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, strat));
                        }

                        //string fullReport = "";
                        //foreach (string p in pages)
                        //{
                        //    fullReport = fullReport + p;
                        //}

                        realistSubject.GetSubjectInfo(pages.ToString());
                        GlobalVar.theSubjectProperty.myRealistReport = realistSubject;
                        this.statusTextBox.AppendText("Realist Loaded...");
                    }

                    if (f.ToLower().Contains("report-") || f.ToLower().Contains("connectmls"))
                    {
                        StringBuilder pages = new StringBuilder();
                        PdfReader pdfReader = new PdfReader(f);
                        ITextExtractionStrategy strat = new PdfHelper.LocationTextExtractionStrategyEx();
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            pages.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, strat));
                        }

                        mlsSubject.GetSubjectInfo(pages.ToString());

                        this.statusTextBox.AppendText("MLS listing Loaded...");
                    }

                    if (f.ToLower().Contains("algonquin"))
                    {
                        StringBuilder pages = new StringBuilder();
                        PdfReader pdfReader = new PdfReader(f);
                        ITextExtractionStrategy strat = new PdfHelper.LocationTextExtractionStrategyEx();
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            pages.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, strat));
                        }

                        // AlgonquinTownshipReport ar = new AlgonquinTownshipReport(this, fullReport);
                        // subjectTownshipRecord = ar.DeepCopy();

                        subjectTownshipRecord = new AlgonquinTownshipReport(this, pages.ToString());
                        subjectBasementGlaTextbox.Text = subjectTownshipRecord.basementGla.ToString();
                        subjectFinishedBasementGlaTextBox.Text = subjectTownshipRecord.finishedBasementGla.ToString();
                        this.statusTextBox.AppendText("Assesor Record Loaded...");
                    }

                    if (f.ToLower().Contains("lakecountyil") || f.ToLower().Contains("lake county"))
                    {
                        StringBuilder pages = new StringBuilder();
                        PdfReader pdfReader = new PdfReader(f);
                        ITextExtractionStrategy strat = new PdfHelper.LocationTextExtractionStrategyEx();
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            pages.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, strat));
                        }


                        subjectTownshipRecord = new LakeCountyTownshipReport(this, pages.ToString());
                        subjectBasementGlaTextbox.Text = subjectTownshipRecord.basementGla.ToString();
                        subjectFinishedBasementGlaTextBox.Text = subjectTownshipRecord.finishedBasementGla.ToString();
                        this.statusTextBox.AppendText("Assesor Record Loaded...");
                    }

                    if (f.ToLower().Contains("mchenrytownship"))
                    {
                        StringBuilder pages = new StringBuilder();
                        PdfReader pdfReader = new PdfReader(f);
                        ITextExtractionStrategy strat = new PdfHelper.LocationTextExtractionStrategyEx();
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            pages.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, strat));
                        }


                        subjectTownshipRecord = new McHenryTownshipReport(this, pages.ToString());
                        subjectBasementGlaTextbox.Text = subjectTownshipRecord.basementGla.ToString();
                        subjectFinishedBasementGlaTextBox.Text = subjectTownshipRecord.finishedBasementGla.ToString();
                        this.statusTextBox.AppendText("Assesor Record Loaded...");
                    }

                    if (f.ToLower().Contains("elgin") && f.ToLower().Contains("township"))
                    {
                        StringBuilder pages = new StringBuilder();
                        PdfReader pdfReader = new PdfReader(f);
                        ITextExtractionStrategy strat = new PdfHelper.LocationTextExtractionStrategyEx();
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            pages.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, strat));
                        }


                        subjectTownshipRecord = new ElginTownshipReport(this, pages.ToString());
                        subjectBasementGlaTextbox.Text = subjectTownshipRecord.basementGla.ToString();
                        subjectFinishedBasementGlaTextBox.Text = subjectTownshipRecord.finishedBasementGla.ToString();
                        this.statusTextBox.AppendText("Assesor Record Loaded...");
                    }







                }


                try
                {
                    //
                    //check for missing data
                    //
                    if (subjectAboveGlaTextbox.Text == "0")
                    {
                        subjectAboveGlaTextbox.Text = subjectTownshipRecord.aboveGradeGla.ToString();
                    }

                    if (subjectYearBuiltTextbox.Text == "")
                    {
                        subjectYearBuiltTextbox.Text = subjectTownshipRecord.yearBuilt.ToString();
                    }
                }
                catch
                {
                    //no township record loaded.
                }

                subjectTownshipTextBox.Text = realistSubject.Township;
                subjectTaxAmountTextBox.Text = GlobalVar.theSubjectProperty.PropertyTax;

                #endregion
            }

            try
            {
                decimal dec;
                Decimal.TryParse(ndAvgDomTextBox.Text, out  dec);

                SubjectNeighborhood.avgDom = Decimal.ToInt32(dec);

                SubjectNeighborhood.minSalePrice = Convert.ToDecimal(ndMinSalePriceTextBox.Text);
                SubjectNeighborhood.maxSalePrice = Convert.ToDecimal(ndMaxSalePriceTextBox.Text);
                SubjectNeighborhood.medianSalePrice = Convert.ToDouble(ndMedianSalePriceTextBox.Text);

                SubjectNeighborhood.minListPrice = Convert.ToDecimal(ndMinListPriceTextBox.Text);
                SubjectNeighborhood.medianListPrice = Convert.ToDouble(ndMedianListPriceTextBox.Text);
                SubjectNeighborhood.maxListPrice = Convert.ToDecimal(ndMaxListPriceTextBox.Text);


                SubjectNeighborhood.numberActiveListings = Convert.ToInt32(ndNumberOfActiveListingTextBox.Text);
                SubjectNeighborhood.numberREOListings = Convert.ToInt32(ndNumberActiveReoListingsTextBox.Text);
                SubjectNeighborhood.numberOfShortSaleListings = Convert.ToInt32(ndNumberActiveShortListingsTextBox.Text);

                SubjectNeighborhood.numberSoldListings = Convert.ToInt32(ndNumberOfSoldListingTextBox.Text);
                SubjectNeighborhood.numberREOSales = Convert.ToInt32(ndNumberSoldReoListingsTextBox.Text);
                SubjectNeighborhood.numberShortSales = Convert.ToInt32(ndNumberSoldShortListingsTextBox.Text);


                Decimal.TryParse(cdAvgDomTextBox.Text, out  dec);

                SetOfComps.avgDom = Decimal.ToInt32(dec);

                SetOfComps.minSalePrice = Convert.ToDecimal(cdMinSalePriceTextBox.Text);
                SetOfComps.maxSalePrice = Convert.ToDecimal(cdMaxSalePriceTextBox.Text);
                SetOfComps.medianSalePrice = Convert.ToDouble(cdMedianSalePriceTextBox.Text);

                SetOfComps.minListPrice = Convert.ToDecimal(cdMinListPriceTextBox.Text);
                SetOfComps.medianListPrice = Convert.ToDouble(cdMedianListPriceTextBox.Text);
                SetOfComps.maxListPrice = Convert.ToDecimal(cdMaxListPriceTextBox.Text);


                SetOfComps.numberActiveListings = Convert.ToInt32(cdNumberOfActiveListingTextBox.Text);
                SetOfComps.numberREOListings = Convert.ToInt32(cdNumberActiveReoListingsTextBox.Text);
                SetOfComps.numberOfShortSaleListings = Convert.ToInt32(cdNumberActiveShortListingsTextBox.Text);

                SetOfComps.numberSoldListings = Convert.ToInt32(cdNumberOfSoldListingTextBox.Text);
                SetOfComps.numberREOSales = Convert.ToInt32(cdNumberSoldReoListingsTextBox.Text);
                SetOfComps.numberShortSales = Convert.ToInt32(cdNumberSoldShortListingsTextBox.Text);



                // //Decimal.TryParse(cdAvgDomTextBox.Text, out  dec);
                // //SetOfComps.avgDom = Decimal.ToInt32(dec);
                // SetOfComps.minSalePrice = Convert.ToDecimal(cdMinSalePriceTextBox.Text);
                // SetOfComps.maxSalePrice = Convert.ToDecimal(cdMaxSalePriceTextBox.Text);
                // SetOfComps.numberActiveListings = Convert.ToInt32(cdNumberOfActiveListingTextBox.Text);
                //// SetOfComps.numberREOListings = Convert.ToInt32(cdNumberActiveReoListingsTextBox.Text);
                // SetOfComps.numberSoldListings = Convert.ToInt32(cdNumberOfSoldListingTextBox.Text);


                //ndAvgDomTextBox.Text = SubjectNeighborhood.avgDom.ToString();
                //        ndMaxSalePriceTextBox.Text = SubjectNeighborhood.maxSalePrice.ToString();
                //        ndMinSalePriceTextBox.Text = SubjectNeighborhood.minSalePrice.ToString();
                //        ndNumberOfActiveListingTextBox.Text = SubjectNeighborhood.numberActiveListings.ToString();
                //        ndNumberActiveReoListingsTextBox.Text = SubjectNeighborhood.numberREOListings.ToString();
                //        ndNumberActiveShortListingsTextBox.Text = SubjectNeighborhood.numberOfShortSaleListings.ToString();
                //        ndNumberOfSoldListingTextBox.Text = SubjectNeighborhood.numberOfSales.ToString();
                //        ndNumberSoldReoListingsTextBox.Text = SubjectNeighborhood.numberREOSales.ToString();
                //        ndNumberSoldShortListingsTextBox.Text = SubjectNeighborhood.numberShortSales.ToString();


                //    cdAvgDomTextBox.Text = SetOfComps.avgDom.ToString();
                //    cdMaxSalePriceTextBox.Text = SetOfComps.maxSalePrice.ToString();
                //    cdMinSalePriceTextBox.Text = SetOfComps.minSalePrice.ToString();
                //    cdNumberOfActiveListingTextBox.Text = SetOfComps.numberActiveListings.ToString();
                //    cdNumberActiveReoListingsTextBox.Text = SetOfComps.numberREOListings.ToString();
                //    cdNumberActiveShortListingsTextBox.Text = SetOfComps.numberOfShortSaleListings.ToString();
                //    cdNumberOfSoldListingTextBox.Text = SetOfComps.numberOfSales.ToString();
                //    cdNumberSoldReoListingsTextBox.Text = SetOfComps.numberREOSales.ToString();
                //    cdNumberSoldShortListingsTextBox.Text = SetOfComps.numberShortSales.ToString();

            }
            catch (Exception ex)
            {

            }
            SubjectBrokerPhone = subjectBrokerPhoneTextBox.Text;
            SubjectListingAgent = subjectListingAgentTextBox.Text;
            SubjectCurrentListPrice = subjectCurrentListPriceTextBox.Text;
           // SubjectMlsStatus = subjectMlsStatusTextBox.Text;

            SubjectAssessmentInfo.managementContact = subjectHoaContactTextBox.Text;
            SubjectAssessmentInfo.frequency = subjectHoaFrequencyTextBox.Text;
            SubjectAssessmentInfo.includes = subjectHoaIncludesTextBox.Text;
            SubjectAssessmentInfo.managementPhone = subjectHoaPhoneTextBox.Text;
            SubjectAssessmentInfo.amount = subjectHoaTextBox.Text;
            SubjectDom = subjectDomTextBox.Text;

            if (String.IsNullOrWhiteSpace(subjectProximityToOfficeTextbox.Text))
            {
                subjectProximityToOfficeTextbox.Text = Get_Distance(subjectFullAddressTextbox.Text);
            }
            if (!string.IsNullOrEmpty(subjectFullAddressTextbox.Text))
            {
                GlobalVar.subjectPoint = Geocode(subjectFullAddressTextbox.Text);
            }
            GlobalVar.theSubjectProperty.County = SubjectCounty;
            GlobalVar.theSubjectProperty.MainForm = this;
        }
    }
}