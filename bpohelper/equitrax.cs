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

    class Equitrax
    {
        //public Dictionary<string, string> typeField = new Dictionary<string, string>()
        //{
        //     //{"2 Unit", ""},
        //     // {"3 Unit", ""},
        //     //  {"4 Unit", ""},
        //     //   {"Vacant Land",""},
        //     //    {"Other", ""},
        //          {"1 Story", "SF Detach"},
        //     //      {,"SF Attach"},
        //     //       {"Condo", ""},
        //     //    {"Townhouse", ""},
        //     //    {"Co-Op", ""},
        //     //    {"Modular", ""},
        //     //    {"Mobile Home", ""},
        //     //   {"Multi Family", ""}
        //    // {"A", "1 Story"}, 
        //    //{"B", "1.5 Story"},
        //    //{"C", "2 Stories"},
        //    //{"D", "3 Stories"},
        //    //{"E", "4+ Stories"},
        //    //{"F", "Coach House"},
        //    //{"G", "Earth"},
        //    //{"H", "Hillside"},
        //    //{"I", "Raised Ranch"},
        //    //{"J", "Split Level"},
        //    //{"K", @"Split Level w/Sub"},
        //    //{"L", "Other"},
        //    //{"M", "Tear Down"}
        //};

        public Dictionary<string, string> formFields = new Dictionary<string, string>()
        {

        };

        public Dictionary<string, string> FieldsOnForm = new Dictionary<string, string>()
        {

        };

        private void WriteScript(string path, string filename, StringBuilder script)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(path + "\\" + filename))
            {
                file.Write(script);
            }


        }
         

        private string formType = "z";

        public string GarageString(MLSListing m)
        {
            string s = "None";
            if (m.GarageExsists())
            {
                s = m.NumberGarageStalls() + "<SP>" + m.GarageType(); 
            }
            //sCSalesGar_2 CONTENT=2<SP>Attached
            return s;

        }

        public string TypeString(MLSListing m)
        {
            string s = "SF<SP>Detach";
            
            

            if (m.PropertyType().Contains("Attached"))
            {
            switch (formType)
            {
                case "z" :
                    break;
                case "n":
                     s = "SF<SP>Attach";
                    break;
                case "x":
                    break;
            }

            }

            return s;

        }

        public string ReportType(iMacros.App iim)
        {
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"FRAME NAME=main");
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPOForm ATTR=NAME:SpecialInst EXTRACT=TXT");
            iim.iimPlayCode(macro.ToString(), 30);
            string reportType = iim.iimGetLastExtract();
            if (reportType.Contains("Fannie Mae"))
            {
                if (reportType.Contains("NREIS"))
                {
                    formType = "n";
                }
                else
                {
                    formType = "z";
                }
               
            }
            else
            {
                formType = "x";
               
            }
            return formType;
        }

        public void Prefill(iMacros.App iim, Form1 form)
        {
            StringBuilder macro = new StringBuilder();
            Dictionary<string, string> fieldList = new Dictionary<string, string>();

            //form z
            //Fannie Mae
            //status = iim.iimPlayCode(macroCode, timeout);
            //macro.AppendLine(@"FRAME NAME=main");
            //macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPOForm ATTR=NAME:SpecialInst EXTRACT=TXT");
            //iim.iimPlayCode(macro.ToString(), 30);
            //string reportType = iim.iimGetLastExtract();
            //macro.Clear();

            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");
            macro.AppendLine(@"FRAME NAME=main");

            if (this.ReportType(iim) == "z" || this.ReportType(iim) == "n")
            {
                //
                //Form Z enhancements
                //
               
                //
                //Selection Boxes
                //
                //Owner Pride
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sOwnerPrideStatus");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Excellent");
                //Landscaping
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropLL");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Good");
                //Location
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropLoc");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Good");
                //Quality of Construction
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropQualConst");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Good");
                //Fee Simple
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropLeasehold");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Fee<SP>Simple");
                //Beds
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropBR");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:" + form.SubjectBedroomCount);
                if (formType == "n")
                {
                    //Baths *form n*
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropBA");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:" + form.SubjectBathroomCount[0]);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropBAH");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:" + form.SubjectBathroomCount[2]);
                }
                //Total Rooms
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropTR");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:" + form.SubjectRoomCount);
                //Property Type
                if (form.SubjectAttached)
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropType");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:SF<SP>Attach");
                }
                else if (form.SubjectDetached)
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropType");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:SF<SP>Detach");
                }
               //Functional Utility
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropFcnUtil");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Yes");
                //Prop Cond
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropCond");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Good");
                //Occupied By
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropOccOwn");
                if (formType == "n")
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Occupied");
                }
                else
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Unknown");
                }
                //HOA?
                if (string.IsNullOrWhiteSpace(form.SubjectAssessmentInfo.amount))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropHOAYesNo");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:No");
                }
                else
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropHOAYesNo");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Yes");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropHOA CONTENT=" + form.SubjectAssessmentInfo.amount);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropHOACurrent");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Yes");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropHOAOwed CONTENT=0");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropHOATerm");
                    try
                    {
                        if (form.SubjectAssessmentInfo.frequency.ToLower().Contains("month"))
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Month");
                        }
                        else if (form.subjectAssessmentInfo.frequency.ToLower().Contains("year"))
                        {
                            macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Month");
                        }
                    }
                    catch
                    {

                    }
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropAMC CONTENT=" + form.SubjectAssessmentInfo.managementContact.Replace(" ","<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropHOAPhone CONTENT=" + form.SubjectAssessmentInfo.managementPhone.Replace(" ","<SP>"));
                   
                    if (form.subjectAssessmentInfo.includes.ToLower().Contains("insurance"))
                    {
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:BPOForm ATTR=NAME:PropHOAIns CONTENT=YES");
                    }
                    if (form.subjectAssessmentInfo.includes.ToLower().Contains("lawn"))
                    {
                         macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:BPOForm ATTR=NAME:PropHOALL CONTENT=YES");
                    }
                    if (form.subjectAssessmentInfo.includes.ToLower().Contains("pool"))
                    {
                         macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:BPOForm ATTR=NAME:PropHOAPool CONTENT=YES");
                    }
                    if (form.subjectAssessmentInfo.includes.ToLower().Contains("tennis"))
                    {
                         macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:BPOForm ATTR=NAME:PropHOATennis CONTENT=YES");
                    }
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropHOAOther CONTENT=NA");
                }
                
                //All Financing?
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropFinYesNo");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Yes");
                //Data Source
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropSource");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:County<SP>Tax");
                //Marketing Condition
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sNDataMktCond");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Stable");
                //Employment Condition
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sNDataEmpCond");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Stable");
                //Market Price Slope Direction
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sNDataMktPriceStat");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Increasing");
                //Marketing Supply
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sNDataMktSupply");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Average");
                //Value Compared to Neighborhood
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sNDataCN");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Average");
                //Recommended Listing
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropRecList");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:As-Is");
                //Likely Buyer
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropLikelyBuyer");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Owner");
                //Listed last 12
                if (form.SubjectMlsStatus.ToLower().Contains("actv")  || form.SubjectMlsStatus.ToLower().Contains("ctg")  || form.SubjectMlsStatus.ToLower().Contains("pend")  ||form.SubjectMlsStatus.ToLower().Contains("temp"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sIsListed");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Yes");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sListingStatus");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Still<SP>Listed");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropOLPrice CONTENT=" + form.SubjectCurrentListPrice);

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropOLDOM CONTENT=" + form.SubjectDom);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:ListingCompany CONTENT=" + form.SubjectListingAgent.Replace(" ","<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:ListingCompanyPhone CONTENT=" + form.SubjectBrokerPhone);
                    macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPOForm ATTR=NAME:NotSaleComm CONTENT=Price<SP>and<SP>location.");

                }
                else 
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sIsListed");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:No");
                }
              

                //
                //TextBoxes
                //
                //Subject Patio,Porch,Deck
                macro.AppendLine(@"TAG POS=2 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropExt CONTENT=Unk");
                //Subject Fence,Pool
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropFencePool CONTENT=Unk");
                //Subject Lot Size
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropSite CONTENT=" + form.SubjectLotSize + "ac");
                //Energy Eff Items
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropEnergyEff CONTENT=NA");
                //View
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropView CONTENT=Residential");
                //Heat/AC
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropAC CONTENT=GFA.C");
                //Other
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropOther CONTENT=NA");
                //Special Financing
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropFin CONTENT=NA");
                //Subject Basement SF
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropBasementSF CONTENT=" + form.SubjectBasementGLA);
                //Subject Basement Finished SF
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropBasement CONTENT=" + form.SubjectBasementFinishedGLA);
                //Subject SF
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropSqFt CONTENT=" + form.SubjectAboveGLA.Replace(",", ""));
                //Subject year blt
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropYrBlt CONTENT=" + form.SubjectYearBuilt.Replace(",", ""));
                //Subject Parcel #
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:ParcelNumber CONTENT=" + form.SubjectPin);
                //Inspection Date
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:InspectionDate CONTENT=" + DateTime.Now.ToShortDateString());
                //Owner Occ %
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataPctOccOwn CONTENT=90");
                //Market Price Slope
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataRateIncreased CONTENT=1");
                //Market Price For How Long
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataMktPriceMonthsInc CONTENT=6");
                //# Boarded
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataBoarded CONTENT=0");
                //Subject Source ID
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropSourceID CONTENT=" + form.SubjectPin);
                //% New Const
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataContstruct CONTENT=0");
                //Asis Sale Value
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:ProbSaleAsIs CONTENT=" + form.SubjectMarketValue);
                //Asis Suggested List
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:SuggValAsIs CONTENT=" + form.SubjectMarketValue);
                //Quick Sale Value
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PR30DQS CONTENT=" + form.SubjectQuickSaleValue);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:MktValQSAsIs CONTENT=" + form.SubjectQuickSaleValue);
                //Quick List
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:SuggListQSAsIs CONTENT=" + form.SubjectQuickSaleValue);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:SuggListQSRep CONTENT=" + form.SubjectQuickSaleValue);
                //Quick Sale Repaird
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:MktValQSRep CONTENT=" + form.SubjectQuickSaleValue);
                //Repaired Value
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:ProbSaleRep CONTENT=" + form.SubjectMarketValue);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:MktValRep CONTENT=" + form.SubjectMarketValue);
                //Suggested Repair List
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:SuggValRep CONTENT=" + form.SubjectMarketValue);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:SuggListRep CONTENT=" + form.SubjectMarketValue);
                //Land Value
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PRLO CONTENT=10000");
                //Rent
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:MarketRent CONTENT=" + form.SubjectRent);
                //neighborhood data
                try
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataADOM CONTENT=" + form.SubjectNeighborhood.avgDom.ToString());
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataRVMin CONTENT=" + form.SubjectNeighborhood.minSalePrice.ToString());
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataRVMax CONTENT=" + form.SubjectNeighborhood.maxSalePrice.ToString());
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataCompSupply CONTENT=" + form.SubjectNeighborhood.numberActiveListings.ToString());
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataREOCorp CONTENT=" + form.SubjectNeighborhood.numberREOListings.ToString());
                }
                catch
                {

                }


                //
                //TextAreas
                //
                //Environmental Problems
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPOForm ATTR=NAME:NDataEPComm CONTENT=" + "No<SP>known<SP>issues<SP>noted<SP>during<SP>drive-by<SP>inspection.");
                //General Market Comments
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPOForm ATTR=NAME:GeneralMarketComments CONTENT=" + "Following<SP>is<SP>the<SP>MLS<SP>market<SP>stats<SP>within<SP>1<SP>mile<SP>radius<SP>of<SP>subject,<SP>6<SP>months<SP>in<SP>time.<BR><LF>Prices seem to have stabilized, demand is higher under 150k:".Replace(" ", "<SP>"));
                //Resale Comments
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPOForm ATTR=NAME:ResaleComments CONTENT=" +  "Subject typical of the area. It is maintained and landscaped. Unable to determine interior condition due to exterior inspection only, so subject was assumed to be in average condition for this report. No adverse conditions were noted at the time of inspection based on exterior observations.".Replace(" ","<SP>"));
                //BPO comments
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPOForm ATTR=NAME:BComment CONTENT=" + "No<SP>known<SP>special<SP>concerns,<SP>encroachments,<SP>easements,<SP>water<SP>rights,<SP>environmental<SP>concerns,<SP>flood<SP>zones.<BR><LF>Searched<SP>a<SP>distance<SP>of<SP>at<SP>least<SP>1<SP>miles,<SP>up<SP>to<SP>6<SP>months<SP>in<SP>time.<SP>The<SP>comps<SP>bracket<SP>the<SP>subject<SP>in<SP>age,<SP>SF,<SP>and<SP>lot<SP>size,<SP>as<SP>well<SP>as<SP>use<SP>comps<SP>in<SP>same<SP>condition<SP>with<SP>similar<SP>features,<SP>and<SP>from<SP>the<SP>subjects<SP>market<SP>area.<SP>All<SP>the<SP>comps<SP>are<SP>Reasonable<SP>substitute<SP>for<SP>the<SP>subject<SP>property,<SP>similar<SP>in<SP>most<SP>areas.<SP>Price<SP>opinion<SP>was<SP>based<SP>off<SP>comparable<SP>statistics.");




                //
                //translation logic needed 
                //
                //Design Appeal
                if (form.SubjectMlsType.ToLower().Contains("1 story"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropAppeal");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Single<SP>Story");
                }
                else if (form.SubjectMlsType.ToLower().Contains("2 stories") || form.SubjectMlsType.ToLower().Contains("townhouse"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropAppeal");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:2-Story<SP>Conv");
                }
                else if (form.SubjectMlsType.ToLower().Contains("raised ranch"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropAppeal");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Split/Bi-Level");
                }
                else if (form.SubjectMlsType.ToLower().Contains("split") || form.SubjectMlsType.ToLower().Contains("other"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropAppeal");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Tri/Muilti-Level");
                }
                else if (form.SubjectMlsType.ToLower().Contains("1.5"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropAppeal");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Cape");
                }
                if (formType == "z")
                {
                    //Baths
                    #region bathroom
                    string tBath = form.SubjectBathroomCount;
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
                        macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropBA");
                        macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:" + tBath);
                    }
                    catch
                    {

                    }

                    #endregion
                }
                //Garage
                #region garage
                if (form.SubjectParkingType.ToLower().Contains("gar"))
                {
                    
                    string contentString = "";

                    if (form.SubjectParkingType.ToLower().Contains("att"))
                    {
                        contentString = (Regex.Match(form.SubjectParkingType, @"\d").Value + "<SP>Attached");
                    }
                    else if (form.SubjectParkingType.ToLower().Contains("det"))
                    {
                        contentString = (Regex.Match(form.SubjectParkingType, @"\d").Value + "<SP>Detached");
                    }
                    macro.AppendLine(@"WAIT SECONDS=0.5");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropGar");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:" + contentString);
                    macro.AppendLine(@"WAIT SECONDS=0.5");
                }
                else
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropGar");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:None");
                }
                #endregion
            }

            else  //default to formX
            {
                #region formx
                //
                //property info section
                //

                //Default subject comments
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPOForm ATTR=NAME:BComment CONTENT=" + GlobalVar.defaultSubjectComments.Replace(" ", "<SP>"));
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPOForm ATTR=NAME:RepComment CONTENT=" + GlobalVar.defaultRepairComments.Replace(" ", "<SP>"));



                //sPropLoc
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropLoc");

                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Good");
                macro.AppendLine(@"DS CMD=CLICK X={{!TAGX}} Y={{!TAGY}}");

                //sPropQualConst
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropQualConst");

                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Good");
                macro.AppendLine(@"DS CMD=CLICK X={{!TAGX}} Y={{!TAGY}}");

                //
                //currently unknown or unavailable data
                //

                //currently listed, default no
                //listed last 12 months, default no

                fieldList.Add("sIsListed", "n");
                fieldList.Add("sIsListed12", "n");

                //location
                fieldList.Add("sNDataAreaType", "s");

                //Property data source 
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:BPOForm ATTR=NAME:PropSourceAssessor CONTENT=NO");
                macro.AppendLine(@"DS CMD=LDOWN X={{!TAGX}} Y={{!TAGY}} CONTENT=");
                macro.AppendLine(@"WAIT SECONDS=0.015");
                macro.AppendLine(@"DS CMD=LUP X={{!TAGX}} Y={{!TAGY}} CONTENT=");
                macro.AppendLine(@"WAIT SECONDS=0.015");


                //property values
                fieldList.Add("sNDataHS", "s");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sNDataHS CONTENT=STABLE");

                //Demand/Supply:
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sNDataComp CONTENT=INBALANCE");
                fieldList.Add("sNDataComp", "i");
                //Conformity
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sNDataCN CONTENT=AVERAGE");
                fieldList.Add("sNDataCN", "a");
                //Curb Appeal:
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropCA CONTENT=AVERAGE");
                fieldList.Add("sPropCA", "a");

                //Change/Year
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataChange CONTENT=0");

                //basement
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropBsmtSqFt CONTENT=" + form.SubjectBasementGLA);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropBsmtFinSqFt CONTENT=" + form.SubjectBasementFinishedGLA);


                //
                //translation logic
                //  
                #region garage


                if (form.SubjectParkingType.ToLower().Contains("gar"))
                {
                    string contentString = "O";

                    if (form.SubjectParkingType.ToLower().Contains("1"))
                    {
                        contentString = "1";
                    }
                    else if (form.SubjectParkingType.ToLower().Contains("2"))
                    {
                        contentString = "2";
                    }
                    else if (form.SubjectParkingType.ToLower().Contains("3"))
                    {
                        contentString = "3";
                    }
                    else if (form.SubjectParkingType.ToLower().Contains("4"))
                    {
                        contentString = "4";
                    }

                    // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropGar CONTENT=" + contentString);
                    fieldList.Add("sPropGar", contentString);

                }
                else
                {
                    fieldList.Add("sPropGar", "n");
                }
                #endregion

                #region style
                string cStr = "";

                if (form.SubjectMlsType == "1 Story" || form.SubjectMlsType.Contains("Condo") || form.SubjectMlsType.Contains("Townhouse‐Ranch"))
                {
                    cStr = "1";
                }
                else if (form.SubjectMlsType == "1.5 Story")
                {
                    cStr = "11";
                }
                else if (form.SubjectMlsType == "2 Stories" || form.SubjectMlsType.Contains("Townhouse"))
                {
                    cStr = "2";
                }
                else if (form.SubjectMlsType.Contains("Raised") || form.SubjectMlsType.Contains("Split"))
                {
                    cStr = "mm";
                }
                // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropStyle  CONTENT=" + cStr);
                fieldList.Add("sPropStyle", cStr);
                #endregion

                #region bathroom
                string tBath = form.SubjectBathroomCount;
                try
                {
                    //if there is a half bath, assuming /d./d pattern, hence the try block, incase /d format 
                    if (Convert.ToInt16(tBath[2].ToString()) > 0)
                    {
                        tBath = tBath[0].ToString() + tBath[0].ToString();

                    }
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropBA CONTENT=" + tBath);
                    fieldList.Add("sPropBA", tBath);
                }
                catch
                {

                }

                #endregion


                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropSubdivision CONTENT=" + form.SubjectSubdivision.Replace(" ", "<SP>"));

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropAPN CONTENT=" + form.SubjectPin);

                if (form.SubjectAttached)
                {
                    // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropType CONTENT=Condo");
                    fieldList.Add("sPropType", "c");
                }
                else
                {
                    //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropType CONTENT=Single<SP>Family");
                    fieldList.Add("sPropType", "s");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropHOA CONTENT=0/mo");
                }

                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropPool CONTENT=None");
                fieldList.Add("sPropPool", "n");

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropGLASqFt CONTENT=" + form.SubjectAboveGLA.Replace(",", ""));

                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropBR CONTENT=" + form.SubjectBedroomCount);
                fieldList.Add("sPropBR", form.SubjectBedroomCount);




                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:LotSize CONTENT=" + form.SubjectLotSize + "ac");
                // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropTR CONTENT=" + form.SubjectRoomCount);


                fieldList.Add("sPropTR", form.SubjectRoomCount);



                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropBsmtPerFin");
                //macro.AppendLine(@"TAG POS=7 TYPE=A FORM=NAME:BPOForm ATTR=CLASS:sbutt");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropGar");
                //macro.AppendLine(@"TAG POS=4 TYPE=A ATTR=CLASS:sitem");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropYrBlt CONTENT=" + form.SubjectYearBuilt);

                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sIsListed");

                //       macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sIsListed12");


                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropSDistrict CONTENT=" + form.SubjectSchoolDistrict.Replace(" ", "<SP>"));




                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropCond CONTENT=Average");
                fieldList.Add("sPropCond", "a");

                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropConstruct CONTENT=Average");
                fieldList.Add("PropConstruct", "a");

                // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropLoc CONTENT=Residential");
                fieldList.Add("sPropLoc", "r");


                // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropOccOwn CONTENT=Owner");
                fieldList.Add("sPropOccOwn", "o");













                // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropExtAdd CONTENT=na");

                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropSource CONTENT=County<SP>Tax");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropDSourceDoc CONTENT=Yes");
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
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropIntPlumbing CONTENT=0");

                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropIntElectrical CONTENT=0");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropExtGarage CONTENT=0");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropExtTrim CONTENT=0");

                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropRecRep CONTENT=No");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sNDataCN CONTENT=No");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropLL CONTENT=No");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sNDataEP CONTENT=No");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sNDataHS CONTENT=Stable");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sNDataCV CONTENT=5-10%");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sNDataNT CONTENT=Owner");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sNDataMktST CONTENT=Mixed");
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sPropOcc CONTENT=AsIs");
                //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:FinanceCost CONTENT=0");

                //
                //Brokers Market Analysis:
                //
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:ProbSaleAsIs CONTENT=" + form.SubjectMarketValue);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:ProbSaleMkt CONTENT=" + form.SubjectMarketValue);
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:SuggValAsIs CONTENT=" + form.SubjectMarketValue);

                foreach (string field in fieldList.Keys)
                {
                    macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:*{0}\r\n", field);
                    macro.AppendLine(@"DS CMD=KEY X={{!TAGX}} Y={{!TAGY}} CONTENT=" + fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                }

                #endregion
            }


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
            macro.AppendLine(@"SET !REPLAYSPEED FAST");

           

           foreach (string field in fieldList.Keys)
           {
               if (field.Contains("*"))
               {
                   //drop down box
                   if (formType == "x")
                   {
                       macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sC{0}{1}_{2}\r\n", sol, field.Replace("*", ""), Regex.Match(compNum, @"\d").Value);
                       macro.AppendLine(@"DS CMD=KEY X={{!TAGX}} Y={{!TAGY}} CONTENT=" + fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                   }
                   else
                   {
                       //sCSalesGar_2 CONTENT=2<SP>Attached
                       macro.AppendLine(@"FRAME NAME=main");
                       //macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}{1}_{2}", sol, field.Replace("*", ""), Regex.Match(compNum, @"\d").Value);
                       //macro.AppendLine();
                       //macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:" + fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>") );
                       macro.AppendFormat(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sC{0}{1}_{2} CONTENT={3}", sol, field.Replace("*", ""), Regex.Match(compNum, @"\d").Value, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                       macro.AppendLine();
                   }
               }
               else
               {
                    
                   macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:C{0}{1}_{2} CONTENT={3}\r\n", sol, field, Regex.Match(compNum, @"\d").Value, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
               }
               
             //  macro.AppendLine(@"DS CMD=CLICK X={{!TAGX}} Y={{!TAGY}}");
           }
            //
            //TBD
            //
             //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListOLDate_1");
             //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListCLPrice_1");
             //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListCLDate_1");

           if (formType == "z" || formType == "n")
           {
               //
               //Defaults
               //
               macro.AppendLine(@"FRAME NAME=main"); 
               //
               //Selection Boxes
               //
               //State
               macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}State_{1}", sol, Regex.Match(compNum, @"\d").Value);
               macro.AppendLine("");
               macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Illinois");
               //REO?
               macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}REOCorp_{1}", sol, Regex.Match(compNum, @"\d").Value);
               macro.AppendLine("");
               macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:No");
               //Location
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sCListLoc_1");
               macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}Loc_{1}", sol, Regex.Match(compNum, @"\d").Value);
               macro.AppendLine("");
               macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Good");
               //Fee Simple
               macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}Leasehold_{1}", sol, Regex.Match(compNum, @"\d").Value);
               macro.AppendLine("");
               macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Fee<SP>Simple");
               //Quality of Construction
               macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}QualConst_{1}", sol, Regex.Match(compNum, @"\d").Value);
               macro.AppendLine("");
               macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Good");
               //Functional Utility
               macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}FcnUtil_{1}", sol, Regex.Match(compNum, @"\d").Value);
               macro.AppendLine("");
               macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Yes");
               //Overall Condition
               macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}Cond_{1}", sol, Regex.Match(compNum, @"\d").Value);
               macro.AppendLine("");
               macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Good");
               //Data Source
               macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}Source_{1}", sol, Regex.Match(compNum, @"\d").Value);
               macro.AppendLine("");
               macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:MLS");
               //Design/Appeal Corrections
               macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}Appeal_{1}", sol, Regex.Match(compNum, @"\d").Value);
               macro.AppendLine("");
               macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:" + fieldList["*Appeal"]);
               //RoomCount Corrections
               macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}TR_{1}", sol, Regex.Match(compNum, @"\d").Value);
               macro.AppendLine("");
               macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:" + fieldList["*TR"]);



               //
               //Text Boxes
               //
               //View

               macro.AppendFormat(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:C{0}View_{1} CONTENT=Residential", sol, Regex.Match(compNum, @"\d").Value);
               macro.AppendLine("");
               macro.AppendLine(@"FRAME NAME=main");
               macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=TXT:Get<SP>Proximity");

               macro.AppendLine(@"FRAME NAME=main");
               macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=TXT:Get<SP>Proximity");

               //Translation Logic
               //Design Appeal

            
               ////Subject Patio,Porch,Deck
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropExt CONTENT=Unk");
               ////Subject Fence,Pool
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropFencePool CONTENT=Unk");
               ////Subject Lot Size
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropSite CONTENT=" + form.SubjectLotSize);
               ////Energy Eff Items
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropEnergyEff CONTENT=NA");

               ////Heat/AC
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropAC CONTENT=GFA.C");
               ////Other
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropOther CONTENT=NA");
               ////Special Financing
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropFin CONTENT=NA");
               ////Subject Basement SF
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropBasementSF CONTENT=" + form.SubjectBasementGLA);
               ////Subject Basement Finished SF
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropBasement CONTENT=" + form.SubjectBasementFinishedGLA);
               ////Subject SF
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropSqFt CONTENT=" + form.SubjectAboveGLA.Replace(",", ""));
               ////Subject Parcel #
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:ParcelNumber CONTENT=" + form.SubjectPin);
               ////Inspection Date
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropFin CONTENT=9/7/2013");
               ////Owner Occ %
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataPctOccOwn CONTENT=90");
               ////Market Price Slope
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataRat* CONTENT=1");
               ////Market Price For How Long
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataMktPriceMonth* CONTENT=6");
               ////# Boarded
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataBoarded CONTENT=0");
               ////% New Const
               //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataContstruct CONTENT=0");


             

           }

           WriteScript(fieldList["filepath"], sol + compNum + ".iim", macro);
            string macroCode = macro.ToString();
            
             
             iim.iimPlayCode(macroCode, 60);
        }
    }
}
