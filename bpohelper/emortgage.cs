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

    class EmortgageForm
    {
        public string formName = "";


    }


    class Emortgage
    {
       

        private EmortgageForm formType;

        public Emortgage()
        {
            formType = new EmortgageForm();
            formType.formName = "e"; 
        }
        public Emortgage(string formLetter)
        {
            formType = new EmortgageForm();
            formType.formName = formLetter;
        }

        private void WriteScript(string path, string filename, StringBuilder script)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(path + "\\" + filename))
            {
                file.Write(script);
            }


        }

        public string BathString(MLSListing m)
        {
            string s = "None";

      

            return s;

        }

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
                switch (formType.formName)
                {
                    case "z":
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

        public void Prefill(iMacros.App iim, Form1 form)
        {
            string imacroCommand = "";
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");
            macro.AppendLine(@"FRAME NAME=main");

          
            if (formType.formName == "e")
            {
                //
                //Selection Boxes - in order, left to right, start from top left
                //

                //sPropStyle --> translation logic section

                //sPropIntInsp - interior inspected?
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropIntInsp");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:No");
                
                //sPropOcc - is property occupied
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropOcc");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Yes");

                //sPropOccOwn - occupied by whom
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropOccOwn");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Unknown");

                //sPropCond
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropCond");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Average");
                
                //sPropType --> translation logic section
               
                //sPropTR
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropTR");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:" + form.SubjectRoomCount);

                //sPropBR
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropBR");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:" + form.SubjectBedroomCount);

                //sPropBA --> translation logic section

                //sPropGar --> translation logic section

                //sPropHOAYesNo --> translation logic section

                //sPropSource
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropSource");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:County<SP>Tax");

                //sPropDSourceDoc - Can you provide documentation of your data source?
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropDSourceDoc");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Yes");

                //sBasement --> translation logic section

                //sGuest
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sGuest");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:No");

                //sPropFireplace --> translation logic section

                //sPropPorch --> translation logic section

                //sPropPatio --> translation logic section

                //sPropDeck --> translation logic section

                //sPropPoolList
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropPoolList");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:None");

                //sPropZone
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropZone");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Residential");

                //sPropUnit
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropUnit");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:1");

                //sIsListed --> translation logic section

                //sIsListed12 --> translation logic section

                //sPropSold   --> translation logic section

                //sPropSold3Years --> translation logic section

                //sPropCA
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropCA");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Average");

                //sPropPM
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropPM");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Average");

                //sPropLL
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropLL");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Average");

                //sNDataCN - neighborhood conformity
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sNDataCN");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Average");


                //sNDataHS - housing supply
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sNDataHS");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Increasing");

                //sNDataCV - crime and valdalism risk level
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sNDataCV");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Low<SP>Risk");

                //sNDataNT - neighborhood tread
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sNDataNT");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Stable");

                //sEconTrend
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sEconTrend");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Stable");

                //sMarketTime
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sMarketTime");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:3<SP>to<SP>6<SP>Mos.");

                //sNDataMkt
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sNDataMkt");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Appreciating");

                //sNDataEP - environmental problems
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sNDataEP");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:No");

                //sNDataREODriven
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sNDataREODriven");
                macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:No");

                //sPropLocation --> translation logic section


                
              


                //
                //TextBoxes
                //
             

                //Inspection Date
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:InspectionDate CONTENT=" + DateTime.Now.ToShortDateString());

                //PropRent - Rent
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropRent CONTENT=" + form.SubjectRent);

                //Subject year blt
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropYrBlt CONTENT=" + form.SubjectYearBuilt.Replace(",", ""));

                //Subject SF
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropSqFt CONTENT=" + form.SubjectAboveGLA.Replace(",", ""));

                //LotSize - Subject Lot Size
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:LotSize CONTENT=" + form.SubjectLotSize + "ac");

                //PropBsmtPerFin - % finished  --> translation logic section

                //Subject Parcel #
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropAPN CONTENT=" + form.SubjectPin);

                //PropSourceID - page, mls, or other id
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropSourceID CONTENT=" + form.SubjectPin);

                //PropOther
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropOther CONTENT=NA" );

                macro.AppendLine(@"FRAME NAME=main");
                //PropSDistrict
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropSDistrict CONTENT=" + form.SubjectSchoolDistrict.Replace(" ","<SP>"));

                //PropSDivision
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropSDivision CONTENT=" + form.SubjectSubdivision.Replace(" ", "<SP>"));

                //PropMLSArea
                //todo

                //PropCounty
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropCounty CONTENT=" + form.SubjectCounty);

                //PropCensusTract
                //todo


                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropBasementSF CONTENT=" + form.SubjectBasementGLA);
                ////Energy Eff Items
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropEnergyEff CONTENT=NA");
                ////View
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropView CONTENT=Residential");
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
                
                ////Subject Parcel #
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:ParcelNumber CONTENT=" + form.SubjectPin);
              
                ////Owner Occ %
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataPctOccOwn CONTENT=90");
                ////Market Price Slope
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataRateIncreased CONTENT=1");
                ////Market Price For How Long
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataMktPriceMonthsInc CONTENT=6");
                ////# Boarded
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataBoarded CONTENT=0");
                ////Subject Source ID
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropSourceID CONTENT=" + form.SubjectPin);
                ////% New Const
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataContstruct CONTENT=0");
                ////Asis Sale Value
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:ProbSaleAsIs CONTENT=" + form.SubjectMarketValue);
                ////Asis Suggested List
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:SuggValAsIs CONTENT=" + form.SubjectMarketValue);
                ////Quick Sale Value
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PR30DQS CONTENT=" + form.SubjectQuickSaleValue);
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:MktValQSAsIs CONTENT=" + form.SubjectQuickSaleValue);
                ////Quick List
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:SuggListQSAsIs CONTENT=" + form.SubjectQuickSaleValue);
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:SuggListQSRep CONTENT=" + form.SubjectQuickSaleValue);
                ////Quick Sale Repaird
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:MktValQSRep CONTENT=" + form.SubjectQuickSaleValue);
                ////Repaired Value
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:ProbSaleRep CONTENT=" + form.SubjectMarketValue);
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:MktValRep CONTENT=" + form.SubjectMarketValue);
                ////Suggested Repair List
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:SuggValRep CONTENT=" + form.SubjectMarketValue);
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:SuggListRep CONTENT=" + form.SubjectMarketValue);
                ////Land Value
                //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PRLO CONTENT=10000");
                
                ////neighborhood data
                //try
                //{
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataADOM CONTENT=" + form.SubjectNeighborhood.avgDom.ToString());
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataRVMin CONTENT=" + form.SubjectNeighborhood.minSalePrice.ToString());
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataRVMax CONTENT=" + form.SubjectNeighborhood.maxSalePrice.ToString());
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataCompSupply CONTENT=" + form.SubjectNeighborhood.numberActiveListings.ToString());
                //    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:NDataREOCorp CONTENT=" + form.SubjectNeighborhood.numberREOListings.ToString());
                //}
                //catch
                //{

                //}


                //
                //TextAreas
                //
                //Environmental Problems
                //macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPOForm ATTR=NAME:NDataEPComm CONTENT=No<SP>known<SP>issues<SP>noted<SP>during<SP>drive-by<SP>inspection.");
                //General Market Comments
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPOForm ATTR=NAME:NeighborComment CONTENT=Following<SP>is<SP>the<SP>MLS<SP>market<SP>stats<SP>within<SP>1<SP>mile<SP>radius<SP>of<SP>subject,<SP>6<SP>months<SP>in<SP>time.<BR><LF>Prices<SP>seem<SP>to<SP>have<SP>stabilized<SP>at<SP>these<SP>lower<SP>levels,<SP>but<SP>demand<SP>it<SP>still<SP>weak<SP>overall:");
                //Resale Comments
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPOForm ATTR=NAME:NDataNComm CONTENT=Subject<SP>is<SP>maintained<SP>and<SP>landscaped.<BR><LF>No<SP>adverse<SP>conditions<SP>were<SP>noted<SP>at<SP>the<SP>time<SP>of<SP>inspection<SP>based<SP>on<SP>exterior<SP>observations.<SP>");
                //BPO comments
                macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPOForm ATTR=NAME:BComment CONTENT=No<SP>known<SP>special<SP>concerns,<SP>encroachments,<SP>easements,<SP>water<SP>rights,<SP>environmental<SP>concerns,<SP>flood<SP>zones.<BR><LF>Searched<SP>a<SP>distance<SP>of<SP>at<SP>least<SP>1<SP>miles,<SP>up<SP>to<SP>6<SP>months<SP>in<SP>time.<SP>The<SP>comps<SP>bracket<SP>the<SP>subject<SP>in<SP>age,<SP>SF,<SP>and<SP>lot<SP>size,<SP>as<SP>well<SP>as<SP>use<SP>comps<SP>in<SP>same<SP>condition<SP>with<SP>similar<SP>features,<SP>and<SP>from<SP>the<SP>subjects<SP>market<SP>area.<SP>All<SP>the<SP>comps<SP>are<SP>Reasonable<SP>substitute<SP>for<SP>the<SP>subject<SP>property,<SP>similar<SP>in<SP>most<SP>areas.<SP>Price<SP>opinion<SP>was<SP>based<SP>off<SP>comparable<SP>statistics.");






                //
                //translation logic needed 
                //
                //Style sPropStyle
                macro.AppendLine(@"FRAME NAME=main");
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropStyle");
                if (form.SubjectMlsType.ToLower().Contains("1 story"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Single<SP>Story");
                }
                else if (form.SubjectMlsType.ToLower().Contains("2 stories") || form.SubjectMlsType.ToLower().Contains("townhouse"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:2-Story<SP>Conv");
                }
                else if (form.SubjectMlsType.ToLower().Contains("raised ranch"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Split/Bi-Level");
                }
                else if (form.SubjectMlsType.ToLower().Contains("split") || form.SubjectMlsType.ToLower().Contains("other"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Tri/Muilti-Level");
                }
                else if (form.SubjectMlsType.ToLower().Contains("1.5"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Cape");
                }

                //sPropFireplace
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sPropFireplace");
                if (String.IsNullOrWhiteSpace(form.SubjectNumberFireplaces))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:No");
                }
                else
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Yes");
                }

                macro.AppendLine(@"FRAME NAME=main");
            
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

                    macro.AppendLine(@"FRAME NAME=main");
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

                    macro.AppendLine(@"FRAME NAME=main");
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

                macro.AppendLine(@"FRAME NAME=main");
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

                //HOA
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
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropAMC CONTENT=" + form.SubjectAssessmentInfo.managementContact.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropHOAPhone CONTENT=" + form.SubjectAssessmentInfo.managementPhone.Replace(" ", "<SP>"));

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

                //Listed last 12
                if (form.SubjectMlsStatus.ToLower().Contains("actv") || form.SubjectMlsStatus.ToLower().Contains("ctg") || form.SubjectMlsStatus.ToLower().Contains("pend") || form.SubjectMlsStatus.ToLower().Contains("temp"))
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sIsListed");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Yes");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sListingStatus");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Still<SP>Listed");
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropOLPrice CONTENT=" + form.SubjectCurrentListPrice);

                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:PropOLDOM CONTENT=" + form.SubjectDom);
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:ListingCompany CONTENT=" + form.SubjectListingAgent.Replace(" ", "<SP>"));
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:ListingCompanyPhone CONTENT=" + form.SubjectBrokerPhone);
                    macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPOForm ATTR=NAME:NotSaleComm CONTENT=Price<SP>and<SP>location.");

                }
                else
                {
                    macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sIsListed");
                    macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:No");
                }
            }
            else
            {
                #region form_I
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

                #endregion
            }


            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 50);


        }

        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string>fieldList)
        {
            string sol;
            string formType = "e";

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
               if (field.Contains("*"))
               {
                   //drop down box
                   if (formType == "n")
                   {
                       macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:sC{0}{1}_{2}\r\n", sol, field.Replace("*", ""), Regex.Match(compNum, @"\d").Value);
                       macro.AppendLine(@"DS CMD=KEY X={{!TAGX}} Y={{!TAGY}} CONTENT=" + fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                   }
                   else
                   {
                       //sCSalesGar_2 CONTENT=2<SP>Attached
                       macro.AppendLine(@"FRAME NAME=main");
                       macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}{1}_{2}", sol, field.Replace("*", ""), Regex.Match(compNum, @"\d").Value);
                       macro.AppendLine();
                       macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:" + fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                   }
               }
               else
               {
                   //orignal way using * instead of C
                   //macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:*{0}{1}_{2} CONTENT={3}\r\n", sol, field, Regex.Match(compNum, @"\d").Value, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                   macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:C{0}{1}_{2} CONTENT={3}\r\n", sol, field, Regex.Match(compNum, @"\d").Value, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
               }



              
           }

           macro.AppendLine(@"FRAME NAME=main");
           //
           //Selection Boxes
           //
           //State
           macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}State_{1}", sol, Regex.Match(compNum, @"\d").Value);
           macro.AppendLine("");
           macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Illinois");
           //REO?
           macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}REO_{1}", sol, Regex.Match(compNum, @"\d").Value);
           macro.AppendLine("");
           macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:No");
           //Units
           macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}Units_{1}", sol, Regex.Match(compNum, @"\d").Value);
           macro.AppendLine("");
           macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:1");
           //Sales Type = Fair Market
           macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}REOCORP_{1}", sol, Regex.Match(compNum, @"\d").Value);
           macro.AppendLine("");
           macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Fair<SP>Market");
           //Pool = no
           macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}Pool_{1}", sol, Regex.Match(compNum, @"\d").Value);
           macro.AppendLine("");
           macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:No");
           //Spa = no
           macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}Spa_{1}", sol, Regex.Match(compNum, @"\d").Value);
           macro.AppendLine("");
           macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:No");
           //Porch = no
           macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}Porch_{1}", sol, Regex.Match(compNum, @"\d").Value);
           macro.AppendLine("");
           macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:No");
           //Patio = no
           macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}Patio_{1}", sol, Regex.Match(compNum, @"\d").Value);
           macro.AppendLine("");
           macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:No");
           //Deck = no
           macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}Deck_{1}", sol, Regex.Match(compNum, @"\d").Value);
           macro.AppendLine("");
           macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:No");
           //Overall Condition
           macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}Cond_{1}", sol, Regex.Match(compNum, @"\d").Value);
           macro.AppendLine("");
           macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Average");


           ////Location
           ////macro.AppendLine(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=name:sCListLoc_1");
           //macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}Loc_{1}", sol, Regex.Match(compNum, @"\d").Value);
           //macro.AppendLine("");
           //macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Good");
           ////Fee Simple
           //macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}Leasehold_{1}", sol, Regex.Match(compNum, @"\d").Value);
           //macro.AppendLine("");
           //macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Fee<SP>Simple");
           ////Quality of Construction
           //macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}QualConst_{1}", sol, Regex.Match(compNum, @"\d").Value);
           //macro.AppendLine("");
           //macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Good");
           ////Functional Utility
           //macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}FcnUtil_{1}", sol, Regex.Match(compNum, @"\d").Value);
           //macro.AppendLine("");
           //macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:Yes");
           
           //Data Source
           macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}Source_{1}", sol, Regex.Match(compNum, @"\d").Value);
           macro.AppendLine("");
           macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:MLS");
 
           ////RoomCount Corrections
           //macro.AppendFormat(@"TAG POS=1 TYPE=INPUT FORM=NAME:BPOForm ATTR=NAME:sC{0}TR_{1}", sol, Regex.Match(compNum, @"\d").Value);
           //macro.AppendLine("");
           //macro.AppendLine(@"TAG POS=1 TYPE=A FORM=NAME:BPOForm ATTR=title:" + fieldList["*TR"]);



           //
           //Text Boxes
           //

           //
           macro.AppendFormat(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:C{0}HOA_{1} CONTENT=0/mo", sol, Regex.Match(compNum, @"\d").Value);
           macro.AppendLine("");

           //Fin
           macro.AppendFormat(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:C{0}Fin_{1} CONTENT=NA", sol, Regex.Match(compNum, @"\d").Value);
           macro.AppendLine("");

           //View
           macro.AppendFormat(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:C{0}View_{1} CONTENT=Residential", sol, Regex.Match(compNum, @"\d").Value);
           macro.AppendLine("");



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
            //
            //TBD
            //
             //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListOLDate_1");
             //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListCLPrice_1");
             //  macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:BPOForm ATTR=NAME:CListCLDate_1");

           WriteScript(fieldList["filepath"], sol + compNum + ".iim", macro);

        
            string macroCode = macro.ToString();
             iim.iimPlayCode(macroCode, 60);
        }
    }
}
