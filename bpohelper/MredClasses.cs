using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using System.Windows.Forms;


namespace bpohelper
{
    public class MREDHtmlField
    {
        public MREDHtmlField(string h)
        {
            htmlLabel = h;
        }

         public MREDHtmlField(string h, MLSListing.ListingSheetFieldName k )
        {
            htmlLabel = h;
            MREDFieldName = k;
        }
        public string htmlLabel;
        public string value;
        string displayName;
        public MLSListing.ListingSheetFieldName MREDFieldName;
    }

    

         [System.Xml.Serialization.XmlInclude(typeof(DetachedListing))]
         [System.Xml.Serialization.XmlInclude(typeof(AttachedListing))]
        public class MLSListing
        {
             public enum ListingSheetFieldName
             {
                 Undefined, MSLNumber, TypeDetached, TypeAttached, Taxes
             };

            protected HtmlAgilityPack.HtmlDocument doc;
            internal Dictionary<ListingSheetFieldName, MREDHtmlField> Fields = new Dictionary<ListingSheetFieldName, MREDHtmlField>()
            {
                {ListingSheetFieldName.TypeDetached,  new MREDHtmlField("Type:")}, 
                {ListingSheetFieldName.TypeAttached,  new MREDHtmlField("Type:")}, 
            };

            #region mlsHtmlFields
            internal Dictionary<string, MREDHtmlField> mlsHtmlFields = new Dictionary<string, MREDHtmlField>()
            {
                
                 {"address", new MREDHtmlField("Address:")},
              
                 {"additionalSalesInfo", new MREDHtmlField("Addl.&nbsp;Sales&nbsp;Info.:")}, 
                 
                 {"acerage", new MREDHtmlField("Acreage:")}, 
                 {"assessmentIncludes", new MREDHtmlField("Asmt&nbsp;Incl:")},
                 {"assessmentAmount", new MREDHtmlField("Amount:")},
                 {"mlsGla", new MREDHtmlField("Appx&nbsp;SF:")},
                 {"basement", new MREDHtmlField("Basement:")}, 
                 {"basementDetails", new MREDHtmlField("Basement&nbsp;Details:")}, 
                 
                 {"bathrooms", new MREDHtmlField(@"Bathrooms (full/half):")},
                 {"bedrooms", new MREDHtmlField("Bedrooms:")}, 
                 {"broker", new MREDHtmlField("Broker:")}, 
                 
                 {"builtBefore78", new MREDHtmlField(@"Blt&nbsp;Before&nbsp;78:")},
                 {"closedDate", new MREDHtmlField("Closed:")},
                 {"contingency", new MREDHtmlField("Contingency:")},
                 {"managementContactName", new MREDHtmlField("Contact&nbsp;Name:")},
                 {"managementContactPhone", new MREDHtmlField("Phone:")},
                 {"contractDate", new MREDHtmlField("Contract:")},
                 {"currentlyLeased", new MREDHtmlField("Curr.&nbsp;Leased:")},
                 {"daysOnMarket", new MREDHtmlField("Lst.&nbsp;Mkt.&nbsp;Time:")},
                 {"assessmentFrequency", new MREDHtmlField("Frequency:")},
                 {"garageType", new MREDHtmlField("Garage&nbsp;Type:")}, 
                 {"lotDimensions", new MREDHtmlField("Dimensions:")}, 
                 {"drivingDirections", new MREDHtmlField("Directions:")},

                 {"financing", new MREDHtmlField("Financing:")},
                 {"mlsNumber", new MREDHtmlField("MLS&nbsp;#:")},
                 {"listAgent", new MREDHtmlField("List&nbsp;Agent:")}, 
                 
                 {"listDate", new MREDHtmlField("List&nbsp;Date:")},
                 {"listPrice", new MREDHtmlField("List&nbsp;Price:")},
                 {"openingBid", new MREDHtmlField("Opening Bid:")},
                 {"mlsAgeRange", new MREDHtmlField("Age:")},
                 {"mlsArea", new MREDHtmlField("Area:")},
                 {"mlsStyle", new MREDHtmlField("Style:")},
                 {"mlsType", new MREDHtmlField("Type:", ListingSheetFieldName.TypeDetached)},
                 {"numFireplaces", new MREDHtmlField("#&nbsp;Fireplaces:")},
                {"numSpaces", new MREDHtmlField("#&nbsp;Spaces:")}, 
                  {"offMarketDate", new MREDHtmlField("Off&nbsp;Market:")},
                 {"origListPrice", new MREDHtmlField("Orig&nbsp;List&nbsp;Price:")},
                 {"ownershipType", new MREDHtmlField("Ownership:")}, 
                 {"parking", new MREDHtmlField("Parking:")},
                 {"pin", new MREDHtmlField("PIN:")}, 
                 {"phone", new MREDHtmlField("Ph&nbsp;#:")}, 
                 {"points", new MREDHtmlField("Points:")},
                 {"remarks", new MREDHtmlField("Remarks:")},
                 {"roomCount", new MREDHtmlField("Rooms:")},
                 {"status", new MREDHtmlField("Status:")},
                 {"soldBy", new MREDHtmlField("Sold&nbsp;by:")},
                 {"soldPrice", new MREDHtmlField("Sold&nbsp;Price:")},
                 {"yearBulit", new MREDHtmlField("Year&nbsp;Built:")},
                 {"subdivision", new MREDHtmlField("Subdivision:")},
                 {"waterFront", new MREDHtmlField("Waterfront:")}            
            };
#endregion 

             public  MLSListing()
             {
                 
             }

            public  MLSListing(string htmlCode)
            {

                rawData = htmlCode;
                doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(rawData);
                //SetMlsFields();
            }

             public void ReloadValues()
             {
                 doc = new HtmlAgilityPack.HtmlDocument();
                 doc.LoadHtml(rawData);
                 SetMlsFields();
             }


            virtual protected void SetAboveGradeLevels()
            {
                mlsMainLevelRooms = Regex.Matches(rawData, @"class=value>Main Level</TD>").Count;
                mls2ndLevelRooms = Regex.Matches(rawData, @"class=value>2nd Level</TD>").Count;
            }

             public string PropertyType() 
            {
                //var queryLondonCustomers = from field in mlsHtmlFields
                //           where field.Value.Key ==  ListingSheetFieldName.TypeDetached
                //           select field;

                return mlsHtmlFields.FirstOrDefault(field => field.Value.MREDFieldName == ListingSheetFieldName.TypeDetached).Value.value;
 
            }

            protected void SetParkingParameters()
            {
                Int32.TryParse(Regex.Match(mlsHtmlFields["numSpaces"].value.TrimEnd(), @"Gar.(\d+)").Groups[1].Value, out numberOfGarageStalls);
                if (mlsHtmlFields["garageType"].value.ToLower().Contains("none") || String.IsNullOrWhiteSpace(mlsHtmlFields["garageType"].value))
                {
                    attachedGarage = false;
                    detachedGarage = false;
                    garageExists = false;
                    numberOfGarageStalls = 0;
                } 
                else if (mlsHtmlFields["garageType"].value.ToLower().Contains("attached"))
                {
                    attachedGarage = true;
                    detachedGarage = false;
                    garageExists = true;

                } 
                else if (mlsHtmlFields["garageType"].value.ToLower().Contains("detached"))
                {
                    attachedGarage = false;
                    detachedGarage = true;
                    garageExists = true;
                } 
            }

            public string BasementGLA()
            {
                return calculatedBasementGLA.ToString();
            }

            public string BasementFinishedGLA()
            {
                return calculatedBasementFinishedGLA.ToString();
            }

            public string  BasementFinishedPercentage()
            {
                return calculatedBasementFinishedPercentage.ToString();
            }

            public bool GarageExsists()
            {
                return garageExists;
            }
            public bool DetachedGarage()
            {
                return detachedGarage;
            }
            public bool AttachedGarage()
            {
                return attachedGarage;
            }

            public int intNumberGarageStalls()
            {
                return numberOfGarageStalls;
            }

            public string NumberGarageStalls()
            {
                return numberOfGarageStalls.ToString();
            }

            public string GarageType()
            {
                return mlsHtmlFields["garageType"].value;
            }


            protected void SetBasementProperties()
            {
                //<TD style="" class=value>Lower</TD>
                //<TD style="" class=value>Basement</TD>
                //<TD style="" class=value>Walkout Basement</TD>
                //<TD class=value>Lower</TD>

                //mlsLowerLevelRooms = rawData.Count(p => p.Equals(@"<TD style="" class=value>Lower</TD>"));
                mlsLowerLevelRooms = Regex.Matches(rawData, @"class=value>Lower</TD>").Count;
                mlsBasementRooms = Regex.Matches(rawData, @"class=value>Basement</TD>").Count;
                mlsWalkoutBasementRooms = Regex.Matches(rawData, @"class=value>Walkout Basement</TD>").Count;

                if (mlsBasementRooms > 0 || mlsWalkoutBasementRooms > 0)
                {
                    calculatedNumberOfBasementRooms = mlsBasementRooms + mlsWalkoutBasementRooms;
                }
                else
                {
                    calculatedNumberOfBasementRooms = mlsLowerLevelRooms;
                }



                if (mlsHtmlFields["basement"].value.ToLower().Contains("none") || String.IsNullOrWhiteSpace(mlsHtmlFields["basement"].value))
                {
                    calculatedBasementGLA = 0;
                    calculatedNumberOfBasementRooms = 0;
                }

                if (mlsHtmlFields["basement"].value.ToLower().Contains("full") | mlsHtmlFields["basement"].value.ToLower().Contains("english"))
                {
                    calculatedBasementGLA = intMlsGla / basementGLADivisionFactor;
                }

                if (mlsHtmlFields["basement"].value.ToLower().Contains("partial"))
                {
                    calculatedBasementGLA = intMlsGla / (basementGLADivisionFactor * 2);
                }

                if (mlsHtmlFields["basementDetails"].value.ToLower().Contains("unfinished") || String.IsNullOrWhiteSpace(mlsHtmlFields["basementDetails"].value))
                {
                    calculatedNumberOfBasementRooms = 0;
                    calculatedBasementFinishedPercentage = 0;
                    fullFinishedBasement = false;
                    partiallyFinishedBasement = false;
                    finishedBasement = false;


                }
                else if (mlsHtmlFields["basementDetails"].value.ToLower().Contains("partially"))
                {

                    calculatedBasementFinishedPercentage = calculatedNumberOfBasementRooms / intMlsTotalRooms;
                    fullFinishedBasement = false;
                    partiallyFinishedBasement = true;
                    finishedBasement = true;


                }
                else if (mlsHtmlFields["basementDetails"].value.ToLower().Contains("finished"))
                {
                    calculatedBasementFinishedPercentage = 100;
                    fullFinishedBasement = true;
                    partiallyFinishedBasement = false;
                    finishedBasement = true;
                }


            }

            internal void SetMlsFields()
            {
              //old way 
              //  mlsNumber = Regex.Match(rawData, @"<TD class=Label>MLS&nbsp;#:</TD>  <TD class=value>(\d+)</TD>").Groups[1].Value;
               //
               // IEnumerable<HtmlNode> columns = doc.DocumentNode.Descendants("td").Where(x => x.GetAttributeValue("class", "").Equals("Label"));
               // IEnumerable<HtmlNode> mynodes = doc.DocumentNode.Descendants("td").Where(x => x.InnerText.Equals("MLS&nbsp;#:"));
                //mlsNumber = doc.DocumentNode.Descendants("td").First(x => x.InnerText.Equals("MLS&nbsp;#:")).NextSibling.NextSibling.InnerText;
                foreach (string key in mlsHtmlFields.Keys)
                {
                    try
                    {
                        mlsHtmlFields[key].value = doc.DocumentNode.Descendants("td").First(x => x.InnerText.Equals((mlsHtmlFields[key]).htmlLabel)).NextSibling.NextSibling.InnerText.Replace("&nbsp;", " ");
                    }
                    catch
                    {
                        try
                        {
                            mlsHtmlFields[key].value = doc.DocumentNode.Descendants("span").First(x => x.InnerText.Equals((mlsHtmlFields[key]).htmlLabel)).NextSibling.NextSibling.InnerText.Replace("&nbsp;", " ");
                        }
                        catch
                        {
                            try
                            {
                                mlsHtmlFields[key].value = doc.DocumentNode.Descendants("span").First(x => x.InnerText.Equals((mlsHtmlFields[key]).htmlLabel)).NextSibling.InnerText.Replace("&nbsp;", " ");
                            }
                            catch
                            {
                               // MessageBox.Show("Field Not Found: " + mlsHtmlFields[key].htmlLabel);
                            }
                        }
                    }
                }
                try
                {
                    Int32.TryParse(mlsHtmlFields["mlsGla"].value, out intMlsGla);
                }
                catch
                {
                    intMlsGla = -1;
                }
                try
                {
                    Int32.TryParse(mlsHtmlFields["roomCount"].value, out intMlsTotalRooms);
                }
                catch
                {
                    intMlsTotalRooms = -1;
                }

               

                SetAboveGradeLevels();
                SetBasementProperties();
                SetParkingParameters();

            }


            private void SetListPrice()
            {
                listPrice = doc.DocumentNode.Descendants("td").First(x => x.InnerText.Equals("MLS&nbsp;#:")).NextSibling.NextSibling.InnerText;
            }

            public double proximityToSubject;
            public bool compareableToSubject = false;
            public string rawData;



            //required per MRED rules Revision Date: March 1, 2012
            private string mlsNumber;
            private string mlsArea;
            private string pin;
            private string streetNumber;
            private string streetName;
            private string streetSuffix;
            private string zip;
            private string city;
            private string state;
            private string county;
            private string township;
            private string corporateLimits;
            private string listPrice;
            private string listDate;
            private string expDate;
            private string listingAgentId;
            private string listingOfficeId;
            private string directions;
            private string elementrySchoolDistrict;
            private string jrMiddleSchoolDistrict;
            private string highSchoolDistrict;
            private string ownershipType;
            private string approxYearBuilt;
            private bool builtBefore1978;
            private string recentRehab;
            private bool waterfront;
            private InteriorFeatures interiorFeatures;
             private string mlstax;
            protected int mlsMainLevelRooms = 0;
            protected int mls2ndLevelRooms = 0;

            protected int mlsBasementRooms = 0;
            protected int mlsLowerLevelRooms = 0;
            protected int mlsWalkoutBasementRooms = 0;

            //non-mls helper fields
            protected int numberOfAboveGradeLevels = 1;
            protected int calculatedBasementGLA;
            protected int calculatedBasementFinishedGLA;
            protected int calculatedBasementFinishedPercentage;
            protected int basementGLADivisionFactor = 0;
            protected bool fullFinishedBasement = false;
            protected bool partiallyFinishedBasement = false;
            protected bool finishedBasement = false;
            protected int intMlsGla = 0;
            protected int calculatedNumberOfBasementRooms = 0;
            protected int calculatedNumberOfMainLevelRooms = 0;
            protected int calculatedNumberOf2ndLevelRooms = 0;
            protected int intMlsTotalRooms = 0;
            protected bool attachedGarage = false;
            protected bool detachedGarage = false;
            protected bool garageExists = false;
            protected int numberOfGarageStalls = 0;
            protected string mlsPropType = "Unk";
            private string realistGla = "-1";
            private string realistLotsize = "-1";

            //public string MlsTaxAmount
            //{
            //    //<TD class=Label>Amount:</TD>  <TD class=value>$21,384</TD>
            //   // get { return Regex.Matches(rawData, @"Amount:\s*.(\d+.*\d+.\d*)")[1].Value; }
            //    set {  mlstax = value; }
            //}

            public string Status
            {
                get { return mlsHtmlFields["status"].value; }
               // set { mlsNumber = value; }
            }

            public int YearBuilt
            {
                get
                {
                    int x = -1;

                    Int32.TryParse(mlsHtmlFields["yearBulit"].value, out x);
                    return x;
                }
                set { mlsHtmlFields["yearBulit"].value = value.ToString(); }
            }
            
             public int RealistGLA
            {
                get
                {
                    int x = -1;
                    
                    Int32.TryParse(realistGla, out x);
                    return x;
                }
                set { realistGla = value.ToString(); }
            }

             public int GLA
            {
                get 
                { 
                    int x = -1;
                    string s = mlsHtmlFields["mlsGla"].value;
                    Int32.TryParse(s, out x);
                    return x;

                }
                set { mlsHtmlFields["mlsGla"].value = value.ToString(); }
            }
            public string MlsNumber
            {
                get { return mlsHtmlFields["mlsNumber"].value; }
                set { mlsNumber = value; }
            }

            public string Type
            {
                get { return mlsHtmlFields["mlsType"].value; }
                set { mlsPropType = value; }
            }

            public double ProximityToSubject
            {
                get { return proximityToSubject; }
                set { proximityToSubject = value; }
            }

            public int intLevels()
            {
                return numberOfAboveGradeLevels;
            }

            public string Levels()
            {
                return numberOfAboveGradeLevels.ToString();
            }

            public double Lotsize
            {
                get
                {
                    double x = -1;
                    string s = mlsHtmlFields["acerage"].value;
                    Double.TryParse(s, out x);
                    return x;

                }
                set { mlsHtmlFields["acerage"].value = value.ToString(); }
            }

            public double RealistLotSize
            {
                get
                {
                    double x = -1;
                  
                    Double.TryParse(realistLotsize, out x);
                    return x;

                }
                set { realistLotsize = value.ToString(); }
            }

            public double SalePrice
            {
                get
                {
                    double x = -1;
                    string s = Regex.Match(mlsHtmlFields["soldPrice"].value, @"\d*,*\d+,.\d*").Value;
                    Double.TryParse(s, out x);
                    return x;

                }
                set { mlsHtmlFields["soldPrice"].value = value.ToString(); }
            }

            public double CurrentListPrice
            {
                get
                {
                    string s;
                    double x = -1;
                    try
                    {
                         s = Regex.Match(mlsHtmlFields["listPrice"].value, @"\d*,*\d+,.\d*").Value;
                    }
                    catch
                    {
                         s = Regex.Match(mlsHtmlFields["openingBid"].value, @"\d*,*\d+,.\d*").Value;
                    }
                    Double.TryParse(s, out x);
                    return x;

                }
                set { mlsHtmlFields["listPrice"].value = value.ToString(); }
            }

            public DateTime SalesDate
            {
                get
                {
                    DateTime x;
                    string s = mlsHtmlFields["closedDate"].value;
                    DateTime.TryParse(s, out x);
                    return x;

                }
                set { mlsHtmlFields["closedDate"].value = value.ToString(); }
            }

        }

        public class AttachedListing : MLSListing
        {
            //required per MRED rules Revision Date: March 1, 2012
            private bool isListedForRent;
            private string unitNumber;
            private string numberDaysForBoardApproval;
            private bool petsAllowed;
            private string typeAttached;
            private string mlsNumberOfStories;

            public AttachedListing()
            {

            }

            public AttachedListing(string s)
                : base(s)
            {
                mlsHtmlFields.Remove("bathrooms");
                mlsHtmlFields.Remove("mlsStyle");
                mlsHtmlFields.Remove("additionalSalesInfo");
                //mlsHtmlFields.Remove("acerage");  /*some attached units have lot sizes in realist, ie duplexes.  */

                mlsHtmlFields["additionalSalesInfo"] = new MREDHtmlField("Additional&nbsp;Sales&nbsp;Information:"); 
                mlsHtmlFields["bathrooms"] = new MREDHtmlField(@"Bathrooms (Full/Half):");
                mlsHtmlFields["offMarketDate"] = new MREDHtmlField(@"Off&nbsp;Mkt:");
                mlsHtmlFields["numberOfStories"] = new MREDHtmlField("#&nbsp;Stories:");
                

               // rawData = rawData + @"<TR>  <TD class=Label>Acerage:</TD>  <TD class=value>03/19/2012</TD> </TR>";

                mlsPropType = "Attached Single";

                SetMlsFields();
            }

            public string MlsPropertyType()
            {
                return mlsPropType;
            }
            //above grade levels
            protected override void SetAboveGradeLevels()
            {

                numberOfAboveGradeLevels = 1;
                basementGLADivisionFactor = 0;
                mlsNumberOfStories = mlsHtmlFields["numberOfStories"].value;

                if (mlsNumberOfStories == "1" && mlsHtmlFields["mlsType"].value.ToLower().Contains("ranch"))
                {
                    numberOfAboveGradeLevels = 1;
                    basementGLADivisionFactor = 1;


                }
                else if (mlsNumberOfStories == "2" && (mlsHtmlFields["mlsType"].value.ToLower().Contains("2 stories") || mlsHtmlFields["mlsType"].value.ToLower().Contains("2 story")))
                {
                    numberOfAboveGradeLevels = 2;
                    basementGLADivisionFactor = 2;
                }
                else
                {
                    // if main, 2nd and lower level with a basement, most likely a mistake
                    if (rawData.Contains("Main Level") && rawData.Contains("2nd Level") && rawData.Contains("Lower") && mlsHtmlFields["basement"].value.ToLower() != "none")
                    {
                        //3 levels, 2 below grade
                        //most likely a trilevel
                        basementGLADivisionFactor = 3;
                        numberOfAboveGradeLevels = 1;
                    }

                    if (rawData.Contains("Main Level") && rawData.Contains("2nd Level") && rawData.Contains("Lower") && mlsHtmlFields["basement"].value.ToLower() != "full")
                    {
                        //2 levels and basement
                        //most likely 2 story

                        numberOfAboveGradeLevels = 2;
                        basementGLADivisionFactor = 2;
                    }

                    // if it has a main, lower AND a basement
                    if (rawData.Contains("Main Level") && !rawData.Contains("2nd Level") && rawData.Contains("Lower") && mlsHtmlFields["basement"].value.ToLower() != "none")
                    {
                        //1 level, 
                        //most likely a bilevel or raised ranch
                        numberOfAboveGradeLevels = 1;
                        basementGLADivisionFactor = 1;
                    }

                    //catch-all when listing indicates there is a basement, but other listing data may be incorrect
                    if (basementGLADivisionFactor == 0 && mlsHtmlFields["basement"].value.ToLower() != "none")
                    {
                        basementGLADivisionFactor = 1;
                    }
                }
            }
        }

        public class DetachedListing : MLSListing
        {
            public DetachedListing()
            {

            }
            static private TypeDetachedList tl = new TypeDetachedList();  
            public DetachedListing(string s)
                : base(s)
            {
                mlsPropType = "Detached Single";
                //mlsHtmlFields["bathrooms"] = new MREDHtmlField(@"Bathrooms (full/half):");
                SetMlsFields();
            }

             public string MlsPropertyType()
            {
                return mlsPropType;
            }

             protected override void SetAboveGradeLevels()
             {
                 base.SetAboveGradeLevels();
                 switch (mlsHtmlFields["mlsType"].value)
                 {
                     case "1 Story" :
                     case "Raised Ranch":
                     case "Coach House":
                     case "Hillside":
                         numberOfAboveGradeLevels = 1;
                         basementGLADivisionFactor = 1;
                         break;
                     case "2 Stories":
                         numberOfAboveGradeLevels = 2;
                         basementGLADivisionFactor = 2;
                         break;
                     case "3 Stories":
                         numberOfAboveGradeLevels = 3;
                         basementGLADivisionFactor = 3;
                         break;
                     case "4+ Stories":
                         numberOfAboveGradeLevels = 4;
                         basementGLADivisionFactor = 4;
                         break;
                     case "1.5 Story":
                         numberOfAboveGradeLevels = 2;
                         basementGLADivisionFactor = 3;
                         break;
                     case "Split Level":
                         numberOfAboveGradeLevels = 1;
                         basementGLADivisionFactor = 3;
                         break;
                     case @"Split Level w/Sub":
                          numberOfAboveGradeLevels = 1;
                         basementGLADivisionFactor = 2;
                         break;
                     case "Other":
                     case "Tear Down":
                     default :
                          numberOfAboveGradeLevels = 1;
                          basementGLADivisionFactor = -1;
                          break;

                 }

                 // if main, 2nd and lower level with a basement, most likely a mistake
                 if (rawData.Contains("Main Level") && rawData.Contains("2nd Level") && rawData.Contains("Lower") && mlsHtmlFields["basement"].value.ToLower() != "none")
                 {
                     //3 levels, 2 below grade
                     //most likely a trilevel
                     basementGLADivisionFactor = 3;
                     numberOfAboveGradeLevels = 1;
                 }


                 if (rawData.Contains("Main Level") && rawData.Contains("2nd Level") && rawData.Contains("Lower") && mlsHtmlFields["basement"].value.ToLower() != "full")
                 {
                     //2 levels and basement
                     //most likely 2 story

                     numberOfAboveGradeLevels = 2;
                     basementGLADivisionFactor = 2;
                 }

                 // if it has a main, lower AND a basement
                 if (rawData.Contains("Main Level") && !rawData.Contains("2nd Level") && rawData.Contains("Lower") && mlsHtmlFields["basement"].value.ToLower() != "none")
                 {
                     //1 level, 
                     //most likely a bilevel or raised ranch
                     numberOfAboveGradeLevels = 1;
                     basementGLADivisionFactor = 1;
                 }
                       
             }

          
        }

        public class Parking
        {
            //Garage=Interior Parking; Parking=Exterior Parking
            //NOTE: Both Garage (Interior Parking) & Exterior Space may be selected.
            //Depending upon data entered in this field, either Garage Detail fields or
            //Exterior Parking Detail fields will display, or both.
            private Dictionary<string, string> parking = new Dictionary<string, string>()
            {
                 {"G", "Garage(s)"}, 
                 {"S", "Exterior Space(s)"},
                 {"N", "None"}
            };
            private bool parkingIncludedInPrice;
           
        }

        public class GarageParking : Parking
        {
            private Dictionary<string, string> garageOwnership = new Dictionary<string, string>()
            {
                 {"A", "Owned"}, 
                 {"B", "Transferrable Lease"},
                 {"C", "Deeded Sold Separately"},
                 {"D", @"Fee/Leased"}, 
                 {"N", @"N/A"}
            };
            private bool garageOnSite;
            private Dictionary<string, string> garageType = new Dictionary<string, string>()
            {
                 {"A", "Attached"}, 
                 {"B", "Detached"},
                 {"C", "None"},
            };
        }

        public class OwnershipTypeList
        {
            public Dictionary<string, string> mlsOwnerShipType;
            public OwnershipTypeList()
            {
                mlsOwnerShipType = new Dictionary<string, string>()
                 {
                  {"CD", "Condo"}, 
                      {"CO", "Co-op"},
                  {"FS", "Fee Simple"},
                 {"HA", "Fee Simple with H.O. Assn."}
               };
            }
       

        }

        public class StreetSuffixList
        {
            public Dictionary<string, string> mlsListingstreetSuffix;

            public StreetSuffixList()
            {

                mlsListingstreetSuffix = new Dictionary<string, string>()
        {
            {"AVE", "Avenue"}, 
            {"BLVD", "Boulevard"},
            {"CIR", "Circle"},
            {"CT", "Court"},
            {"DR", "Drive"},
            {"HWY", "Highway"},
            {"LN", "Lane"},
            {"PKWY", "Parkway"},
            {"PL", "Place"},
            {"PLZ", "Plaza"},
            {"PL", "Point"},
            {"PT", "Place"},
            {"RD", "Road"},
            {"SQ", "Square"},
            {"ST", "Road"},
            {"TER", "Terrace"},
            {"TRL", "Trail"},
            {"WAY", "Way"}
           
        };
            }
        }
    
        public class TypeAttachedList
        {
            //MRED up to 3 selections per Revision Date: March 1, 2012
            public Dictionary<string, string> mlsTypeAttached = new Dictionary<string, string>()
            {
            {"A", "1/2 Duplex"}, 
            {"B", "Cluster"},
            {"C", "Condo"},
            {"D", "Condo‐Duplex"},
            {"E", "Condo‐Loft"},
            {"F", "Corridor"},
            {"G", "Courtyard"},
            {"H", "Flat"},
            {"I", "Garden Unit"},
            {"J", "Garden Complex"},
            {"K", @"Manor Home/Coach House/Villa"},
            {"L", @"Low Rise (1‐3 Stories)"},
            {"M", @"Mid Rise (4‐6 Stories)"},
            {"N", @"High Rise (7+ Stories)"},
            {"O1", @"Quad‐Ranch"},
            {"O2", @"Quad‐Split Level"},
            {"O3", @"Quad‐2 Story"},
            {"O4", @"Quad‐Penthouse"},
            {"O", "Quad"},  //USED FOR SEARCH ONLY
            {"R", "Studio"},
            {"S", "Timeshare"},
            {"T1", @"Townhouse‐Ranch"},
            {"T2", @"Townhouse-2 Story"},
            {"T3", @"Townhouse‐ 3+ Stories"},
            {"T4", @"Townhouse‐ TriLevel"},
            {"T", "Townhouse"}, //USED FOR SEARCH ONLY
            {"U", "Vintage"},
            {"V", "Other"},
            {"W", "Ground Level Ranch"},
            {"X", "Penthouse"}
        };
        }

        public class TypeDetachedList
        {
            //MRED Maximum of 1 selection + Hillside, Earth,
            //Coach House or Tear Down, if applicable per Revision Date: March 1, 2012
            public Dictionary<string, string> mlsTypeDetached = new Dictionary<string, string>()
            {
            {"A", "1 Story"}, 
            {"B", "1.5 Story"},
            {"C", "2 Stories"},
            {"D", "3 Stories"},
            {"E", "4+ Stories"},
            {"F", "Coach House"},
            {"G", "Earth"},
            {"H", "Hillside"},
            {"I", "Raised Ranch"},
            {"J", "Split Level"},
            {"K", @"Split Level w/Sub"},
            {"L", "Other"},
            {"M", "Tear Down"}
       
        };
        }
        public class InteriorFeatures
        {
            private string approxSf;
            private int numberOfRooms;
            private int numberOfBedrooms;
            private int numberofFullBaths;
            private int numberofHalfBaths;
            

        }

        public class Basement
        {
            //MRED Basement field
         public Dictionary<string, string> mredBasementTypeList = new Dictionary<string, string>()
            {
               {"A", "Full"}, 
                {"B", "Partial"},
                {"C", "Walkout"},
                {"D", "English"},
                {"N", "None"}
            };

         public Dictionary<string, string> mredBasementDescriptionList = new Dictionary<string, string>()
            {
             {"A", "Finished"}, 
            {"B", "Partially Finished"},
            {"C", "Unfinished"},
            {"D", "Crawl"},
            {"E", "Cellar"},
            {"G", "Sub‐Basement"},
            {"H", "Slab"},
            {"I", "Exterior Access"},
            {"J", "Other"},
            {"K", "Rough‐In"},
            {"N", "None"}
            };

         public bool basementBaths;


        }

        

    

}