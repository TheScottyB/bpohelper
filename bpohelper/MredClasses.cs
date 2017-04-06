using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using System.Windows.Forms;


namespace bpohelper
{

    //  retsUsername = 'RETS_O_74601_6', // our RETS credentials for mred

//    retsPassword = 'mr8sng62yd',


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


             private string mlsStatus;


            #region mlsHtmlFields
            internal Dictionary<string, MREDHtmlField> mlsHtmlFields = new Dictionary<string, MREDHtmlField>()
            {
                
                 {"address", new MREDHtmlField("Address:")},

                 {"schoolDistrict", new MREDHtmlField("High&nbsp;School:")},
              
                 {"additionalSalesInfo", new MREDHtmlField("Addl.&nbsp;Sales&nbsp;Info.:")}, 
                 
                 {"acerage", new MREDHtmlField("Acreage:")}, 
                 {"assessmentIncludes", new MREDHtmlField("Asmt&nbsp;Incl:")},
                 {"assessmentAmount", new MREDHtmlField("Amount:")},
                 {"mlsGla", new MREDHtmlField("Appx&nbsp;SF:")},
                 {"basement", new MREDHtmlField("Basement:")}, 
                 {"basementDetails", new MREDHtmlField("Basement&nbsp;Details:")}, 
                   {"county", new MREDHtmlField(@"County:")},
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
                  {"exterior", new MREDHtmlField("Exterior:")},
                 {"financing", new MREDHtmlField("Financing:")},
                 {"mlsNumber", new MREDHtmlField("MLS&nbsp;#:")},
                 {"listAgent", new MREDHtmlField("List&nbsp;Agent:")}, 
                 {"listAgentEmail", new MREDHtmlField("Email:")}, 
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
                 {"phoneAgent", new MREDHtmlField("Ph&nbsp;#:")}, 
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

             public string DistressedSaleYesNo()
            {
                 if (TransactionType == "REO" || TransactionType == "Short Sale")
                 {
                     return "Yes";
                 }
                 return "No";
            }

             public bool DistressedSale()
             {
                 if (TransactionType == "REO" || TransactionType == "Short Sale")
                 {
                     return true;
                 }
                 return false;
             }

             public string AdditionalSalesInfo()
             {
                 return mlsHtmlFields["additionalSalesInfo"].value;
             }

            public bool ListedInLast12Months()
            {
                TimeSpan ts = DateTime.Now - ListDate;
                return (ts.TotalDays < 365);
            }

             public bool FinishedBasement
            {
                get { return finishedBasement; }
            }

            public string BasementGLA()
            {
                return Math.Abs(calculatedBasementGLA).ToString();
            }

            public string BasementFinishedGLA()
            {
                return calculatedBasementFinishedGLA.ToString();
            }

            public string  BasementFinishedPercentage()
            {
                return calculatedBasementFinishedPercentage.ToString();
            }

            public string NumberOfBasementRooms()
            {
                return mlsBasementRooms.ToString();
            }

            public string BasementType
            {
                get {  return mlsHtmlFields["basement"].value;}
            }

            public string GarageString()
            {
                string returnString = "Unknown";
                if (garageExists)
                {
                    returnString = NumberGarageStalls() + " " + mlsHtmlFields["garageType"].value;
                } 
                else
                {
                    returnString = "None";
                }
                return returnString ;
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
                if (string.IsNullOrWhiteSpace(mlsHtmlFields["garageType"].value))
                {
                    mlsHtmlFields["garageType"].value = "None";
                }

                return mlsHtmlFields["garageType"].value;
            }

            public string WaterFrontYesNo()
            {
                return mlsHtmlFields["waterFront"].value;
            }

            public bool Waterfront
            {
                get
                {
                    if (WaterFrontYesNo().ToLower() == "yes")
                    {
                        return true;
                    }

                    return false;
                }
            }

             
             public string ListingAgentPhone()
            {
                return mlsHtmlFields["phone"].value;
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
                        //mlsHtmlFields[key].value = doc.DocumentNode.Descendants("td").First(x => x.InnerText.Equals((mlsHtmlFields[key]).htmlLabel)).NextSibling.NextSibling.InnerText.Replace("&nbsp;", " ");
                        mlsHtmlFields[key].value = doc.DocumentNode.Descendants("td").First(x => x.InnerText.Equals((mlsHtmlFields[key]).htmlLabel)).NextSibling.InnerText.Replace("&nbsp;", " ");
                         if (string.IsNullOrWhiteSpace(mlsHtmlFields[key].value))
                         {
                             mlsHtmlFields[key].value = doc.DocumentNode.Descendants("td").First(x => x.InnerText.Equals((mlsHtmlFields[key]).htmlLabel)).NextSibling.NextSibling.InnerText.Replace("&nbsp;", " ");
                            // MessageBox.Show("Field Not Found: " + mlsHtmlFields[key].htmlLabel);
                         }
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
                                
                               //MessageBox.Show("Field Not Found: " + mlsHtmlFields[key].htmlLabel);
                            }
                        }
                    }
                }

                
                try
                {
                    mlsHtmlFields["listAgentEmail"].value = Regex.Matches(rawData, "Email:.*<a.*href..mailto.([^\\\"]*)")[0].Groups[1].Value;
                }
                catch
                {
                    mlsHtmlFields["listAgentEmail"].value = "";
                }

                try
                {
                    mlsHtmlFields["phoneAgent"].value = Regex.Matches(rawData, @".td class..Label..Ph.nbsp.#...td..td class..value..(.\d\d\d. \d\d\d-\d\d\d\d)..td.")[0].Groups[1].Value;
                }
                catch
                {
                    mlsHtmlFields["phoneAgent"].value = "";
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

                if (mlsHtmlFields["numSpaces"].value.Contains(" "))
                {
                    try
                    {
                        mlsHtmlFields["numSpaces"].value = Regex.Matches(rawData, @"(Gar:\d*)")[0].Value;
                    }
                    catch
                    {
                        try
                        {
                            mlsHtmlFields["numSpaces"].value = Regex.Matches(rawData, @"(Par:\d*)")[0].Value;
                        }
                        catch
                        { }
                    }


                }               

                SetAboveGradeLevels();
                SetBasementProperties();
                SetParkingParameters();
                SetAddress();
            }


            private void SetListPrice()
            {
                listPrice = doc.DocumentNode.Descendants("td").First(x => x.InnerText.Equals("MLS&nbsp;#:")).NextSibling.NextSibling.InnerText;
            }

            public double proximityToSubject;
            public bool compareableToSubject = false;
            public string rawData;



            #region private members
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
             private string geocode;
             private string listingBrokerageName;
             private double currentListPrice;
             private double origListPrice;

            #endregion

            #region protected members
             protected int mlsMainLevelRooms = 0;
            protected int mls2ndLevelRooms = 0;

            protected int mlsBasementRooms = 0;
            protected int mlsLowerLevelRooms = 0;
            protected int mlsWalkoutBasementRooms = 0;

            //non-mls helper fields
            protected string transactionType = "Arms Length";
            protected DateTime dateLastPriceChange;
            protected int numberOfPriceChanges = 0;
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

            #endregion


            //public string MlsTaxAmount
            //{
            //    //<TD class=Label>Amount:</TD>  <TD class=value>$21,384</TD>
            //   // get { return Regex.Matches(rawData, @"Amount:\s*.(\d+.*\d+.\d*)")[1].Value; }
            //    set {  mlstax = value; }
            //}

            public string ParcelNumber
            {
                get { return Regex.Match(mlsHtmlFields["pin"].value, @"\d+").Value; }
            }


            #region Address and Location Related Fields
            public string Area
            {
                get { return mlsHtmlFields["mlsArea"].value; }
            }

            protected void SetAddress()
            {
                //fuill line address
                //Address:#NEXT#2627 Sycamore Dr , Waukegan, Illinois 60085#NEXT#
                //string city = tempstrarry[1];
                //
                //With unit # as line 2
                //2805  Glacier Way,   Unit C, Wauconda, Illinois 60084
                //        0               1       2            3


                if  (string.IsNullOrWhiteSpace(mlsHtmlFields["address"].value))
                {
                    mlsHtmlFields["address"].value = Regex.Match(rawData,"<td class=\"Label\">Address:</td>.*?;\">([^<]*)").Groups[1].Value;
                }

                var temparr = mlsHtmlFields["address"].value.Split(',');

                if (temparr.Count() == 3)
                {
                    city = mlsHtmlFields["address"].value.Split(',')[1];
                    zip = mlsHtmlFields["address"].value.Split(',')[2].Split(' ')[2];
                } else
                {
                    city = mlsHtmlFields["address"].value.Split(',')[2];
                    zip = mlsHtmlFields["address"].value.Split(',')[3].Split(' ')[2];
                    mlsHtmlFields["address"].value = temparr[0] + "," + temparr[2] + "," + temparr[3];
                } 
                
              

                if (zip.Contains("-"))
                {
                    zip = zip.Remove(zip.IndexOf("-"));
                }
            }

            public string County
            {
                get { return mlsHtmlFields["county"].value; }
            }
            

            public string City
            {
                get { return city.Trim(); }
            }
            public string Zipcode
            {
                get { return zip; }
            }

            public string State
            {
                get { return "Illinois"; }
            }

            public string StreetAddress
            {
                get { return mlsHtmlFields["address"].value.Split(',')[0]; }
            }
            public string Subdivision
            {
                get 
                {
                    if (String.IsNullOrWhiteSpace(mlsHtmlFields["subdivision"].value))
                        return "NA";
                    return mlsHtmlFields["subdivision"].value; 
                }
            }
            public string GeoPointGd
            {
                get 
                { 
                  return geocode; 
                }

                set
                {
                    geocode = value;
                }
            }


            #endregion

            #region Listing and Price History, Dates, Prices and related data

             public string PointsMlsString
            {
                get
                {
                    string returnString = "Unknown";
                     if ( !string.IsNullOrWhiteSpace(mlsHtmlFields["points"].value))
                     {
                         returnString = mlsHtmlFields["points"].value;
                     }
                    return returnString;
                }
                 
            }
             public string PointsInDollars
            {
                get
                {
                    string returnString = "0";

                    if (!(PointsMlsString == "Unknown"))
                     {
                        //TODO:
                        //return dollar amount based on string
                         //returnString = mlsHtmlFields["points"].value;
                     }
                    return returnString;
                }
                 
            }

             

             public string FinancingMlsString
             {
                 get
                 {
                     string returnString = "Unknown";
                     if (!string.IsNullOrWhiteSpace(mlsHtmlFields["financing"].value))
                     {
                         returnString = mlsHtmlFields["financing"].value;
                     }
                     return returnString;
                 }

             }

            public int NumberOfPriceChanges
            {
                get { return numberOfPriceChanges; }
                set { numberOfPriceChanges = value; }
            }

            public double OriginalListPrice
            {
                get
                {
                    string s;
                    double x = -1;
                    if (origListPrice == 0)
                    {
                        try
                        {
                            s = Regex.Match(mlsHtmlFields["origListPrice"].value, @"\d*,*\d+,.\d*").Value;
                        }
                        catch
                        {
                            s = Regex.Match(mlsHtmlFields["openingBid"].value, @"\d*,*\d+,.\d*").Value;
                        }
                        Double.TryParse(s, out x);
                    }
                    else
                    {
                        x = origListPrice;
                    }
                   
                    return x;
                }
                set { origListPrice = value; }
            }

            public string ContractDate
            {
                get { return mlsHtmlFields["contractDate"].value; }
                set { mlsHtmlFields["contractDate"].value = value; }
            }

            public string TransactionType
            {
                get 

                {
                    if (string.IsNullOrWhiteSpace(mlsHtmlFields["soldPrice"].value))
                    {
                       //active listing
                        if ( mlsHtmlFields["additionalSalesInfo"].value.Contains("REO"))
                        {
                            transactionType = "REO";
                        }
                        if (mlsHtmlFields["additionalSalesInfo"].value.Contains("Short Sale"))
                        {
                            transactionType = "ShortSale";
                        }

                    }
                    else
                    {
                        //sold listing
                        if (mlsHtmlFields["soldPrice"].value.Contains("S"))
                            transactionType = "ShortSale";
                        if (mlsHtmlFields["soldPrice"].value.Contains("F"))
                            transactionType = "REO";
                        if (mlsHtmlFields["soldPrice"].value.Contains("C"))
                            transactionType = "Corp";
                    }
                    return transactionType; 
                }
                set { transactionType = value; }
            }

            public string DOM
            {
                get { return mlsHtmlFields["daysOnMarket"].value; }
                set { mlsHtmlFields["daysOnMarket"].value = value; }
            }

            public DateTime DateOfLastPriceChange
            {
                get {return dateLastPriceChange;}
                set { dateLastPriceChange = value; }
            }

            public string Status
            {
                get 
                {
                    if (String.IsNullOrWhiteSpace(mlsStatus))
                    {
                        return mlsHtmlFields["status"].value;
                    }

                    return mlsStatus;
                    
                }
                set { mlsStatus = value; }
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
                    string s = null;
                    double x = -1;
                    if (currentListPrice == 0)
                    {
                        
                        try
                        {
                            s = Regex.Match(mlsHtmlFields["listPrice"].value, @"\d*,*\d+,.\d*").Value;
                        }
                        catch
                        {
                            try
                            {
                                s = Regex.Match(mlsHtmlFields["openingBid"].value, @"\d*,*\d+,.\d*").Value;
                            }
                            catch { }
                        }
                        Double.TryParse(s, out x);
                    }
                    else
                    {
                        x = currentListPrice;
                    }
      
                    return x;

                }
                set { currentListPrice = value; }
            }


            public DateTime ListDate
            {
                get
                {
                    DateTime x;
                    string s = mlsHtmlFields["listDate"].value;
                    DateTime.TryParse(s, out x);
                    return x;

                }
                set { mlsHtmlFields["listDate"].value = value.ToString(); }
            }

            public string ListDateString
            {
                get
                {
                    DateTime x;
                    string s = mlsHtmlFields["listDate"].value;
                    DateTime.TryParse(s, out x);
                    return x.ToShortDateString();

                }
                set { mlsHtmlFields["listDate"].value = value; }
            }

            public DateTime OffMarketDate
            {
                get
                {
                    DateTime x;
                    string s = mlsHtmlFields["offMarketDate"].value;
                    DateTime.TryParse(s, out x);
                    return x;

                }
                 set { mlsHtmlFields["offMarketDate"].value = value.ToString(); }
            }

            public string OffMarketDateString
            {
                get
                {
                    DateTime x;
                    string s = mlsHtmlFields["offMarketDate"].value;
                    DateTime.TryParse(s, out x);
                    return x.ToShortDateString();

                }
                set { mlsHtmlFields["offMarketDate"].value = value; }
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

            #endregion

            #region roomcounts

            public string FullBathCount
            {
                get { return  mlsHtmlFields["bathrooms"].value[0].ToString(); }
            }
            public string HalfBathCount
            {
                get 
                {
                    string theCount = "0";
                    try
                    {
                        theCount = mlsHtmlFields["bathrooms"].value.Replace(" ", "").Replace(@"/", ".")[2].ToString();
                    }
                    catch { }

                     return theCount;
                }
            }
            public string BathroomCount
            {
                get { return mlsHtmlFields["bathrooms"].value.Replace(" ", "").Replace(@"/","."); }
            }

            public string BedroomCount
            {
                get 
                {
                    return mlsHtmlFields["bedrooms"].value[0].ToString(); 
                }
            }
             


             public int TotalRoomCount
            {
                get { return intMlsTotalRooms; }
            }

             public string YearBuiltString
             {
                 get
                 {
                     return mlsHtmlFields["yearBulit"].value;
                 }
             }

            public int YearBuilt
            {
                get
                {
                    int x = -1;

                    Int32.TryParse(mlsHtmlFields["yearBulit"].value, out x);

                    //if (x==-1)
                    //{ 
                    //    x = this.RealistGLA; 

                    //}

                    return x;
                }
                set { mlsHtmlFields["yearBulit"].value = value.ToString(); }
            }

            #endregion

            #region interior features
            public string NumberOfFireplaces
            {
                get 
                {
                    string theCount = "0";
                    try
                    {
                        theCount = mlsHtmlFields["numFireplaces"].value;
                    }
                    catch { }

                    return theCount;   
                }
            }

            #endregion

            public string SchoolDistrict
            {
                get { return mlsHtmlFields["schoolDistrict"].value; }
            }

            public string MredParkingString
            {
                get { return mlsHtmlFields["parking"].value; }
            }

            public string ListingAgentId
            {
                get { return Regex.Match(mlsHtmlFields["listAgent"].value, @"\((\d+)\)").Groups[1].Value; }
            }
            public string ListingAgentName
            {
                get { return mlsHtmlFields["listAgent"].value; }
            }
            public string ListingBrokerageId
            {
                get { return Regex.Match(mlsHtmlFields["broker"].value, @"\((\d+)\)").Groups[1].Value; }
            }
            public string ListingBrokerageName
            {
                get 
                {
                    if (String.IsNullOrWhiteSpace(listingBrokerageName))
                    {
                        listingBrokerageName = mlsHtmlFields["broker"].value;
                    }
                    return listingBrokerageName;
                }
                set { listingBrokerageName = value; }
            }
            public string ListingAgentNameEmail
            {
                get { return mlsHtmlFields["listAgentEmail"].value; }
            }
            public string ListingAgentNamePhone
            {
                get { return mlsHtmlFields["phoneAgent"].value; }
            }

            public int Age
            {
                get
                {

                    int x = -1;
                    string s = mlsHtmlFields["yearBulit"].value;
                    Int32.TryParse(s, out x);
                    if (x == -1 || x == 0)
                    { return -1; }

                  

                    DateTime myAge = new DateTime((Convert.ToInt32(x)), 1, 1);

                    TimeSpan ts = DateTime.Now - myAge;

                    return ts.Days / 365;
                }
            }


            public int RealistGLA
            {
                get
                {
                    int x = -1;

                    Int32.TryParse(this.realistGla, out x);
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
                    if (x == -1 || x == 0)
                    { x = this.RealistGLA; }
                    return x;

                }
                set { mlsHtmlFields["mlsGla"].value = value.ToString(); }
            }

           

             public string ProperGla(string subjectGla)
            {
                int correctGla = this.GLA;
                try
                {
                    int x = -1;
                    Int32.TryParse(subjectGla, out x);
                    int diffToRealist = Math.Abs(x - this.RealistGLA);
                    int diffToMls = Math.Abs(x - this.GLA);

                    if (this.RealistGLA == -1 | diffToMls <= diffToRealist)
                    {
                        correctGla = this.GLA;
                    }
                    else
                    {
                        correctGla = this.RealistGLA;
                    }

                }
                catch
                {

                }

                return correctGla.ToString();

            }

            public string MlsNumber
            {
                get { return mlsHtmlFields["mlsNumber"].value; }
                set 
                
                {
                    mlsNumber = value;
                    mlsHtmlFields["mlsNumber"].value = mlsNumber;
                }
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
                    if (x == -1)
                    {
                        x = this.RealistLotSize;
                    }
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

            public string ExteriorMlsString
            {
                get
                {
                    return  mlsHtmlFields["exterior"].value;

                }
            }

            public string Heating
            {
                get
                {
                    return Regex.Match(rawData, @"Heating:</span><span class=.value.>(.*?)</span>").Groups[1].Value;

                }
            }

            public string Cooling
            {
                get
                {
                    return Regex.Match(rawData, @"Air&nbsp;Cond:</span><span class=.value.>(.*?)</span>").Groups[1].Value;

                }
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
                                                                   //     Additional&nbsp;Sales&nbsp;Information:    
                mlsHtmlFields["additionalSalesInfo"] = new MREDHtmlField(@"Additional&nbsp;Sales&nbsp;Information:"); 
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