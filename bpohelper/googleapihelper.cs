using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.Apis.Fusiontables.v1;
//using DotNetOpenAuth.OAuth2;
//using Google.Apis.Authentication;
using Google.Apis.Auth.OAuth2;
//using Google.Apis.Authentication.OAuth2;
//using Google.Apis.Authentication.OAuth2.DotNetOpenAuth;
using Google.Apis.Samples.Helper;
using Google.Apis.Util;
using System.Threading;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using Google.Apis.Datastore.v1beta2;
using System.Windows.Forms;
using Google.Apis.Services;

using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Util.Store;





namespace bpohelper
{
    class GoogleFusionTable
    {
        private string tableId;
        public string m_rowid;
        private Google.Apis.Fusiontables.v1.Data.ColumnList cl;
        private Google.Apis.Fusiontables.v1.Data.Sqlresponse lastSqlResponse;
         static bpohelper.Form1 callingForm;
        public  Dictionary<string, string> curRec = new Dictionary<string, string>();
        public Dictionary<string, string> subjectRec = new Dictionary<string, string>();
        static FusiontablesService service;
        private FusiontablesService newService;
        private string lastPin = "";

        public async Task helper_OAuthFusion()
        {
            UserCredential credential;

            credential =  await GoogleWebAuthorizationBroker.AuthorizeAsync(
              new ClientSecrets
              {
                  ClientId = "982969682733.apps.googleusercontent.com",
                  ClientSecret = "ADjXMNxzAqa35fB4Z12UORM9"
                  // APIKey ="AIzaSyDsAcyc5WOw0IqvR93VR8cVsWCHJ8ZoDq4"
              },
              new[] { FusiontablesService.Scope.Fusiontables},
              "user",
              CancellationToken.None);


            //service = new FusiontablesService(auth);
            service = new FusiontablesService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Fusion API Sample",

            });

            var tableColums = service.Column.List(tableId);
            tableColums.MaxResults = 150;
           cl = tableColums.Execute();
            //cl = tableColums.
            newService = service;

        }


         public GoogleFusionTable()
        {
            //callingForm = (bpohelper.Form1) Form1.ActiveForm;
            callingForm = GlobalVar.theSubjectProperty.MainForm;



            // Register the authenticator.
            //var provider = new NativeApplicationClient(GoogleAuthenticationServer.Description);
            //provider.ClientIdentifier = "982969682733.apps.googleusercontent.com";
            //provider.ClientSecret = "ADjXMNxzAqa35fB4Z12UORM9";
            //var auth = new OAuth2Authenticator<NativeApplicationClient>(provider, GetAuthentication);



           

              

            //service = new FusiontablesService(auth);
            //service = new FusiontablesService();
           


            //var service = new BooksService(new BaseClientService.Initializer()
            //{
            //    HttpClientInitializer = credential,
            //    ApplicationName = "Books API Sample",
            //});

            //// Create the service.
            //var service = new DiscoveryService(new BaseClientService.Initializer
            //{
            //    ApplicationName = "Discovery Sample",
            //    APIKey = "[YOUR_API_KEY_HERE]",
            //});
            
        }

        public GoogleFusionTable(string t) : this()
            
        {
            tableId = t;
            //var tableColums = service.Column.List(tableId);
            //tableColums.MaxResults = 150;
            // cl = tableColums.Execute();
            ////cl = tableColums.
            //newService = service;
        }

       

        public bool ColumnExsists(string columnName)
        {
            return cl.Items.Any(x => x.Name == columnName);
        }

        public void AddColumn(string name)
        {
            Google.Apis.Fusiontables.v1.Data.Column newCol = new Google.Apis.Fusiontables.v1.Data.Column();
            newCol.Name = name;
            newCol.Type = "STRING";
            service.Column.Insert(newCol, tableId).Execute();
            cl.Items.Add(newCol);
        }

        public void AddRecord(string sPin, string sAddress)
        {

            string sqlQuery = @"INSERT INTO " + tableId + " ('Parcel ID:', Date, Location) VALUES (" + sPin + ", '" + DateTime.Now.ToShortDateString() +"' ,'" + sAddress +"')";
            string  sqlGetQuery = @"SELECT * FROM " + tableId + " WHERE 'Parcel ID:' = '" + sPin +"'";

            callingForm.GoogleApiCall("bpohelper.GoogleFusionTable.AddRecord: Record Already Exsists?---->");
            Google.Apis.Fusiontables.v1.Data.Sqlresponse ttt = new Google.Apis.Fusiontables.v1.Data.Sqlresponse();
            try
            {
                ttt = service.Query.SqlGet(sqlGetQuery).Execute();
            }
            catch
            {
                callingForm.GoogleApiCall("Failed, Waiting to retry....");
                Thread.Sleep(1000);
                ttt = service.Query.SqlGet(sqlGetQuery).Execute();
            }
            if (ttt.Rows == null)
            {
                var theRequest = service.Query.Sql(sqlQuery);
                Google.Apis.Fusiontables.v1.Data.Sqlresponse theResponse = new Google.Apis.Fusiontables.v1.Data.Sqlresponse();
                callingForm.GoogleApiCall("No, Adding Record---->");
                try
                {
                    theResponse = theRequest.Execute();
                }
                catch
                {
                    callingForm.GoogleApiCall("Failed, Waiting to retry 2420ms....");
                    Thread.Sleep(2420);
                    try
                    {
                        var theRequest2 = service.Query.Sql(sqlQuery);
                        var theResponse2 = theRequest2.Execute();
                        theResponse = theResponse2;
                    }
                    catch
                    {
                        callingForm.GoogleApiCall("Failed, Waiting to retry 16800ms....");
                        Thread.Sleep(16800);
                        var theRequest3 = service.Query.Sql(sqlQuery);
                        var theResponse3 = theRequest3.Execute();
                        theResponse = theResponse3;
                    }
                   
                }
                    
                var rowID = theResponse.Rows[0][0];
                    callingForm.GoogleApiCall("rowid<----" + rowID.ToString());
                    m_rowid = rowID.ToString();
                    curRec["Parcel ID:"] = sPin;
                    curRec["Latitude"] = "";
                    curRec["Longitude"] = "";
              
              //  string x = rep.ToString();
               
            }
            else
            {
                if (ttt.Rows[0][0].ToString() == "")
                {
                    this.UpdateRecord(sPin, "Location", sAddress);
                    this.UpdateRecord(sPin, "Date", DateTime.Now.ToShortDateString());

                }

                var temp = ttt.Rows;
                var temp2 = temp.ToArray();
                string  temp3 = temp2[0].ToString();


                var what = ttt.Columns.Zip(ttt.Rows[0], (first, second) => new KeyValuePair<string, string>(first, second.ToString()));
                 curRec = what.ToDictionary(k => k.Key, v => v.Value );
                callingForm.GoogleApiCall("<----Yes");
            }
            RecordExists(sPin);

        }

        public void AddMlsRecord(string mlsNum, string htmlCode)
        {
            string sqlQuery = @"INSERT INTO " + tableId + " ('mlsNum', Date, htmlCode) VALUES (" + mlsNum + ", '" + DateTime.Now.ToShortDateString() + "' ,'" + htmlCode + "')";
            string sqlGetQuery = @"SELECT * FROM " + tableId + " WHERE 'mlsNum' = '" + mlsNum + "'";

            callingForm.GoogleApiCall("Record Already Exsists?---->");
            Google.Apis.Fusiontables.v1.Data.Sqlresponse ttt = new Google.Apis.Fusiontables.v1.Data.Sqlresponse();
            try
            {
                ttt = service.Query.SqlGet(sqlGetQuery).Execute();
            }
            catch
            {
                callingForm.GoogleApiCall("Failed, Waiting to retry....");
                Thread.Sleep(1000);
                ttt = service.Query.SqlGet(sqlGetQuery).Execute();
            }
            if (ttt.Rows.Count == 0)
            {
                callingForm.GoogleApiCall("No, Adding Record---->");
                service.Query.Sql(sqlQuery).Execute();

            }
            else
            {
                //if (ttt.Rows[0][0] == "")
                //{
                //    this.UpdateRecord(sPin, "Location", sAddress);
                //    this.UpdateRecord(sPin, "Date", DateTime.Now.ToShortDateString());

                //}


                //callingForm.GoogleApiCall("<----Yes");
            }
            //RecordExists(sPin);
        }

        public Dictionary<string, string> Geocode(string sAddress)
        {
            //
            //geocode address
            //ex:
            //http://maps.googleapis.com/maps/api/geocode/json?address=1600+Amphitheatre+Parkway,+Mountain+View,+CA&sensor=true_or_false


            Dictionary<string, string> lngLat = new Dictionary<string, string>()
            {
             {"Latitude", ""},
             {"Longitude", ""}
            };

            if (string.IsNullOrEmpty(curRec["Latitude"]) || string.IsNullOrEmpty(curRec["Longitude"]))
            {


                string googlestr = @" http://maps.googleapis.com/maps/api/geocode/json?address=" + sAddress.Replace(" ", "+") + "&sensor=false";

                //// Create a request for the URL. 		
                WebRequest request = WebRequest.Create(googlestr);

                //// Get the response.
                
                HttpWebResponse response;

                try
                {
                    response = (HttpWebResponse)request.GetResponse();
                } 
                catch
                {
                    try
                    {
                        HttpWebResponse response2 = (HttpWebResponse)request.GetResponse();
                        response = response2;
                    }
                    catch
                    {
                        Thread.Sleep(5000);
                        
                        try
                        {
                            HttpWebResponse response3 = (HttpWebResponse)request.GetResponse();
                            response = response3;
                        }
                        catch
                        {
                            Thread.Sleep(10000);
                           
                            try
                            {
                                HttpWebResponse response4 = (HttpWebResponse)request.GetResponse();
                                response = response4;
                            }
                            catch
                            {
                               
                                Thread.Sleep(30000);
                                HttpWebResponse response5 = (HttpWebResponse)request.GetResponse();
                                response = response5;
                            }
                          
                        }
                       
                    }
                }

                // Get the stream containing content returned by the server.
                Stream dataStream = response.GetResponseStream();
                // Open the stream using a StreamReader for easy access.
                StreamReader reader = new StreamReader(dataStream);
                // Read the content.
                string responseFromServer = reader.ReadToEnd();
                // Display the content.
                //MessageBox.Show(response.ContentLength.ToString());
                //MessageBox.Show(responseFromServer);
                ////Console.WriteLine(responseFromServer);
                //// Cleanup the streams and the response.

                reader.Close();
                dataStream.Close();
                response.Close();

                //ex response:
                //  "formatted_address" : "107 Glenwood Dr, Round Lake Beach, IL 60073, USA",
                //  "geometry" : {
                //   "location" : {
                //   "lat" : 42.3653120,
                //   "lng" : -88.08547999999999

                string pattern = @"lat. : (\d+.\d+)";
                Match match = Regex.Match(responseFromServer, pattern);
                string propertyLatitude = match.Groups[1].Value;

                pattern = @"lng. : (-\d+.\d+)";
                match = Regex.Match(responseFromServer, pattern);
                string propertyLongitude = match.Groups[1].Value;

                lngLat["Longitude"] = propertyLongitude;
                lngLat["Latitude"] = propertyLatitude;

                UpdateRecord(curRec["Parcel ID:"], lngLat);
            }
            else
            {
                lngLat["Longitude"] = curRec["Longitude"];
                lngLat["Latitude"] = curRec["Latitude"];
            }
            return lngLat;

        }

        public void UpdateRecord(string pin, string name, string value)
        {
            string rowid = "";
            string rowIdQurey = "SELECT ROWID FROM " + tableId + " WHERE 'Parcel ID:'=" + pin;
           

            var r = service.Query.Sql(rowIdQurey).Execute();
            try
            {
                rowid = r.Rows[0][0].ToString();
            }
            catch
            {
                Thread.Sleep(1600);
                r = service.Query.Sql(rowIdQurey).Execute();
                rowid = r.Rows[0][0].ToString();
            }
            string sqlQuery = @"UPDATE " + tableId + " SET '" + name + "' = '" + value + "' WHERE ROWID = '" + rowid + "'";

            service.Query.Sql(sqlQuery).Execute();



        }
        public void UpdateRecord(string pin, Dictionary<string, string> fieldListRealist)
        {
            string rowid = ""; 
            string rowIdQurey = "SELECT ROWID FROM " + tableId + " WHERE 'Parcel ID:'=" + pin;
            string sqlGetQuery = @"SELECT * FROM " + tableId + " WHERE 'Parcel ID:' = '" + pin + "'";

            

            
            string nameValuePairs = "";
            Dictionary<string, string> fieldListFusion = new Dictionary<string, string>();
            
            //we will need the rowid of the record we want to update later
            int sleepTime = 350;
          
            callingForm.GoogleApiCall("Get Row ID---->");

            //if (m_rowid.IsNullOrEmpty())
            if (string.IsNullOrEmpty(m_rowid))
            {





                var newrecord = newService.Query.SqlGet(rowIdQurey);
                var ccc = service.Query.SqlGet(rowIdQurey);
                var r = ccc.Execute();
               // var ar = service.Query.SqlGet(rowIdQurey).Execute();
                
                //r = ar;
                //string fff = newrecord.Rows[0][0];


                // r = service.Query.Sql(rowIdQurey).Execute();
                //Thread.Sleep(sleepTime+1000);
                while (r.Rows == null || r.Rows.Count < 1)
                {

                    callingForm.GoogleApiCall("Failed, Waiting to retry....");
                    Thread.Sleep(sleepTime);
                    callingForm.GoogleApiCall("Get Row ID---->");

                    r = newService.Query.Sql(rowIdQurey).Execute();
                    //Thread.Sleep(sleepTime);
                    sleepTime = sleepTime + 1000;

                }


                rowid = r.Rows[0][0].ToString();
                callingForm.GoogleApiCall("<----Row ID: " + rowid);
            }
            else
            {
                //if (curRec["Parcel ID:"] == pin)
                //{

                //}
                rowid = m_rowid;
                //m_rowid = "";

            }
            
            //catch
            //{
            //    r = service.Query.Sql(rowIdQurey).Execute();
            //    rowid = r.Rows[0][0];
            //}


            //this query will pull the current record from the fusion table based on parcel id
            string getRecSql = @"SELECT * FROM " + tableId + " WHERE 'Parcel ID:' = '" + pin + "'";
            //execute sql statement, this is the response
            callingForm.GoogleApiCall("Get Record---->");
            Google.Apis.Fusiontables.v1.Data.Sqlresponse returnRespone = new Google.Apis.Fusiontables.v1.Data.Sqlresponse();


            var returnResponeList = returnRespone.Rows;

            if (curRec["Parcel ID:"] == pin)
            {
                fieldListFusion = curRec;
            }
            else
            {

                try
                {
                    returnRespone = service.Query.Sql(getRecSql).Execute();
                }
                catch
                {
                    callingForm.GoogleApiCall("Failed, Waiting to retry....");
                    Thread.Sleep(1000);
                    returnRespone = service.Query.Sql(getRecSql).Execute();
                }
                returnResponeList = returnRespone.Rows;
                //var returnRespone = service.Query.Sql(getRecSql).Execute();
                //var returnResponeList = returnRespone.Rows;



                while (returnRespone.Rows == null || returnRespone.Rows.Count < 1)
                {
                    callingForm.GoogleApiCall("Failed, Waiting to retry....");
                    Thread.Sleep(1000);
                    returnRespone = service.Query.Sql(getRecSql).Execute();
                    returnResponeList = returnRespone.Rows;
                }

                callingForm.GoogleApiCall("<---- Got record for: " + pin);
                //this works to see if the current record in the table and the one just read are equal
                var what = returnRespone.Columns.Zip(returnResponeList[0],  (first, second) => new KeyValuePair<string, string>(first, second.ToString()));
                fieldListFusion = what.ToDictionary(k => k.Key, v => v.Value);
            }

            if (!fieldListRealist.SequenceEqual(fieldListFusion))
            {
                //row needs updates
                Dictionary<string, string> updates = new Dictionary<string, string>();

                //var u = fieldListFusion.Except(fieldListRealist);
                var u = fieldListRealist.Except(fieldListFusion);
                updates = u.ToDictionary(k => k.Key, v => v.Value);

                //if we are here, something in row needs updating, change date to current date.
                updates.Add("Date", DateTime.Now.ToShortDateString());

                bool updateSent = false;
                string sqlQuery = "";
                nameValuePairs.Insert(0, "Date = " + DateTime.Now.ToShortDateString() + "',");

                foreach (string key in updates.Keys)
                {
                    updateSent = false;
                    sqlQuery = "";
                    nameValuePairs = nameValuePairs.Insert(0, "'" + key + "' = '" + updates[key].Replace("'","") + "',");
                   // string sqlQuery = @"UPDATE " + tableId + " SET " + "'" + key + "' = '" + updates[key] + "'" + " WHERE ROWID = '" + rowid + "'";

                   // service.Query.Sql(sqlQuery).Execute();
                    //Thread.Sleep(300);

                    // string sqlQuery = @"INSERT INTO " + tableId + " ('Parcel ID:', Date, Location) VALUES (" + sPin + ", '" + DateTime.Now.ToShortDateString() +"' ,'" + sAddress +"')";
          
                    sqlQuery = @"UPDATE " + tableId + " SET " + nameValuePairs.Remove(nameValuePairs.Length - 1) + " WHERE ROWID = '" + rowid + "'";
                    //this is to stop stings from going over get/post character limits
                    if (nameValuePairs.Length > 400)
                    {
                         //sqlQuery = @"UPDATE " + tableId + " SET " + nameValuePairs.Remove(nameValuePairs.Length - 1) + " WHERE ROWID = '" + rowid + "'";
                        
                         
                        try
                        {
                            callingForm.GoogleApiCall("Update Record---->");
                            service.Query.Sql(sqlQuery).Execute();
                        }
                        catch (Exception ex)
                        {
                            if (ex.Message.Contains("503"))
                            {

                                callingForm.GoogleApiCall("Failed, 503. Waiting to retry...");
                                Thread.Sleep(1000);
                                try
                                {
                                    service.Query.Sql(sqlQuery).Execute();
                                }
                                catch
                                {
                                    callingForm.GoogleApiCall("Failed, 503. Waiting to retry...");
                                    Thread.Sleep(5000);
                                    service.Query.Sql(sqlQuery).Execute();
                                }
                               
                            }
                        }
                        callingForm.GoogleApiCall("<---- Success");
                        nameValuePairs = "";
                        updateSent = true;
                    }

                  
                }
                if (!updateSent)
                {
                    try
                    {
                        callingForm.GoogleApiCall("Update Record---->");
                        service.Query.Sql(sqlQuery).Execute();
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("503"))
                        {

                            callingForm.GoogleApiCall("Failed, 503. Waiting to retry...");
                            Thread.Sleep(1500);
                            try
                            {
                                service.Query.Sql(sqlQuery).Execute();
                            }
                            catch
                            {
                                callingForm.GoogleApiCall("Failed, 503. Waiting to retry...");
                                Thread.Sleep(5000);
                                try
                                {
                                    service.Query.Sql(sqlQuery).Execute();
                                }
                                catch
                                {
                                    callingForm.GoogleApiCall("Failed, 503. Waiting to retry...");
                                    Thread.Sleep(15000);
                                    service.Query.Sql(sqlQuery).Execute();
                                }
                            }

                        }
                    }
                    callingForm.GoogleApiCall("<---- Success");
                    nameValuePairs = "";
                    updateSent = true;
                }
            }

            RecordExists(pin);

           
           // var ddddd = what;

          //  bool aequalb = fieldList.SequenceEqual(fieldList2);

          // string sqlQuery = @"UPDATE " + tableId + " SET " + nameValuePairs.Remove(nameValuePairs.Length - 1) + " WHERE ROWID = '" + rowid + "'";

          //  service.Query.Sql(sqlQuery).Execute();


        }

        public bool RecordExists(string pin)
        {
            //this query will pull the current record from the fusion table based on parcel id

            try
            {
                if (curRec["Parcel ID:"] == pin)
                    return true;
            }
            catch
            {
            }

            string getRecSql = @"SELECT * FROM " + tableId + " WHERE 'Parcel ID:' = '" + pin + "'";
            //execute sql statement, this is the response

            //if (m_rowid.IsNotNullOrEmpty())
            //{
            //    getRecSql = @"SELECT * FROM " + tableId + " WHERE ROWID = '" + m_rowid + "'";
            //}


            //var returnRespone = service.Query.Sql(getRecSql).Execute();
            //var returnResponeList = returnRespone.Rows;
            Google.Apis.Fusiontables.v1.Data.Sqlresponse returnRespone = new Google.Apis.Fusiontables.v1.Data.Sqlresponse();
            
            try
            {
                callingForm.GoogleApiCall("Record Exsists?---->");
                returnRespone = service.Query.Sql(getRecSql).Execute();
            }
            catch
            {
                callingForm.GoogleApiCall("Failed, Waiting to retry....");
                Thread.Sleep(1000);
                returnRespone = service.Query.Sql(getRecSql).Execute();
            }
            var returnResponeList = returnRespone.Rows;
            if (returnRespone.Rows != null)
            {
                var what = returnRespone.Columns.Zip(returnResponeList[0], (first, second) => new KeyValuePair<string, string>(first, second.ToString()));
                curRec = what.ToDictionary(k => k.Key, v => v.Value);
                callingForm.GoogleApiCall("<----Yes");
                lastPin = pin;

                return true;
            }
            else
            {
                callingForm.GoogleApiCall("<----No, Record does not exsist.");
                return false;
            }
          
            
        }

        //private static IAuthorizationState GetAuthentication(NativeApplicationClient client)
        //{
        //    // You should use a more secure way of storing the key here as
        //    // .NET applications can be disassembled using a reflection tool.
        //    const string STORAGE = "google.samples.dotnet.fusiontable";
        //    const string KEY = "fusion";

        //    string scope = FusiontablesService.Scope.Fusiontables.ToString();


        //    // Check if there is a cached refresh token available.
        //    IAuthorizationState state = AuthorizationMgr.GetCachedRefreshToken(STORAGE, KEY);
        //    if (state != null)
        //    {
        //        try
        //        {
        //            client.RefreshToken(state);
        //            return state; // Yes - we are done.
        //        }
        //        catch (DotNetOpenAuth.Messaging.ProtocolException ex)
        //        {
        //            CommandLine.WriteError("Using existing refresh token failed: " + ex.Message);
        //        }
        //    }




        //    // Retrieve the authorization from the user.
        //    state = AuthorizationMgr.RequestNativeAuthorization(client, scope);
        //    AuthorizationMgr.SetCachedRefreshToken(STORAGE, KEY, state);

        //    return state;
        //}


    }

    public class FuTableRow : IEquatable<FuTableRow>
    {
        public string Name { get; set; }
        public string  Value { get; set; }

        public bool Equals(FuTableRow other)
        {

            //Check whether the compared object is null. 
            if (Object.ReferenceEquals(other, null)) return false;

            //Check whether the compared object references the same data. 
            if (Object.ReferenceEquals(this, other)) return true;

            //Check whether the products' properties are equal. 
            return Value.Equals(other.Value) && Name.Equals(other.Name);
        }

        // If Equals() returns true for a pair of objects  
        // then GetHashCode() must return the same value for these objects. 

        public override int GetHashCode()
        {

            //Get hash code for the Name field if it is not null. 
            int hashProductName = Name == null ? 0 : Name.GetHashCode();

            //Get hash code for the Code field. 
            int hashProductCode = Value.GetHashCode();

            //Calculate the hash code for the product. 
            return hashProductName ^ hashProductCode;
        }
    }

class GoogleCloudDatastore
{

    static bpohelper.Form1 callingForm;

    static DatastoreService service;

    public void StoreMredListing(MLSListing mredListing)
    {
        Google.Apis.Datastore.v1beta2.Data.CommitRequest cr = new Google.Apis.Datastore.v1beta2.Data.CommitRequest();
        
         Google.Apis.Datastore.v1beta2.Data.BeginTransactionRequest btr = new Google.Apis.Datastore.v1beta2.Data.BeginTransactionRequest();

         Google.Apis.Datastore.v1beta2.Data.Entity ml = new Google.Apis.Datastore.v1beta2.Data.Entity();

       Google.Apis.Datastore.v1beta2.Data.KeyPathElement myKeyPath = new Google.Apis.Datastore.v1beta2.Data.KeyPathElement();

        myKeyPath.Kind = "MRED_Listing";
        myKeyPath.Name = mredListing.MlsNumber;

        Google.Apis.Datastore.v1beta2.Data.Key myEntityKey = new Google.Apis.Datastore.v1beta2.Data.Key();
        Google.Apis.Datastore.v1beta2.Data.Key rawHtmlKey = new Google.Apis.Datastore.v1beta2.Data.Key();

        Google.Apis.Datastore.v1beta2.Data.PartitionId myPartitionId = new Google.Apis.Datastore.v1beta2.Data.PartitionId();
       
      

        myPartitionId.DatasetId = "active-century-477";
       myEntityKey.PartitionId = myPartitionId;

      //  Google.Apis.Datastore.v1beta2.Data.Property myProperty = new Google.Apis.Datastore.v1beta2.Data.Property();
      
       // myProperty.StringValue = mredListing.rawData;
       
       ml.Key = myEntityKey;

        

   // ml.Key.PartitionId.DatasetId = "active-century-477";

         //ml.Properties["rawhtml"].StringValue = mredListing.rawData;
        ml.Key.Path = new List<Google.Apis.Datastore.v1beta2.Data.KeyPathElement>();

       ml.Key.Path.Add(myKeyPath);

        //Google.Apis.Datastore.v1beta2.Data. 
      //   Dictionary<string, Google.Apis.Datastore.v1beta2.Data.Property> myPropertyDictionary = new Dictionary<string, Google.Apis.Datastore.v1beta2.Data.Property>();
       //  myPropertyDictionary.Add("raw_html", myProperty);
        // ml.Properties = myPropertyDictionary;

         //foreach (KeyValuePair<string, Google.Apis.Datastore.v1beta2.Data.Property> kvp in myPropertyDictionary)
         //{
         //      ml.Properties.Add(kvp.Key, kvp.Value); 
             
         //}

       
         cr.Mutation = new Google.Apis.Datastore.v1beta2.Data.Mutation();
        // cr.Mutation.Upsert = new  List<Google.Apis.Datastore.v1beta2.Data.Entity>();
        cr.Mutation.Insert = new List<Google.Apis.Datastore.v1beta2.Data.Entity>();
      //   cr.Mutation.Upsert.Add(ml);

        cr.Mutation.Insert.Add(ml);
        
         //service.Datasets.BeginTransaction(btr, mredListing.MlsNumber);
        var btRequest = service.Datasets.BeginTransaction(btr, "active-century-477");
        var btResponse = btRequest.Execute();

        cr.Transaction = btResponse.Transaction;

         var request = service.Datasets.Commit(cr, "active-century-477");

          var response = request.Execute();

          MessageBox.Show(response.ToString());
        
        }

    public  async Task DataStoreTester()
    //public void   DataStoreTester()
        {
            callingForm = (bpohelper.Form1) Form1.ActiveForm;
            // Register the authenticator.
           // var provider = new NativeApplicationClient(GoogleAuthenticationServer.Description);
          //  provider.ClientIdentifier = "982969682733.apps.googleusercontent.com";
          //  provider.ClientSecret = "ADjXMNxzAqa35fB4Z12UORM9";
         


            UserCredential credential;

            credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
              new ClientSecrets
              {
                  ClientId = "22551269451-rfgki0102ravo9scfu4d2l8i7mbi2q6b.apps.googleusercontent.com",
                  ClientSecret = "uluoVq2gn7SGucJzV9UVAuko"
                  // APIKey ="AIzaSyDsAcyc5WOw0IqvR93VR8cVsWCHJ8ZoDq4"
              },
              new[] { DatastoreService.Scope.Datastore, DatastoreService.Scope.UserinfoEmail},
              "beilsco@gmail.com",
              CancellationToken.None
                //new FileDataStore("Books.ListMyLibrary"));
              );

            //service = new FusiontablesService(auth);
           service = new DatastoreService(new BaseClientService.Initializer() 
            {
                HttpClientInitializer = credential,
                ApplicationName = "Datastore API Sample",
                
            });

      

        	//	Google.Apis.Datastore.v1beta2.Data.RunQueryRequest body = new Google.Apis.Datastore.v1beta2.Data.RunQueryRequest();

              //  Google.Apis.Datastore.v1beta2.Data.Query myq = new Google.Apis.Datastore.v1beta2.Data.Query();

               

             //   body.Query = myq;

              //  Google.Apis.Datastore.v1beta2.DatasetsResource.RunQueryRequest rqr = new Google.Apis.Datastore.v1beta2.DatasetsResource.RunQueryRequest(service, body, "active-century-477");

       //    Google.Apis.Datastore.v1beta2.Data.RunQueryRequest myQuery = new Google.Apis.Datastore.v1beta2.Data.RunQueryRequest();
       //    myQuery.Query = new  Google.Apis.Datastore.v1beta2.Data.Query();
                   
       //   var ent = service.Datasets.RunQuery(myQuery, "active-century-477");

        
        
       //// MessageBox.Show(ent.ToString());
      
        
       // Google.Apis.Datastore.v1beta2.Data.RunQueryResponse res;

         
       //     res =   ent.Execute();
      
       //    string str = res.Batch.EntityResults.Count.ToString();
       //    GlobalVar.sandbox = str;

      //    MessageBox.Show( res.Batch.EntityResults.Count.ToString());
          
     //   res.Batch.EntityResults[0].Entity.Properties[0]

         //  var ttt = rqr.Execute();

           ///var ttt = service.Datasets.RunQuery(r, "active-century-477").Execute();
        
      //       MessageBox.Show(res.ToString());

            //var service = new BooksService(new BaseClientService.Initializer()
            //{
            //    HttpClientInitializer = credential,
            //    ApplicationName = "Books API Sample",
            //});

            //// Create the service.
            //var service = new DiscoveryService(new BaseClientService.Initializer
            //{
            //    ApplicationName = "Discovery Sample",
            //    APIKey = "AIzaSyDsAcyc5WOw0IqvR93VR8cVsWCHJ8ZoDq4",
            //});
            
        }

    // public void  DataStoreTester ()
    //{
    //     Google.Apis.Datastore.v1beta2.Data.RunQueryRequest r = new Google.Apis.Datastore.v1beta2.Data.RunQueryRequest();
    //     //r.Query = new Google.Apis.Datastore.v1beta2.Data.Query();



    //     var ttt = service.Datasets.RunQuery(r, "active-century-477").Execute();


    //     MessageBox.Show(ttt.ToString());


    //}

}

}
