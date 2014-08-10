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




namespace bpohelper
{

  

    public static class GlobalVar
    {
        public const string defaultSubjectComments = "Subject is maintained and landscaped. No adverse conditions were noted at the time of inspection based on exterior observations.";
        public const string defaultRepairComments = "No repairs noted from drive-by.";
        public enum CompCompareSystem
        {
            NABPOP,
            USER
        }
        public static CompCompareSystem ccc = new CompCompareSystem();

        public static decimal upperGLA = 0;
        public static decimal lowerGLA = 0;
        public static decimal lowerLotSize = 0;
        public static decimal upperLotSize = 0;
        public static GeoPoint subjectPoint;

        public static List<MLSListing> searchCacheMlsListings = new List<MLSListing>();
        public static SubjectProperty theSubjectProperty = new SubjectProperty();

        public static List<MLSListing> listingsFromLastSearch = new List<MLSListing>();

        public static IEnumerable<MLSListing> listingsSearchResults;

        public static string sandbox = "";


    }



    public class BpoOrder
    {
        public DateTime timestamp { get; set; }
        public string address { get; set; }
        public string duedate { get; set; }
        public string pics { get; set; }
        public string completed { get; set; }
        public string strippedSite { get; set; }
        public string url { get; set; }
        public string twitter { get; set; }
        public string ordernum { get; set; }
    }

    public class DataItem
    {
        public string Key;

        public string Value;

        public DataItem()
        {

        }

        public DataItem(string key, string value)
        {
            Key = key;
            Value = value;
        }
    }

    sealed class XmlDocumentSample
    {
        private XmlDocumentSample() { }

        static XmlReader reader;
        static String filename = @"C:\Users\Scott\Documents\My Dropbox\BPOs\317 BRIERHILL DR\test-completed-form.xml";

        public static void MyMethod()
        {

            ValidationEventHandler eventHandler = new ValidationEventHandler(XmlDocumentSample.ValidationCallback);

            try
            {
                // Create the validating reader and specify DTD validation.
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.DtdProcessing = DtdProcessing.Prohibit;
                settings.ValidationType = ValidationType.None;
                settings.ValidationEventHandler += eventHandler;

                reader = XmlReader.Create(filename, settings);
     
                // Pass the validating reader to the XML document. 
                // Validation fails due to an undefined attribute, but the  
                // data is still loaded into the document.
                XmlDocument doc = new XmlDocument();
                doc.Load(reader);
                XDocument anotherDoc = XDocument.Load(XmlReader.Create(filename, settings));

                anotherDoc.Element("FORMINFO").Element("SUBJECT").Element("PROPTYPE").SetValue("TTT");

                var data = from item in anotherDoc.Descendants("LISTINGCOMP")
                           select new
                           {
                               listComp1 = item.Element("LISTCOMPS").Element("PROPTYPE").Value,
                               listComp2 = item.Element("LISTCOMPS").Element("PROPTYPE").Value
                              // listComp3 = item.
                              // moneySpent = item.Element("PROPTYPE").Value,
                               //zipCode = item.Element("personalInfo").Element("zip").Value
                           };
                foreach (var p in data)
                {
                    Console.WriteLine(p.ToString());

                }
                Console.WriteLine(doc.OuterXml);
            }
            finally
            {
                if (reader != null)
                    reader.Close();
            }
        }

        // Display the validation error. 
        private static void ValidationCallback(object sender, ValidationEventArgs args)
        {
            Console.WriteLine("Validation error loading: {0}", filename);
            Console.WriteLine(args.Message);
        }
    }


    public enum DistanceType { Miles, Kilometers };

    public struct GeoPoint
    {
        public double Latitude;
        public double Longitude;
    }

    class Haversine
    {

        public double Distance(GeoPoint pos1, GeoPoint pos2, DistanceType type)
        {
            double preDlat = pos2.Latitude - pos1.Latitude;
            double preDlon = pos2.Longitude - pos1.Longitude;
            double R = (type == DistanceType.Miles) ? 3960 : 6371;
            double dLat = this.toRadian(preDlat);
            double dLon = this.toRadian(preDlon);
            double a = Math.Sin(dLat / 2) * Math.Sin(dLat / 2) +
            Math.Cos(this.toRadian(pos1.Latitude)) * Math.Cos(this.toRadian(pos2.Latitude)) *
            Math.Sin(dLon / 2) * Math.Sin(dLon / 2);
            double c = 2 * Math.Asin(Math.Min(1, Math.Sqrt(a)));
            double d = R * c;
            return d;
        }

        private double toRadian(double val)
        {
            return (Math.PI / 180) * val;
        }
    }

    public static class RichTextBoxExtensions
    {
        public static void AppendText(this RichTextBox box, string text, Color color)
        {
            box.SelectionStart = box.TextLength;
            box.SelectionLength = 0;

            box.SelectionColor = color;
            box.AppendText(text);
            box.SelectionColor = box.ForeColor;
        }
    }

    public partial class Form1
    {

        public int RoundTo50(string number)
        {
            double p = 0;

            Double.TryParse(number, out p);

            return (int)Math.Round(p / 50) * 50;

        }

        public IList<string> GenerateVersions(string original)
        {
            Dictionary<string, string> versions = new Dictionary<string, string>();
            //Define the versions to generate and their filename suffixes.
            versions.Add("_thumb", "width=100&height=100&crop=auto&format=jpg"); //Crop to square thumbnail
            versions.Add("_medium", "maxwidth=400&maxheight=400format=jpg"); //Fit inside 400x400 area, jpeg
            versions.Add("_large", "maxwidth=1900&maxheight=1900&format=jpg"); //Fit inside 1900x1200 area


            string basePath = ImageResizer.Util.PathUtils.RemoveExtension(original);

            //To store the list of generated paths
            List<string> generatedFiles = new List<string>();

            //Generate each version
            foreach (string suffix in versions.Keys)
                //Let the image builder add the correct extension based on the output file type
                generatedFiles.Add(ImageBuilder.Current.Build(original, basePath + suffix,
                  new ResizeSettings(versions[suffix]), false, true));

            return generatedFiles;
        }

       public void CreateDirectory(string path)
       {
           try
           {
               // Determine whether the directory exists. 
               if (Directory.Exists(path))
               {
                   Console.WriteLine("That path exists already.");
                   return;
               }

               // Try to create the directory.
               DirectoryInfo di = Directory.CreateDirectory(path);
               Console.WriteLine("The directory was created successfully at {0}.", Directory.GetCreationTime(path));

              
           }
           catch (Exception e)
           {
               Console.WriteLine("The process failed: {0}", e.ToString());
           }
           finally { }

       }

        private string dropBoxFolderName;
        public string DropBoxFolder
        {
            get
            {
                if (string.IsNullOrEmpty(dropBoxFolderName))
                {
                    var dbPath = System.IO.Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Dropbox\\host.db");

                    string[] lines = System.IO.File.ReadAllLines(dbPath);
                    byte[] dbBase64Text = Convert.FromBase64String(lines[1]);
                    string folderPath = System.Text.ASCIIEncoding.ASCII.GetString(dbBase64Text);
                    dropBoxFolderName = folderPath;
                }
                return dropBoxFolderName;
            }

            set { dropBoxFolderName = value; }
        }

        public GeoPoint Geocode(string sAddress)
        {
            //
            //geocode address
            //ex:
            //http://maps.googleapis.com/maps/api/geocode/json?address=1600+Amphitheatre+Parkway,+Mountain+View,+CA&sensor=true_or_false



            GeoPoint pt = new GeoPoint();


            string googlestr = @" http://maps.googleapis.com/maps/api/geocode/json?address=" + sAddress.Replace(" ", "+").Replace("#", "") + "&sensor=false";

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
                Thread.Sleep(550);
                try
                {
                  
                    response = (HttpWebResponse)request.GetResponse();
                }
                catch
                {
                    Thread.Sleep(1100);
                    response = (HttpWebResponse)request.GetResponse();
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
            try
            {

                pt.Latitude = Convert.ToDouble(propertyLatitude);
                pt.Longitude = Convert.ToDouble(propertyLongitude);
            }

            catch
            {

            }
            return pt;



        }


    }

    public class ObjectShredder<T>
    {
        private System.Reflection.FieldInfo[] _fi;
        private System.Reflection.PropertyInfo[] _pi;
        private System.Collections.Generic.Dictionary<string, int> _ordinalMap;
        private System.Type _type;

        // ObjectShredder constructor.
        public ObjectShredder()
        {
            _type = typeof(T);
            _fi = _type.GetFields();
            _pi = _type.GetProperties();
            _ordinalMap = new Dictionary<string, int>();
        }

        /// <summary>
        /// Loads a DataTable from a sequence of objects.
        /// </summary>
        /// <param name="source">The sequence of objects to load into the DataTable.</param>
        /// <param name="table">The input table. The schema of the table must match that 
        /// the type T.  If the table is null, a new table is created with a schema 
        /// created from the public properties and fields of the type T.</param>
        /// <param name="options">Specifies how values from the source sequence will be applied to 
        /// existing rows in the table.</param>
        /// <returns>A DataTable created from the source sequence.</returns>
        public DataTable Shred(IEnumerable<T> source, DataTable table, LoadOption? options)
        {
            // Load the table from the scalar sequence if T is a primitive type.
            if (typeof(T).IsPrimitive)
            {
                return ShredPrimitive(source, table, options);
            }

            // Create a new table if the input table is null.
            if (table == null)
            {
                table = new DataTable(typeof(T).Name);
            }

            // Initialize the ordinal map and extend the table schema based on type T.
            table = ExtendTable(table, typeof(T));

            // Enumerate the source sequence and load the object values into rows.
            table.BeginLoadData();
            using (IEnumerator<T> e = source.GetEnumerator())
            {
                while (e.MoveNext())
                {
                    if (options != null)
                    {
                        table.LoadDataRow(ShredObject(table, e.Current), (LoadOption)options);
                    }
                    else
                    {
                        table.LoadDataRow(ShredObject(table, e.Current), true);
                    }
                }
            }
            table.EndLoadData();

            // Return the table.
            return table;
        }

        public DataTable ShredPrimitive(IEnumerable<T> source, DataTable table, LoadOption? options)
        {
            // Create a new table if the input table is null.
            if (table == null)
            {
                table = new DataTable(typeof(T).Name);
            }

            if (!table.Columns.Contains("Value"))
            {
                table.Columns.Add("Value", typeof(T));
            }

            // Enumerate the source sequence and load the scalar values into rows.
            table.BeginLoadData();
            using (IEnumerator<T> e = source.GetEnumerator())
            {
                Object[] values = new object[table.Columns.Count];
                while (e.MoveNext())
                {
                    values[table.Columns["Value"].Ordinal] = e.Current;

                    if (options != null)
                    {
                        table.LoadDataRow(values, (LoadOption)options);
                    }
                    else
                    {
                        table.LoadDataRow(values, true);
                    }
                }
            }
            table.EndLoadData();

            // Return the table.
            return table;
        }

        public object[] ShredObject(DataTable table, T instance)
        {

            FieldInfo[] fi = _fi;
            PropertyInfo[] pi = _pi;

            if (instance.GetType() != typeof(T))
            {
                // If the instance is derived from T, extend the table schema
                // and get the properties and fields.
                ExtendTable(table, instance.GetType());
                fi = instance.GetType().GetFields();
                pi = instance.GetType().GetProperties();
            }

            // Add the property and field values of the instance to an array.
            Object[] values = new object[table.Columns.Count];
            foreach (FieldInfo f in fi)
            {
                values[_ordinalMap[f.Name]] = f.GetValue(instance);
            }

            foreach (PropertyInfo p in pi)
            {
                values[_ordinalMap[p.Name]] = p.GetValue(instance, null);
            }

            // Return the property and field values of the instance.
            return values;
        }

        public DataTable ExtendTable(DataTable table, Type type)
        {
            // Extend the table schema if the input table was null or if the value 
            // in the sequence is derived from type T.            
            foreach (FieldInfo f in type.GetFields())
            {
                if (!_ordinalMap.ContainsKey(f.Name))
                {
                    // Add the field as a column in the table if it doesn't exist
                    // already.
                    DataColumn dc = table.Columns.Contains(f.Name) ? table.Columns[f.Name]
                        : table.Columns.Add(f.Name, f.FieldType);

                    // Add the field to the ordinal map.
                    _ordinalMap.Add(f.Name, dc.Ordinal);
                }
            }
            foreach (PropertyInfo p in type.GetProperties())
            {
                if (!_ordinalMap.ContainsKey(p.Name))
                {
                    // Add the property as a column in the table if it doesn't exist
                    // already.
                    DataColumn dc = table.Columns.Contains(p.Name) ? table.Columns[p.Name]
                        : table.Columns.Add(p.Name, p.PropertyType);

                    // Add the property to the ordinal map.
                    _ordinalMap.Add(p.Name, dc.Ordinal);
                }
            }

            // Return the table.
            return table;
        }
    }

    public static class CustomLINQtoDataSetMethods
    {
        public static DataTable CopyToDataTable<T>(this IEnumerable<T> source)
        {
            return new ObjectShredder<T>().Shred(source, null, null);
        }

        public static DataTable CopyToDataTable<T>(this IEnumerable<T> source,
                                                    DataTable table, LoadOption? options)
        {
            return new ObjectShredder<T>().Shred(source, table, options);
        }

    }


}
