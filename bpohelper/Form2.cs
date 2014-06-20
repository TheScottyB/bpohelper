using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Google.Apis.Fusiontables.v1;
using DotNetOpenAuth.OAuth2;
using Google.Apis.Authentication;
using Google.Apis.Authentication.OAuth2;
using Google.Apis.Authentication.OAuth2.DotNetOpenAuth;
using Google.Apis.Samples.Helper;
using Google.Apis.Util;
using System.Threading;




namespace bpohelper
{
    public partial class equiTraxView : Form
    {

        public string pin;
        public string style;

        public equiTraxView()
        {
            InitializeComponent();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_VisibleChanged(object sender, EventArgs e)
        {
            if (styleComboBox.Items.Contains(style))
            {
                styleComboBox.Text = style;
            }
            parcelIdTextBox.Text = pin;


        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            //
            //Google fusion table trial
            //

            // Register the authenticator.
            var provider = new NativeApplicationClient(GoogleAuthenticationServer.Description);
            provider.ClientIdentifier = "982969682733.apps.googleusercontent.com";
            provider.ClientSecret = "ADjXMNxzAqa35fB4Z12UORM9";
            var auth = new OAuth2Authenticator<NativeApplicationClient>(provider, GetAuthentication);

            // Create the service.
            //var service = new PredictionService(auth);

            var service = new FusiontablesService();

            string sqlQuery = @"INSERT INTO 1cU8NHmtVbE3KWpGa2qSTkvIpwhC5KIJX3hcvlIg (OrderNumber, SubjectPin) VALUES (" + orderNumComboBox.Text + ", " + parcelIdTextBox.Text + ")";

            service.Query.Sql(sqlQuery).Execute();

            sqlQuery = @"SELECT * FROM 1UeTFdijs1WJRyAgAURozDkWph7np85r_6wg3tx0 WHERE 'MLS #' = '8086017'";
            Google.Apis.Fusiontables.v1.Data.Sqlresponse ttt = service.Query.SqlGet(sqlQuery).Execute();

            MessageBox.Show(ttt.Kind);


            Thread.Sleep(20000);

            


            var tttt = service.Table.List().Execute();

           


          // MessageBox.Show(tttt.Items.Count.ToString());




        }
        private static IAuthorizationState GetAuthentication(NativeApplicationClient client)
        {
            // You should use a more secure way of storing the key here as
            // .NET applications can be disassembled using a reflection tool.
            const string STORAGE = "google.samples.dotnet.fusiontable";
            const string KEY = "fusion";
            
            //string scope = FusiontablesService.Scope.Fusiontables.GetStringValue();
            string scope = FusiontablesService.Scope.Fusiontables.ToString();

            // Check if there is a cached refresh token available.
            IAuthorizationState state = AuthorizationMgr.GetCachedRefreshToken(STORAGE, KEY);
            if (state != null)
            {
                try
                {
                    client.RefreshToken(state);
                    return state; // Yes - we are done.
                }
                catch (DotNetOpenAuth.Messaging.ProtocolException ex)
                {
                    CommandLine.WriteError("Using existing refresh token failed: " + ex.Message);
                }
            }

         
         

            // Retrieve the authorization from the user.
            state = AuthorizationMgr.RequestNativeAuthorization(client, scope);
            AuthorizationMgr.SetCachedRefreshToken(STORAGE, KEY, state);
          
            return state;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //
            //Google fusion table trial
            //

            // Register the authenticator.
            var provider = new NativeApplicationClient(GoogleAuthenticationServer.Description);
            provider.ClientIdentifier = "982969682733.apps.googleusercontent.com";
            provider.ClientSecret = "ADjXMNxzAqa35fB4Z12UORM9";
            var auth = new OAuth2Authenticator<NativeApplicationClient>(provider, GetAuthentication);

            // Create the service.
            //var service = new PredictionService(auth);

            var service = new FusiontablesService();

            //add check if coloums exist, if not add
            //Name:	realist_bpohelper
            //Encrypted ID:	1UKrOVmhPWrgLP5d5bDCsiW9whMIK8aLxKhcyOaI
            var tableColums = service.Column.List("1UKrOVmhPWrgLP5d5bDCsiW9whMIK8aLxKhcyOaI").Execute();

            //tableColums.Items.Contains(

            Google.Apis.Fusiontables.v1.Data.Column myc = new Google.Apis.Fusiontables.v1.Data.Column();
            myc.Name = "pin";

            foreach (Google.Apis.Fusiontables.v1.Data.Column c in tableColums.Items)
            {
                if (c.Name == myc.Name)
                {
                    MessageBox.Show(tableColums.Kind);
                }
            }


            string sqlQuery = @"INSERT INTO 1cU8NHmtVbE3KWpGa2qSTkvIpwhC5KIJX3hcvlIg (OrderNumber, SubjectPin) VALUES (" + orderNumComboBox.Text + ", " + parcelIdTextBox.Text + ")";

            //service.Query.Sql(sqlQuery).Fetch();

            sqlQuery = @"SELECT * FROM 1UeTFdijs1WJRyAgAURozDkWph7np85r_6wg3tx0 WHERE 'MLS #' = '8086017'";
            Google.Apis.Fusiontables.v1.Data.Sqlresponse ttt = service.Query.SqlGet(sqlQuery).Execute();

            MessageBox.Show(tableColums.Kind);


            //Thread.Sleep(20000);




            var tttt = service.Table.List().Execute();




            // MessageBox.Show(tttt.Items.Count.ToString());
        }
    }
}
