using H5AuthData;
using System;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using Thinktecture.IdentityModel.Client;

namespace AuthenticationWindow
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string REDIRECT_URL = "";
        private string IONAPIBASE_URL = "";
        private string AUTHORIZATION_ENDPOINT = "";
        private string TOKEN_ENDPOINT = "";
        private string CLIENT_ID = "";
        private string CLIENT_SECRET = "";

        private string access_code = "";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void GetCodeButton_Click(object sender, RoutedEventArgs e)
        {
            string state = Guid.NewGuid().ToString("N");
            string nonce = Guid.NewGuid().ToString("N");

            try
            {
                var client = new OAuth2Client(new Uri(AUTHORIZATION_ENDPOINT));
                var startUrl = client.CreateCodeFlowUrl(
                   clientId: CLIENT_ID,
                   redirectUri: REDIRECT_URL,
                   state: state,
                   nonce: nonce);

                LoginWebView webView = new LoginWebView();
                webView.Owner = this;
                webView.Done += _login_Done;
                webView.Show();
                webView.Start(new Uri(startUrl), new Uri(REDIRECT_URL));
            }
            catch (System.UriFormatException ex)
            {
                MessageBox.Show("Invalid AuthEndPoint", "Url Error", MessageBoxButton.OK);
            }
        }

        public void _login_Done(object sender, AuthorizeResponse response)
        {
            //this.outputBox.AppendText(response.Raw);

            if (!String.IsNullOrWhiteSpace(response.Code))
            {
                //this.CodeTextBox.Text = response.Code;
                this.access_code = response.Code;

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                var client = new OAuth2Client(new Uri(TOKEN_ENDPOINT), CLIENT_ID, CLIENT_SECRET);
                var authCodeResponse = client.RequestAuthorizationCodeAsync(this.access_code, REDIRECT_URL).Result;

                this.Close();

                _handleTokenResponse(authCodeResponse);
            }
        }

        private void GetAccessTokenButton_Click(object sender, RoutedEventArgs e)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            var client = new OAuth2Client(new Uri(TOKEN_ENDPOINT), CLIENT_ID, CLIENT_SECRET);
            // var response = client.RequestAuthorizationCodeAsync(this.CodeTextBox.Text, REDIRECT_URL).Result;
            var response = client.RequestAuthorizationCodeAsync(this.access_code, REDIRECT_URL).Result;

            _handleTokenResponse(response);
        }

        private void _handleTokenResponse(TokenResponse tr)
        {
            //this.outputBox.AppendText(tr.Raw);

            if (!String.IsNullOrWhiteSpace(tr.AccessToken))
            {
                //this.AccessTokenTextBox.Text = tr.AccessToken;
                AuthData.AccessToken = tr.Raw;
            }
        }

        private void CallServiceButton_Click(object sender, RoutedEventArgs e)
        {
        }

        private void envSelection_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var selectedItem = (ComboBoxItem)envSelection.SelectedItem;
            var env = selectedItem.Content.ToString();

            if (env == "AGGREKO_TST")
            {
                this.REDIRECT_URL = AuthData.tst_Redirect_url;
                this.IONAPIBASE_URL = AuthData.tst_IonApiBase_url;
                this.AUTHORIZATION_ENDPOINT = AuthData.tst_Authorization_endPoint;
                this.TOKEN_ENDPOINT = AuthData.tst_Token_endPoint;
                this.CLIENT_ID = AuthData.tst_Client_id;
                this.CLIENT_SECRET = AuthData.tst_Client_secret;
                AuthData.SelectedEnvironment = env;
            }
            else if (env == "AGGREKO_TRN")
            {
                this.REDIRECT_URL = AuthData.trn_Redirect_url;
                this.IONAPIBASE_URL = AuthData.trn_IonApiBase_url;
                this.AUTHORIZATION_ENDPOINT = AuthData.trn_Authorization_endPoint;
                this.TOKEN_ENDPOINT = AuthData.trn_Token_endPoint;
                this.CLIENT_ID = AuthData.trn_Client_id;
                this.CLIENT_SECRET = AuthData.trn_Client_secret;
                AuthData.SelectedEnvironment = env;
            }
            else if (env == "AGGREKO_PRD")
            {
                this.REDIRECT_URL = AuthData.prd_Redirect_url;
                this.IONAPIBASE_URL = AuthData.prd_IonApiBase_url;
                this.AUTHORIZATION_ENDPOINT = AuthData.prd_Authorization_endPoint;
                this.TOKEN_ENDPOINT = AuthData.prd_Token_endPoint;
                this.CLIENT_ID = AuthData.prd_Client_id;
                this.CLIENT_SECRET = AuthData.prd_Client_secret;
                AuthData.SelectedEnvironment = env;
            }
            else
            {
                //MessageBox.Show($"{env} is not a valid M3 Environment");
                GetCodeButton.IsEnabled = false;
                return;
                //AuthData.SelectedEnvironment = "Invalid Env";
            }
        }
    }
}