using System;
using System.Windows;
using System.Windows.Navigation;
using Thinktecture.IdentityModel.Client;

namespace AuthenticationWindow
{
    /// <summary>
    /// Interaction logic for LoginWebView.xaml
    /// </summary>
    public partial class LoginWebView : Window
    {
        public AuthorizeResponse AuthorizeResponse { get; set; }

        public event EventHandler<AuthorizeResponse> Done;

        private Uri _callbackUri;

        public LoginWebView()
        {
            InitializeComponent();
            webView.Navigating += webView_Navigating;
            Closing += LoginWebView_Closing;
        }

        public void LoginWebView_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.Visibility = Visibility.Hidden;
        }

        public void Start(Uri startUri, Uri callbackUrl)
        {
            _callbackUri = callbackUrl;
            webView.Navigate(startUri);
        }

        public void webView_Navigating(object sender, NavigatingCancelEventArgs e)
        {
            if (e.Uri.ToString().StartsWith(_callbackUri.AbsoluteUri))
            {
                AuthorizeResponse = new AuthorizeResponse(e.Uri.AbsoluteUri);
                e.Cancel = true;
                this.Visibility = System.Windows.Visibility.Hidden;

                if (Done != null)
                {
                    Done.Invoke(this, AuthorizeResponse);
                }
            }

            if (e.Uri.ToString().Equals("javascript:void(0)"))
            {
                e.Cancel = true;
            }
        }
    }
}