using ExcelDna.Integration.CustomUI;
using H5AuthData;
using System.Configuration;
using System.Runtime.InteropServices;

namespace H5Net.ExcelOp
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        private string tokenType = "";
        /*
         * public Bitmap GetImageDetails(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "apiSettingsButton":
                    return new Bitmap(H5Net.Properties.Resources.settings);

                default:
                    return null;
            }
        }
        */

        public override string GetCustomUI(string RibbonID)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
                      <ribbon>
                        <tabs>
                          <tab id='h5tab' label='H5 CE'>
                            <group id='h5group' label='H5 CE'>
                              <button id='h5authbtn' imageMso='AccessRequests' label='H5 Auth' onAction='OnH5AuthButtonPressed'/>
                            </group >
                          </tab>
                        </tabs>
                      </ribbon>
                    </customUI>";
        }

        public void OnH5AuthButtonPressed(IRibbonControl control)
        {
            //CTPManager.ShowCTP();

            //var clientID = ConfigurationManager.AppSettings["k1"];
            //AuthData.clientID = clientID;

            AuthData.tst_Redirect_url = ConfigurationManager.AppSettings["TST_REDIRECT_URL"];
            AuthData.tst_IonApiBase_url = ConfigurationManager.AppSettings["TST_IONAPIBASE_URL"];
            AuthData.tst_Authorization_endPoint = ConfigurationManager.AppSettings["TST_AUTHORIZATION_ENDPOINT"];
            AuthData.tst_Token_endPoint = ConfigurationManager.AppSettings["TST_TOKEN_ENDPOINT"];
            AuthData.tst_Client_id = ConfigurationManager.AppSettings["TST_CLIENT_ID"];
            AuthData.tst_Client_secret = ConfigurationManager.AppSettings["TST_CLIENT_SECRET"];

            AuthData.trn_Redirect_url = ConfigurationManager.AppSettings["TRN_REDIRECT_URL"];
            AuthData.trn_IonApiBase_url = ConfigurationManager.AppSettings["TRN_IONAPIBASE_URL"];
            AuthData.trn_Authorization_endPoint = ConfigurationManager.AppSettings["TRN_AUTHORIZATION_ENDPOINT"];
            AuthData.trn_Token_endPoint = ConfigurationManager.AppSettings["TRN_TOKEN_ENDPOINT"];
            AuthData.trn_Client_id = ConfigurationManager.AppSettings["TRN_CLIENT_ID"];
            AuthData.trn_Client_secret = ConfigurationManager.AppSettings["TRN_CLIENT_SECRET"];

            AuthData.prd_Redirect_url = ConfigurationManager.AppSettings["PRD_REDIRECT_URL"];
            AuthData.prd_IonApiBase_url = ConfigurationManager.AppSettings["PRD_IONAPIBASE_URL"];
            AuthData.prd_Authorization_endPoint = ConfigurationManager.AppSettings["PRD_AUTHORIZATION_ENDPOINT"];
            AuthData.prd_Token_endPoint = ConfigurationManager.AppSettings["PRD_TOKEN_ENDPOINT"];
            AuthData.prd_Client_id = ConfigurationManager.AppSettings["PRD_CLIENT_ID"];
            AuthData.prd_Client_secret = ConfigurationManager.AppSettings["PRD_CLIENT_SECRET"];

            var authWindow = new AuthenticationWindow.MainWindow();
            authWindow.ShowDialog();

            /*
            if (!string.IsNullOrEmpty(AuthData.AccessToken) || !string.IsNullOrWhiteSpace(AuthData.AccessToken))
            {
                var str = AuthData.AccessToken;
                JObject access_token = JObject.Parse(AuthData.AccessToken);
                tokenType = access_token["token_type"].Value<string>();
            }
            */
            tokenType = AuthData.GetTokenType();

            if (tokenType == "Bearer")
            {
                CTPManager.ShowCTP();
            }
        }

        /*public void OnButtonPressed(IRibbonControl ctrl)
        {
            var authWindow = new AuthenticationWindow.MainWindow();
            authWindow.ShowDialog();
        }
        */
    }
}