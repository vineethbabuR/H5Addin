using Newtonsoft.Json.Linq;

namespace H5AuthData
{
    public class AuthData
    {
        public static string AccessToken;
        public static string SelectedEnvironment;

        public static string tst_Redirect_url;
        public static string tst_IonApiBase_url;
        public static string tst_Authorization_endPoint;
        public static string tst_Token_endPoint;
        public static string tst_Client_id;
        public static string tst_Client_secret;

        public static string trn_Redirect_url;
        public static string trn_IonApiBase_url;
        public static string trn_Authorization_endPoint;
        public static string trn_Token_endPoint;
        public static string trn_Client_id;
        public static string trn_Client_secret;

        public static string prd_Redirect_url;
        public static string prd_IonApiBase_url;
        public static string prd_Authorization_endPoint;
        public static string prd_Token_endPoint;
        public static string prd_Client_id;
        public static string prd_Client_secret;

        public static string GetAccessToken()
        {
            var bearer_token = "";
            if (!string.IsNullOrEmpty(AccessToken) || !string.IsNullOrWhiteSpace(AccessToken))
            {
                var str = AccessToken;
                JObject access_token = JObject.Parse(AccessToken);
                bearer_token = access_token["access_token"].Value<string>();
            }
            return bearer_token;
        }

        public static string GetRefreshToken()
        {
            var refresh_token = "";
            if (!string.IsNullOrEmpty(AccessToken) || !string.IsNullOrWhiteSpace(AccessToken))
            {
                var str = AccessToken;
                JObject access_token = JObject.Parse(AccessToken);
                refresh_token = access_token["refresh_token"].Value<string>();
            }
            return refresh_token;
        }

        public static string GetTokenType()
        {
            var token_type = "";
            if (!string.IsNullOrEmpty(AccessToken) || !string.IsNullOrWhiteSpace(AccessToken))
            {
                var str = AccessToken;
                JObject access_token = JObject.Parse(AccessToken);
                token_type = access_token["token_type"].Value<string>();
            }
            return token_type;
        }
    }
}