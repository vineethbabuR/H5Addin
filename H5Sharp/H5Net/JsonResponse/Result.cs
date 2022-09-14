using System.Collections.Generic;

namespace H5Net.JsonResponse
{
    public class Result
    {
        public string results { get; set; }
        public string transaction { get; set; }

        // public List<string> records { get; set; }
        public string errorMessage { get; set; }

        public string errorType { get; set; }
        public string errorCode { get; set; }
        public string errorCfg { get; set; }
        public  string errorField { get; set; }
        public List<object> records { get; set; } // this is to handle batch api transactions
    }
}