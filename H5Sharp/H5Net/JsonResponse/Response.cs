using System.Collections.Generic;

namespace H5Net.JsonResponse
{
    public class Response
    {
        public List<Result> results { get; set; }
        public bool wasTerminated { get; set; }
        public int nrOfSuccessfullTransactions { get; set; }
        public int nrOfFailedTransactions { get; set; }
        public string terminationReason { get; set; }
    }
}