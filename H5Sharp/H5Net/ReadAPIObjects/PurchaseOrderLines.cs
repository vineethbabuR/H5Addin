namespace H5Net.ReadAPIObjects
{
    public class PurchaseOrderLines
    {
        public PurchaseOrderLinesResult[] results { get; set; }
        public bool wasTerminated { get; set; }
        public int nrOfSuccessfullTransactions { get; set; }
        public int nrOfFailedTransactions { get; set; }
    }

    public class PurchaseOrderLinesResult
    {
        public string transaction { get; set; }
        public PurchaseOrderLinesRecord[] records { get; set; }
    }

    public class PurchaseOrderLinesRecord
    {
        public string CONO { get; set; }
        public string PUNO { get; set; }
        public string PNLI { get; set; }
        public string PNLS { get; set; }
        public string ITNO { get; set; }
        public string SITE { get; set; }
        public string CODT { get; set; }
        public string CFQA { get; set; }
        public string PUUN { get; set; }
        public string ITDS { get; set; }
        public string EACD { get; set; }
        public string FACI { get; set; }
        public string WHLO { get; set; }
        public string POTC { get; set; }
        public string PUST { get; set; }
        public string PUSL { get; set; }
        public string SUNO { get; set; }
        public string PRCS { get; set; }
        public string SUFI { get; set; }
        public string PITD { get; set; }
        public string PITT { get; set; }
        public string SORN { get; set; }
        public string PUPR { get; set; }
        public string ODI1 { get; set; }
        public string ODI2 { get; set; }
        public string ODI3 { get; set; }
        public string CPPR { get; set; }
        public string CFD1 { get; set; }
        public string CFD2 { get; set; }
        public string CFD3 { get; set; }
        public string PPUN { get; set; }
        public string PUCD { get; set; }
        public string CPUC { get; set; }
        public string LNAM { get; set; }
        public string PTCD { get; set; }
        public string DWDT { get; set; }
        public string ORQA { get; set; }
        public string ADQA { get; set; }
        public string TNQA { get; set; }
        public string RVQA { get; set; }
        public string RCDT { get; set; }
        public string CAQA { get; set; }
        public string RJQA { get; set; }
        public string SDQA { get; set; }
        public string IVQA { get; set; }
        public string IDAT { get; set; }
        public string PLPN { get; set; }
        public string PLPS { get; set; }
        public string PURC { get; set; }
        public string BUYE { get; set; }
        public string GRMT { get; set; }
        public string PACT { get; set; }
        public string TXID { get; set; }
        public string ECVE { get; set; }
        public string OURR { get; set; }
        public string OURT { get; set; }
        public string VTCD { get; set; }
        public string ATNR { get; set; }
        public string RORC { get; set; }
        public string RORN { get; set; }
        public string RORL { get; set; }
        public string RORX { get; set; }
        public string PRIP { get; set; }
        public string IRCV { get; set; }
        public string PROD { get; set; }
        public string SDPC { get; set; }
        public string MODL { get; set; }
        public string TEDL { get; set; }
        public string TEL1 { get; set; }
        public string HAFE { get; set; }
        public string TIHM { get; set; }
        public string UNMS { get; set; }
        public string ORQT { get; set; }
        public string LOCD { get; set; }
        public string CUPR { get; set; }
        public string PUQT { get; set; }
        public string SAAM { get; set; }
        public string RVQT { get; set; }
        public string CFQT { get; set; }
        public string CUCP { get; set; }
        public string DUPL { get; set; }
        public string DUPO { get; set; }
        public string TRRC { get; set; }
        public string TRRN { get; set; }
        public string TRRL { get; set; }
        public string TRRX { get; set; }
        public string RASN { get; set; }
        public string PIAD { get; set; }
        public string SDES { get; set; }
        public string CIAD { get; set; }
        public string CDES { get; set; }
        public string WSCA { get; set; }
        public string OWHL { get; set; }
        public string OFCI { get; set; }
        public string ORAD { get; set; }
        public string GETY { get; set; }
        public string LNA2 { get; set; }
        public string CPRD { get; set; }
        public string ASNE { get; set; }
        public string UCA1 { get; set; }
        public string UCA2 { get; set; }
        public string UCA3 { get; set; }
        public string UCA4 { get; set; }
        public string UCA5 { get; set; }
        public string UCA6 { get; set; }
        public string UCA7 { get; set; }
        public string UCA8 { get; set; }
        public string UCA9 { get; set; }
        public string UCA0 { get; set; }
        public string UDN1 { get; set; }
        public string UDN2 { get; set; }
        public string UDN3 { get; set; }
        public string UDN4 { get; set; }
        public string UDN5 { get; set; }
        public string UDN6 { get; set; }
        public string UID1 { get; set; }
        public string UID2 { get; set; }
        public string UID3 { get; set; }
        public string UCT1 { get; set; }
        public string SCHN { get; set; }
        public string MUFT { get; set; }
        public string PORG { get; set; }

    }
}