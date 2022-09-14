namespace H5Net.ReadAPIObjects
{
    public class AgreementLines
    {
        public AgreementLineResult[] results { get; set; }
        public bool wasTerminated { get; set; }
        public int nrOfSuccessfullTransactions { get; set; }
        public int nrOfFailedTransactions { get; set; }
    }

    public class AgreementLineResult
    {
        public string transaction { get; set; }
        public AgreementLineRecord[] records { get; set; }
    }

    public class AgreementLineRecord
    {
        public string CONO { get; set; }
        public string DIVI { get; set; }
        public string FACI { get; set; }
        public string AGNB { get; set; }
        public string PONR { get; set; }
        public string POSX { get; set; }
        public string VERS { get; set; }
        public string CUPL { get; set; }
        public string IVAD { get; set; }
        public string SAID { get; set; }
        public string SCNM { get; set; }
        public string SAD1 { get; set; }
        public string SAD2 { get; set; }
        public string SAD3 { get; set; }
        public string SAD4 { get; set; }
        public string IYRF { get; set; }
        public string IPHN { get; set; }
        public string FVDT { get; set; }
        public string LVDT { get; set; }
        public string TEDA { get; set; }
        public string ITDT { get; set; }
        public string ADPW { get; set; }
        public string ANOS { get; set; }
        public string ITNO { get; set; }
        public string BANO { get; set; }
        public string DIP1 { get; set; }
        public string DIP2 { get; set; }
        public string DIP3 { get; set; }
        public string DIP4 { get; set; }
        public string DIP5 { get; set; }
        public string DIA1 { get; set; }
        public string DIA2 { get; set; }
        public string DIA3 { get; set; }
        public string DIA4 { get; set; }
        public string DIA5 { get; set; }
        public string DIA6 { get; set; }
        public string CCAP { get; set; }
        public string ANOH { get; set; }
        public string PDAP { get; set; }
        public string PPCA { get; set; }
        public string AMAI { get; set; }
        public string PDAN { get; set; }
        public string PNCA { get; set; }
        public string ASTH { get; set; }
        public string DMOD { get; set; }
        public string COMD { get; set; }
        public string DLDT { get; set; }
        public string COLD { get; set; }
        public string PROJ { get; set; }
        public string ELNO { get; set; }
        public string FWHL { get; set; }
        public string TWHL { get; set; }
        public string ARCC { get; set; }
        public string ARCT { get; set; }
        public string SAPR { get; set; }
        public string ORQA { get; set; }
        public string MIHP { get; set; }
        public string MRTP { get; set; }
        public string UDAY { get; set; }
        public string WHLO { get; set; }
        public string LTYP { get; set; }
        public string ORQT { get; set; }
        public string ALQT { get; set; }
        public string SOQT { get; set; }
        public string DLQT { get; set; }
        public string REQ1 { get; set; }
        public string SUNO { get; set; }
        public string PUNO { get; set; }
        public string DOND { get; set; }
        public string DONR { get; set; }
        public string DOLR { get; set; }
        public string INVM { get; set; }
        public string NODT { get; set; }
        public string IIYR { get; set; }
        public string IIMO { get; set; }
        public string IIDA { get; set; }
        public string UCOS { get; set; }
        public string WXSP { get; set; }
        public string SINN { get; set; }
        public string MODE { get; set; }
        public string CIDF { get; set; }
        public string CIDT { get; set; }
        public string PRRF { get; set; }
        public string AGTP { get; set; }
        public string CMOD { get; set; }
        public string DOLD { get; set; }
        public string PUAG { get; set; }
        public string PUPR { get; set; }
        public string XHUN { get; set; }
        public string NOPR { get; set; }
        public string IPNO { get; set; }
        public string QRPO { get; set; }
        public string GRAC { get; set; }
        public string COAD { get; set; }
        public string STRT { get; set; }
        public string SUFI { get; set; }
    }
}