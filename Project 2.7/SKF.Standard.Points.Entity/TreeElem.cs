namespace SKF.Standard.Points.Entity
{
    public class TreeElem
    {

        public uint TREEELEMID { get; set; }
        public uint HIERARCHYID { get; set; }

        public uint BRANCHLEVEL { get; set; }
        public uint SLOTNUMBER { get; set; }

        public uint TBLSETID { get; set; }

        public string NAME { get; set; }

        public uint CONTAINERTYPE { get; set; }

        public string DESCRIPTION { get; set; }


        public bool ELEMENTENABLE { get; set; }
        public bool PARENTENABLE { get; set; }

        public uint HIERARCHYTYPE { get; set; }
        public uint ALARMFLAGS { get; set; }
        public uint PARENTID { get; set; }
        public uint PARENTREFID { get; set; }
        public uint REFERENCEID { get; set; }

        public int GOOD { get; set; }

        public int ALERT { get; set; }
        public int DANGER { get; set; }
        public int OVERDUE { get; set; }
        public bool CHANNELENABLE { get; set; }
    }
}
