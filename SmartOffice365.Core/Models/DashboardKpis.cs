namespace SmartOffice365.Core.Models
{
    public class DashboardKpis
    {
        public int TotalOT { get; set; }
        public int OTTermines { get; set; }
        public int OTEnCours { get; set; }
        public int OTBloques { get; set; }
        public int OTEnRetard { get; set; }
        public double AvancementGlobal { get; set; }
        public int TotalContacts { get; set; }
        public int TotalAffaires { get; set; }
    }
}
