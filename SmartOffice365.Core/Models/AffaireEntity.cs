using System;
namespace SmartOffice365.Core.Models
{
    public class AffaireEntity
    {
        public int Id { get; set; }
        public string Titre { get; set; } = string.Empty;
        public string CodeUniteSAP { get; set; } = string.Empty;
        public DateTime DateDebutPrevue { get; set; }
        public DateTime DateFinPrevue { get; set; }
        public string Statut { get; set; } = string.Empty; // Planifié, En cours, Terminé, Annulé
        public string Responsable { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public int AvancementGlobal { get; set; } // 0-100
        public DateTime DateCreation { get; set; }
        public string SharePointItemId { get; set; } = string.Empty;
    }
}
