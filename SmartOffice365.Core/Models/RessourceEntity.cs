using System;
namespace SmartOffice365.Core.Models
{
    public class RessourceEntity
    {
        public int Id { get; set; }
        public int OrdreDeTravailId { get; set; }
        public string NumeroOT { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty; // Main d'œuvre, Matériel, Outillage spécial
        public string EntreprisePrestataire { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public int QuantitePrevue { get; set; }
        public int QuantiteReelle { get; set; }
        public string Unite { get; set; } = string.Empty; // h, pcs, kg...
        public bool EstDisponible { get; set; }
        public string Commentaire { get; set; } = string.Empty;
        public DateTime DateCreation { get; set; }
        public string SharePointItemId { get; set; } = string.Empty;
    }
}
