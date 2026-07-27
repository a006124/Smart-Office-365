using System;
namespace SmartOffice365.Core.Models
{
    public class PrerequsEntity
    {
        public int Id { get; set; }
        public int OrdreDeTravailId { get; set; }
        public string NumeroOT { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty; // Consignation électrique, Permis de feu, Permis de travail, ATEX...
        public bool EstValide { get; set; }
        public DateTime DateValidation { get; set; }
        public string Signataire { get; set; } = string.Empty;
        public string Commentaire { get; set; } = string.Empty;
        public DateTime DateExpiration { get; set; }
        public DateTime DateCreation { get; set; }
        public string SharePointItemId { get; set; } = string.Empty;
    }
}
