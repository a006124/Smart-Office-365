using System;
namespace SmartOffice365.Core.Models
{
    public class HabilitationEntity
    {
        public int Id { get; set; }
        public int ContactId { get; set; }
        public string NomContact { get; set; } = string.Empty;
        public string TypeHabilitation { get; set; } = string.Empty; // CACES, Électrique B1, Travail en hauteur...
        public string Niveau { get; set; } = string.Empty;
        public DateTime DateObtention { get; set; }
        public DateTime DateExpiration { get; set; }
        public bool EstValide => DateTime.Now < DateExpiration;
        public string Organisme { get; set; } = string.Empty;
        public string Commentaire { get; set; } = string.Empty;
        public string SharePointItemId { get; set; } = string.Empty;
    }
}
