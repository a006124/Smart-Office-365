using System;
namespace SmartOffice365.Core.Models
{
    public class ContactEntity
    {
        public int Id { get; set; }
        public string Nom { get; set; } = string.Empty;
        public string Prenom { get; set; } = string.Empty;
        public string Role { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string Telephone { get; set; } = string.Empty;
        public string CompteTeams { get; set; } = string.Empty;
        public string CodeVendorSapLIFNR { get; set; } = string.Empty;
        public string Entreprise { get; set; } = string.Empty;
        public DateTime DateCreation { get; set; }
        public DateTime DateModification { get; set; }
        public string SharePointItemId { get; set; } = string.Empty;
    }
}
