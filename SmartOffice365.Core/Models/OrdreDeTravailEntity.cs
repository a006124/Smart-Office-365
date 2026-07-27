using System;
namespace SmartOffice365.Core.Models
{
    public class OrdreDeTravailEntity
    {
        public int Id { get; set; }
        public string NumeroOT_Aufnr { get; set; } = string.Empty;
        public string NumeroEquipement_EQUNR { get; set; } = string.Empty;
        public string PosteTechnique_TPLNR { get; set; } = string.Empty;
        public string PosteTravail_ARBPL { get; set; } = string.Empty;
        public string Titre { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public int Avancement { get; set; } // 0-100
        public string StatutShutdown { get; set; } = string.Empty; // À Faire, En cours, Bloqué, Terminé
        public string Priorite { get; set; } = string.Empty; // Critique, Haute, Normale, Basse
        public int AffaireId { get; set; }
        public string Affaire { get; set; } = string.Empty;
        public string Responsable { get; set; } = string.Empty;
        public string EntreprisePrestataire { get; set; } = string.Empty;
        public DateTime DateDebutPrevue { get; set; }
        public DateTime DateFinPrevue { get; set; }
        public DateTime? DateDebutReelle { get; set; }
        public DateTime? DateFinReelle { get; set; }
        public string MotifsBlockage { get; set; } = string.Empty;
        public DateTime DateCreation { get; set; }
        public DateTime DateModification { get; set; }
        public string SharePointItemId { get; set; } = string.Empty;
    }
}
