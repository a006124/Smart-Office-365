using System;

namespace SmartOffice365.Core.Models
{
    /// <summary>
    /// Représente les indicateurs clés de performance (KPI) pour le tableau de bord, Outlook et Teams
    /// </summary>
    public class DashboardKpis
    {
        // --- Propriétés requises pour le rapport Outlook ---

        /// <summary>Pourcentage d'avancement global (ex: 75.5)</summary>
        public double AvancementGlobal { get; set; }

        /// <summary>Nombre d'ordres de travail bloqués</summary>
        public int OTBloques { get; set; }

        /// <summary>Nombre d'ordres de travail en retard</summary>
        public int OTEnRetard { get; set; }

        /// <summary>Nombre d'ordres de travail terminés</summary>
        public int OTTermines { get; set; }


        // --- Propriétés requises pour les notifications Teams ---

        /// <summary>Nombre total d'ordres de travail</summary>
        public int TotalOT { get; set; }

        /// <summary>Nombre d'ordres de travail en cours</summary>
        public int OTEnCours { get; set; }


        // --- Autres indicateurs globaux ---

        public int TotalContacts { get; set; }
        public int TotalAffaires { get; set; }
        public int TotalRessources { get; set; }
    }
}
