using SmartOffice365.Core.Models;

namespace SmartOffice365.Core.Interfaces
{
    /// <summary>
    /// Service d'envoi de notifications Teams via Adaptive Cards
    /// </summary>
    public interface ITeamsNotificationService
    {
        /// <summary>Envoie une Adaptive Card d'alerte pour un OT bloqué</summary>
        Task SendOTBlockedAlertAsync(OrdreDeTravailEntity ot);

        /// <summary>Envoie un message général dans le canal configuré</summary>
        Task SendChannelMessageAsync(string title, string body);

        /// <summary>Envoie une Adaptive Card d'avancement journalier</summary>
        Task SendDailyProgressCardAsync(DashboardKpis kpis);
    }
}
