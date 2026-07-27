using SmartOffice365.Core.Models;

namespace SmartOffice365.Core.Interfaces
{
    /// <summary>
    /// Service de génération et envoi de rapports par email Outlook
    /// </summary>
    public interface IOutlookReportService
    {
        /// <summary>Envoie le rapport quotidien d'avancement</summary>
        Task SendDailyReportAsync(string[] recipients, DashboardKpis kpis, List<OrdreDeTravailEntity> otsEnRetard);

        /// <summary>Envoie une alerte email pour un OT critique bloqué</summary>
        Task SendCriticalOTAlertAsync(string[] recipients, OrdreDeTravailEntity ot);
    }
}
