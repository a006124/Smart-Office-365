using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Me.SendMail;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Models;

namespace SmartOffice365.Core.Services
{
    public class OutlookReportService : IOutlookReportService
    {
        private readonly IGraphAuthService _authService;

        public OutlookReportService(IGraphAuthService authService)
        {
            _authService = authService;
        }

        public async Task SendDailyReportAsync(string[] recipients, DashboardKpis kpis, List<OrdreDeTravailEntity> otsEnRetard)
        {
            var htmlBody = BuildDailyReportHtml(kpis, otsEnRetard);
            await SendEmailAsync(
                recipients,
                $"[Smart Office 365] Rapport Journalier — {DateTime.Now:dd/MM/yyyy}",
                htmlBody);
        }

        public async Task SendCriticalOTAlertAsync(string[] recipients, OrdreDeTravailEntity ot)
        {
            var htmlBody = $@"
🚨 Alerte — OT Critique Bloqué
-----------------------------

|  |  |
| --- | --- |
| **N° OT (Aufnr)** | {ot.NumeroOT_Aufnr} |
| **Équipement (EQUNR)** | {ot.NumeroEquipement_EQUNR} |
| **Poste technique** | {ot.PosteTechnique_TPLNR} |
| **Priorité** | {ot.Priorite} |
| **Avancement** | {ot.Avancement}% |
| **Motif de blocage** | {ot.MotifsBlockage} |
| **Responsable** | {ot.Responsable} |
";
            await SendEmailAsync(recipients, $"[Smart Office 365] 🚨 OT Bloqué : {ot.NumeroOT_Aufnr} — {ot.Titre}", htmlBody);
        }

        private async Task SendEmailAsync(string[] recipients, string subject, string htmlContent)
        {
            var client = await _authService.GetAuthenticatedClientAsync();
            var toRecipients = recipients.Select(r => new Recipient
            {
                EmailAddress = new EmailAddress { Address = r }
            }).ToList();

            var message = new Message
            {
                Subject = subject,
                Body = new ItemBody { ContentType = BodyType.Html, Content = htmlContent },
                ToRecipients = toRecipients
            };

            await client.Me.SendMail.PostAsync(new SendMailPostRequestBody
            {
                Message = message,
                SaveToSentItems = true
            });
        }

        private string BuildDailyReportHtml(DashboardKpis kpis, List<OrdreDeTravailEntity> otsEnRetard)
        {
            var retardSection = "";
            if (otsEnRetard != null && otsEnRetard.Any())
            {
                var retardRows = string.Join("", otsEnRetard.Select(o =>
                    $"| {o.NumeroOT_Aufnr} | {o.Titre} | {o.Responsable} | {o.Avancement}% |\n"));

                retardSection = $@"
### ⚠️ OT en retard ({otsEnRetard.Count})

| N° OT | Titre | Responsable | Avancement |
| --- | --- | --- | --- |
{retardRows}
";
            }

            return $@"
Smart Office 365
================

Rapport Journalier — {DateTime.Now:dddd dd MMMM yyyy}
-----------------------------------------------------

{kpis.AvancementGlobal:F0}%
Avancement global

{kpis.OTBloques}
OT Bloqués

{kpis.OTEnRetard}
OT En retard

{kpis.OTTermines}
OT Terminés

{retardSection}
";
        }
    }
}
