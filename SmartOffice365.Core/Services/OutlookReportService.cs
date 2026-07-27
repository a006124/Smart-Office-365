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
            var htmlBody = $"""
            <html><body style="font-family: Segoe UI, sans-serif; background:#f5f5f5; padding:20px;">
            <div style="background:#fff; border-radius:12px; padding:24px; border-left:4px solid #ef4444;">
              <h2 style="color:#ef4444;">🚨 Alerte — OT Critique Bloqué</h2>
              <table style="width:100%; border-collapse:collapse;">
                <tr><td><b>N° OT (Aufnr)</b></td><td>{ot.NumeroOT_Aufnr}</td></tr>
                <tr><td><b>Équipement (EQUNR)</b></td><td>{ot.NumeroEquipement_EQUNR}</td></tr>
                <tr><td><b>Poste technique</b></td><td>{ot.PosteTechnique_TPLNR}</td></tr>
                <tr><td><b>Priorité</b></td><td style="color:#ef4444;font-weight:bold;">{ot.Priorite}</td></tr>
                <tr><td><b>Avancement</b></td><td>{ot.Avancement}%</td></tr>
                <tr><td><b>Motif de blocage</b></td><td>{ot.MotifsBlockage}</td></tr>
                <tr><td><b>Responsable</b></td><td>{ot.Responsable}</td></tr>
              </table>
            </div></body></html>
            """;
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
            var retardRows = string.Join("", otsEnRetard.Select(o =>
                $"<tr><td>{o.NumeroOT_Aufnr}</td><td>{o.Titre}</td><td>{o.Responsable}</td><td style='color:#ef4444;'>{o.Avancement}%</td></tr>"));

            return $"""
            <html><body style="font-family: Segoe UI, sans-serif; background:#0f0f10; color:#e3e2e6; padding:20px;">
            <div style="max-width:700px; margin:auto; background:#1e1f20; border-radius:16px; padding:32px;">
              <h1 style="background:linear-gradient(135deg,#7cacf8,#e8b3ff); -webkit-background-clip:text; -webkit-text-fill-color:transparent;">Smart Office 365</h1>
              <h2>Rapport Journalier — {DateTime.Now:dddd dd MMMM yyyy}</h2>
              <div style="display:flex; gap:16px; margin:24px 0;">
                <div style="flex:1; background:#28292a; border-radius:12px; padding:16px; text-align:center;">
                  <div style="font-size:36px; font-weight:bold; color:#a8c7fa;">{kpis.AvancementGlobal:F0}%</div>
                  <div style="color:#8e918f;">Avancement global</div>
                </div>
                <div style="flex:1; background:#28292a; border-radius:12px; padding:16px; text-align:center;">
                  <div style="font-size:36px; font-weight:bold; color:#ef4444;">{kpis.OTBloques}</div>
                  <div style="color:#8e918f;">OT Bloqués</div>
                </div>
                <div style="flex:1; background:#28292a; border-radius:12px; padding:16px; text-align:center;">
                  <div style="font-size:36px; font-weight:bold; color:#f59e0b;">{kpis.OTEnRetard}</div>
                  <div style="color:#8e918f;">OT En retard</div>
                </div>
                <div style="flex:1; background:#28292a; border-radius:12px; padding:16px; text-align:center;">
                  <div style="font-size:36px; font-weight:bold; color:#22c55e;">{kpis.OTTermines}</div>
                  <div style="color:#8e918f;">OT Terminés</div>
                </div>
              </div>
              {(otsEnRetard.Any() ? $"""<h3 style='color:#f59e0b;'>⚠️ OT en retard ({otsEnRetard.Count})</h3>
              <table style='width:100%; border-collapse:collapse; background:#28292a; border-radius:8px;'>
              <thead><tr style='color:#8e918f;'><th>N° OT</th><th>Titre</th><th>Responsable</th><th>Avancement</th></tr></thead>
              <tbody>{retardRows}</tbody></table>""" : "")}
            </div></body></html>
            """;
        }
    }
}
