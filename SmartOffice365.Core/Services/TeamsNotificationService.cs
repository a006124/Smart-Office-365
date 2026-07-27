using Microsoft.Graph;
using Microsoft.Graph.Models;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Models;

namespace SmartOffice365.Core.Services
{
    public class TeamsNotificationService : ITeamsNotificationService
    {
        private readonly IGraphAuthService _authService;
        private readonly string _teamId;
        private readonly string _channelId;

        public TeamsNotificationService(IGraphAuthService authService, string teamId, string channelId)
        {
            _authService = authService;
            _teamId = teamId;
            _channelId = channelId;
        }

        public async Task SendOTBlockedAlertAsync(OrdreDeTravailEntity ot)
        {
            var adaptiveCardJson = $$"""
            {
              "type": "AdaptiveCard",
              "version": "1.4",
              "body": [
                { "type": "TextBlock", "text": "🚨 OT BLOQUÉ — {{ot.NumeroOT_Aufnr}}", "weight": "Bolder", "size": "Large", "color": "Attention" },
                { "type": "FactSet", "facts": [
                  { "title": "Intitulé", "value": "{{ot.Titre}}" },
                  { "title": "Équipement", "value": "{{ot.NumeroEquipement_EQUNR}}" },
                  { "title": "Priorité", "value": "{{ot.Priorite}}" },
                  { "title": "Avancement", "value": "{{ot.Avancement}}%" },
                  { "title": "Motif", "value": "{{ot.MotifsBlockage}}" },
                  { "title": "Responsable", "value": "{{ot.Responsable}}" }
                ]}
              ],
              "actions": [
                { "type": "Action.OpenUrl", "title": "Voir dans Smart Office 365", "url": "https://smart-office365/ot/{{ot.Id}}" }
              ]
            }
            """;

            await SendChannelMessageAsync(
                $"🚨 OT Bloqué : {ot.NumeroOT_Aufnr}",
                $"<attachment contentType=\"application/vnd.microsoft.card.adaptive\">{adaptiveCardJson}</attachment>");
        }

        public async Task SendChannelMessageAsync(string title, string body)
        {
            var client = await _authService.GetAuthenticatedClientAsync();
            var message = new ChatMessage
            {
                Subject = title,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = body
                }
            };
            await client.Teams[_teamId].Channels[_channelId].Messages.PostAsync(message);
        }

        public async Task SendDailyProgressCardAsync(DashboardKpis kpis)
        {
            var body = $"<b>📊 Rapport Journalier Smart Office 365</b><br/>OT Total: {kpis.TotalOT} | Avancement global: {kpis.AvancementGlobal:F1}%<br/>✅ Terminés: {kpis.OTTermines} | 🔵 En cours: {kpis.OTEnCours} | 🔴 Bloqués: {kpis.OTBloques} | ⚠️ En retard: {kpis.OTEnRetard}";
            await SendChannelMessageAsync("📊 Rapport journalier d'avancement", body);
        }
    }
}
