using System.Threading.Tasks;
using Microsoft.Graph;

namespace SmartOffice365.Core.Interfaces
{
    public interface IGraphAuthService
    {
        Task<GraphServiceClient> GetAuthenticatedClientAsync();
        Task<string> GetAccessTokenAsync();

        Task<bool> IsAuthenticatedAsync();
        Task<string> GetCurrentUserDisplayNameAsync();

        /// <summary>
        /// Met à jour l'URL SharePoint cible et réinitialise la connexion si le Tenant change
        /// </summary>
        void UpdateSharePointUrl(string url); // ◄--- AJOUTÉ
    }
}
