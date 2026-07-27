using Microsoft.Graph;

namespace SmartOffice365.Core.Interfaces
{
    /// <summary>
    /// Service d'authentification Microsoft 365 interactive (compte utilisateur direct sans ClientId)
    /// </summary>
    public interface IGraphAuthService
    {
        /// <summary>Obtient un client Graph authentifié via le compte Office 365 connecté</summary>
        Task<GraphServiceClient> GetAuthenticatedClientAsync();

        /// <summary>Déclenche la connexion interactive avec le compte Office 365</summary>
        Task<bool> SignInAsync();

        /// <summary>Déconnecte l'utilisateur courant</summary>
        Task SignOutAsync();

        /// <summary>Vérifie si une session valide existe</summary>
        Task<bool> IsAuthenticatedAsync();

        /// <summary>Retourne le nom de l'utilisateur connecté</summary>
        Task<string> GetCurrentUserDisplayNameAsync();
    }
}
