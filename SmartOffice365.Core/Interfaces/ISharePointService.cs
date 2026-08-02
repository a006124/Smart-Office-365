using System.Collections.Generic;
using System.Threading.Tasks;
using SmartOffice365.Core.Models;


namespace SmartOffice365.UI.ViewModels
{
    public interface ISharePointService
    {
        Task<List<ArretModel>> GetArretsAsync();
        Task<ArretModel> CreateArretAsync(ArretModel arret);
        Task UpdateArretAsync(ArretModel arret);
        Task DeleteArretAsync(int id);
    }
}
