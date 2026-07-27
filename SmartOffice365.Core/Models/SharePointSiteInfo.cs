namespace SmartOffice365.Core.Models
{
    /// <summary>
    /// Information sur un site SharePoint
    /// </summary>
    public class SharePointSiteInfo
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string WebUrl { get; set; } = string.Empty;
    }
}