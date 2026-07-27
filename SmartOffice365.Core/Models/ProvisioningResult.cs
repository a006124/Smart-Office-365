namespace SmartOffice365.Core.Models
{
    /// <summary>
    /// Résultat d'une opération de provisioning SharePoint
    /// </summary>
    public class ProvisioningResult
    {
        public bool Success { get; set; }
        public List<string> ActionsPerformed { get; set; } = new();
        public List<string> Errors { get; set; } = new();
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }
}
