namespace FastMailSearch.Core.Models
{
    public enum IndexerHealth
    {
        Healthy,
        Starting,
        Indexing,
        Updating,
        Paused,
        Degraded,
        SafeMode,
        Error
    }
}
