namespace CustomerApi.Models;

/// <summary>
/// Represents a customer record as stored in the dbo.Customers SQL table.
/// This is the response DTO returned by the API.
/// </summary>
public class Customer
{
    /// <summary>Auto-generated SQL primary key.</summary>
    public int Id { get; set; }

    /// <summary>SharePoint list item ID — unique business key.</summary>
    public int SharePointId { get; set; }

    /// <summary>Customer or company name.</summary>
    public string? Name { get; set; }

    /// <summary>Customer address.</summary>
    public string? Address { get; set; }

    /// <summary>Number of oil rigs owned by the customer.</summary>
    public int? NumberOfRigs { get; set; }

    /// <summary>Number of jack-up rigs owned by the customer.</summary>
    public int? NumberOfJackUps { get; set; }

    /// <summary>Number of Mobile Offshore Drilling Units (MODUs).</summary>
    public int? NumberOfModus { get; set; }

    /// <summary>Comma-separated list of customer site URLs.</summary>
    public string? SiteUrls { get; set; }

    /// <summary>Timestamp of the last modification in SharePoint.</summary>
    public DateTime? SpLastModified { get; set; }

    /// <summary>Timestamp of the last sync from SharePoint to SQL.</summary>
    public DateTime SyncedAt { get; set; }

    /// <summary>Timestamp when the SQL row was first created.</summary>
    public DateTime CreatedAt { get; set; }
}
