using System.ComponentModel.DataAnnotations;

namespace CustomerApi.Models;

/// <summary>
/// Request body for creating a new customer via POST /api/customers.
/// </summary>
public class CreateCustomerRequest
{
    /// <summary>
    /// SharePoint list item ID — must be unique across all customers.
    /// </summary>
    [Required(ErrorMessage = "SharePointId is required.")]
    public int SharePointId { get; set; }

    /// <summary>
    /// Customer or company name. Required.
    /// </summary>
    [Required(ErrorMessage = "Name is required.")]
    [StringLength(255, ErrorMessage = "Name cannot exceed 255 characters.")]
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Customer address. Optional, max 1000 characters.
    /// </summary>
    [StringLength(1000, ErrorMessage = "Address cannot exceed 1000 characters.")]
    public string? Address { get; set; }

    /// <summary>
    /// Number of oil rigs. Must be 0 or greater if provided.
    /// </summary>
    [Range(0, int.MaxValue, ErrorMessage = "NumberOfRigs must be 0 or greater.")]
    public int? NumberOfRigs { get; set; }

    /// <summary>
    /// Number of jack-up rigs. Must be 0 or greater if provided.
    /// </summary>
    [Range(0, int.MaxValue, ErrorMessage = "NumberOfJackUps must be 0 or greater.")]
    public int? NumberOfJackUps { get; set; }

    /// <summary>
    /// Number of MODUs. Must be 0 or greater if provided.
    /// </summary>
    [Range(0, int.MaxValue, ErrorMessage = "NumberOfModus must be 0 or greater.")]
    public int? NumberOfModus { get; set; }

    /// <summary>
    /// Comma-separated list of customer site URLs.
    /// </summary>
    public string? SiteUrls { get; set; }

    /// <summary>
    /// Timestamp of the last modification in SharePoint.
    /// </summary>
    public DateTime? SpLastModified { get; set; }
}
