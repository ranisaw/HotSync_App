using CustomerApi.Models;
using Microsoft.Data.SqlClient;

namespace CustomerApi.Services;

/// <summary>
/// Handles all database operations for the Customers table using raw ADO.NET.
/// No ORM is used — all queries are parameterised SqlCommand instances.
/// </summary>
public class CustomerRepository
{
    private readonly string _connectionString;

    /// <summary>
    /// Initialises a new instance of <see cref="CustomerRepository"/>.
    /// </summary>
    /// <param name="connectionString">Azure SQL connection string read from configuration.</param>
    public CustomerRepository(string connectionString)
    {
        _connectionString = connectionString;
    }

    // -------------------------------------------------------------------------
    // Private helpers
    // -------------------------------------------------------------------------

    /// <summary>
    /// Maps a single <see cref="SqlDataReader"/> row to a <see cref="Customer"/> object.
    /// </summary>
    private static Customer MapRow(SqlDataReader reader)
    {
        return new Customer
        {
            Id              = reader.GetInt32(reader.GetOrdinal("Id")),
            SharePointId    = reader.GetInt32(reader.GetOrdinal("SharePointId")),
            Name            = reader.IsDBNull(reader.GetOrdinal("Name"))            ? null : reader.GetString(reader.GetOrdinal("Name")),
            Address         = reader.IsDBNull(reader.GetOrdinal("Address"))         ? null : reader.GetString(reader.GetOrdinal("Address")),
            NumberOfRigs    = reader.IsDBNull(reader.GetOrdinal("NumberOfRigs"))    ? null : reader.GetInt32(reader.GetOrdinal("NumberOfRigs")),
            NumberOfJackUps = reader.IsDBNull(reader.GetOrdinal("NumberOfJackUps")) ? null : reader.GetInt32(reader.GetOrdinal("NumberOfJackUps")),
            NumberOfModus   = reader.IsDBNull(reader.GetOrdinal("NumberOfModus"))   ? null : reader.GetInt32(reader.GetOrdinal("NumberOfModus")),
            SiteUrls        = reader.IsDBNull(reader.GetOrdinal("SiteUrls"))        ? null : reader.GetString(reader.GetOrdinal("SiteUrls")),
            SpLastModified  = reader.IsDBNull(reader.GetOrdinal("SpLastModified"))  ? null : reader.GetDateTime(reader.GetOrdinal("SpLastModified")),
            SyncedAt        = reader.GetDateTime(reader.GetOrdinal("SyncedAt")),
            CreatedAt       = reader.GetDateTime(reader.GetOrdinal("CreatedAt")),
        };
    }

    // -------------------------------------------------------------------------
    // CRUD operations
    // -------------------------------------------------------------------------

    /// <summary>
    /// Retrieves all customers from the database, ordered by Id ascending.
    /// </summary>
    /// <returns>A list of all <see cref="Customer"/> records.</returns>
    public async Task<List<Customer>> GetAllAsync()
    {
        const string sql = @"
            SELECT Id, SharePointId, Name, Address,
                   NumberOfRigs, NumberOfJackUps, NumberOfModus,
                   SiteUrls, SpLastModified, SyncedAt, CreatedAt
            FROM dbo.Customers
            ORDER BY Id ASC;";

        var customers = new List<Customer>();

        await using var conn = new SqlConnection(_connectionString);
        await conn.OpenAsync();

        await using var cmd = new SqlCommand(sql, conn);
        await using var reader = await cmd.ExecuteReaderAsync();

        while (await reader.ReadAsync())
            customers.Add(MapRow(reader));

        return customers;
    }

    /// <summary>
    /// Retrieves a single customer by their SQL primary key Id.
    /// </summary>
    /// <param name="id">The SQL primary key of the customer.</param>
    /// <returns>The <see cref="Customer"/> if found; otherwise <c>null</c>.</returns>
    public async Task<Customer?> GetByIdAsync(int id)
    {
        const string sql = @"
            SELECT Id, SharePointId, Name, Address,
                   NumberOfRigs, NumberOfJackUps, NumberOfModus,
                   SiteUrls, SpLastModified, SyncedAt, CreatedAt
            FROM dbo.Customers
            WHERE Id = @Id;";

        await using var conn = new SqlConnection(_connectionString);
        await conn.OpenAsync();

        await using var cmd = new SqlCommand(sql, conn);
        cmd.Parameters.AddWithValue("@Id", id);

        await using var reader = await cmd.ExecuteReaderAsync();

        if (await reader.ReadAsync())
            return MapRow(reader);

        return null;
    }

    /// <summary>
    /// Inserts a new customer row and returns the newly assigned SQL Id.
    /// </summary>
    /// <param name="request">The validated create request body.</param>
    /// <returns>The new customer's SQL Id.</returns>
    public async Task<int> CreateAsync(CreateCustomerRequest request)
    {
        const string sql = @"
            INSERT INTO dbo.Customers
                (SharePointId, Name, Address, NumberOfRigs, NumberOfJackUps, NumberOfModus,
                 SiteUrls, SpLastModified, SyncedAt, CreatedAt)
            VALUES
                (@SharePointId, @Name, @Address, @NumberOfRigs, @NumberOfJackUps, @NumberOfModus,
                 @SiteUrls, @SpLastModified, GETUTCDATE(), GETUTCDATE());
            SELECT SCOPE_IDENTITY();";

        await using var conn = new SqlConnection(_connectionString);
        await conn.OpenAsync();

        await using var cmd = new SqlCommand(sql, conn);
        cmd.Parameters.AddWithValue("@SharePointId",    request.SharePointId);
        cmd.Parameters.AddWithValue("@Name",            request.Name);
        cmd.Parameters.AddWithValue("@Address",         (object?)request.Address         ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@NumberOfRigs",    (object?)request.NumberOfRigs    ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@NumberOfJackUps", (object?)request.NumberOfJackUps ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@NumberOfModus",   (object?)request.NumberOfModus   ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@SiteUrls",        (object?)request.SiteUrls        ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@SpLastModified",  (object?)request.SpLastModified  ?? DBNull.Value);

        var result = await cmd.ExecuteScalarAsync();
        return Convert.ToInt32(result);
    }

    /// <summary>
    /// Updates an existing customer record. Returns the number of rows affected (0 = not found).
    /// </summary>
    /// <param name="id">The SQL primary key of the customer to update.</param>
    /// <param name="request">The validated update request body.</param>
    /// <returns>Number of rows affected; 0 means the Id did not exist.</returns>
    public async Task<int> UpdateAsync(int id, UpdateCustomerRequest request)
    {
        const string sql = @"
            UPDATE dbo.Customers
            SET Name            = @Name,
                Address         = @Address,
                NumberOfRigs    = @NumberOfRigs,
                NumberOfJackUps = @NumberOfJackUps,
                NumberOfModus   = @NumberOfModus,
                SiteUrls        = @SiteUrls,
                SpLastModified  = @SpLastModified,
                SyncedAt        = GETUTCDATE()
            WHERE Id = @Id;";

        await using var conn = new SqlConnection(_connectionString);
        await conn.OpenAsync();

        await using var cmd = new SqlCommand(sql, conn);
        cmd.Parameters.AddWithValue("@Id",              id);
        cmd.Parameters.AddWithValue("@Name",            request.Name);
        cmd.Parameters.AddWithValue("@Address",         (object?)request.Address         ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@NumberOfRigs",    (object?)request.NumberOfRigs    ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@NumberOfJackUps", (object?)request.NumberOfJackUps ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@NumberOfModus",   (object?)request.NumberOfModus   ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@SiteUrls",        (object?)request.SiteUrls        ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@SpLastModified",  (object?)request.SpLastModified  ?? DBNull.Value);

        return await cmd.ExecuteNonQueryAsync();
    }

    /// <summary>
    /// Deletes a customer by their SQL primary key. Returns rows affected (0 = not found).
    /// </summary>
    /// <param name="id">The SQL primary key of the customer to delete.</param>
    /// <returns>Number of rows deleted; 0 means the Id did not exist.</returns>
    public async Task<int> DeleteAsync(int id)
    {
        const string sql = "DELETE FROM dbo.Customers WHERE Id = @Id;";

        await using var conn = new SqlConnection(_connectionString);
        await conn.OpenAsync();

        await using var cmd = new SqlCommand(sql, conn);
        cmd.Parameters.AddWithValue("@Id", id);

        return await cmd.ExecuteNonQueryAsync();
    }
}
