using CustomerApi.Models;
using CustomerApi.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace CustomerApi.Controllers;

/// <summary>
/// Provides full CRUD operations for the Customers resource.
/// </summary>
[ApiController]
[Route("api/[controller]")]
[Produces("application/json")]
public class CustomersController : ControllerBase
{
    private readonly CustomerRepository _repository;
    private readonly ExcelExportService _excelExportService;
    private readonly WordExportService _wordExportService;
    private readonly ILogger<CustomersController> _logger;

    /// <summary>
    /// Initialises a new instance of <see cref="CustomersController"/>.
    /// </summary>
    public CustomersController(
        CustomerRepository repository,
        ExcelExportService excelExportService,
        WordExportService wordExportService,
        ILogger<CustomersController> logger)
    {
        _repository         = repository;
        _excelExportService = excelExportService;
        _wordExportService  = wordExportService;
        _logger             = logger;
    }

    // -------------------------------------------------------------------------
    // GET /api/customers
    // -------------------------------------------------------------------------

    /// <summary>
    /// Returns all customers ordered by Id ascending.
    /// </summary>
    /// <returns>200 OK with a JSON array of customers.</returns>
    [HttpGet]
    [ProducesResponseType(typeof(List<Customer>), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> GetAll()
    {
        try
        {
            var customers = await _repository.GetAllAsync();
            return Ok(customers);
        }
        catch (SqlException ex)
        {
            _logger.LogError(ex, "SQL error retrieving all customers.");
            return StatusCode(StatusCodes.Status500InternalServerError, new { error = "A database error occurred." });
        }
    }

    // -------------------------------------------------------------------------
    // GET /api/customers/{id}
    // -------------------------------------------------------------------------

    /// <summary>
    /// Returns a single customer by their SQL primary key.
    /// </summary>
    /// <param name="id">The SQL primary key of the customer.</param>
    /// <returns>200 OK with the customer, or 404 Not Found.</returns>
    [HttpGet("{id:int}")]
    [ProducesResponseType(typeof(Customer), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> GetById(int id)
    {
        try
        {
            var customer = await _repository.GetByIdAsync(id);

            if (customer is null)
                return NotFound(new { error = $"Customer with Id {id} was not found." });

            return Ok(customer);
        }
        catch (SqlException ex)
        {
            _logger.LogError(ex, "SQL error retrieving customer Id={Id}.", id);
            return StatusCode(StatusCodes.Status500InternalServerError, new { error = "A database error occurred." });
        }
    }

    // -------------------------------------------------------------------------
    // POST /api/customers
    // -------------------------------------------------------------------------

    /// <summary>
    /// Creates a new customer record.
    /// </summary>
    /// <param name="request">JSON body with customer fields. Name and SharePointId are required.</param>
    /// <returns>201 Created with a Location header pointing to the new resource.</returns>
    [HttpPost]
    [ProducesResponseType(typeof(Customer), StatusCodes.Status201Created)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    [ProducesResponseType(StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> Create([FromBody] CreateCustomerRequest request)
    {
        if (!ModelState.IsValid)
            return BadRequest(ModelState);

        try
        {
            var newId    = await _repository.CreateAsync(request);
            var customer = await _repository.GetByIdAsync(newId);

            return CreatedAtAction(nameof(GetById), new { id = newId }, customer);
        }
        catch (SqlException ex) when (ex.Number == 2627 || ex.Number == 2601)
        {
            // Unique constraint violation on SharePointId
            _logger.LogWarning(ex, "Duplicate SharePointId {SharePointId} on create.", request.SharePointId);
            return Conflict(new { error = $"A customer with SharePointId {request.SharePointId} already exists." });
        }
        catch (SqlException ex)
        {
            _logger.LogError(ex, "SQL error creating customer.");
            return StatusCode(StatusCodes.Status500InternalServerError, new { error = "A database error occurred." });
        }
    }

    // -------------------------------------------------------------------------
    // PUT /api/customers/{id}
    // -------------------------------------------------------------------------

    /// <summary>
    /// Updates an existing customer by their SQL primary key.
    /// </summary>
    /// <param name="id">The SQL primary key of the customer to update.</param>
    /// <param name="request">JSON body with updated customer fields.</param>
    /// <returns>200 OK with the updated customer, or 404 Not Found.</returns>
    [HttpPut("{id:int}")]
    [ProducesResponseType(typeof(Customer), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> Update(int id, [FromBody] UpdateCustomerRequest request)
    {
        if (!ModelState.IsValid)
            return BadRequest(ModelState);

        try
        {
            var rowsAffected = await _repository.UpdateAsync(id, request);

            if (rowsAffected == 0)
                return NotFound(new { error = $"Customer with Id {id} was not found." });

            var updated = await _repository.GetByIdAsync(id);
            return Ok(updated);
        }
        catch (SqlException ex)
        {
            _logger.LogError(ex, "SQL error updating customer Id={Id}.", id);
            return StatusCode(StatusCodes.Status500InternalServerError, new { error = "A database error occurred." });
        }
    }

    // -------------------------------------------------------------------------
    // DELETE /api/customers/{id}
    // -------------------------------------------------------------------------

    /// <summary>
    /// Deletes a customer by their SQL primary key.
    /// </summary>
    /// <param name="id">The SQL primary key of the customer to delete.</param>
    /// <returns>204 No Content on success, or 404 Not Found.</returns>
    [HttpDelete("{id:int}")]
    [ProducesResponseType(StatusCodes.Status204NoContent)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> Delete(int id)
    {
        try
        {
            var rowsAffected = await _repository.DeleteAsync(id);

            if (rowsAffected == 0)
                return NotFound(new { error = $"Customer with Id {id} was not found." });

            return NoContent();
        }
        catch (SqlException ex)
        {
            _logger.LogError(ex, "SQL error deleting customer Id={Id}.", id);
            return StatusCode(StatusCodes.Status500InternalServerError, new { error = "A database error occurred." });
        }
    }

    // -------------------------------------------------------------------------
    // TASK 7: GET /api/customers/export/excel
    // -------------------------------------------------------------------------

    /// <summary>
    /// Exports all customer records in tabular layout to an Excel (.xlsx) file using OpenXML.
    /// </summary>
    /// <returns>A downloadable .xlsx file stream.</returns>
    [HttpGet("export/excel")]
    [Produces("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")]
    [ProducesResponseType(typeof(FileResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> ExportToExcel()
    {
        try
        {
            var customers = await _repository.GetAllAsync();
            var fileBytes = _excelExportService.ExportCustomersToExcel(customers);

            var fileName = $"Fennex_Customers_Export_{DateTime.UtcNow:yyyyMMdd_HHmmss}.xlsx";
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        catch (SqlException ex)
        {
            _logger.LogError(ex, "SQL error while exporting customers to Excel.");
            return StatusCode(StatusCodes.Status500InternalServerError, new { error = "A database error occurred." });
        }
    }

    // -------------------------------------------------------------------------
    // TASK 8: GET /api/customers/export/word
    // -------------------------------------------------------------------------

    /// <summary>
    /// Exports all customer records, executive summaries, and fleet tables to a Word (.docx) document using OpenXML.
    /// </summary>
    /// <returns>A downloadable .docx file stream.</returns>
    [HttpGet("export/word")]
    [Produces("application/vnd.openxmlformats-officedocument.wordprocessingml.document")]
    [ProducesResponseType(typeof(FileResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> ExportToWord()
    {
        try
        {
            var customers = await _repository.GetAllAsync();
            var fileBytes = _wordExportService.ExportCustomersToWord(customers);

            var fileName = $"Fennex_Fleet_Report_{DateTime.UtcNow:yyyyMMdd_HHmmss}.docx";
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
        }
        catch (SqlException ex)
        {
            _logger.LogError(ex, "SQL error while exporting customers to Word.");
            return StatusCode(StatusCodes.Status500InternalServerError, new { error = "A database error occurred." });
        }
    }

    // -------------------------------------------------------------------------
    // TASK 8 (Extended): GET /api/customers/{id}/export/word
    // -------------------------------------------------------------------------

    /// <summary>
    /// Exports an individual customer's executive profile to a Word (.docx) document using OpenXML.
    /// </summary>
    /// <param name="id">The SQL primary key of the customer.</param>
    /// <returns>A downloadable .docx file stream, or 404 Not Found.</returns>
    [HttpGet("{id:int}/export/word")]
    [Produces("application/vnd.openxmlformats-officedocument.wordprocessingml.document")]
    [ProducesResponseType(typeof(FileResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> ExportCustomerProfileToWord(int id)
    {
        try
        {
            var customer = await _repository.GetByIdAsync(id);
            if (customer is null)
                return NotFound(new { error = $"Customer with Id {id} was not found." });

            var fileBytes = _wordExportService.ExportSingleCustomerToWord(customer);
            var safeName = string.Join("_", (customer.Name ?? "Customer").Split(Path.GetInvalidFileNameChars())).Replace(" ", "_");
            var fileName = $"Fennex_Customer_{safeName}_{id}.docx";

            return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
        }
        catch (SqlException ex)
        {
            _logger.LogError(ex, "SQL error exporting customer profile Id={Id} to Word.", id);
            return StatusCode(StatusCodes.Status500InternalServerError, new { error = "A database error occurred." });
        }
    }
}
