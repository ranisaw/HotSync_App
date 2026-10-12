using CustomerApi.Middleware;
using CustomerApi.Services;

var builder = WebApplication.CreateBuilder(args);

// ─────────────────────────────────────────────────────────────────────────────
// Configuration
// ─────────────────────────────────────────────────────────────────────────────
var connectionString = builder.Configuration["Database:ConnectionString"]
    ?? throw new InvalidOperationException(
        "Database:ConnectionString is missing from configuration. " +
        "Add it to appsettings.json under the 'Database' section.");

// ─────────────────────────────────────────────────────────────────────────────
// Services
// ─────────────────────────────────────────────────────────────────────────────
builder.Services.AddSingleton(new CustomerRepository(connectionString));
builder.Services.AddSingleton<ExcelExportService>();
builder.Services.AddSingleton<WordExportService>();

builder.Services.AddControllers();

// Swagger / OpenAPI — Swashbuckle 6.x with Microsoft.OpenApi 1.x
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
{
    options.SwaggerDoc("v1", new Microsoft.OpenApi.Models.OpenApiInfo
    {
        Title       = "Customer API",
        Version     = "v1",
        Description = "CRUD API for managing Fennex customer data synced from SharePoint into Azure SQL."
    });

    // Include XML comments from the generated documentation file
    var xmlFile = $"{System.Reflection.Assembly.GetExecutingAssembly().GetName().Name}.xml";
    var xmlPath = Path.Combine(AppContext.BaseDirectory, xmlFile);
    if (File.Exists(xmlPath))
        options.IncludeXmlComments(xmlPath);
});

// ─────────────────────────────────────────────────────────────────────────────
// Build
// ─────────────────────────────────────────────────────────────────────────────
var app = builder.Build();

// ─────────────────────────────────────────────────────────────────────────────
// Middleware pipeline
// ─────────────────────────────────────────────────────────────────────────────

// Global exception handler — must be first so it wraps everything
app.UseMiddleware<ErrorHandlingMiddleware>();

// Swagger UI — available in all environments for easy testing
app.UseSwagger();
app.UseSwaggerUI(c =>
{
    c.SwaggerEndpoint("/swagger/v1/swagger.json", "Customer API v1");
    c.RoutePrefix = "swagger";
});

app.UseDefaultFiles();
app.UseStaticFiles();

app.UseHttpsRedirection();
app.UseAuthorization();
app.MapControllers();

app.Run();
