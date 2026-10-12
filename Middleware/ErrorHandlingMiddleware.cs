using System.Net;
using System.Text.Json;

namespace CustomerApi.Middleware;

/// <summary>
/// Global exception handling middleware that catches unhandled exceptions
/// and returns a safe, structured JSON error response. Raw stack traces
/// are never exposed to the API caller.
/// </summary>
public class ErrorHandlingMiddleware
{
    private readonly RequestDelegate _next;
    private readonly ILogger<ErrorHandlingMiddleware> _logger;

    /// <summary>
    /// Initialises a new instance of <see cref="ErrorHandlingMiddleware"/>.
    /// </summary>
    public ErrorHandlingMiddleware(RequestDelegate next, ILogger<ErrorHandlingMiddleware> logger)
    {
        _next   = next;
        _logger = logger;
    }

    /// <summary>
    /// Invokes the middleware. Wraps the next delegate in a try/catch
    /// and converts any unhandled exception to a 500 JSON response.
    /// </summary>
    public async Task InvokeAsync(HttpContext context)
    {
        try
        {
            await _next(context);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unhandled exception for request {Method} {Path}.",
                context.Request.Method, context.Request.Path);

            await WriteErrorResponseAsync(context, ex);
        }
    }

    private static async Task WriteErrorResponseAsync(HttpContext context, Exception ex)
    {
        context.Response.StatusCode  = (int)HttpStatusCode.InternalServerError;
        context.Response.ContentType = "application/json";

        var payload = new
        {
            error   = "An unexpected error occurred. Please try again later.",
            traceId = context.TraceIdentifier
        };

        var json = JsonSerializer.Serialize(payload);
        await context.Response.WriteAsync(json);
    }
}
