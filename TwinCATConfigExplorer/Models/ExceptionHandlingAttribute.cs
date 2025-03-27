using System.Net;
using System.Net.Http;
using System.Web.Http.Filters;

public class ExceptionHandlingAttribute : ExceptionFilterAttribute
{
    public override void OnException(HttpActionExecutedContext context)
    {
        // Log the exception (optional)
        // Your custom exception handling logic here
        context.Response = context.Request.CreateErrorResponse(
            HttpStatusCode.InternalServerError,
            "An unexpected error occurred. Please try again later."
        );
    }
}