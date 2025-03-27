using System.Web.Http; // For Web API configuration
using System.Web.Http.Cors; // For CORS
using TwinCATXmlGenerator.Services; // Namespace for your services
using Microsoft.Extensions.DependencyInjection; // Dependency Injection support
using System; // For Exception handling and Console usage

namespace TwinCATXmlGenerator
{
    public class Startup
    {
        /// <summary>
        /// Configures services for dependency injection.
        /// </summary>
        public void ConfigureServices(IServiceCollection services)
        {
            try
            {
                // Register services
                services.AddSingleton<ITwinCATService, TwinCATService>();
                services.AddSingleton<IXmlProducerService, XmlProducerService>();

                Console.WriteLine("Services configured successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error configuring services: {ex.Message}");
            }
        }

        /// <summary>
        /// Configures the Web API routes and other settings.
        /// </summary>
        public void Configure(HttpConfiguration config)
        {
            try
            {
                // Enable CORS
                var cors = new EnableCorsAttribute("*", "*", "*"); // Replace "*" with specific origins in production
                config.EnableCors(cors);

                // Configure routes
                config.MapHttpAttributeRoutes();
                config.Routes.MapHttpRoute(
                    name: "DefaultApi",
                    routeTemplate: "api/{controller}/{action}/{id}",
                    defaults: new { id = RouteParameter.Optional }
                );

                // Configure JSON formatting
                config.Formatters.JsonFormatter.SerializerSettings.Formatting = Newtonsoft.Json.Formatting.Indented;

                // Initialize configuration
                config.EnsureInitialized();

                Console.WriteLine("API configuration completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error configuring API: {ex.Message}");
            }
        }
    }
}