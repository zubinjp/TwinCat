﻿using System.Web.Http;
using System.Web.Http.Filters;

public static class WebApiConfig
{
    public static void Register(HttpConfiguration config)
    {
        // Enable attribute-based routing (e.g., [Route("api/xml/generate")])
        config.MapHttpAttributeRoutes();

        // Define default API routes
        config.Routes.MapHttpRoute(
            name: "DefaultApi",
            routeTemplate: "api/{controller}/{action}/{id}",
            defaults: new { id = RouteParameter.Optional }
        );

        // Enable JSON formatting by default
        config.Formatters.JsonFormatter.SerializerSettings.Formatting = Newtonsoft.Json.Formatting.Indented;

        // Allow CORS (Cross-Origin Resource Sharing) for external access
        config.EnableCors();

        // Configure global exception handling
        // Ensure the custom ExceptionHandlingAttribute is correctly implemented and accessible
        config.Filters.Add(new ExceptionHandlingAttribute());

        // Ensure the configuration is initialized
        config.EnsureInitialized();

        // Add other global configurations if required in the future
    }
}