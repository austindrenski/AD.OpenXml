using System;
using System.IO;
using System.Linq;
using AD.ApiExtensions;
using JetBrains.Annotations;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ApiExplorer;
using Microsoft.AspNetCore.ResponseCompression;
using Microsoft.DotNet.PlatformAbstractions;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Swashbuckle.AspNetCore.Swagger;
using Swashbuckle.AspNetCore.SwaggerUI;

namespace CompilerAPI
{
    /// <summary>
    /// Represents the startup configuration for a web host.
    /// </summary>
    [PublicAPI]
    public class Startup : IStartup
    {
        /// <summary>
        /// Represents the application configuration properties available during construction.
        /// </summary>
        [NotNull] readonly IConfiguration _configuration;

        /// <summary>
        /// Constructs an <see cref="IStartup"/> for configuration.
        /// </summary>
        /// <param name="configuration">
        /// The application configuration properties available from the <see cref="IWebHostBuilder"/>.
        /// </param>
        /// <exception cref="ArgumentNullException"><paramref name="configuration"/></exception>
        public Startup([NotNull] IConfiguration configuration) => _configuration = configuration;

        /// <summary>
        /// This method gets called by the runtime. Use this method to add services to the container.
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="services"/></exception>
        [Pure]
        [NotNull]
        public IServiceProvider ConfigureServices([NotNull] IServiceCollection services)
            => services.AddLogging(x => x.AddConsole())
                       .AddResponseCompression(x => x.Providers.Add<GzipCompressionProvider>())
                       .AddRouting(x => x.LowercaseUrls = true)
                       .AddAntiforgery(
                           x =>
                           {
                               x.HeaderName = "x-xsrf-token";
                               x.Cookie.HttpOnly = true;
                               x.Cookie.SecurePolicy = CookieSecurePolicy.SameAsRequest;
                           })
                       .AddApiVersioning(
                           x =>
                           {
                               x.AssumeDefaultVersionWhenUnspecified = true;
                               x.DefaultApiVersion = new ApiVersion(2, 0);
                               x.ReportApiVersions = true;
                           })
                       .AddVersionedApiExplorer()
                       .AddSwaggerGen(
                           x =>
                           {
                               x.DescribeAllEnumsAsStrings();
                               x.IgnoreObsoleteActions();
                               x.IgnoreObsoleteProperties();
                               x.IncludeXmlComments(Path.Combine(ApplicationEnvironment.ApplicationBasePath, $"{nameof(CompilerAPI)}.xml"), true);
                               x.OrderActionsBy(y => y.RelativePath);

                               var descriptions =
                                   services.BuildServiceProvider()
                                           .GetRequiredService<IApiVersionDescriptionProvider>()
                                           .ApiVersionDescriptions;

                               foreach (var description in descriptions)
                               {
                                   x.SwaggerDoc(
                                       description.ApiVersion.ToString(),
                                       new Info
                                       {
                                           Title = "Reports API",
                                           Version = description.ApiVersion.ToString(),
                                           Contact = new Contact { Name = "Austin Drenski", Email = "austin.drenski@gmail.com" },
                                           License = new License { Name = "MIT", Url = "https://github.com/austindrenski/AD.OpenXml/blob/master/LICENSE" },
                                           Description = "The Reports API provides document combination and normalization."
                                       });
                               }
                           })
                       .AddMvc(
                           x =>
                           {
                               x.FormatterMappings.SetMediaTypeMappingForFormat("xml", "application/xml");
                               x.FormatterMappings.SetMediaTypeMappingForFormat("html", "text/html");
                               x.RespectBrowserAcceptHeader = true;
                           })
                       .SetCompatibilityVersion(CompatibilityVersion.Version_2_2)
                       .AddJsonOptions(
                           x =>
                           {
                               x.SerializerSettings.Converters.Add(new StringEnumConverter());
                               x.SerializerSettings.NullValueHandling = NullValueHandling.Ignore;
                               x.SerializerSettings.ReferenceLoopHandling = ReferenceLoopHandling.Serialize;
                               x.SerializerSettings.PreserveReferencesHandling = PreserveReferencesHandling.None;
                           })
                       .Services
                       .BuildServiceProvider();

        /// <summary>
        /// This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="app"/></exception>
        public void Configure([NotNull] IApplicationBuilder app)
            => app.UseServerHeader()
                  .UseHeadMethod()
                  .UseResponseCompression()
                  .UseStaticFiles()
                  .UseSwagger(x => x.RouteTemplate = "docs/{documentName}/swagger.json")
                  .UseSwaggerUI(
                      x =>
                      {
                          x.DefaultModelRendering(ModelRendering.Model);
                          x.DisplayRequestDuration();
                          x.DocExpansion(DocExpansion.None);
                          x.DocumentTitle = "Reports API Documentation";
                          x.RoutePrefix = "docs";

                          var descriptions =
                              app.ApplicationServices
                                 .GetRequiredService<IApiVersionDescriptionProvider>()
                                 .ApiVersionDescriptions
                                 .OrderByDescending(y => y.ApiVersion);

                          foreach (var description in descriptions)
                          {
                              x.SwaggerEndpoint($"{description.ApiVersion}/swagger.json", description.ApiVersion.ToString());
                          }
                      })
                  .UseMvc();
    }
}