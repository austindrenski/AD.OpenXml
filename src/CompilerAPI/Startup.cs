using System;
using System.IO;
using System.IO.Compression;
using AD.ApiExtensions.Conventions;
using AD.ApiExtensions.Filters;
using AD.ApiExtensions.Primitives;
using JetBrains.Annotations;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.ResponseCompression;
using Microsoft.DotNet.PlatformAbstractions;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Swashbuckle.AspNetCore.Swagger;

namespace CompilerAPI
{
    // TODO: document Startup.
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public class Startup : IStartup
    {
        /// <summary>
        ///
        /// </summary>
        private IConfiguration Configuration { get; }

        /// <summary>
        ///
        /// </summary>
        /// <param name="configuration"></param>
        /// <exception cref="ArgumentNullException"></exception>
        public Startup([NotNull] IConfiguration configuration) => Configuration = configuration;

        /// <summary>
        /// This method gets called by the runtime. Use this method to add services to the container.
        /// </summary>
        public IServiceProvider ConfigureServices([NotNull] IServiceCollection services)
            => services.AddLogging(x => x.AddConsole())
                       .AddApiVersioning(
                           x =>
                           {
                               x.AssumeDefaultVersionWhenUnspecified = true;
                               x.DefaultApiVersion = new ApiVersion(1, 0);
                           })
                       .AddResponseCompression(x => x.Providers.Add<GzipCompressionProvider>())
                       .Configure<GzipCompressionProviderOptions>(x => x.Level = CompressionLevel.Fastest)
                       .AddRouting(x => x.LowercaseUrls = true)
                       .AddAntiforgery(
                           x =>
                           {
                               x.HeaderName = "x-xsrf-token";
                               x.Cookie.HttpOnly = true;
                               x.Cookie.SecurePolicy = CookieSecurePolicy.SameAsRequest;
                           })
                       .Configure<CookiePolicyOptions>(
                           x =>
                           {
                               x.CheckConsentNeeded = context => true;
                               x.MinimumSameSitePolicy = SameSiteMode.None;
                           })
                       .AddSwaggerGen(
                           x =>
                           {
                               x.DescribeAllEnumsAsStrings();
                               x.MapType<GroupingValues<string, string>>(() => new Schema { Type = "string" });
                               x.IncludeXmlComments(Path.Combine(ApplicationEnvironment.ApplicationBasePath, $"{nameof(CompilerAPI)}.xml"));
                               x.IgnoreObsoleteActions();
                               x.IgnoreObsoleteProperties();
                               x.SwaggerDoc("v1", new Info { Title = "Reports API", Version = "v1" });
                               x.OperationFilter<SwaggerOptionalOperationFilter>();
                           })
                       .AddMvc(
                           x =>
                           {
                               x.Conventions.Add(new KebabControllerModelConvention());
                               x.FormatterMappings.SetMediaTypeMappingForFormat("xml", "application/xml");
                               x.FormatterMappings.SetMediaTypeMappingForFormat("html", "text/html");
                               x.ModelMetadataDetailsProviders.Add(new KebabBindingMetadataProvider());
                               x.RespectBrowserAcceptHeader = true;
                           })
                       .SetCompatibilityVersion(CompatibilityVersion.Version_2_1)
                       .AddJsonOptions(
                           x =>
                           {
                               x.SerializerSettings.ReferenceLoopHandling = ReferenceLoopHandling.Serialize;
                               x.SerializerSettings.PreserveReferencesHandling = PreserveReferencesHandling.Objects;
                               x.SerializerSettings.ContractResolver = new KebabContractResolver();
                           })
                       .Services
                       .BuildServiceProvider();

        /// <summary>
        /// This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        /// </summary>
        public void Configure([NotNull] IApplicationBuilder app)
            => app.Use(
                      async (context, next) =>
                      {
                          context.Response.Headers.Add("referrer-policy", "no-referrer");
                          context.Response.Headers.Add("x-content-type-options", "nosniff");
                          context.Response.Headers.Add("x-frame-options", "deny");
                          context.Response.Headers.Add("x-xss-protection", "1; mode=block");
                          await next();
                      })
                  .UseStaticFiles()
                  .UseSwagger(x => x.RouteTemplate = "docs/{documentName}/swagger.json")
                  .UseSwaggerUI(
                      x =>
                      {
                          x.RoutePrefix = "docs";
                          x.DocumentTitle = "Reports API Documentation";
                          x.HeadContent = "Reports API Documentation";
                          x.SwaggerEndpoint("v1/swagger.json", "Reports API Documentation");
                          x.DocExpansion(DocExpansion.None);
                      })
                  .UseResponseCompression()
                  .UseMvc();
    }
}