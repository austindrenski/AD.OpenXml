using System.IO.Compression;
using JetBrains.Annotations;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.ResponseCompression;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

namespace CompilerAPI
{
    [PublicAPI]
    public class Startup
    {
        public IConfigurationRoot Configuration { get; }

        private bool IsDevelopment { get; set; }

        public Startup(IHostingEnvironment env)
        {
            IsDevelopment = env.IsDevelopment();

            IConfigurationBuilder builder =
                new ConfigurationBuilder()
                    .SetBasePath(env.ContentRootPath)
                    .AddJsonFile($"Properties/appsettings.{env.EnvironmentName}.json", true, true)
                    .AddEnvironmentVariables();

            Configuration = builder.Build();
        }

        /// <summary>
        /// This method gets called by the runtime. Use this method to add services to the container.
        /// </summary>
        public void ConfigureServices(IServiceCollection services)
        {
            // Add framework services. 
            services.AddApplicationInsightsTelemetry(Configuration)
                    .AddApiVersioning(
                        x =>
                        {
                            x.AssumeDefaultVersionWhenUnspecified = true;
                            x.DefaultApiVersion = new ApiVersion(1, 0);
                        })
                    .AddAuthorization(
                        x =>
                        {
                            x.DefaultPolicy = new AuthorizationPolicyBuilder().RequireRole("ITCNET\\ALL_ITC").Build();
                        })
                    .AddLogging()
                    .AddMemoryCache()
                    .AddResponseCompression(
                        x =>
                        {
                            x.Providers.Add<GzipCompressionProvider>();
                        })
                    .AddRouting(
                        x =>
                        {
                            x.LowercaseUrls = true;
                        })
                    .AddMvc();

            services.Configure<GzipCompressionProviderOptions>(x => x.Level = CompressionLevel.Fastest);
        }

        /// <summary>
        /// This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        /// </summary>
        public void Configure(IApplicationBuilder app, IHostingEnvironment env, ILoggerFactory loggerFactory, IMemoryCache memoryCache)
        {
            loggerFactory.AddConsole(LogLevel.Information, false);

            app.Use(
                   async (context, next) =>
                   {
                       context.Response.Headers.Add("X-Content-Type-Options", "nosniff");
                       context.Response.Headers.Add("X-Frame-Options", "DENY");
                       context.Response.Headers.Add("X-Xss-Protection", "1; mode=block");
                       context.Response.Headers.Add("Referrer-Policy", "no-referrer");
                       context.Response.Headers.Add("X-Permitted-Cross-Domain-Policies", "none");
                       context.Response.Headers.Remove("X-Powered-By");
                       await next();
                   })
               .UseStaticFiles()
               .UseResponseCompression()
               .UseWhen(
                   x => env.IsDevelopment(),
                   x => x.UseDeveloperExceptionPage())
               .UseWhen(
                   x => env.IsProduction(),
                   x => x.UseExceptionHandler(
                       y =>
                       {
                           y.Run(
                               async context =>
                               {
                                   context.Response.StatusCode = 500;
                                   context.Response.ContentType = "text/html";
                                   await
                                       context.Response
                                              .WriteAsync("An internal server error has occured. Contact Austin.Drenski@usitc.gov.");
                               });
                       }))
               .UseMvc(
                   x =>
                   {
                       x.MapRoute("default", "api/reports/{controller}/{action}");
                   });
        }
    }
}