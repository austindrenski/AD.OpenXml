using System.IO;
using JetBrains.Annotations;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Net.Http.Server;

namespace CompilerAPI
{
    [PublicAPI]
    public static class Program
    {
        public static void Main(string[] args)
        {
            string environment = args.Length > 0 ? args[0] : "Production";

            IConfigurationRoot configuration =
                new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile($"./{nameof(CompilerAPI)}/Properties/hosting.{environment}.json", true, true)
                    .Build();

            IWebHost host =
                new WebHostBuilder()
                    .UseConfiguration(configuration)
                    .UseContentRoot(Directory.GetCurrentDirectory())
                    .UseStartup<Startup>()
                    .UseWebListener(
                        options =>
                        {
                            options.ListenerSettings.Authentication.Schemes = AuthenticationSchemes.NTLM;
                            options.ListenerSettings.Authentication.AllowAnonymous = true;
                        })
                    .UseApplicationInsights()
                    .Build();

            host.Run();
        }
    }
}