using System.IO;
using JetBrains.Annotations;
using Microsoft.AspNetCore;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;

namespace CompilerAPI
{
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public static class Program
    {
        /// <summary>
        ///
        /// </summary>
        /// <param name="args"></param>
        public static void Main([NotNull] [ItemNotNull] string[] args)
        {
            BuildWebHost(args).Run();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public static IWebHost BuildWebHost([NotNull] [ItemNotNull] string[] args)
        {
            IConfigurationRoot configuration =
                new ConfigurationBuilder()
                    .AddCommandLine(args)
                    .Build();

            return
                WebHost.CreateDefaultBuilder(args)
                       .UseHttpSys()
                       .UseConfiguration(configuration)
                       .UseContentRoot(Directory.GetCurrentDirectory())
                       .UseStartup<Startup>()
                       .Build();
        }
    }
}