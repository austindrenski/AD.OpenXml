using System.IO;
using JetBrains.Annotations;
using Microsoft.AspNetCore.Hosting;
using AD.ApiExtensions;

namespace CompilerAPI
{
    /// <summary>
    /// The entry point for the <see cref="CompilerAPI"/> application.
    /// </summary>
    [PublicAPI]
    public static class Program
    {
        /// <summary>
        /// The main entry point called by the CLR upon startup.
        /// </summary>
        /// <param name="args">The command line arguments.</param>
        public static void Main([NotNull] [ItemNotNull] string[] args) => CreateWebHostBuilder(args).Build().Run();

        /// <summary>
        /// Creates a <see cref="IWebHostBuilder"/> for the current application.
        /// </summary>
        /// <param name="args">The command line arguments.</param>
        /// <returns>
        /// An instance of <see cref="IWebHostBuilder"/>.
        /// </returns>
        [Pure]
        [NotNull]
        public static IWebHostBuilder CreateWebHostBuilder([NotNull] string[] args)
            => new WebHostBuilder()
               .UseContentRoot(Directory.GetCurrentDirectory())
               .UseStartup<Startup>(args)
               .UseHttpSys();
    }
}