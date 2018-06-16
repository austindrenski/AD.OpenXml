using System.IO;
using JetBrains.Annotations;
using Microsoft.AspNetCore.Hosting;
using AD.ApiExtensions.Hosting;

namespace EditorAPI
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
        public static void Main([NotNull] [ItemNotNull] string[] args) => CreateWebHostBuilder(args).Build().Run();

        /// <summary>
        ///
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public static IWebHostBuilder CreateWebHostBuilder([NotNull] string[] args)
            => new WebHostBuilder()
               .UseContentRoot(Directory.GetCurrentDirectory())
               .UseStartup<Startup>(args)
               .UseHttpSys();
    }
}