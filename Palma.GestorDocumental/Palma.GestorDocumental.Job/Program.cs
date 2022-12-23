using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Palma.GestorDocumental.Repository.Component.Implementation;
using System;
using System.IO;

namespace Palma.GestorDocumental.Job
{
    public class Program
    {
        static void Main(string[] args)
        {
            var builder = new ConfigurationBuilder();
            ConfigurationBuilder(builder);

            var host = Host.CreateDefaultBuilder()
                .ConfigureServices((context, services) =>
                {
                    Palma.GestorDocumental.Repository.Main.AddMigrationWordLink(services);

                })
                .Build();

            var flow = ActivatorUtilities.CreateInstance<Component>(host.Services);
            flow.Run();
            Console.WriteLine("Proceso Completado!!!");
            Console.ReadKey();
            
        }

        static void ConfigurationBuilder(IConfigurationBuilder builder)
        {
            builder.SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .AddEnvironmentVariables();
        }
    }
}
