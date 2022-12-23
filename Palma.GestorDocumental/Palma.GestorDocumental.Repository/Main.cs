using System;
using Microsoft.Extensions.DependencyInjection;
using Palma.GestorDocumental.Repository.Component.Interface;
using Palma.GestorDocumental.Repository;


namespace Palma.GestorDocumental.Repository
{
    public class Main
    {
        public static void AddMigrationWordLink(IServiceCollection services)
        {
            services.AddTransient<IComponent, Component.Implementation.Component>();
            services.AddMemoryCache();

        }
    }
}
