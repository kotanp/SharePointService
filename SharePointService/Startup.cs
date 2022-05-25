using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SharePointService.Models;

namespace SharePointService
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddControllersWithViews();
            var settings = Configuration.GetSection("Settings");
            services.Configure<Settings>(settings);
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseExceptionHandler("/Home/Error");
            }
            app.UseStaticFiles();

            app.UseRouting();

            app.UseAuthorization();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllerRoute(
                    name: "sharepointupload",
                    pattern: "sharepoint/upload",
                    defaults: new { controller = "Upload", action = "Index" });
                endpoints.MapControllerRoute(
                    name: "sharepointdelete",
                    pattern: "sharepoint/delete",
                    defaults: new { controller = "Upload", action = "Delete" });
                endpoints.MapControllerRoute(
                    name: "sharepointconvert",
                    pattern: "sharepoint/convert",
                    defaults: new { controller = "Upload", action = "ConvertToPdf" });
                // endpoints.MapControllerRoute(
                //     name: "pdfconvert",
                //     pattern: "sharepoint/convert",
                //     defaults: new { controller = "PdfConverter", action = "Index" });
            });
        }
    }
}
