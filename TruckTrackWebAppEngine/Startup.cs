using Microsoft.Owin;
using Microsoft.Owin.FileSystems;
using Microsoft.Owin.StaticFiles;
using Owin;
using System;
using System.Diagnostics;
using System.IO;
using System.Web.Http;
using System.Web.Http.Cors;

namespace TruckTrackWebAppEngine
{
    public class Startup
    {
        // This code configures Web API. The Startup class is specified as a type
        // parameter in the WebApp.Start method.
        public void Configuration(IAppBuilder app)
        {
            try
            {
                // configure webapi
                HttpConfiguration config = new HttpConfiguration();
                config.EnableCors(new EnableCorsAttribute("*", "*", "*"));
                config.MapHttpAttributeRoutes();
                config.Routes.MapHttpRoute(
                    name: "DefaultApi",
                    routeTemplate: "api/{controller}/{id}",
                    defaults: new { id = RouteParameter.Optional }
                );
                app.UseWebApi(config);

                // setup fileserver
                string fileSaveDirectory = AppCommon.GetFileSaveDirectory();
                Directory.CreateDirectory(fileSaveDirectory);
                if (!Directory.Exists(fileSaveDirectory))
                {
                    throw new Exception("Invalid FileSaveDirectory specified <" + fileSaveDirectory + ">.");
                }
                else
                {
                    FileServerOptions fsOptions = new FileServerOptions();
                    fsOptions.RequestPath = PathString.Empty;
                    fsOptions.FileSystem = new PhysicalFileSystem(fileSaveDirectory);
                    app.UseFileServer(fsOptions);
                }
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("Configuration: " + e.Message, e);
                throw new Exception(message);
            }

        }
    }
}

