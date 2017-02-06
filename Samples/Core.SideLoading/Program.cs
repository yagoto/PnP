//*********************************************************
// THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY
// IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
// PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
//*********************************************************

using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Core.SideLoading
{

    class Program
    {
        static void Main(string[] args)
        {
            // Unique ID for side loading feature
            Guid sideloadingFeature = new Guid("AE3A1339-61F5-4f8f-81A7-ABD2DA956A7D");
            // Prompt for URL
            string url = GetUserInput("Please provide URL for the site where app is being installed: \n");
            // Prompt for Credentials 
            Console.WriteLine("Enter Credentials for {0}", url);
            string userName = GetUserInput("SharePoint username: ");
            SecureString pwd = GetPassword();

            // Get path to the location of the app file in file system
            string path = GetUserInput("Please provide full path to your app package: \n");



            url = "https://sbtwf2.sharepoint.com/sites/OpSite/Top/";
            userName = "admin@sbtwf2.onmicrosoft.com";
            path = @"C:\Users\yagoto\Source\Repos\WfLab\SpAddin1\SpAddin1\bin\Debug\app.publish\1.0.6.6\SpAddin1.app";



            // Create context for SharePoint online
            ClientContext ctx = new ClientContext(url);
            ctx.AuthenticationMode = ClientAuthenticationMode.Default;
            ctx.Credentials = new SharePointOnlineCredentials(userName, pwd);

            // Get variables for the operations
            Site site = ctx.Site;
            Web web = ctx.Web;

            try
            {
                // Make sure we have side loading enabled. 
                // Using PnP Nuget package extensions.
                site.ActivateFeature(sideloadingFeature);
                try
                {
                    bool isInstall = false;

                    if (isInstall)
                    {
                        // Load .app file and install that to site
                        var appstream = System.IO.File.OpenRead(path);
                        AppInstance app = web.LoadAndInstallApp(appstream);
                        ctx.Load(app);
                        ctx.ExecuteQuery();
                        Console.WriteLine("installed success.");
                    }
                    else
                    {
                        var appstream = System.IO.File.OpenRead(path);

                        var productId = new Guid("6C31CD20-2ECA-468E-A4E0-4E7A994F8BC6");

                        //var apps = web.GetAppInstances();

                        var apps = web.GetAppInstancesByProductId(productId);
                        ctx.Load(apps);
                        ctx.ExecuteQuery();

                        if (apps.Any())
                        {
                            AppInstance app = apps.First();
                            app.Upgrade(appstream);

                            ctx.Load(app);
                            ctx.ExecuteQuery();
                            Console.WriteLine("upgrade success. status = " + app.Status.ToString());

                        }
                    }
                }
                catch
                {
                    throw;
                }
                // Disable side loading feature using 
                // PnP Nuget package extensions. 
                site.DeactivateFeature(sideloadingFeature);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(string.Format("Exception!"), ex.ToString());
                Console.WriteLine("Press any key to continue.");
                Console.Read();
            }
        }


        /// <summary>
        /// Helper to get User Input from the console
        /// </summary>
        /// <returns></returns>
        public static string GetUserInput(string message)
        {
            // TODO:後で消すこと。
            return "";


            string path = string.Empty;
            Console.Write(message);
            path = Console.ReadLine();
           
            return path;
        }
        /// <summary>
        /// Helper to return the password
        /// </summary>
        /// <returns>SecureString representing the password</returns>
        public static SecureString GetPassword()
        {
            SecureString sStrPwd = new SecureString();

            try
            {
                Console.Write("SharePoint Password: ");

                for (ConsoleKeyInfo keyInfo = Console.ReadKey(true); keyInfo.Key != ConsoleKey.Enter; keyInfo = Console.ReadKey(true))
                {
                    if (keyInfo.Key == ConsoleKey.Backspace)
                    {
                        if (sStrPwd.Length > 0)
                        {
                            sStrPwd.RemoveAt(sStrPwd.Length - 1);
                            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                            Console.Write(" ");
                            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                        }
                    }
                    else if (keyInfo.Key != ConsoleKey.Enter)
                    {
                        Console.Write("*");
                        sStrPwd.AppendChar(keyInfo.KeyChar);
                    }

                }
                Console.WriteLine("");
            }
            catch (Exception e)
            {
                sStrPwd = null;
                Console.WriteLine(e.Message);
            }

            return sStrPwd;
        }


    }
}
