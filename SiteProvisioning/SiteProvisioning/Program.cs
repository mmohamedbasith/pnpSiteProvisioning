using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Security;
using System.Threading;


namespace SiteProvisioning
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                ConsoleColor defaultForeground = Console.ForegroundColor;

                // Collect information 
                // string templateWebUrl = GetInput("Enter the URL of the template site: ", false, defaultForeground);
                string targetWebUrl = GetInput("Enter the URL of the target site: ", false, defaultForeground);
                string userName = GetInput("Enter your user name:", false, defaultForeground);
                string pwdS = GetInput("Enter your password:", true, defaultForeground);
                string filepath = GetInput("Get XMl Path:", false, defaultForeground);
                string filename = GetInput("Get XMl filename:", false, defaultForeground);
                Console.Write("\nPress 1 for Default authentication . Press 2 for Multi-Factor Authentication ");
                string auth = Console.ReadLine();
                SecureString pwd = new SecureString();

                foreach (char c in pwdS.ToCharArray()) pwd.AppendChar(c);
                if (auth.Equals("2"))
                {
                    var authenticationManager = new OfficeDevPnP.Core.AuthenticationManager();
                    ClientContext pnpcontext = authenticationManager.GetWebLoginClientContext(targetWebUrl, null);
                    Web pnpweb = pnpcontext.Web;
                    pnpcontext.Load(pnpweb);
                    pnpcontext.ExecuteQuery();
                    XMLTemplateProvider provider =
                                 new XMLFileSystemTemplateProvider(filepath, "");


                    ProvisioningTemplate template = provider.GetTemplate(filename);
                    ProvisioningTemplateApplyingInformation ptai =
                    new ProvisioningTemplateApplyingInformation();
                    ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                    {
                        Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                    };



                    template.Connector = provider.Connector;

                    pnpweb.ApplyProvisioningTemplate(template, ptai);
                }
                else
                {

                    using (var context = new ClientContext(targetWebUrl))
                    {
                        context.Credentials = new SharePointOnlineCredentials(userName, pwd);
                        Web web = context.Web;
                        context.Load(web, w => w.Title);
                        context.ExecuteQueryRetry();
                        XMLTemplateProvider provider =
                                 new XMLFileSystemTemplateProvider(filepath, "");
                        ProvisioningTemplate template = provider.GetTemplate(filename);
                        ProvisioningTemplateApplyingInformation ptai =
                        new ProvisioningTemplateApplyingInformation();
                        ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                        {
                            Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                        };
                        template.Connector = provider.Connector;
                        web.ApplyProvisioningTemplate(template, ptai);
                    }
                    // Get the template from existing site and serialize that (not really needed)
                    // Just to pause and indicate that it's all done
                }
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("We are all done. Press enter to continue.");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }

        
     
        private static string GetInput(string label, bool isPassword, ConsoleColor defaultForeground)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("{0} : ", label);
            Console.ForegroundColor = defaultForeground;

            string value = "";

            for (ConsoleKeyInfo keyInfo = Console.ReadKey(true); keyInfo.Key != ConsoleKey.Enter; keyInfo = Console.ReadKey(true))
            {
                if (keyInfo.Key == ConsoleKey.Backspace)
                {
                    if (value.Length > 0)
                    {
                        value = value.Remove(value.Length - 1);
                        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                        Console.Write(" ");
                        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                    }
                }
                else if (keyInfo.Key != ConsoleKey.Enter)
                {
                    if (isPassword)
                    {
                        Console.Write("*");
                    }
                    else
                    {
                        Console.Write(keyInfo.KeyChar);
                    }
                    value += keyInfo.KeyChar;

                }

            }
            Console.WriteLine("");

            return value;
        }
    }
}
