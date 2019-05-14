using System;
using System.Collections.Generic;
using System.Security;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;

namespace BookonSync
{
    class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("SharePoint Online site URL:");
            string webSPOUrl = "https://bookon.dkbs.dk";
            //string webSPOUrl = "http://rnd06:55001";

            //Console.WriteLine("User Name:");
            //string userName = Console.ReadLine();

            //Console.WriteLine("Password:");
            //SecureString password = FetchPasswordFromConsole();

            //Console.WriteLine("List or Document library Title:");
            string listName = "Customers";

            //Console.WriteLine("Start from ID:");
            //int itemIDStart = Convert.ToInt32(Console.ReadLine());

            //Console.WriteLine("End on ID:");
            //int itemIDEnd = Convert.ToInt32(Console.ReadLine());


            try
            {

                //ClientContext context = new ClientContext(webSPOUrl);
                //Web web = context.Web;
                //context.ExecutingWebRequest += clientContext_ExecutingWebRequest;
                //context.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;
                //context.ExecuteQuery();

                Console.WriteLine("Connecting to site...");
                using (ClientContext context = new ClientContext(webSPOUrl))
                {
                    ////context.ExecutingWebRequest += clientContext_ExecutingWebRequest;
                    //context.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;
                    //Web web = context.Web;
                    ////context.Load(web.Lists);
                    ////context.Load(web);
                    ////context.ExecuteQuery();
                    ///

                    
                    context.AuthenticationMode = ClientAuthenticationMode.FormsAuthentication;
                    context.FormsAuthenticationLoginInfo =
                    new FormsAuthenticationLoginInfo("CRM Automation", "9LEkTny4");
                    context.ExecuteQuery();
                    Console.WriteLine("Successfully connected.");
                    Web web = context.Web;
                    Console.WriteLine("Getting "+ listName + " list...");
                    List lst = web.Lists.GetByTitle(listName);
                    ContentTypeCollection ctColl = lst.ContentTypes;
                    context.Load(ctColl);
                    context.ExecuteQuery();
                    Console.WriteLine("List loaded.");
                    
                    string contentTypeName = args[0];
                    List<FieldMataData> itemMetaData = JsonConvert.DeserializeObject<List<FieldMataData>>(args[1]);

                    //foreach (ContentType ct in ctColl)
                    //{
                    //    if (ct.Name == contentTypeName)
                    //    {
                    //        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    //        ListItem newItem = lst.AddItem(itemCreateInfo);
                    //        newItem["ContentTypeId"] = ct.Id;
                    //        foreach (FieldMataData fm in itemMetaData)
                    //        {
                    //            newItem[fm.fieldName] = fm.value;
                    //        }
                    //        newItem.Update();
                    //        context.ExecuteQuery();
                    //        break;
                    //    }
                    //}
                    
                }
                Console.WriteLine("End");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error is: " + ex.Message);
                Console.WriteLine("Try to run script again.");
                Console.ReadLine();
            }
        }
        static void clientContext_ExecutingWebRequest(object sender, WebRequestEventArgs e)
        {
            try
            {
                e.WebRequestExecutor.WebRequest.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
            }
            catch
            { throw; }

        }
        private static SecureString FetchPasswordFromConsole()
        {
            string password = "";
            ConsoleKeyInfo info = Console.ReadKey(true);
            while (info.Key != ConsoleKey.Enter)
            {
                if (info.Key != ConsoleKey.Backspace)
                {
                    Console.Write("*");
                    password += info.KeyChar;
                }
                else if (info.Key == ConsoleKey.Backspace)
                {
                    if (!string.IsNullOrEmpty(password))
                    {
                        password = password.Substring(0, password.Length - 1);
                        int pos = Console.CursorLeft;
                        Console.SetCursorPosition(pos - 1, Console.CursorTop);
                        Console.Write(" ");
                        Console.SetCursorPosition(pos - 1, Console.CursorTop);
                    }
                }
                info = Console.ReadKey(true);
            }
            Console.WriteLine();
            var securePassword = new SecureString();
            //Convert string to secure string  
            foreach (char c in password)
                securePassword.AppendChar(c);
            securePassword.MakeReadOnly();
            return securePassword;
        }
    }

    public class FieldMataData
    {
        public string fieldName { get; set; }
        public string value { get; set; }
    }
}
