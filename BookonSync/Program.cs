using System;
using System.Collections.Generic;
using System.Security;
using System.Reflection;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;

namespace BookonSync
{
    class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("SharePoint Online site URL:");
            //string webSPOUrl = "https://bookon.dkbs.dk";
            string webSPOUrl = "http://rnd06:55001";

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
                    //context.ExecutingWebRequest += clientContext_ExecutingWebRequest;
                    //context.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;
                    //Web web = context.Web;
                    //context.Load(web.Lists);
                    //context.Load(web);
                    //context.ExecuteQuery();



                    context.AuthenticationMode = ClientAuthenticationMode.FormsAuthentication;
                    context.FormsAuthenticationLoginInfo =
                    new FormsAuthenticationLoginInfo("CRM Automation", "9LEkTny4");
                    context.ExecuteQuery();
                    Console.WriteLine("Successfully connected.");
                    Web web = context.Web;
                    Console.WriteLine("Getting " + listName + " list...");
                    List lst = web.Lists.GetByTitle(listName);
                    ContentTypeCollection ctColl = lst.ContentTypes;
                    context.Load(ctColl);
                    context.ExecuteQuery();
                    Console.WriteLine("List loaded.");

                    string contentTypeName = args[0];
                    string actionType = args[1]; ;
                    List<FieldMataData> itemMetaData = JsonConvert.DeserializeObject<List<FieldMataData>>(args[2]);

                    foreach (ContentType ct in ctColl)
                    {
                        if (ct.Name == contentTypeName && contentTypeName == "Organisation")
                        {
                            string zipCodeItemId = null;
                            string industryCodeItemId = null;

                            if (itemMetaData.Find(x => x.FieldName == "postNumber") != null)
                            {
                                //filter by zipcode item ID
                                List zipsList = web.Lists.GetByTitle("TownZipCodes");
                                CamlQuery query = new CamlQuery();
                                string zipCode = itemMetaData.Find(x => x.FieldName == "postNumber").Value;
                                query.ViewXml = @"<View><Query><Where><Eq><FieldRef Name='WorkZip' /><Value Type='Text'>" + zipCode + @"</Value></Eq></Where></Query>
                                                        <ViewFields><FieldRef Name='ID'/></ViewFields></View>";
                                ListItemCollection zipColl = zipsList.GetItems(query);
                                context.Load(zipColl);
                                context.ExecuteQuery();
                                if (zipColl.Count == 1)
                                {
                                    zipCodeItemId = zipColl[0].Id.ToString() + ";#";
                                }
                            }

                            if (itemMetaData.Find(x => x.FieldName == "industryCode") != null)
                            {
                                string industryTitle = itemMetaData.Find(x => x.FieldName == "industryCode").Value;
                                List industryLst = web.Lists.GetByTitle("IndustryCode");
                                CamlQuery query = new CamlQuery();
                                query.ViewXml = @"<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + industryTitle + @"</Value></Eq></Where></Query>
                                                        <ViewFields><FieldRef Name='ID'/></ViewFields></View>";
                                ListItemCollection industryColl = industryLst.GetItems(query);
                                context.Load(industryColl);
                                context.ExecuteQuery();
                                if (industryColl.Count == 1)
                                {
                                    industryCodeItemId = industryColl[0].Id.ToString() + ";#";
                                }
                            }

                            if (actionType == "Create")
                            {
                                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                                ListItem newItem = lst.AddItem(itemCreateInfo);
                                newItem["ContentTypeId"] = ct.Id;

                                if (itemMetaData.Find(x => x.FieldName == "companyName") != null)
                                {
                                    newItem["Title"] = itemMetaData.Find(x => x.FieldName == "companyName").Value;
                                }

                                if (itemMetaData.Find(x => x.FieldName == "address1") != null)
                                {
                                    newItem["Address"] = itemMetaData.Find(x => x.FieldName == "address1").Value;
                                }

                                if (itemMetaData.Find(x => x.FieldName == "address2") != null)
                                {
                                    newItem["Address2"] = itemMetaData.Find(x => x.FieldName == "address2").Value;
                                }                                
                                
                                if (itemMetaData.Find(x => x.FieldName == "postNumber") != null)
                                {
                                    newItem["ZipMachingFilter"] = zipCodeItemId;
                                }

                                if (itemMetaData.Find(x => x.FieldName == "country") != null)
                                {
                                    string country = itemMetaData.Find(x => x.FieldName == "country").Value;
                                    switch (country)
                                    {
                                        case "Denmark":
                                            newItem["Country"] = "1;#";
                                            break;
                                        case "Germany":
                                            newItem["Country"] = "2;#";
                                            break;
                                        case "Sweden":
                                            newItem["Country"] = "3;#";
                                            break;
                                        case "Andora":
                                            newItem["Country"] = "4;#";
                                            break;
                                        default:
                                            break;
                                    }
                                }

                                if (itemMetaData.Find(x => x.FieldName == "phoneNumber") != null)
                                {
                                    newItem["Phone"] = itemMetaData.Find(x => x.FieldName == "phoneNumber").Value;
                                }

                                if (itemMetaData.Find(x => x.FieldName == "accountId") != null)
                                {
                                    newItem["accountID"] = itemMetaData.Find(x => x.FieldName == "accountId").Value;
                                }

                                if (itemMetaData.Find(x => x.FieldName == "industryCode") != null)
                                {
                                    newItem["IndustryCode"] = industryCodeItemId;
                                }
                                newItem.Update();
                                context.ExecuteQuery();
                            }
                            else if(actionType == "Update")
                            {
                                if (itemMetaData.Find(x => x.FieldName == "accountId") != null)
                                {
                                    string accountId = itemMetaData.Find(x => x.FieldName == "accountId").Value;
                                    ListItem updatableItem = getCustomerItem(context, lst, accountId, "accountID");

                                    if (updatableItem != null)
                                    {
                                        if (itemMetaData.Find(x => x.FieldName == "companyName") != null)
                                        {
                                            updatableItem["Title"] = itemMetaData.Find(x => x.FieldName == "companyName").Value;
                                        }

                                        if (itemMetaData.Find(x => x.FieldName == "address1") != null)
                                        {
                                            updatableItem["Address"] = itemMetaData.Find(x => x.FieldName == "address1").Value;
                                        }

                                        if (itemMetaData.Find(x => x.FieldName == "address2") != null)
                                        {
                                            updatableItem["Address2"] = itemMetaData.Find(x => x.FieldName == "address2").Value;
                                        }                                
                                
                                        if (itemMetaData.Find(x => x.FieldName == "postNumber") != null)
                                        {
                                            updatableItem["ZipMachingFilter"] = zipCodeItemId;
                                        }

                                        if (itemMetaData.Find(x => x.FieldName == "country") != null)
                                        {
                                            string country = itemMetaData.Find(x => x.FieldName == "country").Value;
                                            switch (country)
                                            {
                                                case "Denmark":
                                                    updatableItem["Country"] = "1;#";
                                                    break;
                                                case "Germany":
                                                    updatableItem["Country"] = "2;#";
                                                    break;
                                                case "Sweden":
                                                    updatableItem["Country"] = "3;#";
                                                    break;
                                                case "Andora":
                                                    updatableItem["Country"] = "4;#";
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }

                                        if (itemMetaData.Find(x => x.FieldName == "phoneNumber") != null)
                                        {
                                            updatableItem["Phone"] = itemMetaData.Find(x => x.FieldName == "phoneNumber").Value;
                                        }
                                        
                                        if (itemMetaData.Find(x => x.FieldName == "industryCode") != null)
                                        {
                                            updatableItem["IndustryCode"] = industryCodeItemId;
                                        }
                                        updatableItem.Update();
                                        context.ExecuteQuery();
                                    }
                                }
                            }
                            break;
                        }

                        if (ct.Name == contentTypeName && contentTypeName == "Kontaktperson")
                        {
                            if (actionType == "Create")
                            {
                                string relatedOrgId = null;
                                if (itemMetaData.Find(x => x.FieldName == "accountId") != null)
                                {
                                    string accountId = itemMetaData.Find(x => x.FieldName == "accountId").Value;                                    
                                    ListItem organization = getCustomerItem(context, lst, accountId, "accountID");
                                    if(organization != null)
                                    {
                                        relatedOrgId = organization.Id.ToString() + ";#";
                                    }                                    
                                }


                                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                                ListItem newItem = lst.AddItem(itemCreateInfo);
                                newItem["ContentTypeId"] = ct.Id;

                                if (itemMetaData.Find(x => x.FieldName == "firstName") != null || itemMetaData.Find(x => x.FieldName == "lastName") != null)
                                {
                                    string firstName = itemMetaData.Find(x => x.FieldName == "firstName") != null ? itemMetaData.Find(x => x.FieldName == "firstName").Value : "";
                                    string lastName = itemMetaData.Find(x => x.FieldName == "lastName") != null ? itemMetaData.Find(x => x.FieldName == "lastName").Value : "";
                                    string joinedName = firstName + " " + lastName;
                                    newItem["Title"] = joinedName;
                                }

                                if (itemMetaData.Find(x => x.FieldName == "email") != null)
                                {
                                    newItem["Email"] = itemMetaData.Find(x => x.FieldName == "email").Value;
                                }

                                if (itemMetaData.Find(x => x.FieldName == "telephone") != null)
                                {
                                    newItem["Phone"] = itemMetaData.Find(x => x.FieldName == "telephone").Value;
                                }

                                if (itemMetaData.Find(x => x.FieldName == "mobilePhone") != null)
                                {
                                    newItem["CellPhone"] = itemMetaData.Find(x => x.FieldName == "mobilePhone").Value;
                                }

                                if (itemMetaData.Find(x => x.FieldName == "accountId") != null)
                                {
                                    string accountId = itemMetaData.Find(x => x.FieldName == "accountId").Value;
                                    newItem["accountID"] = accountId;
                                    newItem["RelatedOrganization"] = relatedOrgId;
                                }
                                if (itemMetaData.Find(x => x.FieldName == "contactId") != null)
                                {
                                    newItem["contactID"] = itemMetaData.Find(x => x.FieldName == "contactId").Value;
                                }
                                newItem.Update();
                                context.ExecuteQuery();
                            }
                            else if (actionType == "Update")
                            {
                                string relatedOrgId = null;
                                
                                if (itemMetaData.Find(x => x.FieldName == "accountId") != null)
                                {
                                    string accountId = itemMetaData.Find(x => x.FieldName == "accountId").Value;
                                    ListItem organization = getCustomerItem(context, lst, accountId, "accountID");
                                    if (organization != null)
                                    {
                                        relatedOrgId = organization.Id.ToString() + ";#";
                                    }
                                }

                                ListItem updatableItem = null;
                                if (itemMetaData.Find(x => x.FieldName == "contactId") != null)
                                {
                                    string contactId = itemMetaData.Find(x => x.FieldName == "contactId").Value;
                                    updatableItem = getCustomerItem(context, lst, contactId, "contactID");                                    
                                }

                                if(updatableItem != null)
                                {
                                    if (itemMetaData.Find(x => x.FieldName == "firstName") != null || itemMetaData.Find(x => x.FieldName == "lastName") != null)
                                    {
                                        string firstName = itemMetaData.Find(x => x.FieldName == "firstName") != null ? itemMetaData.Find(x => x.FieldName == "firstName").Value : "";
                                        string lastName = itemMetaData.Find(x => x.FieldName == "lastName") != null ? itemMetaData.Find(x => x.FieldName == "lastName").Value : "";
                                        string joinedName = firstName + " " + lastName;
                                        updatableItem["Title"] = joinedName;
                                    }

                                    if (itemMetaData.Find(x => x.FieldName == "email") != null)
                                    {
                                        updatableItem["Email"] = itemMetaData.Find(x => x.FieldName == "email").Value;
                                    }

                                    if (itemMetaData.Find(x => x.FieldName == "telephone") != null)
                                    {
                                        updatableItem["Phone"] = itemMetaData.Find(x => x.FieldName == "telephone").Value;
                                    }

                                    if (itemMetaData.Find(x => x.FieldName == "mobilePhone") != null)
                                    {
                                        updatableItem["CellPhone"] = itemMetaData.Find(x => x.FieldName == "mobilePhone").Value;
                                    }

                                    if (itemMetaData.Find(x => x.FieldName == "accountId") != null)
                                    {
                                        string accountId = itemMetaData.Find(x => x.FieldName == "accountId").Value;
                                        updatableItem["accountID"] = accountId;
                                        updatableItem["RelatedOrganization"] = relatedOrgId;
                                    }
                                    updatableItem.Update();
                                    context.ExecuteQuery();
                                }
                            }

                            break;
                        }
                    }

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
        private static ListItem getCustomerItem(ClientContext context, List customersLst, string accountId, string searchableField)
        {
            ListItem result = null;
            CamlQuery query = new CamlQuery();
            query.ViewXml = @"<View><Query><Where><Eq><FieldRef Name='"+searchableField+@"' /><Value Type='Text'>" + accountId + @"</Value></Eq></Where></Query>
                                                        <ViewFields><FieldRef Name='ID'/><FieldRef Name='Title'/><FieldRef Name='Address'/><FieldRef Name='Address2'/><FieldRef Name='ZipMachingFilter'/><FieldRef Name='Country'/><FieldRef Name='Phone'/><FieldRef Name='IndustryCode'/></ViewFields></View>";
            ListItemCollection customerColl = customersLst.GetItems(query);
            context.Load(customerColl);
            context.ExecuteQuery();
            if (customerColl.Count == 1)
            {
                result = customerColl[0];
            }
            else
            {
                CamlQuery query2 = new CamlQuery();
                query2.ViewXml = @"<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Text'>" + accountId + @"</Value></Eq></Where></Query>
                                                        <ViewFields><FieldRef Name='Title'/><FieldRef Name='Address'/><FieldRef Name='Address2'/><FieldRef Name='ZipMachingFilter'/><FieldRef Name='Country'/><FieldRef Name='Phone'/><FieldRef Name='IndustryCode'/></ViewFields></View>";
                ListItemCollection customerColl2 = customersLst.GetItems(query2);
                context.Load(customerColl2);
                context.ExecuteQuery();
                if (customerColl2.Count == 1)
                {
                    result = customerColl2[0];
                }
            }

            return result;
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
        public string FieldName { get; set; }
        public string Value { get; set; }
    }

    
}
