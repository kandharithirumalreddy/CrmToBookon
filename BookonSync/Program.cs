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
            string webSPOUrl = "https://bookon.dkbs.dk";
            //string webSPOUrl = "http://rnd06:55001";
                        
            try
            {
                Console.WriteLine("Connecting to site...");
                using (ClientContext context = new ClientContext(webSPOUrl))
                {
                    //context.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;
                    //Web web = context.Web;


                    context.AuthenticationMode = ClientAuthenticationMode.FormsAuthentication;
                    context.FormsAuthenticationLoginInfo = new FormsAuthenticationLoginInfo("CRM Automation", "9LEkTny4");
                    context.ExecuteQuery();
                    Console.WriteLine("Successfully connected.");
                    Web web = context.Web;

                    string contentTypeName = args[0];
                    string actionType = args[1];
                    List<FieldMataData> itemMetaData = JsonConvert.DeserializeObject<List<FieldMataData>>(args[2]);
                    if (contentTypeName == "Partner")
                    {
                        List lst = web.Lists.GetByTitle("Partnere");
                        ContentTypeCollection ctColl = lst.ContentTypes;
                        context.Load(ctColl);
                        context.ExecuteQuery();
                        foreach (ContentType ct in ctColl)
                        {
                            if (ct.Name == contentTypeName)
                            {
                                string zipCodeId = null;
                                if(itemMetaData.Find(x => x.FieldName == "postNumber") != null)
                                {
                                    zipCodeId = getZipCodeId(context, itemMetaData.Find(x => x.FieldName == "postNumber").Value) + ";#";
                                }

                                if (actionType == "Create")
                                {
                                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                                    ListItem newItem = lst.AddItem(itemCreateInfo);
                                    newItem["ContentTypeId"] = ct.Id;
                                    foreach (FieldMataData field in itemMetaData)
                                    {
                                        switch (field.FieldName)
                                        {
                                            case "accountId":
                                                newItem["accountID"] = field.Value;
                                                break;
                                            case "partnertype":
                                                newItem["PartnerType"] = getPartnerTypeID(field.Value) + ";#";
                                                break;
                                            //case "membershipType":
                                            //    break;
                                            case "partnerName":
                                                newItem["CompanyName"] = field.Value;
                                                break;
                                            case "cvr":
                                                newItem["VatNumber"] = field.Value;
                                                break;
                                            case "telefon":
                                                newItem["Phone"] = field.Value + "#@#";
                                                break;
                                            case "centertype":
                                                newItem["CenterType"] = getCenterTypeFormatedValue(field.Value);
                                                break;
                                            case "address1":
                                                newItem["Address1"] = field.Value;
                                                break;
                                            case "address2":
                                                newItem["Address2"] = field.Value;
                                                break;
                                            //case "town":
                                            case "postNumber":
                                                newItem["ZipMachingFilter"] = zipCodeId;
                                                break;
                                            case "land":
                                                newItem["Country"] = getLandId(field.Value);
                                                break;
                                            //case "stateAgreement":
                                            //    break;
                                            case "debitornummer":
                                                newItem["DebtorNumber"] = field.Value;
                                                break;
                                            case "debitornummer2":
                                                newItem["DebtorNumber2"] = field.Value;
                                                break;
                                            case "regions":
                                                //region values have to be sepaprated by => ;#
                                                newItem["Region"] = field.Value;
                                                break;
                                            case "membershipStartDate":
                                                //have to be provided in UTC format string
                                                newItem["MembershipStartDate"] = field.Value;
                                                break;
                                            case "publicURL":
                                                newItem["PublicURL"] = field.Value;
                                                break;
                                            case "email":
                                                newItem["EmailAddress"] = field.Value;
                                                break;
                                            case "website":
                                                newItem["Website"] = field.Value + ", " + field.Value;
                                                break;
                                            case "panoramView":
                                                newItem["PanoramaView"] = field.Value + ", " + field.Value;
                                                break;
                                            case "recommandedNPGRT60":
                                                newItem["Recommended"] = field.Value;
                                                break;
                                            case "qualityAssuredNPSGRD30":
                                                newItem["Quality"] = field.Value;
                                                break;
                                            default:
                                                break;
                                        }
                                    }
                                    newItem.Update();
                                    context.ExecuteQuery();
                                }
                                else if (actionType == "Update")
                                {
                                    if (itemMetaData.Find(x => x.FieldName == "accountId") != null)
                                    {
                                        string accountId = itemMetaData.Find(x => x.FieldName == "accountId").Value;
                                        ListItem updatableItem = getPartnerItem(context, lst, accountId);

                                        if (updatableItem != null)
                                        {
                                            foreach (FieldMataData field in itemMetaData)
                                            {
                                                switch (field.FieldName)
                                                {
                                                    case "partnertype":
                                                        updatableItem["PartnerType"] = getPartnerTypeID(field.Value) + ";#";
                                                        break;
                                                    //case "membershipType":
                                                    //    break;
                                                    case "partnerName":
                                                        updatableItem["CompanyName"] = field.Value;
                                                        break;
                                                    case "cvr":
                                                        updatableItem["VatNumber"] = field.Value;
                                                        break;
                                                    case "telefon":
                                                        updatableItem["Phone"] = field.Value + "#@#";
                                                        break;
                                                    case "centertype":
                                                        updatableItem["CenterType"] = getCenterTypeFormatedValue(field.Value);
                                                        break;
                                                    case "address1":
                                                        updatableItem["Address1"] = field.Value;
                                                        break;
                                                    case "address2":
                                                        updatableItem["Address2"] = field.Value;
                                                        break;
                                                    //case "town":
                                                    case "postNumber":
                                                        updatableItem["ZipMachingFilter"] = zipCodeId;
                                                        break;
                                                    case "land":
                                                        updatableItem["Country"] = getLandId(field.Value);
                                                        break;
                                                    //case "stateAgreement":
                                                    //    break;
                                                    case "debitornummer":
                                                        updatableItem["DebtorNumber"] = field.Value;
                                                        break;
                                                    case "debitornummer2":
                                                        updatableItem["DebtorNumber2"] = field.Value;
                                                        break;
                                                    case "regions":
                                                        //region values have to be sepaprated by => ;#
                                                        updatableItem["Region"] = field.Value;
                                                        break;
                                                    case "membershipStartDate":
                                                        //have to be provided in UTC format string
                                                        updatableItem["MembershipStartDate"] = field.Value;
                                                        break;
                                                    case "publicURL":
                                                        updatableItem["PublicURL"] = field.Value;
                                                        break;
                                                    case "email":
                                                        updatableItem["EmailAddress"] = field.Value;
                                                        break;
                                                    case "website":
                                                        updatableItem["Website"] = field.Value + ", " + field.Value;
                                                        break;
                                                    case "panoramView":
                                                        updatableItem["PanoramaView"] = field.Value + ", " + field.Value;
                                                        break;
                                                    case "recommandedNPGRT60":
                                                        updatableItem["Recommended"] = field.Value;
                                                        break;
                                                    case "qualityAssuredNPSGRD30":
                                                        updatableItem["Quality"] = field.Value;
                                                        break;
                                                    default:
                                                        break;
                                                }
                                            }
                                            updatableItem.Update();
                                            context.ExecuteQuery();
                                        }
                                    }
                                }
                                break;
                            }
                        }
                    }
                    else
                    {
                        List lst = web.Lists.GetByTitle("Customers");
                        ContentTypeCollection ctColl = lst.ContentTypes;
                        context.Load(ctColl);
                        context.ExecuteQuery();

                        foreach (ContentType ct in ctColl)
                        {
                            if (ct.Name == contentTypeName && contentTypeName == "Organisation")
                            {
                                string zipCodeItemId = null;
                                string industryCodeItemId = null;

                                if (itemMetaData.Find(x => x.FieldName == "postNumber") != null)
                                {
                                    string zipCode = itemMetaData.Find(x => x.FieldName == "postNumber").Value;                                    
                                    zipCodeItemId = getZipCodeId(context, zipCode);                                   
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
                                        newItem["Country"] = getLandId(country);
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
                                else if (actionType == "Update")
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
                                        if (organization != null)
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

                                    if (updatableItem != null)
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

        private static ListItem getPartnerItem(ClientContext context, List customersLst, string accountId)
        {
            ListItem result = null;
            CamlQuery query = new CamlQuery();
            query.ViewXml = @"<View><Query><Where><Eq><FieldRef Name='accountID' /><Value Type='Text'>" + accountId + @"</Value></Eq></Where></Query>
                                                        <ViewFields><FieldRef Name='ID'/><FieldRef Name='accountID'/><FieldRef Name='PartnerType'/><FieldRef Name='CompanyName'/><FieldRef Name='VatNumber'/><FieldRef Name='Phone'/><FieldRef Name='CenterType'/><FieldRef Name='Address1'/><FieldRef Name='Address2'/><FieldRef Name='ZipMachingFilter'/><FieldRef Name='Country'/><FieldRef Name='DebtorNumber'/><FieldRef Name='DebtorNumber2'/><FieldRef Name='Region'/><FieldRef Name='MembershipStartDate'/><FieldRef Name='PublicURL'/><FieldRef Name='EmailAddress'/><FieldRef Name='Website'/><FieldRef Name='PanoramaView'/><FieldRef Name='Recommended'/><FieldRef Name='Quality'/></ViewFields></View>";
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
                                                        <ViewFields><FieldRef Name='ID'/><FieldRef Name='accountID'/><FieldRef Name='PartnerType'/><FieldRef Name='CompanyName'/><FieldRef Name='VatNumber'/><FieldRef Name='Phone'/><FieldRef Name='CenterType'/><FieldRef Name='Address1'/><FieldRef Name='Address2'/><FieldRef Name='ZipMachingFilter'/><FieldRef Name='Country'/><FieldRef Name='DebtorNumber'/><FieldRef Name='DebtorNumber2'/><FieldRef Name='Region'/><FieldRef Name='MembershipStartDate'/><FieldRef Name='PublicURL'/><FieldRef Name='EmailAddress'/><FieldRef Name='Website'/><FieldRef Name='PanoramaView'/><FieldRef Name='Recommended'/><FieldRef Name='Quality'/></ViewFields></View>";
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

        private static string getPartnerTypeID(string partnereTypeTitle)
        {
            string result = null;
            switch (partnereTypeTitle)
            {
                case "Gold":
                    result = "1";
                    break;
                case "Ikke-partner":
                    result = "2";
                    break;
                case "Samarbejdspartner":
                    result = "3";
                    break;
                case "Bronze":
                    result = "4";
                    break;
                case "Silver":
                    result = "5";
                    break;
                case "Deaktiverede partnere":
                    result = "6";
                    break;
                case "Preferred partner":
                    result = "7";
                    break;
                default:
                    break;
            }
            return result;
        }
        private static string getCenterTypeFormatedValue(string centerTypes)
        {
            string result = null;
            string[] centerTitles = centerTypes.Split(new string[] { ";#" }, StringSplitOptions.RemoveEmptyEntries);
            for(int i = 0; i < centerTitles.Length; i++)
            {
                if(centerTitles[i] == "Slotte & Herregårde")
                {
                    result = result + "8;#" + centerTitles[i] + ";#";
                }
                if (centerTitles[i] == "Ud i naturen")
                {
                    result = result + "9;#" + centerTitles[i] + ";#";
                }
                if (centerTitles[i] == "Teambuilding & aktiviteter")
                {
                    result = result + "10;#" + centerTitles[i] + ";#";
                }
                if (centerTitles[i] == "Det kulinariske arrangement")
                {
                    result = result + "11;#" + centerTitles[i] + ";#";
                }
                if (centerTitles[i] == "Strand & vand")
                {
                    result = result + "12;#" + centerTitles[i] + ";#";
                }
                if (centerTitles[i] == "Det skæve arrangement")
                {
                    result = result + "13;#" + centerTitles[i] + ";#";
                }
                if (centerTitles[i] == "By & seværdigheder")
                {
                    result = result + "14;#" + centerTitles[i] + ";#";
                }
                if (centerTitles[i] == "Firmafester")
                {
                    result = result + "15;#" + centerTitles[i] + ";#";
                }
            }

            return result;
        }

        private static string getZipCodeId(ClientContext context, string zipCode)
        {
            string result = null;
            List zipsList = context.Web.Lists.GetByTitle("TownZipCodes");
            CamlQuery query = new CamlQuery();            
            query.ViewXml = @"<View><Query><Where><Eq><FieldRef Name='WorkZip' /><Value Type='Text'>" + zipCode + @"</Value></Eq></Where></Query>
                                                        <ViewFields><FieldRef Name='ID'/></ViewFields></View>";
            ListItemCollection zipColl = zipsList.GetItems(query);
            context.Load(zipColl);
            context.ExecuteQuery();
            if (zipColl.Count == 1)
            {
                result = zipColl[0].Id.ToString() + ";#";
            }
            return result;
        }

        private static string getLandId(string landTitle)
        {
            string result = null;
            switch (landTitle)
            {
                case "Denmark":
                    result = "1;#";
                    break;
                case "Germany":
                    result = "2;#";
                    break;
                case "Sweden":
                    result = "3;#";
                    break;
                case "Andora":
                    result = "4;#";
                    break;
                default:
                    break;
            }
            return result;
        }
    }

    public class FieldMataData
    {
        public string FieldName { get; set; }
        public string Value { get; set; }
    }

    
}
