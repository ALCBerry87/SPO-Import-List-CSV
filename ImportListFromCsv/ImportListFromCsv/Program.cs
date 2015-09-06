using System;
using System.IO;
using System.Configuration;
using System.Collections.Generic;
using System.Reflection;
using System.Diagnostics;
using System.Linq;
using System.Text;
using CsvHelper;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using Microsoft.SharePoint.Client.Taxonomy;

namespace ImportListFromCSV
{
    class Program
    {
        const string csvPath = "C:\filePath.csv";
        static void Main()
        {
            try
            {
                //Get site URL and credentials values from config
                Uri siteUri = new Uri(ConfigurationManager.AppSettings["SiteUrl"].ToString());
                var accountName = ConfigurationManager.AppSettings["AccountName"];
                char[] pwdChars = ConfigurationManager.AppSettings["AccountPwd"].ToCharArray();

                //Convert password to secure string
                System.Security.SecureString accountPwd = new System.Security.SecureString();
                for (int i = 0; i < pwdChars.Length; i++)
                {
                    accountPwd.AppendChar(pwdChars[i]);
                }

                //Connect to SharePoint Online
                using (var clientContext = new ClientContext(siteUri.ToString())
                {
                    Credentials = new SharePointOnlineCredentials(accountName, accountPwd)
                })
                {
                    if (clientContext != null)
                    {
                        //Map records from CSV file to C# list
                        List<CsvRecord> records = GetRecordsFromCsv();
                        //Get config-specified list
                        List spList = clientContext.Web.Lists.GetByTitle(ConfigurationManager.AppSettings["ListName"]);

                        foreach (CsvRecord record in records)
                        {
                            //Check for existing record based on title (assumes Title should be unique per record)
                            CamlQuery query = new CamlQuery();
                            query.ViewXml = String.Format("@<View><Query><Where><Eq><FieldRef Name=\"Title\" />" +
                                "<Value Type=\"Text\">{0}</Value></Eq></Where></Query></View>", record.Title);
                            var existingMappings = spList.GetItems(query);
                            clientContext.Load(existingMappings);
                            clientContext.ExecuteQuery();

                            switch (existingMappings.Count)
                            {
                                case 0:
                                    //No records found, needs to be added
                                    AddNewListItem(record, spList, clientContext);
                                    break;
                                default:
                                    //An existing record was found - continue with next item
                                    continue;
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError("Failed: " + ex.Message);
                Trace.TraceError("Stack Trace: " + ex.StackTrace);
            }
        }

        private static void AddNewListItem(CsvRecord record, List spList, ClientContext clientContext)
        {
            //Instantiate dictionary to temporarily store field values
            Dictionary<string, object> itemFieldValues = new Dictionary<string, object>();
            //Use reflection to iterate through the record's properties
            PropertyInfo[] properties = typeof(CsvRecord).GetProperties();
            foreach (PropertyInfo property in properties)
            {
                //Get property value
                object propValue = property.GetValue(record, null);
                //Only set field if the property has a value
                if (!String.IsNullOrEmpty(propValue.ToString()))
                {
                    //Get site column that matches the property name
                    //ASSUMPTION: Your property names match the internal names of the corresponding site columns
                    Field matchingField = spList.Fields.GetByInternalNameOrTitle(property.Name);
                    clientContext.Load(matchingField);
                    clientContext.ExecuteQuery();

                    //Switch on the field type
                    switch (matchingField.FieldTypeKind)
                    {
                        case FieldType.User:
                            FieldUserValue userFieldValue = GetUserFieldValue(propValue.ToString(), clientContext);
                            if (userFieldValue != null)
                                itemFieldValues.Add(matchingField.InternalName, userFieldValue);
                            else
                                throw new Exception("User field value could not be added: " + propValue.ToString());
                            break;
                        case FieldType.Lookup:
                            FieldLookupValue lookupFieldValue = GetLookupFieldValue(propValue.ToString(),
                                ConfigurationManager.AppSettings["LookupListName"].ToString(),
                                clientContext);
                            if (lookupFieldValue != null)
                                itemFieldValues.Add(matchingField.InternalName, lookupFieldValue);
                            else
                                throw new Exception("Lookup field value could not be added: " + propValue.ToString());
                            break;
                        case FieldType.Invalid:
                            switch (matchingField.TypeAsString)
                            {
                                case "TaxonomyFieldType":
                                    TaxonomyFieldValue taxFieldValue = GetTaxonomyFieldValue(propValue.ToString(), matchingField, clientContext);
                                    if (taxFieldValue != null)
                                        itemFieldValues.Add(matchingField.InternalName, taxFieldValue);
                                    else
                                        throw new Exception("Taxonomy field value could not be added: " + propValue.ToString());
                                    break;
                                default:
                                    //Code for publishing site columns not implemented
                                    continue;
                            }
                            break;
                        default:
                            itemFieldValues.Add(matchingField.InternalName, propValue);
                            break;
                    }
                }
            }

            //Add new item to list
            ListItemCreationInformation creationInfo = new ListItemCreationInformation();
            ListItem oListItem = spList.AddItem(creationInfo);

            foreach (KeyValuePair<string, object> itemFieldValue in itemFieldValues)
            {
                //Set each field value
                oListItem[itemFieldValue.Key] = itemFieldValue.Value;
            }
            //Persist changes
            oListItem.Update();
            clientContext.ExecuteQuery();
        }

        private static List<CsvRecord> GetRecordsFromCsv()
        {
            List<CsvRecord> records = new List<CsvRecord>();
            using (var sr = new StreamReader(csvPath))
            {
                var reader = new CsvReader(sr);
                records = reader.GetRecords<CsvRecord>().ToList();
            }

            return records;
        }

        private static FieldUserValue GetUserFieldValue(string userName, ClientContext clientContext)
        {
            //Returns first principal match based on user identifier (display name, email, etc.)
            ClientResult<PrincipalInfo> principalInfo = Utility.ResolvePrincipal(
                clientContext, //context
                clientContext.Web, //web
                userName, //input
                PrincipalType.User, //scopes
                PrincipalSource.All, //sources
                null, //usersContainer
                false); //inputIsEmailOnly
            clientContext.ExecuteQuery();
            PrincipalInfo person = principalInfo.Value;

            if (person != null)
            {
                //Get User field from login name
                User validatedUser = clientContext.Web.EnsureUser(person.LoginName);
                clientContext.Load(validatedUser);
                clientContext.ExecuteQuery();

                if (validatedUser != null && validatedUser.Id > 0)
                {
                    //Sets lookup ID for user field to the appropriate user ID
                    FieldUserValue userFieldValue = new FieldUserValue();
                    userFieldValue.LookupId = validatedUser.Id;
                    return userFieldValue;
                }
            }
            return null;
        }

        public static FieldLookupValue GetLookupFieldValue(string lookupName, string lookupListName, ClientContext clientContext)
        {
            //Ref: Karine Bosch - https://karinebosch.wordpress.com/2015/05/11/setting-the-value-of-a-lookup-field-using-csom/
            var lookupList = clientContext.Web.Lists.GetByTitle(lookupListName);
            CamlQuery query = new CamlQuery();
            string lookupFieldName = ConfigurationManager.AppSettings["LookupFieldName"].ToString();
            string lookupFieldType = ConfigurationManager.AppSettings["LookupFieldType"].ToString();

            query.ViewXml = string.Format(@"<View><Query><Where><Eq><FieldRef Name='{0}'/><Value Type='{1}'>{2}</Value></Eq>" +
                                            "</Where></Query></View>", lookupFieldName, lookupFieldType, lookupName);

            ListItemCollection listItems = lookupList.GetItems(query);
            clientContext.Load(listItems, items => items.Include
                                                (listItem => listItem["ID"],
                                                listItem => listItem[lookupFieldName]));
            clientContext.ExecuteQuery();

            if (listItems != null)
            {
                ListItem item = listItems[0];
                FieldLookupValue lookupValue = new FieldLookupValue();
                lookupValue.LookupId = int.Parse(item["ID"].ToString());
                return lookupValue;
            }

            return null;
        }

        public static TaxonomyFieldValue GetTaxonomyFieldValue(string termName, Field mmField, ClientContext clientContext)
        {
            //Ref: Steve Curran - http://sharepointfieldnotes.blogspot.com/2013_06_01_archive.html
            //Cast field to TaxonomyField to get its TermSetId
            TaxonomyField taxField = clientContext.CastTo<TaxonomyField>(mmField);
            //Get term ID from name and term set ID
            string termId = GetTermIdForTerm(termName, taxField.TermSetId, clientContext);
            if (!string.IsNullOrEmpty(termId))
            {
                //Set TaxonomyFieldValue
                TaxonomyFieldValue termValue = new TaxonomyFieldValue();
                termValue.Label = termName;
                termValue.TermGuid = termId;
                termValue.WssId = -1;
                return termValue;
            }
            return null;
        }

        public static string GetTermIdForTerm(string term, Guid termSetId, ClientContext clientContext)
        {
            //Ref: Steve Curran - http://sharepointfieldnotes.blogspot.com/2013_06_01_archive.html
            string termId = string.Empty;

            //Get term set from ID
            TaxonomySession tSession = TaxonomySession.GetTaxonomySession(clientContext);
            TermStore ts = tSession.GetDefaultSiteCollectionTermStore();
            TermSet tset = ts.GetTermSet(termSetId);

            LabelMatchInformation lmi = new LabelMatchInformation(clientContext);

            lmi.Lcid = 1033;
            lmi.TrimUnavailable = true;
            lmi.TermLabel = term;

            //Search for matching terms in the term set based on label
            TermCollection termMatches = tset.GetTerms(lmi);
            clientContext.Load(tSession);
            clientContext.Load(ts);
            clientContext.Load(tset);
            clientContext.Load(termMatches);

            clientContext.ExecuteQuery();

            //Set term ID to first match
            if (termMatches != null && termMatches.Count() > 0)
                termId = termMatches.First().Id.ToString();

            return termId;
        }
    }
}

