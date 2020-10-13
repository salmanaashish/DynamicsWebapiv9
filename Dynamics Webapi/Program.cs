using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net;
using System.Xml;

namespace Dynamics_Webapi
{
    internal class Program
    {
        private const string V = "?$id=";

        //Get configuration data from App.config connectionStrings
        private static readonly string connectionString = ConfigurationManager.ConnectionStrings["Connect"].ConnectionString;

        private static readonly ServiceConfig config = new ServiceConfig(connectionString);

        #region Basic Operation
        public static Uri CreateMethod()
        {

            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {
                    #region  Basic Create operations

                    Console.WriteLine("--Section 1 started--");
                    //Create a contact
                    var contact1 = new JObject
                        {
                            { "firstname", "Rafel" },
                            { "lastname", "Shillo" }
                        };
                    Uri contact1Uri = svc.PostCreate("contacts", contact1);
                    Console.WriteLine($"Contact '{contact1["firstname"]} " +
                        $"{contact1["lastname"]}' created.");
                    Console.WriteLine($"Contact URI: {contact1Uri}");
                    return contact1Uri;
                    #endregion Basic Create operations

                }
            }
            catch (Exception)
            {
                throw;
            }

        }

        public void UpdateMethod(Uri contact1Uri)
        {

            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {
                    #region  Basic Update operations
                    var contact1 = new JObject
                        {
                            { "firstname", "Rafel" },
                            { "lastname", "Shillo" }
                        };
                    //Update a contact
                    JObject contact1Add = new JObject
                    {
                        { "annualincome", 80000 },
                        { "jobtitle", "Junior Developer" }
                    };
                    svc.Patch(contact1Uri, contact1Add);
                    Console.WriteLine(
                    $"Contact '{contact1["firstname"]} {contact1["lastname"]}' " +
                    $"updated with jobtitle and annual income");

                    #endregion Basic Update operations
                }
            }
            catch (Exception)
            {
                throw;
            }

        }

        public static void RetrieveRecord(Uri contact1Uri)
        {
            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {
                    #region retrievingarecord

                    //Retrieve a contact
                    var retrievedcontact1 = svc.Get(contact1Uri.ToString() +
                        "?$select=fullname,annualincome,jobtitle,description");
                    Console.WriteLine($"Contact '{retrievedcontact1["fullname"]}' retrieved: \n" +
                    $"\tAnnual income: {retrievedcontact1["annualincome"]}\n" +
                    $"\tJob title: {retrievedcontact1["jobtitle"]} \n" +
                    //description is initialized empty.
                    $"\tDescription: {retrievedcontact1["description"]}.");

                    #endregion retrievingarecord
                }
            }
            catch (Exception)
            {
                throw;
            }

        }
        public static void ChangeJustoneProperty(Uri contact1Uri)
        {
            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {
                    #region updatingonlyonefield 
                    //Retrieve a contact
                    var retrievedcontact1 = svc.Get(contact1Uri.ToString() +
                        "?$select=fullname,annualincome,jobtitle,description");
                    Console.WriteLine($"Contact '{retrievedcontact1["fullname"]}' retrieved: \n" +
                    $"\tAnnual income: {retrievedcontact1["annualincome"]}\n" +
                    $"\tJob title: {retrievedcontact1["jobtitle"]} \n" +
                    //description is initialized empty.
                    $"\tDescription: {retrievedcontact1["description"]}.");

                    #endregion updatingonlyonefield
                }
            }
            catch (Exception)
            {
                throw;
            }

        }

        public static void retrievejustthesingleproperty(Uri contact1Uri)
        {
            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {
                    #region retrievesinglerecord
                    //Retrieve a contact
                    var telephone1Value = svc.Get($"{contact1Uri}/telephone1");
                    Console.WriteLine($"Contact's telephone # is: {telephone1Value["value"]}.");

                    #endregion retrievesinglerecord
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void Createrecordassociatedtoanotherentity(Uri contact1Uri)
        {
            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {
                    #region Create record associated to another


                    //Create a new account and associate with existing contact in one operation.
                    var account1 = new JObject
                    {
                        { "name", "Contoso Ltd" },
                        { "telephone1", "555-5555" },
                        { "primarycontactid@odata.bind", contact1Uri }
                    };
                    var account1Uri = svc.PostCreate("accounts", account1);
                
                    Console.WriteLine($"Account '{account1["name"]}' created.");
                    Console.WriteLine($"Account URI: {account1Uri}");
                    //Retrieve account name and primary contact info
                    JObject retrievedAccount1 = svc.Get($"{account1Uri}?$select=name," +
                        $"&$expand=primarycontactid($select=fullname,jobtitle,annualincome)") as JObject;

                    Console.WriteLine($"Account '{retrievedAccount1["name"]}' has primary contact " +
                        $"'{retrievedAccount1["primarycontactid"]["fullname"]}':");
                    Console.WriteLine($"\tJob title: {retrievedAccount1["primarycontactid"]["jobtitle"]} \n" +
                        $"\tAnnual income: {retrievedAccount1["primarycontactid"]["annualincome"]}");

                    #endregion : Create record associated to another

                    #region  Create related entities

                    //Build the Account object inside-out, starting with most nested type(s)
                    JArray tasks = new JArray();
                    JObject task1 = new JObject
                    {
                        { "subject", "Sign invoice" },
                        { "description", "Invoice #12321" },
                        { "scheduledend", DateTimeOffset.Parse("4/19/2019") }
                    };
                    tasks.Add(task1);
                    JObject task2 = new JObject
                    {
                        { "subject", "Setup new display" },
                        { "description", "Theme is - Spring is in the air" },
                        { "scheduledstart", DateTimeOffset.Parse("4/20/2019") }
                    };
                    tasks.Add(task2);
                    JObject task3 = new JObject
                    {
                        { "subject", "Conduct training" },
                        { "description", "Train team on making our new blended coffee" },
                        { "scheduledstart", DateTimeOffset.Parse("6/1/2019") }
                    };
                    tasks.Add(task3);

                    JObject contact2 = new JObject
                    {
                        { "firstname", "Susie" },
                        { "lastname", "Curtis" },
                        { "jobtitle", "Coffee Master" },
                        { "annualincome", 48000 },
                        //Add related tasks using corresponding navigation property
                        { "Contact_Tasks", tasks }
                    };

                    JObject account2 = new JObject
                    {
                        { "name", "Fourth Coffee" },
                        //Add related contacts using corresponding navigation property
                        { "primarycontactid", contact2 }
                    };

                    //Create the account and related records
                    Uri account2Uri = svc.PostCreate("accounts", account2);
                    Console.WriteLine($"Account '{account2["name"]}  created.");
                 
                    Console.WriteLine($"Contact URI: {account2Uri}");

                    //Retrieve account, primary contact info, and assigned tasks for contact.
                    //CDS only supports querying-by-expansion one level deep, so first query
                    // account-primary contact.
                    var retrievedAccount2 = svc.Get($"{account2Uri}?$select=name," +
                        $"&$expand=primarycontactid($select=fullname,jobtitle,annualincome)");

                    Console.WriteLine($"Account '{retrievedAccount2["name"]}' " +
                        $"has primary contact '{retrievedAccount2["primarycontactid"]["fullname"]}':");

                    Console.WriteLine($"\tJob title: {retrievedAccount2["primarycontactid"]["jobtitle"]} \n" +
                        $"\tAnnual income: {retrievedAccount2["primarycontactid"]["annualincome"]}");

                    //Next retrieve same contact and its assigned tasks.
                    //Don't have a saved URI for contact 'Susie Curtis', so create one
                    // from base address and entity ID.
                    Uri contact2Uri = new Uri($"{svc.BaseAddress}contacts({retrievedAccount2["primarycontactid"]["contactid"]})");
                    //Retrieve the contact
                    var retrievedcontact2 = svc.Get($"{contact2Uri}?$select=fullname," +
                        $"&$expand=Contact_Tasks($select=subject,description,scheduledstart,scheduledend)");

                    Console.WriteLine($"Contact '{retrievedcontact2["fullname"]}' has the following assigned tasks:");
                    foreach (JToken tk in retrievedcontact2["Contact_Tasks"])
                    {
                        Console.WriteLine(
                            $"Subject: {tk["subject"]}, \n" +
                            $"\tDescription: {tk["description"]}\n" +
                            $"\tStart: {tk["scheduledstart"].Value<DateTime>().ToString("d")}\n" +
                            $"\tEnd: {tk["scheduledend"].Value<DateTime>().ToString("d")}\n");
                    }

                    #endregion   Create related entities


                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static void AssociateandDisassociateentities(Uri contact1Uri,Uri account2Uri)
        {
            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {
                    #region Associate and Disassociate entities

                    //Retrieve a contact
                 

                  
                    /// <summary>
                    /// Demonstrates associating and disassociating of existing entity instances.
                    /// </summary>
                    Console.WriteLine("\n--Section 4 started--");
                    //Add 'Rafel Shillo' to the contact list of 'Fourth Coffee',
                    // a 1-to-N relationship.
                    JObject rel1 = new JObject
                    {
                        { "@odata.id", contact1Uri }
                    }; //relationship object for msg content
                    Uri navUri1 = new Uri($"{account2Uri}/contact_customer_accounts/$ref");
                    //Create relationship
                    svc.Post(navUri1.ToString(), rel1);
                 
                    //Retrieve and output all contacts for account 'Fourth Coffee'.
                    var retrievedContactList1 = svc.Get($"{account2Uri}/contact_customer_accounts?" +
                        $"$select=fullname,jobtitle");

                

                    foreach (JToken ct in retrievedContactList1["value"])
                    {
                        Console.WriteLine($"\tName: {ct["fullname"]}, Job title: {ct["jobtitle"]}");
                    }

                    //Dissociate the contact from the account.  For a collection-valued
                    // navigation property, must append URI of referenced entity.
                    Uri dis1Uri = new Uri($"{navUri1}?$id={contact1Uri}");
                     
                    //Equivalently, could have dissociated from the other end of the
                    // relationship, using the single-valued navigation ref, located in
                    // the contact 'Peter Cambel'.  This dissociation URI has a simpler form:
                    // [Org URI]/api/data/v9.1/contacts([contactid#])/parentcustomerid_account/$ref

                    svc.Delete(dis1Uri.ToString());

                    //'Rafel Shillo' was removed from the the contact list of 'Fourth Coffee'

                    //Associate an opportunity to a competitor, an N-to-N relationship.
                    //First, create the required entity instances.
                    JObject comp1 = new JObject
                    {
                        { "name", "Adventure Works" },
                        {
                            "strengths",
                            "Strong promoter of private tours for multi-day outdoor adventures"
                        }
                    };
                    Uri comp1Uri = svc.PostCreate("competitors", comp1);
                    string comp1Uri1 = comp1Uri.ToString();
                    JObject oppor1 = new JObject
                    {
                        ["name"] = "River rafting adventure",
                        ["description"] = "Sales team on a river-rafting offsite and team building"
                    };
                    Uri oppor1Uri = svc.PostCreate("opportunities", oppor1);
                  
                    //Associate opportunity to competitor via opportunitycompetitors_association.
                    // navigation property.
                    JObject rel2 = new JObject
                    {
                        { "@odata.id", comp1Uri }
                    };
                    Uri navUri2 = new Uri($"{oppor1Uri}/opportunitycompetitors_association/$ref");
                   string navUri3 = navUri2.ToString();
                    svc.Post(navUri2.ToString(), rel2);
                    Console.WriteLine($"Opportunity '{oppor1["name"]}' associated with competitor '{comp1["name"]}'.");

                    //Retrieve all opportunities for competitor 'Adventure Works'.
                    var retrievedOpporList1 = svc.Get($"{comp1Uri}?$select=name,&$expand=opportunitycompetitors_association($select=name,description)");

                    Console.WriteLine($"Competitor '{retrievedOpporList1["name"]}' has the following opportunities:");
                    foreach (JToken op in
                        retrievedOpporList1["opportunitycompetitors_association"])
                    {
                        Console.WriteLine($"\tName: {op["name"]}, \n" +
                            $"\tDescription: {op["description"]}");
                    }

                    //Dissociate opportunity from competitor.
                    //svc.Delete(new Uri(navUri2.ToString() + V + comp1Uri.ToString()));

                    // 'River rafting adventure' opportunity disassociated with 'Adventure Works' competitor


                    #endregion  Associate and Disassociate entities
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static void DeleteRecords(Uri contact1Uri)
        {
            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {
                    svc.Delete(contact1Uri.ToString());
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        #endregion Basic Operation


        #region Querydata

        public static void retrieverecords(Uri account1Uri)
        {
            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {
                    //Get the id and name of the account created to use as a filter.
                    var contoso = svc.Get($"{account1Uri}?$select=accountid,name");
                    var contosoId = Guid.Parse(contoso["accountid"].ToString());
                    string contosoName = (string)contoso["name"];

                    // Basic query: Query using $select against a contact entity to get the properties you want.
                    // For performance best practice, always use $select, otherwise all properties are returned
                    Console.WriteLine("-- Basic Query --");


                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static void retrieverecordswithformattedvalues(Uri contact1Uri)
        {
            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {
                    //Header required to include formatted values
                    var formattedValueHeaders = new Dictionary<string, List<string>> {
                        { "Prefer", new List<string>
                            { "odata.include-annotations=\"OData.Community.Display.V1.FormattedValue\"" }
                        }
                    };

                    var contact1 = svc.Get(
                        $"{contact1Uri}?$select=fullname,jobtitle,annualincome",
                        formattedValueHeaders);

                    Console.WriteLine($"Contact basic info:\n" +
                        $"\tFullname: {contact1["fullname"]}\n" +
                        $"\tJobtitle: {contact1["jobtitle"]}\n" +
                        $"\tAnnualincome (unformatted): {contact1["annualincome"]} \n" +
                        $"\tAnnualincome (formatted): {contact1["annualincome@OData.Community.Display.V1.FormattedValue"]} \n");



                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static void retrieverecordswithformattedvalueswithfilters(Uri contact1Uri, Uri account1Uri)
        {
            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {

                    //Get the id and name of the account created to use as a filter.
                    var contoso = svc.Get($"{account1Uri}?$select=accountid,name");
                    var contosoId = Guid.Parse(contoso["accountid"].ToString());
                    string contosoName = (string)contoso["name"];

                    #region Selecting specific properties

                    // Basic query: Query using $select against a contact entity to get the properties you want.
                    // For performance best practice, always use $select, otherwise all properties are returned
                    Console.WriteLine("-- Basic Query --");

                    //Header required to include formatted values
                    var formattedValueHeaders = new Dictionary<string, List<string>> {
                        { "Prefer", new List<string>
                            { "odata.include-annotations=\"OData.Community.Display.V1.FormattedValue\"" }
                        }
                    };

                    var contact1 = svc.Get(
                        $"{contact1Uri}?$select=fullname,jobtitle,annualincome",
                        formattedValueHeaders);
                    // Filter criteria:
                    // Applying filters to get targeted data.
                    // 1) Using standard query functions (e.g.: contains, endswith, startswith)
                    // 2) Using CDS query functions (e.g.: LastXhours, Last7Days, Today, Between, In, ...)
                    // 3) Using filter operators and logical operators (e.g.: eq, ne, gt, and, or, etc…)
                    // 4) Set precedence using parenthesis (e.g.: ((criteria1) and (criteria2)) or (criteria3)
                    // For more info, see:
                    //https://docs.microsoft.com/en-us/powerapps/developer/common-data-service/webapi/query-data-web-api#filter-results

                    Console.WriteLine("-- Filter Criteria --");
                    //Filter 1: Using standard query functions to filter results.  In this operation, we
                    //will query for all contacts with fullname containing the string "(sample)".
                    JToken containsSampleinFullNameCollection = svc.Get("contacts?" +
                        "$select=fullname,jobtitle,annualincome&" +
                        "$filter=contains(fullname,'(sample)') and " +
                        $"_parentcustomerid_value eq {contosoId.ToString()}",
                        formattedValueHeaders);

                    WriteContactResultsTable(
                        "Contacts filtered by fullname containing '(sample)':",
                        containsSampleinFullNameCollection["value"]);

                    //Filter 2: Using CDS query functions to filter results. In this operation, we will query
                    //for all contacts that were created in the last hour. For complete list of CDS query
                    //functions, see: https://docs.microsoft.com/dynamics365/customer-engagement/web-api/queryfunctions

                    JToken createdInLastHourCollection = svc.Get("contacts?" +
                    "$select=fullname,jobtitle,annualincome&" +
                    "$filter=Microsoft.Dynamics.CRM.LastXHours(PropertyName='createdon',PropertyValue='1') and " +
                    $"_parentcustomerid_value eq {contosoId.ToString()}",
                    formattedValueHeaders);

                    WriteContactResultsTable(
                        "Contacts that were created within the last 1hr:",
                        createdInLastHourCollection["value"]);

                    //Filter 3: Using operators. Building on the previous operation, we further limit
                    //the results by the contact's income. For more info on standard filter operators,
                    //https://docs.microsoft.com/powerapps/developer/common-data-service/webapi/query-data-web-api#filter-results

                    JToken highIncomeContacts = svc.Get("contacts?" +
                        "$select=fullname,jobtitle,annualincome&" +
                        "$filter=contains(fullname,'(sample)') and " +
                        "annualincome gt 55000  and " +
                        $"_parentcustomerid_value eq {contosoId.ToString()}",
                        formattedValueHeaders);


                    WriteContactResultsTable(
                       "Contacts with '(sample)' in name and income above $55,000:",
                       highIncomeContacts["value"]);

                    //Filter 4: Set precedence using parentheses. Continue building on the previous
                    //operation, we further limit results by job title. Parentheses and the order of
                    //filter statements can impact results returned.

                    JToken seniorOrSpecialistsCollection = svc.Get("contacts?" +
                        "$select=fullname,jobtitle,annualincome&" +
                        "$filter=contains(fullname,'(sample)') and " +
                        "(contains(jobtitle, 'senior') or " +
                        "contains(jobtitle,'manager')) and " +
                        "annualincome gt 55000 and " +
                        $"_parentcustomerid_value eq {contosoId.ToString()}",
                        formattedValueHeaders);

                    WriteContactResultsTable(
                        "Contacts with '(sample)' in name senior jobtitle or high income:",
                        seniorOrSpecialistsCollection["value"]);

                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static void retrieverecordswithformattedvalueswithfiltersandordingandalias(dynamic contosoId, dynamic formattedValueHeaders)
        {
            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {


                    #region Ordering and aliases

                    //Results can be ordered in descending or ascending order.
                    Console.WriteLine("\n-- Order Results --");

                    JToken orderedResults = svc.Get("contacts?" +
                        "$select=fullname,jobtitle,annualincome&" +
                        "$filter=contains(fullname,'(sample)')and " +
                        $"_parentcustomerid_value eq {contosoId.ToString()}&" +
                        "$orderby=jobtitle asc, annualincome desc",
                        formattedValueHeaders);

                    WriteContactResultsTable(
                        "Contacts ordered by jobtitle (Ascending) and annualincome (descending)",
                        orderedResults["value"]);

                    //Parameterized aliases can be used as parameters in a query. These parameters can be used
                    //in $filter and $orderby options. Using the previous operation as basis, parameterizing the
                    //query will give us the same results. For more info, see:
                    //https://docs.microsoft.com/powerapps/developer/common-data-service/webapi/use-web-api-functions#passing-parameters-to-a-function

                    Console.WriteLine("\n-- Parameterized Aliases --");

                    JToken orderedResultsWithParams = svc.Get("contacts?" +
                        "$select=fullname,jobtitle,annualincome&" +
                        "$filter=contains(@p1,'(sample)') and " +
                        "@p2 eq @p3&" +
                        "$orderby=@p4 asc, @p5 desc&" +
                        "@p1=fullname&" +
                        "@p2=_parentcustomerid_value&" +
                        $"@p3={contosoId.ToString()}&" +
                        "@p4=jobtitle&" +
                        "@p5=annualincome",
                        formattedValueHeaders);

                    WriteContactResultsTable(
                        "Contacts ordered by jobtitle (Ascending) and annualincome (descending)",
                        orderedResultsWithParams["value"]);

                    #endregion Ordering and aliases

                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static void retrieverecordswithLimitresults(dynamic contosoId, dynamic formattedValueHeaders)
        {
            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {



                    #region Limit results

                    //To limit records returned, use the $top query option.  Specifying a limit number for $top
                    //returns at most that number of results per request. Extra results are ignored.
                    //For more information, see:
                    // https://docs.microsoft.com/powerapps/developer/common-data-service/webapi/query-data-web-api#use-top-query-option
                    Console.WriteLine("\n-- Top Results --");

                    JToken topFive = svc.Get("contacts?" +
                        "$select=fullname,jobtitle,annualincome&" +
                        "$filter=contains(fullname,'(sample)') and " +
                        $"_parentcustomerid_value eq {contosoId.ToString()}&" +
                        "$top=5",
                        formattedValueHeaders);

                    WriteContactResultsTable("Contacts top 5 results:", topFive["value"]);

                    //Result count - count the number of results matching the filter criteria.
                    //Tip: Use count together with the "odata.maxpagesize" to calculate the number of pages in
                    //the query.  Note: CDS has a max record limit of 5000 records per response.
                    Console.WriteLine("\n-- Result Count --");
                    //1) Get a count of a collection without the data.
                    JToken count = svc.Get($"contacts/$count");
                    Console.WriteLine($"\nThe contacts collection has {count} contacts.");
                    //  2) Get a count along with the data.

                    JToken countWithData = svc.Get("contacts?" +
                        "$select=fullname,jobtitle,annualincome&" +
                        "$filter=(contains(jobtitle,'senior') or contains(jobtitle, 'manager')) and " +
                        $"_parentcustomerid_value eq {contosoId.ToString()}" +
                        "&$count=true",
                        formattedValueHeaders);

                    WriteContactResultsTable($"{countWithData["@odata.count"]} " +
                        $"Contacts with 'senior' or 'manager' in job title:",
                        countWithData["value"]);

                    #endregion Limit results

                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static void retrieverecordswithExpandingresults(dynamic contosoId, dynamic formattedValueHeaders, dynamic account1Uri, dynamic contact1Uri)
        {
            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {
                    #region Expanding results

                    //The expand option retrieves related information.
                    //To retrieve information on associated entities in the same request, use the $expand
                    //query option on navigation properties.
                    //  1) Expand using single-valued navigation properties (e.g.: via the 'primarycontactid')
                    //  2) Expand using partner property (e.g.: from contact to account via the 'account_primary_contact')
                    //  3) Expand using collection-valued navigation properties (e.g.: via the 'contact_customer_accounts')
                    //  4) Expand using multiple navigation property types in a single request.
                    //  5) Multi-level expands

                    // Tip: For performance best practice, always use $select statement in an expand option.
                    Console.WriteLine("\n-- Expanding Results --");

                    //1) Expand using the 'primarycontactid' single-valued navigation property of account1.

                    JToken account1 = svc.Get($"{account1Uri}?" +
                        "$select=name&" +
                        "$expand=primarycontactid($select=fullname,jobtitle,annualincome)");

                    Console.WriteLine($"\nAccount {account1["name"]} has the following primary contact person:\n" +
                     $"\tFullname: {account1["primarycontactid"]["fullname"]} \n" +
                     $"\tJobtitle: {account1["primarycontactid"]["jobtitle"]} \n" +
                     $"\tAnnualincome: { account1["primarycontactid"]["annualincome"]}");

                    //2) Expand using the 'account_primary_contact' partner property.

                    JToken contact2 = svc.Get($"{contact1Uri}?$select=fullname,jobtitle,annualincome&" +
                    "$expand=account_primary_contact($select=name)");

                    Console.WriteLine($"\nContact '{contact2["fullname"]}' is the primary contact for the following accounts:");
                    foreach (JObject account in contact2["account_primary_contact"])
                    {
                        Console.WriteLine($"\t{account["name"]}");
                    }

                    //3) Expand using the collection-valued 'contact_customer_accounts' navigation property.

                    JToken account2 = svc.Get($"{account1Uri}?" +
                        "$select=name&" +
                        "$expand=contact_customer_accounts($select=fullname,jobtitle,annualincome)",
                        formattedValueHeaders);

                    WriteContactResultsTable(
                        $"Account '{account2["name"]}' has the following contact customers:",
                        account2["contact_customer_accounts"]);

                    //4) Expand using multiple navigation property types in a single request, specifically:
                    //   primarycontactid, contact_customer_accounts, and Account_Tasks.

                    Console.WriteLine("\n-- Expanding multiple property types in one request -- ");

                    JToken account3 = svc.Get($"{account1Uri}?$select=name&" +
                        "$expand=primarycontactid($select=fullname,jobtitle,annualincome)," +
                        "contact_customer_accounts($select=fullname,jobtitle,annualincome)," +
                        "Account_Tasks($select=subject,description)",
                        formattedValueHeaders);

                    Console.WriteLine($"\nAccount {account3["name"]} has the following primary contact person:\n" +
                                    $"\tFullname: {account3["primarycontactid"]["fullname"]} \n" +
                                    $"\tJobtitle: {account3["primarycontactid"]["jobtitle"]} \n" +
                                    $"\tAnnualincome: {account3["primarycontactid"]["annualincome"]}");

                    WriteContactResultsTable(
                        $"Account '{account3["name"]}' has the following contact customers:",
                        account3["contact_customer_accounts"]);

                    Console.WriteLine($"\nAccount '{account3["name"] }' has the following tasks:");

                    foreach (JObject task in account3["Account_Tasks"])
                    {
                        Console.WriteLine($"\t{task["subject"]}");
                    }

                    // 5) Multi-level expands

                    //The following query applies nested expands to single-valued navigation properties
                    // starting with Task entities related to contacts created for this sample.
                    JToken contosoTasks = svc.Get($"tasks?" +
                        $"$select=subject&" +
                        $"$filter=regardingobjectid_contact_task/_accountid_value eq {contosoId}" +
                        $"&$expand=regardingobjectid_contact_task($select=fullname;" +
                        $"$expand=parentcustomerid_account($select=name;" +
                        $"$expand=createdby($select=fullname)))",
                        formattedValueHeaders);

                    Console.WriteLine("\nExpanded values from Task:");

                    DisplayExpandedValuesFromTask(contosoTasks["value"]);

                    #endregion Expanding results
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static void retrieverecordswithAggregateresults(dynamic contosoId, dynamic formattedValueHeaders, dynamic account1Uri, dynamic contact1Uri)
        {
            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {

                    #region Aggregate results

                    //Get aggregated salary information about Contacts working for Contoso

                    Console.WriteLine("\nAggregated Annual Income information for Contoso contacts:");

                    JToken contactData = svc.Get($"{account1Uri}/contact_customer_accounts?" +
                        $"$apply=aggregate(annualincome with average as average, " +
                        $"annualincome with sum as total, " +
                        $"annualincome with min as minimum, " +
                        $"annualincome with max as maximum)", formattedValueHeaders);

                    Console.WriteLine($"\tAverage income: {contactData["value"][0]["average@OData.Community.Display.V1.FormattedValue"]}");
                    Console.WriteLine($"\tTotal income: {contactData["value"][0]["total@OData.Community.Display.V1.FormattedValue"]}");
                    Console.WriteLine($"\tMinium income: {contactData["value"][0]["minimum@OData.Community.Display.V1.FormattedValue"]}");
                    Console.WriteLine($"\tMaximum income: {contactData["value"][0]["maximum@OData.Community.Display.V1.FormattedValue"]}");



                    #endregion Aggregate results
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static void retrieverecordswithFetchXMLqueries(dynamic contosoId, dynamic formattedValueHeaders, dynamic account1Uri, dynamic contact1Uri)
        {
            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {
                    #region FetchXML queries

                    //Use FetchXML to query for all contacts whose fullname contains '(sample)'.
                    //Note: XML string must be URI encoded. For more information, see:
                    //https://docs.microsoft.com/powerapps/developer/common-data-service/webapi/retrieve-and-execute-predefined-queries#use-custom-fetchxml
                    Console.WriteLine("\n-- FetchXML -- ");
                    string fetchXmlQuery =
                        "<fetch mapping='logical' output-format='xml-platform' version='1.0' distinct='false'>" +
                          "<entity name ='contact'>" +
                            "<attribute name ='fullname' />" +
                            "<attribute name ='jobtitle' />" +
                            "<attribute name ='annualincome' />" +
                            "<order descending ='true' attribute='fullname' />" +
                            "<filter type ='and'>" +
                              "<condition value ='%(sample)%' attribute='fullname' operator='like' />" +
                              $"<condition value ='{contosoId.ToString()}' attribute='parentcustomerid' operator='eq' />" +
                            "</filter>" +
                          "</entity>" +
                        "</fetch>";
                    JToken contacts = svc.Get(
                        $"contacts?fetchXml={WebUtility.UrlEncode(fetchXmlQuery)}",
                        formattedValueHeaders);

                    WriteContactResultsTable($"Contacts Fetched by fullname containing '(sample)':", contacts["value"]);

                    #endregion FetchXML queries

                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static void retrieverecordswithpredefinedqueries(dynamic contosoId, dynamic formattedValueHeaders, dynamic account1Uri, dynamic contact1Uri)
        {
            try
            {
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {
                    #region Using predefined queries

                    //Use predefined queries of the following two types:
                    //  1) Saved query (system view)
                    //  2) User query (saved view)
                    //For more info, see:
                    //https://docs.microsoft.com/powerapps/developer/common-data-service/webapi/retrieve-and-execute-predefined-queries#predefined-queries

                    //1) Saved Query - retrieve "Active Accounts", run it, then display the results.
                    Console.WriteLine("\n-- Saved Query -- ");

                    JToken savedqueryid = svc.Get("savedqueries?" +
                        "$select=name,savedqueryid&" +
                        "$filter=name eq 'Active Accounts'")["value"][0]["savedqueryid"];

                    var activeAccounts = svc.Get(
                        $"accounts?savedQuery={savedqueryid}",
                        formattedValueHeaders)["value"] as JArray;

                    DisplayFormattedEntities("\nActive Accounts", activeAccounts, new string[] { "name" });

                    //2) Create a user query, then retrieve and execute it to display its results.
                    //For more info, see:
                    //https://docs.microsoft.com/powerapps/developer/common-data-service/saved-queries
                    Console.WriteLine("\n-- User Query -- ");
                    var userQuery = new JObject
                    {
                        ["name"] = "My User Query",
                        ["description"] = "User query to display contact info.",
                        ["querytype"] = 0,
                        ["returnedtypecode"] = "contact",
                        ["fetchxml"] = @"<fetch mapping='logical' output-format='xml-platform' version='1.0' distinct='false'>
                    <entity name ='contact'>
                        <attribute name ='fullname' />
                        <attribute name ='contactid' />
                        <attribute name ='jobtitle' />
                        <attribute name ='annualincome' />
                        <order descending ='false' attribute='fullname' />
                        <filter type ='and'>
                            <condition value ='%(sample)%' attribute='fullname' operator='like' />
                            <condition value ='%Manager%' attribute='jobtitle' operator='like' />
                            <condition value ='55000' attribute='annualincome' operator='gt' />
                        </filter>
                    </entity>
                 </fetch>"
                    };

                    //Create the saved query
                    Uri myUserQueryUri = svc.PostCreate("userqueries", userQuery);

                    //Retrieve the userqueryid
                    JToken myUserQueryId = svc.Get($"{myUserQueryUri}/userqueryid")["value"];
                    //Use the query to return results:
                    JToken myUserQueryResults = svc.Get($"contacts?userQuery={myUserQueryId}", formattedValueHeaders)["value"];

                    WriteContactResultsTable($"Contacts Fetched by My User Query:", myUserQueryResults);

                    #endregion Using predefined queries

                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        private static void WriteContactResultsTable(string message, JToken collection)
        {
            //Display column widths for contact results table
            const int col1 = -27;
            const int col2 = -35;
            const int col3 = -15;

            Console.WriteLine($"\n{message}");
            //header
            Console.WriteLine($"\t|{"Full Name",col1}|" +
                $"{"Job Title",col2}|" +
                $"{"Annual Income",col3}");
            Console.WriteLine($"\t|{new string('-', col1 * -1),col1}|" +
                $"{new string('-', col2 * -1),col2}|" +
                $"{new string('-', col3 * -1),col3}");
            //rows
            foreach (JObject contact in collection)
            {
                Console.WriteLine($"\t|{contact["fullname"],col1}|" +
                    $"{contact["jobtitle"],col2}|" +
                    $"{contact["annualincome@OData.Community.Display.V1.FormattedValue"],col3}");
            }
        }
        private static void DisplayFormattedEntities(string label, JArray entities, string[] properties)
        {
            /// <summary> Displays formatted entity collections to the console. </summary>
            /// <param name="label">Descriptive text output before collection contents </param>
            /// <param name="collection"> JObject containing array of entities to output by property </param>
            /// <param name="properties"> Array of properties within each entity to output. </param>
            Console.Write(label);
            int lineNum = 0;
            foreach (JObject entity in entities)
            {
                lineNum++;
                List<string> propsOutput = new List<string>();
                //Iterate through each requested property and output either formatted value if one
                //exists, otherwise output plain value.
                foreach (string prop in properties)
                {
                    string propValue;
                    string formattedProp = prop + "@OData.Community.Display.V1.FormattedValue";
                    if (null != entity[formattedProp])
                    {
                        propValue = entity[formattedProp].ToString();
                    }
                    else
                    {
                        if (null != entity[prop])
                        {
                            propValue = entity[prop].ToString();
                        }
                        else
                        {
                            propValue = "NULL";
                        }
                    }
                    propsOutput.Add(propValue);
                }
                Console.Write("\n\t{0}) {1}", lineNum, string.Join(", ", propsOutput));
            }
            Console.Write("\n");
        }
        private static void DisplayExpandedValuesFromTask(JToken collection)
        {

            //Display column widths for task Lookup Values Table
            const int col1 = -30;
            const int col2 = -30;
            const int col3 = -25;
            const int col4 = -20;

            Console.WriteLine($"\t|{"Subject",col1}|" +
            $"{"Contact",col2}|" +
            $"{"Account",col3}|" +
            $"{"Account CreatedBy",col4}");
            Console.WriteLine($"\t|{new string('-', col1 * -1),col1}|" +
                $"{new string('-', col2 * -1),col2}|" +
                $"{new string('-', col3 * -1),col3}|" +
                $"{new string('-', col4 * -1),col4}");
            //rows
            foreach (JObject task in collection)
            {
                Console.WriteLine($"\t|{task["subject"],col1}|" +
                    $"{task["regardingobjectid_contact_task"]["fullname"],col2}|" +
                    $"{task["regardingobjectid_contact_task"]["parentcustomerid_account"]["name"],col3}|" +
                    $"{task["regardingobjectid_contact_task"]["parentcustomerid_account"]["createdby"]["fullname"],col4}");

                //Console.WriteLine($"\n\tSubject: " +
                //    $"{task["subject"]}");
                //Console.WriteLine($"\t\tContact: " +
                //    $"{task["regardingobjectid_contact_task"]["fullname"]}");
                //Console.WriteLine($"\t\t\tAccount: " +
                //    $"{task["regardingobjectid_contact_task"]["parentcustomerid_account"]["name"]}");
                //Console.WriteLine($"\t\t\t\tAccount Created by: " +
                //    $"{task["regardingobjectid_contact_task"]["parentcustomerid_account"]["createdby"]["fullname"]}");

            }
        }


        #endregion Querydata







#region Conditional operations
        public static void Conditionaloperations()
        {

            // Save the URIs for entity records created in this sample. so they
            // can be deleted later.
            List<Uri> entityUris = new List<Uri>();

            try
            {
                // Use the wrapper class that handles message processing, error handling, and more.
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {
                    Console.WriteLine("--Starting conditional operations demonstration--\n");

                    #region Create required records
                    // Create an account record
                    var account1 = new JObject {
                        { "name", "Contoso Ltd" },
                        { "telephone1", "555-0000" }, //Phone number value increments with each update attempt
                        { "revenue", 5000000},
                        { "description", "Parent company of Contoso Pharmaceuticals, etc."} };

                    Uri account1Uri = svc.PostCreate("accounts", account1);
                    entityUris.Add(account1Uri); // Track any created records

                    // Retrieve the account record that was just created.
                    string queryOptions = "?$select=name,revenue,telephone1,description";
                    var retrievedaccount1 = svc.Get(account1Uri.ToString() + queryOptions);

                    // Store the ETag value from the retrieved record
                    string initialAcctETagVal = retrievedaccount1["@odata.etag"].ToString();

                    Console.WriteLine("Created and retrieved the initial account, shown below:");
                    Console.WriteLine(retrievedaccount1.ToString((Newtonsoft.Json.Formatting)Formatting.Indented));
                    #endregion Create required records

                    #region Conditional GET
                    Console.WriteLine("\n** Conditional GET demonstration **");

                    // Attempt to retrieve the account record using a conditional GET defined by a message header with
                    // the current ETag value.
                    try
                    {
                        retrievedaccount1 = svc.Get(
                            path: account1Uri.ToString() + queryOptions,
                            headers: new Dictionary<string, List<string>> {
                            { "If-None-Match", new List<string> {initialAcctETagVal}}}
                        );

                        // Not expected; the returned response contains content.
                        Console.WriteLine("Instance retrieved using ETag: {0}", initialAcctETagVal);
                        Console.WriteLine(retrievedaccount1.ToString((Newtonsoft.Json.Formatting)Formatting.Indented));
                    }
                    catch (AggregateException ae) // Message was not successful
                    {
                        ae.Handle((x) =>
                        {
                            if (x is ServiceException) // This we know how to handle.
                            {
                                var e = x as ServiceException;
                                if (e.StatusCode == (int)HttpStatusCode.NotModified) // Expected result
                                {
                                    Console.WriteLine("Account record retrieved using ETag: {0}", initialAcctETagVal);
                                    Console.WriteLine("Expected outcome: Entity was not modified so nothing was returned.");
                                    return true;
                                }
                            }
                            return false; // Let anything else stop the application.
                        });
                    }

                    // Modify the account instance by updating the telephone1 attribute
                    svc.Put(account1Uri.ToString(), "telephone1", "555-0001");
                    Console.WriteLine("\n\bAccount telephone number updated to '555-0001'.\n");

                    // Re-attempt to retrieve using conditional GET defined by a message header with
                    // the current ETag value.
                    try
                    {
                        retrievedaccount1 = svc.Get(
                            path: account1Uri.ToString() + queryOptions,
                            headers: new Dictionary<string, List<string>> {
                            { "If-None-Match", new List<string> {initialAcctETagVal}}}
                        );

                        // Expected result
                        Console.WriteLine("Modified account record retrieved using ETag: {0}", initialAcctETagVal);
                        Console.WriteLine("Notice the update ETag value and telephone number");
                    }
                    catch (ServiceException e)
                    {
                        if (e.StatusCode == (int)HttpStatusCode.NotModified) // Not expected
                        {
                            Console.WriteLine("Unexpected outcome: Entity was modified so something should be returned.");
                        }
                        else { throw e; }
                    }

                    // Save the updated ETag value
                    var updatedAcctETagVal = retrievedaccount1["@odata.etag"].ToString();

                    // Display ty updated record
                    Console.WriteLine(retrievedaccount1.ToString((Newtonsoft.Json.Formatting)Formatting.Indented));
                    #endregion Conditional GET

                    #region Optimistic concurrency on delete and update
                    Console.WriteLine("\n** Optimistic concurrency demonstration **");

                    // Attempt to delete original account (if matches original ETag value).
                    // If you replace "initialAcctETagVal" with "updatedAcctETagVal", the delete will
                    // succeed. However, we want the delete to fail for now to demonstrate use of the ETag.
                    Console.WriteLine("Attempting to delete the account using the original ETag value");

                    try
                    {
                        svc.Delete(
                            account1Uri.ToString(),
                            headers: new Dictionary<string, List<string>> {
                               { "If-Match", new List<string> {initialAcctETagVal}}}
                        );

                        // Not expected; this code should not execute.
                        Console.WriteLine("Account deleted");
                    }
                    catch (ServiceException e)
                    {
                        if (e.StatusCode == (int)HttpStatusCode.PreconditionFailed) // Expected result
                        {
                            Console.WriteLine("Expected Error: The version of the account record no" +
                                " longer matches the original ETag.");
                            Console.WriteLine("\tAccount not deleted using ETag '{0}', status code: '{1}'.",
                                initialAcctETagVal, e.StatusCode);
                        }
                        else { throw e; }
                    }

                    Console.WriteLine("Attempting to update the account using the original ETag value");
                    JObject accountUpdate = new JObject() {
                        { "telephone1", "555-0002" },
                        { "revenue", 6000000 }
                    };

                    try
                    {
                        svc.Patch(
                            uri: account1Uri,
                            body: accountUpdate,
                            headers: new Dictionary<string, List<string>> {
                            { "If-Match", new List<string> {initialAcctETagVal}}}
                        );

                        // Not expected; this code should not execute.
                        Console.WriteLine("Account updated using original ETag {0}", initialAcctETagVal);
                    }
                    catch (ServiceException e)
                    {
                        if (e.StatusCode == (int)HttpStatusCode.PreconditionFailed) // Expected error
                        {
                            Console.WriteLine("Expected Error: The version of the existing record doesn't "
                                + "match the ETag property provided.");
                            Console.WriteLine("\tAccount not updated using ETag '{0}', status code: '{1}'.",
                              initialAcctETagVal, (int)e.StatusCode);
                        }
                        else { throw e; }
                    }

                    // Reattempt update if matches current ETag value.
                    accountUpdate["telephone1"] = "555-0003";
                    Console.WriteLine("Attempting to update the account using the current ETag value");
                    try
                    {
                        svc.Patch(
                            uri: account1Uri,
                            body: accountUpdate,
                            headers: new Dictionary<string, List<string>> {
                                { "If-Match", new List<string> { updatedAcctETagVal }} }
                        );

                        // Expected program flow; this code should execute.
                        Console.WriteLine("\nAccount successfully updated using ETag: {0}.",
                            updatedAcctETagVal);
                    }
                    catch (ServiceException e)
                    {
                        if (e.StatusCode == (int)HttpStatusCode.PreconditionFailed) // Not expected
                        {
                            Console.WriteLine("Unexpected status code: '{0}'", (int)e.StatusCode);
                        }
                        else { throw e; }
                    }

                    // Retrieve and output current account state.
                    retrievedaccount1 = svc.Get(account1Uri.ToString() + queryOptions);

                    Console.WriteLine("\nBelow is the final state of the account");
                    Console.WriteLine(retrievedaccount1.ToString((Newtonsoft.Json.Formatting)Formatting.Indented));
                    #endregion Optimistic concurrency on delete and update

                    #region Delete created records

                    // Delete (or keep) all the created entity records.
                    Console.Write("\nDo you want these entity records deleted? (y/n) [y]: ");
                    String answer = Console.ReadLine().Trim();

                    if (!(answer.StartsWith("y") || answer.StartsWith("Y") || answer == string.Empty))
                        entityUris.Clear();

                    foreach (Uri entityUrl in entityUris) svc.Delete(entityUrl.ToString());

                    #endregion Delete created records 
                }
            }
            catch (ServiceException e)
            {
                Console.WriteLine("Message send response: status code {0}, {1}",
                    e.StatusCode, e.ReasonPhrase);
            }
        }
        #endregion Conditional operations

        #region   Functions and Actions

        public static void FunctionsandActions(Uri account1Uri)
        {
            Console.Title = "Function and Actions demonstration";

            // Track entity instance URIs so those records can be deleted prior to exit.
            Dictionary<string, Uri> entityUris = new Dictionary<string, Uri>();

            try
            {
                // Get environment configuration information from the connection string in App.config.
                ServiceConfig config = new ServiceConfig(
                    ConfigurationManager.ConnectionStrings["Connect"].ConnectionString);

                // Use the service class that handles HTTP messaging, error handling, and
                // performance optimizations.
                using (CDSWebApiService svc = new CDSWebApiService(config))
                {
                   

                    #region Call an unbound function with no parameters
                    Console.WriteLine("\n* Call an unbound function with no parameters.");

                    // Retrieve the current user's full name from the WhoAmI function:
                    Console.Write("\tGetting information on the current user..");
                    JToken currentUser = svc.Get("WhoAmI");

                    // Obtain the user's ID and full name
                    JToken user = svc.Get("systemusers(" + currentUser["UserId"] + ")?$select=fullname");

                    Console.WriteLine("completed.");
                    Console.WriteLine("\tCurrent user's full name is '{0}'.", user["fullname"]);
                    #endregion Call an unbound function with no parameters

                    #region Call an unbound function that requires parameters
                    Console.WriteLine("\n* Call an unbound function that requires parameters");

                    // Retrieve the code for the specified time zone
                    int localeID = 1033;
                    string timeZoneName = "Pacific Standard Time";

                    // Define the unbound function and its parameters
                    string[] uria = new string[] {
                        "GetTimeZoneCodeByLocalizedName",
                        "(LocalizedStandardName=@p1,LocaleId=@p2)",
                        "?@p1='" + timeZoneName + "'&@p2=" + localeID };

                    // This would also work:
                    // string[] uria = ["GetTimeZoneCodeByLocalizedName", "(LocalizedStandardName='" + 
                    //    timeZoneName + "',LocaleId=" + localeId + ")"]; 

                    JToken localizedName = svc.Get(string.Join("", uria));
                    string timeZoneCode = localizedName["TimeZoneCode"].ToString();

                    Console.WriteLine(
                        "\tThe time zone '{0}' has the code '{1}'.", timeZoneName, timeZoneCode);
                    #endregion Call an unbound function that requires parameters

                    #region Call a bound function   
                    Console.WriteLine("\n* Call a bound function");

                    // Retrieve the total time (minutes) spent on all tasks associated with 
                    // incident "Sample Case".
                    string boundUri = entityUris["Sample Case"] +
                        @"/Microsoft.Dynamics.CRM.CalculateTotalTimeIncident()";

                    JToken cttir = svc.Get(boundUri);
                    string totalTime = cttir["TotalTime"].ToString();

                    Console.WriteLine("\tThe total duration of tasks associated with the incident " +
                        "is {0} minutes.", totalTime);
                    #endregion Call a bound function 

                    #region Call an unbound action that requires parameters
                    Console.WriteLine("\n* Call an unbound action that requires parameters");

                    // Close the existing opportunity "Opportunity to win" and mark it as won.
                    JObject opportClose = new JObject()
                    {
                        { "subject", "Won Opportunity" },
                        { "opportunityid@odata.bind", entityUris["Opportunity to win"] }
                    };

                    JObject winOpportParams = new JObject()
                    {
                        { "Status", "3" },
                        { "OpportunityClose", opportClose }
                    };

                    JObject won = svc.Post("WinOpportunity", winOpportParams);

                    Console.WriteLine("\tOpportunity won.");
                    #endregion Call an unbound action that requires parameters

                    #region Call a bound action that requires parameters
                    Console.WriteLine("\n* Call a bound action that requires parameters");

                    // Add a new letter tracking activity to the current user's queue.
                    // First create a letter tracking instance.
                    JObject letterAttributes = new JObject()
                    {
                        {"subject", "Example letter" },
                        {"description", "Body of the letter" }
                    };

                    Console.Write("\tCreating letter 'Example letter'..");

                    Uri letterUri = svc.PostCreate("letters", letterAttributes);
                    entityUris.Add("Example letter", letterUri);

                    Console.WriteLine("completed.");

                    //Retrieve the ID associated with this new letter tracking activity.
                    JToken letter = svc.Get(letterUri + "?$select=activityid,subject");
                    string letterActivityId = (string)letter["activityid"];

                    // Retrieve the URL to current user's queue.
                    string myUserId = (string)currentUser["UserId"];

                    JToken queueRef = svc.Get("systemusers(" + myUserId + ")/queueid/$ref");
                    string myQueueUri = (string)queueRef["@odata.id"];

                    //Add the letter activity to current user's queue, then return its queue ID.
                    JObject targetUri = JObject.Parse(
                      @"{activityid: '" + letterActivityId + @"', '@odata.type': 'Microsoft.Dynamics.CRM.letter' }");

                    JObject addToQueueParams = new JObject()
                    {
                        { "Target", targetUri }
                    };

                    string queueItemId = (string)svc.Post(
                        myQueueUri + "/Microsoft.Dynamics.CRM.AddToQueue", addToQueueParams)["QueueItemId"];

                    Console.WriteLine("\tLetter 'Example letter' added to current user's queue.");
                    Console.WriteLine("\tQueueItemId returned from AddToQueue action: {0}", queueItemId);
                    #endregion Call a bound action that requires parameters

                    #region Call a bound custom action that requires parameters
                    Console.WriteLine("\n* Call a bound custom action that requires parameters");

                    // Add a note to a specified contact. Uses the custom action sample_AddNoteToContact, which
                    // is bound to the contact to annotate, and takes a single param, the note to add. It also  
                    // returns the URI to the new annotation. 

                    JObject note = JObject.Parse(
                        @"{NoteTitle: 'Sample note', NoteText: 'The text content of the note.'}");
                    string actionUri = entityUris["Jon Fogg"].ToString() + "/Microsoft.Dynamics.CRM.sample_AddNoteToContact";

                    JObject contact = svc.Post(actionUri, note);
                    Uri annotationUri = new Uri(svc.BaseAddress + "annotations(" + contact["annotationid"] + ")");
                    entityUris.Add((string)note["NoteTitle"], annotationUri);

                    Console.WriteLine("\tA note with the title '{0}' was created and " +
                        "associated with the contact 'Jon Fogg'.", note["NoteTitle"]);
                    #endregion Call a bound custom action that requires parameters

                    #region Call an unbound custom action that requires parameters
                    Console.WriteLine("\n* Call an unbound custom action that requires parameters");

                    // Create a customer of the specified type, using the custom action sample_CreateCustomer,
                    // which takes two parameters: the type of customer ('account' or 'contact') and the name of 
                    // the new customer.
                    string customerName = "New account customer (sample)";
                    JObject customerAttributes = JObject.Parse(
                        @"{CustomerType: 'account', AccountName: '" + customerName + "'}");

                    JObject response = svc.Post("sample_CreateCustomer", customerAttributes);
                    Console.WriteLine("\tThe account '" + customerName + "' was created.");

                    // Because the CreateCustomer custom action does not return any data about the created instance, 
                    // we must query the customer instance to figure out its URI.
                    JToken customer = svc.Get("accounts?$filter=name eq 'New account customer (sample)'&$select=accountid&$top=1");
                    Uri customerUri = new Uri(svc.BaseAddress + "accounts(" + customer["value"][0]["accountid"] + ")");
                    entityUris.Add(customerName, customerUri);

                    // Try to call the same custom action with invalid parameters, here the same name is
                    // not valid for a contact. (ContactFirstname and ContactLastName parameters are  
                    // required when CustomerType is contact.
                    customerAttributes = JObject.Parse(
                        @"{CustomerType: 'contact', AccountName: '" + customerName + "'}");

                    try
                    {
                        customerUri = svc.PostCreate("sample_CreateCustomer", customerAttributes);
                        Console.WriteLine("\tCall to the custom CreateCustomer action succeeded, which was not expected.");
                    }
                    catch (AggregateException e)
                    {
                        Console.WriteLine("\tCall to the custom CreateCustomer action did not succeed (as was expected).");
                        foreach (Exception inner in (e as AggregateException).InnerExceptions)
                        { Console.WriteLine("\t  -" + inner.Message); }
                    }
                    #endregion Call an unbound custom action that requires parameters

                  
                }
            }
            catch (Exception e)
            {
                Console.BackgroundColor = ConsoleColor.Red; // Highlight exceptions
                if (e is AggregateException)
                {
                    foreach (Exception inner in (e as AggregateException).InnerExceptions)
                    { Console.WriteLine("\n" + inner.Message); }
                }
                else if (e is ServiceException)
                {
                    var ex = e as ServiceException;
                    Console.WriteLine("\nMessage send response: status code {0}, {1}",
                        ex.StatusCode, ex.ReasonPhrase);
                }
                Console.ReadKey(); // Pause terminal
            }
        }

        #endregion Functions and Actions


        static void Main(string[] args)
        {
        }
    }
}
#endregion complete