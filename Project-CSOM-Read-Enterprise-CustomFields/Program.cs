using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;

/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
namespace ProjectCSOMReadEnterpriseCustomFields
{
    class GetEnterpriseCustomFields
    {
        const int PROJECT_BLOCK_SIZE = 20;

        private static readonly string SiteUrl = "https://contoso.sharepoint.com/sites/pwa";

        private static ProjectContext projContext =
            new ProjectContext(SiteUrl);

        //Dictionary of ECFs uses FieldType to drive processing of the custom fields
        static Dictionary<String, CustomField> pwaECF = new Dictionary<string, CustomField>();

        static void Main(string[] args)
        {
            using (projContext)
            {
                // Get login cookie using WebLogin
                var cookies = WebLogin.GetWebLoginCookie(new Uri(SiteUrl));

                projContext.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
                {
                    e.WebRequestExecutor.WebRequest.CookieContainer = new CookieContainer();
                    e.WebRequestExecutor.WebRequest.CookieContainer.SetCookies(new Uri(SiteUrl), cookies);
                };

                // 1. List the defined Enterprise Custom Fields in the PWA instance.
                ListPWACustomFields();

                // 2. Get project list with minimal information
                projContext.Load(projContext.Projects, qp => qp.Include(qr => qr.Id));
                projContext.ExecuteQuery();

                var allIds = projContext.Projects.Select(p => p.Id).ToArray();

                int numBlocks = allIds.Length / PROJECT_BLOCK_SIZE + 1;

                // Query all the child objects in blocks of PROJECT_BLOCK_SIZE
                for (int i = 0; i < numBlocks; i++)
                {
                    var idBlock = allIds.Skip(i * PROJECT_BLOCK_SIZE).Take(PROJECT_BLOCK_SIZE);
                    Guid[] block = new Guid[PROJECT_BLOCK_SIZE];
                    Array.Copy(idBlock.ToArray(), block, idBlock.Count());

                    // 2. Retrieve and save project basic and custom field properties in an IEnumerable collection.
                    var projBlk = projContext.LoadQuery(
                        projContext.Projects
                        .Where(p =>   // some elements will be Zero'd guids at the end
                            p.Id == block[0] || p.Id == block[1] ||
                            p.Id == block[2] || p.Id == block[3] ||
                            p.Id == block[4] || p.Id == block[5] ||
                            p.Id == block[6] || p.Id == block[7] ||
                            p.Id == block[8] || p.Id == block[9] ||
                            p.Id == block[10] || p.Id == block[11] ||
                            p.Id == block[12] || p.Id == block[13] ||
                            p.Id == block[14] || p.Id == block[15] ||
                            p.Id == block[16] || p.Id == block[17] ||
                            p.Id == block[18] || p.Id == block[19]
                        )
                        .Include(p => p.Id,
                            p => p.Name,
                            p => p.IncludeCustomFields,
                            p => p.IncludeCustomFields.CustomFields,
                            P => P.IncludeCustomFields.CustomFields.IncludeWithDefaultProperties(
                                lu => lu.LookupTable,
                                lu => lu.LookupEntries
                            )
                        )
                    );

                    projContext.ExecuteQuery();

                    foreach (PublishedProject pubProj in projBlk)
                    {

                        // Set up access to custom field collection of published project
                        var projECFs = pubProj.IncludeCustomFields.CustomFields;

                        // Set up access to custom field values of published project
                        Dictionary<string, object> ECFValues = pubProj.IncludeCustomFields.FieldValues;

                        Console.WriteLine("Name:\t{0}",pubProj.Name);
                        Console.WriteLine("Id:\t{0}", pubProj.Id);
                        Console.WriteLine("ECF count: {0}\n", ECFValues.Count);

                        Console.WriteLine("\n\tType\t   Name\t\t       L.UP   Value                  Description");
                        Console.WriteLine("\t--------   ----------------    ----   --------------------   -----------");

                        foreach (CustomField cf in projECFs)
                        {

                            // 3A. Distinguish CF values that are simple from those that use entries in lookup tables.
                            if (!cf.LookupTable.ServerObjectIsNull.HasValue ||
                                                (cf.LookupTable.ServerObjectIsNull.HasValue && cf.LookupTable.ServerObjectIsNull.Value))
                            {
                                if (ECFValues[cf.InternalName] == null)
                                {   // 3B. Partial implementation. Not usable.
                                    String textValue = "is not set";
                                    Console.WriteLine("\t{0, -8}   {1, -20}        ***{2}",
                                        cf.FieldType, cf.Name, textValue);
                                }
                                else   // 3C. Simple, friendly value for the user
                                {
                                    // CustomFieldType is a CSOM enumeration of ECF types.
                                    switch (cf.FieldType)
                                    {
                                        
                                        case CustomFieldType.COST:
                                            decimal costValue = (decimal)ECFValues[cf.InternalName];
                                            Console.WriteLine("\t{0, -8}   {1, -20}        {2, -22}",
                                                cf.FieldType, cf.Name, costValue.ToString("C"));
                                            break;

                                        case CustomFieldType.DATE:
                                        case CustomFieldType.FINISHDATE:
                                        case CustomFieldType.DURATION:
                                        case CustomFieldType.FLAG:
                                        case CustomFieldType.NUMBER:
                                        case CustomFieldType.TEXT:
                                            Console.WriteLine("\t{0, -8}   {1, -20}        {2, -22}",
                                                cf.FieldType, cf.Name, ECFValues[cf.InternalName]);
                                            break;

                                    }

                                }
                            }
                            else         //3D. The ECF uses a Lookup table to store the values.
                            {
                                Console.Write("\t{0, -8}   {1, -20}", cf.FieldType, cf.Name);

                                String[] entries = (String[])ECFValues[cf.InternalName];

                                if (entries != null)
                                {
                                    foreach (String entry in entries)
                                    {
                                        var luEntry = projContext.LoadQuery(cf.LookupTable.Entries
                                                .Where(e => e.InternalName == entry));

                                        projContext.ExecuteQuery();

                                        Console.WriteLine(" Yes    {0, -22}  {1}", luEntry.First().FullValue, luEntry.First().Description);
                                    }
                                }
                            }
                        }

                        Console.WriteLine("     ------------------------------------------------------------------------\n");

                    }

                }

            }    //end of using


            Console.Write("\nPress any key to exit: ");
            Console.ReadKey(false);

        } //end of Main


        private static void ListPWACustomFields()
        {
            // Retrieves and lists Enterprise Custom Fields (ECFs) defined in the Project 
            // PWA instance. 

            // The ECF properties retrieved include the following: 
            // - InternalName is the identifier for the custom field.
            // - Name is the friendly name recognized by users.
            // - FieldType denotes data type: Text, Cost, Number, etc.

            var allECFields = projContext.LoadQuery(projContext.CustomFields.Include(
                    qp => qp.InternalName,
                    qp => qp.Name,
                    qp => qp.FieldType,
                    qp => qp.LookupTable,
                    qp => qp.EntityType.Name
                )
                /*  Filter for project-associated ECFs: Project vs Task vs Resource */
                /*  .Where(qr => qr.EntityType.Name == projContext.EntityTypes.ProjectEntity.Name)  */
                .OrderBy(qp => qp.EntityType.Name));

            projContext.ExecuteQuery();

            Console.WriteLine("\n     Enterprise Custom Field (ECF) definitions for Projects, Resources, and Tasks: {0}", allECFields.Count());

            Console.WriteLine("\n     ECF Name\t\t   InternalName\t\t\t\t      ECF Type   Association");
            Console.WriteLine("     -------------------   ----------------------------------------   --------   -----------");

            int i = 1;

            foreach (CustomField ECF in allECFields)
            {
                pwaECF[ECF.InternalName] = ECF;

                Console.WriteLine("{0,3}. {1, -22}{2}    {3,-7}    {4}", i++, ECF.Name, ECF.InternalName, ECF.FieldType, ECF.EntityType.Name);
            }

            Console.WriteLine("     --------------------------------------------------------------------------------------\n");

        }   // End of ListPWACustomFields


    }
}
