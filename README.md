# Project CSOM Read Enterprise CustomFields

The Project CSOM Read Enterprise CustomFields sample uses C# and Project CSOM to demonstrate how to read enterprise custom fields (ECFs) that are defined at the Project Online Web App (PWA) instance and read the ECFs that are defined in each project.

Users typically access ECFs by viewing the Project Details page for a specific project stored in the PWA instance.  

## Scenario

As a Project/Program/Portfolio portfolio manager, I would like to use an app that displays the custom values my company has associated with our projects.


## Using the app

1.	Add the Project CSOM client package [here](https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM/).
2.	Update the PWA site.
3.	Update the login/password to your PWA site.
4.	Run the app.

## Prerequisites
To use this code sample, you need the following:

* PWA Site (Project Online, Project Server 2013 or Project Server 2016)
* Visual Studio 2013 or later 
* Project CSOM client library.  It is available as a Nuget Package from [here](https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM/)
* One or more project stored in the PWA instance that use ECFs.


## How the sample affects your tenant data
This sample runs CSOM methods that reads all projects in the PWA instance for the specified user. Tenant data will not be changed nor deleted.

## Additional resources
* [Local and Enterprise Custom Fields](https://msdn.microsoft.com/en-us/library/office/ms447495(v=office.14).aspx)

* [ProjectContext class](https://msdn.microsoft.com/en-us/library/office/microsoft.projectserver.client.projectcontext_di_pj14mref.aspx)

* [Client-side object model (CSOM) for Project 2013](https://aka.ms/project-csom-docs)

* [SharePoint (Project) CSOM Client library](https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM/)

## Copyright

Copyright (c) 2016 Microsoft. All rights reserved.



This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
