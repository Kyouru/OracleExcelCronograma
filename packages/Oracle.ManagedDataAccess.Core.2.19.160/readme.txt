Oracle.ManagedDataAccess.Core NuGet Package 2.19.160 README
===========================================================
Release Notes: Oracle Data Provider for .NET Core

June 2022

This provider supports .NET Core 3.1 and .NET 6.

This document provides information that supplements the Oracle Data Provider for .NET (ODP.NET) documentation.

You have downloaded Oracle Data Provider for .NET. The license agreement is available here:
https://www.oracle.com/downloads/licenses/distribution-license.html


New Features
============
- Azure Active Directory

ODP.NET now supports Azure Active Directory (AAD) authentication when connecting to Oracle Database. 
ODP.NET will then use an access token to authenticate instead of a username and password.

This feature benefits applications and services that use AAD for centralized user authentication 
with Oracle database. Those services can include Azure and Microsoft 365-based cloud services, 
such as Microsoft Power BI service, that rely on AAD for user management.

Using token-based authentication is more secure and simpler for the end user. It becomes unnecessary 
to specify credentials each time the user accesses a resource.  Moreover, the resource never needs 
to handle and manage individual user credentials.

NOTE: The application project will need to explictly add the System.Text.Json 
nuget package version 5.0.2 (or higher) to utilize the AAD authentication feature.

- TLS 1.3

ODP.NET now supports TLS 1.3.


Bug Fixes since Oracle.ManagedDataAccess.Core NuGet Package 2.19.140
====================================================================
Bug 34322469 - CONNECTION POOL THROWS "CONNECTION REQUEST TIMED OUT" EXCEPTION DUE TO LOOPING WITHIN POOLMANAGER.GET() 
Bug 34143076 - UDT: EVOLVED TYPE USAGE LEADS TO ORA-00932 INCONSISTENT DATATYPES
Bug 33948562 - EMPTY LOB RETURNS AS NULL VALUE WHEN INITIALLOBFETCHSIZE IS SET TO -1 
Bug 33509136 - "UNEXPECTED PACKET RECEIVED" EXCEPTION UNEXPECTEDLY THROWN WITH "ORA-01013 USER CANCELLED REQUEST" IN THE TRACES
Bug 32843859 - ORA-01006: BIND VARIABLE DOES NOT EXIST ERROR OCCURS WHEN DERIVEPARAMETERS USED WITH DIFFERENT DB VERSIONS
Bug 31806772 - ORACLECOMMANDBUILDER WITH BINDBYNAME IN CONFIG CAUSES EXCEPTION


Known Issues and Limitations
============================
1) BindToDirectory throws NullReferenceException on Linux when LdapConnection AuthType is Anonymous

https://github.com/dotnet/runtime/issues/61683

This issue is observed when using System.DirectoryServices.Protocols, version 6.0.0.
To workaround the issue, use System.DirectoryServices.Protocols, version 5.0.1.

 Copyright (c) 2021, 2022, Oracle and/or its affiliates. 
