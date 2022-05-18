# Leave Request SPFx adaptive card example

## Summary

This solution provides a basic Leave Request system, which stores its data in two lists:

* LeaveRequest
* LeaveBalance

When the adaptive card loads for the first time, the solution will attempt to provision the two lists in the current site.  

Once active, the solution will:

* show the current user a balance of their leave according to a list item in the **LeaveBalance** list
* clicking the card opens a quick view to enter start date and return to work date
* clicking the "Calculate hours" button calculates the amount of leave to be taken (calculated at 7.8 hours per day, taking weekends into account)
* the user can then click the "Submit for approval" button, which creates a list item in the **LeaveRequest** list

![Screen grab of the Leave Request Adaptive Card in action on a Viva Connections Dashboard page](./demo.gif)

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.13-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

### Viva Connections

Viva Connections must be configured for your tenant. [Follow these instructions to complete the Viva Connections configuration](https://docs.microsoft.com/viva/connections/guide-to-setting-up-viva-connections).

### First use and permissions to provision SharePoint lists

When the Adpative Card is first rendered, the solution will attempt to provision two lists to the current site if they don't already exist.  The current user has access to perform this task the first time round.  The Adaptive Card will not attempt to provision the lists if they are already provisioned.


## Solution

Solution|Author(s)
--------|---------
adaptiveCardExtensions\leaveRequest | Mark Powney [@mpowney](https://twitter.com/mpowney)

## Version history

Version|Date|Comments
-------|----|--------
0.1|May 18, 2022|First commit

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Browse 

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- topic 1
- topic 2
- topic 3

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development