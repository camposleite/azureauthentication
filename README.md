## Azure Authentication with React, MSAL (Microsoft Authentication Library) and Microsoft Graph

#### Tools

- Microsoft Authentication Library: `npm install msal`
- Microsoft Graph JavaScript Client Library: `npm install @microsoft/microsoft-graph-client`
- Microsoft Graph TypeScript Types: `npm install @types/microsoft-graph`

#### App Registration

![App Registration](images/appreg01.jpg)

#### The Registered App with Its Application Id

![Registered App](images/appreg02.jpg)

#### API Permissions

These are the scopes you request a token for

![API Permissions](images/apiperm.jpg)

#### Enable Implicit Flow in Manifest

![Implicit Flow](images/implicitflow.jpg)

#### Final thoughts

There's more than one way to skin a cat. This is just a simple way to to use MSAL to authenticate in Azure and query MS Graph.

Keep on!
