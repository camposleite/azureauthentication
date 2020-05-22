## Azure Authentication with React, MSAL (Microsoft Authentication Library) and Microsoft Graph

This article is a compilation and simplified version of the basic code and configuration needed for an React application to authenticate in Microsoft Azure Active Directory Single Sign-On. 

#### Tools

Front-end
- Microsoft Authentication Library: `npm install msal`
- Microsoft Graph JavaScript Client Library: `npm install @microsoft/microsoft-graph-client`
- Microsoft Graph TypeScript Types: `npm install @types/microsoft-graph`

Back-end (Installed from Visual Studio NuGet Packages)
- Microsoft Identity Client
- Microsoft Graph

### Azure Configuration

#### App Registration

First, you need to register your app.

![App Registration](images/appreg01.jpg)

#### The Registered App with Its Application Id

![Registered App](images/appreg02.jpg)

#### API Permissions

These are the scopes you request a token for

![API Permissions](images/apiperm.jpg)

#### Enable Implicit Flow in Manifest

![Implicit Flow](images/implicitflow.jpg)

#### Client Secrets

This is useful when you want to query MS Graph with permission scopes of type Application, not Delegated. When you want
to query things in the context of the application and not the user's context and token.
For example: when querying a list of users with the scope User.ReadBasic.All, maybe you would not want to give your users the permission to read other users profiles and want the application to be responsible for that.

![Client secret](images/secret.jpg)

#### Note from Microsoft

![Note](images/note.jpg)

So, you should use your back-end to query scopes of type Application, using the app secret.

#### Querying Calendars

![Implicit Flow](images/getcalendar.jpg)

#### Querying From The Back-end (C#)

Here is an example of how you can querya a list of users with the context of your application (api permission/scope of type Application), using the app secret. Remember: here we are not using the user's credentials with a scope permission of type Delegated.

   ```csharp
   
   private List<Microsoft.Graph.User> GetUserByName(string name)
        {
            GraphServiceClient graphServiceClient = GetGraphServiceClient();

            string filter = string.Format("startswith(displayName,'{0}')
            or startswith(givenName,'{0}')
            or startswith(surname,'{0}')
            or startswith(mail,'{0}')
            or startswith(userPrincipalName,'{0}')", name);

            IGraphServiceUsersCollectionPage users = graphServiceClient.Users.Request()
                .Select(u => new
                {
                    u.DisplayName,
                    u.UserPrincipalName,
                    u.Mail,
                    u.OnPremisesDistinguishedName,
                    u.OnPremisesSamAccountName
                })
                .Filter(filter)
                .GetAsync().Result;

            return users.ToList();            
        }
        
  private GraphServiceClient GetGraphServiceClient()
        {
            var confidentialClient = ConfidentialClientApplicationBuilder.Create(_client_id)
                            .WithAuthority(AzureCloudInstance.AzurePublic, _tenant_id)
                            .WithClientSecret(_client_secret)
                            .Build();

            string[] scopes = new string[] { _graph_default_url };

            GraphServiceClient graphServiceClient =
             new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
             {

                 // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
                 var authResult = await confidentialClient
                     .AcquireTokenForClient(scopes)
                     .ExecuteAsync();

                 // Add the access token in the Authorization header of the API request.
                 requestMessage.Headers.Authorization =
                     new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
             })
             );
            return graphServiceClient;
        }
        
```

#### Final thoughts

There's more than one way to skin a cat. This is just a simple way to to use MSAL to authenticate in Azure and query MS Graph.

Keep on!
