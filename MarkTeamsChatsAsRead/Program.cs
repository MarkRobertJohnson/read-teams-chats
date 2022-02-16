using Azure.Identity;
using CommandLine;
using MarkTeamsChatsAsRead;
using Microsoft.Graph;

var scopes = new[] { "openid", "profile", "User.Read", "Chat.ReadBasic", "Chat.Read", "Chat.ReadWrite" };
var appId = "905e5f03-4d28-4e09-903a-bab4ac2f613a";
var tenantId = "db84cc25-a34d-48ae-81b9-7d26ab3769eb";

var options = new InteractiveBrowserCredentialOptions
{
    TenantId = tenantId,
    ClientId = appId,
    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
    // MUST be http://localhost or http://localhost:PORT
    // See https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/System-Browser-on-.Net-Core
    RedirectUri = new Uri("http://localhost"),
};

var parsedOptions = Parser.Default.ParseArguments<CommandLineOptions>(args)
            .WithParsed<CommandLineOptions>(o =>
           {
               if (o.All)
               {
                   Console.WriteLine("Force reading all chats...");
               }
           });

var interactiveCredential = new InteractiveBrowserCredential(options);
var graphClient = new GraphServiceClient(interactiveCredential, scopes);
var user = await graphClient.Me.Request().GetAsync();


Console.WriteLine($"Performing all Teams operations as {user.UserPrincipalName}");
var allChats = await TeamsUtils.GetAllUserChats(graphClient);

var readChats = await TeamsUtils.MarkAllUnreadChatsAsRead(graphClient, allChats, user.Id, tenantId, parsedOptions.Value.All, parsedOptions.Value.Parellelism);

Console.Write($"All unread chats have been marked as read ({readChats.Count}).  {allChats.Count} total chats.  Hit enter to continue.");
Console.ReadLine();
