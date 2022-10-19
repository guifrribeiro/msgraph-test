using System.Collections.Generic;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Extensions.Configuration;
using Helpers;
using Azure.Identity;
using Azure.Core;

namespace msgraph
{
  class Program
  {
    static async Task Main(string[] args)
    {
      Console.WriteLine("-------------------------------MS Graph Tests-------------------------------");
      var config = LoadAppSettings();
      if (config == null)
      {
        Console.WriteLine("Invalid appsettings.json file");
        return;
      }

      Console.WriteLine("---------------OWNER Project informations---------------");
      var clientUser = GetAuthenticatedGraphClient(config);

      var userRequest = clientUser.Users.Request();

      var userResults = userRequest.GetAsync().Result;
      var userId = string.Empty;

      foreach(var user in userResults)
      {
        Console.WriteLine($"{user.Id}: {user.UserPrincipalName} <{user.Mail}>");
        userId = user.Id;
      }

      Console.WriteLine("\nGraph Request");
      Console.WriteLine(userRequest.GetHttpRequestMessage().RequestUri);

      Console.WriteLine("---------------MAIL Informations---------------");
      var clientMessage = GetAuthenticatedGraphClient(config);
      var messageRequest = clientMessage.Users["mariadmin@3330sc.onmicrosoft.com"].MailFolders["Inbox"].Messages.Request();
      var results = messageRequest.GetAsync().Result;

      foreach (var msg in results)
      {
        Console.WriteLine($"{msg.Subject}: {msg.Body} <{msg.Sender}>");
      }

      Console.WriteLine("\nMessage Request");
      Console.WriteLine(messageRequest.GetHttpRequestMessage().RequestUri);

      Console.WriteLine("---------------SEND Mail---------------");
      var clientSend = GetAuthenticatedGraphClient(config);
      
      var message = new Message
      {
        Subject = "Teste Microsoft Graph",
        Body = new ItemBody
        {
          ContentType = BodyType.Html,
          Content = "Esta é apenas uma mensagem simples de testes de uso do pacote Microsoft.Graph para envio de e-mails."
        },
        ToRecipients = new List<Recipient>()
        {
          new Recipient
          {
            EmailAddress = new EmailAddress
            {
              Address = "guifrribeiro@outlook.com"
            }
          }
        },
        InternetMessageHeaders = new List<InternetMessageHeader>()
        {
          new InternetMessageHeader
          {
            Name = "x-custom-header-group-name",
            Value = "Nevada"
          },
          new InternetMessageHeader
          {
            Name = "x-custom-header-group-id",
            Value = "NV001"
          }
        }
      };

      var saveToSentItems = true;
      var sendRequest = clientSend.Users["mariadmin@3330sc.onmicrosoft.com"].SendMail(message, saveToSentItems).Request();
      
      Console.WriteLine("Sending...");
      var response = sendRequest.PostResponseAsync().GetAwaiter().GetResult();
      if (response.StatusCode == System.Net.HttpStatusCode.Accepted)
      {
        Console.WriteLine("Mensagem enviada com sucesso");
      }
      else
      {
        Console.WriteLine("Não foi possível enviar a mensagem");
      }
      
      Console.WriteLine("\nMessage Request");
      Console.WriteLine(sendRequest.GetHttpRequestMessage().RequestUri);
    }

    private static GraphServiceClient? _graphClient;
    private static GraphServiceClient graphClientTest;

    private static IConfigurationRoot? LoadAppSettings()
    {
      try
      {
        var config = new ConfigurationBuilder()
                        .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                        .AddJsonFile("appsettings.json", false, true)
                        .Build();

        if (string.IsNullOrEmpty(config["applicationId"]) ||
            string.IsNullOrEmpty(config["applicationSecret"]) ||
            string.IsNullOrEmpty(config["redirectUri"]) ||
            string.IsNullOrEmpty(config["tenantId"]))
        {
          return null;
        }

        return config;
      }
      catch (System.IO.FileNotFoundException)
      {
        return null;
      }
    }

    private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
    {
      var clientId = config["applicationId"];
      var clientSecret = config["applicationSecret"];
      var redirectUri = config["redirectUri"];
      var tenantId = config["tenantId"];
      var authority = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token&grant_type=client_credentials&resource=https://graph.microsoft.com";

      List<string> scopes = new List<string>();
      scopes.Add("https://graph.microsoft.com/.default");

      var cca = ConfidentialClientApplicationBuilder.Create(clientId)
                                                  .WithAuthority(authority)
                                                  .WithRedirectUri(redirectUri)
                                                  .WithClientSecret(clientSecret)
                                                  .Build();

      return new MsalAuthenticationProvider(cca, scopes.ToArray());
    }

    private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
    {
      var authenticationProvider = CreateAuthorizationProvider(config);
      _graphClient = new GraphServiceClient(authenticationProvider);

      return _graphClient;
    }

    private static GraphServiceClient GetAuthenticatedGraphClientTest(IConfigurationRoot config)
    {
      var scopes = new[] { "https://graph.microsoft.com/.default" };
      var tenantId = config["tenantId"];
      var clientId = config["applicationId"];
      var clientSecret = config["applicationSecret"];

      var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);
      graphClientTest = new GraphServiceClient(clientSecretCredential, scopes);

      var tokenRequestContext = new TokenRequestContext(scopes);
      var token = clientSecretCredential.GetTokenAsync(tokenRequestContext).Result.Token;
      return graphClientTest;
    }
  }
}