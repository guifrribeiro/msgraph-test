using System.Net.Http.Headers;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace Helpers
{
  public class MsalAuthenticationProvider : IAuthenticationProvider
  {
    private IConfidentialClientApplication _clientApplication;
    private string[] _scopes;

    public MsalAuthenticationProvider(IConfidentialClientApplication clientApplication, string[] scopes)
    {
      _clientApplication = clientApplication;
      _scopes = scopes;
    }

    public async Task AuthenticateRequestAsync(HttpRequestMessage request)
    {
      var token = await GetTokenAsync();
      request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
    }

    private async Task<string> GetTokenAsync()
    {
      AuthenticationResult authResult;
      authResult = await _clientApplication.AcquireTokenForClient(_scopes).ExecuteAsync();

      return authResult.AccessToken;
    }
  }
}