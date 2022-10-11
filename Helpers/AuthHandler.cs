using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Graph;
using System.Threading;

namespace Helpers
{
  public class AuthHandler : DelegatingHandler
  {
    private IAuthenticationProvider authenticationProvider;

    public AuthHandler(IAuthenticationProvider authProvider, HttpMessageHandler innerHandler)
    {
      InnerHandler = innerHandler;
      authenticationProvider = authProvider;
    }

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
      await authenticationProvider.AuthenticateRequestAsync(request);
      return await base.SendAsync(request, cancellationToken);
    }
  }
}