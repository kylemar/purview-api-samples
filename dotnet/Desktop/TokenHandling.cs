// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Identity.Client;
using System.Windows;
using System.Windows.Interop;

namespace PurviewAPIExp
{
    public partial class MainWindow : Window
    {
        private async Task<string?> GetToken(TokenType type, string[] scopes, string? claimsChallenge = null, bool silent = false, bool forceRefresh = false)
        {
            IAccount? firstAccount;
            IntPtr myWindow;

            if (silent == false)
            {
                myWindow = new WindowInteropHelper(this).Handle;
            }
            else
            {
                myWindow = new IntPtr(0);
            }

            if (null == MSALPublicClientApp)
            {
                return null;
            }
            var accounts = await MSALPublicClientApp.GetAccountsAsync();
            if (accounts.Any())
            {
                firstAccount = accounts.FirstOrDefault();
            }
            else
            {
                if (useBroker == true)
                {
                    firstAccount = PublicClientApplication.OperatingSystemAccount;
                }
                else
                { 
                    firstAccount = accounts.FirstOrDefault();
                }
            }

            AuthenticationResult? authResult;
            try
            {
                authResult = await MSALPublicClientApp.AcquireTokenSilent(scopes, firstAccount)
                    .WithClaims(claimsChallenge)
                    .WithForceRefresh(forceRefresh)
                    .ExecuteAsync()
                    .ConfigureAwait(false);
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilent. 
                // This indicates you need to call AcquireTokenInteractive to acquire a token

                if (silent == true)
                {
                    return null;
                }
                else
                {
                    firstAccount = null;

                    string[] interactiveScopes;
                    string[] extraScopes;
                    if (type == TokenType.ID)
                    {
                        interactiveScopes = new string[] { "openid" };
                        extraScopes = scopes;
                    }
                    else
                    {
                        interactiveScopes = scopes;
                        extraScopes = new string[] { };
                    }

                    try
                    {
                        authResult = await MSALPublicClientApp.AcquireTokenInteractive(scopes)
                        .WithClaims(claimsChallenge ?? ex.Claims)
                        .WithParentActivityOrWindow(myWindow)
                        .WithAccount(firstAccount)
                        .WithExtraScopesToConsent(extraScopes)
                        .ExecuteAsync()
                        .ConfigureAwait(false);
                    }
                    catch (MsalException msalex)
                    {
                        Console.WriteLine($"Error Acquiring Token Interactivly: {msalex.Message}");
                        authResult = null;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error Acquiring Token Silently: {ex.Message}");
                return null;
            }

            if (null != authResult)
            {
                if (type == TokenType.Access)
                {
                    return authResult.AccessToken;
                }
                else
                {
                    return authResult.IdToken;
                }
            }
            else
            {
                return null;
            }
        }
    }

    internal enum TokenType
    {
        ID,
        Access
    }
}
