// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using active_directory_dotnet_native_uwp_v2;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Identity.Client;
using Microsoft.IdentityModel.Abstractions;
using Microsoft.Kiota.Abstractions.Authentication;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace Native_UWP_V2
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        //Set the scope for API call to user.read
        private string[] scopes = new[] { "User.Read" };
        private string[] DriveScopes = new string[] { "Files.ReadWrite.All" };

        // Below are the clientId (Application Id) of your app registration and the tenant information. 
        // You have to replace:
        // - the content of ClientID with the Application Id for your app registration
        // - The content of Tenant by the information about the accounts allowed to sign-in in your application:
        //   - For Work or School account in your org, use your tenant ID, or domain
        //   - for any Work or School accounts, use organizations
        //   - for any Work or School accounts, or Microsoft personal account, use common
        //   - for Microsoft Personal account, use consumers
        private const string ClientId = "0f8721fb-edac-41c1-a18b-fa82e22d7260";
        private const string ClientSecret = "Jag8Q~kKkZcmHdHXXcncTU1pRTEuG44VfEI3Ocvi";
        private const string TenantId = "2b3575a3-eb8d-49a1-8c90-f22961a3986e"; // Alternatively "[Enter your tenant, as obtained from the azure portal, e.g. kko365.onmicrosoft.com]"
        private const string Authority = "https://login.microsoftonline.com/" + TenantId;
        private static string MSGraphURL = "https://graph.microsoft.com/v1.0/";
        InteractiveBrowserCredentialOptions Options = new InteractiveBrowserCredentialOptions
        {
            
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            TenantId = TenantId,
            ClientId = ClientId,
            RedirectUri = new Uri("http://localhost")
        };

        // The MSAL Public client app
        private static IPublicClientApplication PublicClientApp;
        private IAccount _currentUserAccount;
        private GraphServiceClient graphClient;
        private AuthenticationResult authResult;

        public MainPage()
        {
            this.InitializeComponent();

            // Initialize the MSAL library by building a public client application
            //PublicClientApp = PublicClientApplicationBuilder.Create(ClientId)
            //    .WithAuthority(Authority)
            //    .WithBroker(true)
            //    //this is the currently recommended way to log MSAL message. For more info refer to https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/logging
            //    .WithLogging(new IdentityLogger(EventLogLevel.Warning), enablePiiLogging: false) //set Identity Logging level to Warning which is a middle ground
            //    .Build();

            //_currentUserAccount = Task.Run(async () => await PublicClientApp.GetAccountsAsync()).Result.FirstOrDefault();

            //if (_currentUserAccount != null)
            //{
            //    this.CallGraphButton.Content = "Call Microsoft Graph API";
            //    this.SignOutButton.Visibility = Visibility.Visible;
            //}

        }

        /// <summary>
        /// Call AcquireTokenAsync - to acquire a token requiring user to sign-in
        /// </summary>
        private async void CallGraphButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Sign-in user using MSAL and obtain an access token for MS Graph
                var graphClient = await SignInAndInitializeGraphServiceClient(scopes);
                var a = await graphClient.Me.GetAsync(res => {  res.Headers.Add("Accept", @"text/html, application/json" ); });
                graphClient = await SignInAndInitializeGraphServiceClient(DriveScopes);

                var b = await graphClient.Me.Drive.GetAsync();

            }
            catch (MsalException msalEx)
            {
                await DisplayMessageAsync($"Error Acquiring Token:{System.Environment.NewLine}{msalEx}");
            }
            catch (ODataError ex)
            {
                await DisplayMessageAsync($"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}");
                return;
            }
        }

        /// <summary>
        /// Signs in the user and obtains an Access token for MS Graph
        /// </summary>
        /// <param name="scopes"></param>
        /// <returns> Access Token</returns>
        //
        private async Task<string> SignInUserAndGetTokenUsingMSAL(string[] scopes)
        {
            // returns smth like S-1-15-2-2601115387-131721061-1180486061-1362788748-631273777-3164314714-2766189824
            string sid = Windows.Security.Authentication.Web.WebAuthenticationBroker.GetCurrentApplicationCallbackUri().Host.ToUpper();

            // This is redirect uri you need to register in the app registration portal. The app config does not need it.
            string redirectUri = $"ms-appx-web://microsoft.aad.brokerplugin/{sid}";

            // Initialize the MSAL library by building a public client application
            PublicClientApp = PublicClientApplicationBuilder.Create(ClientId)
                .WithAuthority(Authority)
                .WithUseCorporateNetwork(false)
                .WithRedirectUri("https://login.live.com/oauth20_desktop.srf")
                //this is the currently recommended way to log MSAL message. For more info refer to https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/logging
                .WithLogging(
                (level, message, containsPii) =>
                {
                    Debug.WriteLine($"MSAL: {level} {message} ");
                },
                LogLevel.Warning, enablePiiLogging: false, enableDefaultPlatformLogging: true) //set Identity Logging level to Warning which is a middle ground
                .Build();

            _currentUserAccount = _currentUserAccount ?? (await PublicClientApp.GetAccountsAsync()).FirstOrDefault();

            try
            {
                authResult = await PublicClientApp.AcquireTokenSilent(scopes, _currentUserAccount)
                                                  .ExecuteAsync();

                await Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, () =>
                {
                    this.CallGraphButton.Content = "Call Microsoft Graph API";
                });

            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilentAsync. This indicates you need to call AcquireTokenAsync to acquire a token
                Debug.WriteLine($"MsalUiRequiredException: {ex.Message}");

                // Must be called from UI thread
                authResult = await PublicClientApp.AcquireTokenInteractive(scopes)
                                                  .ExecuteAsync()
                                                  .ConfigureAwait(false);

            }

            return authResult.AccessToken;
        }

        /// <summary>
        /// Sign in user using MSAL and obtain a token for MS Graph
        /// </summary>
        /// <returns>GraphServiceClient</returns>
        private async Task<GraphServiceClient> SignInAndInitializeGraphServiceClient(string[] scopes)
        {
            //GraphServiceClient graphClient = new GraphServiceClient(
            //    new DelegateAuthenticationProvider(async (requestMessage) =>
            //    {
            //        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", await SignInUserAndGetTokenUsingMSAL(scopes));
            //    }), MSGraphURL);
            var tokenProvider = new TokenProvider(SignInUserAndGetTokenUsingMSAL, scopes);
            return new GraphServiceClient(
                new BaseBearerTokenAuthenticationProvider(tokenProvider), MSGraphURL);
            // Callback function that receives the user prompt
            // Prompt contains the generated device code that you must
            // enter during the auth process in the browser


            // This is the incoming token to exchange using on-behalf-of flow
            //var oboToken = await SignInUserAndGetTokenUsingMSAL(scopes);

            //var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider(ClientId, TenantId));
            //return new GraphServiceClient(authenticationProvider);



        }

        /// <summary>
        /// Sign out the current user
        /// </summary>
        private async void SignOutButton_Click(object sender, RoutedEventArgs e)
        {
            IEnumerable<IAccount> accounts = await PublicClientApp.GetAccountsAsync().ConfigureAwait(false);
            IAccount firstAccount = accounts.FirstOrDefault();

            try
            {
                await PublicClientApp.RemoveAsync(firstAccount).ConfigureAwait(false);
                await Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, () =>
                {
                    ResultText.Text = "User has signed-out";
                    TokenInfoText.Text = string.Empty;
                    this.CallGraphButton.Visibility = Visibility.Visible;
                    this.SignOutButton.Visibility = Visibility.Collapsed;
                    this.CallGraphButton.Content = "Sign-In and Call Microsoft Graph API";
                });
            }
            catch (MsalException ex)
            {
                ResultText.Text = $"Error signing-out user: {ex.Message}";
            }
        }

        /// <summary>
        /// Display basic information contained in the token. Needs to be called from the UI thead.
        /// </summary>
        private void DisplayBasicTokenInfo(AuthenticationResult authResult)
        {
            TokenInfoText.Text = "";
            if (authResult != null)
            {
                TokenInfoText.Text += $"User Name: {authResult.Account.Username}" + Environment.NewLine;
                TokenInfoText.Text += $"Token Expires: {authResult.ExpiresOn.ToLocalTime()}" + Environment.NewLine;
                TokenInfoText.Text += $"Token : {authResult.AccessToken}" + Environment.NewLine;
            }
        }

        /// <summary>
        /// Displays a message in the ResultText. Can be called from any thread.
        /// </summary>
        private async Task DisplayMessageAsync(string message)
        {
            await Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal,
                   () =>
                   {
                       ResultText.Text = message;
                   });
        }
    }
}