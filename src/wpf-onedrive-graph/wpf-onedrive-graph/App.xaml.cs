using Microsoft.Identity.Client;
using System.Linq.Expressions;
using System.Windows;

namespace wpf_onedrive_graph
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>

    // To change from Microsoft public cloud to a national cloud, use another value of AzureCloudInstance
    public partial class App : Application
    {
        static App()
        {
            CreateApplication(true);
        }

        public static void CreateApplication(bool useWam)
        {
            var builder = PublicClientApplicationBuilder.Create(ClientId)
                .WithAuthority($"{Instance}{Tenant}")
                .WithDefaultRedirectUri();

            _clientApp = builder.Build();
        }

        // Below are the clientId (Application Id) of your app registration and the tenant information. 
        // You have to replace:
        // - the content of ClientID with the Application Id for your app registration
        // - The content of Tenant by the information about the accounts allowed to sign-in in your application:
        //   - For Work or School account in your org, use your tenant ID, or domain
        //   - for any Work or School accounts, use organizations
        //   - for any Work or School accounts, or Microsoft personal account, use common
        //   - for Microsoft Personal account, use consumers
        private static string ClientId = ClientInfo.ClientId;

        // Note: Tenant is important for the quickstart.
        private static string Tenant = "common";
        private static string Instance = "https://login.microsoftonline.com/";
        private static IPublicClientApplication _clientApp;

        public static IPublicClientApplication PublicClientApp { get { return _clientApp; } }
    }
}
