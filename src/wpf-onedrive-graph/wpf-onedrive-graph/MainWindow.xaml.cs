using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;

namespace wpf_onedrive_graph
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>

    public partial class MainWindow : Window
    {
        //Set the API Endpoint to Graph 'me' endpoint. 
        // To change from Microsoft public cloud to a national cloud, use another value of graphAPIEndpoint.
        // Reference with Graph endpoints here: https://docs.microsoft.com/graph/deployments#microsoft-graph-and-graph-explorer-service-root-endpoints
        string graphAPIEndpoint = "https://graph.microsoft.com/v1.0/me";

        //Set the scope for API calls
        string[] scopes = new string[] { "user.read", "Files.Read", "Files.Read.All", "Files.ReadWrite", "Files.ReadWrite.All" };

        GraphServiceClient _graphClient;

        public MainWindow()
        {
            InitializeComponent();
        }

        private async void DownloadButton_Click(object sender, RoutedEventArgs e)
        {
            ResultText.Text = "";
            if (_graphClient == null)
            {
                ResultText.Text = "Please sign-in first";
                return;
            }

            try
            {
                var driveResult = await _graphClient.Me.Drive.Request().GetAsync();
                var itemResult = await _graphClient.Me.Drive.Root.ItemWithPath(FilePathDownload.Text).Request().GetAsync();
                ResultText.Text = "Downloading...";
                var stream = await _graphClient.Me.Drive.Items[itemResult.Id].Content.Request().GetAsync();
                using (var memoryStream = new System.IO.MemoryStream())
                {
                    stream.CopyTo(memoryStream);
                    System.IO.File.WriteAllBytes(LocalFilePath.Text + "\\" + itemResult.Name, memoryStream.ToArray());
                }
                ResultText.Text = "Download Complete";
            }
            catch (Microsoft.Graph.ServiceException ex)
            {
                ResultText.Text = ex.Error.ToString();
            }
            catch (MsalUiRequiredException ex)
            {
                ResultText.Text = "Please sign-in first";
            }
            catch (Exception ex)
            {
                ResultText.Text = ex.Message;
            }
        }

        private async void UploadButton_Click(object sender, RoutedEventArgs e)
        {
            ResultText.Text = "";
            if (_graphClient == null)
            {
                ResultText.Text = "Please sign-in first";
                return;
            }
            try
            {
                string fileName = Path.GetFileName(FilePathUpload.Text);

                using (var fileStream = System.IO.File.OpenRead(FilePathUpload.Text))
                {
                    // Use properties to specify the conflict behavior
                    // in this case, replace
                    var uploadProps = new DriveItemUploadableProperties
                    {
                        ODataType = null,
                        AdditionalData = new Dictionary<string, object>
                        {
                            { "@microsoft.graph.conflictBehavior", "replace" }
                        }
                    };

                    // Create the upload session
                    // itemPath does not need to be a path to an existing item
                    var uploadSession = await _graphClient.Me.Drive.Root
                        .ItemWithPath(OneDrivePath.Text + "/" + fileName)
                        .CreateUploadSession(uploadProps)
                        .Request()
                        .PostAsync();

                    // Max slice size must be a multiple of 320 KiB
                    int maxSliceSize = 320 * 1024;
                    var fileUploadTask =
                        new LargeFileUploadTask<DriveItem>(uploadSession, fileStream, maxSliceSize);

                    // Create a callback that is invoked after each slice is uploaded
                    IProgress<long> progress = new Progress<long>(prog =>
                    {
                        ResultText.Text = $"Uploaded {prog} bytes of {fileStream.Length} bytes";
                    });

                    try
                    {
                        // Upload the file
                        var uploadResult = await fileUploadTask.UploadAsync(progress);

                        if (uploadResult.UploadSucceeded)
                        {
                            // The ItemResponse object in the result represents the
                            // created item.
                            ResultText.Text = $"Upload complete, item ID: {uploadResult.ItemResponse.Id}";
                        }
                        else
                        {
                            ResultText.Text = "Upload failed";
                        }
                    }
                    catch (ServiceException ex)
                    {
                        ResultText.Text = $"Error uploading: {ex.ToString()}";
                    }
                }
            }
            catch (Microsoft.Graph.ServiceException ex)
            {
                ResultText.Text = ex.Error.ToString();
            }
            catch (MsalUiRequiredException ex)
            {
                ResultText.Text = "Please sign-in first";
            }
            catch (Exception ex)
            {
                ResultText.Text = ex.Message;
            }
        }

        private async void Login_Click(object sender, RoutedEventArgs e)
        {
            AuthenticationResult authResult = null;
            var app = App.PublicClientApp;
            IAccount firstAccount;

            var accounts = await app.GetAccountsAsync();
            firstAccount = accounts.FirstOrDefault();
            InteractiveAuthenticationProvider i = new InteractiveAuthenticationProvider(app, scopes);

            try
            {
                authResult = await app.AcquireTokenSilent(scopes, firstAccount)
                    .ExecuteAsync();
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilent. 
                // This indicates you need to call AcquireTokenInteractive to acquire a token
                System.Diagnostics.Debug.WriteLine($"MsalUiRequiredException: {ex.Message}");

                try
                {
                    authResult = await app.AcquireTokenInteractive(scopes)
                        .WithAccount(firstAccount)
                        .WithParentActivityOrWindow(new WindowInteropHelper(this).Handle) // optional, used to center the browser on the window
                        .WithPrompt(Microsoft.Identity.Client.Prompt.SelectAccount)
                        .ExecuteAsync();
                }
                catch (MsalException msalex)
                {
                    ResultText.Text = $"Error Acquiring Token:{System.Environment.NewLine}{msalex}";
                }
            }
            catch (Exception ex)
            {
                ResultText.Text = $"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}";
                return;
            }

            if (authResult != null)
            {
                _graphClient = new GraphServiceClient(i);
                ResultText.Text = await GetHttpContentWithToken(graphAPIEndpoint, authResult.AccessToken);
                DisplayBasicTokenInfo(authResult);
                this.SignOutButton.Visibility = Visibility.Visible;
                this.SignInButton.Visibility = Visibility.Collapsed;
            }
        }

        /// <summary>
        /// Perform an HTTP GET request to a URL using an HTTP Authorization header
        /// </summary>
        /// <param name="url">The URL</param>
        /// <param name="token">The token</param>
        /// <returns>String containing the results of the GET operation</returns>
        public async Task<string> GetHttpContentWithToken(string url, string token)
        {
            var httpClient = new System.Net.Http.HttpClient();
            System.Net.Http.HttpResponseMessage response;
            try
            {
                var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url);
                //Add the token in Authorization header
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                response = await httpClient.SendAsync(request);
                var content = await response.Content.ReadAsStringAsync();
                return content;
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

        /// <summary>
        /// Sign out the current user
        /// </summary>
        private async void SignOutButton_Click(object sender, RoutedEventArgs e)
        {
            await SignOut();
        }

        private async void Window_Closing(object sender, CancelEventArgs e)
        {
            await SignOut();
        }

        private async Task SignOut()
        {
            var accounts = await App.PublicClientApp.GetAccountsAsync();
            if (accounts.Any())
            {
                try
                {
                    await App.PublicClientApp.RemoveAsync(accounts.FirstOrDefault());
                    _graphClient = null;
                    this.ResultText.Text = "User has signed-out";
                    this.SignOutButton.Visibility = Visibility.Collapsed;
                    this.SignInButton.Visibility = Visibility.Visible;
                }
                catch (MsalException ex)
                {
                    ResultText.Text = $"Error signing-out user: {ex.Message}";
                }
            }
        }

        /// <summary>
        /// Display basic information contained in the token
        /// </summary>
        private void DisplayBasicTokenInfo(AuthenticationResult authResult)
        {
            TokenInfoText.Text = "";
            if (authResult != null)
            {
                TokenInfoText.Text += $"Username: {authResult.Account.Username}" + Environment.NewLine;
                TokenInfoText.Text += $"Token Expires: {authResult.ExpiresOn.ToLocalTime()}" + Environment.NewLine;
            }
        }
    }
}
