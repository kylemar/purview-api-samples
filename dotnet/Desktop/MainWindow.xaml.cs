// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace PurviewAPIExp
{
    using Microsoft.Identity.Client;
    using Microsoft.Identity.Client.Broker;
    using Newtonsoft.Json;
    using System.Globalization;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Windows;
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// API URL formats
        /// </summary>
        private const string UserPsUrlFormat = "{0}/me/dataSecurityAndGovernance/protectionScopes/compute";
        private const string PcUrlFormat = "{0}/me/dataSecurityAndGovernance/processContent";
        private const string ActivityUploadFormat = "{0}/me/dataSecurityAndGovernance/activities/contentActivities";

        /// <summary>
        /// Cache the ProtectionScopes ID
        /// </summary>
        private string scopeIdentifier = string.Empty;

        /// <summary>
        /// going to need some random numbers for the data 
        /// </summary>
        private static Random random = new Random();
        private int applicationNumber = random.Next(0, Int32.MaxValue);

        /// <summary>
        /// Do we need to evaluate inline or offline?
        /// 
        /// 
        /// </summary>
        private bool isOffline = true;

        /// <summary
        /// Simple logger
        /// </summary>         
        private static readonly StringBuilder sbLog = new();
        private readonly Logger logger = new(sbLog);

        /// <summary>
        /// Are the elements in this window initializaed
        /// </summary>
        private bool isInitialized = false;

        /// <summary>
        /// HTTP Operation, GET or POST
        /// </summary>
        private string httpOperation = string.Empty;

        /// <summary>
        /// Entra ID values
        /// </summary>
        private string clientId = "83ef198a-0396-4893-9d4f-d36efbffc8bd";
        private string userId = string.Empty;
        private string tenantId = string.Empty;

        /// <summary>
        /// Windows OS values
        /// </summary>
        private string WindowsVersion = Environment.OSVersion.Version.ToString();
        private string WindowsMarketingName = string.Empty;

        /// <summary>
        /// current display url values
        /// </summary>
        private string userPsUrl = string.Empty;
        private string pcUrl = string.Empty;
        private string activityUrl = string.Empty;
        private string baseUrl = "https://graph.microsoft.com/beta";

        /// <summary>
        /// Microsoft Authentication Library Public Client
        /// Used to sign in the user and acquire access tokens to call Purview API
        /// </summary>
        private static IPublicClientApplication? MSALPublicClientApp;

        /// <summary>
        /// All the permissions the app needs. Listed together for a single consent prompt on first sign in
        /// as opposed incrementally adding consent with multiple prompts as the user explores APIs
        /// </summary>
        private string[] signInScopes = new string[] { "user.read", "Content.Process.User", "ProtectionScopes.Compute.User", "ContentActivity.Write", "SensitivityLabel.Read" };

        /// <summary>
        /// Initializes a new instance of the <see cref="MainWindow"/> class.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            this.isInitialized = true;

            // Updated selected drop boxes
            ApiSelectBox.SelectedIndex = 0;

            // Get Windows Marketing Name
            int build = Environment.OSVersion.Version.Build;
            if (build >= 22000)
            {
                WindowsMarketingName = "Windows 11";
            }
            else
            {
                WindowsMarketingName = "Windows 10";
            }
        }

        private async Task SignInUser()
        {
            PublicClientApplicationBuilder builder = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(AadAuthorityAudience.AzureAdMultipleOrgs)
                // Use Continous Access Evaluation (CAE)
                .WithClientCapabilities(new[] { "cp1" });

            if (UseBroker.IsChecked == true)
            {
                // This sample uses the auth broker by default. 
                //
                // https://devblogs.microsoft.com/identity/improved-windows-broker-support-with-msal-net/
                // https://learn.microsoft.com/en-us/entra/msal/dotnet/acquiring-tokens/desktop-mobile/wam
                //
                BrokerOptions options = new BrokerOptions(BrokerOptions.OperatingSystems.Windows)
                {
                    Title = "Purview API Explorer"
                };
                builder.WithBroker(options);
                builder.WithDefaultRedirectUri();
            }
            else
            {
                builder.WithRedirectUri("http://localhost");
            }

            MSALPublicClientApp = builder.Build();
            TokenCacheHelper.EnableSerialization(MSALPublicClientApp.UserTokenCache);

            try
            {
                string? IDToken = await GetToken(TokenType.ID, signInScopes);
                IEnumerable<IAccount> accounts = await MSALPublicClientApp.GetAccountsAsync();
                IAccount? account = accounts.FirstOrDefault();
                if (account != null) 
                {
                    userId = account.HomeAccountId.ObjectId;
                    tenantId = account.HomeAccountId.TenantId;
                    this.SetUpApiUrls();
                }
            }
            catch (Exception ex)
            {
                StatusBox.Text = "Error signing in user";
                MessageBox.Show($"Error: {ex.Message}");
                logger.Log("Error Acquiring Token: " + ex.Message);
                Environment.Exit(1);
            }
        }

        private async Task<Dictionary<string, string>> GetResponseStringAsync(HttpResponseMessage response)
        {
            string prettyJson = "";
            ComboBoxItem selectedItem = (ComboBoxItem)ApiSelectBox.SelectedItem;
            string? API = selectedItem.Content.ToString();

            Dictionary<string, string> responseDict = new Dictionary<string, string>(3, StringComparer.OrdinalIgnoreCase);

            string responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(true);

            if (responseContent != null)
            {
                try
                {
                    dynamic? parsedJson = JsonConvert.DeserializeObject<dynamic>(responseContent);
                    if (parsedJson != null) 
                    {
                        prettyJson = JsonConvert.SerializeObject(parsedJson, Newtonsoft.Json.Formatting.Indented);
                    }
                }
                catch (JsonReaderException ex)
                {
                    prettyJson = responseContent;
                    logger.Log("Error parsing JSON: " + ex.Message);
                }
            }

            if (selectedItem != null && selectedItem.Content != null)
            {
                if (API != null && API.StartsWith("Process Content") )
                {
                    if (isOffline)
                    {
                        responseDict["StatusCode"] = $"StatusCode: {(int)response.StatusCode} - {response.ReasonPhrase} for offline processing.";
                    }
                    else
                    {
                        responseDict["StatusCode"] = $"StatusCode: {(int)response.StatusCode} - {response.ReasonPhrase}";
                    }
                }
                else if (API != null && API.StartsWith("Protection Scopes"))
                {
                    isOffline = true;
                    if (responseContent != null && responseContent.Contains("evaluateInline"))
                    {
                        isOffline = false;
                    }
                    responseDict["StatusCode"] = $"StatusCode: {(int)response.StatusCode} - {response.ReasonPhrase}";
                }
                else
                {
                    responseDict["StatusCode"] = $"StatusCode: {(int)response.StatusCode} - {response.ReasonPhrase}";
                }
            }

            if (response.Headers.ETag != null)
            {
                scopeIdentifier = response.Headers.ETag.ToString();
            }
            responseDict["Headers"] = $"{response.Headers}";
            responseDict["Content"] = $"{prettyJson}";

            return responseDict;
        }

        private void ApiSelectBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ( userId == string.Empty)
            {
                return;
            }
            this.SetUpApiViews();
        }

        private void SetUpApiUrls()
        {
            if (userId == string.Empty)
            {
                return;
            }
            this.userPsUrl = string.Format(CultureInfo.InvariantCulture, UserPsUrlFormat, this.baseUrl);
            this.pcUrl = string.Format(CultureInfo.InvariantCulture, PcUrlFormat, this.baseUrl);
            this.activityUrl = string.Format(CultureInfo.InvariantCulture, ActivityUploadFormat, this.baseUrl);
            this.SetUpApiViews();
        }

        /// <summary>
        /// Conversation attributes for Process Content API
        /// Allows for a continued conversation
        /// </summary>
        private string conversationId = string.Empty;
        private int conversationSequence = 0;

        private void SetUpApiViews()
        {
            if (userId == string.Empty)
            {
                UrlTextBox.Text = string.Empty;
                Scope.Text = string.Empty;
                return;
            }

            if (this.isInitialized && ApiSelectBox.SelectedItem is ComboBoxItem selectedItem)
            {
                ResponseTextBox.Text = string.Empty;

                StringBuilder headers = new StringBuilder();
                headers.AppendLine("User-Agent:Purview API Explorer");
                headers.AppendLine($"client-request-id:{Guid.NewGuid().ToString()}");
                headers.AppendLine($"x-ms-client-request-id:{Guid.NewGuid().ToString()}");
                if (scopeIdentifier != null && scopeIdentifier != string.Empty)
                {
                    headers.AppendLine($"If-None-Match:{scopeIdentifier}");
                }
                RequestHeadersTextBox.Text = headers.ToString();

                switch (selectedItem.Content.ToString())
                {
                    case "Process Content - Start Conversation":
                        httpOperation = "POST";
                        UrlTextBox.Text = this.pcUrl;
                        Scope.Text = "Content.Process.User";
                        RequestContentTabControl.SelectedIndex = 0;
                        conversationId = Guid.NewGuid().ToString();
                        conversationSequence = 0;
                        RequestBodyTextBox.Text =
                        $"{{\r\n" +
                        $"    \"contentToProcess\": {{\r\n" +
                        $"       \"contentEntries\": [\r\n" +
                        $"          {{\r\n" +
                        $"             \"@odata.type\": \"microsoft.graph.processConversationMetadata\",\r\n" +
                        $"             \"identifier\": \"{Guid.NewGuid()}\",\r\n" +
                        $"             \"content\": {{\r\n" +
                        $"                \"@odata.type\": \"microsoft.graph.textContent\", \r\n" +
                        $"                \"data\": \"For application {++applicationNumber}, Write an acceptance letter for Alex Wilber with Credit card number 4532667785213500, ssn: 120-98-1437 at One Microsoft Way, Redmond, WA 98052\"\r\n" +
                        $"             }},\r\n" +
                        $"             \"name\":\"PC Purview API Explorer message\",\r\n" +
                        $"             \"correlationId\": \"{conversationId}\",\r\n" +
                        $"             \"sequenceNumber\": {conversationSequence++}, \r\n" +
                        $"             \"isTruncated\": false,\r\n" +
                        $"             \"createdDateTime\": \"{DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss")}\",\r\n" +
                        $"             \"modifiedDateTime\": \"{DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss")}\"\r\n" +
                        $"          }}\r\n" +
                        $"       ],\r\n" +
                        $"       \"activityMetadata\": {{ \r\n" +
                        $"          \"activity\": \"uploadText\"\r\n" +
                        $"       }},\r\n" +
                        $"       \"deviceMetadata\": {{\r\n" +
                        $"          \"deviceType\": \"managed\",\r\n" +
                        $"          \"operatingSystemSpecifications\": {{\r\n" +
                        $"             \"operatingSystemPlatform\": \"{WindowsMarketingName}\",\r\n" +
                        $"             \"operatingSystemVersion\": \"{WindowsVersion}\" \r\n" +
                        $"          }}\r\n" +
                        $"       }},\r\n" +
                        $"       \"protectedAppMetadata\": {{\r\n" +
                        $"          \"name\": \"PC Purview API Explorer\",\r\n" + 
                        $"          \"version\": \"0.2\",\r\n" + 
                        $"          \"applicationLocation\":{{\r\n" + 
                        $"             \"@odata.type\": \"microsoft.graph.policyLocationApplication\",\r\n" + 
                        $"             \"value\": \"{clientId}\"\r\n" + 
                        $"          }}\r\n" +
                        $"       }},\r\n" + 
                        $"       \"integratedAppMetadata\": {{\r\n" +
                        $"          \"name\": \"PC Purview API Explorer\",\r\n" +
                        $"          \"version\": \"0.1\" \r\n" +
                        $"       }}\r\n" +
                        $"    }}\r\n" +
                        $"}}";
                        break;

                    case "Process Content - Continue Conversation with Response":
                        httpOperation = "POST";
                        UrlTextBox.Text = this.pcUrl;
                        Scope.Text = "Content.Process.User";
                        RequestContentTabControl.SelectedIndex = 0;
                        RequestBodyTextBox.Text =
                        $"{{\r\n" +
                        $"    \"contentToProcess\": {{\r\n" +
                        $"       \"contentEntries\": [\r\n" +
                        $"          {{\r\n" +
                        $"             \"@odata.type\": \"microsoft.graph.processConversationMetadata\",\r\n" +
                        $"             \"identifier\": \"{Guid.NewGuid()}\",\r\n" +
                        $"             \"content\": {{\r\n" +
                        $"                \"@odata.type\": \"microsoft.graph.textContent\", \r\n" +
                        $"                \"data\": \"Dear Alex, your application {applicationNumber} is accepted. Your payments will be automatically deducted from your credit card 4532667785213500 and statements mailed to Alex Wilber One Microsoft Way, Redmond, WA 98052. For tax purposes this transaction will be reported with your Social Security number -  120-98-1437\"\r\n" +
                        $"             }},\r\n" +
                        $"             \"name\":\"PC Purview API Explorer message\",\r\n" +
                        $"             \"correlationId\": \"{conversationId}\",\r\n" +
                        $"             \"sequenceNumber\": {conversationSequence++}, \r\n" +
                        $"             \"isTruncated\": false,\r\n" +
                        $"             \"createdDateTime\": \"{DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss")}\",\r\n" +
                        $"             \"modifiedDateTime\": \"{DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss")}\"\r\n" +
                        $"          }}\r\n" +
                        $"       ],\r\n" +
                        $"       \"activityMetadata\": {{ \r\n" +
                        $"          \"activity\": \"downloadText\"\r\n" +
                        $"       }},\r\n" +
                        $"       \"deviceMetadata\": {{\r\n" +
                        $"          \"deviceType\": \"managed\",\r\n" +
                        $"          \"operatingSystemSpecifications\": {{\r\n" +
                        $"             \"operatingSystemPlatform\": \"{WindowsMarketingName}\",\r\n" +
                        $"             \"operatingSystemVersion\": \"{WindowsVersion}\" \r\n" +
                        $"          }}\r\n" +
                        $"       }},\r\n" +
                        $"       \"protectedAppMetadata\": {{\r\n" +
                        $"          \"name\": \"PC Purview API Explorer\",\r\n" +
                        $"          \"version\": \"0.2\",\r\n" +
                        $"          \"applicationLocation\":{{\r\n" +
                        $"             \"@odata.type\": \"microsoft.graph.policyLocationApplication\",\r\n" +
                        $"             \"value\": \"{clientId}\"\r\n" +
                        $"          }}\r\n" +
                        $"       }},\r\n" +
                        $"       \"integratedAppMetadata\": {{\r\n" +
                        $"          \"name\": \"PC Purview API Explorer\",\r\n" +
                        $"          \"version\": \"0.1\" \r\n" +
                        $"       }}\r\n" +
                        $"    }}\r\n" +
                        $"}}"; break;

                    case "Process Content - Continue Conversation with Prompt":
                        httpOperation = "POST";
                        UrlTextBox.Text = this.pcUrl;
                        Scope.Text = "Content.Process.User";
                        RequestContentTabControl.SelectedIndex = 0;
                        RequestBodyTextBox.Text =
                        $"{{\r\n" +
                        $"    \"contentToProcess\": {{\r\n" +
                        $"       \"contentEntries\": [\r\n" +
                        $"          {{\r\n" +
                        $"             \"@odata.type\": \"microsoft.graph.processConversationMetadata\",\r\n" +
                        $"             \"identifier\": \"{Guid.NewGuid()}\",\r\n" +
                        $"             \"content\": {{\r\n" +
                        $"                \"@odata.type\": \"microsoft.graph.textContent\", \r\n" +
                        $"                \"data\": \"How many applications have been accepted today?\"\r\n" +
                        $"             }},\r\n" +
                        $"             \"name\":\"PC Purview API Explorer message\",\r\n" +
                        $"             \"correlationId\": \"{conversationId}\",\r\n" +
                        $"             \"sequenceNumber\": {conversationSequence++}, \r\n" +
                        $"             \"isTruncated\": false,\r\n" +
                        $"             \"createdDateTime\": \"{DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss")}\",\r\n" +
                        $"             \"modifiedDateTime\": \"{DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss")}\"\r\n" +
                        $"          }}\r\n" +
                        $"       ],\r\n" +
                        $"       \"activityMetadata\": {{ \r\n" +
                        $"          \"activity\": \"uploadText\"\r\n" +
                        $"       }},\r\n" +
                        $"       \"deviceMetadata\": {{\r\n" +
                        $"          \"deviceType\": \"managed\",\r\n" +
                        $"          \"operatingSystemSpecifications\": {{\r\n" +
                        $"             \"operatingSystemPlatform\": \"{WindowsMarketingName}\",\r\n" +
                        $"             \"operatingSystemVersion\": \"{WindowsVersion}\" \r\n" +
                        $"          }}\r\n" +
                        $"       }},\r\n" +
                        $"       \"protectedAppMetadata\": {{\r\n" +
                        $"          \"name\": \"PC Purview API Explorer\",\r\n" +
                        $"          \"version\": \"0.2\",\r\n" +
                        $"          \"applicationLocation\":{{\r\n" +
                        $"             \"@odata.type\": \"microsoft.graph.policyLocationApplication\",\r\n" +
                        $"             \"value\": \"{clientId}\"\r\n" +
                        $"          }}\r\n" +
                        $"       }},\r\n" +
                        $"       \"integratedAppMetadata\": {{\r\n" +
                        $"          \"name\": \"PC Purview API Explorer\",\r\n" +
                        $"          \"version\": \"0.1\" \r\n" +
                        $"       }}\r\n" +
                        $"    }}\r\n" +
                        $"}}";
                        break;

                    case "Protection Scopes - Initial Call":
                        httpOperation = "POST";
                        UrlTextBox.Text = this.userPsUrl;
                        Scope.Text = "https://canary.graph.microsoft.com/ProtectionScopes.Compute.User";
                        Scope.Text = "ProtectionScopes.Compute.User";
                        RequestContentTabControl.SelectedIndex = 0;
                        RequestBodyTextBox.Text =
                        "{\r\n" +
                        "   \"activities\": \"uploadText,downloadText\",\r\n" +
                        "   \"locations\": [\r\n" +
                        "      {\r\n" +
                        "         \"@odata.type\": \"microsoft.graph.policyLocationApplication\",\r\n" +
                        $"         \"value\": \"{clientId}\"\r\n" +
                        "      }\r\n" +
                        "   ]\r\n" +
                        "}\r\n";
                        break;

                    case "Content Activity":
                        httpOperation = "POST";
                        UrlTextBox.Text = this.activityUrl;
                        Scope.Text = "ContentActivity.Write";
                        conversationId = Guid.NewGuid().ToString();
                        conversationSequence = 0;
                        RequestContentTabControl.SelectedIndex = 0;
                        RequestBodyTextBox.Text =
                        $"{{\r\n" +
                        $"    \"contentMetadata\": {{\r\n" +
                        $"       \"contentEntries\": [\r\n" +
                        $"          {{\r\n" +
                        $"             \"@odata.type\": \"microsoft.graph.processConversationMetadata\",\r\n" +
                        $"             \"identifier\": \"{Guid.NewGuid()}\",\r\n" +
                        $"             \"name\":\"CA Purview API Explorer message\",\r\n" +
                        $"             \"correlationId\": \"{conversationId}\",\r\n" +
                        $"             \"sequenceNumber\": {conversationSequence++}, \r\n" +
                        $"             \"isTruncated\": false,\r\n" +
                        $"             \"createdDateTime\": \"{DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss")}\",\r\n" +
                        $"             \"modifiedDateTime\": \"{DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss")}\"\r\n" +
                        $"          }}\r\n" +
                        $"       ],\r\n" +
                        $"       \"activityMetadata\": {{ \r\n" +
                        $"          \"activity\": \"downloadText\"\r\n" +
                        $"       }},\r\n" +
                        $"       \"deviceMetadata\": {{\r\n" +
                        $"          \"deviceType\": \"unmanaged\",\r\n" +
                        $"          \"operatingSystemSpecifications\": {{\r\n" +
                        $"             \"operatingSystemPlatform\": \"{WindowsMarketingName}\",\r\n" +
                        $"             \"operatingSystemVersion\": \"{WindowsVersion}\" \r\n" +
                        $"          }}\r\n" +
                        $"       }},\r\n" +
                        $"       \"integratedAppMetadata\": {{\r\n" +
                        $"          \"name\": \"CA Purview API Explorer\",\r\n" +
                        $"          \"version\": \"0.1\" \r\n" +
                        $"       }},\r\n" +
                        $"       \"userId\":\"{userId}\",\r\n" +
                        $"       \"scopeIdentifier\":\"0\"\r\n" +
                        $"    }}\r\n" +
                        $"}}";
                        break;
                }
                StatusBox.Text = $"{selectedItem.Content} API selected";
            }
        }

        private async void SendRequestButton_Click(object sender, RoutedEventArgs e)
        {
            string url = UrlTextBox.Text;
            string requestBody = RequestBodyTextBox.Text;
            ResponseTextBox.Text = string.Empty;
            ResponseHeadersTextBox.Text = string.Empty;

            logger.Log($"Request URL: {url}");

            Dictionary<string, string> response;
            try
            {
                StatusBox.Text = "Sending request...";
                response = await this.SendHttpRequest().ConfigureAwait(true);
                ResponseTextBox.Text = response["Content"];
                ResponseHeadersTextBox.Text = response["StatusCode"] + "\n" + response["Headers"];
                logger.Log($"Response StatusCode: {response["StatusCode"]}");
                logger.Log($"Response: {response["Content"]}");
                logger.Log($"Response Headers: {response["Headers"]}");
                StatusBox.Text = $"Request completed. {response["StatusCode"]}";
            }
            catch (Exception ex)
            {
                ResponseTextBox.Text = $"Error: {ex.Message}";
                logger.Log("Error: " + ex.Message);
                StatusBox.Text = "Request failed...";
            }

            LogTextBox.Text = sbLog.ToString();
            LogTextBox.CaretIndex = LogTextBox.Text.Length;
            LogTextBox.ScrollToEnd();
        }

        private async Task<Dictionary<string, string>> SendHttpRequest()
        {
            using (HttpClient client = new HttpClient())
            {
                HttpRequestMessage httpRequestMessage;
                string? accessToken = string.Empty;

                StatusBox.Text = "Acquiring user token...";
                string[] scopes = [Scope.Text];
                accessToken = await GetToken(TokenType.Access, scopes);
                logger.Log($"Access Token: {accessToken}", LogType.Console);

                if (httpOperation == "POST")
                {
                    httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, new Uri(UrlTextBox.Text));
                    httpRequestMessage.Content = new StringContent(RequestBodyTextBox.Text, Encoding.UTF8, "application/json");
                    logger.Log($"Request Body: {RequestBodyTextBox.Text}");
                    httpRequestMessage.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                }
                else
                {
                    httpRequestMessage = new HttpRequestMessage(HttpMethod.Get, new Uri(UrlTextBox.Text));
                    httpRequestMessage.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                }

                logger.Log($"Request Headers:\r\n");
                // Add headers
                if (!string.IsNullOrWhiteSpace(RequestHeadersTextBox.Text))
                {
                    string[] headersArray = RequestHeadersTextBox.Text.Trim().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
                    try
                    {
                        foreach (string header in headersArray)
                        {
                            string[] keyval = header.Trim().Split(':', StringSplitOptions.RemoveEmptyEntries);
                            AddOrUpdateHeader(httpRequestMessage, keyval[0].Trim(), keyval[1].Trim());
                            logger.Log($"{keyval[0].Trim()}: {keyval[1].Trim()}");
                        }
                    }
                    catch (Exception) { }
                }
                logger.Log($"\r\n");

                httpRequestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                
                StatusBox.Text = $"Sending {httpRequestMessage.Method} request to {httpRequestMessage.RequestUri?.ToString()}...";
                HttpResponseMessage httpResponse = await client.SendAsync(httpRequestMessage).ConfigureAwait(true);
                StatusBox.Text = "Reading response...";
                return await this.GetResponseStringAsync(httpResponse).ConfigureAwait(true);
            }
        }

        private void GraphVersionSelectBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (this.isInitialized && GraphVersionSelectBox.SelectedItem is ComboBoxItem selectedItem)
            {
                switch (selectedItem.Content.ToString())
                {
                    case "Beta":
                        this.baseUrl = "https://graph.microsoft.com/beta";
                        this.SetUpApiUrls();
                        break;

                    case "1.0":
                        this.baseUrl = "https://graph.microsoft.com/v1.0";
                        this.SetUpApiUrls();
                        break;

                    case "Canary":
                        this.baseUrl = "https://canary.graph.microsoft.com/testprodbetadcsdogfood";
                        this.SetUpApiUrls();
                        break;
                }
            }
        }

        private async void SignInBtn_Click(object sender, RoutedEventArgs e)
        {
            if (userId == string.Empty)
            {
                await SignInUser();
                if (userId != string.Empty)
                {
                    SendBtn.IsEnabled = true;
                    NewRequestBtn.IsEnabled = true;
                    ApiSelectBox.IsEnabled = true;
                    ClearTokenCache.IsEnabled = false;
                    SignInBtn.Content = "Sign Out";
                    this.SetUpApiViews();
                }
            }
            else
            {
                if (MSALPublicClientApp != null)
                {
                    var accounts = await MSALPublicClientApp.GetAccountsAsync();
                    if (accounts.Any())
                    {
                        try
                        {
                            await MSALPublicClientApp.RemoveAsync(accounts.FirstOrDefault());
                            userId = string.Empty;
                            tenantId = string.Empty;
                            SendBtn.IsEnabled = false;
                            NewRequestBtn.IsEnabled = false;
                            ApiSelectBox.IsEnabled = false;
                            ClearTokenCache.IsEnabled = true;
                            this.SetUpApiViews();
                            SignInBtn.Content = "Sign In";
                        }
                        catch (MsalException msalex)
                        {
                            logger.Log("Error Acquiring Token: " + msalex.Message);
                        }
                    }
                }
            }
        }

        private static void AddOrUpdateHeader(HttpRequestMessage request, string headerName, string headerValue)
        {
            if (request.Headers.Contains(headerName))
            {
                request.Headers.Remove(headerName);
            }
            request.Headers.Add(headerName, headerValue);
        }

        private void NewRequestBtn_Click(object sender, RoutedEventArgs e)
        {
            SetUpApiViews();
        }

        private async void ClearTokenCache_Click(object sender, RoutedEventArgs e)
        {
            if (MSALPublicClientApp != null)
            {
                var accounts = await MSALPublicClientApp.GetAccountsAsync();
                while (accounts.Any())
                {
                    await MSALPublicClientApp.RemoveAsync(accounts.First());
                    accounts = await MSALPublicClientApp.GetAccountsAsync();
                }
            }
            TokenCacheHelper.ClearCache();
        }
    }
}
