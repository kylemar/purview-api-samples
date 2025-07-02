// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace PurviewAPIExp
{
    using Microsoft.Identity.Client;
    using Microsoft.Identity.Client.Broker;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using Purview_API_Explorer;
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
        private const string SensitivityLabelsFormat = "{0}/security/dataSecurityAndGovernance/sensitivityLabels";

        /// <summary>
        /// Cache the ProtectionScopes State
        /// </summary>
        private string protectionScopeState = string.Empty;
        private bool needToCallProtectionScopes = true;

        /// <summary>
        /// Timer to check for ProtectionScopeState needed to up re-computed
        /// </summary>
        Timer? protectionScopeStateTimer = null;
        private DateTime protectionScopeStateIdleTime = DateTime.MinValue;

        /// <summary>
        /// going to need some random numbers for the data 
        /// </summary>
        private static Random random = new Random();
        private int applicationNumber = random.Next(0, Int32.MaxValue);

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
        private string userEmail = string.Empty;

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
        private string sensitivityLabelsUrl = string.Empty;
        private string baseUrl = "https://graph.microsoft.com/v1.0";

        /// <summary>
        /// Microsoft Authentication Library Public Client
        /// Used to sign in the user and acquire access tokens to call Purview API
        /// </summary>
        private static IPublicClientApplication? MSALPublicClientApp;
        private bool useBroker = true;

        /// <summary>
        /// All the permissions the app needs. Listed together for a single consent prompt on first sign in
        /// as opposed incrementally adding consent with multiple prompts as the user explores APIs
        /// </summary>
        private string[] signInScopes = new string[] { "user.read", "Content.Process.User", "ProtectionScopes.Compute.User", "ContentActivity.Write", "SensitivityLabel.Read", "SensitivityLabel.Evaluate" };

        /// <summary>   
        /// Current API Doc page URL
        /// </summary>
        private string currentApiDocUrl = "https://learn.microsoft.com/en-us/purview/developer/?branch=main";

        /// <summary
        /// Wait for the next API call to return before proceding?
        /// </summary>
        private bool waitForApiCall = true;

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

            protectionScopeStateTimer = new Timer(ProtectionScopeStateTimerCallback, null, TimeSpan.FromMinutes(1), TimeSpan.FromMinutes(1));
        }

        private async Task SignInUser()
        {
            PublicClientApplicationBuilder builder = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(AadAuthorityAudience.AzureAdMultipleOrgs)
                // Use Continous Access Evaluation (CAE)
                .WithClientCapabilities(new[] { "cp1" });

            if (useBroker == true)
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
                    userEmail = account.Username;
                    userName.Text = account.Username;
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

        private void ApiSelectBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (userId == string.Empty)
            {
                return;
            }

            int newIndex = SetUpApiViews();
            if (newIndex != -1)
            {
                ApiSelectBox.SelectedItem = ApiSelectBox.Items[newIndex];
            }
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
            this.sensitivityLabelsUrl = string.Format(CultureInfo.InvariantCulture, SensitivityLabelsFormat, this.baseUrl);
            this.SetUpApiViews();
        }

        /// <summary>
        /// Conversation attributes for Process Content API
        /// Allows for a continued conversation
        /// </summary>
        private string conversationId = string.Empty;
        private int conversationSequence = 0;

        private int SetUpApiViews()
        {
            if (userId == string.Empty)
            {
                UrlTextBox.Text = string.Empty;
                Scope.Text = string.Empty;
                return 0;
            }

            if (this.isInitialized && ApiSelectBox.SelectedItem is ComboBoxItem selectedItem)
            {
                ResponseTextBox.Text = string.Empty;

                StringBuilder headers = new StringBuilder();
                headers.AppendLine("User-Agent:Purview API Explorer");
                headers.AppendLine($"client-request-id:{Guid.NewGuid().ToString()}");
                headers.AppendLine($"x-ms-client-request-id:{Guid.NewGuid().ToString()}");

                switch (selectedItem.Content.ToString())
                {
                    case "Process Content - Start Conversation":
                        if (CheckForProtectionScopeState(headers) == false)
                        {
                            return 0;
                        }

                        if (ProtectionScopeStateCache.Prompts == CallPurviewType.Dont)
                        {
                            var popup = new MessagePopup("The tenant is not configured for your app to call ProcessContent for Prompts. You should offer the options to call Content Activities for prompts.", "Do not call Process Content for prompts");
                            popup.Owner = this; // Set owner for modal behavior
                            popup.ShowDialog();
                            return 4;
                        }
                        else
                        {
                            currentApiDocUrl = "https://learn.microsoft.com/en-us/graph/api/userdatasecurityandgovernance-processcontent";
                            httpOperation = "POST";
                            UrlTextBox.Text = this.pcUrl;
                            Scope.Text = "Content.Process.User";
                            RequestContentTabControl.SelectedIndex = 0;
                            conversationId = Guid.NewGuid().ToString();
                            conversationSequence = 0;
                            if (ProtectionScopeStateCache.Prompts == CallPurviewType.Inline)
                            {
                                waitForApiCall = true;
                            }
                            else
                            {
                                waitForApiCall = false;
                            }
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
                            $"          \"operatingSystemSpecifications\": {{\r\n" +
                            $"             \"operatingSystemPlatform\": \"{WindowsMarketingName}\",\r\n" +
                            $"             \"operatingSystemVersion\": \"{WindowsVersion}\" \r\n" +
                            $"          }},\r\n" +
                            $"          \"ipAddress\": \"127.0.0.1\"\r\n" +
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
                        }
                        break;

                    case "Process Content - Continue Conversation with Response":
                        if (CheckForProtectionScopeState(headers) == false)
                        {
                            return 0;
                        }
                        if (ProtectionScopeStateCache.Responses == CallPurviewType.Dont)
                        {
                            var popup = new MessagePopup("The tenant is not configured for your app to call ProcessContent for Responses. You should offer the options to call Content Activities for Responses.", "Do not call Process Content for Responses");
                            popup.Owner = this; // Set owner for modal behavior
                            popup.ShowDialog();
                            return 4;
                        }
                        else
                        {
                            currentApiDocUrl = "https://learn.microsoft.com/en-us/graph/api/userdatasecurityandgovernance-processcontent";
                            httpOperation = "POST";
                            UrlTextBox.Text = this.pcUrl;
                            Scope.Text = "Content.Process.User";
                            if (conversationId == string.Empty)
                            {
                                conversationId = Guid.NewGuid().ToString();
                                conversationSequence = 0;
                            }
                            RequestContentTabControl.SelectedIndex = 0;
                            if (ProtectionScopeStateCache.Responses == CallPurviewType.Inline)
                            {
                                waitForApiCall = true;
                            }
                            else
                            {
                                waitForApiCall = false;
                            }
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
                            $"          \"operatingSystemSpecifications\": {{\r\n" +
                            $"             \"operatingSystemPlatform\": \"{WindowsMarketingName}\",\r\n" +
                            $"             \"operatingSystemVersion\": \"{WindowsVersion}\" \r\n" +
                            $"          }},\r\n" +
                            $"          \"ipAddress\": \"127.0.0.1\"\r\n" +
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
                        }
                        break;

                    case "Process Content - Continue Conversation with Prompt":
                        if (CheckForProtectionScopeState(headers) == false)
                        {
                            return 0;
                        }
                        if (ProtectionScopeStateCache.Prompts == CallPurviewType.Dont)
                        {
                            var popup = new MessagePopup("The tenant is not configured for your app to call ProcessContent for Prompts. You should offer the options to call Content Activities for prompts.", "Do not call Process Content for prompts");
                            popup.Owner = this; // Set owner for modal behavior
                            popup.ShowDialog();
                            return 4;
                        }
                        else
                        {
                            currentApiDocUrl = "https://learn.microsoft.com/en-us/graph/api/userdatasecurityandgovernance-processcontent";
                            httpOperation = "POST";
                            UrlTextBox.Text = this.pcUrl;
                            Scope.Text = "Content.Process.User";
                            if (conversationId == string.Empty)
                            {
                                conversationId = Guid.NewGuid().ToString();
                                conversationSequence = 0;
                            }
                            RequestContentTabControl.SelectedIndex = 0;
                            if (ProtectionScopeStateCache.Prompts == CallPurviewType.Inline)
                            {
                                waitForApiCall = true;
                            }
                            else
                            {
                                waitForApiCall = false;
                            }
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
                            $"          \"operatingSystemSpecifications\": {{\r\n" +
                            $"             \"operatingSystemPlatform\": \"{WindowsMarketingName}\",\r\n" +
                            $"             \"operatingSystemVersion\": \"{WindowsVersion}\" \r\n" +
                            $"          }},\r\n" +
                            $"          \"ipAddress\": \"127.0.0.1\"\r\n" +
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
                        }
                        break;

                    case "Protection Scopes - Initial Call":
                        if (DateTime.Now < protectionScopeStateIdleTime.AddMinutes(30))
                        {
                            var popup = new ProtectionScopeInfo(ProtectionScopeMessageType.TooSoon);
                            popup.Owner = this; // Set owner for modal behavior
                            popup.ShowDialog();
                        }
                        currentApiDocUrl = "https://learn.microsoft.com/en-us/graph/api/userprotectionscopecontainer-compute";
                        httpOperation = "POST";
                        UrlTextBox.Text = this.userPsUrl;
                        Scope.Text = "ProtectionScopes.Compute.User";
                        RequestContentTabControl.SelectedIndex = 0;
                        waitForApiCall = true;
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
                        currentApiDocUrl = "https://learn.microsoft.com/en-us/graph/api/activitiescontainer-post-contentactivities";
                        UrlTextBox.Text = this.activityUrl;
                        Scope.Text = "ContentActivity.Write";
                        conversationId = Guid.NewGuid().ToString();
                        conversationSequence = 0;
                        RequestContentTabControl.SelectedIndex = 0;
                        waitForApiCall = false;
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
                        $"          \"operatingSystemSpecifications\": {{\r\n" +
                        $"             \"operatingSystemPlatform\": \"{WindowsMarketingName}\",\r\n" +
                        $"             \"operatingSystemVersion\": \"{WindowsVersion}\" \r\n" +
                        $"          }},\r\n" +
                        $"          \"ipAddress\": \"127.0.0.1\"\r\n" +
                        $"       }},\r\n" +
                        $"       \"integratedAppMetadata\": {{\r\n" +
                        $"          \"name\": \"CA Purview API Explorer\",\r\n" +
                        $"          \"version\": \"0.1\" \r\n" +
                        $"       }}\r\n" +
                        $"    }}\r\n" +
                        $"}}";
                        break;

                    case "List all Sensitivity Labels":
                        currentApiDocUrl = "https://learn.microsoft.com/en-us/graph/api/tenantdatasecurityandgovernance-list-sensitivitylabels";
                        httpOperation = "GET";
                        UrlTextBox.Text = this.sensitivityLabelsUrl;
                        Scope.Text = "SensitivityLabel.Read";
                        RequestContentTabControl.SelectedIndex = 0;
                        waitForApiCall = true;
                        RequestBodyTextBox.Text = string.Empty;
                        break;

                    case "Get Label Details For Given Label Id":
                        currentApiDocUrl = "https://learn.microsoft.com/en-us/graph/api/sensitivitylabel-get";
                        httpOperation = "GET";
                        UrlTextBox.Text = $"{this.sensitivityLabelsUrl}/defa4170-0d19-0005-0009-bc88714345d2";
                        Scope.Text = "SensitivityLabel.Read";
                        RequestContentTabControl.SelectedIndex = 0;
                        waitForApiCall = true;
                        RequestBodyTextBox.Text = string.Empty;
                        break;

                    case "Get Rights For Given Label Id for the user":
                        httpOperation = "GET";
                        currentApiDocUrl = "https://learn.microsoft.com/en-us/graph/api/usagerightsincluded-get";
                        UrlTextBox.Text = $"{this.sensitivityLabelsUrl}/defa4170-0d19-0005-0009-bc88714345d2/rights";
                        Scope.Text = "SensitivityLabel.Read";
                        RequestContentTabControl.SelectedIndex = 0;
                        RequestBodyTextBox.Text = string.Empty;
                        waitForApiCall = true;
                        break;

                    case "Compute Inheritance":
                        currentApiDocUrl = "https://learn.microsoft.com/en-us/graph/api/sensitivitylabel-computeinheritance";
                        httpOperation = "GET";
                        UrlTextBox.Text = $"{this.sensitivityLabelsUrl}/computeInheritance(labelIds=[\"defa4170-0d19-0005-0007-bc88714345d2\",\"defa4170-0d19-0005-0001-bc88714345d2\",\"defa4170-0d19-0005-000a-bc88714345d2\"],locale='en-US',contentFormats=[\"File\"])";
                        Scope.Text = "SensitivityLabel.Evaluate";
                        RequestContentTabControl.SelectedIndex = 0;
                        RequestBodyTextBox.Text = string.Empty;
                        waitForApiCall = true;
                        break;

                    case "Compute Rights and Inheritance":
                        currentApiDocUrl = "https://learn.microsoft.com/en-us/graph/api/sensitivitylabel-computerightsandinheritance";
                        httpOperation = "POST";
                        UrlTextBox.Text = $"{this.sensitivityLabelsUrl}/computeRightsAndInheritance";
                        Scope.Text = "SensitivityLabel.Evaluate";
                        RequestContentTabControl.SelectedIndex = 0;
                        waitForApiCall = true;
                        RequestBodyTextBox.Text =
                        $"{{\r\n" +
                        $"   \"delegatedUserEmail\": \"{userEmail}\",\r\n" +
                        $"   \"locale\": \"en-us\",\r\n" +
                        $"   \"protectedContents\": [\r\n" +
                        $"      {{\r\n" +
                        $"         \"LabelId\": \"defa4170-0d19-0005-0007-bc88714345d2\",\r\n" +
                        $"         \"contentFormat\": \"File\",\r\n" +
                        $"         \"contentId\": \"doc-234\"\r\n" +
                        $"      }},\r\n" +
                        $"      {{\r\n" +
                        $"         \"LabelId\": \"defa4170-0d19-0005-0001-bc88714345d2\",\r\n" +
                        $"         \"contentFormat\": \"File\",\r\n" +
                        $"         \"contentId\": \"doc-345\"\r\n" +
                        $"      }},\r\n" +
                        $"      {{\r\n" +
                        $"         \"LabelId\": \"defa4170-0d19-0005-000a-bc88714345d2\",\r\n" +
                        $"         \"contentFormat\": \"File\",\r\n" +
                        $"         \"contentId\": \"doc-456\"\r\n" +
                        $"      }}\r\n" +
                        $"   ],\r\n" +
                        $"   \"supportedContentFormats\": [\r\n" +
                        $"      \"File\"\r\n" +
                        $"   ]\r\n" +
                        $"}}";
                break;

                }
                StatusBox.Text = $"{selectedItem.Content} API selected";
                return -1;
            }
            return 0;
        }

        private bool CheckForProtectionScopeState(StringBuilder headers)
        {
            if (needToCallProtectionScopes == false)
            {
                headers.AppendLine($"If-None-Match:{protectionScopeState}");
                RequestHeadersTextBox.Text = headers.ToString();
                return true;
            }
            else
            {
                var popup = new ProtectionScopeInfo(ProtectionScopeMessageType.NoState);
                popup.Owner = this; // Set owner for modal behavior
                popup.ShowDialog();
                return false;
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
                response = await this.SendHttpRequest(waitForApiCall).ConfigureAwait(true);
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

        private async Task<Dictionary<string, string>> SendHttpRequest(bool waitForResponse)
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
                
                StatusBox.Text = $"Sending {httpRequestMessage.Method} request ...";

                if (waitForResponse)
                {
                    // Wait for the response
                    HttpResponseMessage httpResponse = await client.SendAsync(httpRequestMessage).ConfigureAwait(true);
                    StatusBox.Text = "Reading response...";
                    return await this.GetResponseStringAsync(httpResponse).ConfigureAwait(true);
                }
                else
                {
                    _ = client.SendAsync(httpRequestMessage).ConfigureAwait(false);
                    StatusBox.Text = "Not waiting for a response...";

                    Dictionary<string, string> responseDict = new Dictionary<string, string>(3, StringComparer.OrdinalIgnoreCase);

                    responseDict["StatusCode"] = "Not waiting for a response.";
                    responseDict["Headers"] = "Not waiting for a response.";
                    responseDict["Content"] = "Not waiting for a response.";

                    return responseDict;
                }
            }
        }

        private async Task<Dictionary<string, string>> GetResponseStringAsync(HttpResponseMessage response)
        {
            dynamic? parsedJson = null;
            string prettyJson = "";
            ComboBoxItem selectedItem = (ComboBoxItem)ApiSelectBox.SelectedItem;
            string? API = selectedItem.Content.ToString();

            Dictionary<string, string> responseDict = new Dictionary<string, string>(3, StringComparer.OrdinalIgnoreCase);

            string responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(true);

            if (responseContent != null)
            {
                try
                {
                    parsedJson = JsonConvert.DeserializeObject<dynamic>(responseContent);
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

            responseDict["StatusCode"] = $"StatusCode: {(int)response.StatusCode} - {response.ReasonPhrase}";

            if (API != null && API.StartsWith("Protection Scopes"))
            {
                ProtectionScopeStateCache.ParseProtectionScopeState(parsedJson);
                ProtectionScopeStateBox.Text = $"Prompts:{ProtectionScopeStateCache.Prompts}  Responses:{ProtectionScopeStateCache.Responses}";
                needToCallProtectionScopes = false;
            }

            if (API != null && API.StartsWith("Process Content"))
            {
                if (parsedJson != null)
                {
                    string protectionScopeStateValue = parsedJson?["protectionScopeState"]?.ToString() ?? string.Empty;
                    if (!string.IsNullOrEmpty(protectionScopeStateValue))
                    {
                        if (protectionScopeStateValue == "modified")
                        {
                            protectionScopeState = string.Empty;
                            ProtectionScopeStateBox.Text = "Not cached! - Modified. Call ProtectionScopes/compute";
                            needToCallProtectionScopes = true;

                            var popup = new ProtectionScopeInfo(ProtectionScopeMessageType.Modified);
                            popup.Owner = this; // Set owner for modal behavior
                            popup.ShowDialog();
                        }
                        else
                        {
                            protectionScopeStateIdleTime = DateTime.Now;  // Reset idle timer for protection scope state
                            needToCallProtectionScopes = false;
                        }
                    }

                    JArray? policyActions = parsedJson?["policyActions"] as JArray;
                    if (policyActions != null && policyActions.Count > 0)
                    {
                        foreach (var action in policyActions)
                        {
                            if (action["@odata.type"]?.ToString() == "#microsoft.graph.restrictAccessAction")
                            {
                                string actionType = action["action"]?.ToString() ?? "Unknown";
                                string restrictionAction = action["restrictionAction"]?.ToString() ?? "No value";
                                var popup = new MessagePopup($"Your app needs to take steps required for a {restrictionAction} restriction", $"{actionType}");
                                popup.Owner = this; // Set owner for modal behavior
                                popup.ShowDialog();
                            }
                        }
                    }


                }
            }

            if (response.Headers.ETag != null)
            {
                protectionScopeState = response.Headers.ETag.ToString();
                protectionScopeStateIdleTime = DateTime.Now;
                needToCallProtectionScopes = false;
            }

            responseDict["Headers"] = $"{response.Headers}";
            responseDict["Content"] = $"{prettyJson}";

            return responseDict;
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
                }
            }
        }

        private async void SignInBtn_Click(object sender, RoutedEventArgs e)
        {
            if (userId == string.Empty)
            {
                useBroker = true; // Set to true for using WAM broker
                await SignInUser();
                if (userId != string.Empty)
                {
                    SendBtn.IsEnabled = true;
                    NewRequestBtn.IsEnabled = true;
                    ApiSelectBox.IsEnabled = true;
                    SignedIn.Visibility = Visibility.Visible;
                    SignedOut.Visibility = Visibility.Collapsed;
                    SetUpApiViews();
                }
            }
        }

        private void ProtectionScopeStateTimerCallback(object? state)
        {
            if (DateTime.Now >= protectionScopeStateIdleTime.AddMinutes(30))
            {
                Dispatcher.BeginInvoke(new Action(delegate
                {
                    protectionScopeState = string.Empty;
                    protectionScopeStateIdleTime = DateTime.MinValue;
                    ProtectionScopeStateBox.Text = "Not cached! - 30 minutes of idle time";
                    needToCallProtectionScopes = true;
                    logger.Log("Protection scope state cache cleared after 30 minutes of inactivity.");
                }));
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

        private void Docs_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(currentApiDocUrl))
            {
                return;
            }
            // Open the documentation link in the default web browser
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(currentApiDocUrl) { UseShellExecute = true });
        }

        private async void SignOut_Click(object sender, RoutedEventArgs e)
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
                        SetUpApiViews();
                        SignedIn.Visibility = Visibility.Collapsed;
                        SignedOut.Visibility = Visibility.Visible;
                        userName.Text = "Please sign in...";
                        currentApiDocUrl = "https://learn.microsoft.com/en-us/purview/developer/?branch=main";
                        protectionScopeState = string.Empty;
                        ProtectionScopeStateBox.Text = "Not cached!";
                    }
                    catch (MsalException msalex)
                    {
                        logger.Log("Error Acquiring Token: " + msalex.Message);
                    }
                }
            }
        }

        private async void SignInBrowserBtn_Click(object sender, RoutedEventArgs e)
        {
            if (userId == string.Empty)
            {
                useBroker = false; // Set to false for using browser sign-in
                await SignInUser();
                if (userId != string.Empty)
                {
                    SendBtn.IsEnabled = true;
                    NewRequestBtn.IsEnabled = true;
                    ApiSelectBox.IsEnabled = true;
                    SignedIn.Visibility = Visibility.Visible;
                    SignedOut.Visibility = Visibility.Collapsed;
                    SetUpApiViews();
                }
            }
        }

        private void ResetLog_Click(object sender, RoutedEventArgs e)
        {
            LogTextBox.Text = string.Empty;
            logger.ResetLog();
        }
    }
}
