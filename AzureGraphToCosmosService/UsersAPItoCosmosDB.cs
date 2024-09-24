using Microsoft.Azure.Cosmos;
using Microsoft.Graph;
using Azure.Identity;

namespace AzureAdGraphToCosmos
{
    class Program
    {
        // Azure AD Authentication and Graph API details
        private static string tenantId = "";     // Get from Azure AD App Registration
        private static string clientId = "";     // Get from Azure AD App Registration
        private static string clientSecret = ""; // Created in Azure AD

        // Cosmos DB details
        private static string cosmosEndpoint = "";
        private static string cosmosPrimaryKey = "";
        private static string databaseId = "UserDatabase";
        private static string containerId = "UserContainer";

        static async Task Main(string[] args)
        {
            // Initialize Cosmos DB service
            var cosmosService = new CosmosDbService(cosmosEndpoint, cosmosPrimaryKey, databaseId, containerId);
            await cosmosService.InitializeAsync();

            // Get users from Microsoft Graph API
            var graphClient = GetAuthenticatedGraphClient();
            var users = await GetUsersAsync(graphClient);

            // Store users in Cosmos DB
            foreach (var user in users)
            {
                await cosmosService.AddUserAsync(user);
            }

            Console.WriteLine("User data successfully stored in Cosmos DB.");
        }

        // Function to create and authenticate GraphServiceClient
        private static GraphServiceClient GetAuthenticatedGraphClient()
        {
            // Use Azure.Identity's ClientSecretCredential for authentication
            var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);

            // Initialize GraphServiceClient with the credential
            var graphClient = new GraphServiceClient(clientSecretCredential);
            return graphClient;
        }

        // Function to fetch users from Microsoft Graph API
        private static async Task<List<dynamic>> GetUsersAsync(GraphServiceClient graphClient)
        {
            List<dynamic> userList = new List<dynamic>();
            try
            {
                var users = await graphClient.Users.GetAsync();

                foreach (var user in users.Value)
                {
                   
                    var currentUser = await graphClient.Users[user.Id]
                        .GetAsync((requestConfiguration) =>
                        {
                            requestConfiguration.QueryParameters.Select = new[] { 
                                "displayName",
                                "signInActivity",
                                "accountEnabled",
                                "employeeLeaveDateTime",
                                "createdDateTime",
                                "onPremisesExtensionAttributes",
                                "provisioningStatus",
                                "userType" };
                        });

                    Console.WriteLine("User processed");
                    Console.WriteLine(currentUser);

                    dynamic userObj = new
                    {
                        id = currentUser.Id,
                        displayName = currentUser.DisplayName,
                        userPrincipalName = currentUser.UserPrincipalName,
                        accountEnabled = currentUser.AccountEnabled,
                        signInActivity = currentUser.SignInActivity,
                        employeeLeaveDateTime = currentUser.EmployeeLeaveDateTime,
                        createdDateTime = currentUser.CreatedDateTime,
                        userType = currentUser.UserType
                    };
                    userList.Add(userObj);
                    break;
                }
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error fetching users: {ex.Message}");
            }

            return userList;
        }
    }

    // Cosmos DB Service Class
    public class CosmosDbService
    {
        private CosmosClient _cosmosClient;
        private Database _database;
        private Container _container;
        private readonly string _databaseId;
        private readonly string _containerId;

        public CosmosDbService(string cosmosEndpoint, string cosmosPrimaryKey, string databaseId, string containerId)
        {
            _cosmosClient = new CosmosClient(cosmosEndpoint, cosmosPrimaryKey);
            _databaseId = databaseId;
            _containerId = containerId;
        }

        public async Task InitializeAsync()
        {
            _database = await _cosmosClient.CreateDatabaseIfNotExistsAsync(_databaseId);
            _container = await _database.CreateContainerIfNotExistsAsync(_containerId, "/id");
        }

        public async Task AddUserAsync(dynamic user)
        {
            try
            {
                await _container.CreateItemAsync(user, new PartitionKey(user.id.ToString()));
            }
            catch (CosmosException ex)
            {
                Console.WriteLine($"Error storing data in Cosmos DB: {ex.Message}");
            }
        }
    }
}
