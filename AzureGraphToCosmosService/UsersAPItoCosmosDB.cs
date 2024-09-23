using Microsoft.Azure.Cosmos;
using Microsoft.Graph;
using Azure.Identity;

namespace AzureAdGraphToCosmos
{
    class Program
    {
        // Azure AD Authentication and Graph API details
        private static string tenantId = "7a880873-6b76-4a2c-99f7-9d76741cf744";     // Get from Azure AD App Registration
        private static string clientId = "3eb397be-6c64-4f13-8ae7-541a603d6463";     // Get from Azure AD App Registration
        private static string clientSecret = "aD68Q~UIxDh2dizplCVLIpDGeE4VyH_Go9eedaFD"; // Created in Azure AD

        // Cosmos DB details
        private static string cosmosEndpoint = "https://graphapiresponse.documents.azure.com:443/";
        private static string cosmosPrimaryKey = "wrsCQPdr0uYONFFCGv1rb6z8BnwH9FoTfEInKMelOWx8kkmj4zCp0C7de81ir36xyjceoQI8BwzcACDbrgv9mw==";
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
                            requestConfiguration.QueryParameters.Select = new[] { "displayName", "signInActivity", "accountEnabled", "employeeLeaveDateTime", "createdDateTime"};
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
                        createdDateTime = currentUser.CreatedDateTime
                    };
                    userList.Add(userObj);
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
