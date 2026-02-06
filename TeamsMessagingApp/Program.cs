using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace TeamsMessagingApp
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("Teams Messaging App - Using Azure Credential Manager");
            Console.WriteLine("=====================================================\n");

            // Get the user email from command line argument or prompt
            string? targetUserEmail = args.Length > 0 ? args[0] : null;
            
            if (string.IsNullOrWhiteSpace(targetUserEmail))
            {
                Console.Write("Enter the email address of the user to send a message to: ");
                targetUserEmail = Console.ReadLine();
            }

            if (string.IsNullOrWhiteSpace(targetUserEmail))
            {
                Console.WriteLine("Error: Email address is required.");
                return;
            }

            try
            {
                // Use DefaultAzureCredential to automatically detect available credentials
                // This will try multiple credential sources in order:
                // 1. Environment variables
                // 2. Managed Identity
                // 3. Visual Studio
                // 4. Azure CLI
                // 5. Azure PowerShell
                // 6. Interactive browser
                Console.WriteLine("Detecting Azure credentials...");
                var credential = new DefaultAzureCredential(new DefaultAzureCredentialOptions
                {
                    ExcludeInteractiveBrowserCredential = false // Allow interactive login if needed
                });

                // Define the required scopes for Microsoft Graph
                var scopes = new[] { "https://graph.microsoft.com/.default" };

                // Create the Graph client
                var graphClient = new GraphServiceClient(credential, scopes);

                Console.WriteLine("Credentials detected successfully!");
                Console.WriteLine($"\nLooking up user: {targetUserEmail}");

                // Look up the user by email address
                var user = await FindUserByEmail(graphClient, targetUserEmail);

                if (user == null)
                {
                    Console.WriteLine($"Error: User with email '{targetUserEmail}' not found.");
                    return;
                }

                Console.WriteLine($"Found user: {user.DisplayName} (ID: {user.Id})");

                // Get or create a chat with the user and send a message
                await SendTeamsMessage(graphClient, user.Id!, targetUserEmail);

                Console.WriteLine("\n✓ Message sent successfully!");
            }
            catch (Azure.Identity.AuthenticationFailedException ex)
            {
                Console.WriteLine($"\nAuthentication Error: {ex.Message}");
                Console.WriteLine("\nPlease ensure you have valid Azure credentials configured.");
                Console.WriteLine("You can authenticate using:");
                Console.WriteLine("  - Azure CLI: az login");
                Console.WriteLine("  - Environment variables: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET");
                Console.WriteLine("  - Managed Identity (when running in Azure)");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nError: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner Error: {ex.InnerException.Message}");
                }
            }
        }

        static async Task<User?> FindUserByEmail(GraphServiceClient graphClient, string email)
        {
            try
            {
                // Sanitize the email to prevent OData injection
                string sanitizedEmail = email.Replace("'", "''");
                
                // Search for the user by email
                var users = await graphClient.Users.GetAsync(requestConfig =>
                {
                    requestConfig.QueryParameters.Filter = $"mail eq '{sanitizedEmail}' or userPrincipalName eq '{sanitizedEmail}'";
                    requestConfig.QueryParameters.Select = new[] { "id", "displayName", "mail", "userPrincipalName" };
                });

                return users?.Value?.FirstOrDefault();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error looking up user: {ex.Message}");
                return null;
            }
        }

        static async Task SendTeamsMessage(GraphServiceClient graphClient, string userId, string userEmail)
        {
            try
            {
                // Get the authenticated user's ID (the sender)
                var me = await graphClient.Me.GetAsync();
                if (me?.Id == null)
                {
                    throw new Exception("Could not retrieve authenticated user information");
                }

                Console.WriteLine($"Authenticated as: {me.DisplayName} ({me.Mail ?? me.UserPrincipalName})");

                // Create a new chat or get existing chat between the two users
                var chat = new Chat
                {
                    ChatType = ChatType.OneOnOne,
                    Members = new List<ConversationMember>
                    {
                        new AadUserConversationMember
                        {
                            Roles = new List<string> { "owner" },
                            AdditionalData = new Dictionary<string, object>
                            {
                                { "user@odata.bind", $"https://graph.microsoft.com/v1.0/users('{me.Id}')" }
                            }
                        },
                        new AadUserConversationMember
                        {
                            Roles = new List<string> { "owner" },
                            AdditionalData = new Dictionary<string, object>
                            {
                                { "user@odata.bind", $"https://graph.microsoft.com/v1.0/users('{userId}')" }
                            }
                        }
                    }
                };

                Console.WriteLine($"\nCreating or retrieving chat with {userEmail}...");
                Chat? createdChat = null;

                try
                {
                    createdChat = await graphClient.Chats.PostAsync(chat);
                }
                catch (Microsoft.Graph.Models.ODataErrors.ODataError odataEx)
                {
                    // If chat already exists, try to find it
                    Console.WriteLine("Chat may already exist, searching for existing chat...");
                    var chats = await graphClient.Me.Chats.GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.Filter = "chatType eq 'oneOnOne'";
                        requestConfig.QueryParameters.Expand = new[] { "members" };
                    });

                    if (chats?.Value != null)
                    {
                        foreach (var existingChat in chats.Value)
                        {
                            if (existingChat.Members?.Any(m =>
                            {
                                if (m is AadUserConversationMember aadMember)
                                {
                                    return aadMember.UserId == userId;
                                }
                                return false;
                            }) == true)
                            {
                                createdChat = existingChat;
                                Console.WriteLine("Found existing chat!");
                                break;
                            }
                        }
                    }

                    if (createdChat == null)
                    {
                        throw new Exception($"Could not create or find chat: {odataEx.Error?.Message ?? odataEx.Message}");
                    }
                }

                if (createdChat?.Id == null)
                {
                    throw new Exception("Failed to create or retrieve chat");
                }

                // Send a message in the chat
                var message = new ChatMessage
                {
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Text,
                        Content = $"Hello! This is an automated message sent via Microsoft Graph API using Azure credentials. Sent at {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC."
                    }
                };

                Console.WriteLine("Sending Teams message...");
                await graphClient.Chats[createdChat.Id].Messages.PostAsync(message);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error sending Teams message: {ex.Message}", ex);
            }
        }
    }
}
