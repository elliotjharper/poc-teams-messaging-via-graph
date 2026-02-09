# poc-teams-messaging-via-graph

A .NET Core console application that demonstrates how to use Azure credential manager library to authenticate and send Microsoft Teams messages via the Microsoft Graph API.

## Features

- **Automatic Credential Detection**: Uses `DefaultAzureCredential` from Azure.Identity library to automatically detect and use available credentials
- **User Lookup**: Finds users by their email address using Microsoft Graph API
- **Teams Messaging**: Sends Teams messages to users through one-on-one chats

## Prerequisites

- .NET 8.0 SDK or later
- Azure credentials configured (see Authentication section below)
- Microsoft 365 tenant with Teams enabled
- Appropriate Microsoft Graph API permissions

## Required Permissions

The application requires the following Microsoft Graph API permissions:

- `User.Read.All` - To look up users by email
- `Chat.ReadWrite` - To create chats and send messages
- `ChatMessage.Send` - To send chat messages

These permissions need to be granted to the Azure AD application or user account being used for authentication.

## Authentication

The application uses `DefaultAzureCredential` which tries multiple authentication methods in the following order:

1. **Environment Variables** - `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`
2. **Managed Identity** - When running in Azure (App Service, VM, etc.)
3. **Visual Studio** - When running on a development machine with Visual Studio signed in
4. **Azure CLI** - When authenticated via `az login`
5. **Azure PowerShell** - When authenticated via Azure PowerShell
6. **Interactive Browser** - Falls back to browser-based authentication if needed

### Quick Start with Azure CLI

The easiest way to authenticate for development is using Azure CLI:

```bash
az login
```

## Building the Application

```bash
cd TeamsMessagingApp
dotnet build
```

## Running the Application

### Option 1: Pass email as command-line argument

```bash
cd TeamsMessagingApp
dotnet run user@example.com
```

### Option 2: Interactive prompt

```bash
cd TeamsMessagingApp
dotnet run
```

The application will prompt you for the email address.

## Example Output

```
Teams Messaging App - Using Azure Credential Manager
=====================================================

Detecting Azure credentials...
Credentials detected successfully!

Looking up user: user@example.com
Found user: John Doe (ID: abc123-def456-...)
Authenticated as: Jane Smith (jane.smith@example.com)

Creating or retrieving chat with user@example.com...
Sending Teams message...

✓ Message sent successfully!
```

## Troubleshooting

### Authentication Errors

If you encounter authentication errors:

1. **Verify Azure CLI login**: Run `az login` and ensure you're signed in
2. **Check permissions**: Ensure your account has the required Graph API permissions
3. **Use environment variables**: Set `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, and `AZURE_CLIENT_SECRET` if using service principal

### User Not Found

If the user lookup fails:

- Verify the email address is correct
- Ensure the user exists in your Azure AD tenant
- Check that you have `User.Read.All` permission

### Message Send Failures

If sending messages fails:

- Verify you have `Chat.ReadWrite` and `ChatMessage.Send` permissions
- Ensure Teams is enabled for both users
- Check that the target user can receive Teams messages

## Project Structure

```
TeamsMessagingApp/
├── Program.cs              # Main application code
├── TeamsMessagingApp.csproj # Project file with dependencies
```

## Dependencies

- `Azure.Identity` - For Azure credential management
- `Microsoft.Graph` - For Microsoft Graph API client
- `Microsoft.Extensions.Configuration` - For configuration management
- `Microsoft.Extensions.Configuration.Json` - For JSON configuration support

## License

This is a proof-of-concept project for demonstration purposes.