# Context
I want to create a simple dotnet based cli tool that can authenticate to ms graph once (or maybe once a day, depending on whats available), and I can call it from my terminal to get my emails or chat messages etc etc.

The intent is to eventuall allow Claude to be able to use it to automate some of the things for me. I'll just authenticate it once via the cli, and Claude can then do stuff with it.

# Ticket to IT

```
I need an app registration in Confiz AAD with the following details:

Name: SK Desktop Assistant

Authentication:

Supported Account Types: Single Tenant

Platform: Mobile and Desktop Application

Redirect URIs:

https://login.microsoftonline.com/common/oauth2/nativeclient

http://localhost

API Permissions:

Microsoft Graph Delegated Permissions (these will only allow me to access things that I already have access to and nothing more):

Calendars.Read.Shared

Calendars.ReadWrite

Chat.Create

Chat.ReadWrite

ChatMessage.Read

ChatMessage.Send

Mail.ReadWrite

Mail.Send

Presence.Read.All

Tasks.ReadWrite

User.Read

User.ReadBasic.All

Make sure to “Grant admin consent”

Thanks
```