# Mark Teams Chats as Read Utility

This project was created to "unread" Teams chats.  During an automated migratation from one Office 365 tenant to another, the chat were constantly changing causing old chats to be unread.

## Build and Test

1. Open solution in Visual Studio
1. Build the solution
1. Run the solution

## Details

This project relies on an Azure AD Application that uses delegated permissions to read and write Teams data using the Microsoft Graph.  The first time you run the app you will need to consent that the Application has the requested delegated permissions.
