# Syncro Ticket Creator Add-in for Outlook

This Outlook add-in allows you to create Syncro tickets directly from your emails, streamlining your support workflow.

## Features

- Create Syncro tickets from Outlook emails
- Automatically populate ticket information from email content
- Select customers, contacts, and assets from your Syncro account
- Customize ticket title and message before creation

## Installation

1. Clone this repository or download the source code.
2. Install dependencies by running `npm install` in the project directory.
3. Build the project using `npm run build`.
4. Sideload the add-in in Outlook following Microsoft's instructions for [sideloading Outlook add-ins](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).

## Usage

1. Open an email in Outlook.
2. Click on the "Syncro Ticket Creator" button in the Outlook ribbon.
3. If it's your first time using the add-in, you'll be prompted to enter your Syncro API key and URL. Click "Open Settings" to do so.
4. Once settings are configured, the add-in will load your Syncro customers.
5. Select a customer, contact, and optionally an asset.
6. The ticket title and message will be pre-filled with the email subject and body. You can edit these if needed.
7. Click "Create Ticket" to submit the ticket to Syncro.

## Configuration

To use this add-in, you need to provide your Syncro API key and URL:

1. Open the add-in in Outlook.
2. Click the "Open Settings" button if you haven't configured the settings yet.
3. Enter your Syncro API key and URL (e.g., https://yourcompany.syncromsp.com).
4. Click "Save Settings".

## Development

To run the add-in in development mode:

1. Run `npm start` in the project directory.
2. Sideload the add-in in Outlook using the manifest file in the `appPackage` folder.

## Troubleshooting

- If you encounter any issues with API calls, make sure your Syncro API key and URL are correct in the settings.
- Check the browser console for any error messages that may help identify the problem.

## Support

For any questions or issues, please open an issue in this repository or contact your system administrator.
