# Opsie Email Assistant - Outlook Add-in

This is the Outlook add-in version of the Opsie Email Assistant, providing AI-powered email productivity tools directly within Microsoft Outlook.

## Features

- **Email Summary**: Quickly get an AI-generated summary of email content with key points highlighted.
- **Contact History**: View your history with the sender, including previous exchanges and important information.
- **Smart Reply Generation**: Generate contextually appropriate email replies with customizable tone, length, and language.
- **Email Search**: Search through your past conversations with specific contacts to find relevant information.
- **Save & Organize**: Save important emails and organize them for future reference.

## Development Setup

### Prerequisites

- Microsoft Office/Outlook (desktop or web version)
- Node.js and npm for development

### Installation for Development

1. Clone this repository or download the files.
2. For local development testing, you'll need to configure Outlook to allow add-in sideloading:

   **Windows Registry Settings:**
   ```
   Windows Registry Editor Version 5.00
   
   [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer]
   "AllowUnsecureURLs"=dword:00000001
   "EnableSideLoadingKeys"=dword:00000001
   ```

3. In Outlook, go to Trust Center > Trust Center Settings > Manage Add-ins > Custom Add-ins > Add a custom add-in > Add from file
4. Browse to the `manifest.xml` file and follow the prompts to install the add-in.

### Project Structure

- `manifest.xml` - The add-in manifest file that defines metadata, capabilities, and host settings
- `src/` - Contains the source code for the add-in
  - `taskpane.html` - The main HTML interface for the add-in
  - `taskpane.js` - JavaScript code for the add-in functionality
  - `taskpane.css` - Styles for the add-in interface
- `assets/` - Contains icons and other static assets

## Deployment

For production deployment, you can:

1. Host the add-in files on a web server or content delivery network.
2. Update the `manifest.xml` file with the appropriate URLs for your hosted resources.
3. Deploy the add-in through one of these methods:
   - Microsoft AppSource submission for public distribution
   - Centralized deployment for organization-wide distribution
   - Office Add-in catalog for team or departmental distribution

## Customization

You can customize the add-in by:

1. Modifying the UI in `taskpane.html` and `taskpane.css`
2. Enhancing functionality in `taskpane.js`
3. Updating the manifest file to change capabilities or requirements

## API Integration

This add-in is designed to work with the Opsie Email Assistant API. To integrate with your own backend:

1. Replace the mock API functions in `taskpane.js` with actual API calls
2. Ensure your API endpoints are secure and properly handle authentication
3. Update the allowed domains in the manifest file if needed

## License

Copyright © 2023 Opsie AI Technologies

All rights reserved. This project and its contents are proprietary and confidential.

## Support

For support and feedback, contact support@opsie.io 