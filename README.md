# AzureRibbonTz

AzureRibbonTz is a C# VSTO (Visual Studio Tools for Office) plugin for Microsoft Outlook that seamlessly integrates your email workflow with Azure DevOps. It enables users to create, update, or comment on Azure DevOps work items (such as Bugs and User Stories) directly from emails, streamlining collaboration and issue tracking.

## Features

- **Create Azure DevOps Work Items**: Instantly open new Bugs or User Stories from selected Outlook emails.
- **Update Existing Work Items**: Link an email to an existing Azure DevOps work item and add comments or updates using the email context.
- **Custom Ribbon Controls in Outlook**: User-friendly ribbon buttons for quick access to Azure DevOps actions.
- **Azure DevOps Integration**: Requires Personal Access Token (PAT), Azure DevOps URL, project name, and a default assignee for work item creation.
- **Automatic Context Linking**: Automatically attaches relevant email content to Azure DevOps items for better traceability.

## Prerequisites

- Windows OS with Microsoft Outlook 2013 or newer
- [.NET Framework](https://dotnet.microsoft.com/download/dotnet-framework) 4.7.2 or higher
- [Visual Studio](https://visualstudio.microsoft.com/) (for building from source)
- Azure DevOps account and project
- Azure DevOps Personal Access Token (PAT) with work item access

## Installation

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/ctsahic/AzureRibbonTz.git
   ```
2. **Open the Solution in Visual Studio**:
   - Open `AzureRibbonTz.sln`.
3. **Build the Project**:
   - Go to `Build > Build Solution`.
4. **Deploy the Add-in**:
   - For development, press `F5` to launch Outlook with the add-in loaded.
   - For production, publish the solution and follow [Microsoft’s VSTO deployment guide](https://docs.microsoft.com/en-us/visualstudio/vsto/deploying-office-solutions?view=vs-2022).

## Configuration

Before using the add-in, configure the following settings (prompted on first use or in the add-in options):

- **Azure DevOps Personal Access Token (PAT):**  
  Obtain from [Azure DevOps User Settings > Personal Access Tokens](https://dev.azure.com/).
- **Azure DevOps Organization/URL:**  
  Example: `https://dev.azure.com/yourorganization`
- **Project Name:**  
  The name of the Azure DevOps project to use.
- **Default Assignee:**  
  Azure DevOps username or email to assign new items to by default.

## Usage

1. **Create Work Item from Email:**
   - Select an email in Outlook.
   - Click the AzureRibbonTz ribbon button.
   - Choose to create a Bug or User Story.
   - Fill in any additional details and confirm.
2. **Update or Comment on Existing Work Item:**
   - With an email selected, choose the update option from the ribbon.
   - Enter the work item ID or select from recent items.
   - Add your comment or update, which will be attached to the item in Azure DevOps.

## Security

- Your PAT is stored securely and used only for Azure DevOps API calls.
- Never share your PAT publicly.

## Contributing

Contributions are welcome! To contribute:

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/my-feature`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/my-feature`)
5. Open a pull request

## License

Specify your license here (e.g., MIT, Apache 2.0). If none, write "All rights reserved".

---

> For questions, suggestions, or issues, please open an issue or contact [@ctsahic](https://github.com/ctsahic).