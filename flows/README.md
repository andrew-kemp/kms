# KMS Document Approval Flow

This Power Automate flow provides automated approval functionality for documents in the KMS SharePoint document library.

## Overview

The KMS Approval Flow automates the document approval process by:
1. Monitoring the KMS document library for new or modified files
2. Retrieving file metadata, including the "Approvers" field
3. Starting an approval process assigned to the specified email addresses
4. Updating the document's "Approval Status" field based on the outcome
5. Adding rejection comments when applicable

## Flow Components

The `KMSApproval.zip` file contains the complete Power Automate flow definition with the following components:
- **definition.json**: Main flow logic and workflow definition
- **metadata.json**: Flow metadata and configuration
- **connections.json**: Required connector definitions (SharePoint and Approvals)
- **manifest.json**: Package manifest for import

## Prerequisites

Before importing and setting up this flow, ensure you have:

### SharePoint Library Setup
1. A SharePoint document library named "KMS Document Library" (or modify the flow to match your library name)
2. The following custom columns in your document library:
   - **Approvers** (Person or Group, multiple selection enabled)
   - **Approval Status** (Choice field with options: Pending, Approved, Rejected)
   - **Approval Comments** (Multiple lines of text)

### Required Permissions
- SharePoint site access with read/write permissions to the document library
- Power Automate license that supports premium connectors
- Approvals connector permissions

## Setup Instructions

### Step 1: Import the Flow
1. Download the `KMSApproval.zip` file from this repository
2. Open Power Automate (https://flow.microsoft.com)
3. Navigate to "My flows" → "Import" → "Import package (legacy)"
4. Upload the `KMSApproval.zip` file
5. Review import settings and click "Import"

### Step 2: Configure Connections
After import, you'll need to configure the required connections:

1. **SharePoint Connection**:
   - Click "Create new" for the SharePoint connection
   - Sign in with an account that has access to your SharePoint site
   - Authorize the connection

2. **Approvals Connection**:
   - Click "Create new" for the Approvals connection
   - Sign in with your Power Platform account
   - Authorize the connection

### Step 3: Update Flow Configuration
1. Open the imported flow for editing
2. Update the following parameters in each SharePoint action:
   - Replace `https://your-tenant.sharepoint.com/sites/your-site` with your actual SharePoint site URL
   - Verify the library name matches your document library (default: "KMS Document Library")

### Step 4: Test the Flow
1. Save and enable the flow
2. Upload a test document to your SharePoint library
3. Set the "Approvers" field to include one or more email addresses
4. Verify the approval request is sent and the status updates correctly

## Flow Logic Details

### Trigger: File Created or Modified
- **Type**: SharePoint trigger
- **Frequency**: Checks every minute for new or modified files
- **Scope**: KMS Document Library

### Action 1: Get File Metadata
- Retrieves all metadata for the triggered file
- Extracts the "Approvers" field for routing

### Action 2: Conditional Check
- Verifies that the "Approvers" field is not empty
- If no approvers are specified, the flow terminates with a cancellation status

### Action 3: Start Approval Process
- Creates an approval request with:
  - Document title and link
  - Creator and modification details
  - Assignment to specified approvers
- Waits for approval response

### Action 4: Update Approval Status
Based on the approval outcome:
- **If Approved**: Sets "Approval Status" to "Approved"
- **If Rejected**: Sets "Approval Status" to "Rejected" and adds rejection comments

## Customization Options

### Modifying the Approval Request
To customize the approval request content, edit the "Start an approval" action:
- **Title**: Modify the email subject line
- **Details**: Update the request body content
- **Enable notifications**: Configure email notifications
- **Enable reassignment**: Allow approvers to reassign requests

### Adding Additional Actions
Common enhancements include:
- Sending notification emails to document creators
- Creating audit logs in a separate list
- Integrating with Microsoft Teams for notifications
- Adding escalation logic for overdue approvals

### Field Mappings
If your SharePoint library uses different field names:
1. Update the field references in the flow actions
2. Ensure data types match the flow expectations
3. Test thoroughly after modifications

## Troubleshooting

### Common Issues

**Flow not triggering:**
- Verify the SharePoint connection is active
- Check that the site URL and library name are correct
- Ensure the flow is enabled

**Approval emails not sent:**
- Verify the Approvals connection is authorized
- Check that email addresses in the "Approvers" field are valid
- Confirm that approvers have access to Power Automate

**Status not updating:**
- Verify SharePoint write permissions
- Check that the "Approval Status" field exists and is the correct type
- Review flow run history for specific error messages

### Monitoring and Maintenance
- Regularly check the flow run history for errors
- Monitor approval response times and consider adding escalation logic
- Update connections when credentials change
- Review and update field mappings when SharePoint schema changes

## Security Considerations

- Approvers receive emails with document links; ensure they have appropriate SharePoint access
- The flow runs with the connection owner's permissions
- Consider implementing additional access controls for sensitive documents
- Regularly review and audit approval logs

## Support

For issues with this flow:
1. Check the Power Automate run history for detailed error messages
2. Verify all prerequisites are met
3. Test with a simple document first
4. Contact your SharePoint administrator for permission-related issues

## Version History

- **v1.0**: Initial release with basic approval functionality
- Support for multiple approvers
- Automatic status updates
- Rejection comment capture

---

**Note**: This flow requires appropriate licensing for Power Automate premium connectors. Contact your IT administrator for licensing questions.