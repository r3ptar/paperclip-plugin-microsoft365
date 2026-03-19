# Microsoft 365 Plugin -- Setup Guide

This guide walks you through installing and configuring the Paperclip Microsoft 365 plugin. By the end, you will have a working integration between Paperclip and one or more of these Microsoft 365 services:

- **Planner** -- Bidirectional sync between Paperclip issues and Microsoft Planner tasks
- **SharePoint** -- Search, read, and upload documents in SharePoint document libraries
- **Outlook** -- Calendar events for issue deadlines and daily email digest summaries

---

## Table of Contents

1. [Prerequisites](#1-prerequisites)
2. [Create an Azure AD App Registration](#2-create-an-azure-ad-app-registration)
3. [Install the Plugin](#3-install-the-plugin)
4. [Setup Wizard Walkthrough](#4-setup-wizard-walkthrough)
5. [Conflict Resolution Strategies](#5-conflict-resolution-strategies)
6. [After Setup](#6-after-setup)
7. [Troubleshooting](#7-troubleshooting)

---

## 1. Prerequisites

Before you begin, make sure you have:

- **Azure AD (Entra ID) admin access** -- You need permission to create app registrations and grant admin consent for API permissions in your organization's Azure portal.
- **A running Paperclip instance** -- The Microsoft 365 plugin must be available in your Paperclip plugin marketplace.

No redirect URI or user-interactive login is required. The plugin authenticates using the OAuth 2.0 Client Credentials flow, which means it runs as a background service with application-level permissions.

---

## 2. Create an Azure AD App Registration

This is the one manual step you must complete in the Azure portal before the plugin can connect to your Microsoft 365 tenant.

### 2.1 Register the Application

1. Open the [Azure Portal](https://portal.azure.com) and navigate to **Azure Active Directory** (also called Microsoft Entra ID).
2. In the left sidebar, select **App registrations**, then click **New registration**.
3. Fill in the registration form:
   - **Name**: Choose a descriptive name, for example `Paperclip M365 Integration`.
   - **Supported account types**: Select **Accounts in this organizational directory only** (single tenant).
   - **Redirect URI**: Leave this blank. The plugin uses client credentials and does not need a redirect URI.
4. Click **Register**.

### 2.2 Note Your Identifiers

After the app is created, you will land on its **Overview** page. Copy these two values -- you will need them during the Setup Wizard:

| Field | Where to find it |
|---|---|
| **Application (client) ID** | Shown on the app's Overview page |
| **Directory (tenant) ID** | Shown on the app's Overview page |

### 2.3 Create a Client Secret

1. In the app's left sidebar, select **Certificates & secrets**.
2. Under **Client secrets**, click **New client secret**.
3. Enter a description (for example, `Paperclip plugin`) and choose an expiration period.
4. Click **Add**.
5. **Copy the secret value immediately.** It will not be shown again after you leave this page.

Store this secret securely. You will provide it as a secret reference during setup.

### 2.4 Add API Permissions

1. In the app's left sidebar, select **API permissions**.
2. Click **Add a permission** and choose **Microsoft Graph**.
3. Select **Application permissions** (not Delegated permissions).
4. Add the permissions for the services you plan to use:

**Planner**

| Permission | Purpose |
|---|---|
| `Tasks.ReadWrite.All` | Create, read, and update Planner tasks |
| `Group.Read.All` | List M365 groups and their associated plans |

**SharePoint**

| Permission | Purpose |
|---|---|
| `Sites.Read.All` | List and search SharePoint sites and documents |
| `Files.ReadWrite.All` | Read and upload files to document libraries |

**Outlook**

| Permission | Purpose |
|---|---|
| `Calendars.ReadWrite` | Create and update calendar events for issue deadlines |
| `Mail.Send` | Send daily digest emails to configured recipients |

> **Tip:** You only need to add the permissions for services you actually plan to enable. For example, if you only want Planner sync, add `Tasks.ReadWrite.All` and `Group.Read.All`.

### 2.5 Grant Admin Consent

After adding your permissions, click **Grant admin consent for [your organization]** at the top of the API permissions page. Confirm the prompt. Each permission should show a green checkmark in the **Status** column after consent is granted.

---

## 3. Install the Plugin

1. In your Paperclip instance, open the plugin marketplace.
2. Find **Microsoft 365** and install it.
3. Once installed, open the plugin's **Settings** page. Because this is a fresh installation with no saved configuration, the **Setup Wizard** will appear automatically.

If you are returning to configure the plugin after a previous setup, you can launch the wizard again by clicking the **Reconfigure** button on the Settings page.

---

## 4. Setup Wizard Walkthrough

The Setup Wizard guides you through configuration one step at a time. The wizard adapts based on your choices -- steps for services you do not enable are skipped automatically.

### Step 1: Azure AD Connection

Enter the credentials from the app registration you created earlier:

| Field | Description |
|---|---|
| **Tenant ID** | The Directory (tenant) ID from your app's Overview page. Must be a valid UUID. |
| **Client ID** | The Application (client) ID from your app's Overview page. |
| **Client Secret Reference** | A Paperclip secret reference pointing to the client secret you created. The format is typically `secret-ref://...`. |

After filling in all three fields, click **Test Connection**. The plugin will attempt to acquire an OAuth token from Azure AD. You must get a successful connection before you can proceed.

If the test fails, double-check that:
- The Tenant ID and Client ID are correct.
- The client secret has not expired.
- Admin consent has been granted for the API permissions.

### Step 2: Choose Services

Toggle on the Microsoft 365 services you want to integrate:

- **Planner** -- Sync Paperclip issues with Microsoft Planner tasks. Create, update, and reconcile tasks bidirectionally.
- **SharePoint** -- Search and upload documents to SharePoint document libraries. Attach files directly from agent conversations.
- **Outlook** -- Create calendar events for issue deadlines and send email digest summaries to your team.

You must enable at least one service to proceed. The wizard will add configuration steps only for the services you enable.

### Step 3: Planner Configuration

*This step appears only if you enabled Planner.*

| Field | Description |
|---|---|
| **Microsoft 365 Group** | Select the M365 Group that owns the Planner plan you want to sync with. The dropdown is populated from your tenant. |
| **Planner Plan** | Select the plan within the chosen group. The dropdown updates when you change the group. |
| **Conflict Strategy** | Choose how to resolve conflicts when both Paperclip and Planner have changed the same task. See [Conflict Resolution Strategies](#5-conflict-resolution-strategies) below. The default is **Last Write Wins**. |

### Step 4: SharePoint Configuration

*This step appears only if you enabled SharePoint.*

| Field | Description |
|---|---|
| **SharePoint Site** | Select the SharePoint site to use for document search and upload. |
| **Document Library (Drive)** | Select the document library within the chosen site. The dropdown updates when you change the site. |
| **Upload Folder** (optional) | Select a specific folder within the library where uploaded files will be stored. If left empty, uploads go to the library root. |

### Step 5: Outlook Configuration

*This step appears only if you enabled Outlook.*

| Field | Description |
|---|---|
| **Sender User ID** | The email address or User Principal Name (UPN) of the user whose calendar will hold deadline events and who will send digest emails. For example, `service-account@yourtenant.com`. |
| **Calendar** | Select which of that user's calendars to use for deadline events. The dropdown populates after you enter the Sender User ID. |
| **Digest Recipients** | Add email addresses of the people who should receive the daily digest. Type an address and press Enter to add it. |

### Step 6: Review and Save

The final step shows a summary of everything you configured:

- Azure AD connection details
- Which services are enabled
- Planner group, plan, and conflict strategy (if applicable)
- SharePoint site, library, and upload folder (if applicable)
- Outlook sender, calendar, and digest recipients (if applicable)

Review the summary and click **Save & Activate**. The plugin will validate your configuration, save it, and initialize the enabled services. If there are validation errors, they will appear on this screen with guidance on what to fix.

---

## 5. Conflict Resolution Strategies

When Planner sync is enabled, the plugin performs a bidirectional reconciliation every 15 minutes. If the same task has been modified in both Paperclip and Planner since the last sync, that counts as a conflict. The conflict strategy you choose determines which version is kept.

| Strategy | Behavior |
|---|---|
| **Last Write Wins** | The most recently modified version takes priority, regardless of which system it came from. This is the default and is recommended for most teams. |
| **Paperclip Wins** | Paperclip's version always overwrites Planner when a conflict is detected. |
| **Planner Wins** | Planner's version always overwrites Paperclip when a conflict is detected. |

You can change the conflict strategy at any time from the Settings page without re-running the full wizard.

---

## 6. After Setup

Once configuration is saved and activated, the plugin begins working immediately.

### Dashboard Widget

A **M365 Sync Health** widget appears on your Paperclip dashboard. It shows:

- Connection status (Healthy / Unhealthy / Not configured)
- Number of tracked Planner tasks
- Time of the last reconciliation
- Which services are currently enabled

Use the **Refresh** button on the widget to get the latest status.

### Issue Detail Tab

Each issue gains a **Microsoft 365** tab in its detail view. This tab shows:

- The linked **Planner task** (if any), including its sync status and last sync time
- The linked **calendar event** (if any), including the due date

### Project Detail Tab

Projects show a **SharePoint** tab with guidance on using the `sharepoint-search` agent tool to find documents.

### Scheduled Jobs

The plugin runs several background jobs automatically:

| Job | Schedule | Description |
|---|---|---|
| Planner Reconciliation | Every 15 minutes | Full bidirectional sync between all tracked Paperclip issues and their linked Planner tasks. |
| Token Health Check | Every 30 minutes | Verifies that the OAuth credentials are still valid and can acquire tokens. |
| Graph Subscription Renewal | Every 12 hours | Renews Microsoft Graph webhook subscriptions so real-time Planner change notifications keep flowing. |
| Outlook Email Digest | Weekdays at 9:00 AM | Sends an HTML digest of the previous 24 hours of Paperclip issue activity to your configured recipients. |

### Agent Tools

When enabled, the plugin registers four tools that Paperclip agents can use:

- **sharepoint-search** -- Search documents in your configured SharePoint site by keyword
- **sharepoint-read** -- Read the text content of a specific SharePoint document
- **sharepoint-upload** -- Upload a file to your configured upload folder
- **planner-status** -- Check the linked Planner task status for a given Paperclip issue

### Re-running the Wizard

To change your configuration, open the plugin's **Settings** page. You can either:

- Edit individual fields directly on the Settings form and click **Save**.
- Click the **Reconfigure** button in the top-right corner to re-run the Setup Wizard from scratch.

---

## 7. Troubleshooting

### "Connection failed" during the Test Connection step

- Verify the **Tenant ID** and **Client ID** are correct. The Tenant ID must be a valid UUID.
- Confirm that the **client secret** has not expired in Azure AD.
- Check that **admin consent** has been granted for all required API permissions.

### "No groups found" or "No sites found" in the wizard dropdowns

- Verify the app registration has the correct **Application permissions** (not Delegated).
- Confirm that admin consent was granted. Look for green checkmarks next to each permission in the Azure portal.
- For groups, ensure the groups you expect are **Microsoft 365 Groups** (Unified groups), not security groups or distribution lists.

### Sync is not working or the dashboard shows "Unhealthy"

- Open the **M365 Sync Health** dashboard widget and click **Refresh** to see the current status.
- If the token health check is failing, your client secret may have expired. Create a new secret in Azure AD and update the secret reference in the plugin settings.
- Check that the Planner plan still exists and the app still has permission to access it.

### Digest emails are not being sent

- Verify that at least one **digest recipient** email address is configured.
- Confirm that the **Sender User ID** is a valid user in your tenant and the app has `Mail.Send` permission.
- The digest only runs on weekdays at 9:00 AM. If there was no issue activity in the previous 24 hours, the email will still be sent but will show "No recent activity."

### Calendar events are not being created

- Events are only created for issues that have a **due date** set.
- Verify the **Sender User ID** and **Calendar ID** are configured correctly.
- Confirm the app has `Calendars.ReadWrite` permission with admin consent.
