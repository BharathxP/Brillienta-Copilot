# Set up Brillienta Copilot (for a customer's Microsoft 365 admin)

About 15 minutes. You need a SharePoint site owner for steps 1–2 and a Microsoft 365 admin for step 4.

## 1. Install the setup page

1. Upload `Brillienta-Copilot.sppkg` to the tenant App Catalog and deploy it.
2. On any site (for example the IT site), add a page with the **Brillienta Copilot** web part.

## 2. Find your apps

The page scans the current site. Add every other site that holds a Brillienta app (Helpdesk, Asset Management,
Employee Directory). Each app shows a green tick when it is ready.

If Helpdesk shows a warning, open Helpdesk › Settings › Copilot agent and switch on **Process Copilot requests**.
Without it, tickets and changes made in Copilot are saved but never routed, numbered or applied.

## 3. Sign-in

Until Brillienta's central registration is available, use your own (the same one the Helpdesk agent uses works):

1. Entra admin center › App registrations › New registration, multitenant, redirect URI (Web)
   `https://teams.microsoft.com/api/platform/v1.0/oAuthRedirect`.
2. Microsoft Graph delegated permissions: `Sites.ReadWrite.All`, `User.Read`, `offline_access`. Grant admin consent.
3. Create a client secret.
4. Teams Developer Portal › Tools › OAuth client registration: base URL `https://graph.microsoft.com/v1.0`,
   the app's client id and secret, authorization and token endpoints
   `https://login.microsoftonline.com/<tenant>/oauth2/v2.0/authorize` and `/token`, the three scopes above.
5. Paste the **OAuth client registration ID** in step 3 of the setup page.

## 4. Download and upload

1. Select **Download agent package**.
2. Microsoft 365 admin center › Settings › Integrated apps › Upload custom apps › upload the zip, and deploy it to
   the people who should use it.
3. Users find the agent in Microsoft 365 Copilot and Teams. The first time it reads data, it asks them to sign in.

Upload a new package whenever you add an app or rename the agent. It keeps the same app id, so it updates in place.

## Who sees what

The agent signs in as each person and reads SharePoint with their permissions. Someone who cannot open the Assets
list in SharePoint cannot see assets through Copilot either. Helpdesk requesters see only their own tickets, and
internal notes only reach helpdesk agents.

Users without a Microsoft 365 Copilot licence can use the agent in Copilot Chat. Reading company data then uses
Copilot Credits billed to your organisation. To avoid that, deploy the agent only to licensed users.
