# Microsoft 365 Cross-Tenant Mailbox Migration

A step-by-step runbook for moving Exchange Online mailboxes from one Microsoft 365 tenant (source) to another (target) using the native cross-tenant mailbox move.

> **Official reference:** https://learn.microsoft.com/en-us/microsoft-365/migration/cross-tenant-mailbox-migration
> Portal labels, roles and cmdlet parameters change over time. Check the official documentation before each production run.

---

## Read this first

- This guide covers **native Exchange Online cross-tenant mailbox moves**. It does not cover Microsoft Entra cross-tenant synchronisation.
- A successful move **deletes the source mailbox** and leaves a MailUser in the source tenant. Treat it as a production change.
- Mailboxes on **any hold** (retention, litigation, eDiscovery) cannot be migrated.
- Migrated: email, contacts, calendar, tasks, notes, Recoverable Items.
- Not migrated: Teams chat folder content. Teams meeting URLs in migrated calendar items may be invalid and meetings may need to be recreated.
- Cross-tenant mailbox and calendar permissions are not preserved. Migrate delegates and related mailboxes in the same batch.
- Do not run any command with a placeholder unchanged. Always test with a pilot mailbox first.
- Never store client secrets in scripts, tickets, transcripts or source control.

## Overview

| Step | Where | What |
|------|-------|------|
| 1 | Both | Plan and check prerequisites |
| 2 | Admin PC | Install PowerShell module and connect |
| 3 | Target | Register app, create endpoint and organisation relationship |
| 4 | Source | Accept consent, create scope group, create organisation relationship |
| 5 | Target | Prepare MailUser objects |
| 6 | Target | Validate configuration |
| 7 | Admin PC | Create the migration CSV |
| 8 | Target | Run a pilot batch |
| 9 | Target | Run production batches |
| 10 | Both | Post-migration validation |
| 11 | Both | Rollback and recovery |
| 12 | Both | Cleanup |

---

## Step 1: Plan and prerequisites

- [ ] Record both tenants' tenant IDs (GUIDs), `onmicrosoft.com` names, accepted domains and admin contacts.
- [ ] Confirm both tenants use supported Microsoft 365 cloud environments with Exchange Online provisioned.
- [ ] Assign the **Cross Tenant User Data Migration** licence to each migrating user (one-time, per user). Target users also need an appropriate Exchange Online licence.
- [ ] Inventory every mailbox: primary SMTP, UPN, size, archive, aliases, X500 addresses, delegates, forwarding, inbox rules, holds, hybrid dependencies.
- [ ] Resolve holds only with approval from legal, compliance and the data owner.
- [ ] Decide the target UPN and SMTP domain. A custom domain can be verified in only one tenant at a time, so plan the domain move and mail routing separately.
- [ ] Choose a pilot group and a migration window. Notify users about Outlook and mobile re-authentication, MFA and application sign-in.
- [ ] Record a rollback decision before moving anything.
- [ ] For large projects, run Microsoft's validation script and consider a qualified migration partner.

**Roles:** use least privilege. Admins need rights to register and consent the app, create migration endpoints and organisation relationships, and manage migration batches (Move Mailboxes role).

---

## Step 2: Install the module and connect

Install/update from the official PowerShell Gallery (skip if already installed/updated):

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Update-Module ExchangeOnlineManagement
```

After installing, import the Exchange Online module in PowerShell:

```powershell
Import-Module ExchangeOnlineManagement
```

Check the execution policy:

```powershell
Get-ExecutionPolicy
```

If it returns `Restricted`, change it for the current session only:

```powershell
Set-ExecutionPolicy RemoteSigned -Scope Process
```

Connect to one tenant at a time:

```powershell
# Target
Connect-ExchangeOnline -UserPrincipalName admin@targettenant.onmicrosoft.com

# Source
Connect-ExchangeOnline -UserPrincipalName admin@sourcetenant.onmicrosoft.com
```

> **Note:** Make sure the admin accounts in both the target and source tenants have the Global Administrator role assigned. To verify:

1. Go to **Microsoft 365 admin center > Active users**
2. Select the admin user
3. Click the ... (more actions) menu > **Manage roles** 
4. In **Manage admin roles**, confirm the user has **Admin center access** enabled (i.e., an admin role such as Global Administrator is assigned)

Verify the tenant before every change:

```powershell
Get-OrganizationConfig | Format-List Name,Guid,IsDehydrated
Get-AcceptedDomain | Format-Table Name,DomainName,DomainType
```

### ⚠️ Verification Criteria (What to look for)

Before proceeding with any migration steps, you must confirm the following:

*   **`IsDehydrated` MUST equal `False`.**
    > **Note:** If it returns `True`, Exchange Online is not provisioned in this tenant. Stop the migration, assign an Exchange Online license, and wait for Microsoft to provision the mailboxes.

*   **`Guid` MUST match the Tenant ID listed in the Microsoft Entra Admin Center.**
    > **Note:** If the tenant is dehydrated (`True`), the GUID will be a temporary placeholder and will not match Entra.

*   **`Get-AcceptedDomain` MUST return all expected custom domains and the default `*.onmicrosoft.com` routing domain.** The `DomainType` for all active migration domains should be `Authoritative`.

*   **Admin Contacts MUST be identified and recorded.**
    > **How to check:**
    > 1. Log into the Microsoft 365 Admin Center for the tenant.
    > 2. Go to **Settings** > **Org settings** > **Organization profile** > **Organization information**.

Disconnect when finished:

```powershell
Disconnect-ExchangeOnline -Confirm:$false
```

---

## Step 3: Prepare the target (destination) tenant

### 3A. Register the migration app

1. Sign in to https://entra.microsoft.com using the **target** tenant.
2. Go to **App registrations > New registration**.
3. Name the app, for example `CrossTenantMailboxMigration`.
4. Under **Supported account types**, select **Multiple Entra ID tenants**.
6. Set the Web redirect URI to `https://office.com` and register.
7. Copy the **Application (client) ID**.
8. Go to **API permissions > Add a permission > APIs my organization uses > Office 365 Exchange Online > Application permissions > Mailbox > Mailbox.Migration**.
9. Create a **client secret**. Copy its value immediately, store it securely and note the expiry date.
10. Grant admin consent in the target tenant and confirm it completed.
11. Send the source admin this consent URL (replace every placeholder):

```
https://login.microsoftonline.com/SOURCE_TENANT.onmicrosoft.com/adminconsent?client_id=APP_CLIENT_ID&redirect_uri=https://office.com
```

10. The source admin must accept consent **before** the endpoint is created.

### 3B. Create the migration endpoint (target PowerShell)

```powershell
$AppId = "APP_CLIENT_ID"
$EndpointName = "CrossTenant-Source"
$RemoteTenant = "SOURCE_TENANT.onmicrosoft.com"

# Enter the secret value at the prompt. Do not save it in this file.
$Secret = Read-Host "Enter app client-secret VALUE" -AsSecureString
$Credential = [System.Management.Automation.PSCredential]::new($AppId, $Secret)

New-MigrationEndpoint -RemoteServer outlook.office.com -RemoteTenant $RemoteTenant -Credentials $Credential -ExchangeRemoteMove:$true -Name $EndpointName -ApplicationId $AppId
```

If this fails, check source-tenant consent, app ID, secret value and expiry, and tenant name.

### 3C. Create the organisation relationship (target PowerShell)

```powershell
$SourceTenantId = "SOURCE_TENANT_GUID"
$OrgRelName = "CrossTenant-Source"

$Existing = Get-OrganizationRelationship | Where-Object { $_.DomainNames -contains $SourceTenantId }
if ($Existing) {
    Set-OrganizationRelationship $Existing.Name -Enabled:$true -MailboxMoveEnabled:$true -MailboxMoveCapability Inbound
} else {
    New-OrganizationRelationship $OrgRelName -Enabled:$true -MailboxMoveEnabled:$true -MailboxMoveCapability Inbound -DomainNames $SourceTenantId
}
```

Review existing relationships first. Do not overwrite an unrelated one.

---

## Step 4: Prepare the source tenant

### 4A. Accept consent

Open the consent URL from Step 3A as a source tenant admin and accept. Verify under **Entra admin center > Enterprise applications**.

### 4B. Create the scope group (source Portal)

The group must be **mail-enabled** and a **security** group. It must contain only approved migrating users.

> **Path:** Exchange admin center > Recipients > Groups > Add a group > Mail-enabled security
> 
> **OR:** Microsoft 365 admin center > Teams & groups > Active teams & groups > Security groups > Add a mail-enabled security group

1. **Name it:** `CrossTenant-Migration-Scope`
2. **Add Owners:** (the Global Admin)
3. **Add the members:** (migrating users)
4. **Add Group email address:** `CrossTenant-Migration`
5. **Review and create group**

### 4C. Create the organisation relationship (source PowerShell)

```powershell
$TargetTenantId = "TARGET_TENANT_GUID"
$AppId = "APP_CLIENT_ID"
$Scope = "CrossTenant-Migration-Scope"
$OrgRelName = "CrossTenant-Target"

$Config = Get-OrganizationConfig
if ($Config.IsDehydrated) { Enable-OrganizationCustomization }

$Existing = Get-OrganizationRelationship | Where-Object { $_.DomainNames -contains $TargetTenantId }
if ($Existing) {
    Set-OrganizationRelationship $Existing.Name -Enabled:$true -MailboxMoveEnabled:$true -MailboxMoveCapability RemoteOutbound -OAuthApplicationId $AppId -MailboxMovePublishedScopes $Scope
} else {
    New-OrganizationRelationship $OrgRelName -Enabled:$true -MailboxMoveEnabled:$true -MailboxMoveCapability RemoteOutbound -DomainNames $TargetTenantId -OAuthApplicationId $AppId -MailboxMovePublishedScopes $Scope
}
```

Inspect the result:

```powershell
Get-OrganizationRelationship | Format-List Name,Enabled,DomainNames,MailboxMoveEnabled,MailboxMoveCapability,OAuthApplicationId,MailboxMovePublishedScopes
```

---

## Step 5: Prepare target MailUser objects (critical)

Each migrating user must exist in the **target** tenant as an Exchange Online **MailUser** (not a provisioned mailbox) with:

- `ExchangeGUID` copied from the source mailbox
- `ArchiveGUID` copied from the source, if the source has an archive
- `LegacyExchangeDN` and all source X500 addresses added as `x500:` proxy addresses
- Target `UserPrincipalName` and primary SMTP address on a target-verified domain
- `ExternalEmailAddress` pointing to the source routing address (usually `user@sourcetenant.onmicrosoft.com`)
- Correct Exchange Online and Cross Tenant User Data Migration licences

Get source mailbox details (**source** tenant):

```powershell
Get-Mailbox user@sourcedomain.com | Format-List DisplayName,UserPrincipalName,PrimarySmtpAddress,ExchangeGuid,ArchiveGuid,LegacyExchangeDN,EmailAddresses,ArchiveStatus,RecipientTypeDetails
```

Check the target object (**target** tenant):

```powershell
Get-Recipient user@targetdomain.com | Format-List Name,RecipientType,RecipientTypeDetails,PrimarySmtpAddress,ExternalEmailAddress
Get-MailUser user@targetdomain.com | Format-List Name,UserPrincipalName,PrimarySmtpAddress,ExternalEmailAddress,ExchangeGuid,ArchiveGuid,LegacyExchangeDN,EmailAddresses
```

**Warnings**

- Attribute preparation depends on how target identities are provisioned (cloud-only or directory-synchronised). Follow Microsoft's "Prepare target user objects" section. Do not stamp cloud attributes on synchronised objects.
- **Set up and validate the MailUser before assigning the Exchange licence.** Licensing first can create a new mailbox and cause a GUID mismatch.
- If a target object has a previous or soft-deleted mailbox with a mismatched GUID, stop and investigate. `Set-User -PermanentlyClearPreviousMailboxInfo` is irreversible and can make the old mailbox unrecoverable. Do not run it without verification and approval.
- Do not use SMTP addresses from a domain still attached to the source tenant. Keep old `LegacyExchangeDN` values as X500 addresses.

---

## Step 6: Validate before moving

On the **target** tenant:

```powershell
Get-MigrationEndpoint
Test-MigrationServerAvailability -Endpoint "CrossTenant-Source" -TestMailbox "user@targetdomain.com"
```

Then:

- Run Microsoft's cross-tenant migration validation script (linked from the main Microsoft Learn article) and fix every reported issue.
- Confirm each MailUser has the correct `ExchangeGUID`, `ArchiveGUID`, X500 addresses, primary SMTP, UPN and `ExternalEmailAddress`.
- Confirm each user is in the source scope group and licensing is correct in both tenants.

---

## Step 7: Create the migration CSV

Create `users.csv` with a single column. Values must be the **target** MailUser address, not the source address.

```csv
EmailAddress
user1@targetdomain.com
user2@targetdomain.com
```

Keep a copy of the exact CSV used for each batch, with the batch name and date.

---

## Step 8: Run a pilot batch (target tenant)

### Option 1: Portal

1. Exchange admin center (target) > **Migration > Batch > Add migration batch**.
2. Select **Migration to Exchange Online** and choose the cross-tenant migration type.
3. Select the endpoint `CrossTenant-Source`.
4. Upload `users.csv`.
5. Set the target delivery domain, notifications and start option. Choose **manual start** for the pilot.
6. Review and create the batch, then select it and click **Start**.

### Option 2: PowerShell (target)

```powershell
$BatchName = "CrossTenant-Pilot-2026-10-15"
$EndpointName = "CrossTenant-Source"
$TargetDeliveryDomain = "targettenant.onmicrosoft.com"
$CsvPath = "C:\Migration\users.csv"

New-MigrationBatch -Name $BatchName -SourceEndpoint $EndpointName `
    -CSVData ([System.IO.File]::ReadAllBytes($CsvPath)) -TargetDeliveryDomain $TargetDeliveryDomain
```

This creates the batch without starting it. Inspect it:

```powershell
Get-MigrationBatch -Identity $BatchName | Format-List *
Get-MigrationUser -BatchId $BatchName | Get-MigrationUserStatistics
```

Start it when approved:

```powershell
Start-MigrationBatch -Identity $BatchName
```

Microsoft also documents `-AutoStart` on `New-MigrationBatch`. Prefer create, inspect, then start for the first production use.

### Monitor

```powershell
Get-MigrationBatch -Identity $BatchName | Format-List Identity,Status,TotalCount,State
Get-MigrationUser -BatchId $BatchName
Get-MigrationUserStatistics -Identity user@targetdomain.com -IncludeReport | Format-List *
Get-MoveRequest -Flags "CrossTenant"
```

**Portal:** Exchange admin center > **Migration > Batch** > select the batch to see status per user and download the report.

Use the reports to fix root causes. Do not retry repeatedly without fixing the error.
---

## Step 9: Production batches

- Move a small pilot first, validate end to end, then proceed in controlled batches.
- Group delegates and related mailboxes together.
- Keep users and the help desk informed about sign-in changes and expected interruption.
- Monitor mail flow, migration status, client access and service health during cutover.
- For more than 10,000 users, Microsoft recommends multiple source mail-enabled security groups.
- Do not delete source MailUsers, organisation relationships, endpoints or the migration app until migration, mail routing and support checks are complete.

---

## Step 10: Post-migration validation

- [ ] Batch status is completed with no unresolved errors.
- [ ] Target mailbox has the expected primary SMTP, aliases, folders, recent and older mail, contacts, calendar, tasks, notes and archive.
- [ ] Sample message counts and date ranges match for critical folders.
- [ ] Inbound and outbound mail works, internal and external.
- [ ] Outlook on the web, desktop and mobile sign in correctly (MFA, profile, Autodiscover).
- [ ] Delegate access, shared mailboxes, send-as, send-on-behalf, forwarding, rules and calendar sharing are recreated where needed.
- [ ] No NDRs. Source MailUser `targetAddress` routing is correct.
- [ ] Hybrid on-premises objects and `targetAddress` are updated, if applicable.
- [ ] Teams meetings recreated where links are invalid. Teams chats are not migrated.
- [ ] Retention, eDiscovery, audit and compliance requirements confirmed with the compliance owner.
- [ ] Business owner sign-off obtained.

Useful read-only checks:

```powershell
Get-Mailbox user@targetdomain.com | Format-List DisplayName,PrimarySmtpAddress,ExchangeGuid,ArchiveGuid,ArchiveStatus,RecipientTypeDetails
Get-MigrationBatch -Identity "CrossTenant-Pilot-2026-10-15" | Format-List *
Get-MigrationUser -BatchId "CrossTenant-Pilot-2026-10-15"
Get-MoveRequest -Flags "CrossTenant"
```

---

## Step 11: Rollback and recovery

**Before cutover**

- If a batch has not completed, pause and investigate before making further changes.
- Use `Get-MigrationBatch`, `Get-MigrationUser` and detailed statistics to establish the state.
- Confirm the correct way to stop or remove the batch for its current state in the Microsoft documentation before any destructive cleanup.
- Keep source mailboxes, licences and routing intact until the pilot is signed off.

**After a successful move**

- The source mailbox is deleted. There is no simple undo.
- To move back, configure the tenants in the reverse direction: new migration app and consent in the new target, adjusted organisation relationships, prepared MailUsers and a new reverse batch.
- Treat a reverse move as a new, tested project. Keep the original tenant IDs, identities, routing records, CSVs, reports and approvals.
- If data is missing or the move failed, stop cleanup and contact Microsoft support or your migration partner. Do not run `-PermanentlyClearPreviousMailboxInfo` or delete user objects as a first step.

---

## Troubleshooting

| Problem | What to check |
|---------|---------------|
| `CrossTenantMigrationWithoutLicensePermanentException` | Cross Tenant User Data Migration licence assigned; target user has an Exchange Online licence |
| "Needs approval" or target licensing warnings | Verify MailUser attributes first; ExchangeGUID and proxy addresses must be correct before licensing |
| Authentication error creating the endpoint | Source consent accepted, client ID correct, secret value current, remote tenant correct |
| `Test-MigrationServerAvailability` fails | Endpoint, app permissions, consent in both tenants, organisation relationships, MailUser exists, tenant IDs, scope group |
| Mailbox not in scope or move not allowed | User is in the published security group; source relationship points to the target tenant GUID |
| ExchangeGUID or ArchiveGUID mismatch | Compare source and target values; correct using Microsoft's supported procedure; do not clear previous mailbox info without approval |
| NDR 5.1.1 on replies | Copy `LegacyExchangeDN` and all source X500 addresses to the target MailUser |
| Target SMTP address rejected or removed | Domain must be verified in the target tenant with no source-tenant conflict; licensing can remove unverified proxy addresses |
| Mail flow broken after move | Target delivery domain, source MailUser `ExternalEmailAddress`/`targetAddress`, accepted domains, connectors, DNS/MX, hybrid objects, message trace |
| Mailbox is on hold | Migration is blocked; work with compliance and legal on an approved path |
| Calendar or delegate permissions missing | Not supported as-is; recreate in the target and migrate related mailboxes together |
| Teams links invalid or chats missing | Known limitation; recreate meetings; chat content is not migrated |

---

## Step 12: Cleanup (only after sign-off)

- [ ] All batches completed, reports exported, mail flow stable, business owners signed off.
- [ ] Change records and reports retained according to policy.
- [ ] Migration endpoint and organisation relationships removed only if no further or reverse moves are expected.
- [ ] Migration app secret rotated or removed, and the app and consent removed, only when nothing depends on them.
- [ ] Remaining source MailUsers and routing reviewed; remove or convert them only when coexistence and mail flow allow.
- [ ] Documentation, support runbooks and user records updated.

---

## References

1. [Cross-tenant mailbox migration (main runbook)](https://learn.microsoft.com/en-us/microsoft-365/migration/cross-tenant-mailbox-migration)
2. [Cross-tenant mailbox migration (Enterprise documentation)](https://learn.microsoft.com/en-us/microsoft-365/enterprise/cross-tenant-mailbox-migration?view=o365-worldwide)
3. [Connect to Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell)
4. [New-MigrationBatch](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/new-migrationbatch)
5. [Test-MigrationServerAvailability](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/test-migrationserveravailability)
6. [Entra cross-tenant synchronisation overview](https://learn.microsoft.com/en-us/entra/identity/multi-tenant-organizations/cross-tenant-synchronization-overview) (not a mailbox migration tool)

## Licence and disclaimer

This runbook is provided as guidance only. Test in a non-production or pilot scope first and verify every step against current Microsoft documentation.
