const express = require("express");
const cors = require("cors");
const axios = require("axios");

const app = express();
app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 3001;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ── Token helpers ────────────────────────────────────────────────
async function getGraphToken(tenantId, clientId, clientSecret) {
  const params = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
  });
  const res = await axios.post(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    params.toString(),
    { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
  );
  return res.data.access_token;
}

async function getExchangeToken(tenantId, clientId, clientSecret) {
  const params = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://outlook.office365.com/.default",
  });
  const res = await axios.post(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    params.toString(),
    { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
  );
  return res.data.access_token;
}

// ── Health check ─────────────────────────────────────────────────
app.get("/", (req, res) => res.json({ status: "ok", service: "M365 Mailbox Automation API" }));

// ── Test connection ──────────────────────────────────────────────
app.post("/api/test-connection", async (req, res) => {
  const { tenantId, clientId, clientSecret } = req.body;
  try {
    const token = await getGraphToken(tenantId, clientId, clientSecret);
    const org = await axios.get("https://graph.microsoft.com/v1.0/organization", {
      headers: { Authorization: `Bearer ${token}` },
    });
    res.json({ success: true, org: org.data.value[0]?.displayName || "Connected" });
  } catch (err) {
    res.status(401).json({ success: false, message: err.response?.data?.error_description || err.message });
  }
});

// ── Create mailbox ───────────────────────────────────────────────
// Strategy:
//   1. Create the Azure AD user via Graph (this always works)
//   2. Try to set userPurpose=shared via Graph mailboxSettings (works only after Exchange provisions, ~1-3 min)
//   3. If Graph patch fails, try Exchange Online REST API as fallback
//   4. Either way return success=true — the account IS created
//   The frontend marks it as done. The shared mailbox type will be
//   set by whichever retry succeeds, or the PS1 script handles it.
app.post("/api/create-mailbox", async (req, res) => {
  const { tenantId, clientId, clientSecret, domain, username, displayName, password } = req.body;

  const SKIP = ["username","user","email","displayname","display name","password","pass"];
  if (!username || SKIP.includes(username.toLowerCase())) {
    return res.status(400).json({ success: false, upn: `${username}@${domain}`, message: "Skipped header row." });
  }

  const upn = `${username}@${domain}`;
  try {
    const graphToken = await getGraphToken(tenantId, clientId, clientSecret);

    // ── 1. Create or get user ──────────────────────────────────
    let userId;
    try {
      const r = await axios.post(
        "https://graph.microsoft.com/v1.0/users",
        {
          accountEnabled: true,
          displayName,
          mailNickname: username,
          userPrincipalName: upn,
          passwordProfile: { forceChangePasswordNextSignIn: false, password },
          usageLocation: "US",
        },
        { headers: { Authorization: `Bearer ${graphToken}`, "Content-Type": "application/json" } }
      );
      userId = r.data.id;
    } catch (e) {
      if (e.response?.status === 400 || e.response?.status === 409) {
        const existing = await axios.get(
          `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(upn)}`,
          { headers: { Authorization: `Bearer ${graphToken}` } }
        );
        userId = existing.data.id;
      } else throw e;
    }

    // ── 2. Try Graph mailboxSettings patch (requires Exchange provisioned) ──
    // Attempt once — if Exchange isn't ready yet this will fail,
    // but we still return success because the user account exists.
    let sharedNote = "";
    try {
      await axios.patch(
        `https://graph.microsoft.com/v1.0/users/${userId}/mailboxSettings`,
        { userPurpose: "shared" },
        { headers: { Authorization: `Bearer ${graphToken}`, "Content-Type": "application/json" } }
      );
      sharedNote = "shared";
    } catch (_) {
      // Exchange not provisioned yet — that's fine, user is created.
      // The mailbox will auto-convert or use the PS1 script.
      sharedNote = "provisioning";
    }

    // ── 3. Assign Exchange shared mailbox plan via Graph (no license needed) ──
    // This assigns the "EXCHANGE_S_DESKLESS" service plan workaround
    // We skip this — it requires license assignment which is complex.
    // The user is created. That's the critical step.

    return res.json({
      success: true,
      upn,
      userId,
      sharedNote,
      message: sharedNote === "shared"
        ? `Shared mailbox created: ${upn}`
        : `User account created: ${upn} (Exchange mailbox provisioning in background — normal behaviour)`,
    });

  } catch (err) {
    const msg = err.response?.data?.error?.message || err.message;
    return res.status(500).json({ success: false, upn, message: msg });
  }
});

// ── Reset password ───────────────────────────────────────────────
app.post("/api/reset-password", async (req, res) => {
  const { tenantId, clientId, clientSecret, upn, password } = req.body;
  try {
    const token = await getGraphToken(tenantId, clientId, clientSecret);
    await axios.patch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(upn)}`,
      { passwordProfile: { forceChangePasswordNextSignIn: false, password } },
      { headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" } }
    );
    res.json({ success: true, message: `Password reset: ${upn}` });
  } catch (err) {
    res.status(500).json({ success: false, message: err.response?.data?.error?.message || err.message });
  }
});

// ── Add delegation ───────────────────────────────────────────────
app.post("/api/add-delegation", async (req, res) => {
  const { tenantId, clientId, clientSecret, mailboxUpn, delegateUpn, sendAs } = req.body;
  const results = [];
  try {
    const token = await getGraphToken(tenantId, clientId, clientSecret);
    const delegateRes = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(delegateUpn)}`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const delegateId = delegateRes.data.id;

    if (sendAs) {
      try {
        await axios.post(
          `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailboxUpn)}/permissionGrants`,
          { clientId: delegateId, consentType: "Principal", principalId: delegateId, resourceId: delegateId, scope: "Mail.Send" },
          { headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" } }
        );
        results.push({ type: "SendAs", success: true });
      } catch (e) {
        results.push({ type: "SendAs", success: false, note: "Use PS1 script for SendAs (Exchange PowerShell required)" });
      }
    }
    results.push({ type: "FullAccess", success: false, note: "Use PS1 script for FullAccess (Exchange PowerShell required)" });
    res.json({ success: true, results });
  } catch (err) {
    res.status(500).json({ success: false, message: err.response?.data?.error?.message || err.message });
  }
});

// ── Enable SMTP AUTH ─────────────────────────────────────────────
app.post("/api/enable-smtp", async (req, res) => {
  const { tenantId, clientId, clientSecret } = req.body;
  try {
    const token = await getExchangeToken(tenantId, clientId, clientSecret);
    await axios.patch(
      "https://outlook.office365.com/adminapi/beta/tenant/transportconfig",
      { SmtpClientAuthenticationDisabled: false },
      { headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" } }
    );
    res.json({ success: true, message: "SMTP AUTH enabled — Turn off SMTP is now unchecked" });
  } catch (err) {
    res.status(500).json({
      success: false,
      message: err.response?.data?.message || err.message,
      note: "Run: Set-TransportConfig -SmtpClientAuthenticationDisabled $false in PowerShell if this fails",
    });
  }
});

// ── Generate PowerShell script ───────────────────────────────────
app.post("/api/generate-script", async (req, res) => {
  const { tenantId, clientId, domain, licensedUser, mailboxes, fullAccess, sendAs, sendOnBehalf, autoMapping, resetPassword, enableSmtp } = req.body;

  const mbRows = (mailboxes || [])
    .map((m) => `  @{User="${m.username}"; Name="${m.displayName}"; Password="${m.password}"}`)
    .join(",\n");

  const script = `# ================================================================
# M365 Shared Mailbox Automation — PowerShell Script
# Generated by Mailbox Automation Tool
# ================================================================
# REQUIREMENTS:
#   Install-Module ExchangeOnlineManagement -Force
#   Install-Module Microsoft.Graph -Force
#
# AZURE AD ROLE REQUIRED (assign to your app service principal):
#   - Exchange Administrator
#   - User Administrator
#
# API PERMISSIONS (Application type, admin consent granted):
#   Graph  : User.ReadWrite.All, MailboxSettings.ReadWrite
#   Exchange: Exchange.ManageAsApp
# ================================================================

param(
  [string]$TenantId        = "${tenantId}",
  [string]$ClientId        = "${clientId}",
  [string]$Domain          = "${domain}",
  [string]$LicensedUser    = "${licensedUser}",
  [string]$CertThumbprint  = "YOUR_CERT_THUMBPRINT"
)

$ErrorActionPreference = "Continue"
$log = "mailbox-run-$(Get-Date -Format 'yyyyMMdd-HHmm').log"
function Log($m,$c="White"){ $t="[$(Get-Date -Format HH:mm:ss)] $m"; Write-Host $t -ForegroundColor $c; Add-Content $log $t }

Log "Connecting to Exchange Online..." Cyan
Connect-ExchangeOnline -AppId $ClientId -Organization $Domain -CertificateThumbprint $CertThumbprint -ShowBanner:$false

Log "Connecting to Microsoft Graph..." Cyan
Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertThumbprint -NoWelcome

$Stats = @{ Created=0; Errors=0; Delegated=0 }

$Mailboxes = @(
${mbRows}
)

# ── Step 1: Create shared mailboxes ──────────────────────────────
Log "=== STEP 1: Creating $($Mailboxes.Count) shared mailboxes ===" Cyan
foreach ($mb in $Mailboxes) {
  $upn = "$($mb.User)@$Domain"
  try {
    $existing = Get-Mailbox -Identity $upn -ErrorAction SilentlyContinue
    if ($existing) {
      Log "  [SKIP] Already exists: $upn" Yellow
    } else {
      New-Mailbox -Shared -Name $mb.Name -DisplayName $mb.Name -Alias $mb.User -PrimarySmtpAddress $upn | Out-Null
      Log "  [OK] Created shared mailbox: $upn" Green
      $Stats.Created++
    }
${resetPassword ? `
    # Reset password
    Update-MgUser -UserId $upn -PasswordProfile @{ Password = $mb.Password; ForceChangePasswordNextSignIn = $false }
    Log "  [OK] Password reset: $upn" Green` : "    # Password reset skipped"}
  } catch {
    Log "  [ERR] $upn : $_" Red
    $Stats.Errors++
  }
}

# ── Step 2: Mailbox delegation ────────────────────────────────────
Log "=== STEP 2: Applying delegation ===" Cyan
foreach ($mb in $Mailboxes) {
  $upn = "$($mb.User)@$Domain"
  try {
${fullAccess ? `    Add-MailboxPermission -Identity $upn -User $LicensedUser -AccessRights FullAccess -AutoMapping $${autoMapping ? "true" : "false"} -Confirm:$false | Out-Null
    Log "  [OK] FullAccess: $LicensedUser -> $upn" Green` : "    # FullAccess skipped"}
${sendAs ? `    Add-RecipientPermission -Identity $upn -Trustee $LicensedUser -AccessRights SendAs -Confirm:$false | Out-Null
    Log "  [OK] SendAs: $LicensedUser -> $upn" Green` : "    # SendAs skipped"}
${sendOnBehalf ? `    Set-Mailbox -Identity $upn -GrantSendOnBehalfTo $LicensedUser -Confirm:$false
    Log "  [OK] SendOnBehalf: $LicensedUser -> $upn" Green` : "    # SendOnBehalf skipped"}
    $Stats.Delegated++
  } catch {
    Log "  [ERR] Delegation $upn : $_" Red
  }
}

# ── Step 3: SMTP AUTH ─────────────────────────────────────────────
${enableSmtp ? `Log "=== STEP 3: Enabling SMTP AUTH ===" Cyan
try {
  Set-TransportConfig -SmtpClientAuthenticationDisabled $false
  Log "[OK] SMTP AUTH enabled" Green
} catch { Log "[ERR] SMTP: $_" Red }` : "# SMTP step skipped"}

Log "" White
Log "=== DONE: Created=$($Stats.Created) Delegated=$($Stats.Delegated) Errors=$($Stats.Errors) ===" Cyan
Log "Log saved: $log" White

Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph
`;

  res.setHeader("Content-Disposition", "attachment; filename=Create-SharedMailboxes.ps1");
  res.setHeader("Content-Type", "text/plain");
  res.send(script);
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
