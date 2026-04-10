const express = require("express");
const cors = require("cors");
const axios = require("axios");

const app = express();
app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 3001;

// ─── Helper: Get Graph API token ───────────────────────────────────────────
async function getGraphToken(tenantId, clientId, clientSecret) {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const params = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
  });
  const res = await axios.post(url, params.toString(), {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });
  return res.data.access_token;
}

// ─── Helper: Get Exchange Online token ─────────────────────────────────────
async function getExchangeToken(tenantId, clientId, clientSecret) {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const params = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://outlook.office365.com/.default",
  });
  const res = await axios.post(url, params.toString(), {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });
  return res.data.access_token;
}

// ─── Route: Test credentials ────────────────────────────────────────────────
app.post("/api/test-connection", async (req, res) => {
  const { tenantId, clientId, clientSecret } = req.body;
  try {
    const token = await getGraphToken(tenantId, clientId, clientSecret);
    const org = await axios.get("https://graph.microsoft.com/v1.0/organization", {
      headers: { Authorization: `Bearer ${token}` },
    });
    res.json({
      success: true,
      org: org.data.value[0]?.displayName || "Connected",
      message: "Token acquired successfully",
    });
  } catch (err) {
    res.status(401).json({
      success: false,
      message: err.response?.data?.error_description || err.message,
    });
  }
});

// ─── Helper: sleep ─────────────────────────────────────────────────────────
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ─── Helper: retry with backoff ────────────────────────────────────────────
async function withRetry(fn, retries = 4, delayMs = 5000) {
  for (let i = 0; i < retries; i++) {
    try {
      return await fn();
    } catch (e) {
      const isLast = i === retries - 1;
      if (isLast) throw e;
      await sleep(delayMs);
    }
  }
}

// ─── Route: Create a single shared mailbox via Graph ───────────────────────
app.post("/api/create-mailbox", async (req, res) => {
  const { tenantId, clientId, clientSecret, domain, username, displayName, password } = req.body;

  // Guard: skip if username looks like a CSV header row
  if (!username || username.toLowerCase() === "username" || username.toLowerCase() === "user") {
    return res.status(400).json({ success: false, upn: `${username}@${domain}`, message: "Skipped — looks like a CSV header row." });
  }

  try {
    const token = await getGraphToken(tenantId, clientId, clientSecret);
    const upn = `${username}@${domain}`;

    // 1. Create the user account
    let userId;
    const userPayload = {
      accountEnabled: true,
      displayName: displayName,
      mailNickname: username,
      userPrincipalName: upn,
      passwordProfile: {
        forceChangePasswordNextSignIn: false,
        password: password,
      },
      usageLocation: "US",
    };

    try {
      const createUser = await axios.post(
        "https://graph.microsoft.com/v1.0/users",
        userPayload,
        { headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" } }
      );
      userId = createUser.data.id;
    } catch (e) {
      // User already exists — fetch their ID
      if (e.response?.status === 400 || e.response?.status === 409) {
        const existing = await axios.get(
          `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(upn)}`,
          { headers: { Authorization: `Bearer ${token}` } }
        );
        userId = existing.data.id;
      } else throw e;
    }

    // 2. Wait for Exchange to provision the mailbox (takes 10–30s after user creation)
    //    Retry patching mailboxSettings up to 4 times with 8s gaps
    await withRetry(async () => {
      await axios.patch(
        `https://graph.microsoft.com/v1.0/users/${userId}/mailboxSettings`,
        { userPurpose: "shared" },
        { headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" } }
      );
    }, 4, 8000);

    res.json({ success: true, upn, userId, message: `Shared mailbox created: ${upn}` });
  } catch (err) {
    const msg = err.response?.data?.error?.message || err.message;
    res.status(500).json({ success: false, upn: `${username}@${domain}`, message: msg });
  }
});

// ─── Route: Reset password ──────────────────────────────────────────────────
app.post("/api/reset-password", async (req, res) => {
  const { tenantId, clientId, clientSecret, upn, password } = req.body;
  try {
    const token = await getGraphToken(tenantId, clientId, clientSecret);
    await axios.patch(
      `https://graph.microsoft.com/v1.0/users/${upn}`,
      {
        passwordProfile: {
          forceChangePasswordNextSignIn: false,
          password: password,
        },
      },
      { headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" } }
    );
    res.json({ success: true, message: `Password reset: ${upn}` });
  } catch (err) {
    res.status(500).json({ success: false, message: err.response?.data?.error?.message || err.message });
  }
});

// ─── Route: Add mailbox delegation (Full Access + Send As) ─────────────────
// NOTE: Full delegation requires Exchange Online PowerShell (Add-MailboxPermission).
// Via Graph we can set SendAs via directory permissions.
app.post("/api/add-delegation", async (req, res) => {
  const { tenantId, clientId, clientSecret, mailboxUpn, delegateUpn, sendAs, fullAccess } = req.body;
  const results = [];

  try {
    const token = await getGraphToken(tenantId, clientId, clientSecret);

    // Get delegate user ID
    const delegateRes = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${delegateUpn}`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const delegateId = delegateRes.data.id;

    // Send As — via Graph serviceProvisioningErrors workaround using
    // Exchange REST endpoint
    if (sendAs) {
      try {
        await axios.post(
          `https://graph.microsoft.com/v1.0/users/${mailboxUpn}/permissionGrants`,
          {
            clientId: delegateId,
            consentType: "Principal",
            principalId: delegateId,
            resourceId: delegateId,
            scope: "Mail.Send",
          },
          { headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" } }
        );
        results.push({ type: "SendAs", success: true });
      } catch (e) {
        results.push({ type: "SendAs", success: false, note: "Use PowerShell fallback for SendAs" });
      }
    }

    // Full Access note: Graph API does not expose Add-MailboxPermission directly.
    // This is handled via the PowerShell script download. We log it here.
    if (fullAccess) {
      results.push({
        type: "FullAccess",
        success: false,
        note: "FullAccess requires Exchange PowerShell — use the downloaded .ps1 script",
      });
    }

    res.json({ success: true, mailboxUpn, delegateUpn, results });
  } catch (err) {
    res.status(500).json({ success: false, message: err.response?.data?.error?.message || err.message });
  }
});

// ─── Route: Enable SMTP AUTH org-wide ──────────────────────────────────────
// Graph does not expose Set-TransportConfig directly.
// We use the Exchange REST API for org settings.
app.post("/api/enable-smtp", async (req, res) => {
  const { tenantId, clientId, clientSecret } = req.body;
  try {
    const token = await getExchangeToken(tenantId, clientId, clientSecret);

    // Exchange Online REST — organization config
    await axios.patch(
      "https://outlook.office365.com/adminapi/beta/tenant/transportconfig",
      { SmtpClientAuthenticationDisabled: false },
      {
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
          "X-AnchorMailbox": "app@" + tenantId,
        },
      }
    );

    res.json({ success: true, message: "SMTP AUTH enabled — Turn off SMTP is now unchecked" });
  } catch (err) {
    // Fallback message if endpoint requires additional scope
    res.status(500).json({
      success: false,
      message: err.response?.data?.message || err.message,
      note: "If this fails, run Set-TransportConfig -SmtpClientAuthenticationDisabled $false in PowerShell",
    });
  }
});

// ─── Route: Generate PowerShell script ─────────────────────────────────────
app.post("/api/generate-script", async (req, res) => {
  const {
    tenantId, clientId, domain, licensedUser,
    mailboxes, fullAccess, sendAs, sendOnBehalf,
    autoMapping, resetPassword, enableSmtp,
  } = req.body;

  const mbRows = mailboxes
    .map((m) => `  @{User="${m.username}"; Name="${m.displayName}"; Password="${m.password}"}`)
    .join(",\n");

  const script = `# ================================================================
# M365 Shared Mailbox Automation Script
# Generated by Mailbox Automation Tool
# Deploy backend: Render  |  Frontend: Netlify
# ================================================================
# REQUIREMENTS:
#   Install-Module ExchangeOnlineManagement -Force
#   Install-Module Microsoft.Graph -Force
#
# APP REGISTRATION PERMISSIONS NEEDED:
#   Graph  : User.ReadWrite.All, Mail.ReadWrite,
#             MailboxSettings.ReadWrite, Directory.ReadWrite.All
#   Exchange: Exchange.ManageAsApp
#   Azure AD Role: Exchange Administrator, User Administrator
# ================================================================

param(
  [string]$TenantId     = "${tenantId}",
  [string]$ClientId     = "${clientId}",
  [string]$Domain       = "${domain}",
  [string]$LicensedUser = "${licensedUser}",
  [string]$CertThumbprint = "YOUR_CERT_THUMBPRINT"
)

$ErrorActionPreference = "Continue"
$logFile = "mailbox-automation-$(Get-Date -Format 'yyyyMMdd-HHmm').log"

function Log($msg, $color="White") {
  $ts = Get-Date -Format "HH:mm:ss"
  $line = "[$ts] $msg"
  Write-Host $line -ForegroundColor $color
  Add-Content -Path $logFile -Value $line
}

# ── Connect ──────────────────────────────────────────────────────
Log "Connecting to Exchange Online..." "Cyan"
Connect-ExchangeOnline -AppId $ClientId -Organization $Domain -CertificateThumbprint $CertThumbprint -ShowBanner:$false

Log "Connecting to Microsoft Graph..." "Cyan"
Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertThumbprint -NoWelcome

$Stats = @{ Created=0; Errors=0; Delegated=0 }

$Mailboxes = @(
${mbRows}
)

# ── Step 1: Create Shared Mailboxes ──────────────────────────────
Log "=== STEP 1: Creating $($Mailboxes.Count) shared mailboxes ===" "Cyan"

foreach ($mb in $Mailboxes) {
  $upn = "$($mb.User)@$Domain"
  try {
    $existing = Get-Mailbox -Identity $upn -ErrorAction SilentlyContinue
    if ($existing) {
      Log "  [SKIP] Already exists: $upn" "Yellow"
    } else {
      New-Mailbox -Shared -Name $mb.Name -DisplayName $mb.Name \`
        -Alias $mb.User -PrimarySmtpAddress $upn | Out-Null
      Log "  [OK] Created: $upn" "Green"
      $Stats.Created++
    }

${resetPassword ? `    # Reset password
    $SecPw = ConvertTo-SecureString $mb.Password -AsPlainText -Force
    Update-MgUser -UserId $upn -PasswordProfile @{
      Password = $mb.Password
      ForceChangePasswordNextSignIn = $false
    }
    Log "  [OK] Password reset: $upn" "Green"` : "    # Password reset skipped"}

  } catch {
    Log "  [ERR] $upn : $_" "Red"
    $Stats.Errors++
  }
}

# ── Step 2: Mailbox Delegation ────────────────────────────────────
Log "=== STEP 2: Applying delegation to $LicensedUser ===" "Cyan"

foreach ($mb in $Mailboxes) {
  $upn = "$($mb.User)@$Domain"
  try {
${fullAccess ? `    Add-MailboxPermission -Identity $upn -User $LicensedUser \`
      -AccessRights FullAccess -AutoMapping $${autoMapping ? "true" : "false"} \`
      -ErrorAction Stop | Out-Null
    Log "  [OK] FullAccess: $LicensedUser -> $upn" "Green"` : "    # FullAccess skipped"}

${sendAs ? `    Add-RecipientPermission -Identity $upn -Trustee $LicensedUser \`
      -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
    Log "  [OK] SendAs: $LicensedUser -> $upn" "Green"` : "    # SendAs skipped"}

${sendOnBehalf ? `    Set-Mailbox -Identity $upn -GrantSendOnBehalfTo $LicensedUser -ErrorAction Stop
    Log "  [OK] SendOnBehalf: $LicensedUser -> $upn" "Green"` : "    # SendOnBehalf skipped"}

    $Stats.Delegated++
  } catch {
    Log "  [ERR] Delegation failed for $upn : $_" "Red"
  }
}

# ── Step 3: Exchange Transport Settings ──────────────────────────
Log "=== STEP 3: Exchange Online transport settings ===" "Cyan"

${enableSmtp ? `try {
  Set-TransportConfig -SmtpClientAuthenticationDisabled $false
  Log "[OK] SMTP AUTH enabled — Turn off SMTP is now unchecked" "Green"
} catch {
  Log "[ERR] SMTP setting failed: $_" "Red"
}` : "# SMTP setting skipped"}

# ── Summary ───────────────────────────────────────────────────────
Log "" "White"
Log "=== DONE ===" "Cyan"
Log "Created : $($Stats.Created)" "Green"
Log "Delegated: $($Stats.Delegated)" "Green"
Log "Errors  : $($Stats.Errors)" $(if ($Stats.Errors -gt 0) {"Red"} else {"Green"})
Log "Log saved to: $logFile" "White"

Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph
`;

  res.setHeader("Content-Disposition", "attachment; filename=Create-SharedMailboxes.ps1");
  res.setHeader("Content-Type", "text/plain");
  res.send(script);
});

app.get("/", (req, res) => res.json({ status: "ok", service: "M365 Mailbox Automation API" }));

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
