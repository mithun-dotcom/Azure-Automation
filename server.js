const express = require("express");
const cors    = require("cors");
const axios   = require("axios");

const app = express();
app.use(cors());
app.use(express.json({ limit: "10mb" }));

const PORT  = process.env.PORT || 3001;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ── Token helpers ─────────────────────────────────────────────────
async function getGraphToken(tenantId, clientId, clientSecret) {
  const p = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
  });
  const r = await axios.post(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    p.toString(),
    { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
  );
  return r.data.access_token;
}

async function getExchangeToken(tenantId, clientId, clientSecret) {
  const p = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://outlook.office365.com/.default",
  });
  const r = await axios.post(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    p.toString(),
    { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
  );
  return r.data.access_token;
}

// ── Health ────────────────────────────────────────────────────────
app.get("/", (req, res) =>
  res.json({ status: "ok", version: "4.0", service: "M365 Mailbox Automation" })
);

// ── Test connection ───────────────────────────────────────────────
app.post("/api/test-connection", async (req, res) => {
  const { tenantId, clientId, clientSecret } = req.body;
  try {
    const token = await getGraphToken(tenantId, clientId, clientSecret);
    const org = await axios.get("https://graph.microsoft.com/v1.0/organization", {
      headers: { Authorization: `Bearer ${token}` },
    });
    res.json({ success: true, org: org.data.value[0]?.displayName || "Connected" });
  } catch (err) {
    res.status(401).json({
      success: false,
      message: err.response?.data?.error_description || err.message,
    });
  }
});

// ── Create shared mailbox ─────────────────────────────────────────
// Strategy:
//   1. Create user via Graph API
//   2. Convert to shared mailbox via Graph mailboxSettings
//   3. If Exchange not provisioned yet, retry up to 6x with 10s gaps
app.post("/api/create-mailbox", async (req, res) => {
  const { tenantId, clientId, clientSecret, domain, username, displayName, password } = req.body;

  const skip = ["username", "user", "email", "displayname", "password", "pass"];
  if (!username || skip.includes(username.toLowerCase())) {
    return res.status(400).json({ success: false, upn: `${username}@${domain}`, message: "Skipped header row." });
  }

  const upn = `${username}@${domain}`;
  try {
    const token = await getGraphToken(tenantId, clientId, clientSecret);

    // Step 1: Create or get user
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
        { headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" } }
      );
      userId = r.data.id;
    } catch (e) {
      if (e.response?.status === 400 || e.response?.status === 409) {
        // User already exists
        const ex = await axios.get(
          `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(upn)}`,
          { headers: { Authorization: `Bearer ${token}` } }
        );
        userId = ex.data.id;
      } else throw e;
    }

    // Step 2: Convert to shared mailbox — retry until Exchange provisions it
    let converted = false;
    for (let attempt = 1; attempt <= 6; attempt++) {
      try {
        await axios.patch(
          `https://graph.microsoft.com/v1.0/users/${userId}/mailboxSettings`,
          { userPurpose: "shared" },
          { headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" } }
        );
        converted = true;
        break;
      } catch (e) {
        const msg = e.response?.data?.error?.message || "";
        const notReady = msg.includes("inactive") || msg.includes("soft-deleted") ||
          msg.includes("on-premise") || msg.includes("MailboxNotEnabled") || e.response?.status === 404;
        if (notReady && attempt < 6) {
          await sleep(10000); // wait 10s and retry
          continue;
        }
        break; // non-retryable or max retries — user still created
      }
    }

    res.json({
      success: true,
      upn,
      userId,
      converted,
      message: converted
        ? `Shared mailbox created: ${upn}`
        : `User created: ${upn} — Exchange provisioning in background (run PS1 to finalize)`,
    });
  } catch (err) {
    res.status(500).json({
      success: false,
      upn,
      message: err.response?.data?.error?.message || err.message,
    });
  }
});

// ── Reset password ────────────────────────────────────────────────
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

// ── Mailbox delegation ────────────────────────────────────────────
// Graph does not support Add-MailboxPermission (FullAccess).
// We return a clear message so the frontend shows it correctly.
// The PS1 script handles FullAccess via Exchange PowerShell.
app.post("/api/add-delegation", async (req, res) => {
  res.json({
    success: true,
    results: [
      { type: "FullAccess", success: false, note: "Use Download PS1 script for FullAccess delegation" },
      { type: "SendAs",     success: false, note: "Use Download PS1 script for SendAs delegation" },
    ],
  });
});

// ── Enable SMTP AUTH ──────────────────────────────────────────────
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
    res.status(500).json({ success: false, message: err.response?.data?.message || err.message });
  }
});

// ── Generate PowerShell script ────────────────────────────────────
// This script completes what the web app can't do via Graph API:
//   - Converts users to proper shared mailboxes in Exchange
//   - Delegates FullAccess + SendAs to the licensed user
// Uses Tenant ID + Client ID + Client Secret — no interactive login
app.post("/api/generate-script", (req, res) => {
  const { tenantId, clientId, clientSecret, domain, licensedUser, mailboxes, enableSmtp } = req.body;

  const rows = (mailboxes || [])
    .map((m) => `  @{User="${m.username}"; Name="${m.displayName}"}`)
    .join(",\n");

  const smtpLine = enableSmtp
    ? `try { Set-TransportConfig -SmtpClientAuthenticationDisabled $false\n  Write-Host "[OK] SMTP AUTH enabled" -ForegroundColor Green } catch { Write-Host "[ERR] SMTP: $_" -ForegroundColor Red }`
    : "# SMTP step skipped";

  const ps = `# ============================================================
# M365 Shared Mailbox Script
# Auth: Tenant ID + Client ID + Client Secret
# No interactive login required
# ============================================================
# REQUIREMENTS (run once):
#   Install-Module ExchangeOnlineManagement -Force
# THEN:
#   .\\Create-SharedMailboxes.ps1
# ============================================================

$TenantId     = "${tenantId}"
$ClientId     = "${clientId}"
$ClientSecret = "${clientSecret}"
$Domain       = "${domain}"
$LicensedUser = "${licensedUser}"

$log = "run-$(Get-Date -Format 'yyyyMMdd-HHmm').log"
function L($m, $c="White") {
  $ts = Get-Date -Format "HH:mm:ss"
  Write-Host "[$ts] $m" -ForegroundColor $c
  Add-Content $log "[$ts] $m"
}

# Connect to Exchange Online using Client Secret (no interactive login)
L "Getting Exchange token..." Cyan
$tokenBody = @{
  grant_type    = "client_credentials"
  client_id     = $ClientId
  client_secret = $ClientSecret
  scope         = "https://outlook.office365.com/.default"
}
$tok = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $tokenBody -ContentType "application/x-www-form-urlencoded"
$ExchangeToken = ConvertTo-SecureString $tok.access_token -AsPlainText -Force

L "Connecting to Exchange Online..." Cyan
Connect-ExchangeOnline -AppId $ClientId -Organization $Domain -AccessToken $ExchangeToken -ShowBanner:$false
L "[OK] Connected to Exchange Online" Green

$ok=0; $skip=0; $err=0; $delegated=0

$Mailboxes = @(
${rows}
)

# ── STEP 1: Create shared mailboxes ──────────────────────────────
L "" White
L "=== STEP 1: Creating $($Mailboxes.Count) shared mailboxes ===" Cyan

foreach ($mb in $Mailboxes) {
  $upn = "$($mb.User)@$Domain"
  try {
    if (Get-Mailbox -Identity $upn -ErrorAction SilentlyContinue) {
      L "  [SKIP] Already exists: $upn" Yellow; $skip++
    } else {
      New-Mailbox -Shared -Name $mb.Name -DisplayName $mb.Name -Alias $mb.User -PrimarySmtpAddress $upn | Out-Null
      L "  [OK] Created: $upn" Green; $ok++
    }
  } catch { L "  [ERR] $upn : $_" Red; $err++ }
}

# ── STEP 2: Delegate all mailboxes to licensed user ───────────────
L "" White
L "=== STEP 2: Delegating all mailboxes to $LicensedUser ===" Cyan

foreach ($mb in $Mailboxes) {
  $upn = "$($mb.User)@$Domain"
  try {
    Add-MailboxPermission -Identity $upn -User $LicensedUser -AccessRights FullAccess -AutoMapping $true -Confirm:$false | Out-Null
    Add-RecipientPermission -Identity $upn -Trustee $LicensedUser -AccessRights SendAs -Confirm:$false | Out-Null
    L "  [OK] Delegated: $upn -> $LicensedUser" Green; $delegated++
  } catch { L "  [ERR] $upn : $_" Red }
}

# ── STEP 3: SMTP ──────────────────────────────────────────────────
L "" White
L "=== STEP 3: SMTP AUTH ===" Cyan
${smtpLine}

L "" White
L "=== DONE: Created=$ok  Skipped=$skip  Errors=$err  Delegated=$delegated ===" Cyan
L "Log saved to: $log" White

Disconnect-ExchangeOnline -Confirm:$false
`;

  res.setHeader("Content-Disposition", "attachment; filename=Create-SharedMailboxes.ps1");
  res.setHeader("Content-Type", "text/plain");
  res.send(ps);
});

app.listen(PORT, () => console.log(`M365 API v4.0 on port ${PORT}`));
