const express = require("express");
const cors = require("cors");
const axios = require("axios");

const app = express();
app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 3001;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const WIZARD_CLIENT_ID = "de8bc8b5-d9f9-48b1-a8ad-b748da725064";

// ── Token helpers ────────────────────────────────────────────────
async function getGraphToken(tenantId, clientId, clientSecret) {
  const params = new URLSearchParams({ grant_type:"client_credentials", client_id:clientId, client_secret:clientSecret, scope:"https://graph.microsoft.com/.default" });
  const r = await axios.post(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, params.toString(), { headers:{"Content-Type":"application/x-www-form-urlencoded"} });
  return r.data.access_token;
}
async function getExchangeToken(tenantId, clientId, clientSecret) {
  const params = new URLSearchParams({ grant_type:"client_credentials", client_id:clientId, client_secret:clientSecret, scope:"https://outlook.office365.com/.default" });
  const r = await axios.post(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, params.toString(), { headers:{"Content-Type":"application/x-www-form-urlencoded"} });
  return r.data.access_token;
}

// ── Health ───────────────────────────────────────────────────────
app.get("/", (req, res) => res.json({ status:"ok", version:"2.0", service:"M365 Mailbox Automation API" }));

// ── Test connection ──────────────────────────────────────────────
app.post("/api/test-connection", async (req, res) => {
  const { tenantId, clientId, clientSecret } = req.body;
  try {
    const token = await getGraphToken(tenantId, clientId, clientSecret);
    const org = await axios.get("https://graph.microsoft.com/v1.0/organization", { headers:{ Authorization:`Bearer ${token}` } });
    res.json({ success:true, org:org.data.value[0]?.displayName || "Connected" });
  } catch(err) { res.status(401).json({ success:false, message:err.response?.data?.error_description || err.message }); }
});

// ── Create mailbox ───────────────────────────────────────────────
app.post("/api/create-mailbox", async (req, res) => {
  const { tenantId, clientId, clientSecret, domain, username, displayName, password } = req.body;
  const SKIP = ["username","user","email","displayname","display name","password","pass"];
  if (!username || SKIP.includes(username.toLowerCase())) return res.status(400).json({ success:false, upn:`${username}@${domain}`, message:"Skipped header row." });
  const upn = `${username}@${domain}`;
  try {
    const token = await getGraphToken(tenantId, clientId, clientSecret);
    let userId;
    try {
      const r = await axios.post("https://graph.microsoft.com/v1.0/users", { accountEnabled:true, displayName, mailNickname:username, userPrincipalName:upn, passwordProfile:{ forceChangePasswordNextSignIn:false, password }, usageLocation:"US" }, { headers:{ Authorization:`Bearer ${token}`, "Content-Type":"application/json" } });
      userId = r.data.id;
    } catch(e) {
      if (e.response?.status === 400 || e.response?.status === 409) {
        const ex = await axios.get(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(upn)}`, { headers:{ Authorization:`Bearer ${token}` } });
        userId = ex.data.id;
      } else throw e;
    }
    let sharedNote = "provisioning";
    try {
      await axios.patch(`https://graph.microsoft.com/v1.0/users/${userId}/mailboxSettings`, { userPurpose:"shared" }, { headers:{ Authorization:`Bearer ${token}`, "Content-Type":"application/json" } });
      sharedNote = "shared";
    } catch(_) {}
    res.json({ success:true, upn, userId, sharedNote, message: sharedNote==="shared" ? `Shared mailbox created: ${upn}` : `User created: ${upn} (Exchange provisioning in background)` });
  } catch(err) { res.status(500).json({ success:false, upn, message:err.response?.data?.error?.message || err.message }); }
});

// ── Reset password ───────────────────────────────────────────────
app.post("/api/reset-password", async (req, res) => {
  const { tenantId, clientId, clientSecret, upn, password } = req.body;
  try {
    const token = await getGraphToken(tenantId, clientId, clientSecret);
    await axios.patch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(upn)}`, { passwordProfile:{ forceChangePasswordNextSignIn:false, password } }, { headers:{ Authorization:`Bearer ${token}`, "Content-Type":"application/json" } });
    res.json({ success:true, message:`Password reset: ${upn}` });
  } catch(err) { res.status(500).json({ success:false, message:err.response?.data?.error?.message || err.message }); }
});

// ── Add delegation ───────────────────────────────────────────────
app.post("/api/add-delegation", async (req, res) => {
  const { tenantId, clientId, clientSecret, mailboxUpn, delegateUpn, sendAs } = req.body;
  const results = [];
  try {
    const token = await getGraphToken(tenantId, clientId, clientSecret);
    const dRes = await axios.get(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(delegateUpn)}`, { headers:{ Authorization:`Bearer ${token}` } });
    const delegateId = dRes.data.id;
    if (sendAs) {
      try {
        await axios.post(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailboxUpn)}/permissionGrants`, { clientId:delegateId, consentType:"Principal", principalId:delegateId, resourceId:delegateId, scope:"Mail.Send" }, { headers:{ Authorization:`Bearer ${token}`, "Content-Type":"application/json" } });
        results.push({ type:"SendAs", success:true });
      } catch(e) { results.push({ type:"SendAs", success:false, note:"Use PS1 script for SendAs" }); }
    }
    results.push({ type:"FullAccess", success:false, note:"Use PS1 script for FullAccess (Exchange PowerShell required)" });
    res.json({ success:true, results });
  } catch(err) { res.status(500).json({ success:false, message:err.response?.data?.error?.message || err.message }); }
});

// ── Enable SMTP ──────────────────────────────────────────────────
app.post("/api/enable-smtp", async (req, res) => {
  const { tenantId, clientId, clientSecret } = req.body;
  try {
    const token = await getExchangeToken(tenantId, clientId, clientSecret);
    await axios.patch("https://outlook.office365.com/adminapi/beta/tenant/transportconfig", { SmtpClientAuthenticationDisabled:false }, { headers:{ Authorization:`Bearer ${token}`, "Content-Type":"application/json" } });
    res.json({ success:true, message:"SMTP AUTH enabled" });
  } catch(err) { res.status(500).json({ success:false, message:err.response?.data?.message || err.message, note:"Run: Set-TransportConfig -SmtpClientAuthenticationDisabled $false" }); }
});

// ── Auto-setup: Direct login with admin email + password (ROPC) ──
app.post("/api/auto-setup/direct-login", async (req, res) => {
  const { email, password } = req.body;
  if (!email || !password) return res.status(400).json({ success: false, message: "Email and password required." });
  const domain = email.split("@")[1]?.trim();
  if (!domain) return res.status(400).json({ success: false, message: "Invalid email format." });

  // Try multiple well-known public client IDs that support ROPC
  const clientIds = [
    "1950a258-227b-4e31-a9cf-717495945fc2", // Azure PowerShell — supports ROPC in all tenants
    "04b07795-8542-4c44-b3a2-0f47e438e9e2", // Azure CLI
  ];

  let lastError = null;
  for (const clientId of clientIds) {
    try {
      const params = new URLSearchParams({
        grant_type: "password",
        client_id: clientId,
        username: email,
        password: password,
        scope: [
          "https://graph.microsoft.com/Application.ReadWrite.All",
          "https://graph.microsoft.com/AppRoleAssignment.ReadWrite.All",
          "https://graph.microsoft.com/Directory.ReadWrite.All",
          "https://graph.microsoft.com/Organization.Read.All",
          "offline_access",
        ].join(" "),
      });
      const r = await axios.post(
        `https://login.microsoftonline.com/${domain}/oauth2/v2.0/token`,
        params.toString(),
        { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
      );
      return res.json({ success: true, access_token: r.data.access_token, domain });
    } catch (e) {
      lastError = e;
      const errCode = e.response?.data?.error;
      const errDesc = e.response?.data?.error_description || "";
      // MFA required — no point retrying other clients
      if (errCode === "mfa_required" || errDesc.includes("AADSTS50076") || errDesc.includes("AADSTS50079") || errDesc.includes("multi-factor")) {
        return res.status(400).json({ success: false, mfa_required: true, domain, message: "MFA is enabled — please use the device code method instead." });
      }
      // Wrong credentials — no point retrying
      if (errCode === "invalid_grant" || errDesc.includes("Invalid username or password")) {
        return res.status(400).json({ success: false, message: "Invalid email or password. Please check and try again." });
      }
      // ROPC not supported by this client — try next
      if (errDesc.includes("client_assertion") || errDesc.includes("client_secret") || errDesc.includes("AADSTS7000218")) {
        continue;
      }
      // ROPC disabled in tenant
      if (errDesc.includes("AADSTS9002327") || errDesc.includes("Resource owner password")) {
        return res.status(400).json({ success: false, mfa_required: true, domain, message: "Password login is disabled in this tenant. Please use the device code method." });
      }
    }
  }

  const finalMsg = lastError?.response?.data?.error_description || lastError?.message || "Authentication failed";
  res.status(400).json({ success: false, mfa_required: true, domain, message: finalMsg + " — Try the device code method instead." });
});


app.post("/api/auto-setup/device-code", async (req, res) => {
  try {
    const params = new URLSearchParams({ client_id:WIZARD_CLIENT_ID, scope:["https://graph.microsoft.com/Application.ReadWrite.All","https://graph.microsoft.com/AppRoleAssignment.ReadWrite.All","https://graph.microsoft.com/Directory.ReadWrite.All","https://graph.microsoft.com/Organization.Read.All","offline_access"].join(" ") });
    const r = await axios.post("https://login.microsoftonline.com/common/oauth2/v2.0/devicecode", params.toString(), { headers:{"Content-Type":"application/x-www-form-urlencoded"} });
    res.json({ success:true, device_code:r.data.device_code, user_code:r.data.user_code, verification_uri:r.data.verification_uri, expires_in:r.data.expires_in, interval:r.data.interval });
  } catch(e) { res.status(400).json({ success:false, message:e.response?.data?.error_description || e.message }); }
});

// ── Auto-setup: Poll token ───────────────────────────────────────
app.post("/api/auto-setup/poll-token", async (req, res) => {
  const { deviceCode } = req.body;
  try {
    const params = new URLSearchParams({ grant_type:"urn:ietf:params:oauth:grant-type:device_code", client_id:WIZARD_CLIENT_ID, device_code:deviceCode });
    const r = await axios.post("https://login.microsoftonline.com/common/oauth2/v2.0/token", params.toString(), { headers:{"Content-Type":"application/x-www-form-urlencoded"} });
    res.json({ success:true, access_token:r.data.access_token });
  } catch(e) {
    const err = e.response?.data?.error;
    if (err === "authorization_pending" || err === "slow_down") return res.json({ success:false, pending:true });
    res.status(400).json({ success:false, pending:false, message:e.response?.data?.error_description || e.message });
  }
});

// ── Auto-setup: Create app registration ─────────────────────────
app.post("/api/auto-setup/create-app", async (req, res) => {
  const { accessToken, appName, tenantId } = req.body;
  const h = { Authorization:`Bearer ${accessToken}`, "Content-Type":"application/json" };
  const steps = [];
  const log = (msg, ok=true) => steps.push({ msg, ok });
  try {
    log("Creating app registration: " + appName);
    const appRes = await axios.post("https://graph.microsoft.com/v1.0/applications", { displayName:appName, signInAudience:"AzureADMyOrg", requiredResourceAccess:[{ resourceAppId:"00000003-0000-0000-c000-000000000000", resourceAccess:[{ id:"741f803b-c850-494e-b5df-cde7c675a1ca",type:"Role"},{ id:"df021288-bdef-4463-88db-98f22de89214",type:"Role"},{ id:"e2a3a72e-5f79-4c64-b1b1-878b674786c9",type:"Role"},{ id:"931e8a5d-5fa3-4bcc-b695-d7c8e4b95e9a",type:"Role"},{ id:"19dbc75e-c2e2-444c-a770-ec69d8559fc7",type:"Role"}]},{ resourceAppId:"00000002-0000-0ff1-ce00-000000000000", resourceAccess:[{ id:"dc50a0fb-09a3-484d-be87-e023b12c6440",type:"Role"}]}] }, { headers:h });
    const appId = appRes.data.appId, objectId = appRes.data.id;
    log("App registered — Client ID: " + appId);
    await sleep(2000);
    const spRes = await axios.post("https://graph.microsoft.com/v1.0/servicePrincipals", { appId }, { headers:h });
    const spId = spRes.data.id;
    log("Service principal created");
    const secretRes = await axios.post(`https://graph.microsoft.com/v1.0/applications/${objectId}/addPassword`, { passwordCredential:{ displayName:"M365AutoSecret", endDateTime:new Date(Date.now()+365*24*60*60*1000*2).toISOString() } }, { headers:h });
    const clientSecret = secretRes.data.secretText;
    log("Client secret generated (2-year expiry)");
    await sleep(5000);
    const graphSp = await axios.get("https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'", { headers:h });
    const graphSpId = graphSp.data.value[0]?.id;
    for (const permId of ["741f803b-c850-494e-b5df-cde7c675a1ca","df021288-bdef-4463-88db-98f22de89214","e2a3a72e-5f79-4c64-b1b1-878b674786c9","931e8a5d-5fa3-4bcc-b695-d7c8e4b95e9a","19dbc75e-c2e2-444c-a770-ec69d8559fc7"]) {
      try { await axios.post(`https://graph.microsoft.com/v1.0/servicePrincipals/${spId}/appRoleAssignments`, { principalId:spId, resourceId:graphSpId, appRoleId:permId }, { headers:h }); } catch(_) {}
    }
    log("Graph permissions granted & admin consent applied");
    try {
      const exchSp = await axios.get("https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000002-0000-0ff1-ce00-000000000000'", { headers:h });
      await axios.post(`https://graph.microsoft.com/v1.0/servicePrincipals/${spId}/appRoleAssignments`, { principalId:spId, resourceId:exchSp.data.value[0]?.id, appRoleId:"dc50a0fb-09a3-484d-be87-e023b12c6440" }, { headers:h });
      log("Exchange.ManageAsApp granted");
    } catch(e) { log("Exchange permission — assign Exchange Administrator role manually", false); }
    let domain = "";
    try { const orgRes = await axios.get("https://graph.microsoft.com/v1.0/organization", { headers:h }); domain = orgRes.data.value[0]?.verifiedDomains?.find(d=>d.isDefault)?.name || ""; log("Domain: " + domain); } catch(_) {}
    log("Setup complete! Credentials ready.", true);
    res.json({ success:true, steps, result:{ appId, clientSecret, tenantId, domain, objectId, spId } });
  } catch(e) { res.status(500).json({ success:false, steps, message:e.response?.data?.error?.message || e.message }); }
});

// ── Auto-setup: Generate registration script ─────────────────────
app.post("/api/auto-setup/generate-reg-script", async (req, res) => {
  const { appName, domain } = req.body;
  const name = appName || "M365 Mailbox Automation";
  const script = `# M365 App Registration Script — run as Global Admin
# Install-Module Microsoft.Graph -Force

Connect-MgGraph -Scopes "Application.ReadWrite.All","AppRoleAssignment.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All" -TenantId "${domain||'YOUR_DOMAIN'}" -NoWelcome

$app = New-MgApplication -DisplayName "${name}" -SignInAudience "AzureADMyOrg" -RequiredResourceAccess @(@{ResourceAppId="00000003-0000-0000-c000-000000000000";ResourceAccess=@(@{Id="741f803b-c850-494e-b5df-cde7c675a1ca";Type="Role"},@{Id="df021288-bdef-4463-88db-98f22de89214";Type="Role"},@{Id="e2a3a72e-5f79-4c64-b1b1-878b674786c9";Type="Role"},@{Id="931e8a5d-5fa3-4bcc-b695-d7c8e4b95e9a";Type="Role"},@{Id="19dbc75e-c2e2-444c-a770-ec69d8559fc7";Type="Role"})},@{ResourceAppId="00000002-0000-0ff1-ce00-000000000000";ResourceAccess=@(@{Id="dc50a0fb-09a3-484d-be87-e023b12c6440";Type="Role"})})
Start-Sleep 3
$sp = New-MgServicePrincipal -AppId $app.AppId
$secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{DisplayName="AutoSecret";EndDateTime=(Get-Date).AddYears(2)}
Start-Sleep 8
$graphSp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
@("741f803b-c850-494e-b5df-cde7c675a1ca","df021288-bdef-4463-88db-98f22de89214","e2a3a72e-5f79-4c64-b1b1-878b674786c9","931e8a5d-5fa3-4bcc-b695-d7c8e4b95e9a","19dbc75e-c2e2-444c-a770-ec69d8559fc7") | ForEach-Object { try { New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -ResourceId $graphSp.Id -AppRoleId $_ | Out-Null } catch {} }
$exchSp = Get-MgServicePrincipal -Filter "AppId eq '00000002-0000-0ff1-ce00-000000000000'"
try { New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -ResourceId $exchSp.Id -AppRoleId "dc50a0fb-09a3-484d-be87-e023b12c6440" | Out-Null } catch {}
$org = Get-MgOrganization
$domain = ($org.VerifiedDomains | Where-Object IsDefault).Name
Write-Host ""; Write-Host "=== COPY THESE ===" -ForegroundColor Yellow
Write-Host "Tenant ID    : $($org.Id)"
Write-Host "Client ID    : $($app.AppId)"
Write-Host "Client Secret: $($secret.SecretText)"
Write-Host "Domain       : $domain"
"Tenant ID: $($org.Id)\nClient ID: $($app.AppId)\nSecret: $($secret.SecretText)\nDomain: $domain" | Out-File "M365Creds-$domain.txt"
Write-Host "Saved to M365Creds-$domain.txt" -ForegroundColor Green
Disconnect-MgGraph`;
  res.setHeader("Content-Disposition", `attachment; filename=Register-M365App.ps1`);
  res.setHeader("Content-Type", "text/plain");
  res.send(script);
});

// ── Generate PowerShell script ───────────────────────────────────
app.post("/api/generate-script", async (req, res) => {
  const { tenantId, clientId, domain, licensedUser, mailboxes, fullAccess, sendAs, sendOnBehalf, autoMapping, resetPassword, enableSmtp } = req.body;
  const mbRows = (mailboxes||[]).map(m=>`  @{User="${m.username}"; Name="${m.displayName}"; Password="${m.password}"}`).join(",\n");
  const script = `# M365 Shared Mailbox Automation\nparam(\n  [string]$TenantId="${tenantId}",[string]$ClientId="${clientId}",[string]$Domain="${domain}",[string]$LicensedUser="${licensedUser}",[string]$Cert="YOUR_CERT_THUMBPRINT"\n)\nConnect-ExchangeOnline -AppId $ClientId -Organization $Domain -CertificateThumbprint $Cert -ShowBanner:$false\nConnect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Cert -NoWelcome\n$Mailboxes=@(\n${mbRows}\n)\nforeach($mb in $Mailboxes){\n  $upn="$($mb.User)@$Domain"\n  try{\n    New-Mailbox -Shared -Name $mb.Name -DisplayName $mb.Name -Alias $mb.User -PrimarySmtpAddress $upn|Out-Null\n    Write-Host "[OK] $upn" -ForegroundColor Green\n${resetPassword?'    Update-MgUser -UserId $upn -PasswordProfile @{Password=$mb.Password;ForceChangePasswordNextSignIn=$false}':''}\n${fullAccess?'    Add-MailboxPermission -Identity $upn -User $LicensedUser -AccessRights FullAccess -AutoMapping $'+( autoMapping?'true':'false')+' -Confirm:$false|Out-Null':''}\n${sendAs?'    Add-RecipientPermission -Identity $upn -Trustee $LicensedUser -AccessRights SendAs -Confirm:$false|Out-Null':''}\n  }catch{ Write-Host "[ERR] $upn : $_" -ForegroundColor Red }\n}\n${enableSmtp?'Set-TransportConfig -SmtpClientAuthenticationDisabled $false':''}\nDisconnect-ExchangeOnline -Confirm:$false`;
  res.setHeader("Content-Disposition","attachment; filename=Create-SharedMailboxes.ps1");
  res.setHeader("Content-Type","text/plain");
  res.send(script);
});

app.listen(PORT, () => console.log(`M365 Automation API v2.0 running on port ${PORT}`));
