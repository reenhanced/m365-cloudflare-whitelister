# Reenhanced — Azure → Cloudflare Whitelister

Allow **Power Automate** and **Logic Apps** traffic through your Cloudflare firewall.

Built by [Reenhanced](https://reenhanced.com) — the world's best solution for WordPress + Microsoft.

## What it does

1. Click **Fetch Latest Service Tags** — the widget auto-downloads the latest Azure Service Tags JSON from Microsoft through a lightweight proxy.
2. Select service tags (**AzureConnectors**, **AzureCloud**) and optionally specific Azure regions.
3. Preview the exact Cloudflare rules, IP ranges, curl commands, and manual instructions.
4. **Copy the preview** and hand it to a technical resource — or apply directly to Cloudflare from the widget.

Manual file upload is available as a fallback if the proxy is unavailable.

No backend, no server calls except to Microsoft (via proxy) and directly to Cloudflare when you choose to apply. The Azure Service Tags file is processed entirely in your browser.

## Quick start

### Option A — Docker

```bash
cd demo
docker compose up -d
```

Open <http://localhost:8088>. The included nginx config provides the CORS proxy at `/api/ms-download/` so auto-fetch works out of the box.

### Option B — Any static server

Serve the repo root and open `demo/index.html`.

### Option C — Embed in WordPress

1. Upload `dist/m365-cloudflare-widget.js` and `dist/proxy.php` to your WordPress theme or a public directory.
2. Add this HTML where you want the UI to appear:

```html
<div id="reenhanced-m365-cf-widget"></div>
<script src="https://YOUR_DOMAIN/path/to/m365-cloudflare-widget.js"></script>
<script>
  window.ReenhancedM365CFWidget.init({
    target: "#reenhanced-m365-cf-widget",
    proxyBaseUrl: "https://YOUR_DOMAIN/path/to/proxy.php?url="
  });
</script>
```

The `proxy.php` file acts as a CORS proxy to download.microsoft.com so the widget can auto-fetch the Service Tags JSON.

## Files

| File | Purpose |
|------|---------|
| `dist/m365-cloudflare-widget.js` | Self-contained embeddable widget (no build step needed) |
| `dist/proxy.php` | PHP CORS proxy for WordPress / Apache deployments |
| `demo/index.html` | Demo page that mirrors a WordPress embed |
| `demo/docker-compose.yml` | One-command Docker demo |
| `demo/nginx.conf` | Nginx routing config with CORS proxy |

## Data flow

1. Widget auto-fetches the latest `ServiceTags_Public_*.json` from [Microsoft](https://www.microsoft.com/en-us/download/details.aspx?id=56519) via a CORS proxy (nginx or PHP).
2. Widget parses the file **locally** to extract AzureConnectors and/or AzureCloud CIDR ranges.
3. User selects regions and generates a preview.
4. Preview includes:
   - Full list of CIDR ranges
   - Ready-to-run `curl` command for the Cloudflare API
   - Step-by-step manual dashboard instructions
5. User can **copy** the preview and send it to a colleague, or **apply** directly to Cloudflare.

## Cloudflare requirements

- An [API Token](https://dash.cloudflare.com/profile/api-tokens) with **Zone > Zone WAF** (or broader) permissions.
- The Zone ID for your domain (visible on the zone overview page in Cloudflare).
- **Credentials are only needed if you choose to apply directly.** The preview and copy features work without any credentials.

## Transparency

- All processing happens locally in your browser.
- Credentials are only stored in `localStorage` if you explicitly click "Save Credentials Locally".
- This repository is open-source — audit it, fork it, run it locally.

## Azure Service Tags reference

Service Tags are updated weekly by Microsoft. Re-run this tool periodically to keep your rules current.

- Download page: <https://www.microsoft.com/en-us/download/details.aspx\?id\=56519\>
- Documentation: <https://learn.microsoft.com/en-us/azure/virtual-network/service-tags-overview\>

| Tag | What it covers |
|-----|---------------|
| **AzureConnectors** | Power Automate, Logic Apps, and managed connector infrastructure |
| **AzureCloud** | All Azure datacenter IP ranges (much broader — use with caution) |
