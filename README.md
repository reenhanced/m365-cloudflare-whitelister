# Reenhanced — Fix Power Automate + Cloudflare

**Is Cloudflare blocking your Power Automate flows from reaching your website?** This free tool fixes that.

It creates ready-to-use firewall rules that tell Cloudflare to let Microsoft Power Automate (and Logic Apps) traffic through. You can hand the instructions to your IT team, or apply them yourself if you have Cloudflare access.

Built by [Reenhanced](https://reenhanced.com) — the world's best solution for WordPress + Microsoft.

**Use the live tool now:** <https://reenhanced.com/power-automate-cloudflare/>

## The problem

When you connect Power Automate to a website that uses Cloudflare (a popular security service), Cloudflare often blocks the connection because it doesn't recognize Microsoft's servers. Your flows fail, webhooks don't fire, and HTTP requests time out.

The fix is to tell Cloudflare "these IP addresses belong to Microsoft — let them through." Microsoft publishes a list of these IP addresses, but the list is long, technical, and changes every week. This tool handles all of that for you.

## How it works

The tool walks you through four simple steps:

1. **Get the latest IP addresses** — Click one button and the tool downloads the current list of Microsoft IP addresses automatically.
2. **Choose what to allow** — The default selection (Power Automate & Connectors) is right for most people. Just leave it as-is.
3. **Get instructions for your IT team** — The tool creates a complete set of instructions you can copy and paste into an email or Teams message. Your IT team gets everything they need: the IP addresses, step-by-step Cloudflare dashboard instructions, and an API command for advanced users.
4. **Apply directly (optional)** — If you manage Cloudflare yourself, you can apply the rules directly from the tool. You'll need your Cloudflare API Token and Zone ID.

After you're done, the tool offers a **calendar reminder** so you can re-run it monthly — Microsoft updates these IP addresses regularly.

## Privacy and security

- **Everything happens in your browser.** The IP address list is downloaded from Microsoft, and all processing happens locally. Nothing is sent to any third-party server.
- **Cloudflare credentials are optional.** You only need them if you choose to apply rules directly. The copy-and-send workflow needs no credentials at all.
- **Credentials stay local.** If you save your Cloudflare credentials, they're stored only in your browser's local storage. You can clear them at any time.
- **This project is open-source.** You can audit every line of code in this repository.

## Try it locally

### Option A — Docker (easiest)

If you have [Docker](https://www.docker.com/products/docker-desktop/) installed:

```bash
cd demo
docker compose up -d
```

Then open <http://localhost:8088> in your browser. Everything works out of the box.

### Option B — Any web server

If you already have a local web server, point it at the root of this repository and open `demo/index.html`.

### Option C — Embed in your own website (WordPress, etc.)

1. Copy `dist/m365-cloudflare-widget.js` and `dist/proxy.php` to your web server.
2. Add this snippet to the page where you want the tool to appear:

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

The `proxy.php` file is a small helper that lets the tool download IP addresses from Microsoft. It's needed because browsers block direct cross-site downloads for security reasons (CORS). The Docker setup uses nginx for this instead.

## What's in this repository

| File | What it does |
|------|-------------|
| `dist/m365-cloudflare-widget.js` | The tool itself — a single JavaScript file with no dependencies |
| `dist/proxy.php` | A small PHP helper for WordPress / Apache servers (downloads the IP list from Microsoft) |
| `demo/index.html` | A demo page showing the tool embedded in a web page |
| `demo/docker-compose.yml` | Launches a local demo with one command (`docker compose up`) |
| `demo/nginx.conf` | Web server config for the Docker demo |
| `demo/Dockerfile` | Container build file for the Docker demo |

## Cloudflare requirements (only for direct apply)

If you want the tool to apply rules directly to Cloudflare (Step 4), you'll need:

- A **Cloudflare API Token** with "Zone > Zone WAF" permissions — [create one here](https://dash.cloudflare.com/profile/api-tokens)
- Your **Zone ID** — found on your domain's overview page in the [Cloudflare dashboard](https://dash.cloudflare.com)

If you don't have these, no worries — just use Step 3 to copy the instructions and send them to whoever manages your Cloudflare account.

## Keeping rules up to date

Microsoft updates their IP addresses weekly. If the addresses change and your Cloudflare rules are outdated, Power Automate may get blocked again.

We recommend **re-running this tool once a month**. The tool can add a calendar reminder for you after you generate instructions.

## Background: what are Azure Service Tags?

Microsoft groups their IP addresses into named sets called "Service Tags." This tool uses two of them:

| Tag | What it covers | Who needs it |
|-----|---------------|-------------|
| **AzureConnectors** | Power Automate, Logic Apps, and related connector services | Most people — this is the default and recommended choice |
| **AzureCloud** | All Microsoft Azure IP addresses worldwide | Only use if specifically asked — this is a very large list |

For more details, see Microsoft's [Service Tags documentation](https://learn.microsoft.com/en-us/azure/virtual-network/service-tags-overview) and [download page](https://www.microsoft.com/en-us/download/details.aspx?id=56519).
