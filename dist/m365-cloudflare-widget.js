(function () {
  "use strict";

  /* ------------------------------------------------------------------ */
  /*  Constants                                                          */
  /* ------------------------------------------------------------------ */
  var STORAGE_KEY = "reenhanced_pa_cf_widget_v1";
  var RULE_REF_PREFIX = "reenhanced_azure_allow";
  var SERVICE_TAGS_DOWNLOAD = "https://www.microsoft.com/en-us/download/details.aspx?id=56519";
  var SERVICE_TAGS_GUID_PATH = "download/7/1/d/71d86715-5596-4529-9b13-da13a5de5b63/ServiceTags_Public_";
  var TAG_PREFIXES = {
    AzureConnectors: {
      label: "Power Automate & Connectors",
      shortLabel: "Power Automate",
      description:
        "The IP addresses used by Power Automate and Logic Apps to reach your website. This is what most people need.",
      recommended: true
    },
    AzureCloud: {
      label: "All Azure Services",
      shortLabel: "All Azure",
      description:
        "A much larger set of IP addresses covering all of Microsoft Azure. Only use this if someone specifically told you to.",
      recommended: false
    }
  };

  /* ------------------------------------------------------------------ */
  /*  Utilities                                                          */
  /* ------------------------------------------------------------------ */
  function esc(v) {
    return String(v)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function dedupe(arr) {
    var seen = Object.create(null);
    return arr.filter(function (v) {
      if (seen[v]) return false;
      seen[v] = true;
      return true;
    });
  }

  function chunkArray(arr, size) {
    var out = [];
    for (var i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
    return out;
  }

  function copyText(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(text);
    }
    var ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
    return Promise.resolve();
  }

  function readSettings() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch (e) {
      return null;
    }
  }

  function writeSettings(s) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(s));
  }

  function clearSettings() {
    localStorage.removeItem(STORAGE_KEY);
  }

  /* ------------------------------------------------------------------ */
  /*  Auto-fetch Azure Service Tags                                      */
  /* ------------------------------------------------------------------ */
  function pad2(n) { return n < 10 ? "0" + n : "" + n; }

  function todayDate() {
    var d = new Date();
    return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
  }

  function defaultRuleDesc() {
    return "Allow Power Automate / Azure Connectors (Reenhanced " + todayDate() + ")";
  }

  var LAST_RUN_KEY = "reenhanced_pa_cf_last_run";

  function recordLastRun() {
    localStorage.setItem(LAST_RUN_KEY, todayDate());
  }

  function getLastRun() {
    try { return localStorage.getItem(LAST_RUN_KEY) || null; } catch (e) { return null; }
  }

  function daysSinceLastRun() {
    var last = getLastRun();
    if (!last) return null;
    var parts = last.split("-");
    var then = new Date(+parts[0], +parts[1] - 1, +parts[2]);
    var now = new Date();
    return Math.floor((now - then) / 86400000);
  }

  function generateICS() {
    // Create a calendar event 30 days from now
    var now = new Date();
    var remind = new Date(now.getTime() + 30 * 86400000);
    var remindEnd = new Date(remind.getTime() + 3600000); // 1-hour event

    function icsDate(d) {
      return d.getUTCFullYear() +
        pad2(d.getUTCMonth() + 1) +
        pad2(d.getUTCDate()) + "T" +
        pad2(d.getUTCHours()) +
        pad2(d.getUTCMinutes()) +
        pad2(d.getUTCSeconds()) + "Z";
    }

    var uid = "reenhanced-m365-cf-" + now.getTime() + "@reenhanced.com";
    var lines = [
      "BEGIN:VCALENDAR",
      "VERSION:2.0",
      "PRODID:-//Reenhanced//M365 CF Whitelister//EN",
      "BEGIN:VEVENT",
      "UID:" + uid,
      "DTSTAMP:" + icsDate(now),
      "DTSTART:" + icsDate(remind),
      "DTEND:" + icsDate(remindEnd),
      "SUMMARY:Update Cloudflare rules for Power Automate (Reenhanced)",
      "DESCRIPTION:Microsoft updates their IP addresses regularly. Re-run the Reenhanced whitelister tool to keep your Cloudflare rules current.\\n\\nhttps://reenhanced.com",
      "BEGIN:VALARM",
      "TRIGGER:-PT15M",
      "ACTION:DISPLAY",
      "DESCRIPTION:Time to update your Cloudflare Power Automate rules",
      "END:VALARM",
      "END:VEVENT",
      "END:VCALENDAR"
    ];
    return lines.join("\r\n");
  }

  function downloadICS() {
    var blob = new Blob([generateICS()], { type: "text/calendar;charset=utf-8" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = "update-cloudflare-power-automate.ics";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function candidateDates() {
    // Microsoft publishes weekly, typically on Mondays.
    // Generate candidate dates for the last 8 weeks of Mondays,
    // plus today and the last few days in case of mid-week updates.
    var dates = [];
    var now = new Date();
    // Try today and the last 3 days first
    for (var d = 0; d <= 3; d++) {
      var dt = new Date(now);
      dt.setDate(dt.getDate() - d);
      dates.push(dt.getFullYear() + pad2(dt.getMonth() + 1) + pad2(dt.getDate()));
    }
    // Then try recent Mondays (day 1) going back 8 weeks
    for (var w = 0; w < 8; w++) {
      var monday = new Date(now);
      monday.setDate(monday.getDate() - monday.getDay() + 1 - w * 7);
      var key = monday.getFullYear() + pad2(monday.getMonth() + 1) + pad2(monday.getDate());
      if (dates.indexOf(key) === -1) dates.push(key);
    }
    return dates;
  }

  async function autoFetchServiceTags(proxyBaseUrl) {
    var dates = candidateDates();
    for (var i = 0; i < dates.length; i++) {
      var relPath = SERVICE_TAGS_GUID_PATH + dates[i] + ".json";
      var url = proxyBaseUrl + relPath;
      try {
        var resp = await fetch(url);
        if (!resp.ok) continue;
        var json = await resp.json();
        if (json && Array.isArray(json.values)) return json;
      } catch (e) {
        // Network error or JSON parse — try next candidate
      }
    }
    return null;
  }

  /* ------------------------------------------------------------------ */
  /*  Service-tag parsing                                                */
  /* ------------------------------------------------------------------ */
  function parseServiceTags(json) {
    var tagMap = {};

    Object.keys(TAG_PREFIXES).forEach(function (prefix) {
      tagMap[prefix] = { global: [], regions: {} };
    });

    (json.values || []).forEach(function (entry) {
      var name = entry.name || "";
      var props = entry.properties || {};
      var prefixes = props.addressPrefixes || [];
      if (prefixes.length === 0) return;

      Object.keys(TAG_PREFIXES).forEach(function (tp) {
        if (name === tp) {
          tagMap[tp].global = tagMap[tp].global.concat(prefixes);
        } else if (name.indexOf(tp + ".") === 0) {
          var region = name.substring(tp.length + 1);
          if (!tagMap[tp].regions[region]) tagMap[tp].regions[region] = [];
          tagMap[tp].regions[region] = tagMap[tp].regions[region].concat(prefixes);
        }
      });
    });

    return tagMap;
  }

  function collectCidrs(tagMap, selectedTags, selectedRegions, ipv4Only) {
    var cidrs = [];
    selectedTags.forEach(function (tag) {
      var data = tagMap[tag];
      if (!data) return;

      if (selectedRegions.length === 0) {
        cidrs = cidrs.concat(data.global);
      } else {
        selectedRegions.forEach(function (region) {
          if (data.regions[region]) cidrs = cidrs.concat(data.regions[region]);
        });
      }
    });

    if (ipv4Only) {
      cidrs = cidrs.filter(function (c) {
        return c.indexOf(":") === -1;
      });
    }

    return dedupe(cidrs).sort();
  }

  function allRegions(tagMap, selectedTags) {
    var set = Object.create(null);
    selectedTags.forEach(function (tag) {
      var data = tagMap[tag];
      if (!data) return;
      Object.keys(data.regions).forEach(function (r) {
        set[r] = true;
      });
    });
    return Object.keys(set).sort();
  }

  /* ------------------------------------------------------------------ */
  /*  Preview / instructions generator                                   */
  /* ------------------------------------------------------------------ */
  function buildInstructions(cidrs, opts) {
    var zoneId = opts.zoneId || "YOUR_ZONE_ID";
    var token = opts.token || "YOUR_API_TOKEN";
    var description = opts.description || defaultRuleDesc();
    var tags = opts.tags || [];
    var regions = opts.regions || [];
    var date = new Date().toISOString().split("T")[0];

    var chunks = chunkArray(cidrs, 250);

    var rulesJson = chunks.map(function (chunk, i) {
      return {
        description: description + (chunks.length > 1 ? " - chunk " + (i + 1) + "/" + chunks.length : ""),
        expression: "(ip.src in {" + chunk.join(" ") + "})",
        action: "skip",
        action_parameters: { ruleset: "current" },
        enabled: true
      };
    });

    var curlBody = JSON.stringify({ rules: rulesJson }, null, 2);

    var lines = [];
    lines.push("======================================================================");
    lines.push("  CLOUDFLARE FIREWALL RULES — Allow Power Automate");
    lines.push("  Generated " + date + " via Reenhanced Whitelister");
    lines.push("======================================================================");
    lines.push("");
    lines.push("WHAT THIS IS:");
    lines.push("Power Automate (Microsoft 365) needs to reach this website,");
    lines.push("but Cloudflare's firewall is blocking it. The rules below tell");
    lines.push("Cloudflare to allow traffic from Microsoft's IP addresses.");
    lines.push("");
    lines.push("SUMMARY:");
    lines.push("  Service           : " + tags.join(", "));
    lines.push("  Regions           : " + (regions.length > 0 ? regions.join(", ") : "Global (all regions)"));
    lines.push("  IP addresses      : " + cidrs.length);
    lines.push("  Firewall rules    : " + chunks.length);
    lines.push("");
    lines.push("----------------------------------------------------------------------");
    lines.push("  OPTION A — MANUAL STEPS (Cloudflare Dashboard)");
    lines.push("----------------------------------------------------------------------");
    lines.push("");
    lines.push("1. Log in to the Cloudflare Dashboard (https://dash.cloudflare.com).");
    lines.push("2. Select the zone (domain) for your website.");
    lines.push("3. Go to  Security > WAF > Custom rules.");
    lines.push("4. Click \"Create rule\".");
    lines.push("5. Name the rule: \"" + description + "\"");
    lines.push("6. Click \"Edit expression\" and paste the expression below:");
    lines.push("");
    chunks.forEach(function (chunk, i) {
      if (chunks.length > 1) lines.push("   --- Rule " + (i + 1) + " of " + chunks.length + " ---");
      lines.push("   (ip.src in {" + chunk.join(" ") + "})");
      lines.push("");
    });
    lines.push("7. Set the action to \"Skip\" and check \"Skip all remaining");
    lines.push("   custom rules\".");
    lines.push("8. Click \"Deploy\" to save the rule.");
    lines.push("");
    if (chunks.length > 1) {
      lines.push("   NOTE: Cloudflare limits expressions to ~4KB. This list");
      lines.push("   has been split into " + chunks.length + " rules. Repeat steps 4-8 for each.");
      lines.push("");
    }
    lines.push("----------------------------------------------------------------------");
    lines.push("  OPTION B — API COMMAND (for advanced users)");
    lines.push("----------------------------------------------------------------------");
    lines.push("");
    lines.push("Requirements:");
    lines.push("  - An API Token with Zone > Zone WAF permissions");
    lines.push("    Create one at: https://dash.cloudflare.com/profile/api-tokens");
    lines.push("  - Your Zone ID (found on the zone overview page)");
    lines.push("");
    lines.push("WARNING: This command replaces ALL custom firewall rules in the zone.");
    lines.push("If you have existing custom rules, read them first with a GET to the");
    lines.push("same URL and merge them into the rules array below.");
    lines.push("");
    lines.push("curl -X PUT \\");
    lines.push(
      "  \"https://api.cloudflare.com/client/v4/zones/" + zoneId + "/rulesets/phases/http_request_firewall_custom/entrypoint\" \\"
    );
    lines.push("  -H \"Authorization: Bearer " + token + "\" \\");
    lines.push("  -H \"Content-Type: application/json\" \\");
    lines.push("  -d '" + curlBody + "'");
    lines.push("");
    lines.push("----------------------------------------------------------------------");
    lines.push("  FULL LIST OF IP ADDRESSES");
    lines.push("----------------------------------------------------------------------");
    lines.push("");
    cidrs.forEach(function (c) {
      lines.push("  " + c);
    });
    lines.push("");
    lines.push("----------------------------------------------------------------------");
    lines.push("  IMPORTANT NOTES");
    lines.push("----------------------------------------------------------------------");
    lines.push("");
    lines.push("- Microsoft updates these IP addresses weekly. Re-run this tool");
    lines.push("  periodically (e.g., monthly) to keep your rules current.");
    lines.push("- Source: " + SERVICE_TAGS_DOWNLOAD);
    lines.push("- Tool: https://github.com/reenhanced/m365-cloudflare-whitelister");
    lines.push("");
    lines.push("======================================================================");
    lines.push("  Generated by Reenhanced — reenhanced.com");
    lines.push("  The world's best solution for WordPress + Microsoft");
    lines.push("======================================================================");

    return lines.join("\n");
  }

  /* ------------------------------------------------------------------ */
  /*  Cloudflare apply                                                   */
  /* ------------------------------------------------------------------ */
  function sanitizeRule(rule) {
    var o = {
      action: rule.action,
      description: rule.description,
      enabled: rule.enabled !== false,
      expression: rule.expression
    };
    if (rule.id) o.id = rule.id;
    if (rule.ref) o.ref = rule.ref;
    if (rule.action_parameters) o.action_parameters = rule.action_parameters;
    if (rule.logging) o.logging = rule.logging;
    if (rule.ratelimit) o.ratelimit = rule.ratelimit;
    return o;
  }

  async function applyToCloudflare(cidrs, cfToken, zoneId, description) {
    var url =
      "https://api.cloudflare.com/client/v4/zones/" +
      encodeURIComponent(zoneId) +
      "/rulesets/phases/http_request_firewall_custom/entrypoint";

    var headers = {
      Authorization: "Bearer " + cfToken,
      "Content-Type": "application/json"
    };

    var getResp = await fetch(url, { method: "GET", headers: headers });
    if (!getResp.ok) {
      throw new Error(
        "Cloudflare read failed (HTTP " +
          getResp.status +
          "). Verify token scope (Zone > Zone WAF), zone ID, and account access."
      );
    }
    var getJson = await getResp.json();
    if (!getJson.success) throw new Error("Cloudflare: " + JSON.stringify(getJson.errors));

    var existing = (getJson.result && getJson.result.rules) || [];
    var preserved = existing
      .filter(function (r) {
        return !(r.ref && r.ref.indexOf(RULE_REF_PREFIX) === 0);
      })
      .map(sanitizeRule);

    var chunks = chunkArray(cidrs, 250);
    var newRules = chunks.map(function (chunk, i) {
      return {
        ref: RULE_REF_PREFIX + "_" + (i + 1),
        description: description + (chunks.length > 1 ? " - chunk " + (i + 1) + "/" + chunks.length : ""),
        expression: "(ip.src in {" + chunk.join(" ") + "})",
        action: "skip",
        action_parameters: { ruleset: "current" },
        enabled: true
      };
    });

    var putResp = await fetch(url, {
      method: "PUT",
      headers: headers,
      body: JSON.stringify({ rules: preserved.concat(newRules) })
    });

    if (!putResp.ok)
      throw new Error("Cloudflare update failed (HTTP " + putResp.status + "). Ensure Zone > Zone WAF permissions.");
    var putJson = await putResp.json();
    if (!putJson.success) throw new Error("Cloudflare: " + JSON.stringify(putJson.errors));

    return { ruleCount: newRules.length, cidrCount: cidrs.length };
  }

  /* ------------------------------------------------------------------ */
  /*  Styles                                                             */
  /* ------------------------------------------------------------------ */
  function injectStyles(doc) {
    if (doc.getElementById("rh-pa-cf-style")) return;
    var s = doc.createElement("style");
    s.id = "rh-pa-cf-style";
    s.textContent = [
      /* ── Base ── */
      ".rh-w{--orange:#f2711c;--orange-hover:#e8590c;--dark:#1b1c1d;--blue:#2185d0;--blue-hover:#1678c2;--green:#21ba45;--green-hover:#16ab39;--text:rgba(0,0,0,.87);--muted:rgba(0,0,0,.6);--border:rgba(34,36,38,.15);--card-shadow:0 1px 2px 0 rgba(34,36,38,.15);--link:#4183c4;font-family:'Noto Sans','Helvetica Neue',Arial,Helvetica,sans-serif;max-width:100%;margin:0 auto;padding:0;color:var(--text);line-height:1.5;font-size:1rem}",

      /* ── Segments (cards) — matches Semantic UI .ui.segment ── */
      ".rh-card{background:#fff;border:1px solid var(--border);border-radius:.22222rem;padding:1.25em 1.25em;margin-bottom:1rem;box-shadow:var(--card-shadow)}",
      ".rh-card.raised{box-shadow:0 2px 4px 0 rgba(34,36,38,.12),0 2px 10px 0 rgba(34,36,38,.15)}",

      /* ── Hero — matches inverted masthead ── */
      ".rh-hero{background:var(--dark);color:rgba(255,255,255,.9);border:none;border-radius:.22222rem;text-align:center;padding:2em 1.5em 1.6em}",
      ".rh-hero h1{margin:0;font-size:1.6rem;font-weight:700;line-height:1.3;color:#fff}",
      ".rh-hero-sub{margin:.6rem 0 0;font-size:1.05rem;line-height:1.6;color:rgba(255,255,255,.8)}",
      ".rh-hero-icon{font-size:2.4rem;margin-bottom:.4rem}",
      ".rh-brand-sm{font-size:.82rem;margin-top:.8rem;color:rgba(255,255,255,.5);letter-spacing:.3px}",
      ".rh-brand-sm a{color:rgba(255,255,255,.6);text-decoration:underline}",
      ".rh-brand-sm a:hover{color:rgba(255,255,255,.9)}",

      /* ── Steps ── */
      ".rh-steps{counter-reset:rh-step}",
      ".rh-step{counter-increment:rh-step;position:relative;padding-left:2.8em}",
      ".rh-step::before{content:counter(rh-step);position:absolute;left:.2em;top:.2em;width:1.8em;height:1.8em;background:var(--orange);color:#fff;font-weight:700;font-size:.85rem;border-radius:50%;display:flex;align-items:center;justify-content:center}",
      ".rh-step-done::before{content:'\\2713';background:var(--green)}",
      ".rh-step h3{margin:0 0 .25em;font-size:1.1rem;font-weight:700;color:var(--dark)}",
      ".rh-step-hint{margin:0 0 .7em;font-size:.95rem;line-height:1.55;color:var(--muted)}",

      /* ── Buttons — matches Semantic UI .ui.button ── */
      ".rh-actions{display:flex;flex-wrap:wrap;gap:.5em;margin-top:.6em}",
      ".rh-btn{display:inline-block;padding:.6111em 1.5em;border:none;border-radius:.22222rem;font-weight:700;cursor:pointer;font-size:1rem;font-family:inherit;line-height:1em;text-align:center;transition:background-color .1s ease,opacity .1s ease,color .1s ease;box-shadow:0 0 0 1px transparent inset}",
      ".rh-btn:hover{opacity:.92}",
      ".rh-btn:active{opacity:.85}",
      ".rh-btn.primary{background:var(--orange);color:#fff}",
      ".rh-btn.primary:hover{background:var(--orange-hover);color:#fff}",
      ".rh-btn.big{padding:.8em 2em;font-size:1.1rem}",
      ".rh-btn.secondary{background:var(--blue);color:#fff}",
      ".rh-btn.secondary:hover{background:var(--blue-hover);color:#fff}",
      ".rh-btn.green{background:var(--green);color:#fff}",
      ".rh-btn.green:hover{background:var(--green-hover);color:#fff}",
      ".rh-btn.ghost{background:#e0e1e2;color:rgba(0,0,0,.6)}",
      ".rh-btn.ghost:hover{background:#cacbcd;color:rgba(0,0,0,.8)}",
      ".rh-btn:disabled{opacity:.45;cursor:default}",

      /* ── Status messages — matches Semantic UI .ui.message ── */
      ".rh-status{margin-top:.6em;padding:1em 1.2em;border-radius:.22222rem;font-size:.95rem;line-height:1.5;box-shadow:0 0 0 1px rgba(34,36,38,.22) inset}",
      ".rh-info{background:#f8ffff;color:#276f86;box-shadow:0 0 0 1px #a9d5de inset}",
      ".rh-success{background:#fcfff5;color:#2c662d;box-shadow:0 0 0 1px #a3c293 inset}",
      ".rh-error{background:#fff6f6;color:#9f3a38;box-shadow:0 0 0 1px #e0b4b4 inset}",
      ".rh-warn{background:#fffaf3;color:#573a08;box-shadow:0 0 0 1px #c9ba9b inset}",

      /* ── Form fields — matches Semantic UI .ui.form ── */
      ".rh-field{display:flex;flex-direction:column;gap:.28571rem;margin-bottom:.8em}",
      ".rh-field label{font-size:.92857em;font-weight:700;color:var(--text)}",
      ".rh-field input,.rh-field select{padding:.6em 1em;border:1px solid rgba(34,36,38,.15);border-radius:.22222rem;background:#fff;font-size:1em;font-family:inherit;color:var(--text);outline:none;transition:border-color .1s ease,box-shadow .1s ease}",
      ".rh-field input:focus,.rh-field select:focus{border-color:#85b7d9;box-shadow:0 0 0 0 rgba(34,36,38,.35) inset}",
      ".rh-row{display:flex;flex-wrap:wrap;gap:.4em}",

      /* ── Pills / checkboxes ── */
      ".rh-pill{display:inline-flex;align-items:center;gap:.4em;padding:.45em .7em;border:1px solid rgba(34,36,38,.15);border-radius:.22222rem;background:#fff;cursor:pointer;font-size:.9rem;transition:background .15s,border-color .15s}",
      ".rh-pill:hover{background:#f9fafb;border-color:rgba(34,36,38,.35)}",
      ".rh-pill input{margin:0}",

      /* ── Recommended badge — orange ── */
      ".rh-rec-badge{font-size:.72rem;background:var(--orange);color:#fff;padding:.15em .5em;border-radius:.22222rem;font-weight:700;margin-left:.3em}",

      /* ── Option cards ── */
      ".rh-opt-card{border:1px solid rgba(34,36,38,.15);border-radius:.22222rem;padding:.8em 1em;background:#fff;cursor:pointer;transition:border-color .15s,box-shadow .15s}",
      ".rh-opt-card:hover{border-color:rgba(34,36,38,.35);box-shadow:0 1px 3px 0 rgba(34,36,38,.12)}",
      ".rh-opt-card.selected{border-color:var(--blue);background:#f8ffff;box-shadow:0 0 0 1px var(--blue) inset}",
      ".rh-opt-card label{display:flex;align-items:flex-start;gap:.6em;cursor:pointer}",
      ".rh-opt-card input{margin-top:.25em;flex-shrink:0}",
      ".rh-opt-title{font-weight:700;font-size:.95rem;display:flex;align-items:center;gap:.3em;color:var(--dark)}",
      ".rh-opt-desc{font-size:.88rem;color:var(--muted);line-height:1.45;margin-top:.15em}",

      /* ── Drop zone ── */
      ".rh-drop{border:2px dashed rgba(34,36,38,.25);border-radius:.22222rem;padding:1.2em;text-align:center;cursor:pointer;background:#f9fafb;transition:border-color .15s,background .15s}",
      ".rh-drop.over{border-color:var(--orange);background:#fffaf3}",
      ".rh-drop input[type=file]{display:none}",

      /* ── Preview code block ── */
      ".rh-pre{font-family:Consolas,Monaco,'Courier New',monospace;font-size:.82rem;max-height:400px;overflow:auto;background:var(--dark);color:rgba(255,255,255,.9);border-radius:.22222rem;padding:1em;white-space:pre-wrap;word-break:break-all;line-height:1.55}",

      /* ── Callout ── */
      ".rh-callout{background:#f8ffff;border:1px solid #a9d5de;border-radius:.22222rem;padding:1em 1.2em;margin:.8em 0}",
      ".rh-callout-icon{font-size:1.2rem;margin-right:.4em}",
      ".rh-callout p{margin:0;font-size:.95rem;line-height:1.55;color:#276f86}",

      /* ── Note text ── */
      ".rh-note{font-size:.9rem;line-height:1.5;color:var(--muted)}",
      ".rh-note strong{color:var(--text)}",
      ".rh-note a{color:var(--link)}",
      ".rh-note a:hover{color:#1e70bf}",

      /* ── Region grid ── */
      ".rh-region-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:.4em}",

      /* ── Utility ── */
      ".rh-hidden{display:none}",
      ".rh-disabled{opacity:.45;pointer-events:none;user-select:none}",
      ".rh-btn.global-active{background:var(--blue);color:#fff}",
      ".rh-divider{border:none;border-top:1px solid var(--border);margin:.8em 0}",
      ".rh-cf-fields{display:grid;grid-template-columns:1fr 1fr;gap:.6em}",

      /* ── Details/summary ── */
      ".rh-details{margin-top:.6em}",
      ".rh-details summary{cursor:pointer;font-size:.95rem;font-weight:700;color:var(--link);padding:.3em 0;user-select:none}",
      ".rh-details summary:hover{color:#1e70bf;text-decoration:underline}",
      ".rh-details[open] summary{margin-bottom:.5em}",

      /* ── Copy banner ── */
      ".rh-copy-banner{display:flex;align-items:center;gap:.8em;background:#fcfff5;border:1px solid #a3c293;border-radius:.22222rem;padding:1em 1.2em;margin-top:.7em;box-shadow:0 0 0 1px #a3c293 inset}",
      ".rh-copy-banner-icon{font-size:1.8rem;flex-shrink:0}",
      ".rh-copy-banner p{margin:0;font-size:.95rem;line-height:1.5}",
      ".rh-copy-banner strong{color:#2c662d}",

      /* ── Footer ── */
      ".rh-footer{text-align:center;padding:1em 1.25em;background:#f9fafb}",
      ".rh-footer p{margin:0;font-size:.88rem;color:var(--muted);line-height:1.5}",
      ".rh-footer a{color:var(--link);text-decoration:none}",
      ".rh-footer a:hover{text-decoration:underline}",

      "@media(max-width:700px){.rh-cf-fields{grid-template-columns:1fr}.rh-hero h1{font-size:1.3rem}.rh-step{padding-left:2.4em}}"
    ].join("\n");
    doc.head.appendChild(s);
  }

  /* ------------------------------------------------------------------ */
  /*  Template                                                           */
  /* ------------------------------------------------------------------ */
  function buildTemplate() {
    var tagOptions = Object.keys(TAG_PREFIXES)
      .map(function (key) {
        var t = TAG_PREFIXES[key];
        var checked = t.recommended ? "checked" : "";
        var recClass = t.recommended ? " rh-opt-card selected" : " rh-opt-card";
        var badge = t.recommended ? '<span class="rh-rec-badge">Recommended</span>' : '';
        return (
          '<div class="' + recClass + '" data-tag="' + esc(key) + '">' +
          '<label><input class="rh-tag-cb" type="checkbox" value="' + esc(key) + '" ' + checked + '>' +
          '<div><div class="rh-opt-title">' + esc(t.label) + badge + '</div>' +
          '<div class="rh-opt-desc">' + esc(t.description) + '</div></div>' +
          '</label></div>'
        );
      })
      .join("");

    return [
      '<div class="rh-w">',

      /* Hero header — friendly, outcome-focused */
      '<div class="rh-card rh-hero">',
      '<div class="rh-hero-icon">\uD83D\uDEE1\uFE0F</div>',
      '<h1>Fix Power Automate + Cloudflare Issues</h1>',
      '<p class="rh-hero-sub">Is Cloudflare blocking Power Automate from reaching your website?<br>This tool creates the instructions your IT team needs to fix it.</p>',
      '<p class="rh-brand-sm">by <a href="https://reenhanced.com" target="_blank" rel="noopener">reenhanced</a> &mdash; WordPress + Microsoft</p>',
      '</div>',

      '<div class="rh-steps">',

      /* ── Step 1 — Load data (simple) ── */
      '<div class="rh-card rh-step" id="rh-step1">',
      '<h3>Get the latest IP addresses from Microsoft</h3>',
      '<p class="rh-step-hint">Microsoft publishes a list of IP addresses that Power Automate uses. We\'ll download the latest version so your IT team knows exactly which addresses to allow.</p>',
      '<div class="rh-actions">',
      '<button class="rh-btn primary big" id="rh-fetch-btn" type="button">\uD83D\uDD04 Get Latest IP Addresses</button>',
      '</div>',
      '<div class="rh-status rh-info" id="rh-file-status">Click the button above to get started.</div>',
      '<details class="rh-details"><summary>Having trouble? Upload the file manually</summary>',
      '<p class="rh-note" style="margin:.4rem 0">If the automatic download isn\'t working, you can <a href="' + esc(SERVICE_TAGS_DOWNLOAD) + '" target="_blank" rel="noopener">download the file directly from Microsoft</a> and drop it here.</p>',
      '<div class="rh-drop" id="rh-drop">',
      '<p><strong>Drag &amp; drop</strong> the downloaded file here, or <strong>click to browse</strong></p>',
      '<input type="file" id="rh-file" accept=".json,application/json">',
      '</div>',
      '</details>',
      '</div>',

      /* ── Step 2 — Choose what to allow (simplified) ── */
      '<div class="rh-card rh-step rh-disabled" id="rh-step2">',
      '<h3>Choose what to allow through Cloudflare</h3>',
      '<p class="rh-step-hint">For most people, the default selection below is all you need. Power Automate &amp; Connectors covers the traffic from your flows.</p>',
      '<div class="rh-field"><label>What should Cloudflare allow?</label>',
      '<div style="display:flex;flex-direction:column;gap:.45rem">' + tagOptions + '</div>',
      '</div>',
      '<p class="rh-note" id="rh-tag-desc">\u2705 <strong>Good choice.</strong> Power Automate &amp; Connectors is the right pick for most websites.</p>',

      /* Region selection — hidden in details for non-technical users */
      '<details class="rh-details"><summary>Narrow down by region (optional, advanced)</summary>',
      '<p class="rh-note" style="margin:0 0 .4rem">By default, IP addresses from <strong>all regions worldwide</strong> are included. This is the safest option. Only narrow this down if your IT team specifically asked you to.</p>',
      '<div class="rh-actions" style="margin:0 0 .3rem">',
      '<button class="rh-btn secondary" id="rh-global-btn" type="button">\u2733 All Regions (recommended)</button>',
      '<button class="rh-btn ghost" id="rh-sel-all-regions" type="button">Select all</button>',
      '<button class="rh-btn ghost" id="rh-clear-regions" type="button">Clear all</button>',
      '</div>',
      '<div class="rh-region-grid" id="rh-region-list"><span class="rh-note">Load IP addresses first (Step 1).</span></div>',
      '</details>',

      '<label class="rh-pill" style="margin-top:.4rem"><input id="rh-ipv4-only" type="checkbox" checked>IPv4 only (most common &mdash; recommended)</label>',
      '</div>',

      /* ── Step 3 — Preview & copy for IT team ── */
      '<div class="rh-card rh-step rh-disabled" id="rh-step3">',
      '<h3>Get the instructions for your IT team</h3>',
      '<p class="rh-step-hint">Click the button below to create ready-to-use instructions. You can copy them and email or message them to whoever manages your Cloudflare account.</p>',
      '<div class="rh-actions">',
      '<button class="rh-btn primary big" id="rh-preview-btn" type="button" disabled>\uD83D\uDCCB Create Instructions</button>',
      '</div>',
      '<div class="rh-status rh-info" id="rh-preview-status">Click "Create Instructions" to see the results.</div>',

      /* Copy banner — prominent, friendly */
      '<div class="rh-copy-banner rh-hidden" id="rh-copy-section">',
      '<div class="rh-copy-banner-icon">\uD83D\uDCE7</div>',
      '<div>',
      '<p><strong>Ready to send to your IT team!</strong> The full technical instructions are below. Click the button to copy everything, then paste it into an email or Teams message.</p>',
      '<div class="rh-actions" style="margin-top:.4rem">',
      '<button class="rh-btn green big" id="rh-copy-btn" type="button" disabled>\uD83D\uDCCB Copy Instructions to Clipboard</button>',
      '</div>',
      '</div>',
      '</div>',

      '<details class="rh-details" id="rh-preview-details" style="display:none"><summary>View full technical details</summary>',
      '<pre class="rh-pre" id="rh-preview-box"></pre>',
      '</details>',
      '</div>',

      /* ── Step 4 — Apply directly (advanced, collapsed by default) ── */
      '<div class="rh-card rh-step rh-disabled" id="rh-step4">',
      '<h3>Apply it yourself <span class="rh-note">(if you have Cloudflare access)</span></h3>',
      '<p class="rh-step-hint">Skip this step if you\'re sending the instructions to someone else. This is only for people who manage the Cloudflare account directly.</p>',
      '<details class="rh-details"><summary>I have Cloudflare access &mdash; let me apply directly</summary>',
      '<p class="rh-note" style="margin:0 0 .5rem">You\'ll need your Cloudflare <strong>API Token</strong> and <strong>Zone ID</strong>. If you don\'t know what those are, send the instructions from Step 3 to your IT team instead.</p>',
      '<div class="rh-cf-fields">',
      '<div class="rh-field"><label for="rh-cf-token">Cloudflare API Token</label><input id="rh-cf-token" type="password" placeholder="Paste your API token here" autocomplete="off"></div>',
      '<div class="rh-field"><label for="rh-zone-id">Zone ID</label><input id="rh-zone-id" type="text" placeholder="Found on your Cloudflare dashboard" autocomplete="off"></div>',
      '</div>',
      '<div class="rh-field"><label for="rh-rule-desc">Rule Name</label><input id="rh-rule-desc" type="text" value="' + esc(defaultRuleDesc()) + '"></div>',
      '<p class="rh-note"><strong>Need help?</strong> Your API Token needs <strong>Zone &gt; Zone WAF</strong> permissions. <a href="https://dash.cloudflare.com/profile/api-tokens" target="_blank" rel="noopener">Create one here</a>.</p>',
      '<div class="rh-actions">',
      '<button class="rh-btn primary" id="rh-apply-btn" type="button" disabled>Apply to Cloudflare Now</button>',
      '<button class="rh-btn ghost" id="rh-save-btn" type="button">Remember These Credentials</button>',
      '<button class="rh-btn ghost" id="rh-clear-btn" type="button">Forget Saved Credentials</button>',
      '</div>',
      '<div class="rh-status rh-hidden" id="rh-apply-status"></div>',
      '</details>',
      '</div>',

      '</div>', /* /steps */

      /* Reminder section — shown after instructions are generated */
      '<div class="rh-card rh-hidden" id="rh-reminder-section">',
      '<div style="display:flex;align-items:flex-start;gap:.8em">',
      '<div style="font-size:1.8rem;flex-shrink:0">\uD83D\uDCC5</div>',
      '<div>',
      '<p style="margin:0 0 .3em;font-weight:700;font-size:1rem;color:var(--dark)">Set a reminder to update these rules</p>',
      '<p class="rh-note" style="margin:0 0 .6em">Microsoft updates these IP addresses regularly. We recommend re-running this tool <strong>every month</strong> to keep your Cloudflare rules current. Add a reminder to your calendar so you don\u2019t forget.</p>',
      '<div class="rh-actions">',
      '<button class="rh-btn secondary" id="rh-reminder-btn" type="button">\uD83D\uDCC5 Add Monthly Reminder to Calendar</button>',
      '</div>',
      '<p class="rh-note" style="margin:.5em 0 0" id="rh-last-run-note"></p>',
      '</div>',
      '</div>',
      '</div>',

      /* How it works callout */
      '<div class="rh-callout">',
      '<p><span class="rh-callout-icon">\uD83D\uDD12</span> <strong>Your data stays private.</strong> Everything happens right here in your browser. No data is sent to any server except Microsoft (for the IP list) and Cloudflare (only if you choose to apply directly). This tool is <a href="https://github.com/reenhanced/m365-cloudflare-whitelister" target="_blank" rel="noopener">open-source</a>.</p>',
      '</div>',

      /* Footer */
      '<div class="rh-card rh-footer">',
      '<p>Built by <a href="https://reenhanced.com" target="_blank" rel="noopener">Reenhanced</a> &mdash; the world\'s best solution for WordPress + Microsoft.<br>Azure Service Tags are updated weekly. Run this tool again periodically to keep your rules current.</p>',
      '</div>',

      '</div>'
    ].join("");
  }

  /* ------------------------------------------------------------------ */
  /*  Mount & wire up                                                    */
  /* ------------------------------------------------------------------ */
  function mount(container, opts) {
    var proxyBaseUrl = (opts && opts.proxyBaseUrl) || "/api/ms-download/";
    injectStyles(container.ownerDocument);
    container.innerHTML = buildTemplate();

    /* State */
    var tagMap = null;
    var lastCidrs = null;
    var lastInstructions = null;

    /* DOM refs */
    var $step1 = container.querySelector("#rh-step1");
    var $fetchBtn = container.querySelector("#rh-fetch-btn");
    var $drop = container.querySelector("#rh-drop");
    var $file = container.querySelector("#rh-file");
    var $fileStatus = container.querySelector("#rh-file-status");
    var $tagCbs = container.querySelectorAll(".rh-tag-cb");
    var $tagDesc = container.querySelector("#rh-tag-desc");
    var $regionList = container.querySelector("#rh-region-list");
    var $globalBtn = container.querySelector("#rh-global-btn");
    var $selAll = container.querySelector("#rh-sel-all-regions");
    var $clearRegions = container.querySelector("#rh-clear-regions");
    var $ipv4 = container.querySelector("#rh-ipv4-only");
    var $previewBtn = container.querySelector("#rh-preview-btn");
    var $copyBtn = container.querySelector("#rh-copy-btn");
    var $copySection = container.querySelector("#rh-copy-section");
    var $previewStatus = container.querySelector("#rh-preview-status");
    var $previewBox = container.querySelector("#rh-preview-box");
    var $previewDetails = container.querySelector("#rh-preview-details");
    var $cfToken = container.querySelector("#rh-cf-token");
    var $zoneId = container.querySelector("#rh-zone-id");
    var $ruleDesc = container.querySelector("#rh-rule-desc");
    var $applyBtn = container.querySelector("#rh-apply-btn");
    var $saveBtn = container.querySelector("#rh-save-btn");
    var $clearBtn = container.querySelector("#rh-clear-btn");
    var $applyStatus = container.querySelector("#rh-apply-status");
    var $reminderSection = container.querySelector("#rh-reminder-section");
    var $reminderBtn = container.querySelector("#rh-reminder-btn");
    var $lastRunNote = container.querySelector("#rh-last-run-note");

    /* Tag option card click highlights */
    var $optCards = container.querySelectorAll(".rh-opt-card");
    $optCards.forEach(function (card) {
      var cb = card.querySelector(".rh-tag-cb");
      card.addEventListener("click", function (e) {
        if (e.target !== cb) {
          cb.checked = !cb.checked;
          cb.dispatchEvent(new Event("change", { bubbles: true }));
        }
        card.classList.toggle("selected", cb.checked);
      });
      cb.addEventListener("change", function () {
        card.classList.toggle("selected", cb.checked);
      });
    });

    /* Restore saved credentials */
    var saved = readSettings();
    if (saved) {
      $cfToken.value = saved.cfToken || "";
      $zoneId.value = saved.zoneId || "";
      if (saved.ruleDesc) $ruleDesc.value = saved.ruleDesc;
    }

    /* Show stale-data warning if last run was >30 days ago */
    var _daysSince = daysSinceLastRun();
    if (_daysSince !== null && _daysSince > 30) {
      $fileStatus.className = "rh-status rh-warn";
      $fileStatus.innerHTML = "\u26A0\uFE0F You last updated your Cloudflare rules <strong>" + _daysSince + " days ago</strong>. Microsoft updates their IP addresses regularly &mdash; click the button above to fetch the latest list and update your rules.";
    } else if (_daysSince !== null) {
      $lastRunNote.textContent = "You last ran this tool " + _daysSince + " day" + (_daysSince !== 1 ? "s" : "") + " ago (" + getLastRun() + ").";
    }

    /* Wire up reminder button */
    $reminderBtn.addEventListener("click", function () {
      downloadICS();
      $reminderBtn.textContent = "\u2705 Reminder Downloaded!";
      setTimeout(function () { $reminderBtn.textContent = "\uD83D\uDCC5 Add Monthly Reminder to Calendar"; }, 3000);
    });

    function showReminderSection() {
      $reminderSection.classList.remove("rh-hidden");
      recordLastRun();
      $lastRunNote.textContent = "Last run: " + todayDate() + ". We\u2019ll remind you if it\u2019s been more than 30 days.";
    }

    /* ---- Enable steps after data load ---- */
    function enableSteps() {
      ['#rh-step2','#rh-step3','#rh-step4'].forEach(function (sel) {
        var el = container.querySelector(sel);
        if (el) el.classList.remove('rh-disabled');
      });
    }

    /* ---- Shared data loading ---- */
    function applyJson(json) {
      if (!json.values || !Array.isArray(json.values)) throw new Error("Invalid format");
      tagMap = parseServiceTags(json);
      $fileStatus.className = "rh-status rh-success";
      $fileStatus.textContent =
        "\u2705 Got it! Found " + json.values.length + " entries covering all of Microsoft's services.";
      $step1.classList.add("rh-step-done");
      refreshRegions();
      enableSteps();
      $previewBtn.disabled = false;
      $applyBtn.disabled = false;
    }

    /* ---- Auto-fetch ---- */
    $fetchBtn.addEventListener("click", async function () {
      $fetchBtn.disabled = true;
      $fileStatus.className = "rh-status rh-info";
      $fileStatus.textContent = "Downloading the latest IP addresses from Microsoft\u2026 This may take a moment.";
      try {
        var json = await autoFetchServiceTags(proxyBaseUrl);
        if (!json) throw new Error("Couldn\u2019t download the IP list automatically. Try the manual upload option below.");
        applyJson(json);
      } catch (e) {
        $fileStatus.className = "rh-status rh-error";
        $fileStatus.textContent = e.message;
      } finally {
        $fetchBtn.disabled = false;
      }
    });

    /* ---- File loading (manual fallback) ---- */
    function handleFile(file) {
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function () {
        try {
          applyJson(JSON.parse(reader.result));
        } catch (e) {
          $fileStatus.className = "rh-status rh-error";
          $fileStatus.textContent = "Error: " + e.message;
        }
      };
      reader.readAsText(file);
    }

    $drop.addEventListener("click", function () {
      $file.click();
    });
    $file.addEventListener("change", function () {
      handleFile($file.files[0]);
    });
    $drop.addEventListener("dragover", function (e) {
      e.preventDefault();
      $drop.classList.add("over");
    });
    $drop.addEventListener("dragleave", function () {
      $drop.classList.remove("over");
    });
    $drop.addEventListener("drop", function (e) {
      e.preventDefault();
      $drop.classList.remove("over");
      handleFile(e.dataTransfer.files[0]);
    });

    /* ---- Tag selection ---- */
    function selectedTagKeys() {
      var keys = [];
      $tagCbs.forEach(function (cb) {
        if (cb.checked) keys.push(cb.value);
      });
      return keys;
    }

    function refreshRegions() {
      if (!tagMap) return;
      var tags = selectedTagKeys();
      var regions = allRegions(tagMap, tags);
      if (regions.length === 0) {
        $regionList.innerHTML = '<span class="rh-note">No regional data for selected tags.</span>';
        return;
      }
      $regionList.innerHTML = regions
        .map(function (r) {
          return '<label class="rh-pill"><input class="rh-region-cb" type="checkbox" value="' + esc(r) + '">' + esc(r) + '</label>';
        })
        .join("");
    }

    function updateTagDesc() {
      var keys = selectedTagKeys();
      if (keys.length === 0) {
        $tagDesc.innerHTML = "\u26A0\uFE0F Select at least one option above.";
        return;
      }
      var hasConnectors = keys.indexOf("AzureConnectors") !== -1;
      var hasCloud = keys.indexOf("AzureCloud") !== -1;
      if (hasConnectors && !hasCloud) {
        $tagDesc.innerHTML = "\u2705 <strong>Good choice.</strong> This covers Power Automate, Logic Apps, and related connectors.";
      } else if (hasConnectors && hasCloud) {
        $tagDesc.innerHTML = "\u26A0\uFE0F <strong>Heads up:</strong> \"All Azure Services\" adds a very large number of IP addresses. Only include it if your IT team specifically asked for it.";
      } else if (hasCloud) {
        $tagDesc.innerHTML = "\u26A0\uFE0F <strong>Are you sure?</strong> This is a very broad set. Most people only need \"Power Automate &amp; Connectors\" above.";
      }
    }

    $tagCbs.forEach(function (cb) {
      cb.addEventListener("change", function () {
        refreshRegions();
        updateTagDesc();
      });
    });

    function updateGlobalBtnState() {
      var checked = container.querySelectorAll(".rh-region-cb:checked");
      if (checked.length === 0) {
        $globalBtn.classList.add("global-active");
      } else {
        $globalBtn.classList.remove("global-active");
      }
    }
    // Initial state: global is active (no regions selected)
    $globalBtn.classList.add("global-active");

    function getSelectedRegions() {
      var all = container.querySelectorAll(".rh-region-cb");
      var checked = container.querySelectorAll(".rh-region-cb:checked");
      // If all regions are selected, treat as global (empty array)
      if (all.length > 0 && checked.length === all.length) return [];
      var regions = [];
      checked.forEach(function (cb) { regions.push(cb.value); });
      return regions;
    }

    // Use event delegation on the region grid so dynamically added checkboxes work
    $regionList.addEventListener("change", function (e) {
      if (e.target && e.target.classList.contains("rh-region-cb")) {
        updateGlobalBtnState();
      }
    });

    $globalBtn.addEventListener("click", function () {
      container.querySelectorAll(".rh-region-cb").forEach(function (cb) {
        cb.checked = false;
      });
      updateGlobalBtnState();
    });
    $selAll.addEventListener("click", function () {
      container.querySelectorAll(".rh-region-cb").forEach(function (cb) {
        cb.checked = true;
      });
      updateGlobalBtnState();
    });
    $clearRegions.addEventListener("click", function () {
      container.querySelectorAll(".rh-region-cb").forEach(function (cb) {
        cb.checked = false;
      });
      updateGlobalBtnState();
    });

    /* ---- Preview ---- */
    $previewBtn.addEventListener("click", function () {
      if (!tagMap) return;
      var tags = selectedTagKeys();
      if (tags.length === 0) {
        $previewStatus.className = "rh-status rh-error";
        $previewStatus.textContent = "Please select at least one option in Step 2 above.";
        return;
      }
      var regions = getSelectedRegions();
      var ipv4Only = $ipv4.checked;
      var cidrs = collectCidrs(tagMap, tags, regions, ipv4Only);
      if (cidrs.length === 0) {
        $previewStatus.className = "rh-status rh-error";
        $previewStatus.textContent = "No IP addresses matched your selection. Try selecting different options in Step 2.";
        return;
      }
      lastCidrs = cidrs;
      lastInstructions = buildInstructions(cidrs, {
        zoneId: $zoneId.value.trim() || undefined,
        token: $cfToken.value.trim() || undefined,
        description: $ruleDesc.value.trim() || undefined,
        tags: tags,
        regions: regions
      });
      $previewBox.textContent = lastInstructions;
      $previewDetails.style.display = "";
      $copyBtn.disabled = false;
      $copySection.classList.remove("rh-hidden");

      // Mark step 3 done
      container.querySelector("#rh-step3").classList.add("rh-step-done");
      showReminderSection();

      var warn = "";
      if (tags.indexOf("AzureCloud") !== -1 && cidrs.length > 500) {
        warn = " Note: You\u2019ve included All Azure Services, which adds a lot of addresses (" + cidrs.length + "). That\u2019s fine, but you could reduce this by picking specific regions.";
      }

      $previewStatus.className = "rh-status rh-success";
      $previewStatus.textContent = "\u2705 Instructions ready! " + cidrs.length + " IP addresses to allow." + warn;
    });

    /* ---- Copy ---- */
    $copyBtn.addEventListener("click", function () {
      if (!lastInstructions) return;
      copyText(lastInstructions).then(function () {
        $previewStatus.className = "rh-status rh-success";
        $previewStatus.textContent = "\uD83C\uDF89 Copied! Paste this into an email or Teams message to your IT team.";
        $copyBtn.textContent = "\u2705 Copied!";
        setTimeout(function () { $copyBtn.textContent = "\uD83D\uDCCB Copy Instructions to Clipboard"; }, 3000);
      });
    });

    /* ---- Apply ---- */
    $applyBtn.addEventListener("click", async function () {
      var token = $cfToken.value.trim();
      var zoneId = $zoneId.value.trim();
      var desc = $ruleDesc.value.trim() || defaultRuleDesc();

      if (!token || !zoneId) {
        $applyStatus.className = "rh-status rh-error";
        $applyStatus.classList.remove("rh-hidden");
        $applyStatus.textContent = "Please enter both your API Token and Zone ID above. If you don\u2019t have these, send the instructions from Step 3 to your IT team instead.";
        return;
      }

      var tags = selectedTagKeys();
      var regions = getSelectedRegions();
      var cidrs = collectCidrs(tagMap, tags, regions, $ipv4.checked);
      if (cidrs.length === 0) {
        $applyStatus.className = "rh-status rh-error";
        $applyStatus.classList.remove("rh-hidden");
        $applyStatus.textContent = "No IP addresses to apply. Check your selections in Step 2.";
        return;
      }

      $applyStatus.className = "rh-status rh-info";
      $applyStatus.classList.remove("rh-hidden");
      $applyStatus.textContent = "Applying " + cidrs.length + " IP addresses to your Cloudflare account\u2026 Please wait.";
      $applyBtn.disabled = true;

      try {
        var result = await applyToCloudflare(cidrs, token, zoneId, desc);
        $applyStatus.className = "rh-status rh-success";
        $applyStatus.textContent =
          "\uD83C\uDF89 Done! " + result.ruleCount + " rule(s) created covering " + result.cidrCount + " IP addresses. Power Automate should now be able to reach your website.";
        showReminderSection();
      } catch (e) {
        $applyStatus.className = "rh-status rh-error";
        $applyStatus.textContent = e.message;
      } finally {
        $applyBtn.disabled = false;
      }
    });

    /* ---- Save / clear credentials ---- */
    $saveBtn.addEventListener("click", function () {
      writeSettings({
        cfToken: $cfToken.value.trim(),
        zoneId: $zoneId.value.trim(),
        ruleDesc: $ruleDesc.value.trim()
      });
      $applyStatus.className = "rh-status rh-success";
      $applyStatus.classList.remove("rh-hidden");
      $applyStatus.textContent = "Credentials saved to this browser. They\u2019ll be here next time you visit.";
    });

    $clearBtn.addEventListener("click", function () {
      clearSettings();
      $cfToken.value = "";
      $zoneId.value = "";
      $applyStatus.className = "rh-status rh-success";
      $applyStatus.classList.remove("rh-hidden");
      $applyStatus.textContent = "Saved credentials have been removed from this browser.";
    });
  }

  /* ------------------------------------------------------------------ */
  /*  Public API                                                         */
  /* ------------------------------------------------------------------ */
  window.ReenhancedM365CFWidget = {
    init: function (opts) {
      var target = (opts && opts.target) || "#reenhanced-m365-cf-widget";
      var el = typeof target === "string" ? document.querySelector(target) : target;
      if (!el) throw new Error("ReenhancedM365CFWidget: container not found (" + target + ")");
      mount(el, opts);
    }
  };
})();
