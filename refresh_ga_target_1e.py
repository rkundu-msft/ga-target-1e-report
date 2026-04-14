#!/usr/bin/env python3
"""
Refresh GA_Target_1E LT Review report from ADO using az CLI auth.
Query: Work Item Type NOT IN (Epic), State = [Any], Tags Contains GA_Target_1E
"""

import json
import os
import html
import subprocess
import sys
from urllib.request import Request, urlopen
from datetime import datetime
from collections import defaultdict

# ADO config
ORG = "domoreexp"
PROJECT = "MSTeams"
AZ_CMD = r"C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd"

# KR color map
KR_COLORS = {
    "Fundamentals and Craft": "#4A90D9",
    "Zoom compete/sales blockers": "#E74C3C",
    "CSAT": "#27AE60",
    "Product-led growth and Monetization": "#F39C12",
    "Intelligence for Events": "#8E44AD",
    "High Impact Meeting": "#2C3E50",
}

# State -> color
STATE_COLORS = {
    "Closed": "#27ae60",
    "Resolved": "#8e44ad",
    "Active": "#2980b9",
    "RollingOut": "#f39c12",
    "Proposed": "#95a5a6",
    "New": "#3498db",
    "Design": "#16a085",
    "Cut": "#e74c3c",
    "Removed": "#e74c3c",
}

# KR display order
KR_ORDER = [
    "Fundamentals and Craft",
    "Zoom compete/sales blockers",
    "CSAT",
    "Product-led growth and Monetization",
    "Intelligence for Events",
    "High Impact Meeting",
]


def get_bearer_token():
    """Get ADO token. Uses ADO_PAT env var if set (CI), otherwise az CLI (local)."""
    pat = os.environ.get("ADO_PAT", "").strip()
    if pat:
        # In CI: return PAT as basic auth (handled differently in api calls)
        return ("pat", pat)
    result = subprocess.run(
        [r"C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd",
         "account", "get-access-token",
         "--resource", "499b84ac-1321-427f-aa17-267ca6975798",
         "--query", "accessToken", "-o", "tsv"],
        capture_output=True, text=True, timeout=30, shell=True
    )
    if result.returncode != 0:
        print(f"  Token error: {result.stderr}", file=sys.stderr)
        sys.exit(1)
    return ("bearer", result.stdout.strip())


def _auth_header(token):
    """Return auth header based on token type."""
    import base64
    auth_type, auth_val = token
    if auth_type == "pat":
        encoded = base64.b64encode(f":{auth_val}".encode()).decode()
        return f"Basic {encoded}"
    return f"Bearer {auth_val}"


def api_get(url, token):
    """GET request to ADO API."""
    req = Request(url, headers={
        "Authorization": _auth_header(token),
        "Content-Type": "application/json"
    })
    with urlopen(req, timeout=120) as resp:
        return json.loads(resp.read().decode())


def api_post(url, token, data):
    """POST request to ADO API."""
    req = Request(url, data=json.dumps(data).encode(), headers={
        "Authorization": _auth_header(token),
        "Content-Type": "application/json"
    }, method="POST")
    with urlopen(req, timeout=120) as resp:
        return json.loads(resp.read().decode())


GH_REPO = "rkundu-msft/ga-target-1e-report"
GH_COMMITMENTS_FILE = "commitments.json"


def sync_github_to_ado(token):
    """Read commitments.json from GitHub and push values to ADO."""
    import base64 as b64
    print("  Checking GitHub commitments.json for browser edits...")
    try:
        req = Request(
            f"https://api.github.com/repos/{GH_REPO}/contents/{GH_COMMITMENTS_FILE}",
            headers={"Accept": "application/vnd.github.v3+json"}
        )
        with urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode())
        commitments = json.loads(b64.b64decode(data["content"]).decode())
    except Exception as e:
        print(f"  Could not fetch commitments.json: {e}")
        return 0

    if not commitments:
        print("  No browser edits to sync.")
        return 0

    print(f"  Found {len(commitments)} browser edits. Syncing to ADO...")

    # Fetch current ADO values for these items
    valid_ids = [int(k) for k in commitments.keys() if k.isdigit()]
    if not valid_ids:
        return 0

    ids_str = ",".join(map(str, valid_ids))
    try:
        result = api_get(
            f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/wit/workitems?ids={ids_str}"
            f"&fields=System.Id,Custom.CommittedTargettedCut,Custom.CustomerImpacting&api-version=7.1",
            token
        )
    except Exception:
        print("  Failed to fetch items from ADO.")
        return 0

    ado_map = {}
    ci_map = {}
    for item in result.get("value", []):
        wid = item["id"]
        ado_map[wid] = item["fields"].get("Custom.CommittedTargettedCut", "") or ""
        ci_map[wid] = item["fields"].get("Custom.CustomerImpacting", "") or ""

    updated = 0
    for wid_str, new_val in commitments.items():
        wid = int(wid_str)
        if new_val not in ("Committed", "Targeted", "Cut"):
            continue
        if ado_map.get(wid) == new_val:
            continue  # already matches
        # Patch ADO
        ops = [{"op": "replace", "path": "/fields/Custom.CommittedTargettedCut", "value": new_val}]
        if not ci_map.get(wid):
            ops.insert(0, {"op": "add", "path": "/fields/Custom.CustomerImpacting", "value": "No"})
        try:
            patch_url = f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/wit/workitems/{wid}?api-version=7.1"
            req = Request(patch_url, data=json.dumps(ops).encode(), headers={
                "Authorization": _auth_header(token),
                "Content-Type": "application/json-patch+json"
            }, method="PATCH")
            with urlopen(req, timeout=30) as resp:
                if resp.status == 200:
                    updated += 1
                    print(f"    {wid} -> {new_val}")
        except Exception as e:
            print(f"    {wid} FAILED: {e}")

    print(f"  Synced {updated} items to ADO.")
    return updated


def query_work_items(token):
    """Run WIQL query matching the screenshot filters."""
    wiql = {
        "query": "SELECT [System.Id] FROM WorkItems WHERE [System.WorkItemType] NOT IN ('Epic') AND [System.Tags] CONTAINS 'GA_Target_1E' ORDER BY [System.Id] ASC"
    }
    url = f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/wit/wiql?api-version=7.1"
    result = api_post(url, token, wiql)
    return [item["id"] for item in result.get("workItems", [])]


def get_work_items_batch(token, ids):
    """Fetch work item details in batches of 200."""
    all_items = []
    fields = "System.Id,System.Title,System.WorkItemType,System.State,System.AssignedTo,System.Tags,System.Parent,Custom.CommittedTargettedCut"
    for i in range(0, len(ids), 200):
        batch = ids[i:i+200]
        ids_str = ",".join(map(str, batch))
        url = f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/wit/workitems?ids={ids_str}&fields={fields}&api-version=7.1"
        result = api_get(url, token)
        all_items.extend(result.get("value", []))
    return all_items


# Ground-truth ID->KR mapping from the validated reference report (GA_Target_View.html)
KNOWN_KR_MAP = {
    # Fundamentals and Craft
    4891487: "Fundamentals and Craft", 4916512: "Fundamentals and Craft",
    4971194: "Fundamentals and Craft", 4971219: "Fundamentals and Craft",
    4973775: "Fundamentals and Craft", 4981209: "Fundamentals and Craft",
    5031186: "Fundamentals and Craft", 5031190: "Fundamentals and Craft",
    5031196: "Fundamentals and Craft", 5031233: "Fundamentals and Craft",
    5031237: "Fundamentals and Craft", 5031277: "Fundamentals and Craft",
    5031278: "Fundamentals and Craft", 5089662: "Fundamentals and Craft",
    5101991: "Fundamentals and Craft", 5102071: "Fundamentals and Craft",
    5102927: "Fundamentals and Craft", 5117888: "Fundamentals and Craft",
    5121417: "Fundamentals and Craft", 5121418: "Fundamentals and Craft",
    5124130: "Fundamentals and Craft", 5148706: "Fundamentals and Craft",
    5148709: "Fundamentals and Craft", 5148723: "Fundamentals and Craft",
    5148767: "Fundamentals and Craft", 5153271: "Fundamentals and Craft",
    5153279: "Fundamentals and Craft",
    5153283: "CSAT",  # Moved from F&C - Room Role location field UX
    5190315: "Fundamentals and Craft", 5190317: "Fundamentals and Craft",
    5190320: "Fundamentals and Craft", 5283339: "Fundamentals and Craft",
    5283341: "Fundamentals and Craft",
    # Zoom compete/sales blockers
    4098900: "Zoom compete/sales blockers", 4098904: "Zoom compete/sales blockers",
    4709134: "Zoom compete/sales blockers", 4892306: "Zoom compete/sales blockers",
    5151010: "Zoom compete/sales blockers",
    # CSAT
    4707636: "CSAT", 4709076: "CSAT", 4709192: "CSAT", 4891483: "CSAT",
    4897148: "CSAT", 4949434: "CSAT", 4983592: "CSAT", 4983823: "CSAT",
    5024216: "CSAT", 5031205: "CSAT", 5099678: "CSAT", 5099689: "CSAT",
    5109892: "CSAT", 5118590: "CSAT", 5118591: "CSAT", 5121632: "CSAT",
    5124118: "CSAT", 5124128: "CSAT", 5124134: "CSAT", 5124173: "CSAT",
    5124230: "CSAT", 5148749: "CSAT", 5153263: "CSAT", 5153281: "CSAT",
    5153284: "CSAT", 5153294: "CSAT", 5153313: "CSAT", 5153317: "CSAT",
    5190302: "CSAT", 5190304: "CSAT", 5190308: "CSAT", 5190310: "CSAT",
    5190323: "CSAT", 5190328: "CSAT", 5190330: "CSAT", 5190904: "CSAT",
    5190940: "CSAT", 5254770: "CSAT",
    # Product-led growth and Monetization
    4709062: "Product-led growth and Monetization",
    4863951: "Product-led growth and Monetization",
    4892328: "Product-led growth and Monetization",
    4916482: "Product-led growth and Monetization",
    4916630: "Product-led growth and Monetization",
    4935047: "Product-led growth and Monetization",
    4935048: "Product-led growth and Monetization",
    5016533: "Product-led growth and Monetization",
    5121416: "Product-led growth and Monetization",
    5124135: "Product-led growth and Monetization",
    5153290: "Product-led growth and Monetization",
    5153291: "Product-led growth and Monetization",
    5153295: "Product-led growth and Monetization",
    5190313: "Product-led growth and Monetization",
    # Intelligence for Events
    5153315: "Intelligence for Events",
    # High Impact Meeting
    5124138: "High Impact Meeting",
    5148713: "CSAT",  # Moved from HIM - COS UX feedback
    5148762: "CSAT",  # Moved from HIM - COS UX feedback
    5148766: "High Impact Meeting",
    5153265: "High Impact Meeting", 5153272: "High Impact Meeting",
    5153273: "High Impact Meeting", 5190311: "High Impact Meeting",
    5190325: "High Impact Meeting",
    # --- New items added after reference report (Apr 2026 refresh) ---
    # Fundamentals and Craft - reliability bugs, monitoring, infra
    4272661: "Fundamentals and Craft",  # create_virtual_event failures
    5152869: "Fundamentals and Craft",  # manage load failure
    5152910: "Fundamentals and Craft",  # Create event failure for Broadcast
    5152939: "Fundamentals and Craft",  # Duplicate event failure
    5153526: "Fundamentals and Craft",  # Event landing failures
    5170445: "Fundamentals and Craft",  # RecorderOption MO failure
    5202331: "Fundamentals and Craft",  # Users Not Found in Graph
    5202334: "Fundamentals and Craft",  # Operation Timeout
    5202340: "Fundamentals and Craft",  # On-Prem Exchange mailbox
    5202342: "Fundamentals and Craft",  # Start time in past
    5228957: "Fundamentals and Craft",  # Disable rule action MO
    5241677: "Fundamentals and Craft",  # Discover e2e reliability
    5247373: "Fundamentals and Craft",  # VES Attendee Split On-Prem
    5251279: "Fundamentals and Craft",  # Update event TransientExchangeId
    5251797: "Fundamentals and Craft",  # EnableParticipantRenaming
    5251834: "Fundamentals and Craft",  # Sensitivity Label failure
    5254862: "Fundamentals and Craft",  # Registration limit exceed 1000
    5254886: "Fundamentals and Craft",  # Discover phased rendering timeout
    5270446: "Fundamentals and Craft",  # VES Timeout Discover
    5270863: "Fundamentals and Craft",  # VES Discover PreconditionFailed
    5289759: "Fundamentals and Craft",  # Access Forbidden reliability
    5291136: "Fundamentals and Craft",  # Guardian monitor for create/update
    5297279: "Fundamentals and Craft",  # create_virtual_event RequireStatusFailed
    # CSAT
    5202370: "CSAT",                    # Custom Template feature not allowed
    5254934: "CSAT",                    # Event updates to co-org/presenters
    5312149: "CSAT",                    # Upgrade experience improvements
    3856956: "CSAT",                    # Send separate invites for Attendees
    # --- Apr 14 refresh ---
    4709061: "CSAT",                              # RSVP support on Discover/Attend/Event landing
    5222196: "Product-led growth and Monetization",# Related events suggestion - drive user funnel
    5222284: "Product-led growth and Monetization",# Improved default events in discovery
    5222286: "Product-led growth and Monetization",# Filters in Teams Discover
    5222304: "Product-led growth and Monetization",# Improved share experience across channels
    5222315: "CSAT",                              # Auto Publish of recording
    5222318: "CSAT",                              # Attendee CTA on past events for Recap
    5286747: "Fundamentals and Craft",            # Rename Meet -> Events
}


def classify_kr(item_id, title, tags):
    """
    Classify into one of the 6 Events KRs.
    Uses ground-truth lookup for known items, keyword fallback for new ones.
    """
    # Known item - use validated mapping
    if item_id in KNOWN_KR_MAP:
        return KNOWN_KR_MAP[item_id]

    # Fallback: keyword classifier for new items
    t = (title + " " + tags).lower()

    # Intelligence for Events (narrow, check first)
    if any(kw in t for kw in ["ai-powered", "ai powered", "intelligence", "copilot",
                               "recap", "ai ros", "faq handling", "ai response"]):
        return "Intelligence for Events"

    # High Impact Meeting (specific)
    if any(kw in t for kw in ["broadcast", "collaborative", "session type",
                               "green room", "rtmp", "screen setup",
                               "managed mode", "facilitator", "producer"]):
        return "High Impact Meeting"

    # Zoom compete/sales blockers
    if any(kw in t for kw in ["100k", "100,000", "20k registration",
                               "registration scale", "pre-load video",
                               "breakout xl", "vdi th",
                               "performance and reliability testing"]):
        return "Zoom compete/sales blockers"

    # Product-led growth and Monetization
    if any(kw in t for kw in ["plg", "upsell", "calendar flow", "entry point",
                               "add to calendar", "monetization", "tpre",
                               "fre coachmark", "pin meet", "empty state",
                               "community", "discoverability", "grouping event",
                               "attendee pack", "surface event",
                               "search capability in discover",
                               "filters to sift", "filtering and grouping"]):
        return "Product-led growth and Monetization"

    # CSAT (broad)
    if any(kw in t for kw in ["template", "co-org", "shared mailbox", "email",
                               "invite", "registration", "polls", "q&a lifecycle",
                               "room availability", "notify", "restrict forwarding",
                               "ics file", "rsvp", "location field", "capacity",
                               "approval", "people picker", "sensitivity",
                               "cancel", "manage event", "forward invite",
                               "propose", "recurring", "export", "multi-session",
                               "in person", "offline", "room alias", "room finder",
                               "thankyou", "delegate", "survey", "logo",
                               "custom properties", "email template",
                               "co-organizer", "pop out", "multi window",
                               "terminology", "eventify", "zoom feature for banner",
                               "specific people", "people and group",
                               "banner image.*email", "toggling off"]):
        return "CSAT"

    # Fundamentals and Craft (default)
    return "Fundamentals and Craft"


def compute_stats(kr_groups):
    """Compute cutline stats for leadership view."""
    from collections import Counter
    stats = {}

    # Global counters
    all_items = [item for items in kr_groups.values() for item in items]
    total = len(all_items)
    type_counts = Counter(i["type"] for i in all_items)
    state_counts = Counter(i["state"] for i in all_items)

    closed_states = {"Closed", "Resolved", "RollingOut"}
    done = sum(1 for i in all_items if i["state"] in closed_states)
    active = sum(1 for i in all_items if i["state"] == "Active")
    proposed = total - done - active

    stats["total"] = total
    stats["type_counts"] = dict(type_counts.most_common())
    stats["state_counts"] = dict(state_counts.most_common())
    stats["done"] = done
    stats["active"] = active
    stats["proposed"] = proposed
    stats["pct_done"] = round(done * 100 / total) if total else 0

    # Per-KR breakdown
    kr_stats = {}
    for kr, items in kr_groups.items():
        tc = Counter(i["type"] for i in items)
        sc = Counter(i["state"] for i in items)
        d = sum(1 for i in items if i["state"] in closed_states)
        kr_stats[kr] = {
            "total": len(items),
            "type_counts": dict(tc.most_common()),
            "state_counts": dict(sc.most_common()),
            "done": d,
            "pct_done": round(d * 100 / len(items)) if items else 0,
        }
    stats["kr_stats"] = kr_stats
    return stats


def generate_html(kr_groups, today_str, stats):
    """Generate the HTML report."""
    css = """
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, sans-serif; background: #f0f2f5; color: #1a1a1a; padding: 30px; }
  .header { text-align: center; margin-bottom: 8px; }
  .header h1 { font-size: 22px; margin-bottom: 4px; }
  .header p { color: #666; font-size: 13px; margin-bottom: 20px; }
  .section-title { font-size: 18px; font-weight: 700; margin: 30px auto 12px; max-width: 1300px; padding-bottom: 6px; border-bottom: 2px solid #ddd; }

  .summary-table { max-width: 1300px; margin: 0 auto 36px; border-collapse: collapse; width: 100%; background: #fff; border-radius: 10px; overflow: hidden; box-shadow: 0 1px 6px rgba(0,0,0,0.08); }
  .summary-table th { background: #1a1a1a; color: #fff; padding: 12px 16px; text-align: left; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; }
  .summary-table td { padding: 12px 16px; border-bottom: 1px solid #f0f0f0; font-size: 13px; }
  .summary-table tr:last-child td { border-bottom: none; }
  .summary-table tr:hover td { background: #f9fbfd; }
  .summary-table .kr-name { font-weight: 600; }
  .summary-table .total-row td { background: #fafafa; font-weight: 700; border-top: 2px solid #ddd; }
  .num { text-align: center; font-weight: 600; }

  .badge { display: inline-block; padding: 3px 12px; border-radius: 12px; font-size: 11px; font-weight: 600; color: #fff; }
  .badge-committed { background: #27ae60; }
  .badge-targeted { background: #f39c12; }
  .badge-cut { background: #e74c3c; }

  .detail-section { max-width: 1300px; margin: 0 auto 24px; }
  .kr-header { padding: 14px 20px; background: #fff; border-left: 5px solid; border-radius: 8px 8px 0 0; box-shadow: 0 1px 4px rgba(0,0,0,0.06); display: flex; justify-content: space-between; align-items: center; }
  .kr-header h3 { font-size: 15px; }
  .kr-header .counts { font-size: 12px; color: #888; }
  .detail-table { width: 100%; border-collapse: collapse; background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,0.06); border-radius: 0 0 8px 8px; overflow: hidden; }
  .detail-table th { background: #fafafa; text-align: left; padding: 8px 14px; font-size: 11px; color: #888; text-transform: uppercase; border-bottom: 1px solid #eee; }
  .detail-table td { padding: 8px 14px; font-size: 12px; border-bottom: 1px solid #f5f5f5; }
  .detail-table tr:last-child td { border-bottom: none; }
  .detail-table tr:hover td { background: #f9fbfd; }
  a { color: #4A90D9; text-decoration: none; }
  a:hover { text-decoration: underline; }
  .state-tag { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 10px; font-weight: 600; color: #fff; }
  .type-tag { color: #888; font-size: 11px; }
  .em-owner { font-size: 11px; color: #555; white-space: nowrap; }

  select.commitment-select {
    padding: 3px 8px;
    border-radius: 12px;
    border: 1px solid #ddd;
    font-size: 11px;
    font-weight: 600;
    cursor: pointer;
    background: #f9f9f9;
    appearance: auto;
  }
  select.commitment-select.val-Committed { background: #27ae60; color: #fff; border-color: #27ae60; }
  select.commitment-select.val-Targeted { background: #f39c12; color: #fff; border-color: #f39c12; }
  select.commitment-select.val-Cut { background: #e74c3c; color: #fff; border-color: #e74c3c; }
  select.commitment-select.val- { background: #f9f9f9; color: #999; }

  .save-bar { max-width: 1300px; margin: 0 auto 10px; display: flex; justify-content: flex-end; align-items: center; gap: 10px; }
  .save-status { font-size: 12px; font-weight: 600; padding: 3px 10px; border-radius: 8px; }
  .status-ok { color: #27ae60; }
  .status-saving { color: #f39c12; }
  .status-error { color: #e74c3c; cursor: pointer; }
  .token-btn { font-size: 11px; padding: 4px 12px; border: 1px solid #ddd; border-radius: 8px; background: #fff; cursor: pointer; color: #555; }
  .token-btn:hover { background: #f0f2f5; }
  .refresh-btn { font-size: 11px; padding: 4px 12px; border: 1px solid #ddd; border-radius: 8px; background: #fff; cursor: pointer; color: #555; }
  .refresh-btn:hover { background: #f0f2f5; }

  /* Cutline / Shiproom view */
  .cutline { max-width: 1300px; margin: 0 auto 36px; }
  .cutline-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 24px; }
  .cut-card { background: #fff; border-radius: 10px; padding: 16px 20px; box-shadow: 0 1px 6px rgba(0,0,0,0.08); text-align: center; }
  .cut-card .cut-num { font-size: 36px; font-weight: 800; line-height: 1.1; }
  .cut-card .cut-label { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 4px; }
  .cut-card .cut-sub { font-size: 11px; color: #aaa; margin-top: 2px; }
  .progress-bar-wrap { background: #eee; border-radius: 6px; height: 10px; overflow: hidden; margin-top: 8px; }
  .progress-bar-fill { height: 100%; border-radius: 6px; transition: width 0.3s; }

  .type-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 10px; margin-bottom: 24px; }
  .type-card { background: #fff; border-radius: 8px; padding: 12px 14px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); text-align: center; border-top: 3px solid; }
  .type-card .tc-num { font-size: 28px; font-weight: 700; }
  .type-card .tc-label { font-size: 11px; color: #666; margin-top: 2px; }

  .kr-matrix { border-collapse: collapse; width: 100%; background: #fff; border-radius: 10px; overflow: hidden; box-shadow: 0 1px 6px rgba(0,0,0,0.08); }
  .kr-matrix th { background: #1a1a1a; color: #fff; padding: 10px 14px; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; text-align: center; }
  .kr-matrix th:first-child { text-align: left; }
  .kr-matrix td { padding: 10px 14px; border-bottom: 1px solid #f0f0f0; font-size: 12px; text-align: center; }
  .kr-matrix td:first-child { text-align: left; font-weight: 600; }
  .kr-matrix tr:last-child td { border-bottom: none; }
  .kr-matrix tr:hover td { background: #f9fbfd; }
  .kr-matrix .bar-cell { position: relative; }
  .kr-matrix .mini-bar { display: inline-block; height: 6px; border-radius: 3px; vertical-align: middle; margin-left: 6px; }

  /* State filter bar */
  .filter-bar { max-width: 1300px; margin: 0 auto 16px; display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
  .filter-bar label { font-size: 12px; font-weight: 600; color: #555; margin-right: 4px; }
  .filter-pill { display: inline-block; padding: 5px 14px; border-radius: 16px; font-size: 11px; font-weight: 600; cursor: pointer; border: 2px solid; user-select: none; transition: all 0.15s; }
  .filter-pill.active { color: #fff; }
  .filter-pill:not(.active) { background: #fff; }
  .filter-pill:hover { opacity: 0.85; }
"""

    js = r"""
const REPO = 'rkundu-msft/ga-target-1e-report';
const FILE_PATH = 'commitments.json';
let commitments = {};
let fileSha = null;
let saveTimeout = null;

function getToken() { return localStorage.getItem('gh_token_ga1e'); }
function setToken(t) { localStorage.setItem('gh_token_ga1e', t); }
function promptToken() {
  var t = prompt('Enter a GitHub Personal Access Token (needs repo scope) to enable shared saving:');
  if (t) { setToken(t); return true; }
  return false;
}

async function loadCommitments() {
  var local = localStorage.getItem('ga1e_commitments');
  if (local) {
    try { commitments = JSON.parse(local); applyCommitments(); } catch(e) {}
  }
  try {
    var resp = await fetch('https://api.github.com/repos/' + REPO + '/contents/' + FILE_PATH, {
      headers: { 'Accept': 'application/vnd.github.v3+json', 'Cache-Control': 'no-cache' }
    });
    if (resp.ok) {
      var data = await resp.json();
      fileSha = data.sha;
      var remote = JSON.parse(atob(data.content));
      if (Object.keys(remote).length > 0) {
        commitments = Object.assign(commitments, remote);
        localStorage.setItem('ga1e_commitments', JSON.stringify(commitments));
        applyCommitments();
      }
    }
  } catch(e) { console.error('Remote load failed, using local:', e); }
}

function applyCommitments() {
  document.querySelectorAll('select.commitment-select').forEach(function(sel) {
    var id = sel.closest('tr').querySelector('a').textContent.trim();
    if (commitments[id]) {
      sel.value = commitments[id];
      sel.className = 'commitment-select val-' + commitments[id];
    }
  });
  updateSummary();
}

async function saveCommitments() {
  var token = getToken();
  if (!token) return;
  showStatus('Saving...', 'saving');
  try {
    var getResp = await fetch('https://api.github.com/repos/' + REPO + '/contents/' + FILE_PATH, {
      headers: { 'Authorization': 'token ' + token, 'Accept': 'application/vnd.github.v3+json' }
    });
    if (getResp.ok) { fileSha = (await getResp.json()).sha; }
    var content = btoa(unescape(encodeURIComponent(JSON.stringify(commitments, null, 2))));
    var resp = await fetch('https://api.github.com/repos/' + REPO + '/contents/' + FILE_PATH, {
      method: 'PUT',
      headers: { 'Authorization': 'token ' + token, 'Accept': 'application/vnd.github.v3+json', 'Content-Type': 'application/json' },
      body: JSON.stringify({ message: 'Update commitments', content: content, sha: fileSha })
    });
    if (resp.ok) {
      fileSha = (await resp.json()).content.sha;
      showStatus('Saved', 'ok');
    } else if (resp.status === 401 || resp.status === 403) {
      showStatus('Bad token - click Set Token', 'error');
      localStorage.removeItem('gh_token_ga1e');
    } else { showStatus('Save failed: ' + resp.status, 'error'); }
  } catch(e) { showStatus('Save error', 'error'); }
}

function queueSave() {
  if (saveTimeout) clearTimeout(saveTimeout);
  saveTimeout = setTimeout(saveCommitments, 1000);
}

function showStatus(msg, type) {
  var el = document.getElementById('save-status');
  el.textContent = msg;
  el.className = 'save-status status-' + type;
  if (type === 'ok') setTimeout(function() { el.textContent = ''; el.className = 'save-status'; }, 3000);
}

function updateDropdown(sel) {
  sel.className = 'commitment-select val-' + sel.value;
  var id = sel.closest('tr').querySelector('a').textContent.trim();
  if (sel.value) { commitments[id] = sel.value; } else { delete commitments[id]; }
  localStorage.setItem('ga1e_commitments', JSON.stringify(commitments));
  updateSummary();
  queueSave();
}

function updateSummary() {
  var rows = document.querySelectorAll('tr[data-kr]');
  var counts = {};
  rows.forEach(function(row) {
    var kr = row.getAttribute('data-kr');
    if (!counts[kr]) counts[kr] = {total:0, Committed:0, Targeted:0, Cut:0};
    counts[kr].total++;
    var badge = row.querySelector('.badge');
    var sel = row.querySelector('select.commitment-select');
    var val = '';
    if (badge) val = badge.textContent.trim();
    if (sel) val = sel.value;
    if (val && counts[kr][val] !== undefined) counts[kr][val]++;
  });
  var allKRs = document.querySelectorAll('tr[data-summary-kr]');
  var grandTotal=0, grandC=0, grandT=0, grandX=0;
  allKRs.forEach(function(row) {
    var kr = row.getAttribute('data-summary-kr');
    if (counts[kr]) {
      row.querySelector('.s-total').textContent = counts[kr].total;
      row.querySelector('.s-committed').textContent = counts[kr].Committed;
      row.querySelector('.s-targeted').textContent = counts[kr].Targeted;
      row.querySelector('.s-cut').textContent = counts[kr].Cut;
      grandTotal += counts[kr].total;
      grandC += counts[kr].Committed;
      grandT += counts[kr].Targeted;
      grandX += counts[kr].Cut;
    }
  });
  var totalRow = document.querySelector('tr.total-row');
  if (totalRow) {
    totalRow.querySelector('.s-total').textContent = grandTotal;
    totalRow.querySelector('.s-committed').textContent = grandC;
    totalRow.querySelector('.s-targeted').textContent = grandT;
    totalRow.querySelector('.s-cut').textContent = grandX;
  }
  document.querySelectorAll('.kr-header[data-kr]').forEach(function(hdr) {
    var kr = hdr.getAttribute('data-kr');
    if (counts[kr]) {
      hdr.querySelector('.counts').textContent =
        counts[kr].Committed + 'C / ' + counts[kr].Targeted + 'T / ' + counts[kr].Cut + 'X \u00b7 ' + counts[kr].total + ' total';
    }
  });
}

var filters = { commitment: new Set(), type: new Set() };

function toggleFilter(pill, value, color, group) {
  var s = filters[group];
  if (s.has(value)) {
    s.delete(value);
    pill.classList.remove('active');
    pill.style.background = '#fff';
    pill.style.color = color;
  } else {
    s.add(value);
    pill.classList.add('active');
    pill.style.background = color;
    pill.style.color = '#fff';
  }
  applyFilters();
}

function getRowCommitment(row) {
  var badge = row.querySelector('.badge');
  if (badge) return badge.textContent.trim();
  var sel = row.querySelector('select.commitment-select');
  if (sel && sel.value) return sel.value;
  return 'Unset';
}

function getRowType(row) {
  var el = row.querySelector('.type-tag');
  return el ? el.textContent.trim() : '';
}

function applyFilters() {
  var rows = document.querySelectorAll('tr[data-kr]');
  var hasC = filters.commitment.size > 0;
  var hasT = filters.type.size > 0;
  rows.forEach(function(row) {
    var matchC = !hasC || filters.commitment.has(getRowCommitment(row));
    var matchT = !hasT || filters.type.has(getRowType(row));
    row.style.display = (matchC && matchT) ? '' : 'none';
  });
  var anyFilter = hasC || hasT;
  document.querySelectorAll('.kr-header[data-kr]').forEach(function(hdr) {
    var kr = hdr.getAttribute('data-kr');
    var table = hdr.nextElementSibling;
    if (!table) return;
    var visible = table.querySelectorAll('tr[data-kr]:not([style*="display: none"])').length;
    var total = table.querySelectorAll('tr[data-kr]').length;
    var countsEl = hdr.querySelector('.counts');
    var base = countsEl.getAttribute('data-base');
    if (!base) { countsEl.setAttribute('data-base', countsEl.textContent); base = countsEl.textContent; }
    countsEl.textContent = anyFilter ? visible + '/' + total + ' shown' : base;
  });
}

document.addEventListener('DOMContentLoaded', function() {
  updateSummary();
  loadCommitments();
});
"""

    lines = []
    lines.append('<!DOCTYPE html>')
    lines.append('<html lang="en">')
    lines.append('<head>')
    lines.append('<meta charset="UTF-8">')
    lines.append('<title>GA_Target_1E - GA Target View</title>')
    lines.append(f'<style>{css}</style>')
    lines.append(f'<script>{js}</script>')
    lines.append('</head>')
    lines.append('<body>')
    lines.append('')
    lines.append('<div class="header">')
    lines.append('  <h1>One Events GA (1E) - Commitment View</h1>')
    lines.append(f'  <p>Refreshed: {today_str}</p>')
    lines.append('</div>')
    lines.append('<div class="save-bar">')
    lines.append('  <span id="save-status" class="save-status"></span>')
    lines.append('  <button class="token-btn" onclick="if(promptToken())loadCommitments()">Set GitHub Token</button>')
    lines.append('</div>')
    lines.append('')

    # ── FILTER BARS ──
    lines.append('<div class="filter-bar">')
    lines.append('  <label>Commitment:</label>')
    lines.append('  <span class="filter-pill" data-group="commitment" style="border-color:#27ae60;color:#27ae60" onclick="toggleFilter(this,\'Committed\',\'#27ae60\',\'commitment\')">Committed</span>')
    lines.append('  <span class="filter-pill" data-group="commitment" style="border-color:#f39c12;color:#f39c12" onclick="toggleFilter(this,\'Targeted\',\'#f39c12\',\'commitment\')">Targeted</span>')
    lines.append('  <span class="filter-pill" data-group="commitment" style="border-color:#e74c3c;color:#e74c3c" onclick="toggleFilter(this,\'Cut\',\'#e74c3c\',\'commitment\')">Cut</span>')
    lines.append('  <span class="filter-pill" data-group="commitment" style="border-color:#95a5a6;color:#95a5a6" onclick="toggleFilter(this,\'Unset\',\'#95a5a6\',\'commitment\')">Unset</span>')
    lines.append('  <span style="margin:0 8px;color:#ddd">|</span>')
    lines.append('  <label>Type:</label>')

    type_filter_colors = {
        "Feature": "#2980b9", "Bug": "#e74c3c", "User Feedback": "#f39c12",
        "Task": "#8e44ad", "Requirement": "#16a085",
    }
    all_types = set()
    for items_list in kr_groups.values():
        for item in items_list:
            all_types.add(item["type"])
    for t in ["Feature", "Bug", "User Feedback", "Task", "Requirement"]:
        if t in all_types:
            c = type_filter_colors.get(t, "#555")
            lines.append(f'  <span class="filter-pill" data-group="type" style="border-color:{c};color:{c}" onclick="toggleFilter(this,\'{html.escape(t)}\',\'{c}\',\'type\')">{html.escape(t)}</span>')
    for t in all_types:
        if t not in type_filter_colors:
            lines.append(f'  <span class="filter-pill" data-group="type" style="border-color:#555;color:#555" onclick="toggleFilter(this,\'{html.escape(t)}\',\'#555\',\'type\')">{html.escape(t)}</span>')

    lines.append('</div>')
    lines.append('')

    # ── CUTLINE / SHIPROOM VIEW ──
    total = stats["total"]
    done = stats["done"]
    active = stats["active"]
    proposed = stats["proposed"]
    pct = stats["pct_done"]
    tc = stats["type_counts"]
    sc = stats["state_counts"]

    type_colors = {
        "Feature": "#2980b9", "Bug": "#e74c3c", "User Feedback": "#f39c12",
        "Task": "#8e44ad", "Requirement": "#16a085", "Design Change Request": "#2c3e50",
    }

    lines.append('<div class="cutline">')
    lines.append('<div class="section-title">Shiproom Cutline</div>')

    # Top-level scorecards
    lines.append('<div class="cutline-grid">')
    lines.append(f'  <div class="cut-card"><div class="cut-num" style="color:#1a1a1a">{total}</div><div class="cut-label">Total Items</div><div class="cut-sub">{len([k for k in kr_groups if kr_groups[k]])} KRs active</div></div>')
    lines.append(f'  <div class="cut-card"><div class="cut-num" style="color:#27ae60">{done}</div><div class="cut-label">Done</div><div class="cut-sub">Closed + Resolved + RollingOut</div><div class="progress-bar-wrap"><div class="progress-bar-fill" style="width:{pct}%;background:#27ae60"></div></div></div>')
    lines.append(f'  <div class="cut-card"><div class="cut-num" style="color:#2980b9">{active}</div><div class="cut-label">In Flight</div><div class="cut-sub">Active work items</div></div>')
    lines.append(f'  <div class="cut-card"><div class="cut-num" style="color:#95a5a6">{proposed}</div><div class="cut-label">Not Started</div><div class="cut-sub">Proposed / backlog</div></div>')
    lines.append('</div>')

    # Work item type breakdown cards
    lines.append('<div class="type-grid">')
    type_order = ["Feature", "Bug", "User Feedback", "Task", "Requirement"]
    for t in type_order:
        c = tc.get(t, 0)
        if c > 0:
            color = type_colors.get(t, "#555")
            lines.append(f'  <div class="type-card" style="border-color:{color}"><div class="tc-num" style="color:{color}">{c}</div><div class="tc-label">{html.escape(t)}{"s" if c != 1 else ""}</div></div>')
    # Any other types not in the order
    for t, c in tc.items():
        if t not in type_order and c > 0:
            color = type_colors.get(t, "#555")
            lines.append(f'  <div class="type-card" style="border-color:{color}"><div class="tc-num" style="color:{color}">{c}</div><div class="tc-label">{html.escape(t)}{"s" if c != 1 else ""}</div></div>')
    lines.append('</div>')

    # KR execution matrix
    ordered_krs_for_matrix = [kr for kr in KR_ORDER if kr in kr_groups]
    for kr in kr_groups:
        if kr not in ordered_krs_for_matrix:
            ordered_krs_for_matrix.append(kr)

    lines.append('<table class="kr-matrix">')
    lines.append('  <thead><tr>')
    lines.append('    <th>KR</th><th>Total</th><th>Features</th><th>Bugs</th><th>Feedback</th><th>Tasks</th><th>Done</th><th>Active</th><th>Proposed</th><th>Completion</th>')
    lines.append('  </tr></thead>')
    lines.append('  <tbody>')

    grand = {"total": 0, "feat": 0, "bug": 0, "fb": 0, "task": 0, "done": 0, "active": 0, "proposed": 0}

    for kr in ordered_krs_for_matrix:
        ks = stats["kr_stats"][kr]
        color = KR_COLORS.get(kr, "#555")
        feat = ks["type_counts"].get("Feature", 0)
        bug = ks["type_counts"].get("Bug", 0)
        fb = ks["type_counts"].get("User Feedback", 0)
        task = ks["type_counts"].get("Task", 0) + ks["type_counts"].get("Requirement", 0)
        d = ks["done"]
        a = ks["state_counts"].get("Active", 0)
        p = ks["total"] - d - a
        pct_kr = ks["pct_done"]

        grand["total"] += ks["total"]
        grand["feat"] += feat
        grand["bug"] += bug
        grand["fb"] += fb
        grand["task"] += task
        grand["done"] += d
        grand["active"] += a
        grand["proposed"] += p

        bar_color = "#27ae60" if pct_kr >= 50 else "#f39c12" if pct_kr >= 25 else "#e74c3c"
        lines.append(f'  <tr>')
        lines.append(f'    <td style="color:{color}">{html.escape(kr)}</td>')
        lines.append(f'    <td><strong>{ks["total"]}</strong></td>')
        lines.append(f'    <td>{feat or "-"}</td>')
        lines.append(f'    <td>{bug or "-"}</td>')
        lines.append(f'    <td>{fb or "-"}</td>')
        lines.append(f'    <td>{task or "-"}</td>')
        lines.append(f'    <td style="color:#27ae60;font-weight:600">{d}</td>')
        lines.append(f'    <td style="color:#2980b9;font-weight:600">{a}</td>')
        lines.append(f'    <td style="color:#95a5a6">{p}</td>')
        lines.append(f'    <td class="bar-cell">{pct_kr}%<span class="mini-bar" style="width:{max(pct_kr, 4)}px;background:{bar_color}"></span></td>')
        lines.append(f'  </tr>')

    # Total row
    gpct = round(grand["done"] * 100 / grand["total"]) if grand["total"] else 0
    gbar = "#27ae60" if gpct >= 50 else "#f39c12" if gpct >= 25 else "#e74c3c"
    lines.append(f'  <tr style="background:#fafafa;font-weight:700;border-top:2px solid #ddd">')
    lines.append(f'    <td>TOTAL</td>')
    lines.append(f'    <td>{grand["total"]}</td>')
    lines.append(f'    <td>{grand["feat"]}</td>')
    lines.append(f'    <td>{grand["bug"]}</td>')
    lines.append(f'    <td>{grand["fb"]}</td>')
    lines.append(f'    <td>{grand["task"]}</td>')
    lines.append(f'    <td style="color:#27ae60">{grand["done"]}</td>')
    lines.append(f'    <td style="color:#2980b9">{grand["active"]}</td>')
    lines.append(f'    <td style="color:#95a5a6">{grand["proposed"]}</td>')
    lines.append(f'    <td class="bar-cell">{gpct}%<span class="mini-bar" style="width:{max(gpct, 4)}px;background:{gbar}"></span></td>')
    lines.append(f'  </tr>')

    lines.append('  </tbody>')
    lines.append('</table>')
    lines.append('</div>')
    lines.append('')

    # ── COMMITMENT SUMMARY (existing) ──
    lines.append('<div class="section-title">Commitment Summary by KR</div>')
    lines.append('<table class="summary-table">')
    lines.append('  <thead>')
    lines.append('    <tr>')
    lines.append('      <th>KR</th>')
    lines.append('      <th style="text-align:center">Total</th>')
    lines.append('      <th style="text-align:center"><span class="badge badge-committed">Committed</span></th>')
    lines.append('      <th style="text-align:center"><span class="badge badge-targeted">Targeted</span></th>')
    lines.append('      <th style="text-align:center"><span class="badge badge-cut">Cut</span></th>')
    lines.append('    </tr>')
    lines.append('  </thead>')
    lines.append('  <tbody>')

    # Build ordered KR list
    ordered_krs = [kr for kr in KR_ORDER if kr in kr_groups]
    for kr in kr_groups:
        if kr not in ordered_krs:
            ordered_krs.append(kr)

    grand_total = 0
    grand_c = grand_t = grand_x = 0
    for kr in ordered_krs:
        items = kr_groups[kr]
        total = len(items)
        grand_total += total
        c = sum(1 for i in items if i.get("commitment") == "Committed")
        t = sum(1 for i in items if i.get("commitment") == "Targeted")
        x = sum(1 for i in items if i.get("commitment") == "Cut")
        grand_c += c; grand_t += t; grand_x += x
        color = KR_COLORS.get(kr, "#555")
        lines.append(f'    <tr data-summary-kr="{html.escape(kr)}">')
        lines.append(f'      <td class="kr-name" style="color:{color}">{html.escape(kr)}</td>')
        lines.append(f'      <td class="num s-total">{total}</td>')
        lines.append(f'      <td class="num s-committed" style="color:#27ae60">{c}</td>')
        lines.append(f'      <td class="num s-targeted" style="color:#f39c12">{t}</td>')
        lines.append(f'      <td class="num s-cut" style="color:#e74c3c">{x}</td>')
        lines.append('    </tr>')

    lines.append('    <tr class="total-row">')
    lines.append('      <td>TOTAL</td>')
    lines.append(f'      <td class="num s-total">{grand_total}</td>')
    lines.append(f'      <td class="num s-committed" style="color:#27ae60">{grand_c}</td>')
    lines.append(f'      <td class="num s-targeted" style="color:#f39c12">{grand_t}</td>')
    lines.append(f'      <td class="num s-cut" style="color:#e74c3c">{grand_x}</td>')
    lines.append('    </tr>')
    lines.append('  </tbody>')
    lines.append('</table>')
    lines.append('')

    # Detail sections
    lines.append('<div class="section-title">Feature Detail by KR</div>')

    for kr in ordered_krs:
        items = kr_groups[kr]
        color = KR_COLORS.get(kr, "#555")
        total = len(items)

        lines.append('<div class="detail-section">')
        lines.append(f'  <div class="kr-header" style="border-color:{color}" data-kr="{html.escape(kr)}">')
        lines.append(f'    <h3 style="color:{color}">{html.escape(kr)}</h3>')
        kc = sum(1 for i in items if i.get("commitment") == "Committed")
        kt = sum(1 for i in items if i.get("commitment") == "Targeted")
        kx = sum(1 for i in items if i.get("commitment") == "Cut")
        lines.append(f'    <div class="counts">{kc}C / {kt}T / {kx}X &middot; {total} total</div>')
        lines.append('  </div>')
        lines.append('  <table class="detail-table">')
        lines.append('    <thead><tr><th>ID</th><th>Feature</th><th>Type</th><th>EM Owner</th><th>State</th><th>Commitment</th></tr></thead>')
        lines.append('    <tbody>')

        for item in items:
            wid = item["id"]
            title = html.escape(item["title"])
            wtype = html.escape(item["type"])
            state = item["state"]
            owner = html.escape(item["owner"])
            state_color = STATE_COLORS.get(state, "#95a5a6")
            url = f"https://domoreexp.visualstudio.com/MSTeams/_workitems/edit/{wid}"

            lines.append(f'      <tr data-kr="{html.escape(kr)}">')
            lines.append(f'        <td><a href="{url}" target="_blank">{wid}</a></td>')
            lines.append(f'        <td>{title}</td>')
            lines.append(f'        <td><span class="type-tag">{wtype}</span></td>')
            lines.append(f'        <td><span class="em-owner">{owner}</span></td>')
            lines.append(f'        <td><span class="state-tag" style="background:{state_color}">{html.escape(state)}</span></td>')
            commitment = item.get("commitment", "")
            val_class = f"val-{commitment}" if commitment else "val-"
            lines.append('        <td>')
            lines.append(f'          <select class="commitment-select {val_class}" onchange="updateDropdown(this)">')
            lines.append(f'            <option value=""{" selected" if not commitment else ""}>-- Select --</option>')
            lines.append(f'            <option value="Committed"{" selected" if commitment == "Committed" else ""}>Committed</option>')
            lines.append(f'            <option value="Targeted"{" selected" if commitment == "Targeted" else ""}>Targeted</option>')
            lines.append(f'            <option value="Cut"{" selected" if commitment == "Cut" else ""}>Cut</option>')
            lines.append('          </select>')
            lines.append('        </td>')
            lines.append('      </tr>')

        lines.append('    </tbody>')
        lines.append('  </table>')
        lines.append('</div>')

    lines.append('')
    lines.append('</body>')
    lines.append('</html>')

    return '\n'.join(lines)


def main():
    print("Getting bearer token...")
    token = get_bearer_token()

    # Sync browser edits from GitHub -> ADO before refreshing
    sync_github_to_ado(token)

    print("Querying ADO for GA_Target_1E tagged items...")

    # Step 1: WIQL query
    ids = query_work_items(token)
    print(f"  Found {len(ids)} work items")

    if not ids:
        print("No work items found. Check auth and query.")
        return

    # Step 2: Fetch details
    print("  Fetching work item details...")
    items = get_work_items_batch(token, ids)
    print(f"  Got {len(items)} item details")

    # Step 3: Classify into KRs by title/tags keywords
    kr_groups = defaultdict(list)
    for item in items:
        f = item.get("fields", {})
        assigned = f.get("System.AssignedTo", {})
        owner = assigned.get("displayName", "") if isinstance(assigned, dict) else str(assigned) if assigned else ""
        title = f.get("System.Title", "")
        tags = f.get("System.Tags", "")

        kr = classify_kr(item["id"], title, tags)

        commitment = f.get("Custom.CommittedTargettedCut", "") or ""
        if commitment not in ("Committed", "Targeted", "Cut"):
            commitment = ""

        kr_groups[kr].append({
            "id": item["id"],
            "title": title,
            "type": f.get("System.WorkItemType", ""),
            "state": f.get("System.State", ""),
            "owner": owner,
            "tags": tags,
            "commitment": commitment,
        })

    # Sort items within each KR by ID
    for kr in kr_groups:
        kr_groups[kr].sort(key=lambda x: x["id"])

    print(f"\n  KR breakdown:")
    for kr in KR_ORDER:
        if kr in kr_groups:
            print(f"    {kr}: {len(kr_groups[kr])} items")
    for kr in kr_groups:
        if kr not in KR_ORDER:
            print(f"    {kr}: {len(kr_groups[kr])} items  [NEW KR]")

    # Step 5: Compute stats and generate HTML
    today_str = datetime.now().strftime("%B %d, %Y %I:%M %p")
    stats = compute_stats(dict(kr_groups))
    print(f"\n  Cutline: {stats['done']} done / {stats['active']} active / {stats['proposed']} proposed ({stats['pct_done']}% complete)")
    for t, c in stats["type_counts"].items():
        print(f"    {t}: {c}")
    html_content = generate_html(dict(kr_groups), today_str, stats)

    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "GA_Target_View.html")
    with open(output_path, "w", encoding="utf-8") as fout:
        fout.write(html_content)

    print(f"\n  Report saved to: {output_path}")
    print(f"  Total items: {sum(len(v) for v in kr_groups.values())}")


if __name__ == "__main__":
    main()
