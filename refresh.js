// Cloudflare Pages Function — POST /api/refresh
// Triggers the GitHub Actions email-fetch pipeline (workflow_dispatch) so the
// dashboard's Refresh button pulls brand-new reports, not just the last build.
//
// Requires a Cloudflare Pages environment variable GH_DISPATCH_TOKEN
// (a GitHub token with Actions: read & write on this repo). It is stored
// server-side in Cloudflare and never reaches the browser.
//
// Optional overrides: GH_REPO, GH_WORKFLOW, GH_REF.

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
  });
}

export async function onRequestPost({ env }) {
  const token = env.GH_DISPATCH_TOKEN;
  const repo = env.GH_REPO || "bhaudiu/courtyardvadodara-automation";
  const workflow = env.GH_WORKFLOW || "deploy.yml";
  const ref = env.GH_REF || "main";

  if (!token) return json({ ok: false, error: "not_configured" }, 501);

  const gh = {
    Authorization: `Bearer ${token}`,
    Accept: "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
    "User-Agent": "cyv-refresh",
  };

  // Debounce: if a run is already queued or running, don't stack another —
  // just tell the client to keep polling.
  try {
    const runsR = await fetch(
      `https://api.github.com/repos/${repo}/actions/workflows/${workflow}/runs?per_page=1`,
      { headers: gh }
    );
    if (runsR.ok) {
      const runs = await runsR.json();
      const last = runs.workflow_runs && runs.workflow_runs[0];
      if (last && (last.status === "queued" || last.status === "in_progress")) {
        return json({ ok: true, already: true });
      }
    }
  } catch (e) {
    // ignore — fall through to dispatch
  }

  const disp = await fetch(
    `https://api.github.com/repos/${repo}/actions/workflows/${workflow}/dispatches`,
    {
      method: "POST",
      headers: { ...gh, "Content-Type": "application/json" },
      body: JSON.stringify({ ref }),
    }
  );

  if (disp.status === 204) return json({ ok: true, triggered: true });
  const detail = (await disp.text()).slice(0, 200);
  return json({ ok: false, error: "dispatch_failed", status: disp.status, detail }, 502);
}

export function onRequestGet() {
  return json({ ok: true, hint: "POST here to fetch the latest reports." });
}
