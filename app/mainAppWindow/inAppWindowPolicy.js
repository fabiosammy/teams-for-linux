/**
 * In-app window policy for non-Teams web apps loaded through `app.url`
 * (e.g. Outlook on the web).
 *
 * The root window's `onNewWindow` was written for Teams: every `about:blank`
 * popup is denied (it feeds the hidden-window silent-auth workaround in
 * onBeforeRequestHandler) and every other URL goes to the external browser.
 * Outlook relies on both shapes for its pop-outs: "pop out" compose opens
 * `about:blank` and renders into it through `window.opener`, and "open in new
 * window" opens a URL on the Outlook host. Denying the first and sending the
 * second to the browser breaks pop-outs entirely (the browser has no session).
 *
 * `urlHandling.inAppHosts` lists exact hostnames whose popups stay in the app:
 *
 *   - a popup TO one of those hosts, and
 *   - an `about:blank` popup opened BY a page on one of those hosts
 *
 * are allowed as child windows. A child opened by `window.open` shares its
 * opener's session, so the pop-out is signed in and `window.opener` scripting
 * works. The list is empty by default, so Teams behaviour is unchanged.
 *
 * Matching is exact, like auth.webauthn.extraOrigins: a security decision
 * never depends on pattern matching, so each host is listed in its own right.
 *
 * Pure (no Electron imports) so the decision is unit-testable.
 */

/**
 * Reduce a configured entry to a lowercase hostname, or null if it cannot be
 * one. Bare hostnames only: no scheme, port, path or wildcard.
 * @param {unknown} entry
 * @returns {string|null}
 */
function normalizeHost(entry) {
  if (typeof entry !== "string") return null;
  const host = entry.trim().toLowerCase();
  if (!/^[a-z0-9.-]+$/.test(host)) return null;
  if (host.startsWith(".") || host.endsWith(".") || host.includes("..")) return null;
  return host;
}

/**
 * Build the set of in-app hosts. Malformed entries are dropped with a warning
 * that reports how many were dropped and never the values (a tenant host can
 * identify the organisation).
 * @param {unknown} inAppHosts - urlHandling.inAppHosts
 * @returns {Set<string>}
 */
function buildInAppHosts(inAppHosts) {
  const hosts = new Set();
  if (!Array.isArray(inAppHosts)) return hosts;

  let rejected = 0;
  for (const entry of inAppHosts) {
    const host = normalizeHost(entry);
    if (host) hosts.add(host);
    else rejected += 1;
  }
  if (rejected > 0) {
    console.warn("[IN_APP_WINDOWS] Ignored malformed urlHandling.inAppHosts entries", {
      rejected,
      expected: "bare hostname, e.g. outlook.office.com",
    });
  }
  return hosts;
}

function hostOf(url) {
  try {
    const parsed = new URL(url);
    return parsed.protocol === "https:" ? parsed.hostname : null;
  } catch {
    return null;
  }
}

function isAboutBlank(url) {
  return url === "about:blank" || url.startsWith("about:blank#");
}

/**
 * Decide whether a `window.open` request stays in the app.
 * @param {Set<string>} hosts - from buildInAppHosts
 * @param {string} url - the requested popup URL
 * @param {string} openerUrl - URL of the page calling window.open
 * @returns {boolean}
 */
function isInAppPopup(hosts, url, openerUrl) {
  if (hosts.size === 0 || typeof url !== "string") return false;
  if (url === "about:blank#blocked") return false;
  if (isAboutBlank(url)) return hosts.has(hostOf(openerUrl));
  return hosts.has(hostOf(url));
}

module.exports = { buildInAppHosts, isInAppPopup };
