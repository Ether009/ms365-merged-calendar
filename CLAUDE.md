# CLAUDE.md — MS365 Merged Calendar

Context for continuing development of this WordPress plugin in Claude Code.

## What this is

A single-file WordPress plugin that merges calendars from Microsoft 365 **groups**
and **shared mailboxes** into one filterable, windowed event list, rendered on the
front end via a shortcode. Events load asynchronously through a REST endpoint; the
browser never talks to Microsoft directly.

- Main file: `ms365-merged-calendar.php` (everything in one file, by design).
- Function prefix: `ms365cal_`.
- Shortcode: `[ms365_calendar calendars="slug1,slug2" days="14"]` (no attrs = all calendars, 14-day window).

## Golden rules (do not regress)

- Must pass `php -l` **and** `phpcs --standard=WordPress-Extra` with **zero violations**
  before every commit/handoff.
- Keep syntax **PHP 7.4 compatible** (header says `Requires PHP: 7.4`). No `match`,
  enums, nullsafe `?->`, constructor promotion, `readonly`, etc. It's verified to parse
  on 8.3 but written to the 7.4 baseline.
- All output escaped (`esc_html` / `esc_attr` / `esc_url`), all input sanitized, admin
  POST guarded by the `ms365cal_save` nonce. WPCS enforces this.
- Every function is prefixed `ms365cal_` to avoid collisions.
- The browser only ever sends/receives calendar **slugs**, never Graph identifiers.
  The server maps slugs → sources and intersects against configured calendars. Never
  expose raw group GUIDs / mailbox addresses in unauthenticated output. (The debug
  output does include sources and is therefore gated to `manage_options`.)
- The public REST endpoint is intentionally unauthenticated (it backs a public page).
  It's protected by slug-whitelisting + per-IP rate limiting + server-side caching —
  keep those in place.

## Architecture

**Async model**
1. `ms365cal_shortcode()` outputs only a shell (calendar chips, prev/next nav, loading
   banner) plus a JSON `data-config`. No Graph call happens on page render.
2. Front-end JS (in `ms365cal_assets()`, hooked to `wp_footer`) fetches
   `/wp-json/ms365cal/v1/events` per window and renders the list.
3. `ms365cal_rest_events()` → `ms365cal_events_window()` → `ms365cal_fetch_one()`
   per calendar → merge + sort.

**Auth** — app-only client-credentials (`ms365cal_get_token()`); token cached in the
`ms365cal_token` transient (~55 min). Credentials come from wp-config constants
(`MS365CAL_TENANT_ID` / `MS365CAL_CLIENT_ID` / `MS365CAL_CLIENT_SECRET`) if defined,
otherwise the DB settings.

**Graph calls** — `ms365cal_view_url()` builds `/groups/{id}/calendarView` or
`/users/{addr}/calendarView` (the *only* group-vs-mailbox difference). `$select` =
subject,start,end,location,isAllDay,onlineMeeting,webLink,type,seriesMasterId. The
`Prefer: outlook.timezone` header returns times already in the site timezone.

**Pagination** — `ms365cal_fetch_one()` follows `@odata.nextLink`, capped at
`MS365CAL_MAX_PAGES` (20) × 100 events/page.

**Recurrence** — calendarView returns expanded occurrences *without* the pattern.
`ms365cal_fetch_one()` collects the distinct `seriesMasterId`s while paging, then
`ms365cal_fetch_recurrence_map()` resolves them in one pass via Graph's **`$batch`**
endpoint (20 masters per request, relative URLs from `ms365cal_events_rel_base()`),
bounded per request by `MS365CAL_MAX_MASTERS` (200) shared across calendars via a
static budget. `ms365cal_format_recurrence()` renders a readable string like
"Repeats weekly on Mon, Wed until 1 Dec 2026"; anything unresolved (over budget, HTTP
failure, or no pattern) falls back to a generic "Recurring event".

**Caching / resilience** (`ms365cal_events_window()`)
- Per-window transient keyed by `slugs|start|days|tz`, payload `{fetched, events}`.
- Freshness window = `cache_minutes`; entry kept up to `DAY_IN_SECONDS` so stale data
  can still be served during an outage.
- Global Graph back-off transient `ms365cal_backoff` respects Graph 429/503
  `Retry-After` (capped at `MS365CAL_BACKOFF_MAX` = 300s). While active, Graph is not
  called; stale data is served.
- On throttle: serve stale if available, else `WP_Error('ms365cal_throttled')` →
  REST returns **429 + Retry-After**.

**Rate limiting** — `ms365cal_rate_check()` per-IP fixed window via
`ms365cal_rate_limits()`. Settings `rate_max` / `rate_window` (defaults from the
`MS365CAL_RATE_MAX` / `MS365CAL_RATE_WINDOW` constants); filters `ms365cal_rate_max` /
`ms365cal_rate_window`. IP via `ms365cal_client_ip()` (REMOTE_ADDR only, spoof-safe;
filter `ms365cal_client_ip` if behind a trusted proxy).

**Front-end behavior**
- Prev/next window paging (14-day default). Client-side in-memory window cache. Paging
  buttons disabled 600ms per click. A request-id guard drops superseded responses.
  Bounded auto-retry on 429 using `Retry-After` (max 3).
- Events grouped by day. Multi-day / ongoing events are clamped to the window start for
  grouping but display their real span (with end date/time).
- Event titles are expand/collapse **buttons** (accordion — at most one open). Detail
  panel shows: when, recurrence, location, online-meeting join link, and an optional
  "Open in Outlook" link.

## Settings & constants

**Option `ms365cal_settings`**: `tenant_id`, `client_id`, `client_secret`,
`cache_minutes` (20), `timezone`, `rate_max`, `rate_window`, `show_outlook`
(bool, default **false**), `deploy_key` (self-update secret; blank = endpoint off),
`calendars[]` = `{slug, label, color, type(group|mailbox), source, default}`.

**Constants**: `MS365CAL_OPTION`, `MS365CAL_TOKEN_TRANSIENT`, `MS365CAL_MAX_WINDOW` (62),
`MS365CAL_RATE_MAX` (30), `MS365CAL_RATE_WINDOW` (60), `MS365CAL_BACKOFF_MAX` (300),
`MS365CAL_MAX_PAGES` (20), `MS365CAL_MAX_MASTERS` (200).

**wp-config-only (optional)**: `MS365CAL_TENANT_ID` / `MS365CAL_CLIENT_ID` /
`MS365CAL_CLIENT_SECRET` (win over DB settings), and `MS365CAL_DEPLOY_KEY` — the shared
secret that enables the self-update endpoint (see below). All are undefined by default.

## Debug endpoint

`GET /wp-json/ms365cal/v1/events?debug=1` — **admin only** (`manage_options`).

REST cookie-auth requires a nonce, so a bare URL in the browser is treated as
unauthenticated and returns the normal public response. Use the "Run calendar
diagnostic" link on the settings page (it carries `wp_create_nonce('wp_rest')`), or
call with an `X-WP-Nonce` header.

Every debug run FLUSHES the token + back-off + window caches (`ms365cal_flush_cache()`)
then fetches live. Output: `window`, `token`, `token_roles`, `token_audience`
(via `ms365cal_decode_jwt_claims()`), and per-calendar
`{slug, type, source, http, events, throttled, error}`, plus `flushed: true`.

**Reading it**
- `token_roles: []` → the app-only token carries no application roles → permissions are
  Delegated not Application, OR admin consent not granted, OR the plugin's Client ID
  points at a different app registration than the one you consented.
- calendar `http: 403` "Access is denied..." → Exchange denial (usually the empty-roles
  case above, or an Exchange application access policy excluding the mailbox).
- `http: 404` → wrong source (group *email* instead of the object-ID GUID) or a
  Type/Source mismatch.
- `http: 200, events: 0` → credentials and IDs are fine; widen the window.

## Azure / Graph setup

App registration → **Application** permissions `Calendars.Read` (shared mailboxes) +
`Group.Read.All` (group calendars) → **Grant admin consent for <tenant>**. Add a client
secret (prefer `MS365CAL_CLIENT_SECRET` in wp-config over the DB). No Exchange
application access policy is needed if the app should read all mailboxes; add one
(`New-ApplicationAccessPolicy ... -AccessRight RestrictAccess`) to scope it down.

## Dev / verification tooling

Requires `php-cli` + `php-xml`. WPCS 3.x needs PHP_CodeSniffer ≥ 3.9, so use the
maintained PHPCSStandards fork (4.x phar), **not** the old squizlabs 3.7 phar.

```bash
curl -sSL -o phpcs.phar  https://github.com/PHPCSStandards/PHP_CodeSniffer/releases/latest/download/phpcs.phar
curl -sSL -o phpcbf.phar https://github.com/PHPCSStandards/PHP_CodeSniffer/releases/latest/download/phpcbf.phar
git clone --depth 1 https://github.com/WordPress/WordPress-Coding-Standards.git wpcs
git clone --depth 1 https://github.com/PHPCSStandards/PHPCSUtils.git phpcsutils
git clone --depth 1 https://github.com/PHPCSStandards/PHPCSExtra.git phpcsextra
php phpcs.phar --config-set installed_paths /abs/wpcs,/abs/phpcsutils,/abs/phpcsextra
```

Verify (all three must be clean):

```bash
php -l ms365-merged-calendar.php
php phpcbf.phar --standard=WordPress-Extra --extensions=php ms365-merged-calendar.php   # auto-fix
php phpcs.phar  --standard=WordPress-Extra --extensions=php ms365-merged-calendar.php   # expect zero violations
```

Two intentional suppressions exist and are justified: a
`WordPress.DB.DirectDatabaseQuery` ignore on the cache-flush prefix delete, and a
`...obfuscation_base64_decode` ignore on the JWT-claims decode.

## Packaging / install

WordPress "Upload Plugin" needs a zip whose top folder matches the plugin slug:

```bash
mkdir -p ms365-merged-calendar && cp ms365-merged-calendar.php ms365-merged-calendar/
zip -r ms365-merged-calendar.zip ms365-merged-calendar
```

A single `.php` dropped into `wp-content/plugins/` also works (self-update inactive
in that case — the updater lives in the bundled `plugin-update-checker/` folder). No
DB schema, no activation hooks — deactivation fully reverts.

## Releases & self-update

Repo: **[Ether009/ms365-merged-calendar](https://github.com/Ether009/ms365-merged-calendar)**
(public). The plugin self-updates in wp-admin from GitHub **Releases** via a vendored
[plugin-update-checker](https://github.com/YahnisElsts/plugin-update-checker) v5
(`ms365cal_init_updates()`). `enableReleaseAssets()` prefers the release's attached zip
and falls back to GitHub's source archive.

Two GitHub Actions workflows:
- `.github/workflows/lint.yml` — runs `php -l` + `phpcs --standard=WordPress-Extra` on
  every push/PR (the golden-rule checks). This repo has **no local PHP**, so CI is the
  verification path for the main file.
- `.github/workflows/release.yml` — on a `v*` tag push, builds the installable zip
  (top folder = slug, dev files excluded) and publishes a Release with it attached.

**Cutting a release:** bump the `Version:` header in the plugin file (this is what
drives whether an update is offered — it must match / exceed the tag), then
`git tag vX.Y.Z && git push --tags`. The vendored library is not linted (WPCS runs on
`ms365-merged-calendar.php` only).

**Self-update endpoint** — `POST /wp-json/ms365cal/v1/self-update`
(`ms365cal_rest_self_update()`). Forces a fresh PUC check and installs this repo's
latest release via `Plugin_Upgrader`, so a live site with no shell access can be
updated on demand. **Off unless a deploy key is set** — either `MS365CAL_DEPLOY_KEY` in
wp-config (preferred) or the **Deploy key** field on the settings page (stored in
`ms365cal_settings['deploy_key']`); the constant wins via `ms365cal_cred('deploy_key')`.
The key is sent in the `X-MS365CAL-Deploy-Key` header and compared with `hash_equals()`; a
`ms365cal_selfupdate_lock` transient blocks rapid/concurrent triggers. The source is
fixed to this repo, so it can only ever install the repo's own latest release. The
`release.yml` "Trigger live self-update" step calls it after publishing, gated on a repo
**secret** `SITE_DEPLOY_KEY` (= the wp-config key) and a repo **variable** `DEPLOY_URL`
(site base URL); it's best-effort and never fails the release. Note the endpoint only
exists on a site once it's running 2.0.4+.

## Known items / possible next work

- `MS365CAL_MAX_WINDOW` (62) is a code constant, not a UI field (deliberate safety bound).
- Default shortcode window is 14 days (tested at 60). Consider per-shortcode defaults.
- The 10 categorical calendar colors aren't fully colorblind-safe; the label pills
  mitigate this.
- Recurrence master-fetching adds Graph calls on a cold fetch; amortized by caching.

## Change log (recent, newest last)

1. Async plugin: shell + REST + per-window cache, prev/next paging.
2. Shared-mailbox support (calendar `type` = group | mailbox).
3. Admin-configurable rate limit + prev/next button cooldown.
4. Graph throttling: global back-off + stale-while-revalidate + client auto-retry.
5. Rate-limit values surfaced in the admin UI.
6. Admin-only debug mode (per-calendar live Graph status).
7. Nonce-authenticated diagnostic link (a bare REST URL isn't authenticated).
8. `token_roles` / `token_audience` in debug (diagnosed the empty-roles cause).
9. Debug now flushes token + caches — fixed the live issue (stale pre-consent token).
10. Pagination via `@odata.nextLink` (busy calendars were capped at 100).
11. End date/time labels for multi-day & timed events, ongoing-event grouping,
    recurrence shown in the list, and click-to-expand accordion (replaced the
    title→Outlook link).
12. Outlook link made an opt-in setting (default off).
13. Published to GitHub (Ether009/ms365-merged-calendar, public). Self-update wired via
    bundled plugin-update-checker v5 + a tag-driven release workflow; lint CI enforces
    `php -l` + WordPress-Extra (no local PHP on this box).
14. `webLink` now gated server-side: when `show_outlook` is off, the REST response
    strips `link` so the Outlook URL never leaves the server (was render-only before).
15. Fixed timezone-basis mismatch in date/time labels: `sort`/`dayKey` format via the
    DateTime's own (plugin) zone, but the `wp_date()` labels rendered in the WordPress
    site zone. When the two settings differ, all-day events (pinned to midnight) showed
    the wrong day and disagreed with their group. `wp_date()` now takes the plugin zone
    explicitly. (Minor sibling not yet addressed: the recurrence "until" date in
    `ms365cal_format_recurrence()`.)
16. Recurrence master resolution now uses Graph `$batch` (20/request, budget 200)
    instead of up-to-40 individual GETs, so large multi-calendar views show real
    patterns for far more events instead of the generic "Recurring event" fallback.
    Replaced `ms365cal_recurrence_text()`/`ms365cal_events_base()` with
    `ms365cal_fetch_recurrence_map()`/`ms365cal_events_rel_base()`.
17. Secret-guarded self-update endpoint (`POST /self-update`) + release-workflow
    trigger, so a live site with no shell access can be updated on demand / on release.
    Off unless `MS365CAL_DEPLOY_KEY` is set; installs only this repo's latest release.
18. Deploy key is now also settable on the settings page (**Deploy key** field, password
    input with keep-on-blank + a "Clear" toggle), stored in `ms365cal_settings`. The
    wp-config constant still wins (endpoint reads `ms365cal_cred('deploy_key')`).
