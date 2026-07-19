# MS365 Merged Calendar

A single-file WordPress plugin that merges calendars from Microsoft 365 **groups**
and **shared mailboxes** into one filterable, windowed event list, rendered on the
front end via a shortcode. Events load asynchronously through a REST endpoint; the
browser never talks to Microsoft directly and only ever sees calendar *slugs*, never
Graph identifiers.

## Install

Download `ms365-merged-calendar.zip` from the [latest release](../../releases/latest),
then in wp-admin go to **Plugins → Add New → Upload Plugin**, upload the zip, and
activate. (A single `ms365-merged-calendar.php` dropped into `wp-content/plugins/`
also works, but then the self-updater is inactive — the update checker lives in the
bundled `plugin-update-checker/` folder.)

No database schema and no activation hooks: deactivation fully reverts.

## Configure

1. **Azure / Entra ID** — register an app, grant **Application** permissions
   `Calendars.Read` (shared mailboxes) and `Group.Read.All` (group calendars), grant
   admin consent for the tenant, and add a client secret.
2. **WordPress** — under **Settings → MS365 Calendar**, enter the Tenant ID, Client
   ID, and secret (or, preferred, define `MS365CAL_TENANT_ID` / `MS365CAL_CLIENT_ID` /
   `MS365CAL_CLIENT_SECRET` in `wp-config.php`), then add calendars: a **group** uses
   its object-ID GUID as the source, a **shared mailbox** uses its email address.
3. Use the **Run calendar diagnostic** link on the settings page to verify live Graph
   status per calendar if anything looks empty.

## Usage

```
[ms365_calendar]                                        all calendars, current week (Monday–Sunday)
[ms365_calendar calendars="eng,events"]                 only these slugs are shown at all
[ms365_calendar enabled="eng"]                          all shown, but only "eng" starts checked
[ms365_calendar calendars="eng,events" enabled="eng"]   both together
```

Every calendar is checked (visible) by default; `enabled` scopes that down for a
particular embed — the rest still appear as chips, just unchecked until clicked.

The view is a fixed weekly window; visitors page forward/back one week at a time, and
the current week is the earliest they can go back to.

## Automatic updates

Self-update is **off by default** — no repo is assumed. To enable it, set **Settings →
MS365 Calendar → Update source** to a GitHub repo, as `owner/repo` or a full URL (or
define `MS365CAL_UPDATE_REPO` in `wp-config.php`, which takes precedence over the
field). Once set, the site checks that repo's GitHub Releases for newer versions and
offers a one-click update in **wp-admin → Plugins**, the same way any other plugin
update would appear.

To track this plugin's own official releases, set it to `Ether009/ms365-merged-calendar`
(this repo). To track a fork instead, point it at that fork's `owner/repo`.

For updates to actually be offered from whichever repo you configure:

- The repo must be **public** (the update checker doesn't authenticate).
- Each release needs a **`vX.Y.Z` tag** whose version is equal to or ahead of the
  `Version:` header inside that release's `ms365-merged-calendar.php` — that header
  is what WordPress compares against the installed version, not the tag itself.
- Attaching a **release asset zip** (top-level folder named `ms365-merged-calendar`,
  same layout as this repo's own releases) is recommended — the update checker
  prefers it and it lets you exclude dev-only files from what gets installed.
  Without one, it falls back to GitHub's auto-generated source archive, which works
  but ships the whole repo as-is.

### Instant updates (optional)

WordPress only checks for plugin updates on its own background schedule (roughly every
12 hours), so a release you've just published can sit unnoticed for a while even with
Update source configured. For a site with no shell/SFTP access, the plugin exposes a
secret-guarded endpoint that forces an immediate check-and-install instead of waiting:

```
POST /wp-json/ms365cal/v1/self-update
Header: X-MS365CAL-Deploy-Key: <your key>
```

It re-checks the configured repo right away and, if a newer release exists, installs
it on the spot — response is JSON with `from`/`to` versions and `reactivated` (whether
the plugin had to be reactivated after the file swap, which WordPress does
automatically mid-upgrade). It's a no-op (`updated: false`) if already current, and
disabled outright (404) until a deploy key exists.

**To set it up**, on **this same repo you already pointed Update source at**:

1. **On the WordPress site** — Settings → MS365 Calendar → **Deploy key**: paste a
   long random secret (or define `MS365CAL_DEPLOY_KEY` in `wp-config.php`, which takes
   precedence). Blank keeps the endpoint disabled.
2. **On the GitHub repo** — Settings → Secrets and variables → Actions:
   - **Secret** `SITE_DEPLOY_KEY` — the same value as the deploy key above.
   - **Variable** `DEPLOY_URL` — the site's base URL, e.g. `https://example.com`.

Once both are set, `.github/workflows/release.yml`'s last step calls the endpoint
automatically right after every release is published — pushing a `vX.Y.Z` tag now
takes the site from tag to installed update in seconds, with no manual "check for
updates" click needed. It's best-effort: if the trigger fails or isn't configured, the
release still publishes normally and the site picks it up on its next periodic check.

You can also call the endpoint yourself at any time (`curl`, a different CI system,
etc.) — it doesn't have to go through this repo's Action.

## License

GPL-2.0-or-later. The bundled plugin-update-checker library is MIT licensed
(see `plugin-update-checker/license.txt`).
