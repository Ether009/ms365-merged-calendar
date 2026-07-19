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
[ms365_calendar]                          all calendars, current week (Monday–Sunday)
[ms365_calendar calendars="eng,events"]   specific slugs
```

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

## License

GPL-2.0-or-later. The bundled plugin-update-checker library is MIT licensed
(see `plugin-update-checker/license.txt`).
