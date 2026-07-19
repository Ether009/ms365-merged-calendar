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

The plugin checks this repo's **GitHub Releases** for newer versions and offers a
one-click update in **wp-admin → Plugins**, the same way any other plugin update
would appear — no manual download needed after the initial install.

## License

GPL-2.0-or-later. The bundled plugin-update-checker library is MIT licensed
(see `plugin-update-checker/license.txt`).
