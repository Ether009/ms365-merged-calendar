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
[ms365_calendar]                                   all calendars, 14-day window
[ms365_calendar calendars="eng,events" days="30"]  specific slugs and window size
```

## Automatic updates

The plugin self-updates from this repo's **GitHub Releases** using
[plugin-update-checker](https://github.com/YahnisElsts/plugin-update-checker) (vendored
under `plugin-update-checker/`). WordPress compares each release's plugin `Version:`
header against the installed one and offers a one-click update in wp-admin.

### Cutting a release

1. Bump the `Version:` header in `ms365-merged-calendar.php`.
2. Commit, then tag and push:

   ```bash
   git commit -am "Release 2.0.1"
   git tag v2.0.1
   git push && git push --tags
   ```

3. The `Release` GitHub Action builds `ms365-merged-calendar.zip` and publishes a
   release for the tag. Installed sites pick it up on their next update check (or via
   **Dashboard → Updates → Check again**).

The tag version and the `Version:` header should match; the header is what actually
drives whether an update is offered.

## Development

Everything (settings, REST endpoint, shortcode, admin UI, front-end assets) lives in
`ms365-merged-calendar.php` by design. Before every commit it must pass **both**:

```bash
php -l ms365-merged-calendar.php
php phpcs.phar --standard=WordPress-Extra --extensions=php ms365-merged-calendar.php
```

with zero violations. Keep syntax PHP 7.4 compatible. See [CLAUDE.md](CLAUDE.md) for
architecture, the caching/throttling model, the debug endpoint, and the full WPCS
tooling setup.

## License

GPL-2.0-or-later. The bundled plugin-update-checker library is MIT licensed
(see `plugin-update-checker/license.txt`).
