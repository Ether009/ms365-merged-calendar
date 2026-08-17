<?php
/**
 * Plugin Name:       MS365 Merged Calendar (Async)
 * Description:        Merge calendars from Microsoft 365 groups and shared mailboxes into one filterable, windowed list. Events load asynchronously per view via a REST endpoint; prev/next paging with client-side window caching.
 * Version:           2.28.0
 * Requires PHP:      7.4
 * Author:            You
 * License:           GPL-2.0-or-later
 */

if ( ! defined( 'ABSPATH' ) ) {
	exit;
}

define( 'MS365CAL_OPTION', 'ms365cal_settings' );
define( 'MS365CAL_TOKEN_TRANSIENT', 'ms365cal_token' );
define( 'MS365CAL_MAX_WINDOW', 62 ); // hard cap on days per request
define( 'MS365CAL_RATE_MAX', 30 );   // max REST requests per IP per window (0 = disabled)
define( 'MS365CAL_RATE_WINDOW', 60 ); // rate-limit window, seconds
define( 'MS365CAL_BACKOFF_MAX', 300 ); // max seconds to pause Graph calls after a 429/503
define( 'MS365CAL_MAX_PAGES', 20 );    // safety cap on Graph nextLink pages per calendar (× the configured page size)
define( 'MS365CAL_MAX_MASTERS', 200 ); // recurrence series masters resolved per request (via $batch, 20/call)

/**
 * ---------------------------------------------------------------------------
 *  Automatic updates from GitHub (plugin-update-checker)
 * ---------------------------------------------------------------------------
 *  One-click updates in wp-admin, driven by GitHub Releases of a repo set via the
 *  'update_repo' setting or MS365CAL_UPDATE_REPO (see ms365cal_cred()) — e.g.
 *  "Ether009/ms365-merged-calendar" or a full GitHub URL. No repo configured (the
 *  default) means self-update stays inactive; so does a missing updater library
 *  (e.g. a bare single-file install) or an absent 'plugin-update-checker/' folder.
 *  Either way the plugin still runs, it just won't self-update. WordPress compares
 *  each release's plugin "Version:" header against the installed one.
 */
function ms365cal_init_updates() {
	$loader = __DIR__ . '/plugin-update-checker/plugin-update-checker.php';
	if ( ! is_readable( $loader ) ) {
		return;
	}
	require_once $loader;

	if ( ! class_exists( 'YahnisElsts\\PluginUpdateChecker\\v5\\PucFactory' ) ) {
		return;
	}

	$repo_url = ms365cal_normalize_repo_url( ms365cal_cred( 'update_repo' ) );
	if ( '' === $repo_url ) {
		return; // no repo configured — self-update stays inactive, same as a missing library.
	}

	$checker = \YahnisElsts\PluginUpdateChecker\v5\PucFactory::buildUpdateChecker(
		$repo_url,
		__FILE__,
		'ms365-merged-calendar'
	);

	// Prefer a curated release-asset zip when a release has one (the CI build
	// attaches it); fall back to GitHub's auto-generated source archive.
	$api = $checker->getVcsApi();
	if ( method_exists( $api, 'enableReleaseAssets' ) ) {
		$api->enableReleaseAssets();
	}

	ms365cal_update_checker( $checker ); // stash it for the self-update endpoint.
}
ms365cal_init_updates();

/**
 * Accessor for the plugin-update-checker instance built in ms365cal_init_updates().
 * Pass an instance to store it; call with no argument to retrieve it (or null if the
 * updater library wasn't available). Lets the self-update endpoint force a re-check.
 */
function ms365cal_update_checker( $set = null ) {
	static $checker = null;
	if ( null !== $set ) {
		$checker = $set;
	}
	return $checker;
}

/**
 * Normalise an auto-updater repo setting into the full URL PUC expects. Accepts
 * either a bare "owner/repo" shorthand (the form shown in the settings UI) or a
 * full GitHub URL (e.g. pasted straight from the browser address bar). Returns ''
 * unchanged so the caller can fall back to the built-in default repo.
 */
function ms365cal_normalize_repo_url( $repo ) {
	$repo = trim( (string) $repo );
	if ( '' === $repo ) {
		return '';
	}
	if ( ! preg_match( '#^https?://#i', $repo ) ) {
		$repo = 'https://github.com/' . ltrim( $repo, '/' );
	}
	return untrailingslashit( $repo ) . '/';
}

/**
 * ---------------------------------------------------------------------------
 *  Settings
 * ---------------------------------------------------------------------------
 *  Credentials may be defined in wp-config.php (they win over DB options):
 *      define( 'MS365CAL_TENANT_ID', '...' );
 *      define( 'MS365CAL_CLIENT_ID', '...' );
 *      define( 'MS365CAL_CLIENT_SECRET', '...' );
 *      define( 'MS365CAL_DEPLOY_KEY', '...' );  // self-update endpoint; also settable in the UI
 *      define( 'MS365CAL_UPDATE_REPO', '...' ); // auto-updater source repo; also settable in the UI
 */
function ms365cal_get_settings() {
	$defaults = array(
		'tenant_id'           => '',
		'client_id'           => '',
		'client_secret'       => '',
		'cache_minutes'       => 20,
		'cache_grace_minutes' => 60,
		'timezone'            => wp_timezone_string(),
		'rate_max'            => MS365CAL_RATE_MAX,
		'rate_window'         => MS365CAL_RATE_WINDOW,
		'show_outlook'        => false,
		'show_recurrence'     => false,
		'timeline_interval'   => 15,
		'events_top'          => 100,
		'lazy_body'           => true,
		'deploy_key'          => '',
		'update_repo'         => '',
		'calendars'           => array(),
	);
	$saved    = get_option( MS365CAL_OPTION, array() );
	return wp_parse_args( is_array( $saved ) ? $saved : array(), $defaults );
}

function ms365cal_cred( $key ) {
	$const = 'MS365CAL_' . strtoupper( $key );
	if ( defined( $const ) && constant( $const ) ) {
		return constant( $const );
	}
	$s = ms365cal_get_settings();
	return isset( $s[ $key ] ) ? $s[ $key ] : '';
}

/**
 * ---------------------------------------------------------------------------
 *  Graph auth (app-only / client credentials)
 * ---------------------------------------------------------------------------
 */
function ms365cal_get_token() {
	$cached = get_transient( MS365CAL_TOKEN_TRANSIENT );
	if ( $cached ) {
		return $cached;
	}

	$tenant = ms365cal_cred( 'tenant_id' );
	$client = ms365cal_cred( 'client_id' );
	$secret = ms365cal_cred( 'client_secret' );

	if ( ! $tenant || ! $client || ! $secret ) {
		return new WP_Error( 'ms365cal_no_creds', 'Azure credentials are not configured.' );
	}

	$resp = wp_remote_post(
		"https://login.microsoftonline.com/{$tenant}/oauth2/v2.0/token",
		array(
			'timeout' => 15,
			'body'    => array(
				'grant_type'    => 'client_credentials',
				'client_id'     => $client,
				'client_secret' => $secret,
				'scope'         => 'https://graph.microsoft.com/.default',
			),
		)
	);

	if ( is_wp_error( $resp ) ) {
		return $resp;
	}

	$body = json_decode( wp_remote_retrieve_body( $resp ), true );
	if ( empty( $body['access_token'] ) ) {
		$msg = isset( $body['error_description'] ) ? $body['error_description'] : 'Token request failed.';
		return new WP_Error( 'ms365cal_token_failed', $msg );
	}

	$expires = isset( $body['expires_in'] ) ? (int) $body['expires_in'] : 3600;
	set_transient( MS365CAL_TOKEN_TRANSIENT, $body['access_token'], max( 60, $expires - 300 ) );

	return $body['access_token'];
}

/**
 * Build the calendarView URL. This is the only place group vs. mailbox differs.
 * $top is the configurable page size (settings['events_top']) — a calendar's page
 * only carries @odata.nextLink (triggering a pagination/straggler fallback) once its
 * event count in the window exceeds this value, so raising it above the busiest
 * calendar's normal event count lets most fetches skip pagination entirely. When
 * $lazy_body is true, 'body' is left out of $select entirely — the list view never
 * needs it, only the on-demand event-body endpoint does (see
 * ms365cal_rest_event_body()) — which cuts both the Graph response payload and the
 * per-event HTML-sanitisation work on every cold list fetch. 'id' is always selected
 * explicitly (rather than relying on Graph's default inclusion) since lazy mode needs
 * it to fetch a specific event's body later. Lazy mode selects the much smaller
 * 'bodyPreview' (a ~255-char plain-text truncation) instead, purely so the row can
 * flag whether there's anything to fetch at all — see ms365cal_build_rows_from_page().
 */
function ms365cal_view_url( $cal, $start_iso, $end_iso, $top, $lazy_body ) {
	$source = rawurlencode( $cal['source'] );
	$base   = ( 'mailbox' === $cal['type'] )
		? "https://graph.microsoft.com/v1.0/users/{$source}/calendarView"
		: "https://graph.microsoft.com/v1.0/groups/{$source}/calendarView";

	$select = 'id,subject,start,end,location,isAllDay,onlineMeeting,webLink,type,seriesMasterId';
	if ( $lazy_body ) {
		$select .= ',bodyPreview';
	} else {
		$select .= ',body';
	}

	return add_query_arg(
		array(
			'startDateTime' => $start_iso,
			'endDateTime'   => $end_iso,
			'$orderby'      => 'start/dateTime',
			'$top'          => (int) $top,
			'$select'       => $select,
		),
		$base
	);
}

/**
 * Auto-detect group vs. shared-mailbox from the source value itself, instead
 * of making the admin pick a type: a group's source is always an object-ID
 * GUID (8-4-4-4-12 hex), a shared mailbox's is always an email address —
 * the two formats never overlap, so this is unambiguous.
 */
function ms365cal_detect_calendar_type( $source ) {
	$is_guid = (bool) preg_match( '/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i', trim( $source ) );
	return $is_guid ? 'group' : 'mailbox';
}

/**
 * Relative ($batch-friendly) base for a calendar's /events collection: the path
 * without the Graph host, e.g. "/groups/{id}/events" or "/users/{addr}/events".
 * Used to build $batch sub-request URLs when resolving recurring series masters.
 */
function ms365cal_events_rel_base( $cal ) {
	$source = rawurlencode( $cal['source'] );
	return ( 'mailbox' === $cal['type'] )
		? "/users/{$source}/events"
		: "/groups/{$source}/events";
}

/**
 * Human-readable (Swedish) recurrence summary from a Graph recurrence object, e.g.
 * "Upprepas varje vecka på mån, ons" or "Upprepas varje månad den 15:e till 1 dec 2026".
 */
function ms365cal_format_recurrence( $rec ) {
	if ( empty( $rec['pattern']['type'] ) ) {
		return '';
	}
	$p        = $rec['pattern'];
	$interval = isset( $p['interval'] ) ? max( 1, (int) $p['interval'] ) : 1;

	$abbr = array(
		'monday'    => 'mån',
		'tuesday'   => 'tis',
		'wednesday' => 'ons',
		'thursday'  => 'tor',
		'friday'    => 'fre',
		'saturday'  => 'lör',
		'sunday'    => 'sön',
	);
	$ord  = array(
		'first'  => 'första',
		'second' => 'andra',
		'third'  => 'tredje',
		'fourth' => 'fjärde',
		'last'   => 'sista',
	);
	$dow  = array();
	$days = array();
	if ( ! empty( $p['daysOfWeek'] ) && is_array( $p['daysOfWeek'] ) ) {
		foreach ( $p['daysOfWeek'] as $d ) {
			$key    = strtolower( $d );
			$dow[]  = $key;
			$days[] = isset( $abbr[ $key ] ) ? $abbr[ $key ] : $key;
		}
	}
	$day_list = implode( ', ', $days );
	$weekdays = array( 'monday', 'tuesday', 'wednesday', 'thursday', 'friday' );
	if ( 5 === count( $dow ) && ! array_diff( $weekdays, $dow ) ) {
		$day_list = 'vardagar';
	}
	$idx_key = isset( $p['index'] ) ? $p['index'] : 'first';
	$index   = isset( $ord[ $idx_key ] ) ? $ord[ $idx_key ] : 'första';

	switch ( $p['type'] ) {
		case 'daily':
			$s = $interval > 1 ? "var {$interval}:e dag" : 'dagligen';
			break;
		case 'weekly':
			$base = $interval > 1 ? "var {$interval}:e vecka" : 'varje vecka';
			$s    = $day_list ? "{$base} på {$day_list}" : $base;
			break;
		case 'absoluteMonthly':
			$base = $interval > 1 ? "var {$interval}:e månad" : 'varje månad';
			$dom  = isset( $p['dayOfMonth'] ) ? (int) $p['dayOfMonth'] : 0;
			$s    = $dom ? "{$base} den {$dom}:e" : $base;
			break;
		case 'relativeMonthly':
			$base = $interval > 1 ? "var {$interval}:e månad" : 'varje månad';
			$s    = trim( "{$base} den {$index} {$day_list}" );
			break;
		case 'absoluteYearly':
			$s = 'årligen';
			break;
		case 'relativeYearly':
			$s = trim( "årligen den {$index} {$day_list}" );
			break;
		default:
			$s = 'enligt schema';
	}

	$label = 'Upprepas ' . $s;

	if ( isset( $rec['range']['type'] ) && 'endDate' === $rec['range']['type'] && ! empty( $rec['range']['endDate'] ) ) {
		$ts = strtotime( $rec['range']['endDate'] );
		if ( $ts ) {
			$label .= ' till ' . wp_date( 'j M Y', $ts );
		}
	}

	return $label;
}

/**
 * Resolve a set of recurring-series master IDs (for one calendar) to human-readable
 * recurrence strings using Microsoft Graph's $batch endpoint — up to 20 masters per
 * request instead of one GET each. The caller passes de-duplicated IDs; a per-request
 * budget (MS365CAL_MAX_MASTERS, shared across calendars via a static) bounds how many
 * are resolved so a pathological window can't fan out without limit. Anything not
 * resolved (over budget, HTTP failure, or a master with no pattern) is simply omitted
 * from the returned map, and the caller shows a generic "Återkommande händelse".
 *
 * @param array  $cal        Calendar config.
 * @param array  $master_ids Series-master IDs to resolve.
 * @param string $token      Graph bearer token.
 * @return array Map of master_id => recurrence text, for the ones resolved.
 */
function ms365cal_fetch_recurrence_map( $cal, $master_ids, $token ) {
	static $budget = MS365CAL_MAX_MASTERS;

	$map        = array();
	$master_ids = array_values( array_unique( $master_ids ) );
	if ( empty( $master_ids ) || $budget <= 0 ) {
		return $map;
	}

	$master_ids = array_slice( $master_ids, 0, $budget );
	$budget     = $budget - count( $master_ids );
	$base       = ms365cal_events_rel_base( $cal );

	foreach ( array_chunk( $master_ids, 20 ) as $chunk ) {
		$requests = array();
		$id_map   = array(); // batch sub-request id => master_id.
		foreach ( $chunk as $i => $mid ) {
			$rid            = (string) $i;
			$id_map[ $rid ] = $mid;
			$requests[]     = array(
				'id'     => $rid,
				'method' => 'GET',
				'url'    => $base . '/' . rawurlencode( $mid ) . '?$select=recurrence',
			);
		}

		$resp = wp_remote_post(
			'https://graph.microsoft.com/v1.0/$batch',
			array(
				'timeout' => 20,
				'headers' => array(
					'Authorization' => 'Bearer ' . $token,
					'Content-Type'  => 'application/json',
					'Accept'        => 'application/json',
				),
				'body'    => wp_json_encode( array( 'requests' => $requests ) ),
			)
		);

		if ( is_wp_error( $resp ) || 200 !== (int) wp_remote_retrieve_response_code( $resp ) ) {
			continue; // whole chunk unresolved; caller falls back to the generic label.
		}

		$body = json_decode( wp_remote_retrieve_body( $resp ), true );
		if ( empty( $body['responses'] ) || ! is_array( $body['responses'] ) ) {
			continue;
		}

		foreach ( $body['responses'] as $sub ) {
			$rid = isset( $sub['id'] ) ? (string) $sub['id'] : '';
			if ( '' === $rid || ! isset( $id_map[ $rid ] ) ) {
				continue;
			}

			$status = isset( $sub['status'] ) ? (int) $sub['status'] : 0;
			$rec    = isset( $sub['body']['recurrence'] ) ? $sub['body']['recurrence'] : null;
			if ( 200 !== $status || ! is_array( $rec ) ) {
				continue;
			}

			$text = ms365cal_format_recurrence( $rec );
			if ( '' !== $text ) {
				$map[ $id_map[ $rid ] ] = $text;
			}
		}
	}

	return $map;
}

/**
 * Compact "when" label for the list, showing the end so multi-day and timed
 * events don't look like single all-day entries.
 */
function ms365cal_when_label( $start, $end, $all_day, $time_fmt ) {
	// Render in the same zone the times were parsed in (the plugin's configured
	// timezone), not wp_date()'s default of the WordPress site timezone. Those two
	// can differ, and for an all-day event pinned to midnight the difference flips
	// the displayed day — so the label would disagree with the day it's grouped under.
	$tz = $start->getTimezone();
	if ( $all_day ) {
		// Graph's all-day end is exclusive midnight, so the last real day is -1.
		$end_incl = ( clone $end )->modify( '-1 day' );
		if ( $end_incl->format( 'Y-m-d' ) <= $start->format( 'Y-m-d' ) ) {
			return 'Heldag';
		}
		return 'Heldag · ' . wp_date( 'j M', $start->getTimestamp(), $tz ) . ' – ' . wp_date( 'j M', $end_incl->getTimestamp(), $tz );
	}

	$s_time = wp_date( $time_fmt, $start->getTimestamp(), $tz );
	$e_time = wp_date( $time_fmt, $end->getTimestamp(), $tz );

	if ( $start->format( 'Y-m-d' ) === $end->format( 'Y-m-d' ) ) {
		return $s_time . ' – ' . $e_time;
	}

	return wp_date( 'j M', $start->getTimestamp(), $tz ) . ' ' . $s_time
		. ' – ' . wp_date( 'j M', $end->getTimestamp(), $tz ) . ' ' . $e_time;
}

/**
 * Strip the auto-inserted Teams/online-meeting join block from a plain-text event
 * body. Outlook wraps that block between long horizontal-rule lines (a run of
 * underscore characters); everything from the first such rule to the last is
 * boilerplate (join instructions, dial-in, legal notice) that's already covered by
 * the separate join link and location shown in the UI, so it's just noise in the
 * description. Text before/after the rule(s) — the organiser's own notes — is kept.
 */
function ms365cal_strip_meeting_boilerplate( $text ) {
	if ( '' === $text || false === strpos( $text, '___' ) ) {
		return $text;
	}

	if ( ! preg_match_all( '/_{10,}/', $text, $matches, PREG_OFFSET_CAPTURE ) ) {
		return $text;
	}

	$first = $matches[0][0];
	$last  = end( $matches[0] );

	if ( count( $matches[0] ) === 1 ) {
		// A single rule with no closing line: the block runs to the end of the body.
		$stripped = substr( $text, 0, $first[1] );
	} else {
		$stripped = substr( $text, 0, $first[1] ) . substr( $text, $last[1] + strlen( $last[0] ) );
	}

	$stripped = preg_replace( '/\n{3,}/', "\n\n", $stripped );
	return trim( (string) $stripped );
}

/**
 * Collapse an immediately-repeated link in an event body down to one occurrence.
 * Graph's HTML-to-plain-text conversion (and, it turns out, the HTML body itself)
 * can duplicate a link back to back — e.g. a link embedded as an Outlook "Smart
 * Link" preview card renders once for the visible link and once more for the card
 * reference, invisible in Outlook's own rendered view. In HTML this shows up as two
 * adjacent <a> tags sharing the same href; kept as a secondary pass is the older
 * plain-text shape (the same bare URL twice in a row) in case a body ever arrives
 * that way. Only merges when the *exact same* link repeats with nothing but
 * whitespace (or a stray <br>) between — intentionally repeated links elsewhere in
 * the body are left alone. Runs before sanitisation, on Graph's raw markup.
 */
function ms365cal_dedupe_repeated_links( $text ) {
	if ( '' === $text || false === stripos( $text, 'http' ) ) {
		return $text;
	}

	$anchor_pattern = '/(<a\s[^>]*href="([^"]+)"[^>]*>.*?<\/a>)(?:\s|&nbsp;|<br\s*\/?>)*<a\s[^>]*href="\2"[^>]*>.*?<\/a>/is';
	$prev           = null;
	while ( $prev !== $text ) {
		$prev   = $text;
		$merged = preg_replace( $anchor_pattern, '$1', $text );
		if ( null !== $merged ) {
			$text = $merged;
		}
	}

	$plain = preg_replace( '/(<?https?:\/\/[^\s<>]+>?)(\s+\1)+/', '$1', $text );
	return null !== $plain ? $plain : $text;
}

/**
 * Sanitise event-body HTML from Graph before it reaches the browser as real markup
 * (rather than escaped plain text), so a link keeps its original anchor text instead
 * of showing as a bare URL. Event bodies can be set by external meeting organisers,
 * so this goes through wp_kses() with a conservative allowlist — basic text
 * structure and links only; no style/class/id/data-* attributes, images, or
 * scripts — the same trusted sanitiser WordPress uses for post content, just
 * narrower. Every surviving link is then forced to open safely in a new tab
 * (target="_blank" rel="noopener noreferrer"), overriding whatever — if anything —
 * Graph's source HTML set, since these links point wherever the organiser put them.
 */
function ms365cal_sanitize_event_html( $html ) {
	if ( '' === $html ) {
		return '';
	}

	$allowed = array(
		'a'      => array( 'href' => true ),
		'p'      => array(),
		'br'     => array(),
		'div'    => array(),
		'span'   => array(),
		'b'      => array(),
		'strong' => array(),
		'i'      => array(),
		'em'     => array(),
		'u'      => array(),
		'ul'     => array(),
		'ol'     => array(),
		'li'     => array(),
	);

	$clean = wp_kses( $html, $allowed );

	// The allowlist above only ever lets 'href' through, so every surviving <a> is
	// exactly "<a href=...>" at this point — safe to target unconditionally.
	$with_target = preg_replace( '/<a\s+href=/i', '<a target="_blank" rel="noopener noreferrer" href=', $clean );
	if ( null !== $with_target ) {
		$clean = $with_target;
	}

	// Outlook's editor tends to leave runs of empty paragraphs/line breaks; collapse
	// them so the description doesn't end up mostly whitespace.
	$collapsed = preg_replace(
		array( '/(?:<div>(?:&nbsp;|\s)*<\/div>\s*){2,}/i', '/(?:<br\s*\/?>\s*){3,}/i' ),
		array( '', '<br><br>' ),
		$clean
	);
	if ( null !== $collapsed ) {
		$clean = $collapsed;
	}

	return trim( $clean );
}

/**
 * Build event rows from one page of raw Graph calendarView results (the 'value'
 * array). Shared by the sequential per-calendar fetch (ms365cal_fetch_one(), used
 * for the admin debug diagnostic and as a pagination fallback) and the batched
 * fetch (ms365cal_fetch_calendars_batched()), so the event-shaping logic — time
 * labels, the left-column t1/t2, longSpan, body sanitisation, recurrence-master
 * collection — lives in exactly one place. $need_masters accumulates
 * seriesMasterId => true for any recurring occurrence found, by reference. Every
 * recurring row still gets a cheap boolean 'recurring' flag from Graph's initial
 * calendarView data; when $show_recur is false, only the expensive recurrence text
 * collection is skipped (every row's 'recur' stays '') so the caller never issues
 * the extra $batch calls to resolve it.
 * When $lazy_body is true, $raw_events never carried 'body' in the first place (see
 * ms365cal_view_url()), so every row's 'body' stays '' and the dedupe/boilerplate/
 * sanitise pipeline never runs here; the row instead carries the raw Graph event 'id'
 * (so the client can fetch that one event's body on demand via
 * ms365cal_rest_event_body()) and a 'hasBody' flag derived from the cheap
 * 'bodyPreview' field, so the client can skip rendering an expand affordance at all
 * on events with nothing to show, instead of discovering that only after a click.
 * Every row also carries a plain-text 'preview' snippet (~160 chars, from
 * bodyPreview in lazy mode or the sanitised body otherwise) for the Expanded
 * layout, and a 'carriedOver' flag — true when an event is pinned to today
 * (still running) but actually *started* before this window, so the client can
 * surface it in a separate "ongoing since before" section instead of letting it
 * read as a same-day event.
 *
 * @return array Row objects; a row with a series master still carries a temporary
 *               '_master' key for the caller to resolve and strip afterward.
 */
function ms365cal_build_rows_from_page( $cal, $raw_events, $zone, $time_fmt, $window_start, $today0, $window_end, &$need_masters, $show_recur, $lazy_body ) {
	$out = array();

	foreach ( $raw_events as $e ) {
		$raw_start = isset( $e['start']['dateTime'] ) ? $e['start']['dateTime'] : '';
		if ( '' === $raw_start ) {
			continue;
		}
		try {
			$st = new DateTime( $raw_start, $zone );
		} catch ( Exception $ex ) {
			continue;
		}

		$en      = null;
		$raw_end = isset( $e['end']['dateTime'] ) ? $e['end']['dateTime'] : '';
		if ( '' !== $raw_end ) {
			try {
				$en = new DateTime( $raw_end, $zone );
			} catch ( Exception $ex ) {
				$en = null;
			}
		}

		$allday = ! empty( $e['isAllDay'] );

		// True for a multi-day event whose real end falls after this window (e.g.
		// a months-long ongoing series). Sent to the client so it can tell a
		// genuinely-empty week from one that only "has" a long-running background
		// event — used to decide whether to default to next week instead.
		$multiday  = $en && ( $en->format( 'Y-m-d' ) !== $st->format( 'Y-m-d' ) );
		$long_span = $multiday && ( $en > $window_end );

		// Weekly grouping: an event still running today (started earlier, ends
		// later) is pinned to today so it shows as current rather than under a
		// past day; one that started before the window but already ended is
		// pinned to the window start; everything else keeps its own start day.
		$pinned_today = ( null !== $en && $en > $today0 && $st < $today0 );
		// A pinned-to-today event whose *actual* start predates this window
		// entirely (started last week or earlier) looks, once pinned, exactly
		// like a same-day event — nothing marked it as having started earlier.
		// Flagged so the client can pull it out into its own "ongoing since
		// before" section above the day list instead of burying that fact.
		$carried_over = $pinned_today && ( $st < $window_start );
		if ( $pinned_today ) {
			$eff = clone $today0;
		} elseif ( $st < $window_start ) {
			$eff = clone $window_start;
		} else {
			$eff = $st;
		}

		$when = $en
			? ms365cal_when_label( $st, $en, $allday, $time_fmt )
			: ( $allday ? 'Heldag' : wp_date( $time_fmt, $st->getTimestamp(), $zone ) );

		// Left-hand time column that flanks the rail. Single-day events show the
		// times (or "Heldag"); multi-day events show the start date/time at the
		// top and the end date/time at the bottom so the span stays visible
		// without expanding (all-day skips the time).
		$s_time = wp_date( $time_fmt, $st->getTimestamp(), $zone );
		if ( $allday ) {
			// Graph's all-day end is exclusive midnight; the last real day is -1.
			$end_incl = $en ? ( clone $en )->modify( '-1 day' ) : clone $st;
			if ( $end_incl->format( 'Y-m-d' ) > $st->format( 'Y-m-d' ) ) {
				$t1 = wp_date( 'j M', $st->getTimestamp(), $zone );
				$t2 = wp_date( 'j M', $end_incl->getTimestamp(), $zone );
			} else {
				$t1 = 'Heldag';
				$t2 = '';
			}
		} elseif ( $en && $en->format( 'Y-m-d' ) !== $st->format( 'Y-m-d' ) ) {
			$t1 = wp_date( 'j M', $st->getTimestamp(), $zone ) . ' ' . $s_time;
			$t2 = wp_date( 'j M', $en->getTimestamp(), $zone ) . ' ' . wp_date( $time_fmt, $en->getTimestamp(), $zone );
		} else {
			$t1 = $s_time;
			$t2 = $en ? wp_date( $time_fmt, $en->getTimestamp(), $zone ) : '';
		}

		// Recurrence: calendarView already tells us whether an event is part of a
		// series (type + seriesMasterId), which is cheap enough for the icon. The
		// human-readable pattern still lives on the series master, so only collect
		// master IDs for the optional slower lookup when that setting is enabled.
		$recur     = '';
		$master_id = isset( $e['seriesMasterId'] ) ? $e['seriesMasterId'] : '';
		$etype     = isset( $e['type'] ) ? $e['type'] : '';
		$recurring = ( '' !== $master_id ) || in_array( $etype, array( 'occurrence', 'exception', 'seriesMaster' ), true );
		if ( $show_recur ) {
			if ( '' !== $master_id ) {
				$recur = 'Återkommande händelse';

				$need_masters[ $master_id ] = true;
			} elseif ( in_array( $etype, array( 'occurrence', 'exception' ), true ) ) {
				$recur = 'Återkommande händelse';
			}
		}

		// A local $event_body — kept distinct from the caller's response-JSON
		// $body/$page_body variable, which a same-named local here used to shadow
		// (broke pagination: isset() on a non-numeric string offset is false, so
		// @odata.nextLink was never seen once any event on the page had reached
		// this point).
		$is_online  = ! empty( $e['onlineMeeting'] );
		$event_body = '';
		// Plain-text, ~1-2 line snippet for the Expanded layout — computed either
		// way, since lazy mode already has bodyPreview in hand for the has-body
		// check below and non-lazy mode already has the full sanitised body.
		// wp_html_excerpt() strips any remaining tags and truncates cleanly.
		$preview = '';
		if ( $lazy_body ) {
			// bodyPreview is a plain-text truncation of the *raw* body, taken before
			// our boilerplate-stripping would run — so an online meeting whose only
			// content is the auto-inserted join block can show a non-empty preview
			// even though the sanitised result would end up empty. Treated as an
			// acceptable false positive (still clickable, just finds nothing) rather
			// than fetching the real body here just to check.
			$raw_preview = isset( $e['bodyPreview'] ) ? trim( (string) $e['bodyPreview'] ) : '';
			if ( '' !== $raw_preview ) {
				$preview = wp_html_excerpt( $raw_preview, 160, '…' );
			}
		} else {
			$event_body = isset( $e['body']['content'] ) ? trim( (string) $e['body']['content'] ) : '';
			if ( '' !== $event_body ) {
				$event_body = ms365cal_dedupe_repeated_links( $event_body );
				if ( $is_online ) {
					// The join link and "Teams meeting" location are shown separately in
					// the UI, so the auto-inserted join/dial-in block in the body is just
					// noise — strip it, keeping any real notes the organiser wrote.
					$event_body = ms365cal_strip_meeting_boilerplate( $event_body );
				}
				$event_body = ms365cal_sanitize_event_html( $event_body );
				$preview    = wp_html_excerpt( $event_body, 160, '…' );
			}
		}
		$has_body = '' !== $preview;

		$row = array(
			'cal'         => $cal['slug'],
			'title'       => isset( $e['subject'] ) && '' !== $e['subject'] ? $e['subject'] : '(ingen rubrik)',
			'sort'        => $eff->format( 'Y-m-d\TH:i:s' ),
			'dayKey'      => $eff->format( 'Y-m-d' ),
			'dayLabel'    => wp_date( 'D j M', $eff->getTimestamp(), $zone ),
			't1'          => $t1,
			't2'          => $t2,
			'when'        => $when,
			'recur'       => $recur,
			'recurring'   => $recurring,
			'longSpan'    => $long_span,
			'carriedOver' => $carried_over,
			'body'        => $event_body,
			'preview'     => $preview,
			'location'    => isset( $e['location']['displayName'] ) ? $e['location']['displayName'] : '',
			'online'      => $is_online,
			'joinUrl'     => isset( $e['onlineMeeting']['joinUrl'] ) ? $e['onlineMeeting']['joinUrl'] : '',
			'link'        => isset( $e['webLink'] ) ? $e['webLink'] : '',
			// Raw start/end (site-timezone-naive, same format as 'sort') and the
			// all-day flag — 'sort'/'dayKey' are the *pinned* (carried-over/today-
			// pinned) values used for list grouping, not necessarily the event's
			// real timing, so the Day/Week calendar-grid views (which position
			// events by actual clock time) need these instead.
			'start'       => $st->format( 'Y-m-d\TH:i:s' ),
			'end'         => $en ? $en->format( 'Y-m-d\TH:i:s' ) : null,
			'allDay'      => $allday,
		);
		if ( $show_recur && '' !== $master_id ) {
			$row['_master'] = $master_id; // temp key, stripped after resolution.
		}
		if ( $lazy_body ) {
			$row['id']      = isset( $e['id'] ) ? $e['id'] : '';
			$row['hasBody'] = $has_body;
		}
		$out[] = $row;
	}

	return $out;
}

/**
 * Single Graph GET with one short inline retry for a tiny Retry-After. Longer
 * throttles fall through to the caller's back-off handling.
 */
function ms365cal_http_get( $url, $args ) {
	$resp = wp_remote_get( $url, $args );
	if ( ! is_wp_error( $resp ) ) {
		$code = (int) wp_remote_retrieve_response_code( $resp );
		if ( 429 === $code || 503 === $code ) {
			$ra = (int) wp_remote_retrieve_header( $resp, 'retry-after' );
			if ( $ra > 0 && $ra <= 2 ) {
				sleep( $ra );
				$resp = wp_remote_get( $url, $args );
			}
		}
	}
	return $resp;
}

/**
 * Fetch one calendar, following Graph's @odata.nextLink so busy calendars
 * return everything (up to MS365CAL_MAX_PAGES × page size), not just page one.
 *
 * @return array{events:array,throttled:bool,retry_after:int,status:int,error:string}
 *         'throttled' is true when Graph returned 429/503; 'retry_after' carries
 *         Graph's Retry-After (seconds) when present. 'status' is the HTTP code
 *         (0 for a transport error) and 'error' a short reason, both for debug.
 */
function ms365cal_fetch_one( $cal, $token, $start_iso, $end_iso, $tz, $top, $show_recur, $lazy_body ) {
	$args = array(
		'timeout' => 20,
		'headers' => array(
			'Authorization' => 'Bearer ' . $token,
			'Prefer'        => 'outlook.timezone="' . $tz . '", outlook.body-content-type="html"',
		),
	);

	try {
		$zone = new DateTimeZone( $tz );
	} catch ( Exception $e ) {
		$zone = new DateTimeZone( 'UTC' );
	}
	$time_fmt = get_option( 'time_format', 'H:i' );

	try {
		$window_start = new DateTime( $start_iso, $zone );
	} catch ( Exception $e ) {
		$window_start = new DateTime( 'now', $zone );
	}
	try {
		$window_end = new DateTime( $end_iso, $zone );
	} catch ( Exception $e ) {
		$window_end = clone $window_start;
	}

	$today0 = new DateTime( 'now', $zone );
	$today0->setTime( 0, 0, 0 );

	$out          = array();
	$need_masters = array(); // seriesMasterId => true, resolved in one $batch pass after paging.
	$url          = ms365cal_view_url( $cal, $start_iso, $end_iso, $top, $lazy_body );
	$page         = 0;

	while ( $url && $page < MS365CAL_MAX_PAGES ) {
		++$page;
		$resp = ms365cal_http_get( $url, $args );

		if ( is_wp_error( $resp ) ) {
			return array(
				'events'      => array(),
				'throttled'   => false,
				'retry_after' => 0,
				'status'      => 0,
				'error'       => $resp->get_error_message(),
			);
		}

		$code = (int) wp_remote_retrieve_response_code( $resp );

		if ( 429 === $code || 503 === $code ) {
			$ra = (int) wp_remote_retrieve_header( $resp, 'retry-after' );
			return array(
				'events'      => array(),
				'throttled'   => true,
				'retry_after' => $ra > 0 ? $ra : 30,
				'status'      => $code,
				'error'       => '',
			);
		}

		if ( 200 !== $code ) {
			// 403/404/etc. — no events, but not a throttle. Capture Graph's reason.
			$err_body = json_decode( wp_remote_retrieve_body( $resp ), true );
			$err_msg  = isset( $err_body['error']['message'] )
				? $err_body['error']['message']
				: substr( (string) wp_remote_retrieve_body( $resp ), 0, 300 );
			return array(
				'events'      => array(),
				'throttled'   => false,
				'retry_after' => 0,
				'status'      => $code,
				'error'       => $err_msg,
			);
		}

		$page_body = json_decode( wp_remote_retrieve_body( $resp ), true );

		if ( ! empty( $page_body['value'] ) && is_array( $page_body['value'] ) ) {
			$rows = ms365cal_build_rows_from_page( $cal, $page_body['value'], $zone, $time_fmt, $window_start, $today0, $window_end, $need_masters, $show_recur, $lazy_body );
			$out  = array_merge( $out, $rows );
		}

		// Follow pagination. The nextLink is an absolute URL carrying the skiptoken
		// plus the original $select/$top, so pass it through unchanged.
		$url = isset( $page_body['@odata.nextLink'] ) ? $page_body['@odata.nextLink'] : '';
	}

	// One batched pass to turn the collected series-master IDs into readable
	// recurrence strings, then drop the temp key so it never reaches the client.
	$recur_map = ms365cal_fetch_recurrence_map( $cal, array_keys( $need_masters ), $token );
	foreach ( $out as $idx => $row ) {
		if ( isset( $row['_master'] ) ) {
			if ( isset( $recur_map[ $row['_master'] ] ) ) {
				$out[ $idx ]['recur'] = $recur_map[ $row['_master'] ];
			}
			unset( $out[ $idx ]['_master'] );
		}
	}

	return array(
		'events'      => $out,
		'throttled'   => false,
		'retry_after' => 0,
		'status'      => 200,
		'error'       => '',
	);
}

/**
 * Strip the Graph API host+version prefix from an absolute URL (e.g. one built by
 * ms365cal_view_url()), for use as a $batch sub-request's relative "url" field.
 */
function ms365cal_graph_relative_url( $url ) {
	return (string) preg_replace( '#^https://graph\.microsoft\.com/v[0-9.]+#', '', $url );
}

/**
 * Resolve recurring-series master IDs across MULTIPLE calendars in as few Graph
 * $batch calls as possible — one shared budget (MS365CAL_MAX_MASTERS, own static,
 * independent of ms365cal_fetch_recurrence_map()'s), 20 masters per $batch request,
 * mixing different calendars' relative bases within the same request since $batch
 * sub-requests are independent of each other. $refs is a list of [slug, master_id]
 * pairs (already de-duplicated by the caller). $cals_by_slug maps slug => calendar
 * config, needed to build each master's relative URL. Returns a map keyed
 * "slug|master_id" => recurrence text, for whichever were resolved.
 */
function ms365cal_fetch_recurrence_map_multi( $cals_by_slug, $refs, $token ) {
	static $budget = MS365CAL_MAX_MASTERS;

	$map = array();
	if ( empty( $refs ) || $budget <= 0 ) {
		return $map;
	}

	$refs   = array_slice( $refs, 0, $budget );
	$budget = $budget - count( $refs );

	foreach ( array_chunk( $refs, 20 ) as $chunk ) {
		$requests = array();
		$id_map   = array(); // batch sub-request id => "slug|master_id".

		foreach ( $chunk as $i => $ref ) {
			list( $slug, $mid ) = $ref;
			if ( ! isset( $cals_by_slug[ $slug ] ) ) {
				continue;
			}
			$rid            = (string) $i;
			$id_map[ $rid ] = $slug . '|' . $mid;
			$requests[]     = array(
				'id'     => $rid,
				'method' => 'GET',
				'url'    => ms365cal_events_rel_base( $cals_by_slug[ $slug ] ) . '/' . rawurlencode( $mid ) . '?$select=recurrence',
			);
		}
		if ( empty( $requests ) ) {
			continue;
		}

		$resp = wp_remote_post(
			'https://graph.microsoft.com/v1.0/$batch',
			array(
				'timeout' => 20,
				'headers' => array(
					'Authorization' => 'Bearer ' . $token,
					'Content-Type'  => 'application/json',
					'Accept'        => 'application/json',
				),
				'body'    => wp_json_encode( array( 'requests' => $requests ) ),
			)
		);

		if ( is_wp_error( $resp ) || 200 !== (int) wp_remote_retrieve_response_code( $resp ) ) {
			continue; // whole chunk unresolved; caller falls back to the generic label.
		}

		$body = json_decode( wp_remote_retrieve_body( $resp ), true );
		if ( empty( $body['responses'] ) || ! is_array( $body['responses'] ) ) {
			continue;
		}

		foreach ( $body['responses'] as $sub ) {
			$rid = isset( $sub['id'] ) ? (string) $sub['id'] : '';
			if ( '' === $rid || ! isset( $id_map[ $rid ] ) ) {
				continue;
			}

			$status = isset( $sub['status'] ) ? (int) $sub['status'] : 0;
			$rec    = isset( $sub['body']['recurrence'] ) ? $sub['body']['recurrence'] : null;
			if ( 200 !== $status || ! is_array( $rec ) ) {
				continue;
			}

			$text = ms365cal_format_recurrence( $rec );
			if ( '' !== $text ) {
				$map[ $id_map[ $rid ] ] = $text;
			}
		}
	}

	return $map;
}

/**
 * Fetch every configured calendar's first calendarView page together via Graph's
 * $batch endpoint (up to 20 sub-requests per call) instead of one sequential
 * request per calendar — the dominant cost on a cache miss (measured: ~5.4s for 10
 * calendars sequentially, vs. this bringing it down to close to a single Graph
 * round-trip). A calendar whose page has more than $top (settings['events_top'],
 * default 100) events in the window — has @odata.nextLink — is rare for a typical
 * weekly window, especially once $top is raised past a calendar's normal event
 * count; when it happens, that one calendar falls back to the existing sequential
 * ms365cal_fetch_one(), which resolves its own pagination and recurrence, so its
 * rows are excluded from the shared recurrence-consolidation pass below (they're
 * already resolved).
 *
 * @return array{all:array,throttled:bool,retry_after:int}
 */
function ms365cal_fetch_calendars_batched( $wanted, $token, $start_iso, $end_iso, $tz, $top, $show_recur, $lazy_body ) {
	try {
		$zone = new DateTimeZone( $tz );
	} catch ( Exception $e ) {
		$zone = new DateTimeZone( 'UTC' );
	}
	$time_fmt = get_option( 'time_format', 'H:i' );

	try {
		$window_start = new DateTime( $start_iso, $zone );
	} catch ( Exception $e ) {
		$window_start = new DateTime( 'now', $zone );
	}
	try {
		$window_end = new DateTime( $end_iso, $zone );
	} catch ( Exception $e ) {
		$window_end = clone $window_start;
	}
	$today0 = new DateTime( 'now', $zone );
	$today0->setTime( 0, 0, 0 );

	$cals_by_slug = array();
	$url_by_slug  = array();
	foreach ( $wanted as $cal ) {
		$cals_by_slug[ $cal['slug'] ] = $cal;
		$url_by_slug[ $cal['slug'] ]  = ms365cal_view_url( $cal, $start_iso, $end_iso, $top, $lazy_body );
	}

	$prefer       = 'outlook.timezone="' . $tz . '", outlook.body-content-type="html"';
	$events       = array();
	$stragglers   = array(); // slugs needing the full sequential fallback.
	$need_masters = array(); // list of [slug, master_id] across the batched calendars.

	foreach ( array_chunk( $url_by_slug, 20, true ) as $chunk ) {
		$requests = array();
		$id_map   = array(); // batch sub-request id => slug.
		$i        = 0;
		foreach ( $chunk as $slug => $view_url ) {
			$rid            = (string) $i;
			$id_map[ $rid ] = $slug;
			$requests[]     = array(
				'id'      => $rid,
				'method'  => 'GET',
				'url'     => ms365cal_graph_relative_url( $view_url ),
				'headers' => array( 'Prefer' => $prefer ),
			);
			++$i;
		}

		$resp = wp_remote_post(
			'https://graph.microsoft.com/v1.0/$batch',
			array(
				'timeout' => 25,
				'headers' => array(
					'Authorization' => 'Bearer ' . $token,
					'Content-Type'  => 'application/json',
					'Accept'        => 'application/json',
				),
				'body'    => wp_json_encode( array( 'requests' => $requests ) ),
			)
		);

		if ( is_wp_error( $resp ) ) {
			// Whole chunk unreachable — fall back to the sequential fetch per
			// calendar rather than silently dropping them.
			foreach ( $id_map as $slug ) {
				$stragglers[] = $slug;
			}
			continue;
		}

		$code = (int) wp_remote_retrieve_response_code( $resp );
		if ( 429 === $code || 503 === $code ) {
			$ra = (int) wp_remote_retrieve_header( $resp, 'retry-after' );
			return array(
				'all'         => array(),
				'throttled'   => true,
				'retry_after' => $ra > 0 ? $ra : 30,
			);
		}
		if ( 200 !== $code ) {
			foreach ( $id_map as $slug ) {
				$stragglers[] = $slug;
			}
			continue;
		}

		$batch_body = json_decode( wp_remote_retrieve_body( $resp ), true );
		if ( empty( $batch_body['responses'] ) || ! is_array( $batch_body['responses'] ) ) {
			foreach ( $id_map as $slug ) {
				$stragglers[] = $slug;
			}
			continue;
		}

		foreach ( $batch_body['responses'] as $sub ) {
			$rid = isset( $sub['id'] ) ? (string) $sub['id'] : '';
			if ( '' === $rid || ! isset( $id_map[ $rid ] ) ) {
				continue;
			}
			$slug   = $id_map[ $rid ];
			$status = isset( $sub['status'] ) ? (int) $sub['status'] : 0;

			if ( 429 === $status || 503 === $status ) {
				$ra = isset( $sub['headers']['Retry-After'] ) ? (int) $sub['headers']['Retry-After'] : 0;
				return array(
					'all'         => array(),
					'throttled'   => true,
					'retry_after' => $ra > 0 ? $ra : 30,
				);
			}

			if ( 200 !== $status ) {
				continue; // 403/404/etc. for this one calendar — contributes zero events.
			}

			$page_body = isset( $sub['body'] ) && is_array( $sub['body'] ) ? $sub['body'] : array();

			if ( ! empty( $page_body['@odata.nextLink'] ) ) {
				// More than one page for this calendar — rare; fall back to the
				// existing sequential fetch (it resolves its own recurrence too)
				// rather than juggling batched pagination for an edge case.
				$stragglers[] = $slug;
				continue;
			}

			if ( ! empty( $page_body['value'] ) && is_array( $page_body['value'] ) ) {
				$cal_need_masters = array();
				$rows             = ms365cal_build_rows_from_page(
					$cals_by_slug[ $slug ],
					$page_body['value'],
					$zone,
					$time_fmt,
					$window_start,
					$today0,
					$window_end,
					$cal_need_masters,
					$show_recur,
					$lazy_body
				);
				$events           = array_merge( $events, $rows );
				foreach ( array_keys( $cal_need_masters ) as $mid ) {
					$need_masters[] = array( $slug, $mid );
				}
			}
		}
	}

	// Resolve recurrence for the batched calendars' collected series masters,
	// together, in as few $batch calls as possible. $need_masters is always empty
	// when $show_recur is off (ms365cal_build_rows_from_page() never populates it),
	// so this block — and its $batch calls — is a natural no-op in that case.
	if ( $show_recur && ! empty( $need_masters ) ) {
		$recur_map = ms365cal_fetch_recurrence_map_multi( $cals_by_slug, $need_masters, $token );
		foreach ( $events as $idx => $row ) {
			if ( isset( $row['_master'] ) ) {
				$key = $row['cal'] . '|' . $row['_master'];
				if ( isset( $recur_map[ $key ] ) ) {
					$events[ $idx ]['recur'] = $recur_map[ $key ];
				}
				unset( $events[ $idx ]['_master'] );
			}
		}
	}

	// Rare stragglers (busy calendars needing pagination, or a chunk-level
	// transport/HTTP failure): fetch them the old way, one at a time. Their own
	// recurrence is already resolved internally by ms365cal_fetch_one().
	foreach ( array_unique( $stragglers ) as $slug ) {
		$res = ms365cal_fetch_one( $cals_by_slug[ $slug ], $token, $start_iso, $end_iso, $tz, $top, $show_recur, $lazy_body );
		if ( ! empty( $res['throttled'] ) ) {
			return array(
				'all'         => array(),
				'throttled'   => true,
				'retry_after' => (int) $res['retry_after'],
			);
		}
		$events = array_merge( $events, $res['events'] );
	}

	return array(
		'all'         => $events,
		'throttled'   => false,
		'retry_after' => 0,
	);
}

/**
 * Merge + sort + cache a window of events for a set of calendar slugs.
 *
 * $force_live skips both cache short-circuits below (fresh hit and
 * stale-while-revalidate) and goes straight to the live-fetch path — used only
 * by ms365cal_rest_refresh(), the background-refresh trigger fired from the
 * stale-while-revalidate branch itself, so that call actually performs the
 * live fetch instead of recursing into "serve stale, trigger a refresh" again.
 *
 * @return array|WP_Error
 */
function ms365cal_events_window( $slugs, $start_date, $days, $force_live = false ) {
	$settings   = ms365cal_get_settings();
	$tz         = $settings['timezone'] ? $settings['timezone'] : 'UTC';
	$days       = min( MS365CAL_MAX_WINDOW, max( 1, (int) $days ) );
	$top        = max( 10, min( 999, (int) $settings['events_top'] ) );
	$show_recur = ! empty( $settings['show_recurrence'] );
	$lazy_body  = ! empty( $settings['lazy_body'] );

	$wanted = array_values(
		array_filter(
			$settings['calendars'],
			function ( $c ) use ( $slugs ) {
				return in_array( $c['slug'], $slugs, true );
			}
		)
	);
	if ( empty( $wanted ) ) {
		return array();
	}

	// Normalise the window to whole days in site tz.
	try {
		$zone  = new DateTimeZone( $tz );
		$start = DateTime::createFromFormat( 'Y-m-d', $start_date, $zone );
		if ( ! $start ) {
			$start = new DateTime( 'now', $zone );
		}
		$start->setTime( 0, 0, 0 );
		$end = ( clone $start )->modify( '+' . $days . ' days' );
	} catch ( Exception $e ) {
		return new WP_Error( 'ms365cal_tz', 'Invalid timezone.' );
	}

	$start_iso = $start->format( 'Y-m-d\TH:i:s' );
	$end_iso   = $end->format( 'Y-m-d\TH:i:s' );

	$cache_key = 'ms365cal_w_' . md5( implode( ',', $slugs ) . '|' . $start_iso . '|' . $days . '|' . $tz );

	$now       = time();
	$fresh     = max( 1, (int) $settings['cache_minutes'] ) * MINUTE_IN_SECONDS;
	$cached    = get_transient( $cache_key ); // stored payload with fetched-time + events, or false
	$has_stale = is_array( $cached ) && isset( $cached['events'] );

	// Fresh cache hit.
	if ( ! $force_live && $has_stale && ( $now - (int) $cached['fetched'] ) < $fresh ) {
		return $cached['events'];
	}

	// Stale-while-revalidate: expired, but not by more than the grace window
	// (settings['cache_grace_minutes']). Serve the old data immediately — the
	// visitor never waits on a live Graph call for this — and fire a
	// non-blocking background request to refresh it, so the next request finds
	// a warm cache. Self-sufficient: any real visitor landing here while
	// stale-but-in-grace is what triggers the refresh — no external scheduled
	// trigger needed, just occasional organic traffic within the grace window.
	$grace = max( 0, (int) $settings['cache_grace_minutes'] ) * MINUTE_IN_SECONDS;
	if ( ! $force_live && $has_stale && $grace > 0 && ( $now - (int) $cached['fetched'] ) < ( $fresh + $grace ) ) {
		ms365cal_trigger_background_refresh( $slugs, $start_date, $days );
		return $cached['events'];
	}

	// Graph asked us to back off recently — don't call it again yet. Serve stale
	// data if we have any; otherwise report the throttle so the client can wait.
	$backoff_until = (int) get_transient( 'ms365cal_backoff' );
	if ( $backoff_until > $now ) {
		if ( $has_stale ) {
			return $cached['events'];
		}
		return new WP_Error( 'ms365cal_throttled', 'The calendar service is busy.', array( 'retry_after' => $backoff_until - $now ) );
	}

	$token = ms365cal_get_token();
	if ( is_wp_error( $token ) ) {
		return $has_stale ? $cached['events'] : $token;
	}

	$res = ms365cal_fetch_calendars_batched( $wanted, $token, $start_iso, $end_iso, $tz, $top, $show_recur, $lazy_body );

	if ( ! empty( $res['throttled'] ) ) {
		$backoff = max( 1, min( (int) $res['retry_after'], MS365CAL_BACKOFF_MAX ) );
		set_transient( 'ms365cal_backoff', $now + $backoff, $backoff );
		if ( $has_stale ) {
			return $cached['events']; // seamless: slightly old data beats a blank calendar
		}
		return new WP_Error( 'ms365cal_throttled', 'The calendar service is busy.', array( 'retry_after' => $backoff ) );
	}

	$all = $res['all'];

	usort(
		$all,
		function ( $a, $b ) {
			return strcmp( $a['sort'], $b['sort'] );
		}
	);

	// Keep well past the freshness window so stale data stays available to serve
	// during a throttle or outage.
	$keep = max( $fresh, DAY_IN_SECONDS );
	set_transient(
		$cache_key,
		array(
			'fetched' => $now,
			'events'  => $all,
		),
		$keep
	);
	return $all;
}

/**
 * Fire a non-blocking loopback request to /refresh so a stale-but-in-grace
 * window (see ms365cal_events_window()) gets refreshed in the background
 * instead of making the current visitor wait on a live Graph call — the same
 * fire-and-forget technique WordPress's own spawn_cron() uses for wp-cron.php.
 * A short lock transient stops concurrent visitors hitting the same stale
 * window from each firing a duplicate refresh; it isn't released early on
 * success since the lock's own short TTL (well under how long a real refresh
 * takes) is enough to prevent a pile-up without needing that.
 */
function ms365cal_trigger_background_refresh( $slugs, $start_date, $days ) {
	$lock_key = 'ms365cal_refresh_lock_' . md5( implode( ',', $slugs ) . '|' . $start_date . '|' . $days );
	if ( get_transient( $lock_key ) ) {
		return;
	}
	set_transient( $lock_key, 1, 30 );

	wp_remote_post(
		add_query_arg(
			array(
				'cals'  => implode( ',', $slugs ),
				'start' => $start_date,
				'days'  => $days,
			),
			rest_url( 'ms365cal/v1/refresh' )
		),
		array(
			'timeout'   => 0.01,
			'blocking'  => false,
			'sslverify' => apply_filters( 'https_local_ssl_verify', false ),
		)
	);
}

/**
 * Decode the (unverified) claims from a JWT payload. We only read what our own
 * token already contains — no signature check needed — to show the application
 * roles the token actually carries. Returns an associative array of claims.
 */
function ms365cal_decode_jwt_claims( $jwt ) {
	$parts = explode( '.', (string) $jwt );
	if ( count( $parts ) < 2 ) {
		return array();
	}
	$payload = strtr( $parts[1], '-_', '+/' );
	$pad     = strlen( $payload ) % 4;
	if ( $pad ) {
		$payload .= str_repeat( '=', 4 - $pad );
	}
	// phpcs:ignore WordPress.PHP.DiscouragedPHPFunctions.obfuscation_base64_decode
	$json = base64_decode( $payload );
	$data = json_decode( (string) $json, true );
	return is_array( $data ) ? $data : array();
}

/**
 * Drop the cached Graph token, the back-off flag, all window event caches, and all
 * per-event lazy-body caches so the next fetch starts from a blank slate. The debug
 * endpoint calls this to force a brand-new token after a permissions/consent change,
 * rather than reusing a token that was minted before consent (which carries no roles).
 */
function ms365cal_flush_cache() {
	global $wpdb;

	// Exact-key deletes — these work even under a persistent object cache, so the
	// token concern is always covered.
	delete_transient( MS365CAL_TOKEN_TRANSIENT );
	delete_transient( 'ms365cal_backoff' );

	// Window event caches and per-event lazy-body caches use dynamic keys, so
	// clear them by prefix. Best-effort: on a persistent object cache this DB
	// sweep won't reach cached copies, but the diagnostic reads live anyway and
	// the token above is what matters.
	$like_w_val = $wpdb->esc_like( '_transient_ms365cal_w_' ) . '%';
	$like_w_to  = $wpdb->esc_like( '_transient_timeout_ms365cal_w_' ) . '%';
	$like_b_val = $wpdb->esc_like( '_transient_ms365cal_body_' ) . '%';
	$like_b_to  = $wpdb->esc_like( '_transient_timeout_ms365cal_body_' ) . '%';
	// phpcs:ignore WordPress.DB.DirectDatabaseQuery.DirectQuery, WordPress.DB.DirectDatabaseQuery.NoCaching
	$wpdb->query(
		$wpdb->prepare(
			'DELETE FROM ' . $wpdb->options . ' WHERE option_name LIKE %s OR option_name LIKE %s'
			. ' OR option_name LIKE %s OR option_name LIKE %s',
			$like_w_val,
			$like_w_to,
			$like_b_val,
			$like_b_to
		)
	);
}

/**
 * Admin-only diagnostic: fetch each configured calendar live (bypassing the
 * cache and the back-off) and report the real Graph status, so a 403/404 that
 * the normal path swallows becomes visible. Returns array|WP_Error.
 */
function ms365cal_diagnose( $slugs, $start_date, $days ) {
	$settings   = ms365cal_get_settings();
	$tz         = $settings['timezone'] ? $settings['timezone'] : 'UTC';
	$days       = min( MS365CAL_MAX_WINDOW, max( 1, (int) $days ) );
	$top        = max( 10, min( 999, (int) $settings['events_top'] ) );
	$show_recur = ! empty( $settings['show_recurrence'] );
	$lazy_body  = ! empty( $settings['lazy_body'] );

	$wanted = array_values(
		array_filter(
			$settings['calendars'],
			function ( $c ) use ( $slugs ) {
				return in_array( $c['slug'], $slugs, true );
			}
		)
	);

	try {
		$zone  = new DateTimeZone( $tz );
		$start = DateTime::createFromFormat( 'Y-m-d', $start_date, $zone );
		if ( ! $start ) {
			$start = new DateTime( 'now', $zone );
		}
		$start->setTime( 0, 0, 0 );
		$end = ( clone $start )->modify( '+' . $days . ' days' );
	} catch ( Exception $e ) {
		return new WP_Error( 'ms365cal_tz', 'Invalid timezone.' );
	}

	$start_iso = $start->format( 'Y-m-d\TH:i:s' );
	$end_iso   = $end->format( 'Y-m-d\TH:i:s' );

	$out = array(
		'window'    => array(
			'start' => $start_iso,
			'end'   => $end_iso,
			'tz'    => $tz,
		),
		'token'     => 'ok',
		'calendars' => array(),
	);

	if ( empty( $wanted ) ) {
		$out['token'] = 'no calendars configured';
		return $out;
	}

	$token = ms365cal_get_token();
	if ( is_wp_error( $token ) ) {
		$out['token'] = $token->get_error_message();
		return $out;
	}

	// Surface the application roles the token carries. Empty here == the app has
	// no consented Application permissions, which is the usual cause of a blanket
	// "Access is denied" from Exchange even though a token was issued.
	$claims                = ms365cal_decode_jwt_claims( $token );
	$out['token_roles']    = isset( $claims['roles'] ) ? $claims['roles'] : array();
	$out['token_audience'] = isset( $claims['aud'] ) ? $claims['aud'] : '';

	foreach ( $wanted as $cal ) {
		$res                = ms365cal_fetch_one( $cal, $token, $start_iso, $end_iso, $tz, $top, $show_recur, $lazy_body );
		$out['calendars'][] = array(
			'slug'      => $cal['slug'],
			'type'      => $cal['type'],
			'source'    => $cal['source'],
			'http'      => isset( $res['status'] ) ? (int) $res['status'] : null,
			'events'    => count( $res['events'] ),
			'throttled' => ! empty( $res['throttled'] ),
			'error'     => isset( $res['error'] ) ? $res['error'] : '',
		);
	}

	return $out;
}

/**
 * ---------------------------------------------------------------------------
 *  REST endpoint:  GET /wp-json/ms365cal/v1/events?cals=eng,sales&start=2026-07-17&days=14
 * ---------------------------------------------------------------------------
 *  Public (the host page is public). The browser passes calendar *slugs*,
 *  never Graph identifiers, and only configured slugs are ever queried.
 *
 *  Add &debug=1 as a logged-in admin to get per-calendar Graph status instead.
 */
function ms365cal_register_rest() {
	register_rest_route(
		'ms365cal/v1',
		'/events',
		array(
			'methods'             => 'GET',
			'permission_callback' => '__return_true',
			'args'                => array(
				'cals'  => array( 'sanitize_callback' => 'sanitize_text_field' ),
				'start' => array( 'sanitize_callback' => 'sanitize_text_field' ),
				'days'  => array( 'sanitize_callback' => 'absint' ),
				'debug' => array( 'sanitize_callback' => 'absint' ),
			),
			'callback'            => 'ms365cal_rest_events',
		)
	);

	// On-demand single-event body (only registered/used when 'lazy_body' is on — see
	// ms365cal_view_url()). 'cal' is a whitelisted slug, same as /events; 'id' is the
	// raw Graph event id from that event's row, needed to fetch just this one event.
	register_rest_route(
		'ms365cal/v1',
		'/event-body',
		array(
			'methods'             => 'GET',
			'permission_callback' => '__return_true',
			'args'                => array(
				'cal' => array( 'sanitize_callback' => 'sanitize_title' ),
				'id'  => array( 'sanitize_callback' => 'sanitize_text_field' ),
			),
			'callback'            => 'ms365cal_rest_event_body',
		)
	);

	// Background-refresh trigger for the stale-while-revalidate path (see
	// ms365cal_events_window() / ms365cal_trigger_background_refresh()). Not
	// meant to be called directly by a visitor — the plugin fires this at
	// itself as a non-blocking loopback request. Public (like /events) since it
	// performs the exact same fetch-and-cache work a real visitor would trigger
	// anyway, just rate-limited the same way to bound misuse.
	register_rest_route(
		'ms365cal/v1',
		'/refresh',
		array(
			'methods'             => 'POST',
			'permission_callback' => '__return_true',
			'args'                => array(
				'cals'  => array( 'sanitize_callback' => 'sanitize_text_field' ),
				'start' => array( 'sanitize_callback' => 'sanitize_text_field' ),
				'days'  => array( 'sanitize_callback' => 'absint' ),
			),
			'callback'            => 'ms365cal_rest_refresh',
		)
	);

	// Deploy hook: force this plugin to pull + install its latest GitHub release.
	// Disabled unless MS365CAL_DEPLOY_KEY is defined in wp-config; auth is the key
	// (timing-safe compare), so no WP login is required. See ms365cal_rest_self_update().
	register_rest_route(
		'ms365cal/v1',
		'/self-update',
		array(
			'methods'             => 'POST',
			'permission_callback' => '__return_true',
			'callback'            => 'ms365cal_rest_self_update',
		)
	);

	// Admin-only: groups + rooms + mail-enabled users from the tenant, for the
	// "type to search" source picker on the Calendars tab. See ms365cal_fetch_directory().
	register_rest_route(
		'ms365cal/v1',
		'/directory',
		array(
			'methods'             => 'GET',
			'permission_callback' => function () {
				return current_user_can( 'manage_options' );
			},
			'args'                => array(
				'refresh' => array( 'sanitize_callback' => 'absint' ),
			),
			'callback'            => 'ms365cal_rest_directory',
		)
	);
}
add_action( 'rest_api_init', 'ms365cal_register_rest' );

/**
 * Admin-only directory picker: lists Graph GET responses as a flat page of
 * {value:[...], '@odata.nextLink':...}, following nextLink up to
 * MS365CAL_MAX_PAGES pages (same safety cap used for calendarView pagination).
 * Returns whatever was collected plus an error string if a page failed, rather
 * than discarding partial results — the directory picker degrades gracefully
 * (one section missing) instead of failing outright over one bad resource.
 */
function ms365cal_graph_collect( $url, $token ) {
	$items = array();
	$pages = 0;
	while ( $url && $pages < MS365CAL_MAX_PAGES ) {
		$resp = wp_remote_get(
			$url,
			array(
				'timeout' => 20,
				'headers' => array( 'Authorization' => 'Bearer ' . $token ),
			)
		);
		if ( is_wp_error( $resp ) ) {
			return array(
				'items' => $items,
				'error' => $resp->get_error_message(),
			);
		}
		$code = (int) wp_remote_retrieve_response_code( $resp );
		$body = json_decode( wp_remote_retrieve_body( $resp ), true );
		if ( 200 !== $code ) {
			$msg = isset( $body['error']['message'] ) ? $body['error']['message'] : ( 'HTTP ' . $code );
			return array(
				'items' => $items,
				'error' => $msg,
			);
		}
		if ( ! empty( $body['value'] ) && is_array( $body['value'] ) ) {
			$items = array_merge( $items, $body['value'] );
		}
		$url = isset( $body['@odata.nextLink'] ) ? $body['@odata.nextLink'] : '';
		++$pages;
	}
	return array(
		'items' => $items,
		'error' => null,
	);
}

/**
 * Fetches the tenant's M365 groups, room mailboxes, and mail-enabled user
 * accounts for the "type to search" source picker — a flat, filterable list
 * instead of making the admin hand-type a GUID or email. Rooms and users are
 * both stored as 'mailbox' sources (same as ms365cal_detect_calendar_type()
 * would classify them from their address alone); 'kind' here is display-only,
 * so the picker can group results.
 *
 * There is deliberately no attempt to distinguish "shared" mailboxes from
 * regular user mailboxes — Graph's core /users endpoint has no reliable field
 * for that (see CLAUDE.md), so every mail-enabled Member account is listed
 * and the admin recognises the one they want by name.
 *
 * Cached for 15 minutes (directories don't change minute to minute); pass
 * $force_refresh to bypass it. Best-effort: a failure on one resource (e.g.
 * missing Place.Read.All) doesn't block the other two, it's just reported in
 * 'errors'.
 */
function ms365cal_fetch_directory( $force_refresh = false ) {
	$transient_key = 'ms365cal_directory';
	if ( ! $force_refresh ) {
		$cached = get_transient( $transient_key );
		if ( false !== $cached ) {
			return $cached;
		}
	}

	$token = ms365cal_get_token();
	if ( is_wp_error( $token ) ) {
		return array(
			'items'   => array(),
			'errors'  => array( $token->get_error_message() ),
			'fetched' => time(),
		);
	}

	$group_filter = rawurlencode( "groupTypes/any(c:c eq 'Unified')" );
	$resources    = array(
		'group'   => "https://graph.microsoft.com/v1.0/groups?\$filter={$group_filter}&\$select=id,displayName&\$top=999",
		'room'    => 'https://graph.microsoft.com/v1.0/places/microsoft.graph.room?$top=999',
		'mailbox' => 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userType&$top=999',
	);

	$items  = array();
	$errors = array();

	foreach ( $resources as $kind => $url ) {
		$res = ms365cal_graph_collect( $url, $token );
		if ( $res['error'] ) {
			$errors[] = $kind . ': ' . $res['error'];
		}
		foreach ( $res['items'] as $row ) {
			if ( 'group' === $kind ) {
				if ( empty( $row['id'] ) ) {
					continue;
				}
				$items[] = array(
					'kind'   => 'group',
					'source' => $row['id'],
					'label'  => $row['displayName'] ?? $row['id'],
				);
			} elseif ( 'room' === $kind ) {
				if ( empty( $row['emailAddress'] ) ) {
					continue;
				}
				$items[] = array(
					'kind'   => 'room',
					'source' => $row['emailAddress'],
					'label'  => $row['displayName'] ?? $row['emailAddress'],
				);
			} else {
				// Every mail-enabled Member account, not just "shared" ones — see docblock above.
				if ( empty( $row['mail'] ) || 'Member' !== ( $row['userType'] ?? '' ) ) {
					continue;
				}
				$items[] = array(
					'kind'   => 'mailbox',
					'source' => $row['mail'],
					'label'  => $row['displayName'] ?? $row['mail'],
				);
			}
		}
	}

	usort(
		$items,
		function ( $a, $b ) {
			return strcasecmp( $a['label'], $b['label'] );
		}
	);

	$result = array(
		'items'   => $items,
		'errors'  => $errors,
		'fetched' => time(),
	);
	set_transient( $transient_key, $result, 15 * MINUTE_IN_SECONDS );
	return $result;
}

function ms365cal_rest_directory( WP_REST_Request $req ) {
	if ( ! current_user_can( 'manage_options' ) ) {
		return new WP_REST_Response( array( 'error' => 'forbidden' ), 403 );
	}
	$data = ms365cal_fetch_directory( (bool) $req->get_param( 'refresh' ) );
	return new WP_REST_Response( $data, 200 );
}

/**
 * Secret-authenticated self-update. Forces the update checker to re-check GitHub and,
 * if a newer release exists, installs it via the plugin upgrader — so a release can be
 * pushed to a live site without shell access or a wp-admin click.
 *
 * Guardrails: off unless MS365CAL_DEPLOY_KEY is defined; key sent in the
 * X-MS365CAL-Deploy-Key header and compared with hash_equals(); a short transient lock
 * blocks concurrent/rapid triggers; and the update source is whatever repo
 * 'update_repo' / MS365CAL_UPDATE_REPO names (see ms365cal_init_updates()), so the
 * endpoint can only ever install that repo's latest release — never arbitrary code.
 * No repo configured means no update checker was built at all: this returns
 * 'updater_unavailable' (500) rather than silently falling back to anything.
 */
function ms365cal_rest_self_update( WP_REST_Request $req ) {
	// Effective key: wp-config constant MS365CAL_DEPLOY_KEY wins, else the DB setting
	// (Settings → MS365 Calendar). Empty = endpoint disabled.
	$key = (string) ms365cal_cred( 'deploy_key' );
	if ( '' === $key ) {
		return new WP_REST_Response( array( 'error' => 'not_found' ), 404 );
	}

	$provided = (string) $req->get_header( 'X-MS365CAL-Deploy-Key' );
	if ( '' === $provided || ! hash_equals( $key, $provided ) ) {
		return new WP_REST_Response( array( 'error' => 'forbidden' ), 403 );
	}

	if ( get_transient( 'ms365cal_selfupdate_lock' ) ) {
		$resp = new WP_REST_Response( array( 'error' => 'busy' ), 429 );
		$resp->header( 'Retry-After', '60' );
		return $resp;
	}
	set_transient( 'ms365cal_selfupdate_lock', 1, MINUTE_IN_SECONDS );

	$plugin  = plugin_basename( __FILE__ );
	$headers = get_file_data( __FILE__, array( 'Version' => 'Version' ) );
	$before  = isset( $headers['Version'] ) ? $headers['Version'] : '';

	$checker = ms365cal_update_checker();
	if ( ! $checker ) {
		delete_transient( 'ms365cal_selfupdate_lock' );
		return new WP_REST_Response( array( 'error' => 'updater_unavailable' ), 500 );
	}

	// Force a fresh check so the update_plugins transient carries the newest release.
	$checker->checkForUpdates();

	require_once ABSPATH . 'wp-admin/includes/file.php';
	require_once ABSPATH . 'wp-admin/includes/misc.php';
	require_once ABSPATH . 'wp-admin/includes/plugin.php';
	require_once ABSPATH . 'wp-admin/includes/class-wp-upgrader.php';

	$was_active = is_plugin_active( $plugin );

	$skin     = new Automatic_Upgrader_Skin();
	$upgrader = new Plugin_Upgrader( $skin );
	$result   = $upgrader->upgrade( $plugin );

	// Plugin_Upgrader::upgrade() silently deactivates an active plugin before the
	// swap and only skips that during cron (wp_doing_cron()). We run in a REST
	// request, so nothing reactivates it — restore the prior active state ourselves,
	// or a self-update would knock the plugin (and this endpoint) offline.
	$reactivated = false;
	if ( $was_active && ! is_plugin_active( $plugin ) ) {
		activate_plugin( $plugin, '', false, true );
		$reactivated = true;
	}

	$after = get_file_data( __FILE__, array( 'Version' => 'Version' ) );
	$after = isset( $after['Version'] ) ? $after['Version'] : '';

	delete_transient( 'ms365cal_selfupdate_lock' );

	// A new version may change how events are computed/rendered, so drop the cached
	// window payloads (and token) — otherwise stale pre-update data is served until
	// each window's cache naturally expires. Only when a version actually changed.
	if ( $after !== $before ) {
		ms365cal_flush_cache();
	}

	if ( is_wp_error( $result ) ) {
		return new WP_REST_Response(
			array(
				'error'   => 'upgrade_failed',
				'message' => $result->get_error_message(),
				'from'    => $before,
			),
			500
		);
	}

	return new WP_REST_Response(
		array(
			'updated'     => ( $after !== $before ),
			'from'        => $before,
			'to'          => $after,
			'reactivated' => $reactivated,
			'messages'    => $skin->get_upgrade_messages(),
		),
		200
	);
}

/**
 * Best-effort client IP. Uses REMOTE_ADDR only — X-Forwarded-For is spoofable,
 * which would let a caller bypass the limit. If your site sits behind a trusted
 * reverse proxy/CDN, resolve the real client IP into REMOTE_ADDR at that layer
 * (or filter 'ms365cal_client_ip').
 */
function ms365cal_client_ip() {
	$ip = isset( $_SERVER['REMOTE_ADDR'] ) ? trim( wp_unslash( $_SERVER['REMOTE_ADDR'] ) ) : '';
	if ( ! filter_var( $ip, FILTER_VALIDATE_IP ) ) {
		$ip = 'unknown';
	}
	return apply_filters( 'ms365cal_client_ip', $ip );
}

/**
 * Effective rate-limit values. Precedence: filter overrides the saved admin
 * setting, which defaults to the MS365CAL_RATE_* constants on a fresh install.
 *
 * @return array{max:int,window:int}
 */
function ms365cal_rate_limits() {
	$s = ms365cal_get_settings();
	return array(
		'max'    => (int) apply_filters( 'ms365cal_rate_max', (int) $s['rate_max'] ),
		'window' => (int) apply_filters( 'ms365cal_rate_window', (int) $s['rate_window'] ),
	);
}

/**
 * Fixed-window per-IP throttle backed by a transient.
 *
 * @return bool True if the request is allowed, false if over the limit.
 */
function ms365cal_rate_check() {
	$limits = ms365cal_rate_limits();
	$max    = $limits['max'];
	$window = max( 1, $limits['window'] );
	if ( $max <= 0 ) {
		return true; // limiter disabled
	}

	$key   = 'ms365cal_rl_' . md5( ms365cal_client_ip() );
	$count = (int) get_transient( $key );

	if ( $count >= $max ) {
		return false;
	}

	// Sliding window: each allowed hit refreshes the TTL, so the caller stays
	// counted until they pause for a full window. Normal browsing (a handful of
	// clicks) never approaches the limit.
	set_transient( $key, $count + 1, $window );
	return true;
}

function ms365cal_rest_events( WP_REST_Request $req ) {
	if ( ! ms365cal_rate_check() ) {
		$limits = ms365cal_rate_limits();
		$resp   = new WP_REST_Response( array( 'error' => 'rate_limited' ), 429 );
		$resp->header( 'Retry-After', (string) max( 1, $limits['window'] ) );
		return $resp;
	}

	$settings   = ms365cal_get_settings();
	$configured = wp_list_pluck( $settings['calendars'], 'slug' );

	$req_slugs = array_filter( array_map( 'sanitize_title', explode( ',', (string) $req->get_param( 'cals' ) ) ) );
	$slugs     = array_values( array_intersect( $req_slugs, $configured ) );
	if ( empty( $slugs ) ) {
		$slugs = $configured; // default to all
	}

	$start = (string) $req->get_param( 'start' );
	if ( ! preg_match( '/^\d{4}-\d{2}-\d{2}$/', $start ) ) {
		$start = wp_date( 'Y-m-d' );
	}

	$days = $req->get_param( 'days' ) ? (int) $req->get_param( 'days' ) : 7;

	// Admin-only diagnostic. Non-admins passing debug=1 fall through to normal output.
	if ( $req->get_param( 'debug' ) && current_user_can( 'manage_options' ) ) {
		ms365cal_flush_cache(); // start from a blank slate: fresh token + fresh fetch
		$diag = ms365cal_diagnose( $slugs, $start, $days );
		if ( is_wp_error( $diag ) ) {
			return new WP_REST_Response( array( 'error' => $diag->get_error_message() ), 200 );
		}
		$diag['flushed'] = true;
		return new WP_REST_Response( array( 'debug' => $diag ), 200 );
	}

	$events = ms365cal_events_window( $slugs, $start, $days );
	if ( is_wp_error( $events ) ) {
		if ( 'ms365cal_throttled' === $events->get_error_code() ) {
			$data = $events->get_error_data();
			$ra   = isset( $data['retry_after'] ) ? max( 1, (int) $data['retry_after'] ) : 30;
			$resp = new WP_REST_Response(
				array(
					'error'       => 'upstream_busy',
					'retry_after' => $ra,
				),
				429
			);
			$resp->header( 'Retry-After', (string) $ra );
			return $resp;
		}
		return new WP_REST_Response(
			array( 'error' => current_user_can( 'manage_options' ) ? $events->get_error_message() : 'unavailable' ),
			503
		);
	}

	// Data minimisation: the Outlook webLink only needs to reach the browser when
	// the "Open in Outlook" link is enabled. When it's off, strip it server-side so
	// it never leaves the server (the front end hides it either way, but this keeps
	// the raw link out of the unauthenticated response entirely).
	if ( empty( $settings['show_outlook'] ) ) {
		$events = array_map(
			function ( $ev ) {
				unset( $ev['link'] );
				return $ev;
			},
			$events
		);
	}

	return new WP_REST_Response( array( 'events' => $events ), 200 );
}

/**
 * Background-refresh trigger for the stale-while-revalidate path in
 * ms365cal_events_window() — see ms365cal_trigger_background_refresh() for
 * where this gets called (the plugin firing a non-blocking request at itself,
 * never meant to be hit directly by a visitor). $force_live=true makes the
 * inner call actually perform the live fetch instead of finding the same
 * still-stale cache and triggering yet another background refresh. The
 * response is discarded by the caller (fire-and-forget), so what's returned
 * here doesn't matter much beyond a 200.
 */
function ms365cal_rest_refresh( WP_REST_Request $req ) {
	if ( ! ms365cal_rate_check() ) {
		return new WP_REST_Response( array( 'error' => 'rate_limited' ), 429 );
	}

	$settings   = ms365cal_get_settings();
	$configured = wp_list_pluck( $settings['calendars'], 'slug' );

	$req_slugs = array_filter( array_map( 'sanitize_title', explode( ',', (string) $req->get_param( 'cals' ) ) ) );
	$slugs     = array_values( array_intersect( $req_slugs, $configured ) );
	if ( empty( $slugs ) ) {
		$slugs = $configured;
	}

	$start = (string) $req->get_param( 'start' );
	if ( ! preg_match( '/^\d{4}-\d{2}-\d{2}$/', $start ) ) {
		$start = wp_date( 'Y-m-d' );
	}
	$days = $req->get_param( 'days' ) ? (int) $req->get_param( 'days' ) : 7;

	ms365cal_events_window( $slugs, $start, $days, true );
	return new WP_REST_Response( array( 'ok' => true ), 200 );
}

/**
 * ---------------------------------------------------------------------------
 *  REST endpoint:  GET /wp-json/ms365cal/v1/event-body?cal=eng&id=<graph-event-id>
 * ---------------------------------------------------------------------------
 *  On-demand body for a single event, used when 'lazy_body' is on so the list
 *  fetch (ms365cal_events_window()) never has to fetch or sanitise every event's
 *  body up front — only the one a visitor actually expands. 'cal' is validated
 *  against the configured slug whitelist exactly like /events; 'id' is only ever
 *  used to build a Graph URL scoped to that calendar's own /events collection
 *  (rawurlencode()'d, so it can't be used to escape that path), never trusted for
 *  anything else. Failures are soft: any non-throttle error returns 200 with an
 *  empty body rather than surfacing Graph's error detail to an unauthenticated
 *  caller — the front end just shows the event without a description.
 */
function ms365cal_rest_event_body( WP_REST_Request $req ) {
	if ( ! ms365cal_rate_check() ) {
		$limits = ms365cal_rate_limits();
		$resp   = new WP_REST_Response( array( 'error' => 'rate_limited' ), 429 );
		$resp->header( 'Retry-After', (string) max( 1, $limits['window'] ) );
		return $resp;
	}

	$settings = ms365cal_get_settings();
	$slug     = (string) $req->get_param( 'cal' );
	$event_id = (string) $req->get_param( 'id' );
	if ( '' === $slug || '' === $event_id ) {
		return new WP_REST_Response( array( 'error' => 'bad_request' ), 400 );
	}

	$cal = null;
	foreach ( $settings['calendars'] as $c ) {
		if ( $c['slug'] === $slug ) {
			$cal = $c;
			break;
		}
	}
	if ( null === $cal ) {
		return new WP_REST_Response( array( 'error' => 'unknown_calendar' ), 404 );
	}

	// Same global back-off Graph throttling uses elsewhere: don't call Graph again
	// while it's active, and don't cache anything since we have nothing to cache.
	$now           = time();
	$backoff_until = (int) get_transient( 'ms365cal_backoff' );
	if ( $backoff_until > $now ) {
		$resp = new WP_REST_Response(
			array(
				'error'       => 'upstream_busy',
				'retry_after' => $backoff_until - $now,
			),
			429
		);
		$resp->header( 'Retry-After', (string) ( $backoff_until - $now ) );
		return $resp;
	}

	$cache_key = 'ms365cal_body_' . md5( $slug . '|' . $event_id );
	$cached    = get_transient( $cache_key );
	if ( is_array( $cached ) && isset( $cached['body'] ) ) {
		return new WP_REST_Response( array( 'body' => $cached['body'] ), 200 );
	}

	$token = ms365cal_get_token();
	if ( is_wp_error( $token ) ) {
		return new WP_REST_Response( array( 'body' => '' ), 200 );
	}

	$url = add_query_arg(
		array( '$select' => 'body,onlineMeeting' ),
		'https://graph.microsoft.com/v1.0' . ms365cal_events_rel_base( $cal ) . '/' . rawurlencode( $event_id )
	);

	$resp = wp_remote_get(
		$url,
		array(
			'timeout' => 15,
			'headers' => array(
				'Authorization' => 'Bearer ' . $token,
				'Prefer'        => 'outlook.body-content-type="html"',
			),
		)
	);

	if ( is_wp_error( $resp ) ) {
		return new WP_REST_Response( array( 'body' => '' ), 200 );
	}

	$code = (int) wp_remote_retrieve_response_code( $resp );

	if ( 429 === $code || 503 === $code ) {
		$ra      = (int) wp_remote_retrieve_header( $resp, 'retry-after' );
		$backoff = max( 1, min( $ra > 0 ? $ra : 30, MS365CAL_BACKOFF_MAX ) );
		set_transient( 'ms365cal_backoff', $now + $backoff, $backoff );
		$out = new WP_REST_Response(
			array(
				'error'       => 'upstream_busy',
				'retry_after' => $backoff,
			),
			429
		);
		$out->header( 'Retry-After', (string) $backoff );
		return $out;
	}

	if ( 200 !== $code ) {
		// 403/404/etc — the event may have moved or been deleted since the list
		// fetch. Fail soft rather than surfacing Graph's error to the visitor.
		return new WP_REST_Response( array( 'body' => '' ), 200 );
	}

	$data      = json_decode( wp_remote_retrieve_body( $resp ), true );
	$is_online = ! empty( $data['onlineMeeting'] );
	$body      = isset( $data['body']['content'] ) ? trim( (string) $data['body']['content'] ) : '';
	if ( '' !== $body ) {
		$body = ms365cal_dedupe_repeated_links( $body );
		if ( $is_online ) {
			$body = ms365cal_strip_meeting_boilerplate( $body );
		}
		$body = ms365cal_sanitize_event_html( $body );
	}

	$fresh = max( 1, (int) $settings['cache_minutes'] ) * MINUTE_IN_SECONDS;
	set_transient( $cache_key, array( 'body' => $body ), $fresh );

	return new WP_REST_Response( array( 'body' => $body ), 200 );
}

/**
 * ---------------------------------------------------------------------------
 *  Colour palette
 * ---------------------------------------------------------------------------
 *  14 preset colour sets, replacing free-text hex entry — named hues from
 *  Tailwind CSS's default palette (a widely-used, purpose-built set for
 *  exactly this kind of categorical colour-coding), using each hue's 500/800/
 *  100 shade for primary/text/supplement respectively. Deliberately not the
 *  full Tailwind set: near-neighbour hues (the other four greys besides
 *  slate; yellow next to amber; sky next to cyan; violet next to purple;
 *  rose next to red) are dropped so every remaining option stays visually
 *  distinct at a glance — verified pairwise by RGB distance, worst case
 *  (emerald vs teal) still clearly separable, most pairs far more so.
 *    - primary:    the rail / picker-dot colour (Tailwind's *-500).
 *    - text:       for text sitting on that set's own supplement pill —
 *                  chosen per set for contrast (Tailwind's *-800), not
 *                  hardcoded black.
 *    - supplement: the pill background (calendar-picker chip, and the
 *                  category-label pill in the event list) — Tailwind's
 *                  *-100.
 */
function ms365cal_color_palette() {
	return array(
		array(
			'name'       => 'Slate',
			'primary'    => '#64748B',
			'text'       => '#1E293B',
			'supplement' => '#F1F5F9',
		),
		array(
			'name'       => 'Red',
			'primary'    => '#EF4444',
			'text'       => '#991B1B',
			'supplement' => '#FEE2E2',
		),
		array(
			'name'       => 'Orange',
			'primary'    => '#F97316',
			'text'       => '#9A3412',
			'supplement' => '#FFEDD5',
		),
		array(
			'name'       => 'Amber',
			'primary'    => '#F59E0B',
			'text'       => '#92400E',
			'supplement' => '#FEF3C7',
		),
		array(
			'name'       => 'Lime',
			'primary'    => '#84CC16',
			'text'       => '#3F6212',
			'supplement' => '#ECFCCB',
		),
		array(
			'name'       => 'Green',
			'primary'    => '#22C55E',
			'text'       => '#166534',
			'supplement' => '#DCFCE7',
		),
		array(
			'name'       => 'Emerald',
			'primary'    => '#10B981',
			'text'       => '#065F46',
			'supplement' => '#D1FAE5',
		),
		array(
			'name'       => 'Teal',
			'primary'    => '#14B8A6',
			'text'       => '#115E59',
			'supplement' => '#CCFBF1',
		),
		array(
			'name'       => 'Cyan',
			'primary'    => '#06B6D4',
			'text'       => '#155E75',
			'supplement' => '#CFFAFE',
		),
		array(
			'name'       => 'Blue',
			'primary'    => '#3B82F6',
			'text'       => '#1E40AF',
			'supplement' => '#DBEAFE',
		),
		array(
			'name'       => 'Indigo',
			'primary'    => '#6366F1',
			'text'       => '#3730A3',
			'supplement' => '#E0E7FF',
		),
		array(
			'name'       => 'Purple',
			'primary'    => '#A855F7',
			'text'       => '#6B21A8',
			'supplement' => '#F3E8FF',
		),
		array(
			'name'       => 'Fuchsia',
			'primary'    => '#D946EF',
			'text'       => '#86198F',
			'supplement' => '#FAE8FF',
		),
		array(
			'name'       => 'Pink',
			'primary'    => '#EC4899',
			'text'       => '#9D174D',
			'supplement' => '#FCE7F3',
		),
	);
}

/**
 * Resolve a calendar's stored primary hex to its full {primary, text,
 * supplement} set. Falls back to the closest palette entry by RGB distance
 * for anything not an exact match (e.g. a colour saved before this palette
 * existed) so legacy values still render with a sensible, readable pairing
 * instead of breaking.
 */
function ms365cal_color_set( $hex ) {
	$palette = ms365cal_color_palette();
	foreach ( $palette as $set ) {
		if ( 0 === strcasecmp( $set['primary'], (string) $hex ) ) {
			return $set;
		}
	}
	return ms365cal_nearest_color_set( $hex, $palette );
}

/**
 * Nearest palette entry to an arbitrary hex colour by squared RGB distance —
 * used both for legacy-colour front-end rendering (via ms365cal_color_set())
 * and to pre-select a sensible swatch on the admin settings page for a
 * calendar whose stored colour predates this palette.
 */
function ms365cal_nearest_color_set( $hex, $palette = null ) {
	if ( null === $palette ) {
		$palette = ms365cal_color_palette();
	}
	$hex = ltrim( strtoupper( (string) $hex ), '#' );
	if ( 3 === strlen( $hex ) ) {
		$hex = $hex[0] . $hex[0] . $hex[1] . $hex[1] . $hex[2] . $hex[2];
	}
	if ( 6 !== strlen( $hex ) || ! ctype_xdigit( $hex ) ) {
		return $palette[0];
	}
	$r1 = hexdec( substr( $hex, 0, 2 ) );
	$g1 = hexdec( substr( $hex, 2, 2 ) );
	$b1 = hexdec( substr( $hex, 4, 2 ) );

	$best   = $palette[0];
	$best_d = PHP_INT_MAX;
	foreach ( $palette as $set ) {
		$p2 = ltrim( $set['primary'], '#' );
		$r2 = hexdec( substr( $p2, 0, 2 ) );
		$g2 = hexdec( substr( $p2, 2, 2 ) );
		$b2 = hexdec( substr( $p2, 4, 2 ) );
		$d  = ( $r1 - $r2 ) * ( $r1 - $r2 ) + ( $g1 - $g2 ) * ( $g1 - $g2 ) + ( $b1 - $b2 ) * ( $b1 - $b2 );
		if ( $d < $best_d ) {
			$best_d = $d;
			$best   = $set;
		}
	}
	return $best;
}

/**
 * ---------------------------------------------------------------------------
 *  Shortcode — renders the shell only. Events load async via the REST route.
 *  [ms365_calendar calendars="eng,sales" enabled="eng"]  (weekly, Monday-aligned)
 * ---------------------------------------------------------------------------
 */
function ms365cal_shortcode( $atts ) {
	$settings = ms365cal_get_settings();
	$all_cals = $settings['calendars'];

	if ( empty( $all_cals ) ) {
		return '<p>Inga kalendrar har konfigurerats än. Lägg till några under Inställningar &rarr; MS365 Calendar.</p>';
	}

	$atts = shortcode_atts(
		array(
			'calendars' => '',
			'enabled'   => '',
			'days'      => 7,
		),
		$atts,
		'ms365_calendar'
	);

	if ( '' === trim( $atts['calendars'] ) ) {
		$slugs = wp_list_pluck( $all_cals, 'slug' );
	} else {
		$slugs = array_map( 'trim', explode( ',', $atts['calendars'] ) );
	}

	$scoped = array_values(
		array_filter(
			$all_cals,
			function ( $c ) use ( $slugs ) {
				return in_array( $c['slug'], $slugs, true );
			}
		)
	);
	if ( empty( $scoped ) ) {
		return '<p>Ingen av de begärda kalendrarna finns.</p>';
	}

	// All calendars are on by default; 'enabled' scopes that down to a subset for
	// this particular embed (the rest are still shown as chips, just unchecked
	// until clicked) — e.g. [ms365_calendar enabled="eng,events"]. There's no
	// global per-calendar "on by default" setting any more, only this per-embed
	// override.
	$enabled_slugs = '' === trim( $atts['enabled'] ) ? null : array_map( 'trim', explode( ',', $atts['enabled'] ) );

	$window = min( MS365CAL_MAX_WINDOW, max( 1, (int) $atts['days'] ) );
	$tz     = $settings['timezone'] ? $settings['timezone'] : 'UTC';

	// Today, in the site timezone, as the initial window start.
	try {
		$today = new DateTime( 'now', new DateTimeZone( $tz ) );
	} catch ( Exception $e ) {
		$today = new DateTime( 'now' );
	}

	$meta     = array();
	$defaults = array();
	foreach ( $scoped as $c ) {
		$set                    = ms365cal_color_set( $c['color'] );
		$meta[ $c['slug'] ]     = array(
			'label'      => $c['label'],
			'primary'    => $set['primary'],
			'text'       => $set['text'],
			'supplement' => $set['supplement'],
		);
		$defaults[ $c['slug'] ] = null === $enabled_slugs || in_array( $c['slug'], $enabled_slugs, true );
	}

	$config = array(
		'rest'        => esc_url_raw( rest_url( 'ms365cal/v1/events' ) ),
		'bodyRest'    => esc_url_raw( rest_url( 'ms365cal/v1/event-body' ) ),
		'cals'        => wp_list_pluck( $scoped, 'slug' ),
		'meta'        => $meta,
		'defaults'    => $defaults,
		'windowDays'  => $window,
		'startY'      => (int) $today->format( 'Y' ),
		'startM'      => (int) $today->format( 'n' ),
		'startD'      => (int) $today->format( 'j' ),
		'showOutlook' => ! empty( $settings['show_outlook'] ),
		'lazyBody'    => ! empty( $settings['lazy_body'] ),
		'tlInterval'  => (int) $settings['timeline_interval'],
	);

	$uid = 'ms365cal-' . wp_generate_password( 6, false, false );

	ob_start();
	?>
	<div class="ms365cal" id="<?php echo esc_attr( $uid ); ?>" data-config='<?php echo esc_attr( wp_json_encode( $config ) ); ?>'>
		<div class="ms365cal-bar">
			<span class="ms365cal-title">Kalendrar</span>
			<span>
				<button type="button" class="ms365cal-act" data-act="all">Välj alla</button>
				<button type="button" class="ms365cal-act" data-act="none">Rensa</button>
			</span>
		</div>

		<div class="ms365cal-chips"></div>

		<div class="ms365cal-mode-switch" role="group" aria-label="Vytyp">
			<button type="button" class="ms365cal-mode-btn" data-mode="list">Lista</button>
			<button type="button" class="ms365cal-mode-btn" data-mode="calendar">Kalender</button>
		</div>

		<div class="ms365cal-layout-switch" role="group" aria-label="Layout" data-for-mode="list">
			<button type="button" class="ms365cal-layout-btn" data-layout="standard">Standard</button>
			<button type="button" class="ms365cal-layout-btn" data-layout="compact">Kompakt</button>
			<button type="button" class="ms365cal-layout-btn" data-layout="expanded">Utökad</button>
		</div>
		<div class="ms365cal-layout-switch" role="group" aria-label="Vy" data-for-mode="calendar" hidden>
			<button type="button" class="ms365cal-layout-btn" data-calview="day">Dag</button>
			<button type="button" class="ms365cal-layout-btn" data-calview="3day">3 dagar</button>
			<button type="button" class="ms365cal-layout-btn" data-calview="workweek">Arbetsvecka</button>
			<button type="button" class="ms365cal-layout-btn" data-calview="week">Vecka</button>
			<button type="button" class="ms365cal-layout-btn" data-calview="month">Månad</button>
		</div>

		<div class="ms365cal-nav">
			<button type="button" class="ms365cal-page" data-dir="-1" aria-label="Föregående vecka"><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="15 6 9 12 15 18"></polyline></svg></button>
			<span class="ms365cal-range"></span>
			<button type="button" class="ms365cal-page" data-dir="1" aria-label="Nästa vecka"><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="9 6 15 12 9 18"></polyline></svg></button>
		</div>

		<div class="ms365cal-list" aria-live="polite">
			<div class="ms365cal-banner">Laddar&hellip;</div>
		</div>
	</div>
	<?php
	return ob_get_clean();
}
add_shortcode( 'ms365_calendar', 'ms365cal_shortcode' );

/**
 * Front-end CSS + JS.
 */
function ms365cal_assets() {
	?>
	<style>
	.ms365cal{--ms-line:rgba(120,120,125,.22);--ms-soft:rgba(120,120,125,.09);font-size:15px;line-height:1.5;}
	.ms365cal-bar{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:14px;}
	.ms365cal-title{font-weight:700;font-size:1.05em;}
	.ms365cal-act{font-size:12px;padding:5px 12px;margin-left:6px;background:transparent;border:1px solid var(--ms-line);border-radius:999px;cursor:pointer;color:inherit;opacity:.75;transition:background .15s,opacity .15s;}
	.ms365cal-act:hover{background:var(--ms-soft);opacity:1;}
	.ms365cal-chips{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:18px;}
	.ms365cal-chip{display:inline-flex;align-items:center;gap:7px;font-size:13px;padding:6px 13px;border-radius:999px;cursor:pointer;border:1px solid var(--ms-line);background:transparent;color:inherit;opacity:.5;transition:opacity .15s,border-color .15s,background .15s,color .15s;}
	.ms365cal-chip:hover{opacity:.8;}
	.ms365cal-chip.is-on{border-color:var(--cc);background:var(--cbg);color:var(--ct);opacity:1;}
	.ms365cal-dot{width:9px;height:9px;border-radius:50%;background:currentColor;opacity:.35;}
	.ms365cal-chip.is-on .ms365cal-dot{background:var(--cc);opacity:1;}
	/* Tier 1 (List/Calendar): a filled "track" with a pill thumb — the primary, higher-level choice. Tier 2 (below) is deliberately a different shape (squared-off tabs, not pills) and visually lighter, so the two read as different kinds of control rather than one long row of similar buttons. */
	.ms365cal-mode-switch{display:flex;justify-content:flex-end;gap:2px;padding:3px;background:var(--ms-soft);border-radius:999px;margin:0 0 8px auto;width:max-content;}
	.ms365cal-mode-btn{font-size:13px;padding:7px 16px;background:transparent;border:none;border-radius:999px;cursor:pointer;color:inherit;opacity:.55;font-weight:700;transition:opacity .15s,background .15s;}
	.ms365cal-mode-btn:hover{opacity:.85;}
	.ms365cal-mode-btn.is-active{opacity:1;background:var(--ms-line);}
	.ms365cal-layout-switch{display:flex;flex-wrap:wrap;justify-content:flex-end;gap:4px;padding-top:6px;margin-bottom:10px;}
	.ms365cal-layout-switch[hidden]{display:none;}
	.ms365cal-layout-btn{font-size:11px;padding:4px 10px;background:transparent;border:1px solid transparent;border-radius:6px;cursor:pointer;color:inherit;opacity:.5;font-weight:500;transition:opacity .15s,background .15s,border-color .15s;}
	.ms365cal-layout-btn:hover{opacity:.85;background:var(--ms-soft);}
	.ms365cal-layout-btn.is-active{opacity:1;border-color:var(--ms-line);background:var(--ms-soft);font-weight:700;}
	.ms365cal-nav{display:flex;align-items:center;justify-content:center;gap:16px;margin-bottom:16px;padding:8px;border:1px solid var(--ms-line);border-radius:14px;background:var(--ms-soft);}
	.ms365cal-page{width:40px;height:40px;flex:0 0 auto;display:inline-flex;align-items:center;justify-content:center;border:none;border-radius:50%;background:rgba(120,120,125,.16);box-shadow:0 1px 2px rgba(0,0,0,.07);cursor:pointer;color:inherit;opacity:1;transition:background .15s,box-shadow .15s,transform .1s;}
	.ms365cal-page:hover{background:rgba(120,120,125,.28);box-shadow:0 2px 5px rgba(0,0,0,.12);transform:translateY(-1px);}
	.ms365cal-page:active:not(:disabled){transform:translateY(0) scale(.92);box-shadow:0 1px 2px rgba(0,0,0,.07);}
	.ms365cal-page:focus-visible{outline:2px solid currentColor;outline-offset:2px;}
	.ms365cal-page:disabled{opacity:.25;cursor:default;background:transparent;box-shadow:none;}
	.ms365cal-page:disabled:hover{background:transparent;box-shadow:none;transform:none;}
	/* Explicit size/behaviour for the chevron, not left to the svg element's own width/height attributes — a host theme's own "svg{...}" reset (common for responsive-image handling) can silently override those attributes and collapse the icon to 0 width while height still renders, which is exactly what happened here without this rule. */
	.ms365cal-page svg{display:block;width:18px;height:18px;flex:none;overflow:visible;}
	.ms365cal-range{font-size:15px;font-weight:700;opacity:.9;min-width:200px;text-align:center;}
	.ms365cal-list{position:relative;min-height:60px;}
	.ms365cal-list.is-loading{opacity:.55;transition:opacity .15s;}
	.ms365cal-banner,.ms365cal-empty{padding:1.75rem 0;text-align:center;opacity:.5;}
	.ms365cal-error{padding:1.25rem 0;text-align:center;color:#c0392b;}
	.ms365cal-error button{margin-left:8px;font-size:13px;padding:4px 11px;border:1px solid var(--ms-line);border-radius:999px;background:transparent;cursor:pointer;color:inherit;}
	.ms365cal-day{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;opacity:.5;margin:22px 0 4px;}
	.ms365cal-day:first-child{margin-top:2px;}
	.ms365cal-earlier{margin:6px 0 2px;}
	.ms365cal-earlier-tog{display:flex;align-items:center;gap:8px;width:100%;text-align:left;background:none;border:0;padding:6px 12px;margin:0;cursor:pointer;color:inherit;font:inherit;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;opacity:.55;transition:opacity .12s;}
	.ms365cal-earlier-tog:hover{opacity:.9;}
	.ms365cal-earlier-tog .ms365cal-caret{transform:none;}
	.ms365cal-earlier-tog[aria-expanded="true"] .ms365cal-caret{transform:rotate(90deg);}
	.ms365cal-row{border-radius:12px;padding:0 12px;transition:background .12s;}
	.ms365cal-row:hover{background:var(--ms-soft);}
	.ms365cal-head{display:flex;align-items:stretch;gap:10px;}
	.ms365cal-times{flex:0 0 auto;width:82px;display:flex;flex-direction:column;justify-content:space-between;align-items:flex-end;text-align:right;padding:10px 0;font-size:12px;font-variant-numeric:tabular-nums;line-height:1.25;}
	.ms365cal-t1{font-weight:600;opacity:.85;}
	.ms365cal-t2{opacity:.5;}
	.ms365cal-cat{display:inline-block;font-size:11px;font-weight:600;line-height:1.2;overflow-wrap:break-word;max-width:100%;padding:3px 8px;border-radius:999px;}
	.ms365cal-rail{width:4px;border-radius:999px;flex:0 0 auto;margin:10px 0;}
	.ms365cal-hbody{flex:1;min-width:0;padding:10px 0;display:flex;flex-direction:column;}
	.ms365cal-ev,.ms365cal-ev-static{font-size:15px;font-weight:600;background:none;border:0;padding:0;margin:0;text-align:left;text-transform:none;color:inherit;font-family:inherit;display:flex;align-items:baseline;gap:9px;width:100%;line-height:1.35;}
	.ms365cal-ev{cursor:pointer;}
	.ms365cal-ev:hover .ms365cal-title{opacity:.65;}
	.ms365cal-title{transition:opacity .12s;}
	.ms365cal-caret{display:inline-block;flex:0 0 auto;font-size:9px;opacity:.4;transition:transform .15s;}
	.ms365cal-meta-line{margin-top:auto;padding-top:8px;display:flex;align-items:center;gap:10px;font-size:12px;opacity:.6;}
	.ms365cal-loc-line{flex:1 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
	.ms365cal-recur-line{flex:0 0 auto;margin-left:auto;white-space:nowrap;}
	.ms365cal-detail{margin:6px 0 4px;font-size:13px;line-height:1.6;}
	.ms365cal-detail div,.ms365cal-detail p{margin:5px 0;}
	.ms365cal-detail div:first-child,.ms365cal-detail p:first-child{margin-top:0;}
	.ms365cal-desc{opacity:.9;}
	.ms365cal-desc-loading{opacity:.5;}
	.ms365cal-detail a{color:#185fa5;font-weight:600;text-decoration:underline;text-underline-offset:2px;}
	.ms365cal-ongoing{margin:0 0 16px;padding-bottom:8px;border-bottom:1px solid var(--ms-line);}
	.ms365cal-ongoing-label{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;opacity:.55;margin:2px 0 4px;}
	.ms365cal-preview{font-size:12.5px;opacity:.65;margin-top:3px;overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;line-clamp:2;}
	/* Compact layout: times+category is a fixed-width single line — a *fixed* width on .ms365cal-times, not just content that happens to stay on one line, is what actually keeps the rail at the same x position on every row; letting the column size to its own content (width:auto) put the rail at a different x on almost every row depending on how wide that row's time text happened to be. Category label is the flexible element that ellipsises to whatever's left, t1/t2 keep their own width cap so a long multi-day date/time can't blow the budget either. Title+meta is a strictly non-wrapping line (title ellipsises, recurrence drops to just the icon), detail still drops to its own full-width line when opened. Rows sit closer together. */
	.ms365cal-layout-compact .ms365cal-row{padding:0 10px;}
	.ms365cal-layout-compact .ms365cal-times{flex-direction:row;flex-wrap:nowrap;align-items:center;justify-content:flex-start;width:130px;flex:0 0 130px;overflow:hidden;gap:6px;text-align:left;padding:4px 0;font-size:11px;}
	.ms365cal-layout-compact .ms365cal-t1,.ms365cal-layout-compact .ms365cal-t2{flex:0 0 auto;max-width:70px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
	/* A true fixed width (flex:0 0 50px + width, not just a max-width a shorter label could fall under) — otherwise a short category name leaves a variable gap before the rail instead of the times cluster always filling the same footprint edge to edge. */
	.ms365cal-layout-compact .ms365cal-cat{flex:0 0 50px;width:50px;box-sizing:border-box;min-width:0;padding:1px 6px;font-size:10px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
	/* No margin, no rounding: the rail fills the row's full height edge-to-edge, and with zero gap between rows already, adjacent rails butt straight up against each other — reads as one continuous line that changes colour, not a row of separate pills. */
	.ms365cal-layout-compact .ms365cal-rail{margin:0;border-radius:0;}
	.ms365cal-layout-compact .ms365cal-hbody{padding:4px 0;}
	.ms365cal-layout-compact .ms365cal-line{display:flex;flex-wrap:nowrap;align-items:baseline;gap:8px;min-width:0;}
	.ms365cal-layout-compact .ms365cal-ev,.ms365cal-layout-compact .ms365cal-ev-static{flex:1 1 auto;min-width:0;width:auto;}
	.ms365cal-layout-compact .ms365cal-title{display:block;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
	.ms365cal-layout-compact .ms365cal-meta-line{flex:0 0 auto;margin-top:0;padding-top:0;white-space:nowrap;}
	.ms365cal-layout-compact .ms365cal-loc-line{max-width:150px;}
	/* Fixed width whether or not there's actually a pattern to show (see the JS: a recurrence-less row still gets an empty span here) — otherwise the location's start position would shift depending on whether a neighbouring row happened to have a recurrence icon or not. */
	.ms365cal-layout-compact .ms365cal-recur-line{display:inline-block;width:14px;flex:0 0 auto;text-align:center;cursor:default;}
	/* Expanded layout: an always-visible body snippet under the title. */
	.ms365cal-layout-expanded .ms365cal-hbody{padding:12px 0;}
	/* Calendar mode: Day/Week timeline. */
	.ms365cal-tl-header{display:flex;}
	.ms365cal-tl-axis-spacer{flex:0 0 42px;}
	.ms365cal-tl-hcell{flex:1;text-align:center;padding:4px 2px 8px;}
	.ms365cal-tl-hcell.is-today .ms365cal-tl-hdate{background:var(--ms-line);border-radius:50%;}
	.ms365cal-tl-hday{display:block;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;opacity:.55;}
	.ms365cal-tl-hdate{display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;margin-top:2px;font-size:15px;font-weight:600;}
	.ms365cal-tl-allday{display:flex;padding:4px 0;border-bottom:1px solid var(--ms-line);}
	.ms365cal-tl-allday-cell{flex:1;padding:0 2px;display:flex;flex-direction:column;gap:2px;min-width:0;}
	.ms365cal-tl-chip{display:block;width:100%;border:none;font:inherit;text-align:left;text-transform:none;letter-spacing:normal;cursor:pointer;font-size:11px;padding:2px 6px;border-radius:4px;color:#fff;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
	.ms365cal-tl-chip:hover,.ms365cal-tl-chip:focus-visible{filter:brightness(.9);outline:2px solid rgba(255,255,255,.6);outline-offset:-2px;}
	.ms365cal-tl-body{display:flex;max-height:560px;overflow-y:auto;}
	.ms365cal-tl-axis{flex:0 0 42px;}
	.ms365cal-tl-hour{height:72px;box-sizing:border-box;border-top:1px solid var(--ms-line);font-size:10.5px;opacity:.5;text-align:right;padding-right:6px;transform:translateY(-6px);}
	.ms365cal-tl-hour:first-child{border-top:0;}
	.ms365cal-tl-grid{flex:1;display:flex;position:relative;}
	/* One "segment" (whichever interval — 15/30/60min — Settings → Setup → Timeline interval chose) is always SEGMENT_PX=18px tall regardless of which interval is selected, so switching intervals changes how many segments an hour has, not how big any one segment is — HOUR_PX (see renderTimeline()) is derived from this, not the other way around. Purely visual, doesn't change what times events can start/end at. Base rule here is the 60min (hourly-only) case: 1 segment = 1 hour = 18px, no separate fine layer needed since there's nothing finer to draw. */
	.ms365cal-tl-daycol{flex:1;position:relative;min-width:0;border-left:1px solid var(--ms-line);background-image:repeating-linear-gradient(to bottom,transparent,transparent 17px,var(--ms-line) 17px,var(--ms-line) 18px);}
	/* 15min setting: 4 segments/hour, hour line every 72px; 30min: 2 segments/hour, hour line every 36px. Both share the same 18px fine-line layer (SEGMENT_PX is constant) — only the hour-boundary spacing differs between them. */
	.ms365cal-tl-fine-15 .ms365cal-tl-daycol{background-image:repeating-linear-gradient(to bottom,transparent,transparent 71px,var(--ms-line) 71px,var(--ms-line) 72px),repeating-linear-gradient(to bottom,transparent,transparent 17px,rgba(120,120,125,.1) 17px,rgba(120,120,125,.1) 18px);}
	.ms365cal-tl-fine-30 .ms365cal-tl-daycol{background-image:repeating-linear-gradient(to bottom,transparent,transparent 35px,var(--ms-line) 35px,var(--ms-line) 36px),repeating-linear-gradient(to bottom,transparent,transparent 17px,rgba(120,120,125,.1) 17px,rgba(120,120,125,.1) 18px);}
	.ms365cal-tl-daycol.is-today{background-color:var(--ms-soft);}
	.ms365cal-tl-event{display:flex;align-items:flex-start;justify-content:flex-start;border:none;font:inherit;text-align:left;text-transform:none;letter-spacing:normal;cursor:pointer;position:absolute;box-sizing:border-box;border-radius:4px;padding:2px 5px;font-size:11px;line-height:1.3;color:#fff;box-shadow:-2px 0 3px rgba(0,0,0,.18);transition:box-shadow .12s;}
	.ms365cal-tl-event:hover,.ms365cal-tl-event:focus-visible{filter:brightness(.9);outline:2px solid rgba(255,255,255,.6);outline-offset:-2px;}
	.ms365cal-tl-event.ms365cal-tl-event--front{z-index:999!important;box-shadow:0 4px 14px rgba(0,0,0,.35);}
	.ms365cal-tl-event.ms365cal-tl-event--wide{width:auto!important;max-width:260px;}
	.ms365cal-tl-ev-title{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0;}
	/* Calendar mode: Month grid (read-only — no drill-down yet). */
	.ms365cal-month-head{display:grid;grid-template-columns:repeat(7,1fr);margin-bottom:2px;}
	.ms365cal-month-hcell{text-align:center;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;opacity:.55;padding:4px 0;}
	.ms365cal-month-grid{display:grid;grid-template-columns:repeat(7,1fr);grid-auto-rows:minmax(90px,auto);}
	.ms365cal-month-cell{border:1px solid var(--ms-line);margin:-1px 0 0 -1px;padding:4px;box-sizing:border-box;overflow:hidden;}
	.ms365cal-month-cell.is-outside{opacity:.4;}
	.ms365cal-month-daynum{display:inline-block;min-width:20px;text-align:center;font-size:12px;font-weight:600;margin-bottom:3px;}
	.ms365cal-month-cell.is-today .ms365cal-month-daynum{background:var(--ms-line);border-radius:50%;}
	.ms365cal-month-chip{display:block;width:100%;border:none;font:inherit;text-align:left;text-transform:none;letter-spacing:normal;cursor:pointer;font-size:10.5px;padding:1px 5px;border-radius:3px;color:#fff;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin-bottom:2px;}
	.ms365cal-month-chip:hover,.ms365cal-month-chip:focus-visible{filter:brightness(.9);outline:2px solid rgba(255,255,255,.6);outline-offset:-2px;}
	.ms365cal-month-more{display:block;width:100%;border:none;background:none;font:inherit;text-align:left;text-transform:none;letter-spacing:normal;cursor:pointer;font-size:10px;opacity:.6;padding:0;color:inherit;}
	.ms365cal-month-more:hover,.ms365cal-month-more:focus-visible{opacity:1;text-decoration:underline;}
	.ms365cal-tl-event.ms365cal-tl-more{background:#495057;font-weight:600;}
	/* Calendar mode: event-detail popup, opened by clicking a timeline event or a Month/all-day chip — those are small colour blocks with no room for List's inline accordion. A solid background is unavoidable here (unlike the rest of this plugin, which stays transparent to blend with the host theme) since a popup has to stay legible over arbitrary page content behind it; Canvas/CanvasText are the dark-mode-aware system colours for "page background/text", with a plain #fff/#1d2327 fallback for browsers that don't support them. */
	.ms365cal-modal-backdrop{position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:1000;display:flex;align-items:center;justify-content:center;padding:20px;box-sizing:border-box;}
	.ms365cal-modal{position:relative;max-width:480px;width:100%;max-height:85vh;overflow-y:auto;background:#fff;background:Canvas;color:#1d2327;color:CanvasText;border-radius:12px;padding:28px 24px 24px;box-shadow:0 12px 40px rgba(0,0,0,.3);}
	.ms365cal-modal-close{position:absolute;top:12px;right:12px;width:32px;height:32px;border:none;border-radius:50%;background:rgba(120,120,125,.14);cursor:pointer;font-size:14px;color:inherit;text-transform:none;letter-spacing:normal;display:flex;align-items:center;justify-content:center;}
	.ms365cal-modal-close:hover{background:rgba(120,120,125,.28);}
	.ms365cal-modal-cat{display:inline-block;font-size:11px;font-weight:600;padding:3px 9px;border-radius:999px;margin-bottom:10px;}
	.ms365cal-modal-title{margin:0 0 10px;font-size:18px;line-height:1.35;padding-right:24px;}
	.ms365cal-modal-row{font-size:13px;opacity:.75;margin-bottom:4px;}
	.ms365cal-modal-recur{opacity:.6;}
	.ms365cal-modal-detail{margin-top:14px;font-size:13.5px;line-height:1.6;}
	.ms365cal-modal-list-body{display:flex;flex-direction:column;gap:6px;max-height:60vh;overflow-y:auto;}
	.ms365cal-modal-list-item{display:flex;align-items:center;gap:10px;width:100%;box-sizing:border-box;padding:8px 10px;border:1px solid var(--ms-line);border-radius:8px;background:none;font:inherit;text-align:left;text-transform:none;letter-spacing:normal;cursor:pointer;color:inherit;}
	.ms365cal-modal-list-item:hover,.ms365cal-modal-list-item:focus-visible{background:var(--ms-soft);outline:2px solid rgba(0,0,0,.15);outline-offset:-2px;}
	.ms365cal-modal-list-dot{width:10px;height:10px;border-radius:50%;flex:0 0 auto;}
	.ms365cal-modal-list-text{display:flex;flex-direction:column;gap:2px;min-width:0;}
	.ms365cal-modal-list-title{font-weight:600;font-size:14px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
	.ms365cal-modal-list-when{font-size:12px;opacity:.65;}
	</style>
	<script>
	(function(){
	function esc(s){var d=document.createElement('div');d.textContent=s==null?'':s;return d.innerHTML;}
	// esc() alone doesn't escape quotes (safe for text nodes, not for inside a
	// quoted HTML attribute) — used only for the id/slug we embed in data-* below.
	function escAttr(s){return esc(s).replace(/"/g,'&quot;');}
	// Query-string separator for appending params to a REST url that may already
	// carry one — e.g. rest_url() returns "index.php?rest_route=..." on a site
	// without pretty permalinks, so blindly appending "?more=params" would nest a
	// second "?" inside the rest_route value instead of adding a real param.
	function qsep(url){return url.indexOf('?')===-1?'?':'&';}
	function pad(n){return (n<10?'0':'')+n;}
	function iso(d){return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());}
	function fmt(d){return d.toLocaleDateString('sv-SE',{day:'numeric',month:'short'});}
	// ISO-8601 week number (Swedish "vecka" numbering): the week containing that
	// week's Thursday determines both the week number and its ISO year.
	function isoWeek(d){
		var t=new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate()));
		var day=t.getUTCDay()||7;
		t.setUTCDate(t.getUTCDate()+4-day);
		var yearStart=new Date(Date.UTC(t.getUTCFullYear(),0,1));
		return Math.ceil(((t-yearStart)/86400000+1)/7);
	}

	function initCal(root){
		var cfg;
		try{cfg=JSON.parse(root.getAttribute('data-config'));}catch(e){return;}

		var listEl  = root.querySelector('.ms365cal-list');
		var chipsEl = root.querySelector('.ms365cal-chips');
		var rangeEl = root.querySelector('.ms365cal-range');

		var enabled = {};
		cfg.cals.forEach(function(s){enabled[s]=!!cfg.defaults[s];});

		// Monday of the week containing d (weeks start Monday).
		function weekStart(d){var x=new Date(d.getFullYear(),d.getMonth(),d.getDate());x.setDate(x.getDate()-((x.getDay()+6)%7));return x;}
		var today    = new Date(cfg.startY, cfg.startM-1, cfg.startD); // today, site tz
		var todayKey = iso(today);

		// Every fresh page load starts in the cheap, predictable List/Standard
		// view. Visitors can still switch modes for the current page session, but
		// prior cookies/localStorage no longer override the initial render.
		var mode='list',layout='standard',calView='week';

		var refDate=new Date(today);   // anchor date the current window is derived from
		var monthRef=null;             // first-of-month Date, set only while calView==='month'
		var floorOverride=null;        // raised once by the empty-week auto-advance (list mode only)
		var start,days;

		function computeWindow(){
			if(mode==='calendar'&&calView==='day'){
				return {start:new Date(refDate.getFullYear(),refDate.getMonth(),refDate.getDate()),days:1};
			}
			if(mode==='calendar'&&calView==='3day'){
				return {start:new Date(refDate.getFullYear(),refDate.getMonth(),refDate.getDate()),days:3};
			}
			if(mode==='calendar'&&calView==='workweek'){
				return {start:weekStart(refDate),days:5}; // Mon–Fri of the week
			}
			if(mode==='calendar'&&calView==='month'){
				var first=new Date(refDate.getFullYear(),refDate.getMonth(),1);
				var last=new Date(refDate.getFullYear(),refDate.getMonth()+1,0);
				var gridStart=weekStart(first);
				var gridEndExcl=weekStart(last);gridEndExcl.setDate(gridEndExcl.getDate()+7);
				return {start:gridStart,days:Math.round((gridEndExcl-gridStart)/86400000),monthRef:first};
			}
			return {start:weekStart(refDate),days:7}; // list mode, or calendar/week
		}
		function computeMinStart(){
			if(floorOverride)return floorOverride;
			if(mode==='calendar'&&(calView==='day'||calView==='3day'))return new Date(today.getFullYear(),today.getMonth(),today.getDate());
			if(mode==='calendar'&&calView==='month')return weekStart(new Date(today.getFullYear(),today.getMonth(),1));
			return weekStart(today); // list, week, workweek
		}
		function recomputeWindow(){
			var w=computeWindow();
			start=w.start;days=w.days;monthRef=w.monthRef||null;
		}
		recomputeWindow();

		function updateNavLabels(){
			var unit='vecka';
			if(mode==='calendar'){
				if(calView==='day')unit='dag';
				else if(calView==='3day')unit='3 dagar';
				else if(calView==='workweek')unit='arbetsvecka';
				else if(calView==='month')unit='månad';
			}
			var prevBtn=root.querySelector('.ms365cal-page[data-dir="-1"]');
			var nextBtn=root.querySelector('.ms365cal-page[data-dir="1"]');
			if(prevBtn)prevBtn.setAttribute('aria-label','Föregående '+unit);
			if(nextBtn)nextBtn.setAttribute('aria-label','Nästa '+unit);
		}
		function applyView(){
			root.classList.remove('ms365cal-layout-standard','ms365cal-layout-compact','ms365cal-layout-expanded','ms365cal-mode-list','ms365cal-mode-calendar');
			root.classList.add('ms365cal-mode-'+mode);
			if(mode==='list')root.classList.add('ms365cal-layout-'+layout);
			root.querySelectorAll('.ms365cal-mode-btn').forEach(function(b){
				b.classList.toggle('is-active',b.getAttribute('data-mode')===mode);
			});
			root.querySelectorAll('[data-for-mode]').forEach(function(g){
				g.hidden=(g.getAttribute('data-for-mode')!==mode);
			});
			root.querySelectorAll('.ms365cal-layout-btn').forEach(function(b){
				var isActive=mode==='list'?(b.getAttribute('data-layout')===layout):(b.getAttribute('data-calview')===calView);
				b.classList.toggle('is-active',isActive);
			});
			updateNavLabels();
		}
		root.querySelectorAll('.ms365cal-mode-btn').forEach(function(b){
			b.addEventListener('click',function(){
				var next=b.getAttribute('data-mode');
				if(next===mode)return;
				mode=next;
				floorOverride=null;
				// Ground on today rather than carrying over wherever navigation had
				// drifted to in the mode just left — e.g. paging Month forward a few
				// months and then switching to List shouldn't land on some unrelated
				// past week just because refDate was last normalised to a 1st-of-month.
				refDate=new Date(today);
				recomputeWindow();
				applyView();
				load();
			});
		});
		root.querySelectorAll('.ms365cal-layout-btn').forEach(function(b){
			b.addEventListener('click',function(){
				if(mode==='list'){
					var nextLayout=b.getAttribute('data-layout');
					if(!nextLayout||nextLayout===layout)return;
					layout=nextLayout;
					applyView();
					paint(); // expanded needs the preview line (re)rendered
				}else{
					var nextView=b.getAttribute('data-calview');
					if(!nextView||nextView===calView)return;
					calView=nextView;
					refDate=new Date(today); // same reasoning as the mode switch above
					recomputeWindow();
					applyView();
					load();
				}
			});
		});
		applyView();

		var cache = {};                                   // window key -> events[]
		var reqId = 0;
		var retryTimer = null;
		var throttleRetries = 0;
		var MAX_THROTTLE_RETRIES = 3;
		var autoAdvanceChecked = false;                   // only skip an empty first week once
		var calEvents = [];                               // events currently rendered (List or Calendar mode), for click-to-lookup by data-idx
		var clusterGroups = [];                           // groups of events behind a "+N" overflow badge, for click-to-lookup by data-cluster-idx
		var currentModal = null;                          // the open event-detail popup, if any

		function winKey(){return iso(start)+'|'+days;}
		function updateRange(){
			if(mode==='calendar'&&calView==='day'){
				rangeEl.textContent=start.toLocaleDateString('sv-SE',{weekday:'long',day:'numeric',month:'long'});
			}else if(mode==='calendar'&&calView==='3day'){
				var end3=new Date(start);end3.setDate(end3.getDate()+days-1);
				rangeEl.textContent=fmt(start)+' \u2013 '+fmt(end3);
			}else if(mode==='calendar'&&calView==='month'&&monthRef){
				var label=monthRef.toLocaleDateString('sv-SE',{month:'long',year:'numeric'});
				rangeEl.textContent=label.charAt(0).toUpperCase()+label.slice(1);
			}else{
				// Week (List or Calendar), and Workweek \u2014 a specific ISO week either way,
				// just fewer days shown for Workweek (Mon\u2013Fri instead of Mon\u2013Sun).
				var end=new Date(start);end.setDate(end.getDate()+days-1);
				rangeEl.textContent='Vecka '+isoWeek(start)+' \u00b7 '+fmt(start)+' \u2013 '+fmt(end);
			}
		}
		function updateNav(){
			var prev=root.querySelector('.ms365cal-page[data-dir="-1"]');
			if(prev)prev.disabled=(start<=computeMinStart());
		}

		function renderChips(){
			chipsEl.innerHTML='';
			cfg.cals.forEach(function(slug){
				var m=cfg.meta[slug],on=enabled[slug];
				var b=document.createElement('button');
				b.type='button';
				b.className='ms365cal-chip'+(on?' is-on':'');
				b.style.setProperty('--cc',m.primary);
				b.style.setProperty('--cbg',m.supplement);
				b.style.setProperty('--ct',m.text);
				b.innerHTML='<span class="ms365cal-dot"></span>'+esc(m.label);
				b.addEventListener('click',function(){enabled[slug]=!enabled[slug];b.classList.toggle('is-on');paint();});
				chipsEl.appendChild(b);
			});
		}

		// Shared between List mode's accordion and the Calendar-mode popup: the
		// join-link/online/Outlook-link lines plus the body itself if already in
		// hand (eager mode). In lazy mode e.body is '' here — the body arrives
		// later via loadBody(), which inserts it into whatever container holds
		// this HTML, so those synchronous links still need to render up front
		// rather than waiting on the fetch.
		function buildDetailHtml(e){
			var d='';
			if(e.body)d+='<div class="ms365cal-desc">'+e.body+'</div>';
			if(e.joinUrl)d+='<div><a href="'+esc(e.joinUrl)+'" target="_blank" rel="noopener">Anslut till onlinemöte</a></div>';
			else if(e.online)d+='<div>Onlinemöte</div>';
			if(cfg.showOutlook&&e.link)d+='<div><a href="'+esc(e.link)+'" target="_blank" rel="noopener">Öppna i Outlook ↗</a></div>';
			return d;
		}

		function closeEventModal(){
			if(!currentModal)return;
			currentModal.remove();
			currentModal=null;
			document.removeEventListener('keydown',modalKeyHandler);
		}
		function modalKeyHandler(ev){
			if(ev.key==='Escape')closeEventModal();
		}
		// The calendar-grid views (Day/Week/Month) show events as small colour
		// blocks/chips with no room for a body — clicking one opens this instead
		// of the List view's inline accordion, reusing buildDetailHtml()/loadBody()
		// so lazy-body fetching works identically either way.
		function openEventModal(e){
			closeEventModal();
			var m=cfg.meta[e.cal];if(!m)return;
			var locText=e.location?e.location:(e.online?'Online':'');
			var recurShort=e.recur?e.recur.replace(/^Upprepas\s+/,''):'';
			var metaHtml='';
			if(e.when)metaHtml+='<div class="ms365cal-modal-row">'+esc(e.when)+'</div>';
			if(locText)metaHtml+='<div class="ms365cal-modal-row">'+esc(locText)+'</div>';
			if(e.recur)metaHtml+='<div class="ms365cal-modal-row ms365cal-modal-recur" title="'+escAttr(e.recur)+'">↻ '+esc(recurShort)+'</div>';

			var lazyOk=cfg.lazyBody&&e.id&&e.hasBody;
			var lazyAttrs=lazyOk?' data-cal="'+escAttr(e.cal)+'" data-id="'+escAttr(e.id)+'" data-loaded="0"':'';

			var backdrop=document.createElement('div');
			backdrop.className='ms365cal-modal-backdrop';
			backdrop.innerHTML='<div class="ms365cal-modal" role="dialog" aria-modal="true">'
				+'<button type="button" class="ms365cal-modal-close" aria-label="Stäng">✕</button>'
				+'<span class="ms365cal-modal-cat" style="color:'+m.text+';background:'+m.supplement+'">'+esc(m.label)+'</span>'
				+'<h3 class="ms365cal-modal-title">'+esc(e.title)+'</h3>'
				+metaHtml
				+'<div class="ms365cal-detail ms365cal-modal-detail"'+lazyAttrs+'>'+buildDetailHtml(e)+'</div>'
			+'</div>';
			document.body.appendChild(backdrop);
			currentModal=backdrop;
			document.addEventListener('keydown',modalKeyHandler);
			backdrop.addEventListener('click',function(ev){if(ev.target===backdrop)closeEventModal();});
			backdrop.querySelector('.ms365cal-modal-close').addEventListener('click',closeEventModal);
			backdrop.querySelector('.ms365cal-modal-close').focus();
			if(lazyOk)loadBody(backdrop.querySelector('.ms365cal-modal-detail'));
		}

		// A "+N" overflow badge (Month's "+N fler", or a cascaded timeline cluster)
		// opens this instead of jumping straight to one event's detail — a small
		// pick list, since which specific event the user meant isn't known yet. Picking
		// a row opens that event's own detail via openEventModal(), same as clicking
		// it directly would if it weren't covered/overflowed in the first place.
		function openEventListModal(events){
			closeEventModal();
			var rows='';
			events.forEach(function(e,i){
				var m=cfg.meta[e.cal];if(!m)return;
				var whenHtml=e.when?'<span class="ms365cal-modal-list-when">'+esc(e.when)+'</span>':'';
				rows+='<button type="button" class="ms365cal-modal-list-item" data-ev-idx="'+i+'">'
					+'<span class="ms365cal-modal-list-dot" style="background:'+m.primary+'"></span>'
					+'<span class="ms365cal-modal-list-text"><span class="ms365cal-modal-list-title">'+esc(e.title)+'</span>'+whenHtml+'</span>'
				+'</button>';
			});
			var backdrop=document.createElement('div');
			backdrop.className='ms365cal-modal-backdrop';
			backdrop.innerHTML='<div class="ms365cal-modal" role="dialog" aria-modal="true">'
				+'<button type="button" class="ms365cal-modal-close" aria-label="Stäng">✕</button>'
				+'<h3 class="ms365cal-modal-title">'+events.length+' händelser</h3>'
				+'<div class="ms365cal-modal-list-body">'+rows+'</div>'
			+'</div>';
			document.body.appendChild(backdrop);
			currentModal=backdrop;
			document.addEventListener('keydown',modalKeyHandler);
			backdrop.addEventListener('click',function(ev){if(ev.target===backdrop)closeEventModal();});
			backdrop.querySelector('.ms365cal-modal-close').addEventListener('click',closeEventModal);
			backdrop.querySelector('.ms365cal-modal-close').focus();
			backdrop.querySelectorAll('.ms365cal-modal-list-item').forEach(function(btn){
				btn.addEventListener('click',function(){
					var idx=parseInt(btn.getAttribute('data-ev-idx'),10);
					if(!isNaN(idx)&&events[idx])openEventModal(events[idx]);
				});
			});
		}

		function renderDays(list,showDayHeaders){
			if(showDayHeaders===undefined)showDayHeaders=true;
			var html='',lastDay='';
			list.forEach(function(e){
				var m=cfg.meta[e.cal];if(!m)return;
				if(showDayHeaders&&e.dayKey!==lastDay){html+='<div class="ms365cal-day">'+esc(e.dayLabel)+'</div>';lastDay=e.dayKey;}

				// Bottom line: location on the left, recurrence pinned to the right (via
				// margin-left:auto on the recurrence span, so it stays right-aligned
				// whether or not a location is present).
				var isRecurring=!!(e.recurring||e.recur);
				var recurTitle=e.recur||'Återkommande händelse';
				var recurShort=e.recur?e.recur.replace(/^Upprepas\s+/,''):'';
				var locText=e.location?e.location:(e.online?'Online':'');
				var metaBits='';
				if(locText)metaBits+='<span class="ms365cal-loc-line">'+esc(locText)+'</span>';
				if(isRecurring){
					metaBits+='<span class="ms365cal-recur-line" title="'+escAttr(recurTitle)+'">\u21bb'+(recurShort?' '+esc(recurShort):'')+'</span>';
				}
				var metaLine=metaBits?'<div class="ms365cal-meta-line">'+metaBits+'</div>':'';

				// Compact only: the recurrence pattern text is dropped to just the icon
				// (full pattern still available on hover via title=) so the line has room
				// to stay a single line instead of wrapping. Whenever there's a location,
				// the recurrence slot is always present (empty span if there's no actual
				// pattern) \u2014 a fixed-width CSS rule then keeps the location's start
				// position identical whether or not that row happens to recur, instead of
				// it shifting by however wide the icon would've been.
				var compactBits='';
				if(locText){
					compactBits+='<span class="ms365cal-loc-line">'+esc(locText)+'</span>';
					compactBits+=isRecurring
						?'<span class="ms365cal-recur-line" title="'+escAttr(recurTitle)+'">\u21bb</span>'
						:'<span class="ms365cal-recur-line" aria-hidden="true"></span>';
				}else if(isRecurring){
					compactBits+='<span class="ms365cal-recur-line" title="'+escAttr(recurTitle)+'">\u21bb</span>';
				}
				var compactMetaLine=compactBits?'<div class="ms365cal-meta-line">'+compactBits+'</div>':'';

				// No "when" line here \u2014 the times are already in the left column.
				// No location either \u2014 it's on the always-visible meta line above.
				// A title with anything to show opens the same popup Calendar mode
				// uses (buildDetailHtml()/loadBody() are shared, so lazy-body fetching
				// works identically either way) instead of the old inline accordion.
				// In lazy mode (cfg.lazyBody), e.hasBody comes from Graph's cheap
				// bodyPreview field (see ms365cal_build_rows_from_page()) \u2014 known at
				// list-fetch time, so an event with nothing to show (no body, no join
				// link, no online/outlook line) renders as plain non-clickable text
				// instead of an expand button that would do nothing.
				var showDetail=buildDetailHtml(e)||(cfg.lazyBody&&e.id&&e.hasBody);
				// data-idx indexes into calEvents, set to the currently-painted list
				// at the top of paint() — same lookup mechanism the calendar-grid
				// views use for their own popup.
				var titleHtml=showDetail
					?'<button type="button" class="ms365cal-ev" data-idx="'+calEvents.indexOf(e)+'"><span class="ms365cal-title">'+esc(e.title)+'</span></button>'
					:'<div class="ms365cal-ev-static"><span class="ms365cal-title">'+esc(e.title)+'</span></div>';

				// Expanded layout only: a short always-visible snippet of the body
				// (e.preview is plain text from the server — see ms365cal_build_rows_
				// from_page() — so esc() here, not injected as markup like e.body).
				var previewHtml=(layout==='expanded'&&e.preview)?'<div class="ms365cal-preview">'+esc(e.preview)+'</div>':'';

				// Compact wraps title+meta in their own nowrap flex line (.ms365cal-line)
				// so the two can never wrap apart onto separate lines the way plain
				// flex-wrapping siblings could — title ellipsises instead.
				var hbodyInner=(layout==='compact')
					?('<div class="ms365cal-line">'+titleHtml+compactMetaLine+'</div>')
					:(titleHtml+previewHtml+metaLine);

				html+='<div class="ms365cal-row">'
					+'<div class="ms365cal-head">'
						+'<div class="ms365cal-times">'
							+'<span class="ms365cal-t1">'+esc(e.t1)+'</span>'
							+'<span class="ms365cal-cat" style="color:'+m.text+';background:'+m.supplement+'">'+esc(m.label)+'</span>'
							+'<span class="ms365cal-t2">'+esc(e.t2)+'</span>'
						+'</div>'
						+'<div class="ms365cal-rail" style="background:'+m.primary+'"></div>'
						+'<div class="ms365cal-hbody">'+hbodyInner+'</div>'
					+'</div>'
				+'</div>';
			});
			return html;
		}

		// Greedy lane layout for same-day timed events that overlap in time: sort by
		// start, group into connected overlap-clusters (a gap with nothing active
		// breaks the group), then within each cluster give each event the first lane
		// whose previous occupant has already ended, else a new lane. Rendering (below)
		// turns laneIndex/laneCount into a cascade for just the first N visible lanes
		// (N set per render \u2014 see cascadeVisibleMax() below) \u2014 not an equal width
		// split (which on a busy day shrank lanes to unreadable slivers) and not an
		// unbounded cascade either (which stayed technically readable but still
		// looked cluttered once 6+ lanes stacked on a single busy day) \u2014 each visible
		// lane offsetting right by CASCADE_STEP_PCT and layering on top of the
		// previous ones.
		var CASCADE_STEP_PCT=18;
		// How many lanes to cascade before folding the rest into "+N fler" box(es),
		// based on how many days are in view \u2014 a proxy for how wide each day column
		// actually is (measuring real rendered width would need a second pass after
		// insertion, which isn't worth the complexity/flicker here). Per explicit
		// request: keep as many real events visible as reasonably fit rather than
		// collapsing down to "just the one event, plus a badge for the rest" \u2014 an
		// earlier version did exactly that for Week/Workweek specifically (reasoning
		// that those columns can be very narrow, e.g. 64px measured live on a
		// sidebar-constrained layout) and it read as showing *less* real information
		// than the user wanted, not as calmer. Day/3-day still get one lane more than
		// Week/Workweek, since they do have more room, but neither collapses this low.
		function cascadeVisibleMax(numDays){
			return numDays<=3?5:4;
		}
		function layoutLanes(items){
			items.sort(function(a,b){return a.start-b.start;});
			var i=0;
			while(i<items.length){
				var groupEnd=items[i].end;
				var j=i+1;
				while(j<items.length&&items[j].start<groupEnd){
					if(items[j].end>groupEnd)groupEnd=items[j].end;
					j++;
				}
				var lanesEnd=[];
				for(var k=i;k<j;k++){
					var it=items[k],laneIdx=-1;
					for(var li=0;li<lanesEnd.length;li++){
						if(lanesEnd[li]<=it.start){laneIdx=li;break;}
					}
					if(laneIdx===-1){laneIdx=lanesEnd.length;lanesEnd.push(it.end);}
					else lanesEnd[laneIdx]=it.end;
					it.laneIndex=laneIdx;
				}
				var clusterMembers=items.slice(i,j);
				for(var k2=i;k2<j;k2++){items[k2].laneCount=lanesEnd.length;items[k2].clusterMembers=clusterMembers;}
				i=j;
			}
		}

		// Day/Week: a vertical hour axis with events positioned/sized by actual
		// clock time (row.start/row.end \u2014 the real timing, unlike the list view's
		// dayKey/sort which can be pinned to today for display purposes). Colour
		// only, no calendar label chip \u2014 see renderMonth() for the same choice.
		function renderTimeline(events,numDays){
			calEvents=events;
			clusterGroups=[];
			var CASCADE_VISIBLE_MAX=cascadeVisibleMax(numDays);
			// One segment (one interval unit — see the .ms365cal-tl-fine-* CSS rules)
			// is always SEGMENT_PX tall regardless of which interval is selected, so
			// HOUR_PX (how tall a full hour renders) scales with how many segments
			// make up an hour rather than being a fixed constant — switching to a
			// finer interval makes the timeline taller, not the segments smaller.
			var SEGMENT_PX=18;
			var HOUR_PX=SEGMENT_PX*(60/cfg.tlInterval);
			var dayCols=[];
			for(var i=0;i<numDays;i++){
				var d=new Date(start);d.setDate(d.getDate()+i);
				dayCols.push({date:d,key:iso(d),timed:[],allDay:[]});
			}
			var colByKey={};dayCols.forEach(function(c){colByKey[c.key]=c;});
			var winEnd=new Date(start);winEnd.setDate(winEnd.getDate()+numDays);

			events.forEach(function(e){
				if(!e.start)return;
				var st=new Date(e.start);
				var en=e.end?new Date(e.end):new Date(st.getTime()+30*60000);
				if(e.allDay||iso(st)!==iso(en)){
					var cursor=new Date(st<start?start:st);
					var endClamped=new Date(en<winEnd?en:winEnd);
					while(cursor<endClamped){
						var key=iso(cursor);
						if(colByKey[key])colByKey[key].allDay.push(e);
						cursor.setDate(cursor.getDate()+1);
					}
					return;
				}
				var key2=iso(st);
				if(colByKey[key2])colByKey[key2].timed.push({ev:e,start:st,end:en});
			});
			dayCols.forEach(function(col){layoutLanes(col.timed);});

			// Trim the hour axis to roughly bound the events actually present,
			// rather than always showing the full 24h — a day with everything
			// between 9 and 17 shouldn't force scrolling past 9 empty hours above
			// and 9 below it. ~1h of padding on each side; falls back to a plain
			// working-hours window when there's nothing timed to bound against
			// (an all-day-only day, or a genuinely empty one).
			var minStartMin=null,maxEndMin=null;
			dayCols.forEach(function(col){col.timed.forEach(function(it){
				var sM=it.start.getHours()*60+it.start.getMinutes();
				var eM=it.end.getHours()*60+it.end.getMinutes();
				if(eM<=sM)eM=24*60;
				if(minStartMin===null||sM<minStartMin)minStartMin=sM;
				if(maxEndMin===null||eM>maxEndMin)maxEndMin=eM;
			});});
			var PAD_MIN=60;
			var rangeStartHour,rangeEndHour;
			if(minStartMin===null){
				rangeStartHour=7;rangeEndHour=19;
			}else{
				rangeStartHour=Math.max(0,Math.floor((minStartMin-PAD_MIN)/60));
				rangeEndHour=Math.min(24,Math.ceil((maxEndMin+PAD_MIN)/60));
				if(rangeEndHour<=rangeStartHour)rangeEndHour=Math.min(24,rangeStartHour+1);
			}
			var rangeStartMin=rangeStartHour*60;

			var tlIntervalClass=cfg.tlInterval===30?' ms365cal-tl-fine-30':(cfg.tlInterval===60?'':' ms365cal-tl-fine-15');
			var html='<div class="ms365cal-tl'+tlIntervalClass+'">';
			html+='<div class="ms365cal-tl-header"><div class="ms365cal-tl-axis-spacer"></div>';
			dayCols.forEach(function(col){
				html+='<div class="ms365cal-tl-hcell'+(col.key===todayKey?' is-today':'')+'">'
					+'<span class="ms365cal-tl-hday">'+esc(col.date.toLocaleDateString('sv-SE',{weekday:'short'}))+'</span>'
					+'<span class="ms365cal-tl-hdate">'+col.date.getDate()+'</span>'
				+'</div>';
			});
			html+='</div>';

			if(dayCols.some(function(c){return c.allDay.length;})){
				html+='<div class="ms365cal-tl-allday"><div class="ms365cal-tl-axis-spacer"></div>';
				dayCols.forEach(function(col){
					html+='<div class="ms365cal-tl-allday-cell">';
					col.allDay.forEach(function(e){
						var m=cfg.meta[e.cal];if(!m)return;
						html+='<button type="button" class="ms365cal-tl-chip" style="background:'+m.primary+'" title="'+escAttr(e.title)+'" data-idx="'+events.indexOf(e)+'">'+esc(e.title)+'</button>';
					});
					html+='</div>';
				});
				html+='</div>';
			}

			html+='<div class="ms365cal-tl-body"><div class="ms365cal-tl-axis">';
			for(var h=rangeStartHour;h<rangeEndHour;h++)html+='<div class="ms365cal-tl-hour" style="height:'+HOUR_PX+'px">'+pad(h)+':00</div>';
			html+='</div><div class="ms365cal-tl-grid" style="height:'+((rangeEndHour-rangeStartHour)*HOUR_PX)+'px">';
			function endMinOf(x){
				var sm=x.start.getHours()*60+x.start.getMinutes();
				var em=x.end.getHours()*60+x.end.getMinutes();
				return (x.end.getDate()!==x.start.getDate()||em<=sm)?24*60:em;
			}
			dayCols.forEach(function(col){
				html+='<div class="ms365cal-tl-daycol'+(col.key===todayKey?' is-today':'')+'">';
				var badgedClusters=[];
				col.timed.forEach(function(it){
					var m=cfg.meta[it.ev.cal];if(!m)return;
					if(it.laneIndex>=CASCADE_VISIBLE_MAX)return;
					var overflowing=it.laneCount>CASCADE_VISIBLE_MAX;
					var leftPct=it.laneIndex*CASCADE_STEP_PCT;
					// The last visible lane of an overflowing cluster doesn't render
					// this specific event as its own box — instead everything at or
					// past it is split into "+N fler" box(es), computed once per
					// cluster (below) rather than per occurrence, but potentially
					// emitting *several* boxes, not one. A single box spanning the
					// cluster's entire time range read as "busy from here to there"
					// even across long stretches where nothing was actually hidden at
					// that moment — instead, the hidden subset is re-grouped by its own
					// actual overlap (the same connected-touching-or-overlapping
					// technique layoutLanes() itself uses, just applied to only the
					// overflow members), so each box is tightly scoped to when — and
					// exactly what — is genuinely hidden right then. Each box occupies
					// its own dedicated lane slot with the exact same geometry math as
					// a real event, styled the same way (same class, just a neutral
					// colour instead of a specific calendar's) — not a floating overlay
					// button — so it reads as more items in the timeline rather than a
					// decoration bolted on top of one, and can never overlap another
					// box's text since it has its own slot like any other lane.
					if(overflowing&&it.laneIndex===CASCADE_VISIBLE_MAX-1){
						if(badgedClusters.indexOf(it.clusterMembers)===-1){
							badgedClusters.push(it.clusterMembers);
							// clusterMembers is sorted by start time, not by lane
							// assignment, so "what's hidden" isn't just the array's
							// tail — filter on laneIndex itself first.
							var hidden=it.clusterMembers.filter(function(x){return x.laneIndex>=CASCADE_VISIBLE_MAX-1;});
							var hi=0;
							while(hi<hidden.length){
								var subEnd=endMinOf(hidden[hi]);
								var hj=hi+1;
								while(hj<hidden.length&&(hidden[hj].start.getHours()*60+hidden[hj].start.getMinutes())<subEnd){
									var hjEnd=endMinOf(hidden[hj]);
									if(hjEnd>subEnd)subEnd=hjEnd;
									hj++;
								}
								var sub=hidden.slice(hi,hj);
								var subStartMin=sub[0].start.getHours()*60+sub[0].start.getMinutes();
								var moreTop=((subStartMin-rangeStartMin)/60)*HOUR_PX;
								var moreHeight=Math.max(18,((subEnd-subStartMin)/60)*HOUR_PX);
								var cidx=clusterGroups.length;
								clusterGroups.push(sub.map(function(x){return x.ev;}));
								html+='<button type="button" class="ms365cal-tl-event ms365cal-tl-more" style="top:'+moreTop+'px;height:'+moreHeight+'px;left:'+leftPct+'%;width:calc(100% - '+leftPct+'%);z-index:'+(it.laneIndex+1)+'" data-cluster-idx="'+cidx+'" aria-label="'+sub.length+' tilläggshändelser i denna period"><span class="ms365cal-tl-ev-title">+'+sub.length+' fler</span></button>';
								hi=hj;
							}
						}
						return;
					}
					var startMin=it.start.getHours()*60+it.start.getMinutes();
					var top=((startMin-rangeStartMin)/60)*HOUR_PX;
					var height=Math.max(18,((endMinOf(it)-startMin)/60)*HOUR_PX);
					html+='<button type="button" class="ms365cal-tl-event" style="top:'+top+'px;height:'+height+'px;left:'+leftPct+'%;width:calc(100% - '+leftPct+'%);z-index:'+(it.laneIndex+1)+';background:'+m.primary+'" title="'+escAttr(it.ev.title)+'" data-idx="'+events.indexOf(it.ev)+'"><span class="ms365cal-tl-ev-title">'+esc(it.ev.title)+'</span></button>';
				});
				html+='</div>';
			});
			html+='</div></div></div>';
			listEl.innerHTML=html;
		}

		// Month: each cell capped at a few event chips plus a "+N fler" overflow
		// button (clicking it lists just the overflow \u2014 the capped chips are already
		// directly clickable) \u2014 same colour-only/no-label choice as the timeline
		// views. Still no click-through from an empty cell to Day view for that date.
		function renderMonth(events,numDays,ref){
			calEvents=events;
			clusterGroups=[];
			var dayCols=[];
			for(var i=0;i<numDays;i++){
				var d=new Date(start);d.setDate(d.getDate()+i);
				dayCols.push({date:d,key:iso(d),events:[],inMonth:d.getMonth()===ref.getMonth()});
			}
			var colByKey={};dayCols.forEach(function(c){colByKey[c.key]=c;});
			var winEnd=new Date(start);winEnd.setDate(winEnd.getDate()+numDays);

			events.forEach(function(e){
				if(!e.start)return;
				var st=new Date(e.start);
				var en=e.end?new Date(e.end):st;
				var spans=e.allDay||iso(st)!==iso(en);
				if(!spans){
					var key=iso(st);
					if(colByKey[key])colByKey[key].events.push(e);
					return;
				}
				var cursor=new Date(st<start?start:st);
				var endClamped=new Date(en<winEnd?en:winEnd);
				while(cursor<endClamped){
					var k=iso(cursor);
					if(colByKey[k])colByKey[k].events.push(e);
					cursor.setDate(cursor.getDate()+1);
				}
			});

			var WEEKDAYS=['M\u00e5n','Tis','Ons','Tor','Fre','L\u00f6r','S\u00f6n'];
			var CAP=3;
			var html='<div class="ms365cal-month"><div class="ms365cal-month-head">';
			WEEKDAYS.forEach(function(l){html+='<div class="ms365cal-month-hcell">'+l+'</div>';});
			html+='</div><div class="ms365cal-month-grid">';
			dayCols.forEach(function(col){
				html+='<div class="ms365cal-month-cell'+(col.inMonth?'':' is-outside')+(col.key===todayKey?' is-today':'')+'">'
					+'<span class="ms365cal-month-daynum">'+col.date.getDate()+'</span>';
				col.events.slice(0,CAP).forEach(function(e){
					var m=cfg.meta[e.cal];if(!m)return;
					html+='<button type="button" class="ms365cal-month-chip" style="background:'+m.primary+'" title="'+escAttr(e.title)+'" data-idx="'+events.indexOf(e)+'">'+esc(e.title)+'</button>';
				});
				if(col.events.length>CAP){
					var cidx=clusterGroups.length;
					clusterGroups.push(col.events.slice(CAP));
					html+='<button type="button" class="ms365cal-month-more" data-cluster-idx="'+cidx+'">+'+(col.events.length-CAP)+' fler</button>';
				}
				html+='</div>';
			});
			html+='</div></div>';
			listEl.innerHTML=html;
		}

		function paint(){
			var events=cache[winKey()]||[];
			var visible=events.filter(function(e){return enabled[e.cal];});
			if(mode==='calendar'){
				if(calView==='month')renderMonth(visible,days,monthRef||start);
				else renderTimeline(visible,days);
				return;
			}
			if(!visible.length){
				listEl.innerHTML='<p class="ms365cal-empty">Inga h\u00e4ndelser den h\u00e4r veckan.</p>';
				return;
			}
			calEvents=visible;
			// Events pinned to today but actually started before this window (a
			// still-running event from an earlier week) get pulled into their own
			// section above the day list instead of blending into today's day
			// group, where they'd otherwise read as if they started today.
			var carried=[],rest=[];
			visible.forEach(function(e){(e.carriedOver?carried:rest).push(e);});
			var past=[],upcoming=[];
			rest.forEach(function(e){(e.dayKey<todayKey?past:upcoming).push(e);});

			var html='';
			if(carried.length){
				html+='<div class="ms365cal-ongoing">'
					+'<div class="ms365cal-ongoing-label">P\u00e5g\u00e5r sedan tidigare</div>'
					+renderDays(carried,false)
				+'</div>';
			}
			if(past.length){
				html+='<div class="ms365cal-earlier">'
					+'<button type="button" class="ms365cal-earlier-tog" aria-expanded="false">'
						+'<span class="ms365cal-caret">\u25b8</span>Tidigare H\u00e4ndelser'
					+'</button>'
					+'<div class="ms365cal-earlier-body" hidden>'+renderDays(past)+'</div>'
				+'</div>';
			}
			html+=renderDays(upcoming);
			listEl.innerHTML=html;
		}

		// Lazy body fetch (cfg.lazyBody): the detail container already carries
		// data-cal/data-id, set by openEventModal() for List and Calendar mode alike;
		// this is called once per event, the first time its popup opens, guarded by
		// data-loaded so a rapid close/reopen doesn't refetch.
		function loadBody(detail){
			var cal=detail.getAttribute('data-cal'),id=detail.getAttribute('data-id');
			detail.setAttribute('data-loaded','1');
			if(!cal||!id)return;
			var loading=document.createElement('div');
			loading.className='ms365cal-desc-loading';
			loading.textContent='Laddar…';
			detail.insertBefore(loading,detail.firstChild);
			fetch(cfg.bodyRest+qsep(cfg.bodyRest)+'cal='+encodeURIComponent(cal)+'&id='+encodeURIComponent(id),{headers:{'Accept':'application/json'}})
				.then(function(r){return r.ok?r.json():{body:''};})
				.then(function(data){
					if(data&&data.body){
						var desc=document.createElement('div');
						desc.className='ms365cal-desc';
						// Same trust boundary as the eager path: server-sanitised HTML
						// (wp_kses, see ms365cal_sanitize_event_html()), injected as markup.
						desc.innerHTML=data.body;
						loading.replaceWith(desc);
					}else{
						loading.remove();
					}
				})
				.catch(function(){loading.remove();});
		}

		// Independent toggle for the "Tidigare Händelser" group (unrelated to any
		// event popup, just collapses/expands the past-days section).
		listEl.addEventListener('click',function(ev){
			var tog=ev.target.closest('.ms365cal-earlier-tog');
			if(!tog||!listEl.contains(tog))return;
			var eb=tog.parentNode.querySelector('.ms365cal-earlier-body');
			var op=tog.getAttribute('aria-expanded')==='true';
			tog.setAttribute('aria-expanded',op?'false':'true');
			if(eb)eb.hidden=op;
		});

		// Clicking an event — a List row's title (.ms365cal-ev), a Calendar-mode
		// block/chip (Day/Week timeline, Month grid, or the all-day strip) — opens
		// the same popup. List used to expand an inline accordion instead; the
		// calendar-grid views never had room for that in a small colour block, and
		// once both share one popup there's a single, consistent place all detail
		// rendering (buildDetailHtml()/loadBody()) lives instead of two. data-idx
		// indexes into calEvents, set fresh on every paint()/renderTimeline()/
		// renderMonth(). A "+N" overflow badge (data-cluster-idx into clusterGroups)
		// opens the pick-list modal instead — checked first since it's a plain
		// sibling of the events it overflows, not a descendant, so there's no risk
		// of it also matching the event selector.
		listEl.addEventListener('click',function(ev){
			var moreEl=ev.target.closest('.ms365cal-tl-more,.ms365cal-month-more');
			if(moreEl&&listEl.contains(moreEl)){
				var cidx=parseInt(moreEl.getAttribute('data-cluster-idx'),10);
				if(!isNaN(cidx)&&clusterGroups[cidx])openEventListModal(clusterGroups[cidx]);
				return;
			}
			var el=ev.target.closest('.ms365cal-ev,.ms365cal-tl-event,.ms365cal-tl-chip,.ms365cal-month-chip');
			if(!el||!listEl.contains(el))return;
			var idx=parseInt(el.getAttribute('data-idx'),10);
			if(isNaN(idx)||!calEvents[idx])return;
			openEventModal(calEvents[idx]);
		});

		// Bring a covered cascaded timeline event to the front on hover or keyboard
		// focus, and also widen it if its own title is truncated. Bringing to front
		// is unconditional — a covered event can fit its own (nominal, pre-cascade)
		// box width and still be entirely hidden behind a later lane's box drawn on
		// top of it, so gating the front-bump on truncation (an earlier version of
		// this) left some covered events unreachable by hover. Widening only grows
		// the box (never shrinks it), and only when the title is actually truncated
		// (scrollWidth>clientWidth) — a blanket width:auto would shrink an
		// already-full-width event down to its short text's natural width instead of
		// leaving it alone. Hover is skipped entirely on touch/coarse-pointer devices
		// (no real hover state there, and tapping already opens the full-detail
		// popup); keyboard focus stays active everywhere since it's a real,
		// deliberate interaction.
		function bringToFront(el){
			el.classList.add('ms365cal-tl-event--front');
			var span=el.querySelector('.ms365cal-tl-ev-title');
			if(span&&span.scrollWidth>span.clientWidth+0.5)el.classList.add('ms365cal-tl-event--wide');
		}
		function resetFront(el){
			el.classList.remove('ms365cal-tl-event--front','ms365cal-tl-event--wide');
		}
		var supportsHover=window.matchMedia&&window.matchMedia('(hover:hover) and (pointer:fine)').matches;
		if(supportsHover){
			listEl.addEventListener('mouseover',function(ev){
				var el=ev.target.closest('.ms365cal-tl-event');
				if(!el||!listEl.contains(el)||el.contains(ev.relatedTarget))return;
				bringToFront(el);
			});
			listEl.addEventListener('mouseout',function(ev){
				var el=ev.target.closest('.ms365cal-tl-event');
				if(!el||!listEl.contains(el)||el.contains(ev.relatedTarget))return;
				resetFront(el);
			});
		}
		listEl.addEventListener('focusin',function(ev){
			var el=ev.target.closest('.ms365cal-tl-event');
			if(el&&listEl.contains(el))bringToFront(el);
		});
		listEl.addEventListener('focusout',function(ev){
			var el=ev.target.closest('.ms365cal-tl-event');
			if(el&&listEl.contains(el))resetFront(el);
		});

		function load(){
			if(retryTimer){clearTimeout(retryTimer);retryTimer=null;}
			updateRange();updateNav();
			var key=winKey();
			if(cache[key]){paint();return;}          // client-side window cache hit

			var my=++reqId;
			listEl.classList.remove('is-loading');
			// Clear the previous week immediately and show a loading indicator, rather
			// than leaving stale events on screen until the new window arrives.
			listEl.innerHTML='<div class="ms365cal-banner">Laddar\u2026</div>';

			var url=cfg.rest+qsep(cfg.rest)+'cals='+encodeURIComponent(cfg.cals.join(','))
				+'&start='+iso(start)+'&days='+days;

			fetch(url,{headers:{'Accept':'application/json'}})
				.then(function(r){
					if(r.status===429){
						var ra=parseInt(r.headers.get('Retry-After')||'0',10);
						var e=new Error('429');e.retryAfter=(ra>0?ra:30);throw e;
					}
					if(!r.ok)throw new Error(r.status);
					return r.json();
				})
				.then(function(data){
					if(my!==reqId)return;                // a newer request superseded this
					throttleRetries=0;
					cache[key]=data.events||[];

					// If this is the initial (earliest) week and it has nothing left from
					// today onward — ignoring multi-day events that only "cover" the week
					// as part of a longer span (their real end is past this week), and
					// events only present because they carried over from an earlier
					// week — default to next week instead, and treat it as the new
					// earliest week. List mode only: the day/week/month calendar-grid
					// views always show their actual window (that's the point of a grid),
					// they don't skip ahead looking for content.
					if(mode==='list'&&!autoAdvanceChecked&&iso(start)===iso(computeMinStart())){
						autoAdvanceChecked=true;
						var hasUpcoming=cache[key].some(function(e){
							return e.dayKey>=todayKey&&enabled[e.cal]&&!e.longSpan&&!e.carriedOver;
						});
						if(!hasUpcoming){
							refDate=new Date(refDate);refDate.setDate(refDate.getDate()+days);
							recomputeWindow();
							floorOverride=new Date(start);
							load();
							return;
						}
					}

					listEl.classList.remove('is-loading');
					paint();
				})
				.catch(function(err){
					if(my!==reqId)return;
					listEl.classList.remove('is-loading');

					if(''+err.message==='429'){
						var wait=Math.min(err.retryAfter||30,120);
						var auto=throttleRetries<MAX_THROTTLE_RETRIES;
						listEl.innerHTML='<p class="ms365cal-error">Upptaget just nu \u2014 '
							+(auto?'f\u00f6rs\u00f6ker igen om '+wait+'s\u2026':'f\u00f6rs\u00f6k igen om en stund.')
							+'<button type="button" class="ms365cal-retry">F\u00f6rs\u00f6k igen</button></p>';
						var rb=listEl.querySelector('.ms365cal-retry');
						if(rb)rb.addEventListener('click',function(){throttleRetries=0;load();});
						if(auto){
							throttleRetries++;
							retryTimer=setTimeout(load,wait*1000);
						}
						return;
					}

					listEl.innerHTML='<p class="ms365cal-error">Kunde inte ladda kalendern.'
						+'<button type="button" class="ms365cal-retry">F\u00f6rs\u00f6k igen</button></p>';
					var rb2=listEl.querySelector('.ms365cal-retry');
					if(rb2)rb2.addEventListener('click',load);
				});
		}

		// Step refDate by whatever unit the current mode/view actually pages by —
		// a day, a week, or a full calendar month (not a fixed number of days,
		// since months vary in length).
		function stepRef(dir){
			if(mode==='calendar'&&calView==='day'){
				refDate=new Date(refDate.getFullYear(),refDate.getMonth(),refDate.getDate()+dir);
			}else if(mode==='calendar'&&calView==='3day'){
				refDate=new Date(refDate.getFullYear(),refDate.getMonth(),refDate.getDate()+dir*3);
			}else if(mode==='calendar'&&calView==='month'){
				refDate=new Date(refDate.getFullYear(),refDate.getMonth()+dir,1);
			}else{
				// Week (List or Calendar) and Workweek both page by a whole week.
				refDate=new Date(refDate.getFullYear(),refDate.getMonth(),refDate.getDate()+dir*7);
			}
		}

		var COOLDOWN=600; // ms
		var pageBtns=root.querySelectorAll('.ms365cal-page');
		function setPaging(disabled){pageBtns.forEach(function(b){b.disabled=disabled;});}

		pageBtns.forEach(function(btn){
			btn.addEventListener('click',function(){
				if(btn.disabled)return;
				var dir=parseInt(btn.getAttribute('data-dir'),10);
				var prevIso=iso(start),savedRef=new Date(refDate);
				stepRef(dir);
				recomputeWindow();
				if(start<computeMinStart()){refDate=new Date(today);recomputeWindow();}
				if(iso(start)===prevIso){refDate=savedRef;recomputeWindow();return;} // already at the bound
				setPaging(true);
				load();
				setTimeout(function(){setPaging(false);updateNav();},COOLDOWN);
			});
		});
		root.querySelectorAll('.ms365cal-act').forEach(function(btn){
			btn.addEventListener('click',function(){
				var on=btn.getAttribute('data-act')==='all';
				cfg.cals.forEach(function(s){enabled[s]=on;});
				renderChips();paint();
			});
		});

		renderChips();
		load();
	}

	document.querySelectorAll('.ms365cal').forEach(initCal);
	})();
	</script>
	<?php
}
add_action( 'wp_footer', 'ms365cal_assets' );

/**
 * ---------------------------------------------------------------------------
 *  Admin settings
 * ---------------------------------------------------------------------------
 */
function ms365cal_admin_menu() {
	add_options_page( 'MS365 Calendar', 'MS365 Calendar', 'manage_options', 'ms365cal', 'ms365cal_settings_page' );
}
add_action( 'admin_menu', 'ms365cal_admin_menu' );

/**
 * A small "?" icon next to a field label that reveals $html (trusted, hand-
 * written markup — not user input) as a tooltip on hover/focus, instead of a
 * permanent paragraph of description text under every field. $html is echoed
 * unescaped by design, same as the rest of this admin template's own literal
 * markup; never pass it anything derived from user input.
 */
function ms365cal_help( $html ) {
	// phpcs:ignore WordPress.Security.EscapeOutput.OutputNotEscaped -- $html is always a hand-written literal at the call site, never user input; see the docblock above.
	echo '<span class="ms365cal-help dashicons dashicons-editor-help" tabindex="0"><span class="ms365cal-tip">' . $html . '</span></span>';
}

function ms365cal_settings_page() {
	if ( ! current_user_can( 'manage_options' ) ) {
		return;
	}

	if ( isset( $_POST['ms365cal_nonce'] ) && wp_verify_nonce( sanitize_key( $_POST['ms365cal_nonce'] ), 'ms365cal_save' ) ) {
		$new = ms365cal_get_settings();

		$new['tenant_id'] = sanitize_text_field( wp_unslash( $_POST['tenant_id'] ?? '' ) );
		$new['client_id'] = sanitize_text_field( wp_unslash( $_POST['client_id'] ?? '' ) );

		$posted_secret = trim( wp_unslash( $_POST['client_secret'] ?? '' ) );
		if ( '' !== $posted_secret ) {
			$new['client_secret'] = $posted_secret;
		}
		$new['cache_minutes']       = max( 1, (int) ( $_POST['cache_minutes'] ?? 20 ) );
		$new['cache_grace_minutes'] = max( 0, (int) ( $_POST['cache_grace_minutes'] ?? 60 ) );
		$new['timezone']            = sanitize_text_field( wp_unslash( $_POST['timezone'] ?? 'UTC' ) );
		$new['rate_max']            = max( 0, (int) ( $_POST['rate_max'] ?? MS365CAL_RATE_MAX ) );
		$new['rate_window']         = max( 1, (int) ( $_POST['rate_window'] ?? MS365CAL_RATE_WINDOW ) );
		$new['show_outlook']        = ! empty( $_POST['show_outlook'] );
		$new['show_recurrence']     = ! empty( $_POST['show_recurrence'] );
		$posted_interval            = (int) ( $_POST['timeline_interval'] ?? 15 );
		$new['timeline_interval']   = in_array( $posted_interval, array( 15, 30, 60 ), true ) ? $posted_interval : 15;
		$new['events_top']          = max( 10, min( 999, (int) ( $_POST['events_top'] ?? 100 ) ) );
		$new['lazy_body']           = ! empty( $_POST['lazy_body'] );

		// Deploy key: like the client secret, a blank submission keeps the stored value;
		// the explicit "clear" checkbox is the only way to erase it (disable the endpoint).
		if ( ! empty( $_POST['deploy_key_clear'] ) ) {
			$new['deploy_key'] = '';
		} else {
			$posted_deploy = trim( wp_unslash( $_POST['deploy_key'] ?? '' ) );
			if ( '' !== $posted_deploy ) {
				$new['deploy_key'] = $posted_deploy;
			}
		}

		$new['update_repo'] = sanitize_text_field( wp_unslash( $_POST['update_repo'] ?? '' ) );

		$cals = array();
		$rows = isset( $_POST['cal'] ) && is_array( $_POST['cal'] ) ? $_POST['cal'] : array();
		foreach ( $rows as $row ) {
			$slug = sanitize_title( $row['slug'] ?? '' );
			$src  = trim( sanitize_text_field( wp_unslash( $row['source'] ?? '' ) ) );
			if ( '' === $slug || '' === $src ) {
				continue;
			}
			$color = sanitize_hex_color( $row['color'] ?? '' );
			if ( ! $color ) {
				$color = ms365cal_color_palette()[0]['primary'];
			}
			$cals[] = array(
				'slug'   => $slug,
				'label'  => sanitize_text_field( wp_unslash( $row['label'] ?? $slug ) ),
				'color'  => $color,
				'type'   => ms365cal_detect_calendar_type( $src ),
				'source' => $src,
			);
		}
		$new['calendars'] = $cals;

		update_option( MS365CAL_OPTION, $new );
		delete_transient( MS365CAL_TOKEN_TRANSIENT );
		echo '<div class="notice notice-success"><p>Settings saved.</p></div>';
	}

	$s                 = ms365cal_get_settings();
	$has_secret        = (bool) ( $s['client_secret'] || ( defined( 'MS365CAL_CLIENT_SECRET' ) && MS365CAL_CLIENT_SECRET ) );
	$deploy_const      = defined( 'MS365CAL_DEPLOY_KEY' ) && MS365CAL_DEPLOY_KEY;
	$has_deploy        = (bool) ( $s['deploy_key'] || $deploy_const );
	$update_repo_const = defined( 'MS365CAL_UPDATE_REPO' ) && MS365CAL_UPDATE_REPO;
	?>
	<style>
	.ms365cal-help{position:relative;display:inline-block;margin-left:5px;color:#787c82;cursor:help;font-size:16px;vertical-align:text-bottom;}
	.ms365cal-help:hover,.ms365cal-help:focus{color:#2271b1;outline:none;}
	.ms365cal-help .ms365cal-tip{display:none;position:absolute;left:0;top:26px;z-index:10;width:340px;max-width:min(340px,80vw);padding:10px 12px;background:#1d2327;color:#f0f0f1;font-size:12.5px;font-weight:400;line-height:1.6;border-radius:4px;box-shadow:0 2px 10px rgba(0,0,0,.2);}
	.ms365cal-help .ms365cal-tip code{background:rgba(255,255,255,.15);color:inherit;padding:1px 4px;border-radius:3px;}
	.ms365cal-help:hover .ms365cal-tip,.ms365cal-help:focus .ms365cal-tip{display:block;}
	.ms365cal-tab-panel{display:none;}
	.ms365cal-tab-panel.is-active{display:block;}
	.ms365cal-color-picker{position:relative;display:inline-block;}
	.ms365cal-color-trigger{display:flex;align-items:center;gap:8px;width:120px;box-sizing:border-box;padding:5px 10px;border:1px solid #8c8f94;border-radius:4px;background:#fff;cursor:pointer;font-size:13px;color:#1d2327;}
	.ms365cal-color-trigger:hover{border-color:#2271b1;}
	.ms365cal-color-bar{display:inline-block;width:26px;height:4px;border-radius:2px;flex:0 0 auto;}
	.ms365cal-color-label{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
	.ms365cal-color-menu{display:none;position:absolute;top:100%;left:0;z-index:20;margin-top:2px;width:120px;box-sizing:border-box;max-height:260px;overflow-y:auto;background:#fff;border:1px solid #8c8f94;border-radius:4px;box-shadow:0 2px 8px rgba(0,0,0,.15);}
	.ms365cal-color-menu.is-open{display:block;}
	.ms365cal-color-option{display:flex;align-items:center;gap:8px;width:100%;box-sizing:border-box;padding:6px 10px;border:0;background:none;cursor:pointer;font-size:13px;color:#1d2327;text-align:left;}
	.ms365cal-color-option:hover,.ms365cal-color-option:focus{background:#f0f0f1;outline:none;}
	.ms365cal-src-picker{position:relative;}
	.ms365cal-src-menu{display:none;position:absolute;top:100%;left:0;z-index:20;margin-top:2px;width:420px;max-width:90vw;max-height:320px;overflow-y:auto;background:#fff;border:1px solid #8c8f94;border-radius:4px;box-shadow:0 2px 8px rgba(0,0,0,.15);}
	.ms365cal-src-menu.is-open{display:block;}
	.ms365cal-src-section{padding:4px 10px;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.03em;color:#787c82;background:#f6f7f7;}
	.ms365cal-src-option{display:block;width:100%;box-sizing:border-box;padding:6px 10px;border:0;background:none;cursor:pointer;font-size:13px;color:#1d2327;text-align:left;}
	.ms365cal-src-option:hover,.ms365cal-src-option:focus{background:#f0f0f1;outline:none;}
	.ms365cal-src-name{display:block;}
	.ms365cal-src-addr{display:block;font-size:11.5px;color:#787c82;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
	.ms365cal-src-empty,.ms365cal-src-hint{padding:8px 10px;font-size:12px;color:#787c82;}
	.ms365cal-src-error{padding:8px 10px;font-size:12px;color:#b32d2e;}
	</style>
	<div class="wrap">
		<h1>MS365 Merged Calendar</h1>
		<h2 class="nav-tab-wrapper">
			<a href="#calendars" class="nav-tab" data-tab="calendars">Calendars</a>
			<a href="#caching" class="nav-tab" data-tab="caching">Caching</a>
			<a href="#setup" class="nav-tab" data-tab="setup">Setup</a>
			<a href="#diagnostics" class="nav-tab" data-tab="diagnostics">Diagnostics</a>
		</h2>
		<form method="post">
			<?php wp_nonce_field( 'ms365cal_save', 'ms365cal_nonce' ); ?>

			<div class="ms365cal-tab-panel" data-tab="calendars">
				<p>Start typing in the <strong>Source</strong> field to search groups, rooms, and mailboxes
				from your tenant — click a match to fill it in, or just type/paste a GUID or email
				address directly if you already know it. For a <strong>group</strong> the source is its
				object ID (GUID); for a <strong>mailbox</strong> (shared, personal, or room) it's an email
				address. Which one it is gets detected from that format automatically on save — no need to
				say which. <strong>Label</strong> and <strong>Slug</strong> stay disabled until a source is
				set; picking one from the list fills the Label in for you (still editable afterward) and
				assigns a random colour, but the Slug is left for you to fill in yourself.</p>
				<table class="widefat" id="ms365cal-rows">
					<thead><tr>
						<th>Label</th><th>Slug</th><th>Source (GUID or email)</th><th>Colour</th><th></th>
					</tr></thead>
					<tbody>
					<?php
					$palette = ms365cal_color_palette();
					$rows    = ! empty( $s['calendars'] ) ? $s['calendars'] : array(
						array(
							'slug'   => '',
							'label'  => '',
							'color'  => $palette[0]['primary'],
							'source' => '',
						),
					);
					foreach ( $rows as $i => $c ) :
						// The selected option: an exact palette match, or the closest one for a
						// colour saved before this palette existed, so legacy colours still land
						// somewhere sensible on first view instead of showing nothing selected.
						$checked_set = ms365cal_color_set( $c['color'] );
						// Label/slug only make sense once a source is chosen — see the source
						// picker's JS, which flips these back on (and, when picked from the
						// directory, fills the label) the moment the source field gets a value.
						$no_source = ( '' === trim( $c['source'] ) );
						?>
						<tr>
							<td><input type="text" name="cal[<?php echo (int) $i; ?>][label]" value="<?php echo esc_attr( $c['label'] ); ?>" <?php disabled( $no_source ); ?>></td>
							<td><input type="text" name="cal[<?php echo (int) $i; ?>][slug]" value="<?php echo esc_attr( $c['slug'] ); ?>" <?php disabled( $no_source ); ?>></td>
							<td>
								<div class="ms365cal-src-picker">
									<input type="text" name="cal[<?php echo (int) $i; ?>][source]" class="regular-text ms365cal-src-input" autocomplete="off" value="<?php echo esc_attr( $c['source'] ); ?>" placeholder="Type to search, or paste a GUID/email">
									<div class="ms365cal-src-menu" role="listbox" hidden></div>
								</div>
							</td>
							<td>
								<div class="ms365cal-color-picker">
									<input type="hidden" name="cal[<?php echo (int) $i; ?>][color]" value="<?php echo esc_attr( $checked_set['primary'] ); ?>" class="ms365cal-color-value">
									<button type="button" class="ms365cal-color-trigger" aria-haspopup="listbox" aria-expanded="false">
										<span class="ms365cal-color-bar" style="background:<?php echo esc_attr( $checked_set['primary'] ); ?>"></span>
										<span class="ms365cal-color-label"><?php echo esc_html( $checked_set['name'] ); ?></span>
									</button>
									<div class="ms365cal-color-menu" role="listbox" hidden>
										<?php foreach ( $palette as $set ) : ?>
											<button type="button" class="ms365cal-color-option" role="option" data-value="<?php echo esc_attr( $set['primary'] ); ?>" data-name="<?php echo esc_attr( $set['name'] ); ?>">
												<span class="ms365cal-color-bar" style="background:<?php echo esc_attr( $set['primary'] ); ?>"></span><?php echo esc_html( $set['name'] ); ?>
											</button>
										<?php endforeach; ?>
									</div>
								</div>
							</td>
							<td><button type="button" class="button ms365cal-del">Remove</button></td>
						</tr>
					<?php endforeach; ?>
					</tbody>
				</table>
				<p><button type="button" class="button" id="ms365cal-add">Add calendar</button></p>

				<h2>Usage</h2>
				<p><code>[ms365_calendar]</code> &mdash; all calendars, current week (Monday&ndash;Sunday), all checked.<br>
				<code>[ms365_calendar calendars="eng,events"]</code> &mdash; only these slugs are shown at all.<br>
				<code>[ms365_calendar enabled="eng"]</code> &mdash; all shown, but only <code>eng</code> starts checked (the
				rest are still there as chips, just unchecked until clicked).<br>
				<code>[ms365_calendar calendars="eng,events" enabled="eng"]</code> &mdash; both together.<br>
				Navigate forward week by week; the current week is the earliest.</p>
			</div>

			<div class="ms365cal-tab-panel" data-tab="caching">
				<table class="form-table">
					<tr><th>Cache (minutes)</th><td><input type="number" name="cache_minutes" min="1" value="<?php echo esc_attr( $s['cache_minutes'] ); ?>"></td></tr>
					<tr><th>Cache grace (minutes)<?php ms365cal_help( 'Once the cache expires, keep serving the old data for up to this many <em>extra</em> minutes instead of making the visitor wait on a live Graph call &mdash; a background request refreshes it at the same time. Self-sufficient: any real visitor within the grace window triggers the refresh, no external pre-warming needed. <code>0</code> disables this and always fetches live the moment the cache expires.' ); ?></th><td>
						<input type="number" name="cache_grace_minutes" min="0" value="<?php echo esc_attr( $s['cache_grace_minutes'] ); ?>">
					</td></tr>
					<tr><th>Events per page<?php ms365cal_help( 'Number of events requested per calendar per Graph API call (<code>$top</code>). Default <code>100</code>. A calendar only needs a follow-up pagination request when it has more events in the window than this value, so raising it comfortably above your busiest calendar\'s normal event count lets most fetches skip pagination entirely. Larger values mean a bigger response payload; 999 is the practical ceiling.' ); ?></th><td>
						<input type="number" name="events_top" min="10" max="999" value="<?php echo esc_attr( $s['events_top'] ); ?>" style="width:6em">
					</td></tr>
					<tr><th>Event descriptions<?php ms365cal_help( 'On by default. Fetching and sanitising every event\'s full HTML description on every list load adds real cost even when nobody reads most of them. With this on, the list fetch skips the description entirely and a small follow-up request (<code>GET /event-body</code>) fetches just that one event\'s description the first time it\'s expanded, cached briefly afterward. Turning it off restores the old behaviour &mdash; every description is fetched up front, so expanding is instant but cold loads are slower.' ); ?></th><td>
						<label><input type="checkbox" name="lazy_body" <?php checked( ! empty( $s['lazy_body'] ) ); ?>> Only fetch an event's description when a visitor expands it (recommended)</label>
					</td></tr>
					<tr><th>Recurrence pattern<?php ms365cal_help( 'Off by default. <strong>Showing this requires an extra Microsoft Graph lookup per distinct recurring series</strong> to resolve the pattern, which can meaningfully slow down a cold (uncached) load &mdash; on this plugin\'s own benchmarks it cost roughly as much time as fetching the events themselves. Disabling it skips those lookups entirely for a faster cold load.' ); ?></th><td>
						<label><input type="checkbox" name="show_recurrence" <?php checked( ! empty( $s['show_recurrence'] ) ); ?>> Show a recurrence pattern (e.g. &ldquo;Repeats weekly on Mon, Wed&rdquo;) for recurring events</label>
					</td></tr>
				</table>
			</div>

			<div class="ms365cal-tab-panel" data-tab="setup">
				<p>Register an app in Entra ID, grant <code>Application</code> permissions
				<code>Calendars.Read</code> (shared mailboxes) and <code>Group.Read.All</code>
				(group calendars), then grant admin consent and add a client secret.</p>
				<table class="form-table">
					<tr><th>Tenant ID</th><td><input type="text" name="tenant_id" class="regular-text" value="<?php echo esc_attr( $s['tenant_id'] ); ?>"></td></tr>
					<tr><th>Client ID</th><td><input type="text" name="client_id" class="regular-text" value="<?php echo esc_attr( $s['client_id'] ); ?>"></td></tr>
					<tr><th>Client secret<?php ms365cal_help( 'Stored in the database. For better security, define <code>MS365CAL_CLIENT_SECRET</code> in wp-config.php instead.' ); ?></th><td>
						<input type="password" name="client_secret" class="regular-text" value="" placeholder="<?php echo $has_secret ? '&bull;&bull;&bull;&bull;&bull;&bull; (leave blank to keep)' : 'Enter secret'; ?>">
					</td></tr>
					<tr><th>Timezone</th><td><input type="text" name="timezone" class="regular-text" value="<?php echo esc_attr( $s['timezone'] ); ?>"> <span class="description">e.g. Europe/Stockholm</span></td></tr>
					<tr><th>Rate limit<?php ms365cal_help( 'Applies to the public events endpoint. Set requests to <code>0</code> to disable the limiter. Cached views don\'t count toward it.' ); ?></th><td>
						<input type="number" name="rate_max" min="0" value="<?php echo esc_attr( $s['rate_max'] ); ?>" style="width:6em">
						requests per
						<input type="number" name="rate_window" min="1" value="<?php echo esc_attr( $s['rate_window'] ); ?>" style="width:6em">
						seconds, per IP.
					</td></tr>
					<tr><th>Outlook link<?php ms365cal_help( 'Off by default. When enabled, each expanded event includes a link to open it in Outlook on the web.' ); ?></th><td>
						<label><input type="checkbox" name="show_outlook" <?php checked( ! empty( $s['show_outlook'] ) ); ?>> Show an &ldquo;Open in Outlook&rdquo; link in the expanded event details</label>
					</td></tr>
					<tr><th>Timeline interval<?php ms365cal_help( 'How finely the Day/Week/3-day/Workweek calendar-mode timeline is subdivided with faint guide lines, to make it easier to judge where an event starts and ends at a glance. Only affects the grid lines drawn between the hour marks &mdash; not the hour marks themselves, and not what times events can actually start/end at (that\'s whatever Graph returns). Default 15 minutes.' ); ?></th><td>
						<select name="timeline_interval">
							<?php foreach ( array( 15, 30, 60 ) as $opt ) : ?>
								<option value="<?php echo esc_attr( $opt ); ?>" <?php selected( (int) $s['timeline_interval'], $opt ); ?>><?php echo 60 === $opt ? 'Hourly only (no sub-lines)' : esc_html( $opt . ' minutes' ); ?></option>
							<?php endforeach; ?>
						</select>
					</td></tr>
					<tr><th>Deploy key<?php ms365cal_help( 'Enables <code>POST /wp-json/ms365cal/v1/self-update</code> (header <code>X-MS365CAL-Deploy-Key</code>), which makes this site install the plugin\'s latest GitHub release on demand. Leave blank to keep the endpoint disabled. Anyone with the key can trigger a reinstall, so use a long random value and rotate it if it leaks. Defining <code>MS365CAL_DEPLOY_KEY</code> in wp-config.php is more secure and takes precedence.' ); ?></th><td>
						<?php if ( $deploy_const ) : ?>
							<p class="description"><strong>Set via <code>MS365CAL_DEPLOY_KEY</code> in wp-config.php</strong>, which overrides this field.</p>
						<?php else : ?>
							<input type="password" name="deploy_key" class="regular-text" value="" placeholder="<?php echo $has_deploy ? '&bull;&bull;&bull;&bull;&bull;&bull; (leave blank to keep)' : 'Paste a long random secret'; ?>">
							<?php if ( $has_deploy ) : ?>
								<label style="margin-left:8px"><input type="checkbox" name="deploy_key_clear" value="1"> Clear (disable endpoint)</label>
							<?php endif; ?>
						<?php endif; ?>
					</td></tr>
					<tr><th>Update source<?php ms365cal_help( 'GitHub repo to check for plugin updates, as <code>owner/repo</code> or a full URL. <strong>Blank means self-update stays off</strong> &mdash; nothing is assumed. The repo must be <strong>public</strong>, and each release needs a <code>vX.Y.Z</code> tag matching (or ahead of) the <code>Version:</code> header in the release you publish.' ); ?></th><td>
						<?php if ( $update_repo_const ) : ?>
							<p class="description"><strong>Set via <code>MS365CAL_UPDATE_REPO</code> in wp-config.php</strong>, which overrides this field. Current value: <code><?php echo esc_html( MS365CAL_UPDATE_REPO ); ?></code></p>
						<?php else : ?>
							<input type="text" name="update_repo" class="regular-text" value="<?php echo esc_attr( $s['update_repo'] ); ?>" placeholder="owner/repo">
						<?php endif; ?>
					</td></tr>
				</table>
			</div>

			<div class="ms365cal-tab-panel" data-tab="diagnostics">
				<?php
				$debug_url = add_query_arg(
					array(
						'debug'    => 1,
						'days'     => 60,
						'_wpnonce' => wp_create_nonce( 'wp_rest' ),
					),
					rest_url( 'ms365cal/v1/events' )
				);
				?>
				<p>
					<a href="<?php echo esc_url( $debug_url ); ?>" target="_blank" rel="noopener" class="button">Run calendar diagnostic</a>
					<?php
					ms365cal_help(
						'Opens a JSON report of each calendar\'s live Graph status (HTTP code, event count, and any error).
						A <code>403</code> means missing permission or admin consent; <code>404</code> means a wrong source ID
						or type; <code>200</code> with <code>events: 0</code> means the credentials and IDs are fine, so it\'s
						the date window. The link already carries the security nonce that REST cookie-auth requires &mdash;
						typing the URL by hand without it returns the normal (public) empty response instead. The nonce
						expires after a while, so reload this page for a fresh link. Only administrators can see the report.'
					);
					?>
				</p>
			</div>

			<?php submit_button(); ?>
		</form>
	</div>

	<script>
	(function(){
		var body=document.querySelector('#ms365cal-rows tbody');
		var idx=body.querySelectorAll('tr').length;
		var directoryUrl=<?php echo wp_json_encode( esc_url_raw( rest_url( 'ms365cal/v1/directory' ) ) ); ?>;
		var directoryNonce=<?php echo wp_json_encode( wp_create_nonce( 'wp_rest' ) ); ?>;
		document.getElementById('ms365cal-add').addEventListener('click',function(){
			var tr=body.querySelector('tr').cloneNode(true);
			tr.querySelectorAll('input,select').forEach(function(el){
				el.name=el.name.replace(/cal\[\d+\]/,'cal['+idx+']');
				if(el.classList.contains('ms365cal-color-value')){
					return; // handled below, from the cloned menu's own first option
				}
				if(el.type==='checkbox'){
					el.checked=false;
				}else if(el.tagName==='SELECT'){
					el.selectedIndex=0;
				}else{
					el.value='';
				}
			});
			var firstOption=tr.querySelector('.ms365cal-color-option');
			if(firstOption)setColorPicker(tr.querySelector('.ms365cal-color-picker'),firstOption.dataset.value,firstOption.dataset.name);
			updateRowEnablement(tr);
			body.appendChild(tr);idx++;
		});
		body.addEventListener('click',function(e){
			if(e.target.classList.contains('ms365cal-del')){
				if(body.querySelectorAll('tr').length>1)e.target.closest('tr').remove();
			}
		});

		// Custom colour picker: a button that always shows the current colour
		// (bar + name) and toggles a small popup list of the same, instead of a
		// native <select> — a native select's closed state doesn't reliably
		// reflect a styled <option> across browsers, and <option> can't hold
		// more than plain coloured text (no room for a proper colour bar).
		function setColorPicker(picker,value,name){
			if(!picker)return;
			picker.querySelector('.ms365cal-color-value').value=value;
			picker.querySelector('.ms365cal-color-bar').style.background=value;
			picker.querySelector('.ms365cal-color-label').textContent=name;
		}
		// Label/slug are only meaningful once a source is picked, so they stay
		// disabled (greyed out) until the source field has something in it —
		// re-checked on every keystroke/selection, not just once.
		function updateRowEnablement(tr){
			var hasSource=tr.querySelector('.ms365cal-src-input').value.trim()!=='';
			var labelInput=tr.querySelector('input[name$="[label]"]');
			var slugInput=tr.querySelector('input[name$="[slug]"]');
			if(labelInput)labelInput.disabled=!hasSource;
			if(slugInput)slugInput.disabled=!hasSource;
		}
		function closeAllMenus(){
			document.querySelectorAll('.ms365cal-color-menu.is-open').forEach(function(m){
				m.hidden=true;m.classList.remove('is-open');
				var t=m.previousElementSibling;
				if(t)t.setAttribute('aria-expanded','false');
			});
		}
		body.addEventListener('click',function(e){
			var trigger=e.target.closest('.ms365cal-color-trigger');
			if(trigger){
				var menu=trigger.nextElementSibling;
				var wasOpen=menu.classList.contains('is-open');
				closeAllMenus();
				if(!wasOpen){
					menu.hidden=false;menu.classList.add('is-open');
					trigger.setAttribute('aria-expanded','true');
				}
				return;
			}
			var option=e.target.closest('.ms365cal-color-option');
			if(option){
				setColorPicker(option.closest('.ms365cal-color-picker'),option.dataset.value,option.dataset.name);
				closeAllMenus();
			}
		});
		document.addEventListener('click',function(e){
			if(!e.target.closest('.ms365cal-color-picker'))closeAllMenus();
		});
		document.addEventListener('keydown',function(e){
			if(e.key==='Escape')closeAllMenus();
		});

		// Source picker: the Source field itself is the filter — typing narrows a
		// dropdown of groups/rooms/mailboxes fetched from Graph, but nothing stops
		// the admin from just typing/pasting a GUID or email directly, same as
		// before this existed (the field's value is what gets submitted either way).
		var directoryItems=null,directoryPromise=null,directoryErrors=[];
		var KIND_LABEL={group:'Group',room:'Room',mailbox:'Mailbox'};
		var KIND_ORDER=['group','room','mailbox'];
		function loadDirectory(force){
			if(directoryPromise&&!force)return directoryPromise;
			directoryPromise=fetch(directoryUrl+(force?'?refresh=1':''),{headers:{'X-WP-Nonce':directoryNonce}})
				.then(function(r){return r.json();})
				.then(function(data){
					directoryItems=data.items||[];
					directoryErrors=data.errors||[];
					return directoryItems;
				})
				.catch(function(){
					directoryItems=[];
					directoryErrors=['Request failed.'];
					return directoryItems;
				});
			return directoryPromise;
		}
		function renderSrcMenu(menu,query){
			menu.innerHTML='';
			if(directoryItems===null){
				var loading=document.createElement('div');
				loading.className='ms365cal-src-hint';
				loading.textContent='Loading directory…';
				menu.appendChild(loading);
				return;
			}
			if(directoryErrors.length){
				var err=document.createElement('div');
				err.className='ms365cal-src-error';
				err.textContent="Couldn't load some of the directory: "+directoryErrors.join('; ');
				menu.appendChild(err);
			}
			var q=(query||'').trim().toLowerCase();
			var matches=directoryItems.filter(function(it){
				return !q||it.label.toLowerCase().indexOf(q)!==-1||it.source.toLowerCase().indexOf(q)!==-1;
			});
			if(!matches.length){
				var empty=document.createElement('div');
				empty.className='ms365cal-src-empty';
				empty.textContent=directoryItems.length?'No matches — you can still enter a GUID or email manually.':'Directory unavailable — enter a GUID or email manually.';
				menu.appendChild(empty);
				return;
			}
			var LIMIT=40,shown=0;
			KIND_ORDER.forEach(function(kind){
				if(shown>=LIMIT)return;
				var group=matches.filter(function(it){return it.kind===kind;});
				if(!group.length)return;
				var header=document.createElement('div');
				header.className='ms365cal-src-section';
				header.textContent=KIND_LABEL[kind];
				menu.appendChild(header);
				group.slice(0,LIMIT-shown).forEach(function(it){
					var btn=document.createElement('button');
					btn.type='button';
					btn.className='ms365cal-src-option';
					btn.dataset.source=it.source;
					btn.dataset.label=it.label;
					var name=document.createElement('span');
					name.className='ms365cal-src-name';
					name.textContent=it.label;
					var addr=document.createElement('span');
					addr.className='ms365cal-src-addr';
					addr.textContent=it.source;
					btn.appendChild(name);
					btn.appendChild(addr);
					menu.appendChild(btn);
					shown++;
				});
			});
			if(shown<matches.length){
				var hint=document.createElement('div');
				hint.className='ms365cal-src-hint';
				hint.textContent='Showing first '+shown+' of '+matches.length+' — keep typing to narrow.';
				menu.appendChild(hint);
			}
		}
		function closeAllSrcMenus(){
			document.querySelectorAll('.ms365cal-src-menu.is-open').forEach(function(m){
				m.hidden=true;m.classList.remove('is-open');
			});
		}
		body.addEventListener('focusin',function(e){
			var input=e.target.closest('.ms365cal-src-input');
			if(!input)return;
			var menu=input.closest('.ms365cal-src-picker').querySelector('.ms365cal-src-menu');
			closeAllSrcMenus();
			menu.hidden=false;menu.classList.add('is-open');
			renderSrcMenu(menu,input.value);
			loadDirectory(false).then(function(){
				if(document.activeElement===input)renderSrcMenu(menu,input.value);
			});
		});
		body.addEventListener('input',function(e){
			var input=e.target.closest('.ms365cal-src-input');
			if(!input)return;
			var menu=input.closest('.ms365cal-src-picker').querySelector('.ms365cal-src-menu');
			menu.hidden=false;menu.classList.add('is-open');
			renderSrcMenu(menu,input.value);
			updateRowEnablement(input.closest('tr'));
		});
		body.addEventListener('click',function(e){
			var option=e.target.closest('.ms365cal-src-option');
			if(!option)return;
			var row=option.closest('tr');
			row.querySelector('.ms365cal-src-input').value=option.dataset.source;
			var labelInput=row.querySelector('input[name$="[label]"]');
			if(labelInput)labelInput.value=option.dataset.label;
			updateRowEnablement(row);
			var colorOptions=row.querySelectorAll('.ms365cal-color-option');
			if(colorOptions.length){
				var pick=colorOptions[Math.floor(Math.random()*colorOptions.length)];
				setColorPicker(row.querySelector('.ms365cal-color-picker'),pick.dataset.value,pick.dataset.name);
			}
			closeAllSrcMenus();
		});
		document.addEventListener('click',function(e){
			if(!e.target.closest('.ms365cal-src-picker'))closeAllSrcMenus();
		});
		document.addEventListener('keydown',function(e){
			if(e.key==='Escape')closeAllSrcMenus();
		});
	})();
	(function(){
		var STORAGE_KEY='ms365cal_active_tab';
		var tabs=document.querySelectorAll('.nav-tab-wrapper .nav-tab');
		var panels=document.querySelectorAll('.ms365cal-tab-panel');
		function activate(name){
			tabs.forEach(function(t){t.classList.toggle('nav-tab-active',t.getAttribute('data-tab')===name);});
			panels.forEach(function(p){p.classList.toggle('is-active',p.getAttribute('data-tab')===name);});
		}
		tabs.forEach(function(t){
			t.addEventListener('click',function(e){
				e.preventDefault();
				var name=t.getAttribute('data-tab');
				activate(name);
				try{localStorage.setItem(STORAGE_KEY,name);}catch(e){}
			});
		});
		var initial='calendars';
		try{
			var saved=localStorage.getItem(STORAGE_KEY);
			if(saved&&document.querySelector('.ms365cal-tab-panel[data-tab="'+saved+'"]'))initial=saved;
		}catch(e){}
		activate(initial);
	})();
	</script>
	<?php
}
