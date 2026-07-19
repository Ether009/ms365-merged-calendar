<?php
/**
 * Plugin Name:       MS365 Merged Calendar (Async)
 * Description:        Merge calendars from Microsoft 365 groups and shared mailboxes into one filterable, windowed list. Events load asynchronously per view via a REST endpoint; prev/next paging with client-side window caching.
 * Version:           2.2.4
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
define( 'MS365CAL_MAX_PAGES', 20 );    // safety cap on Graph nextLink pages per calendar (×100 events)
define( 'MS365CAL_MAX_MASTERS', 200 ); // recurrence series masters resolved per request (via $batch, 20/call)

/**
 * ---------------------------------------------------------------------------
 *  Automatic updates from GitHub (plugin-update-checker)
 * ---------------------------------------------------------------------------
 *  One-click updates in wp-admin, driven by GitHub Releases of the public repo
 *  Ether009/ms365-merged-calendar. The updater library is vendored under
 *  /plugin-update-checker. If it's absent (e.g. a bare single-file install)
 *  this is a no-op: the plugin still runs, it just won't self-update. WordPress
 *  compares each release's plugin "Version:" header against the installed one.
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

	$checker = \YahnisElsts\PluginUpdateChecker\v5\PucFactory::buildUpdateChecker(
		'https://github.com/Ether009/ms365-merged-calendar/',
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
 * ---------------------------------------------------------------------------
 *  Settings
 * ---------------------------------------------------------------------------
 *  Credentials may be defined in wp-config.php (they win over DB options):
 *      define( 'MS365CAL_TENANT_ID', '...' );
 *      define( 'MS365CAL_CLIENT_ID', '...' );
 *      define( 'MS365CAL_CLIENT_SECRET', '...' );
 *      define( 'MS365CAL_DEPLOY_KEY', '...' );  // self-update endpoint; also settable in the UI
 */
function ms365cal_get_settings() {
	$defaults = array(
		'tenant_id'     => '',
		'client_id'     => '',
		'client_secret' => '',
		'cache_minutes' => 20,
		'timezone'      => wp_timezone_string(),
		'rate_max'      => MS365CAL_RATE_MAX,
		'rate_window'   => MS365CAL_RATE_WINDOW,
		'show_outlook'  => false,
		'deploy_key'    => '',
		'calendars'     => array(),
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
 */
function ms365cal_view_url( $cal, $start_iso, $end_iso ) {
	$source = rawurlencode( $cal['source'] );
	$base   = ( 'mailbox' === $cal['type'] )
		? "https://graph.microsoft.com/v1.0/users/{$source}/calendarView"
		: "https://graph.microsoft.com/v1.0/groups/{$source}/calendarView";

	return add_query_arg(
		array(
			'startDateTime' => $start_iso,
			'endDateTime'   => $end_iso,
			'$orderby'      => 'start/dateTime',
			'$top'          => 100,
			'$select'       => 'subject,start,end,location,isAllDay,onlineMeeting,webLink,type,seriesMasterId,body',
		),
		$base
	);
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
function ms365cal_fetch_one( $cal, $token, $start_iso, $end_iso, $tz ) {
	$args = array(
		'timeout' => 20,
		'headers' => array(
			'Authorization' => 'Bearer ' . $token,
			'Prefer'        => 'outlook.timezone="' . $tz . '", outlook.body-content-type="text"',
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
	$url          = ms365cal_view_url( $cal, $start_iso, $end_iso );
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

		$body = json_decode( wp_remote_retrieve_body( $resp ), true );

		if ( ! empty( $body['value'] ) && is_array( $body['value'] ) ) {
			foreach ( $body['value'] as $e ) {
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
				if ( null !== $en && $en > $today0 && $st < $today0 ) {
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

				// Recurrence: calendarView returns expanded occurrences; the pattern
				// lives on the series master. Collect the master IDs now and resolve
				// them together via $batch after paging (see below); until then, mark
				// the row generic. Occurrences/exceptions with no master stay generic.
				$recur     = '';
				$master_id = isset( $e['seriesMasterId'] ) ? $e['seriesMasterId'] : '';
				$etype     = isset( $e['type'] ) ? $e['type'] : '';
				if ( '' !== $master_id ) {
					$recur = 'Återkommande händelse';

					$need_masters[ $master_id ] = true;
				} elseif ( in_array( $etype, array( 'occurrence', 'exception' ), true ) ) {
					$recur = 'Återkommande händelse';
				}

				$is_online = ! empty( $e['onlineMeeting'] );
				$body      = isset( $e['body']['content'] ) ? trim( (string) $e['body']['content'] ) : '';
				if ( $is_online && '' !== $body ) {
					// The join link and "Teams meeting" location are shown separately in
					// the UI, so the auto-inserted join/dial-in block in the body is just
					// noise — strip it, keeping any real notes the organiser wrote.
					$body = ms365cal_strip_meeting_boilerplate( $body );
				}

				$row = array(
					'cal'      => $cal['slug'],
					'title'    => isset( $e['subject'] ) && '' !== $e['subject'] ? $e['subject'] : '(ingen rubrik)',
					'sort'     => $eff->format( 'Y-m-d\TH:i:s' ),
					'dayKey'   => $eff->format( 'Y-m-d' ),
					'dayLabel' => wp_date( 'D j M', $eff->getTimestamp(), $zone ),
					't1'       => $t1,
					't2'       => $t2,
					'when'     => $when,
					'recur'    => $recur,
					'longSpan' => $long_span,
					'body'     => $body,
					'location' => isset( $e['location']['displayName'] ) ? $e['location']['displayName'] : '',
					'online'   => $is_online,
					'joinUrl'  => isset( $e['onlineMeeting']['joinUrl'] ) ? $e['onlineMeeting']['joinUrl'] : '',
					'link'     => isset( $e['webLink'] ) ? $e['webLink'] : '',
				);
				if ( '' !== $master_id ) {
					$row['_master'] = $master_id; // temp key, stripped after resolution.
				}
				$out[] = $row;
			}
		}

		// Follow pagination. The nextLink is an absolute URL carrying the skiptoken
		// plus the original $select/$top, so pass it through unchanged.
		$url = isset( $body['@odata.nextLink'] ) ? $body['@odata.nextLink'] : '';
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
 * Merge + sort + cache a window of events for a set of calendar slugs.
 *
 * @return array|WP_Error
 */
function ms365cal_events_window( $slugs, $start_date, $days ) {
	$settings = ms365cal_get_settings();
	$tz       = $settings['timezone'] ? $settings['timezone'] : 'UTC';
	$days     = min( MS365CAL_MAX_WINDOW, max( 1, (int) $days ) );

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
	if ( $has_stale && ( $now - (int) $cached['fetched'] ) < $fresh ) {
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

	$all         = array();
	$throttled   = false;
	$retry_after = 30;
	foreach ( $wanted as $cal ) {
		$res = ms365cal_fetch_one( $cal, $token, $start_iso, $end_iso, $tz );
		if ( ! empty( $res['throttled'] ) ) {
			// Graph throttles per app/tenant, so once one call is limited the rest
			// will be too — stop and don't pile on more requests.
			$throttled   = true;
			$retry_after = (int) $res['retry_after'];
			break;
		}
		$all = array_merge( $all, $res['events'] );
	}

	if ( $throttled ) {
		$backoff = max( 1, min( $retry_after, MS365CAL_BACKOFF_MAX ) );
		set_transient( 'ms365cal_backoff', $now + $backoff, $backoff );
		if ( $has_stale ) {
			return $cached['events']; // seamless: slightly old data beats a blank calendar
		}
		return new WP_Error( 'ms365cal_throttled', 'The calendar service is busy.', array( 'retry_after' => $backoff ) );
	}

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
 * Drop the cached Graph token, the back-off flag, and all window event caches so
 * the next fetch starts from a blank slate. The debug endpoint calls this to
 * force a brand-new token after a permissions/consent change, rather than
 * reusing a token that was minted before consent (which carries no roles).
 */
function ms365cal_flush_cache() {
	global $wpdb;

	// Exact-key deletes — these work even under a persistent object cache, so the
	// token concern is always covered.
	delete_transient( MS365CAL_TOKEN_TRANSIENT );
	delete_transient( 'ms365cal_backoff' );

	// Window event caches use dynamic keys, so clear them by prefix. Best-effort:
	// on a persistent object cache this DB sweep won't reach cached copies, but
	// the diagnostic reads live anyway and the token above is what matters.
	$like_val = $wpdb->esc_like( '_transient_ms365cal_w_' ) . '%';
	$like_to  = $wpdb->esc_like( '_transient_timeout_ms365cal_w_' ) . '%';
	// phpcs:ignore WordPress.DB.DirectDatabaseQuery.DirectQuery, WordPress.DB.DirectDatabaseQuery.NoCaching
	$wpdb->query( $wpdb->prepare( "DELETE FROM {$wpdb->options} WHERE option_name LIKE %s OR option_name LIKE %s", $like_val, $like_to ) );
}

/**
 * Admin-only diagnostic: fetch each configured calendar live (bypassing the
 * cache and the back-off) and report the real Graph status, so a 403/404 that
 * the normal path swallows becomes visible. Returns array|WP_Error.
 */
function ms365cal_diagnose( $slugs, $start_date, $days ) {
	$settings = ms365cal_get_settings();
	$tz       = $settings['timezone'] ? $settings['timezone'] : 'UTC';
	$days     = min( MS365CAL_MAX_WINDOW, max( 1, (int) $days ) );

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
		$res                = ms365cal_fetch_one( $cal, $token, $start_iso, $end_iso, $tz );
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
}
add_action( 'rest_api_init', 'ms365cal_register_rest' );

/**
 * Secret-authenticated self-update. Forces the update checker to re-check GitHub and,
 * if a newer release exists, installs it via the plugin upgrader — so a release can be
 * pushed to a live site without shell access or a wp-admin click.
 *
 * Guardrails: off unless MS365CAL_DEPLOY_KEY is defined; key sent in the
 * X-MS365CAL-Deploy-Key header and compared with hash_equals(); a short transient lock
 * blocks concurrent/rapid triggers; and the update source is fixed to this plugin's
 * own GitHub repo, so the endpoint can only ever install this repo's latest release —
 * never arbitrary code.
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
 * ---------------------------------------------------------------------------
 *  Colour helper
 * ---------------------------------------------------------------------------
 */
function ms365cal_rgba( $hex, $alpha ) {
	$hex = ltrim( $hex, '#' );
	if ( 3 === strlen( $hex ) ) {
		$hex = $hex[0] . $hex[0] . $hex[1] . $hex[1] . $hex[2] . $hex[2];
	}
	if ( 6 !== strlen( $hex ) ) {
		$hex = '888888';
	}
	return sprintf(
		'rgba(%d,%d,%d,%s)',
		hexdec( substr( $hex, 0, 2 ) ),
		hexdec( substr( $hex, 2, 2 ) ),
		hexdec( substr( $hex, 4, 2 ) ),
		$alpha
	);
}

/**
 * ---------------------------------------------------------------------------
 *  Shortcode — renders the shell only. Events load async via the REST route.
 *  [ms365_calendar calendars="eng,sales"]  (weekly, Monday-aligned)
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
		$meta[ $c['slug'] ]     = array(
			'label' => $c['label'],
			'color' => $c['color'],
			'bg'    => ms365cal_rgba( $c['color'], 0.12 ),
		);
		$defaults[ $c['slug'] ] = ! isset( $c['default'] ) || $c['default'];
	}

	$config = array(
		'rest'        => esc_url_raw( rest_url( 'ms365cal/v1/events' ) ),
		'cals'        => wp_list_pluck( $scoped, 'slug' ),
		'meta'        => $meta,
		'defaults'    => $defaults,
		'windowDays'  => $window,
		'startY'      => (int) $today->format( 'Y' ),
		'startM'      => (int) $today->format( 'n' ),
		'startD'      => (int) $today->format( 'j' ),
		'showOutlook' => ! empty( $settings['show_outlook'] ),
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

		<div class="ms365cal-nav">
			<button type="button" class="ms365cal-page" data-dir="-1" aria-label="Föregående vecka">&larr;</button>
			<span class="ms365cal-range"></span>
			<button type="button" class="ms365cal-page" data-dir="1" aria-label="Nästa vecka">&rarr;</button>
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
	.ms365cal-chip{display:inline-flex;align-items:center;gap:7px;font-size:13px;padding:6px 13px;border-radius:999px;cursor:pointer;border:1px solid var(--ms-line);background:transparent;color:inherit;opacity:.5;transition:opacity .15s,border-color .15s,background .15s;}
	.ms365cal-chip:hover{opacity:.8;}
	.ms365cal-chip.is-on{border-color:var(--cc);background:var(--cbg);opacity:1;}
	.ms365cal-dot{width:9px;height:9px;border-radius:50%;background:currentColor;opacity:.35;}
	.ms365cal-chip.is-on .ms365cal-dot{background:var(--cc);opacity:1;}
	.ms365cal-nav{display:flex;align-items:center;gap:14px;margin-bottom:6px;}
	.ms365cal-page{width:36px;height:36px;border:1px solid var(--ms-line);border-radius:10px;background:transparent;cursor:pointer;font-size:16px;line-height:1;color:inherit;opacity:.85;transition:background .15s,opacity .15s;}
	.ms365cal-page:hover{background:var(--ms-soft);opacity:1;}
	.ms365cal-page:disabled{opacity:.3;cursor:default;}
	.ms365cal-page:disabled:hover{background:transparent;}
	.ms365cal-range{font-size:14px;font-weight:600;opacity:.7;}
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
	.ms365cal-cat{font-size:11px;font-weight:600;line-height:1.2;overflow-wrap:break-word;max-width:100%;padding:2px 0;}
	.ms365cal-rail{width:4px;border-radius:999px;flex:0 0 auto;margin:10px 0;}
	.ms365cal-hbody{flex:1;min-width:0;padding:10px 0;display:flex;flex-direction:column;}
	.ms365cal-ev{font-size:15px;font-weight:600;background:none;border:0;padding:0;margin:0;text-align:left;cursor:pointer;color:inherit;font-family:inherit;display:flex;align-items:baseline;gap:9px;width:100%;line-height:1.35;}
	.ms365cal-ev:hover .ms365cal-title{opacity:.65;}
	.ms365cal-title{transition:opacity .12s;}
	.ms365cal-caret{display:inline-block;flex:0 0 auto;font-size:9px;opacity:.4;transition:transform .15s;}
	.ms365cal-meta-line{margin-top:auto;padding-top:8px;display:flex;align-items:center;gap:10px;font-size:12px;opacity:.6;}
	.ms365cal-loc-line{flex:1 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
	.ms365cal-recur-line{flex:0 0 auto;margin-left:auto;white-space:nowrap;}
	.ms365cal-detail{margin:6px 0 4px;font-size:13px;line-height:1.6;}
	.ms365cal-detail div{margin:5px 0;}
	.ms365cal-detail div:first-child{margin-top:0;}
	.ms365cal-desc{white-space:pre-wrap;opacity:.9;}
	.ms365cal-detail a{color:#185fa5;font-weight:600;text-decoration:underline;text-underline-offset:2px;}
	</style>
	<script>
	(function(){
	function esc(s){var d=document.createElement('div');d.textContent=s==null?'':s;return d.innerHTML;}
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

		var cache = {};                                   // window key -> events[]
		// Monday of the week containing d (weeks start Monday).
		function weekStart(d){var x=new Date(d.getFullYear(),d.getMonth(),d.getDate());x.setDate(x.getDate()-((x.getDay()+6)%7));return x;}
		var today    = new Date(cfg.startY, cfg.startM-1, cfg.startD); // today, site tz
		var todayKey = iso(today);
		var start    = weekStart(today);
		var minStart = weekStart(today);                  // current week is the earliest
		var days  = 7;                                    // weekly window
		var reqId = 0;
		var retryTimer = null;
		var throttleRetries = 0;
		var MAX_THROTTLE_RETRIES = 3;
		var openDetail = null;
		var autoAdvanceChecked = false;                   // only skip an empty first week once

		function winKey(){return iso(start)+'|'+days;}
		function updateRange(){
			var end=new Date(start);end.setDate(end.getDate()+days-1);
			rangeEl.textContent='Vecka '+isoWeek(start)+' \u00b7 '+fmt(start)+' \u2013 '+fmt(end);
		}
		function updateNav(){
			var prev=root.querySelector('.ms365cal-page[data-dir="-1"]');
			if(prev)prev.disabled=(start<=minStart);
		}

		function renderChips(){
			chipsEl.innerHTML='';
			cfg.cals.forEach(function(slug){
				var m=cfg.meta[slug],on=enabled[slug];
				var b=document.createElement('button');
				b.type='button';
				b.className='ms365cal-chip'+(on?' is-on':'');
				b.style.setProperty('--cc',m.color);
				b.style.setProperty('--cbg',m.bg);
				b.innerHTML='<span class="ms365cal-dot"></span>'+esc(m.label);
				b.addEventListener('click',function(){enabled[slug]=!enabled[slug];b.classList.toggle('is-on');paint();});
				chipsEl.appendChild(b);
			});
		}

		function renderDays(list){
			var html='',lastDay='';
			list.forEach(function(e){
				var m=cfg.meta[e.cal];if(!m)return;
				if(e.dayKey!==lastDay){html+='<div class="ms365cal-day">'+esc(e.dayLabel)+'</div>';lastDay=e.dayKey;}

				// Bottom line: location on the left, recurrence pinned to the right (via
				// margin-left:auto on the recurrence span, so it stays right-aligned
				// whether or not a location is present).
				var recurShort=e.recur?e.recur.replace(/^Upprepas\s+/,''):'';
				var locText=e.location?e.location:(e.online?'Online':'');
				var metaBits='';
				if(locText)metaBits+='<span class="ms365cal-loc-line">'+esc(locText)+'</span>';
				if(e.recur)metaBits+='<span class="ms365cal-recur-line">\u21bb '+esc(recurShort)+'</span>';
				var metaLine=metaBits?'<div class="ms365cal-meta-line">'+metaBits+'</div>':'';

				// No "when" line here \u2014 the start/end times stay visible in the left
				// column (which stretches with the row) while expanded, so repeating
				// them in the body would be redundant. No location either \u2014 it's on
				// the always-visible meta line above.
				var d='';
				if(e.body)d+='<div class="ms365cal-desc">'+esc(e.body)+'</div>';
				if(e.joinUrl)d+='<div><a href="'+esc(e.joinUrl)+'" target="_blank" rel="noopener">Anslut till onlinem\u00f6te</a></div>';
				else if(e.online)d+='<div>Onlinem\u00f6te</div>';
				if(cfg.showOutlook&&e.link)d+='<div><a href="'+esc(e.link)+'" target="_blank" rel="noopener">\u00d6ppna i Outlook \u2197</a></div>';

				html+='<div class="ms365cal-row">'
					+'<div class="ms365cal-head">'
						+'<div class="ms365cal-times">'
							+'<span class="ms365cal-t1">'+esc(e.t1)+'</span>'
							+'<span class="ms365cal-cat" style="color:'+m.color+'">'+esc(m.label)+'</span>'
							+'<span class="ms365cal-t2">'+esc(e.t2)+'</span>'
						+'</div>'
						+'<div class="ms365cal-rail" style="background:'+m.color+'"></div>'
						+'<div class="ms365cal-hbody">'
							+'<button type="button" class="ms365cal-ev" aria-expanded="false">'
								+'<span class="ms365cal-title">'+esc(e.title)+'</span>'
							+'</button>'
							+(d?'<div class="ms365cal-detail" hidden>'+d+'</div>':'')
							+metaLine
						+'</div>'
					+'</div>'
				+'</div>';
			});
			return html;
		}

		function paint(){
			openDetail=null;
			var events=cache[winKey()]||[];
			var visible=events.filter(function(e){return enabled[e.cal];});
			if(!visible.length){
				listEl.innerHTML='<p class="ms365cal-empty">Inga h\u00e4ndelser den h\u00e4r veckan.</p>';
				return;
			}
			var past=[],upcoming=[];
			visible.forEach(function(e){(e.dayKey<todayKey?past:upcoming).push(e);});

			var html='';
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

		// Expand/collapse: accordion for events (one open at a time); independent
		// toggle for the "Tidigare Händelser" group.
		listEl.addEventListener('click',function(ev){
			var tog=ev.target.closest('.ms365cal-earlier-tog');
			if(tog&&listEl.contains(tog)){
				var eb=tog.parentNode.querySelector('.ms365cal-earlier-body');
				var op=tog.getAttribute('aria-expanded')==='true';
				tog.setAttribute('aria-expanded',op?'false':'true');
				if(eb)eb.hidden=op;
				return;
			}
			var btn=ev.target.closest('.ms365cal-ev');
			if(!btn||!listEl.contains(btn))return;
			var row=btn.closest('.ms365cal-row');
			var detail=row?row.querySelector('.ms365cal-detail'):null;
			if(!detail)return;
			if(openDetail===detail){
				detail.hidden=true;btn.setAttribute('aria-expanded','false');openDetail=null;return;
			}
			if(openDetail){
				openDetail.hidden=true;
				var prow=openDetail.closest('.ms365cal-row');
				var prev=prow?prow.querySelector('.ms365cal-ev'):null;
				if(prev)prev.setAttribute('aria-expanded','false');
			}
			detail.hidden=false;btn.setAttribute('aria-expanded','true');openDetail=detail;
		});

		function load(){
			if(retryTimer){clearTimeout(retryTimer);retryTimer=null;}
			updateRange();updateNav();
			var key=winKey();
			if(cache[key]){paint();return;}          // client-side window cache hit

			var my=++reqId;
			openDetail=null;
			listEl.classList.remove('is-loading');
			// Clear the previous week immediately and show a loading indicator, rather
			// than leaving stale events on screen until the new window arrives.
			listEl.innerHTML='<div class="ms365cal-banner">Laddar\u2026</div>';

			var url=cfg.rest+'?cals='+encodeURIComponent(cfg.cals.join(','))
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
					// as part of a longer span (their real end is past this week) — default
					// to next week instead, and treat it as the new earliest week.
					if(!autoAdvanceChecked&&iso(start)===iso(minStart)){
						autoAdvanceChecked=true;
						var hasUpcoming=cache[key].some(function(e){
							return e.dayKey>=todayKey&&enabled[e.cal]&&!e.longSpan;
						});
						if(!hasUpcoming){
							start=new Date(start);start.setDate(start.getDate()+days);
							minStart=new Date(start);
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

		var COOLDOWN=600; // ms
		var pageBtns=root.querySelectorAll('.ms365cal-page');
		function setPaging(disabled){pageBtns.forEach(function(b){b.disabled=disabled;});}

		pageBtns.forEach(function(btn){
			btn.addEventListener('click',function(){
				if(btn.disabled)return;
				var dir=parseInt(btn.getAttribute('data-dir'),10);
				var ns=new Date(start);ns.setDate(ns.getDate()+dir*days);
				if(ns<minStart)ns=new Date(minStart);   // don't page before the current week
				if(iso(ns)===iso(start))return;          // already at the bound
				start=ns;
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
		$new['cache_minutes'] = max( 1, (int) ( $_POST['cache_minutes'] ?? 20 ) );
		$new['timezone']      = sanitize_text_field( wp_unslash( $_POST['timezone'] ?? 'UTC' ) );
		$new['rate_max']      = max( 0, (int) ( $_POST['rate_max'] ?? MS365CAL_RATE_MAX ) );
		$new['rate_window']   = max( 1, (int) ( $_POST['rate_window'] ?? MS365CAL_RATE_WINDOW ) );
		$new['show_outlook']  = ! empty( $_POST['show_outlook'] );

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

		$cals = array();
		$rows = isset( $_POST['cal'] ) && is_array( $_POST['cal'] ) ? $_POST['cal'] : array();
		foreach ( $rows as $row ) {
			$slug = sanitize_title( $row['slug'] ?? '' );
			$src  = trim( sanitize_text_field( wp_unslash( $row['source'] ?? '' ) ) );
			if ( '' === $slug || '' === $src ) {
				continue;
			}
			$color = sanitize_hex_color( $row['color'] ?? '#378add' );
			if ( ! $color ) {
				$color = '#378add';
			}
			$cals[] = array(
				'slug'    => $slug,
				'label'   => sanitize_text_field( wp_unslash( $row['label'] ?? $slug ) ),
				'color'   => $color,
				'type'    => ( 'mailbox' === ( $row['type'] ?? '' ) ) ? 'mailbox' : 'group',
				'source'  => $src,
				'default' => ! empty( $row['default'] ),
			);
		}
		$new['calendars'] = $cals;

		update_option( MS365CAL_OPTION, $new );
		delete_transient( MS365CAL_TOKEN_TRANSIENT );
		echo '<div class="notice notice-success"><p>Settings saved.</p></div>';
	}

	$s            = ms365cal_get_settings();
	$has_secret   = (bool) ( $s['client_secret'] || ( defined( 'MS365CAL_CLIENT_SECRET' ) && MS365CAL_CLIENT_SECRET ) );
	$deploy_const = defined( 'MS365CAL_DEPLOY_KEY' ) && MS365CAL_DEPLOY_KEY;
	$has_deploy   = (bool) ( $s['deploy_key'] || $deploy_const );
	?>
	<div class="wrap">
		<h1>MS365 Merged Calendar</h1>
		<form method="post">
			<?php wp_nonce_field( 'ms365cal_save', 'ms365cal_nonce' ); ?>

			<h2>Azure app registration</h2>
			<p>Register an app in Entra ID, grant <code>Application</code> permissions
			<code>Calendars.Read</code> (shared mailboxes) and <code>Group.Read.All</code>
			(group calendars), then grant admin consent and add a client secret.</p>
			<table class="form-table">
				<tr><th>Tenant ID</th><td><input type="text" name="tenant_id" class="regular-text" value="<?php echo esc_attr( $s['tenant_id'] ); ?>"></td></tr>
				<tr><th>Client ID</th><td><input type="text" name="client_id" class="regular-text" value="<?php echo esc_attr( $s['client_id'] ); ?>"></td></tr>
				<tr><th>Client secret</th><td>
					<input type="password" name="client_secret" class="regular-text" value="" placeholder="<?php echo $has_secret ? '&bull;&bull;&bull;&bull;&bull;&bull; (leave blank to keep)' : 'Enter secret'; ?>">
					<p class="description">Stored in the database. For better security, define <code>MS365CAL_CLIENT_SECRET</code> in wp-config.php instead.</p>
				</td></tr>
				<tr><th>Cache (minutes)</th><td><input type="number" name="cache_minutes" min="1" value="<?php echo esc_attr( $s['cache_minutes'] ); ?>"></td></tr>
				<tr><th>Timezone</th><td><input type="text" name="timezone" class="regular-text" value="<?php echo esc_attr( $s['timezone'] ); ?>"> <span class="description">e.g. Europe/Stockholm</span></td></tr>
				<tr><th>Rate limit</th><td>
					<input type="number" name="rate_max" min="0" value="<?php echo esc_attr( $s['rate_max'] ); ?>" style="width:6em">
					requests per
					<input type="number" name="rate_window" min="1" value="<?php echo esc_attr( $s['rate_window'] ); ?>" style="width:6em">
					seconds, per IP.
					<p class="description">Applies to the public events endpoint. Set requests to <code>0</code> to disable the limiter. Cached views don't count toward it.</p>
				</td></tr>
				<tr><th>Outlook link</th><td>
					<label><input type="checkbox" name="show_outlook" <?php checked( ! empty( $s['show_outlook'] ) ); ?>> Show an &ldquo;Open in Outlook&rdquo; link in the expanded event details</label>
					<p class="description">Off by default. When enabled, each expanded event includes a link to open it in Outlook on the web.</p>
				</td></tr>
				<tr><th>Deploy key</th><td>
					<?php if ( $deploy_const ) : ?>
						<p class="description"><strong>Set via <code>MS365CAL_DEPLOY_KEY</code> in wp-config.php</strong>, which overrides this field.</p>
					<?php else : ?>
						<input type="password" name="deploy_key" class="regular-text" value="" placeholder="<?php echo $has_deploy ? '&bull;&bull;&bull;&bull;&bull;&bull; (leave blank to keep)' : 'Paste a long random secret'; ?>">
						<?php if ( $has_deploy ) : ?>
							<label style="margin-left:8px"><input type="checkbox" name="deploy_key_clear" value="1"> Clear (disable endpoint)</label>
						<?php endif; ?>
					<?php endif; ?>
					<p class="description">
						Enables <code>POST <?php echo esc_html( rest_url( 'ms365cal/v1/self-update' ) ); ?></code>
						(header <code>X-MS365CAL-Deploy-Key</code>), which makes this site install the plugin's
						latest GitHub release on demand. Leave blank to keep the endpoint disabled. Anyone with
						the key can trigger a reinstall, so use a long random value and rotate it if it leaks.
						Defining <code>MS365CAL_DEPLOY_KEY</code> in wp-config.php is more secure and takes precedence.
					</p>
				</td></tr>
			</table>

			<h2>Calendars</h2>
			<p>For a <strong>group</strong>, the source is the group's object ID (GUID).
			For a <strong>shared mailbox</strong>, the source is its email address.</p>
			<table class="widefat" id="ms365cal-rows">
				<thead><tr>
					<th>Label</th><th>Slug</th><th>Type</th><th>Source (GUID or email)</th><th>Colour</th><th>On by default</th><th></th>
				</tr></thead>
				<tbody>
				<?php
				$rows = ! empty( $s['calendars'] ) ? $s['calendars'] : array(
					array(
						'slug'    => '',
						'label'   => '',
						'color'   => '#378add',
						'type'    => 'group',
						'source'  => '',
						'default' => true,
					),
				);
				foreach ( $rows as $i => $c ) :
					?>
					<tr>
						<td><input type="text" name="cal[<?php echo (int) $i; ?>][label]" value="<?php echo esc_attr( $c['label'] ); ?>"></td>
						<td><input type="text" name="cal[<?php echo (int) $i; ?>][slug]" value="<?php echo esc_attr( $c['slug'] ); ?>"></td>
						<td>
							<select name="cal[<?php echo (int) $i; ?>][type]">
								<option value="group"   <?php selected( $c['type'], 'group' ); ?>>Group</option>
								<option value="mailbox" <?php selected( $c['type'], 'mailbox' ); ?>>Shared mailbox</option>
							</select>
						</td>
						<td><input type="text" name="cal[<?php echo (int) $i; ?>][source]" class="regular-text" value="<?php echo esc_attr( $c['source'] ); ?>"></td>
						<td><input type="text" name="cal[<?php echo (int) $i; ?>][color]" value="<?php echo esc_attr( $c['color'] ); ?>" size="8"></td>
						<td style="text-align:center"><input type="checkbox" name="cal[<?php echo (int) $i; ?>][default]" <?php checked( ! empty( $c['default'] ) ); ?>></td>
						<td><button type="button" class="button ms365cal-del">Remove</button></td>
					</tr>
				<?php endforeach; ?>
				</tbody>
			</table>
			<p><button type="button" class="button" id="ms365cal-add">Add calendar</button></p>

			<?php submit_button(); ?>
		</form>

		<h2>Usage</h2>
		<p><code>[ms365_calendar]</code> &mdash; all calendars, current week (Monday&ndash;Sunday).<br>
		<code>[ms365_calendar calendars="eng,events"]</code> &mdash; specific slugs. Navigate forward week by week; the current week is the earliest.</p>

		<h2>Troubleshooting</h2>
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
		<p><a href="<?php echo esc_url( $debug_url ); ?>" target="_blank" rel="noopener" class="button">Run calendar diagnostic</a></p>
		<p class="description">
			Opens a JSON report of each calendar's live Graph status (HTTP code, event count, and any error).
			A <code>403</code> means missing permission or admin consent; <code>404</code> means a wrong source ID
			or type; <code>200</code> with <code>events: 0</code> means the credentials and IDs are fine, so it's the
			date window. The link above already carries the security nonce that REST cookie-auth requires &mdash;
			typing the URL by hand without it returns the normal (public) empty response instead. The nonce expires
			after a while, so reload this settings page to get a fresh link. Only administrators can see the report.
		</p>
	</div>

	<script>
	(function(){
		var body=document.querySelector('#ms365cal-rows tbody');
		var idx=body.querySelectorAll('tr').length;
		document.getElementById('ms365cal-add').addEventListener('click',function(){
			var tr=body.querySelector('tr').cloneNode(true);
			tr.querySelectorAll('input,select').forEach(function(el){
				el.name=el.name.replace(/cal\[\d+\]/,'cal['+idx+']');
				if(el.type==='checkbox')el.checked=false;else if(el.tagName==='SELECT')el.selectedIndex=0;else el.value=(el.name.indexOf('color')>-1?'#378add':'');
			});
			body.appendChild(tr);idx++;
		});
		body.addEventListener('click',function(e){
			if(e.target.classList.contains('ms365cal-del')){
				if(body.querySelectorAll('tr').length>1)e.target.closest('tr').remove();
			}
		});
	})();
	</script>
	<?php
}
