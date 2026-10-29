"use strict";

const { ExtensionSupport } = ChromeUtils.importESModule(
  "resource:///modules/ExtensionSupport.sys.mjs"
);
const Services = globalThis.Services;

const MAIN_WINDOW = "chrome://messenger/content/messenger.xhtml";
const STYLE_ID = "one-click-meeting-join-style";
const CLEANUP_KEY = "__oneClickMeetingJoinCleanup";
const LOG_PREFIX = "[One-Click Meeting Join]";

function getItemSources(item) {
  if (!item) {
    return [];
  }

  return [
    "DESCRIPTION",
    "X-ALT-DESC",
    "URL",
    "LOCATION",
    "CONFERENCE",
    "X-GOOGLE-CONFERENCE",
    "X-MICROSOFT-SKYPETEAMSMEETINGURL",
  ]
    .map(property => {
      try {
        return item[property.toLowerCase()] || item.getProperty?.(property);
      } catch {
        return "";
      }
    })
    .filter(value => typeof value === "string" && value);
}

function findMeetingLink(sources) {
  const links = new Map();

  for (const source of sources) {
    for (const match of source.matchAll(/https?:\/\/[^\s<>"']+/gi)) {
      const raw = match[0]
        .replace(/&amp;/gi, "&")
        .replace(/&gt;$/i, "")
        .replace(/[),.;!?]+$/g, "");
      try {
        const url = new URL(raw);
        if (url.protocol !== "https:" && url.protocol !== "http:") {
          continue;
        }

        const end = match.index + match[0].length;
        const previousUrls = [...source.slice(0, match.index).matchAll(/https?:\/\/[^\s<>"']+/gi)];
        const previousUrl = previousUrls.at(-1);
        const nextUrl = /https?:\/\/[^\s<>"']+/gi.exec(source.slice(end));
        const contextStart = Math.max(
          previousUrl ? previousUrl.index + previousUrl[0].length : 0,
          match.index - 160
        );
        const contextEnd = nextUrl ? end + nextUrl.index : Math.min(source.length, end + 160);
        const linkContext = `${source.slice(contextStart, match.index)} ${source.slice(end, contextEnd)}`
          .replace(/<[^>]*>/g, " ");
        const previous = links.get(url.href);
        links.set(url.href, {
          url: url.href,
          labeledJoin: Boolean(previous?.labeledJoin || /\b(join|attend)\b/i.test(linkContext)),
        });
      } catch {
        // Ignore malformed invitation URLs.
      }
    }
  }

  const candidates = [...links.values()];
  const known = candidates.filter(link => isKnownMeetingLink(link.url));
  if (known.length === 1) {
    return known[0].url;
  }

  const labeled = candidates.filter(link => link.labeledJoin);
  if (labeled.length === 1) {
    return labeled[0].url;
  }

  return candidates.length === 1 ? candidates[0].url : null;
}

function isKnownMeetingLink(value) {
  const url = new URL(value);
  const host = url.hostname.toLowerCase();
  const path = url.pathname.toLowerCase();

  if (host === "teams.microsoft.com" || host.endsWith(".teams.microsoft.com")) {
    return path.includes("/l/meetup-join/") || path.startsWith("/meet/");
  }
  if (host === "teams.live.com") {
    return path.startsWith("/meet/");
  }
  if (host === "meet.google.com") {
    return true;
  }
  if (host === "zoom.us" || host.endsWith(".zoom.us")) {
    return /^\/(j|my)\//.test(path);
  }
  if (host === "webex.com" || host.endsWith(".webex.com")) {
    return /\/(meet|join)\//.test(path);
  }
  return host === "meet.goto.com" || host === "gotomeet.me" || host === "whereby.com";
}

function attachToWindow(window) {
  const document = window.document;
  if (window[CLEANUP_KEY]) {
    return;
  }

  let agendaObserver;
  let rowObserver;
  let style;
  let loggedFirstScan = false;

  const cleanup = () => {
    agendaObserver?.disconnect();
    rowObserver?.disconnect();
    document.querySelectorAll("button[data-one-click-join]").forEach(button => button.remove());
    style?.remove();
    window.removeEventListener("unload", cleanup);
    delete window[CLEANUP_KEY];
  };
  window[CLEANUP_KEY] = cleanup;
  window.addEventListener("unload", cleanup, { once: true });

  const installOnAgenda = () => {
    const agenda = document.getElementById("agenda");
    if (!agenda) {
      return false;
    }

    agendaObserver?.disconnect();
    style = document.createElement("style");
    style.id = STYLE_ID;
    style.textContent = `
    .one-click-meeting-join {
      margin-inline-start: 0.35em;
      padding: 1px 6px;
      border: 0;
      border-radius: 1em;
      background: #1683e9;
      color: white;
      cursor: pointer;
      font: inherit;
      font-size: 0.8em;
      white-space: nowrap;
    }
    .one-click-meeting-join:focus-visible {
      outline: 2px solid currentColor;
      outline-offset: 2px;
    }
  `;
    document.documentElement.appendChild(style);

    const addBadges = () => {
      const rows = agenda.querySelectorAll("li.agenda-listitem:not(.agenda-listitem-end)");
      let items = 0;
      let itemsWithText = 0;
      let urls = 0;
      let matches = 0;
      let added = 0;

      for (const row of rows) {
        if (row.querySelector("button[data-one-click-join]")) {
          continue;
        }

        const item = row.item;
        if (!item) {
          continue;
        }
        items++;

        const sources = getItemSources(item);
        if (sources.length) {
          itemsWithText++;
        }
        urls += sources.reduce(
          (count, source) => count + [...source.matchAll(/https?:\/\/[^\s<>"']+/gi)].length,
          0
        );
        const url = findMeetingLink(sources);
        if (!url) {
          continue;
        }
        matches++;

        const anchor = row.querySelector(".agenda-listitem-relative, .agenda-listitem-title");
        if (!anchor) {
          continue;
        }

        const button = document.createElement("button");
        button.type = "button";
        button.className = "one-click-meeting-join";
        button.dataset.oneClickJoin = "true";
        button.textContent = "Join";
        button.setAttribute("aria-label", "Join meeting");
        button.title = "Join meeting";
        button.addEventListener("click", event => {
          event.preventDefault();
          event.stopPropagation();
          Services.externalProtocolService.loadURI(Services.io.newURI(url));
        });
        button.addEventListener("dblclick", event => event.stopPropagation());
        anchor.after(button);
        added++;
      }

      if (!loggedFirstScan && rows.length) {
        console.info(
          LOG_PREFIX,
          `Today Pane scan: ${rows.length} rows, ${items} items, ${itemsWithText} with link fields, ${urls} URLs, ${matches} meeting links, ${added} badges.`
        );
        loggedFirstScan = true;
      }
    };

    rowObserver = new window.MutationObserver(addBadges);
    rowObserver.observe(agenda, { childList: true, subtree: true });
    addBadges();
    console.info(LOG_PREFIX, "Today Pane agenda found; watching event rows.");
    return true;
  };

  console.info(LOG_PREFIX, `Checking Today Pane in ${window.location.href}.`);
  if (!installOnAgenda()) {
    console.info(LOG_PREFIX, "Agenda not present yet; waiting for it to appear.");
    agendaObserver = new window.MutationObserver(() => installOnAgenda());
    agendaObserver.observe(document.documentElement, { childList: true, subtree: true });
  }
}

class MeetingJoin extends ExtensionCommon.ExtensionAPI {
  onStartup() {
    console.info(LOG_PREFIX, "Experiment startup.");
    this.listenerId = this.extension.id;
    this.windowListener = {
      chromeURLs: [MAIN_WINDOW],
      onLoadWindow: attachToWindow,
    };

    ExtensionSupport.registerWindowListener(this.listenerId, this.windowListener);
    for (const window of ExtensionSupport.openWindows) {
      if (window.location.href.startsWith(MAIN_WINDOW)) {
        attachToWindow(window);
      }
    }
  }

  getAPI() {
    return {};
  }

  onShutdown(isAppShutdown) {
    if (!this.listenerId) {
      return;
    }

    for (const window of ExtensionSupport.openWindows) {
      if (window.location.href.startsWith(MAIN_WINDOW)) {
        window[CLEANUP_KEY]?.();
      }
    }
    ExtensionSupport.unregisterWindowListener(this.listenerId);

    if (!isAppShutdown) {
      Services.obs.notifyObservers(null, "startupcache-invalidate", null);
    }
  }
}

this.MeetingJoin = MeetingJoin;

