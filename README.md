# One-Click Meeting Join

A small Thunderbird add-on prototype. It adds a blue **Join** badge beside events in the Today Pane when it finds a likely meeting URL in the event's description, URL, or location. Clicking the badge opens the link in the system browser.

## Try it in Thunderbird 153

1. Open **Add-ons and Themes**.
2. Open the gear menu and choose **Debug Add-ons**.
3. Choose **Load Temporary Add-on** and select this folder's `manifest.json`.
4. Look at the Today Pane on the right. Add an event with a meeting URL in its description to try the badge.

## Link selection

Teams meeting-join links, Google Meet, Zoom, Webex, GoTo, and Whereby links are recognized. Other providers work when there is only one URL or exactly one URL is labeled with “join” or “attend”. When several unrecognized URLs are ambiguous, the add-on leaves the event alone instead of risking the wrong link, such as a meeting-options URL.

## Compatibility and permissions

This prototype targets Thunderbird 153.6.x, including 153.6.0. It uses a Thunderbird Experiment API to add a control to the built-in Today Pane. Thunderbird grants Experiment APIs **full, unrestricted access** to the application and computer, so it may show a broad permission prompt. Its UI integration also depends on Thunderbird's internal Today Pane markup and may need updates when Thunderbird changes.

