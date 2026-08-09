# LINE WORKS important notice notification

This add-on sends the Portal "重要事項共有専用" text to LINE WORKS and, when configured, normal LINE groups when a post is created or updated.

It does not use or overwrite existing LINE WORKS settings. Add these as Script Properties in the Wonder+ Portal Apps Script project:

- `WONDER_PORTAL_LINEWORKS_BOT_ID`  
  Current Bot ID: `12897792`. If this property is blank, the script uses `12897792`.
- `WONDER_PORTAL_LINEWORKS_CLIENT_ID`
- `WONDER_PORTAL_LINEWORKS_CLIENT_SECRET`
- `WONDER_PORTAL_LINEWORKS_SERVICE_ACCOUNT`
- `WONDER_PORTAL_LINEWORKS_PRIVATE_KEY`
- `WONDER_PORTAL_LINEWORKS_CHANNEL_IDS`

`WONDER_PORTAL_LINEWORKS_CHANNEL_IDS` can be comma-separated. The current targets are:

```text
b0c28afd-26dc-55c3-531d-564ee5088d56
86be568e-ca15-b86a-3623-fbcbabe68431
4bec9198-ca7b-346c-092f-7d1e2d87eaf8
```

The bot must be installed or invited into each target talk room. If LINE WORKS sending fails, the Portal post itself is still saved; the error is written to Apps Script logs.

## Normal LINE group notification

Normal LINE groups use LINE Messaging API, not LINE WORKS API. Add these Script Properties only for the Wonder+ Portal project:

- `WONDER_PORTAL_LINE_CHANNEL_ACCESS_TOKEN`
- `WONDER_PORTAL_LINE_TO_IDS`

`WONDER_PORTAL_LINE_TO_IDS` can be comma-separated. Use LINE group IDs, room IDs, or user IDs that the LINE official account bot can push to.

Required LINE setup:

1. Create or use a LINE official account for Wonder+ Portal.
2. Enable Messaging API and issue a channel access token.
3. Add the official account bot to the target LINE group.
4. Obtain the group ID from webhook events after a message is posted in that group, then put it into `WONDER_PORTAL_LINE_TO_IDS`.

If LINE sending fails, the Portal post itself is still saved; the error is written to Apps Script logs.

## Test functions

- `testImportantNoticeLineWorks()`
- `testImportantNoticeLine()`

After updating `board_api_addon.gs`, deploy a new web app version.
