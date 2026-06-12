# MC Pregame Google Apps Script

This project manages room booking calendars using Google Apps Script.

## Features

- **Calendar Integration**: Automatically creates calendar events based on room bookings
- **Room Management**: Handles multiple room bookings and availability
- **Dynamic Form Validation**: Conditional dropdowns based on room selections
- **Email Notifications**: Sends booking confirmations and updates

## Files

- `Code.gs`: Main calendar management functions
- `ColumnsEdit.gs`: Handles spreadsheet editing and validation logic
- `appsscript.json`: Project configuration

## Setup

1. Install clasp globally:

   ```bash
   npm install -g @google/clasp
   ```

2. Login to your Google account:

   ```bash
   clasp login
   ```

3. Install dependencies:
   ```bash
   npm install
   ```

## Development

### Push local changes to Google Apps Script:

```bash
npm run push
# or
clasp push
```

### Pull changes from Google Apps Script:

```bash
npm run pull
# or
clasp pull
```

### Open the project in Google Apps Script editor:

```bash
npm run open
# or
clasp open
```

### Watch for changes and auto-push:

```bash
npm run watch
# or
clasp push --watch
```

### Deploy the project:

```bash
npm run deploy
# or
clasp deploy
```

### View logs:

```bash
npm run logs
# or
clasp tail-logs
```

## Project Structure

- Room booking data is managed through Google Sheets
- Calendar events are created in draft mode first, then pushed to main calendars
- The system supports recurring events and multi-room bookings
- Email notifications are sent to booking requesters

## Configuration

The project is configured for:

- **Timezone**: America/New_York
- **Runtime**: V8
- **Logging**: Stackdriver

## Access Control

- Administrative functions are restricted to authorized users
- Menu items appear only for authorized email addresses