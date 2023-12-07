# google-sheets-strava

## Overview

This is automation that gives you a button to export Strava activities to Google Sheets, specifically to a sheet titled `raw_data`. What you do with that afterwards is up to you.

I personally have the below set up:
- Google Data Studio Dashboard @ [cycling.agmsb.xyz](https://cycling.agmsb.xyz)
- `monthly_data` roll-up sheet
- `yearly_data` roll-up sheet
- `gear_data` roll-up sheet
- `friends_data` roll-up sheet

## Setting up `raw_data`

Create a sheet titled `raw_data` and add the below columns starting with column A:
- Date	
- Name	
- Duration	
- Distance	
- Elevation	
- Average Speed	
- Max Speed	
- Kilojoules	
- Average Heart Rate	
- Max Heart Rate	
- Gear ID	
- Athlete Count	

In `Tools > Script Editor`, update `code.js` and create a new file:
- `oauth.js`

Update `oauth.js` with your Strava client id and client secret.

Install the OAuth2 library for Google Apps Script, see [setup here](https://github.com/googleworkspace/apps-script-oauth2#setup). In your first time running your code, run the function `getRides` in `Code.js` from `Tools > Script Editor` - it will fail and prompt you to authorize in browser with a provided URL. 

## Setting up `monthly_data` or `yearly_data`

Example: roll-up rides by month
```
// A2 being month formatted as YEAR-MO aka "XXXX-YY"
=COUNTIF(raw_data!A:A, "*"&A2&"*")
```

Example: roll-up mileage by month
```
// A2 being month formatted as YEAR-MO aka "XXXX-YY"
=SUMIF(raw_data!A:A, "*"&A2&"*", raw_data!D:D)
```

## Setting up `gear_data`

It ain't pretty but Strava creates unique IDs for your gear. 

Example: roll-up mileage by gear. 

```
// change gear ID in 2nd arg
=SUMIF(raw_data!K:K, "b7595586", raw_data!D:D)
```

## Setting up `friends_data`

In `raw_data`, I append 3 columns at the end of the sheet:
- `Friends`
- `Lifetime Distance`
- `Lifetime Elevation`

In `Friends`, I manually insert the friends I rode as `First Name` + ` ` + `Last Name`, with each friend separated by comma. 

Then in the sheet `friends_data`, you can roll-up by data by friend using the below examples:

Example: roll-up rides by friend
```
// A2 being Friend Name
=COUNTIF(raw_data!M:M, "*"&A2&"*")
```

Example: roll-up mileage by friend
```
// A2 being Friend Name
=SUMIF(raw_data!M:M, "*"&A2&"*", raw_data!D:D)
```

## TODO

- Add tutorial on setting up Google Data Studio
