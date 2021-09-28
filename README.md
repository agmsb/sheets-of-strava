# google-sheets-strava

## Overview

Every strava activity is written to a sheet titled `raw_data`. What you do with that afterwards is up to you.

I personally have the below set up:
- Google Data Studio Dashboard @ [cycling.agmsb.dev](https://cycling.agmsb.dev)
- `monthly_data` roll-up sheet
- `yearly_data` roll-up sheet
- `gear_data` roll-up sheet
- `friends_data` roll-up sheet

## Setting up `raw_data`

Create a sheet titled `raw_data`.
In `Tools > Script Editor`, create two files:
- `code.gs`
- `oauth.gs`

Update `oauth.gs` with your Strava client id and client secret.

## Setting up `monthly_data` or `yearly_data`

Example: roll-up rides by month
```
=COUNTIF(raw_data!A:A, "*"&A2&"*")
```

Example: roll-up mileage by month
```
=SUMIF(raw_data!A:A, "*2020-06*", raw_data!D:D)
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

In `Friends`, I manually insert the friends I rode as `First Name` `Last Name`, with each friend separated by comma. 

Then in the sheet `friends_data`, you can roll-up by data by friends using the below examples:

Example: roll-up rides by friend
```
=COUNTIF(raw_data!M:M, "*"&A2&"*")
```

Example: roll-up mileage by friend
```
=SUMIF(raw_data!M:M, "*"&A2&"*", raw_data!D:D)
```

## TODO

- Automatically update epoch time to right after last activity in `code.gs`
- Add tutorial on setting up Google Data Studio
