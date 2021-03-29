# Logic Apps

- [Logic Apps](#logic-apps)
  - [EditTeamsSettings](#editteamssettings)
    - [Description](#description)
    - [Example call](#example-call)
    - [Example response](#example-response)
  - [ViewTeamsSettings](#viewteamssettings)
    - [Description](#description-1)
    - [Example call](#example-call-1)
    - [Example response](#example-response-1)

>NOTE: The two `.logicapp.json` files are copies of the gui edited Logic Apps.

## EditTeamsSettings

### Description

Edit the guest access settings for a given team, will return a `200` on success. Correctly formatted requests failing within Microsoft Graph will return a `502 bad gateway` (this most likely indicates that the team name is incorrect).

### Example call

```API
POST
https://prod-26.westus2.logic.azure.com:443/workflows/98f6dff293824b32a1f8a84dcd37cf21/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig={KEY}

{
"TicketID": "{TICKETID}",
"Requestor": "{REQUESTOR}",
"TeamName": "{TEAMNAME}",
"GuestSettingsEnabled": "true"
}
```

- >NOTE: The GuestSettingsEnabled is a string (even though that value is a boolean true/false) this is configured in the same way that Microsoft Graph requires the value when posting changes.

### Example response

```json
{
    "Requestor": "{REQUESTOR}",
    "TeamID": "{TEAMGUID}",
    "TeamName": "{TEAMNAME}",
    "TeamSettings": [
        {
            "id": "{TEAMSETTINGGUID}",
            "displayName": "Group.Unified.Guest",
            "templateId": "08d542b9-071f-4e16-94b0-74abb372e3d9",
            "values": [
                {
                    "name": "AllowToAddGuests",
                    "value": "{TRUE/FALSE}"
                }
            ]
        }
    ],
    "TicketID": "{TICKETID}"
}
```

## ViewTeamsSettings

### Description

View the guest access settings for a given team, will return a `200` on a successful response. Correctly formatted requests failing within Microsoft Graph will return a `502 bad gateway` (this most likely indicates that the team name is incorrect).

### Example call

```API
GET
https://prod-27.westus2.logic.azure.com:443/workflows/d3b2e83f4dad4eaf9ee0f3785fabf99d/triggers/manual/paths/invoke/teams/{TEAMNAME}?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig={KEY}
```

### Example response

```JSON
[
    {
        "TeamID": "{GUID}",
        "TeamName": "{TEAMNAME}",
        "TeamSettings": [
            {
                "name": "AllowToAddGuests",
                "value": "{TRUE/FALSE}"
            }
        ]
    }
]
```
