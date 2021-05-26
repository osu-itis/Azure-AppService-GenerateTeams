# TeamGuestSettings

## About

Get the gest settings of the specified team

## Requirements

No external modules required

## Use

### Get Request

```HTTP REST
GET /api/TeamGuestSettings?code=<AZUREFUNCTIONKEY>&TeamName=<TEAMNAME>
Host: <HOST>
```

#### Example Response

```JSON
{
    "id": "4172d260-9f5b-4a90-8ec2-55315a4b5874",
    "mail": "KeenanTestingTeams1171034411@OregonStateUniversity.onmicrosoft.com",
    "displayName": "Keenan Testing Teams",
    "AllowToAddGuests": "true",
    "GuestSettingsID": "998624a0-1a8c-4201-8b3c-d7718ff26c4f"
}
```

### Post Request

```HTTP REST
POST /api/TeamGuestSettings?code=<AZUREFUNCTIONKEY>
Host: <HOST>
Content-Type: application/json

{
"TicketID": "01234567",
"Requestor": "carrk@oregonstate.edu",
"TeamName": "Keenan Testing Teams",
"GuestSettingsEnabled": "true"
}
```

#### Example Response

```JSON
{
    "id": "4172d260-9f5b-4a90-8ec2-55315a4b5874",
    "mail": "KeenanTestingTeams1171034411@OregonStateUniversity.onmicrosoft.com",
    "displayName": "Keenan Testing Teams",
    "AllowToAddGuests": "true",
    "GuestSettingsID": "998624a0-1a8c-4201-8b3c-d7718ff26c4f"
}
```