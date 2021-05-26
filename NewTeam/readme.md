# NewTeam

## About

Generate a new team

## Requirements

No external modules required

## Use

### Request to generate a new team

```HTTP REST
POST /api/NewTeam?code=<AZUREFUNCTIONKEY>
Host: <HOST>
Content-Type: application/json

{
  "TeamDescription": "Generated through API call",
  "TeamType": "Private+Team",
  "TeamName": "Some Name",
  "TicketID": "00000000",
  "Requestor": "email.address@oregonstate.edu"
}
```

#### Example response

```JSON
{
  "ID": <GUID-of-new-Team>,
  "Description": <Description of Team>,
  "TicketID": "00000000",
  "Status": "SUCCESS",
  "rowKey": <Guid of entry in Azure Table Storage>,
  "Visibility": "Private",
  "Mail": "<Email.Address>@OregonStateUniversity.onmicrosoft.com",
  "DisplayName": <Description>,
  "partitionKey": "TeamsLog"
}
```
