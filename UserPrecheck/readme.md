# UserPrecheck

## About

Check if a user is licensed for Microsoft teams, can be used to check before creating a team for a user that cannot be a member/owner of the team.

## Requirements

No external Modules required.

## Use

### Get Request

```HTTP REST
GET /api/UserPrecheck?code=<AZUREFUNCTIONKEY>&ticket=<TICKETNUMBER>&user=<USEREMAILADDRESS>
Host: <HOST>
```

#### Example Response

```json
{
    "User": "carrk@oregonstate.edu",
    "TeamsEnabled": "Success"
}
```
