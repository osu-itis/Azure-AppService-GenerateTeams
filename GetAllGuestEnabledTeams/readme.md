# GetAllGuestEnabledTeams

## About

This function gathers a (json) list of all of the guest enabled teams

## Requirements

No external modules required

## Use

GET request

## Example results

```json
[
    {
        "Id": "8e2e1094-a1ae-4310-a01c-e165b704dc3f",
        "DisplayName": "TestTeamPleaseIgnore",
        "Mail": "TestTeamPleaseIgnore@OregonStateUniversity.onmicrosoft.com",
        "AllowToAddGuests": "True"
    },
    {
        "Id": "4172d260-9f5b-4a90-8ec2-55315a4b5874",
        "DisplayName": "Keenan Testing Teams",
        "Mail": "KeenanTestingTeams1171034411@OregonStateUniversity.onmicrosoft.com",
        "AllowToAddGuests": "true"
    }
]
```
