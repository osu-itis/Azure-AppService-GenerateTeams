# NOTES

## Graph commands needed

- Get all of the unified groups, and then select the ID and the resource provisioning option. This gives us everything we need to filter for all groups that are teams

```GRAPH
https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(a:a eq 'unified')&$select=id,resourceProvisioningOptions
```

- Getting the information on a team's settings, we can use this to determine if anything needs to be changed

```GRAPH
https://graph.microsoft.com/v1.0/teams/{id}
```

- We can then patch a list of members (up to 20 at a time) into the membership of the group

```GRAPH
PATCH https://graph.microsoft.com/v1.0/groups/{group-id}
Content-type: application/json
Content-length: 30

{
  "members@odata.bind": [
    "https://graph.microsoft.com/v1.0/directoryObjects/{id}",
    "https://graph.microsoft.com/v1.0/directoryObjects/{id}",
    "https://graph.microsoft.com/v1.0/directoryObjects/{id}"
    ]
}
```

## Process

- Timer Trigger (every hour?)
- Gather all of the members of the existing management group
  - Check members for compliance?
    - Make changes as needed
      - Notify the team of changes?
- Get a list of all of the unified groups, filter out non-teams, filter out members of the management group, filter out members of the exceptions group
- Process (or log, unclear...?) all of the teams that are NOT in the management group
  - Send notification to team that settings are changing (if any)
  - Set changes on team
  - Add team to the management group

## Other (need a process in teams to request team is added to exception list?)

## Do we need logging?

- how long to keep logs for?
