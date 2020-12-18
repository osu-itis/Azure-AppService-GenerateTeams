# Timed Group Check

## About

- Gathers all Teams in our organization and adds information about them into an Azure table, this is used for managing and tracking our Teams.
- Teams can have one of the following Partition Keys:
  - **KnownTeams**
    - Teams that have been auto-added to the list and should be treated as a normal team.
  - **ExemptTeams**
    - Teams that have been manually added to the Exempt list and will most likely not have our expected settings.
    - If a team has been added with both the `ExemptTeams` and the `KnownTeams` partition, the duplicate `KnownTeams` will be auto removed.

## Requirements

- This function requires the `MSAL.PS` module which is currently not part of microsoft's default modules that can be auto imported with the `requirements.psd1` file. Because of this, A copy needs to be manually added to `TimedGroupCheck\modules\MSAL.PS\`.

## Use

- This function is automatically ran via a timer trigger and does not need to be manually called.
