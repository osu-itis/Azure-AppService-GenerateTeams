# Timed Group Check

## About

- Gathers all Teams in our organization and adds information about them into an Azure table, this is used for managing and tracking our Teams.

## Requirements

- This function requires the `MSAL.PS` module which is currently not part of microsoft's default modules that can be auto imported with the `requirements.psd1` file. Because of this, A copy needs to be manually added to `\modules\MSAL.PS\`.

## Use

- This function is automatically ran via a timer trigger and does not need to be manually called.
