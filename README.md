# Family life google script documentation

## Summary

This documentation explains how the family life google suite scripts are setup. Basically it involves a Google Form that the family life staff have access to that can update the family life spreadsheet through the help of some Google Apps Scripts found in the repo.

## Artifacts

- Update status Google forms both staging and production (Live in family life Google Suite)
- Status Google sheets both staging and production (Live in family life Google Suite)
- This repo, contains all the scrips

## Requirements

- The owner of the forms must be a Google Suite Admin
- Users must be part of the Google suite domain and part of the directory API

## Setup

### Staging and Production

A staging environment (scripts, forms, sheets) has been setup for family life to experiment and test the existing solution. A production environment (scripts, forms, sheets) has also been setup that can be updated with the staging script code once staging is verified to be working.

### This repo holds both the staging and production scripts

NOTE: The visual GSuite editor should not be used for versioning or to update scripts if possible. Instead use clasp, the link below should be used to push to the family life Google Suite scripts.

To install and use clasp follow these instructions https://developers.google.com/apps-script/guides/clasp

## Running the scripts

### Set up form triggers on Gsuite

Google suite App script triggers, onOpen and onFormSubmit must link to the methods of the same name in the Code.js files.

### Enable required APIs

Currently the only API this relies on is the Google Directory API (Option can be found in the script editor)

### Debugging the code

* Run onOpen method using debug (verify other debug methods work)

## Family life Google Script Use cases

- Family life outreach workers can update their status using Google forms instead of a manual spreadsheet
- Family life outreach workers can receive calendar events and notifications to update their status
- Family life managers can recieve notifications if their workers haven't updated their status after a given time (Work in progress)

## Important details

- Currently the way the family life Google sheet columns are named is relied upon by the script. I.e. the "Email address" column must exist as well as the "Status" and "Notes" columns. Any change to the format or options in these columns would need to be reflected in the script.

- The Family Life staff email addresses must all be present in the "Email address" column, and must be spelled correctly, with no whitespace

- Google suite App script comes with a debug log that can be found in view -> stacklog 

- Google forms should be setup so only family life Google Suite members can access them (This is the default option)
