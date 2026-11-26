<!--
FILE:      CHANGELOG.md
CREATED:   2026-11-05 14:57 +03
TASK:      #939
AUTHOR:    M Alhassoun
WHAT:      Customer-readable change log of the kwtSMS Excel add-in.
WHY:       Every user-visible change gets a line; this folder had none.
USAGE:     Add entries under [Unreleased]; move them to a version heading on release.
DEPENDS:   Nothing.
-->
# Changelog

## [1.0.21] - 2026-11-26

### Added
- The login screen says what the add-in does, how to get a kwtSMS account and how to reach support.
- Requests to kwtSMS carry the integration name and version (no personal data) so kwtSMS can see which integrations are in use.
- If your kwtSMS account is in Test mode, the send result says the SMS was queued as a test and not delivered.

### Security
- Your kwtSMS API login is no longer saved inside the workbook. It stays in Excel on your device, so sharing a file never shares your login. A login saved in a workbook by an older version is removed the next time the file is opened with the add-in, and you are asked to log in again. You log in once per device.

### Changed
- The add-in is now served from addins.kwtsms.com, a kwtSMS server, instead of GitHub Pages.
- When kwtSMS cannot be reached, login shows a clear message instead of "Failed to fetch".

