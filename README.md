# Microsoft 365 Apps packager

A workflow for creating a Microsoft Intune package for the Microsoft 365 Apps.

Requires the following secrets on the repo:

* `COMMIT_EMAIL` - email address used for commits
* `COMMIT_NAME` - user name used for comments
* `GPGKEY` - GPG key to sign commits
* `GPGPASSPHRASE` - passphrase to unlock the GPG key
* `TENANT_ID` - tenant ID used by `new-autopackage.yml`
* `CLIENT_ID` - app registration client ID used by `new-autopackage.yml` to authenticate to the target tenent
* `CLIENT_SECRET` - password used by `new-autopackage.yml` to authenticate to the target tenent

## Scripts

`Create-Win32App.ps1` imports the intunewin package into the target Intune tenant, using `App.json` as the template.
