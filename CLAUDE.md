# Tajneed dashboard

Static pages served by GitHub Pages at https://tmhtwe.github.io/tajneed/. This repo is public.

- `index.html`: Tajneed dashboard. It loads `data/members.enc.json` and decrypts it in the browser with the access key the viewer enters.
- `data/members.enc.json`: all member data, encrypted (PBKDF2-SHA256 + AES-256-GCM). The access key is never stored in the repo.
- `tools/update.js`: rebuilds the encrypted file from a "MembersGroupedByJamaat" Excel export.

## Privacy rules

- Never commit member data in plain form: no Excel, CSV or JSON with names, phones, emails, addresses or DOBs. Only `data/members.enc.json`.
- Never commit, print or write the access key into files. Pass it only through the `TAJNEED_KEY` environment variable.

## Updating the data (when the user gives a new MembersGroupedByJamaat Excel)

1. `cd tools && npm install` (first time only).
2. Dry run to see what changed:
   `TAJNEED_KEY=<key> node tools/update.js --excel "<path to xlsx>" --dry-run`
   It lists new members, members no longer listed, and Majlis changes.
3. If it exits with "NEED DOB AND MUSI", ask the user for each listed member's date of birth and whether they are Musi. Write them to a JSON file **outside the repo**:
   `{ "2603": { "birthday": "1983-02-10", "musi": false } }`
   The same file can change an existing member's DOB or Musi status.
4. Run again without `--dry-run`, adding `--extra <that file>`.
5. Once a year the user gives next year's Ansars-to-be list (Code, Name, Majlis, Birthday, Email, Mobile). Pass it with `--tobe <file>`. Without it the current list is kept.
6. Commit only `data/members.enc.json` and push to `main`. Pages updates in a minute or two; the dashboard's Refresh button reloads it.

DOB and Musi are kept inside the encrypted file, carried over by member code. Members who drop off the list keep their DOB/Musi in an archive inside the file, so they don't need to be re-entered if they come back.

Dashboard rules (in `index.html`): Saff-e-Awwal is over 55 by birth year, Saff-e-Dom otherwise. "New Ansars" this year are those born in (current year − 41).
