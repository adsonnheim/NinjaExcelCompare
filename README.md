# NinjaExcelCompare
PowerShell script to check the devices in both an excel spreadsheet and NinjaOne RMM.

How to use:
- Create a copy of `.env.example` and rename it to `.env.local`
- Fill out parameters in newly created .env.local file of where your users are located in your domain, e.g.:
    - `CLIENT_ID=xxxxxxxxxxxxxxxxxxxx`
    - `CLIENT_SECRET=xxxxxxxxxxxxxxxxxxxx`
    - `RMM_URL=https://xxx.rmmservice.com`
    - `REPORT_URL=/path/to/where/file/is/located`
    - `TRACKED_SHEETS=Sheet 1, Sheet 2, Sheet 3`
- Run `NinjaExcelCompare.ps1`
- View results in terminal