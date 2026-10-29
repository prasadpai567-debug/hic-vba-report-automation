# Source modules

The `.bas` files in this folder are sanitized exports of the VBA modules reviewed from the original workbook.

- `controller/HicController.bas` contains the controller workflow. It uses placeholders for the internal network path, recipients, sender, company branding, and links.
- `report/HicAutomation.bas` contains the report-side refresh and SQL integration module. It uses placeholders for the internal network path, email settings, and SQL connection string.

Before using these modules in an enterprise environment, restore approved deployment configuration through a controlled process. Do not commit production paths, recipients, database identifiers, credentials, report data, or runtime logs.
