# GUI for MDE API sample app
Simple PowerShell GUI for Microsoft Defender for Endpoint API machine actions.

Forked from Microsoft repo (MIT Licensed): https://github.com/microsoft/mde-api-gui/

> [!IMPORTANT]
> This project has nothing to do with Microsoft.

> [!NOTE]
> If you intend to use this with many machines (100+), consider adding throttling handling to avoid API rate limiting. There is already one with 500 milliseconds delay between each request, but it may not be enough. And note that it takes a lot of time to process many machines in sequence=> 100 machines ~ 1min30s

## Pros

- No installation of SDK needed
- Quick to execute and simple GUI
- Very useful in case of critical incident
- Has a file picker for CSV
- Can Add / Remove tags
- Added supported for unmanaged device isolation type (new in MDE API); it will basically do a "Network contain" action on unmanaged devices, since there isn't the agent to enforce isolation.

## Cons

- Will be more difficult to keep up to date
- Disabled Advanced Query research, now has only one option: CSV import.

<img width="943" height="824" alt="image" src="https://github.com/user-attachments/assets/90c1090f-6ab7-4a83-a158-2dc12871afd6" />

## Get started
1. Create Azure AD application as described here: https://learn.microsoft.com/en-us/microsoft-365/security/defender-endpoint/apis-intro?view=o365-worldwide
2. Grant the following API permissions to the application:

| Permission | Description |
|-------------------------|----------------------|
| AdvancedQuery.Read.All	| Run advanced queries |
| Machine.Isolate |	Isolate machine |
| Machine.ReadWrite.All |	Read and write all machine information (used for tagging) |
| Machine.Scan |	Scan machine |

3. Create application secret.
## Usage
1. **Connect** with AAD Tenant ID, Application Id and Application Secret of the application created earlier.
2. **Get Devices** that you want to perform actions on, using one of the following methods:
    * CSV file (single Name column with machine FQDNs)
3. Confirm selection in PowerShell forms pop-up.
4. Choose action that you want to perform on **Selected Devices**, the following actions are currently available:
    * Specify device tag in text box and **Apply tag**.
    * Run **AV Scan**.
    * **Isolate**/Release device.
5. Verify actions result with **Logs** text box.

## TODO (may or may not be implemented)

* Disable device in Entra ID (different API)