# Telefon Excel

Reads a Samsung phone plugged in over USB and writes the details to a single Excel file. Plugging the same phone in again does not add a new row. The existing row is updated by serial number.

The window stays small and on top, so 50–60 phones can be connected one after another.

## What it reads

| Field | Source |
| --- | --- |
| Model | `ro.product.model` |
| Serial | `ro.serialno`, or the ADB serial if that is empty |
| IMEI 1 / IMEI 2 | `iphonesubinfo` service |
| CSC | `ro.csc.sales_code` |
| CSC version | `mdc.omc.config_version` |
| Build number | `ro.build.display.id` |
| Baseband | `gsm.version.baseband` |
| Android | `ro.build.version.release` |
| One UI | `ro.build.version.oneui` (for example `80500` becomes 8.5) |

Excel columns: No., Date Time, Serial, Model, IMEI 1, IMEI 2, CSC, CSC Version, Build Number, Baseband, Android, One UI, Status.

IMEI and serial values are stored as text, so Excel does not turn them into scientific notation.

## How to run

1. On the phone, turn on Developer options and USB debugging.
2. On the first connection, accept **Allow USB debugging** on the phone. Without that approval, nothing is written.
3. Run `Baslat.bat`.

If `openpyxl`, `pyserial`, or `Pillow` is missing, the batch file installs them. `adb.exe` is in this folder, so ADB does not need a separate install.

The default workbook is `telefonlar.xlsx`. **Excel** opens that file. **Gözat** picks a different one. Do not leave the workbook open in Excel while phones are being recorded. If it is locked, the row is written to `telefonlar_yedek.xlsx` and copied back when the file is free.

**Durdur** pauses reading new phones. **Yeniden oku** reads the connected phone again. **Üstte** keeps the window on top.

## Label

**IMEI 1 etiketi bas** prints only IMEI 1 on a paired Nelko P21. The barcode is Code 128, with the number beside it as `XXXXXX XXXXXX XXX`. IMEI 2 is not printed.

The printer speaks Bluetooth SPP over a COM port. Windows Settings showing Connected and then immediately Disconnected is normal. The port stays open only while a label is printing.

Label size is read from the roll in the printer. If that fails, the job uses 14×40 mm with a 5 mm gap.

## Not in the repository

`telefonlar.xlsx`, `telefonlar_yedek.xlsx`, `kayit.log`, and `ayar.json` are local records and settings. They are not committed.
