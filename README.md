# Collecting-info-from-OV-and-iLO : server-vitals.ps1
This sample script is used to collect HW information of servers either through OneView or through iLO, specifically:
   * iLO firmware version
   * System BIOS version
   * NIC information: FW version - Health Status
   * Logical/Physical disk information: Count - Encryption status - Health status
   * NVMe disks (if existed):  Count - Health status - FW version of NVMe backplane controller
   * Memory: Size - Health Status
   * Fan: Health Status
   * PDU: Health status
   * System health Check

It then generates an Excel file with multiple sheets: the first sheet contains list of servers that are in ready state ( no HW issue), other sheets contains list of servers that exhibit specific HW issues ( Fw not compatible, NIC errors, disk errors...)

Finally, it generates AHS log file for each server that potentially has HW challenges

## Prerequisites
   * OneView PowerShell library v5.0
   * HPEiLOcmdlets
   * ImportExcel module from PowerShell gallery
   * Excel file containing racks information ( see Samples)
   * config.ps1 contains various credentials and FW baseline

## Note
The script per se may not run correctly in your environment as it has specific customization. However, you can use it as reference to OneView/iLO functions or capabilities to collect HW information.

## Syntax
```
    .\server-vitals.ps1 -OV_Appliance_IP <OV-IP-Address>  -sourceXLS samples.xlsx

```




