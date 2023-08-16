# Intro
This is a script designed to collect *Summary Report* and *Extended Summary Report* Message Traces for emails handled by an Exchange Online Tenant in the past 10 days.
This script is designed to export data to CSV files on local drive on demand, avoiding limits that might lead to data truncation when using Exchange Online Admin Center.
For more information on Message Traces in Exchange, check Microsoft documentations [here](https://learn.microsoft.com/en-us/exchange/monitoring/trace-an-email-message/message-trace-modern-eac)

# Requirements
Before running the script, you need to be connected to Exchange Online Powershell:
[Exchange Online Powershell Module](https://learn.microsoft.com/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps)

This script was built and tested on Windows 11, with the following versions of Powershell and ExchangeOnline Module:
```
> $PSVersionTable
Name                           Value
----                           -----
PSVersion                      5.1.22621.1778
PSEdition                      Desktop
```
```
> $PSVersionTable
Name                           Value
----                           -----
PSVersion                      7.3.6
PSEdition                      Core
```
```
> Get-Module "ExchangeOnlineManagement"|fl Version
Version : 3.2.0
```
```
> Get-ComputerInfo|fl OsName, OsVersion
OsName    : Microsoft Windows 11 Enterprise
OsVersion : 10.0.22621
```
