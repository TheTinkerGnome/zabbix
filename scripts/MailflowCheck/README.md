# MailflowCheck

## About
This script runs the IMAP Checker utility to get the most recent unread message from it's configured mailbox and reports the date/timestamp to the Zabbix server using the zabbix_sender process

## Requirements
**Zabbix Agent For Windows**

**IMAP Checker Utility** (https://github.com/DavidWGilmore/imap-checker/)

**Windows 7/ Server 2008 or newer** (required by IMAP Checker)

## Configuration
The following variables need to be configured before runing the script

**serverName** - The IP address or hostname of your Zabbix server

**hostName** - The value of the Host Name field for the given item

**zbxSender** - Path to the local zabbix_sender application

**imap_checker** - Path to the IMAP Checker utility

For a list of epoch offset values for various timezones see http://www.epochconverter.com/epoch/timezones.php

**standardTimeOffset** - Epoch offset in seconds

**DSTOffset** - Epoch offset in seconds during DST
