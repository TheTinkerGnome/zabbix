Windows-DomainController
======
This template checks various Performance monitors and services associated with Active Directory, as well as running the dcdiag tool to assess AD health

Usage

1. Place dcdiag_check.vbs in the zabbix script directory of your DHCP server
2. Add the following line to UserParameter section of the Comain Controller's config and point it to the script like this: dcdiag[*],cscript //Nologo C:\bin\Zabbix\scripts\dcdiag_check.vbs $1
3. Restart the zabbix agent service
