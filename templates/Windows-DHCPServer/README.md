Windows-DHCP
======
Based on the script written by Marco Gottardello from http://atsoy.blogspot.com/2010/06/vbs-script-check-dhcp-state.html

Usage

1. Place check_dhcp.vbs in the zabbix script directory of your DHCP server
2. Add a dhcp.free UserParameter to the DHCP server's config like this: UserParameter=dhcp.free,cscript.exe //Nologo c:\bin\zabbix\scripts\check_dhcp.vbs
