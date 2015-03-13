Windows-Updates
======
This template checks to see how many Important and Optional updates need to be installed. Can be used
in an environment using WSUS or retrieving updates directly from Microsoft

Usage

1. Place winupdates.vbs on the Windows machine you wish to monitor
2. Configure the serverName variable with the IP of your Zabbix server
3. Configure the zbxSender variable with the path to zabbix_sender.exe
4. Configure the script to run as a scheduled task
