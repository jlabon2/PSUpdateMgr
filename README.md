# PSUpdateMgr - GUI for PSWindowsUpdate

A frontend for the PSWindowsUpdate module that allows patch deployment to chosen computer groups generated from a WSUS server. Fully multithreaded and supports logging and muti-client remote connections using RDCMan.

<b>Usage</b>
![Alt text](web/PUM-Settings.png "Settings")
Define the WSUS server, logging path, and computer groups to generate a list of clients to monitor.
___
![Alt text](web/PUM-Overview.png "Overview")
Shows a list of all selected WSUS clients and their state.
___
![Alt text](web/PUM-Install.png "Install")
Allows the immediate or scheduled install of updates, allowing automatic restarts (with repatching upo boot). 
___
![Alt text](web/PUM-Logs.png "Logs")
Displays all actions as logged. 
___

Allows for the ability to remotely start Windows updates on one or more clients at a time, while instant WSUS checks can also been initiated. Reboots can be initiated forcefully or as needed. Additionally, single clients can been remotely connected to using MSTSC or in bulk using RDCMan.



