# PoSh-WSUS-Cleanup
WSUS Server Cleanup using PoSh
* Local_WSUS_Server_Cleanup.ps1 - PowerShell script to allow WSUS Cleanup to be scheduled
* Clean_Failed_Cleanup.ps1 - If WSUS Cleanup fails, run this script to cleanup the SQL part that usually times out and causes the whole cleanup to fail. (Now merged into Local_WSUS_Server_Cleanup.ps1)
