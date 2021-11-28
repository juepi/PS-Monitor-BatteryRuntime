# Monitor-BatteryRuntime
PowerShell script to monitor battery runtime in Windows. It automatically installs a Scheduled task at first run that will start at System startup and run as system account.  
The script will create a "BatteryRuntime-Log" Subfolder in the folder where it is called from. beside some helper files (which you should not delete), you can find a Excel compartible `Results.csv` file, which will be permanently updated by the Scheduled task showing you the following information for each battery discharge cycle (time between plugging in AC power):
- Date when the discharge cycle ended
- On-battery runtime
- energy discharged from battery (percentage)
- Estimated runtime for a full battery discharge cycle

## Notes on usage / Limitations
Note that the script will probably only work (correctly) with notebooks that have a single battery. There might be OS limitations, i'm running it on Windows 11, but it should also work on recent Windows 10 installations.  
The script might be called with the `-verbose` parameter from the shell to get some output for debugging purpose.  

Have fun,
Juergen
