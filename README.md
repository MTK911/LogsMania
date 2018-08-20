# LogsMania
Microsoft exchange server automated tracking logs analysis tool. Analyse logs using LogParser tool and with the power of Powershell writes reports in readable formats.

# Dependencies 
+ PowerShell 5 (Optional Ver. 5)
+ PowerShell Modules (ImportCSV, ImportExcel)
+ Log Parser 2+
+ MS Excel (To read reports)


# Configuraitons
1. Change all contents of PS1 files according to your specifcations
2. Emails within Organization and outside Organization are detected based on Domain Name, Change org.com to your own domain names
3. Report templates are configured based on specified cells and sheets. Change them accordingly as you make changes in code or template
4. Place "Logs Mania" directory in "C:\" drive. Also change the name of "LogsMania-master" to "Logs Mania" with a space.

All these configurations are mandetory if you want the script to work properly. Also all these changes can be change according to your need.

In case of any concern, Issue or Question feel free to let me know.

# To Run
1. Place your tracking logs in LOGS directory
2. Run "logman.bat" and everything will takecare of itself hopefully :p

