A simple Powershell-based Windows inventory and survey tool with UI. 

-- Scans defined any windows servers specified in the host field, .txt file or .CSV file and returns datapoints/properties for those hosts into ui and file output.  

--Datapoints Returned
-Winrm enabled?
-User folders list
-Scheduled Tasks (non-system)
-Open web/mail ports (443, 80, 8443, 8080, 25)
-ICMP polling (ping)
-IIS installed?
-SQL installed? + Instance Name

-Runs best in ISE or VScode
-Winrm required for datapoint polling. 
