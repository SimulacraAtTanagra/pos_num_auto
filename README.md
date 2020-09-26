# pos_num_auto
This process has been created for the automatic update to the position number file. 
It takes the latest position status report, strips it, filters it and uses the data to update a file on the shared folder for use. After running successfully, it sends an update e-mail tot he HRIS team. 

Current to-do:
-Updating without needing the position status report (note to self, by benchmarking against report of all position numbers and all currently claimed numbers from other lcoal reports)
-Automating the scheduling for this process (if above is accomplished, daily process)
-Enabling filesharing from Python (note to self, using pyautogui maybe)
