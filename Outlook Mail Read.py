#import wmi module  
import wmi  
   
# Initializise  the wmi constructor  
f = wmi.WMI()  
   
# Print the header   
print("Printing the pid   Process name")  
   
# all the running processes  
for process in f.Win32_Process():  
print(f"{process.ProcessId:<5} {process.Name}")  