# CM-Instant-Deploy
A tool I have made if I need to deploy a specific application to a specific workstation in real time. 


## Steps
- Enter computer name and press Test Connection to ensure workstation can be reached.
- The server name and site code will be prepopulated with the CM server and site code that the workstation the script is ran from is set to.
- Select Application will bring up a screen to select the application. The folder structure will be the same format as the SCCM console.
- Deploy the application.
  
![alt text](ID.png)




## Requirements
Must have administrative rights to workstations. 
<br>
Must have access to CM server (done through an invoke-wmi)
