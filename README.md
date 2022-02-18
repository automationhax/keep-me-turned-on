# keep-me-turned-on
Automation script to keep windows computer screen turned on. In the background the script creates an Excel macro to move the mouse and uses Shell Script to right click on the screen.

# Step 1 - Copy Script
`Navigate to "Start" file in this project`

# Step 2 - Create Text File
`Copy Start file content and paste into text file on Desktop (or any folder)`

# Step 3 - Save Text File as .vbs file
`Save text file with File name as "Start.vbs" and Save as type as "All types (*.*)". 
The file should look blue (vbs file color).`

# Step 4 - DONT SKIP (Shows how to stop the script)
`Create a shortcut on Desktop and paste this in as the path: 
"C:\Windows\System32\taskkill.exe /IM wscript.exe /f". 
Double clicking this shortcut will stop the automation.`

`Or right click on the Taskbar on Desktop > open Task Manager > Right click and close the wscript.exe process `

# Step 5 - Start Script
`Double click Start.vbs to start the automation script`
