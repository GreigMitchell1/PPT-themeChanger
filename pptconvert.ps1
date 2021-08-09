#Renfrew North Hymn Template Changer v1.1
#Greig Mitchell - August 2021

#load required libraries, assign com objects
Add-type -AssemblyName office
$Application = New-Object -ComObject powerpoint.application
$application.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
#Theme to change files to
$themePath = "C:\Users\greig\Desktop\pptps\RenfrewNorth-V1.4.potx"
#Files to Change
$path = "C:\Users\greig\Desktop\pptps\"
#loop to pickup all ppt / pptx files
Get-ChildItem -Path $path -Include "*.ppt", "*.pptx" -Recurse |
ForEach-Object {
 $presentation = $application.Presentations.open($_.fullname)
 $presentation.ApplyTemplate($themePath)
 $presentation.A
 $presentation.Save()
 $presentation.Close()
 #Debug print ppt file to terminal
 "Modifying $_.FullName"
} 

$application.quit()
$application = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()
