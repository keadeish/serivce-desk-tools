**<center>Active Directory Exporting Tool</center>**

A tool which exports groups & sub-groups from Active Directory. This tool's script was created with PowerShell & its GUI was designed with C#; using Windows WPF framework.

**How to run**
1. Create a folder at your desired destination on your PC - e.g. 'C:\Temp\ServiceDeskTools'.
2. Open the 'ServiceDeskTools.ps1' file with PowerShell ISE (**run as administrator**).
3. On **line 3** of ServiceDeskTools.ps1 replace the path of $xamlFile with the path of which you saved the MainWindow.xaml file in the folder onto your PC in step one (e.g. `$xamlFile="C:\Temp\MainWindow.xaml"`) **Don't forget to include MainWindow.xaml!**
4. Run the script by pressing **F5** or the green play icon.
5. Enter the name of the group that you'd like to export.
6. Click export and then select the file path of which you'd like the exported CSV file to be located.

**FAQs**
1. **An error occurs when attempting to run the script** - Make sure that you have the correct permissions. This application uses Microsoft's ActiveDirectory module (https://learn.microsoft.com/en-us/powershell/module/activedirectory/?view=windowsserver2022-ps)

![Service Desk Tools (1)](https://github.com/keadeish/service-desk-tools/assets/90222144/9be6473b-ea16-40db-99c8-d7d4331e57c3)
