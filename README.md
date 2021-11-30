Click here below to open this project in VSCode in your browser  
[Primary module file](https://github1s.com/Bukinnear/NUCAD/blob/master/_Modules/NUC-AD/NUC-AD.psm1)  
[Client script Template](https://github1s.com/Bukinnear/NUCAD/blob/master/_Modules/NUC-AD/NUC-AD.psm1)

### What is it?  
**'New User Creation - Active Directory'**, or the 'NUCAD' Module, is a powershell utility developed to help automate and fast-track new user onboardings.

### How does it work?  
The [module](../master/_Modules/NUC-AD/) contains generic functions intended to be imported and used from a separate, client-specific script.
The client script will handle things that vary across clients (such as username, and email formats, on-premise exchange vs O365, shared drives, if applicable), and calls public functions from the NUCAD module.

### How to use:  
you can manually import the module with `Import-Module -name "C:/Path/to/Module"`, but the included template file under [_Utilities](../master/_Utilities) will automatically import it from the script's directory, and works in Powershell v2 onwards

### What can it do?  
All publically available functions are demonstrated (and commented!) in the template file
