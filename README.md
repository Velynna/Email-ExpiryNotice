# Send-ExpiryNotice.ps1
PowerShell script to email users that their Active Directory password is soon expiring, along with info on how to change it. It is designed to run as a scheduled task on a server with the Active Directory PowerShell module installed.

How to use:
1. After downloading Send-ExpiryNotice.ps1, open the file with your preferred text/code editor and scroll to the Begin block; insert your own data.
2. If you'd like to update the styling on the email to your own, do that in the Begin block as well.
3. At the very bottom of the script is an invokation. It is initially commented out, so you can test from the CLI. Before setting up the task, uncomment this line. A task will not run a function on its own, so the invokation line is critical to functionality.

Recommended to run once per day.

Update notes from 3.0 to 4.0.
+ Formatting updated to be in-line with PowerShell style guidelines (within reason), and best practices.
+ Extraneous files removed.
  + Task scheduling is done in the Task Scheduler instead of leveraging Invoke-Installation.
  + Remove-ScriptVariables is no longer necessarily. All variables clear on their own when a function ends.
  + Required statement takes the place of Set-ModuleStatus.
  + No images are used.
+ Remaining files moved into 1 single file, with a line to call at the end.
  + This makes it easier for a task schedule, and means less files to lose.
+ Dramatic changes to $EmailBody
  + Replaced content with branded email. (Branding has been removed.)
  + Email uses a responsive design with in-line styling.
  + Instead of if/else logic within the email, the email was changed to a solid block with variables, and the variables are used in if/else logic before the EmailBody variable.
  + Variables with company data were updated for the new email template.
  + Variables in a modified style sheet for easy-updating.
+ Updated Notes.
+ Expanded Parameter blocks, and updated CmdletBinding blocks.
+ Removed extraneous variables, and lines that had been commented out.
  + Variables for company info were removed in favor of hard-coded information. Unfortunately the side-effect is that slightly more advanced knowledge is required for use.
+ Added email at the very end of the process block to notify when the task has completed.
+ Updated EventLog text.
+ Added Debug checks.
+ Domain Function Level check has been removed. Windows Server 2008 or higher now required.
