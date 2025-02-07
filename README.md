# Add-EmailsCSVtoDistributionList
Made for those that need to add a bulk amount of emails to a distribution list group quickly from their responses on a Google Form or similar, as opposed to–annoyingly–manually inputting every email address into Exchange.

## Disclaimer:
The -AllowClobber parameter on the installation of the ExchangeOnlineManagement module is present to ensure its functionality; make sure you understand the implications of this and it doesn’t overwrite any PowerShell cmdlets you care about. I recommend installing the module manually if this is of any concern to you.
The author is not responsible for any misuse, malpractice, or damage that may stem from the use of this script. Please review the script and upload the file to https://www.virustotal.com/gui/home/upload if there are any concerns. This script does not require admin rights. RUN AT OWN RISK.

## Requirements:
To be a manager of the distribution group that you wish to add these emails to.
PowerShell version 5.1 or later (this is typically pre-installed on latest Windows versions).
NuGet Package Provider (the script will install this if it detects it as missing).
Exchange Online Management PowerShell module (the script will install this with -AllowClobber if it detects it missing).

## Steps:
Acquire your .csv containing a column full of emails, preferably with a header (first row) that allows you to note in the future that it is the correct column. If coming from Sheets or Excel, in the spreadsheet that you wish to add the emails from, click “File” and download the spreadsheet as a Comma Separated Values (.csv) file.

Download the script. Open File Explorer, navigate to the script in your Downloads folder, right click, and click “Run with PowerShell.”

Follow the instructions in the script. There may be periods of waiting for the script to process or download if applicable. If prompted to install NuGet or ExchangeOnline, please do both. You will be prompted to sign in to your Microsoft account that manages the distribution group you wish to add these emails to.