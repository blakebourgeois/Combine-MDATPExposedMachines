# Combine-MDATPExposedMachines
a gross script to merge exported data from Microsoft Defender ATP

The portal does not provide quick or easy ways to export all of the vulnerabilities, either by machine or in total.
This is a "quick" way to perform AD integration
The script puts out a master file ($Date-MDATP_Full.xlsx) with all the discovered vulnerablities, and tables listing machines by vuln count and vulns by machine count

## Requirements
Must have the ImportExcel tools (Install-Module ImportExcel) 

## Instructions

Set the $pwd to where you are going to save all the exports from MDATP
Save the reports from MDATP:
1. Go to securitycenter.windows.com
2. Go to Threat and Vulnerability Management > Security Recommendations from the sidebar
3. Under filters > Remediation type select whatever you want to look for. I usually leave out configuration change but there could be value in it
4. Under filters > Status make sure it is only "active"
5. Go through the results list and find results you want to aggregate. For example, you might only want exploitable threats or something with a specific threshold of weaknesses. 
5a. Click the desired entry and hit "export" by the Exposed devices drop down
5b. repeat for every security recommendation to whatever threshold is appropriate for your situation
6. Make sure all the reports are in the $pwd folder.
7. Run script and get the exported data.
