Get-ADUser -Filter 'Enabled -eq "True"' -properties *| where -Property Officephone -ne $null|ft Name, department,Officephone,sAMAccountName,Mail,Company,physicalDeliveryOfficeName,msRTCSIP-Line
