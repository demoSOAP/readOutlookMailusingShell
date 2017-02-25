$Outlook = New-Object -ComObject Outlook.Application
$mapi = $Outlook.getnamespace("mapi");

#for reading
$data = $mapi.GetDefaultFolder(6).Items;
#for sending
$Mail = $Outlook.CreateItem(0)

#for writing to csv used pipe to format the result
$data | select SenderName, SentOn, Subject|
Export-Csv -Path "C:\test\SCEEmailList.csv" -NoTypeInformation -Encoding ASCII -Delimiter ','   

#for sending mail
$Mail.To = "demohan@aodobe.com"
$Mail.Subject = "Action"
$Mail.Body ="Pay rise please"
$Mail.Send()