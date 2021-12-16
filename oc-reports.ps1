Set-Location C:\oc\Tomcat_B\reports_client.data\reports #Change working directory to make relative variables easier. 
#Set variables
$PSEmailServer = 'mailserver'
$path = "C:\oc\Tomcat_B\reports_client.data\reports"
$date = Get-Date -UFormat %y_%m_%d
$fulldate = Get-Date -DisplayHint Date
$parent = [System.IO.Path]::GetTempPath() #create & store a temporary name
[string] $name = [System.Guid]::NewGuid() #get current date and format for later searches.
$temp = (Join-Path $parent $name) #store full path\file name.
New-Item -ItemType Directory -Path (Join-Path $parent $name) #create & store a temporary path. Uses current user temp path C:\Users\username\AppData\Local\Temp

Copy-Item -Path $path\*$date* -Destination $temp #copy items from production directory to temporary path.
#Compress-Archive -Path $temp -DestinationPath $parent\$name #Requires windows 10 cmdlets, not available in this enviornment

#Compression Workaround for Windows 7. Invoke .NET 4.5 to compress files in temporary directory.
[Reflection.Assembly]::LoadWithPartialName( "System.IO.Compression.FileSystem" )
[System.AppDomain]::CurrentDomain.GetAssemblies()
$destfile = "$parent\$name.zip"
$compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
$includebasedir = $false
[System.IO.Compression.ZipFile]::CreateFromDirectory($temp, $destfile, $compressionLevel, $includebasedir)

#Email compressed archive to users
Send-MailMessage -To email -From email1 -Bcc email2 -Subject "Client Reports Generated $fulldate" -Body "Clients Reports generated on $fulldate. Compressed Archive of files attached." -Attachments $parent\$name.zip -Credential $cred

#Cleanup
Remove-Item -path $parent\$name.zip
Remove-Item -path $parent\$name -Recurse