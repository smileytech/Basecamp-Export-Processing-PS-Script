function ZipFiles( $zipfilename, $sourcedir )
#Function from http://stackoverflow.com/questions/1153126/how-to-create-a-zip-archive-with-powershell
{
   Add-Type -Assembly System.IO.Compression.FileSystem
   $compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
   [System.IO.Compression.ZipFile]::CreateFromDirectory($sourcedir, $zipfilename, $compressionLevel, $true)
}

#Set basic path variables
$dirpath = 'C:\Basecamp\account-xyz\projects'
$cssfilename = 'application.css'
$csspath = [io.path]::Combine($dirpath, '..', $cssfilename)

#Update stylesheet reference in each html file
'Starting update to stylesheet references'
$oldStylesheetRef = '<link rel="stylesheet" type="text/css" href="../../../application.css">'
$newStylesheetRef = '<link rel="stylesheet" type="text/css" href="../application.css">'
$files = Get-ChildItem $dirpath -Recurse | Where-Object Extension -eq '.html'
$files.Count.ToString() + ' HTML files found'
$files | ForEach-Object {(Get-Content $_.FullName).Replace($oldStylesheetRef, $newStylesheetRef) | Set-Content $_.FullName}
'Finished updating stylesheet references'

#Copy stylesheet to each project
'Starting stylesheet copy to each project'
$dirs = Get-ChildItem $dirpath | Where-Object PSIsContainer -eq $true
$dirs.Count.ToString() + ' projects found'
$dirs | ForEach-Object {Copy-Item -Path $csspath -Destination (Join-Path $_.FullName $cssfilename)}
'Finished stylesheet copy to each project'

#Zip each project directory
'Starting zipping projects'
$zipdir = 'projectzips'
$zippath = [io.path]::Combine($dirpath, '..', "..", $zipdir)
New-Item $zippath -ItemType Directory
$dirs | ForEach-Object {
    $zipfile = Join-Path $zippath ($_.BaseName.Trim() + '.zip')
    ZipFiles $zipfile $_.FullName
    $_.BaseName.Trim() + '.zip file created'
}
'Finished zipping projects'