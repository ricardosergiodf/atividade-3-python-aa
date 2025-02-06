$exclude = @("venv", "ativ-pratica-3-python-aa.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "ativ-pratica-3-python-aa.zip" -Force