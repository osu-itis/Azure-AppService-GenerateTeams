# Attempt to import the required module
if ((Get-Module MSAL.PS).count -ne 1) {
    $attempt = 0
    do {
        $attempt ++
        start-sleep -Seconds 2
        Write-Output "Attempting to import the MSAL.PS module..."
        Import-Module MSAL.PS -Force -ErrorAction SilentlyContinue
    } until (((Get-Module MSAL.PS).count -eq 1)-or($attempt -eq 10))
    switch ($attempt) {
        '10' {throw "Could not import the MSAL.PS module"}
    }
    Remove-Variable "attempt"
}
else {
    Write-Output "MSAL.PS module already loaded"
}
