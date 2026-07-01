$Error.Clear()      # This is a global variable!
$Errors = @()
$Items = Get-ChildItem e:\wpdata -Recurse -ErrorAction SilentlyContinue
ForEach($Err In $Error)
{
    $Errors += $Err.Exception
}
ForEach($Item In $Items)
{
    Try
    {
        $Item | Get-ACL -ErrorAction SilentlyContinue | Out-Null
    }
    Catch
    {
        $Errors += "$($_.Exception.Message) $($Item.FullName)"
    }
}    
Write-Host $Errors.Count "errors encountered."
$Errors | Out-File errors.txt