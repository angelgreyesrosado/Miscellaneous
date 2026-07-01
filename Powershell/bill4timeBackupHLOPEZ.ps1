Write-Host "Starting bill4time data backup script at $(Get-Date)..."
$ie = new-object -ComObject "InternetExplorer.Application"
$requestUri = "https://secure.bill4time.com/b4t2/"
$uriCompanyID = "hlopezlaw"
$uriLogin = "hlopez@hlopezlaw.com"
$uriPwd = "Hlg101480@"
$companyIdFragment = "firmCode";
$userIdFragment = "username";
$passwordIdFragment = "password";
$acceptTermsInputFragment = "weblogin_visitor_accept_terms"
$buttonIdFragment = "weblogin_submit";
# Show the instance of IE. With this line commented out, we get a "headless" browsing experience with no visible window or rendering. 
$ie.visible = $true
$ie.silent = $true
$ie.navigate($requestUri)
while($ie.Busy) { Start-Sleep -Milliseconds 10000 }
#
# Get Home Page elements and perform login
#$doc = $ie.Document
($ie.document.getElementsByName("firmCode") |select -first 1).value = $uriLogin
#$userNameField = $doc.Document.IHTMLDocument3_getElementsByTagName("firmCode")
#$userNameField.value = "$uriLogin"
#$doc.getElementsByTagName("input") | % {
#    if ($_.id -ne $null){
#        if ($_.id.Contains($buttonIdFragment)) { $btn = $_ }
#        if ($_.id.Contains($companyIdFragment)) { $company = $_ }
#        if ($_.id.Contains($userIdFragment)) { $user = $_ }
#        if ($_.id.Contains($passwordIdFragment)) { $pwd = $_ }
#    }
#}
#$company.value = $uriCompanyID
#$user.value = $uriLogin
#$pwd.value = $uriPwd
#$btn.disabled = $false
#$btn.click()
Write-Verbose "Login Complete"
