#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll"
$url="https://hlopezlaw.sharepoint.com/sites/HLopezLaw"
$ReportOutput="C:\Temp\TermStoreData.csv"
#Array to Hold Result - PSObjects
$ResultCollection = @()
$cred =  Get-Credential

$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($url)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($cred.username, $cred.Password)  
$clientContext.Credentials = $credentials 

############# Get your tenant Taxonomy details ##################

$MMS = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($clientcontext)
$clientContext.Load($MMS)
$clientContext.ExecuteQuery()
$TermStores = $MMS.TermStores
$clientContext.Load($TermStores)
$clientContext.ExecuteQuery()
#Bind to Term Store
$TermStore = $TermStores[0]
Write-Host "#################################      Taxonomy on tenant which contain the Site collection   #################################"
Write-Host "#################################      $url    #################################"
$clientContext.Load($TermStore)
$clientContext.ExecuteQuery()
""
$Groups = $TermStore.Groups
$clientContext.Load($groups)
$clientContext.ExecuteQuery()

for ($i = 0; $i -lt $Groups.Count; $i++)
{ 
write-host "For the Group: "$groups[$i].Name " / " $groups[$i].id
$groupname =$groups[$i]
$termsets = $groupname.TermSets
$clientContext.Load($termsets)
$clientContext.ExecuteQuery()
    foreach ($termset in $termsets ){
    write-host "  The TermSet:"$termset.Name " / " $termset.id  -ForegroundColor DarkYellow
	            $TermData = new-object PSObject
                $TermData | Add-member -membertype NoteProperty -name "TermSet" -Value "TermSet"
                $TermData | Add-member -membertype NoteProperty -name "Name" -Value $Termset.Name
                $TermData | Add-member -membertype NoteProperty -name "ID" -Value $termset.id   
                $ResultCollection += $TermData
				$terms = $termset.Terms
				$clientContext.Load($terms)
				$clientContext.ExecuteQuery()
				foreach ($term in $terms ){
				write-host  "    The Term ( level 0): "$term.Name  " / " $term.id  -ForegroundColor Gray
								$TermData = new-object PSObject
								$TermData | Add-member -membertype NoteProperty -name "TermSet" -Value "Level 0"
								$TermData | Add-member -membertype NoteProperty -name "Name" -Value $term.Name
								$TermData | Add-member -membertype NoteProperty -name "ID
" -Value $term.id   
								$ResultCollection += $TermData
								$items = $term.Terms
								$clientContext.Load($Items)
								$clientContext.ExecuteQuery()
								foreach ( $term_level1 in $items ){
								Write-Host "      Level 1:" $term_level1.Name " / "$term_level1.ID -ForegroundColor cyan
									$TermData = new-object PSObject
									$TermData | Add-member -membertype NoteProperty -name "Group" -Value "Level 1"
									$TermData | Add-member -membertype NoteProperty -name "GroupName" -Value $term_level1.Name
									$TermData | Add-member -membertype NoteProperty -name "GroupId" -Value $term_level1.ID   
													$ResultCollection += $TermData
									#$term_level1.TermsCount
									$items2 = $term_level1.Terms
									$clientContext.Load($Items2)
									$clientContext.ExecuteQuery()

									foreach ( $term_level2 in $items2 ){
									# $term_level2.TermsCount
									Write-Host "        Level 2:" $term_level2.Name " / "$term_level2.ID -ForegroundColor green
										$items3 = $term_level2.Terms
										$clientContext.Load($Items3)
										$clientContext.ExecuteQuery()
										foreach ( $term_level3 in $items3 ){
										#$term_level3.TermsCount
										Write-Host "            Level 3:" $term_level3.Name " / "$term_level3.ID -ForegroundColor yellow
											$TermData = new-object PSObject
											$TermData | Add-member -membertype NoteProperty -name "Division" -Value "Level 2"
											$TermData | Add-member -membertype NoteProperty -name "Level2Name" -Value $term_level3.Name
											$TermData | Add-member -membertype NoteProperty -name "Level2id" -Value $term_level3.ID   
											$ResultCollection += $TermData	

									$items4 = $term_level3.Terms
									$clientContext.Load($Items4)
									$clientContext.ExecuteQuery()
									foreach ( $term_level4 in $items4 ){

										#$term_level4.TermsCount
										Write-Host "              Level 4:" $term_level4.Name " / "$term_level4.ID -ForegroundColor blue
										$TermData = new-object PSObject
										$TermData | Add-member -membertype NoteProperty -name "Group" -Value "Level 4"
										$TermData | Add-member -membertype NoteProperty -name "TermSet" -Value $term_level4.Name
										$TermData | Add-member -membertype NoteProperty -name "Term" -Value $term_level4.ID   
										$ResultCollection += $TermData
										$items5 = $term_level4.Terms
										$clientContext.Load($Items5)
										$clientContext.ExecuteQuery()
										foreach ( $term_level5 in $items5 ){
										#$term_level5.TermsCount
										Write-Host "                Level 5:" $term_level5.Name" / "$term_level5.ID -ForegroundColor red
										} 
									}  
								}
							}
						}
					}
				}
			}
    #Export Results to a CSV File
    $ResultCollection | Export-csv $ReportOutput -notypeinformation
