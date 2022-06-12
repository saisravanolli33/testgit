try{$ErrorActionPreference='SilentlyContinue'; $runs = az pipelines runs list --org https://dev.azure.com/office --project crates --status completed |ConvertFrom-Json # --result failed
$releases = az pipelines release list --org https://dev.azure.com/office --project crates | convertfrom-json}catch{'az loaded'}
$Latestruns = $runs | where{$($(get-date)-[datetime]$_.finishtime).totaldays-lt 1}
$Latestreleases = $releases | where{$($(get-date)-[datetime]$_.modifiedon).totaldays-lt 1}
$h = @{Authorization="BasicOmY0ZmRsM3JlbHJibHJ6bGh3aWVldmQyZmZnM3pjN3l5ZXBvbGczM3ZlbWp1ZnM3eTVmcHE=";ContentType="application/json"}
$outpath = "C:\Users\v-dsripuram\Desktop" 
$tempoutpath = "C:\Users\v-dsripuram\Desktop\out.csv"
$tempout = [pscustomobject][ordered]@{Pipeline=$null;Type=$null;Message=$null;'Duration (min)'=$null;LastRun=$null;Status=$null}

$emailhtml = "
<style>
table, th, td {
  border: 1px solid black;
  border-collapse: collapse;
}</style>
<p>Hello,</p><p></p></p><p>Below is summary of Pipeline status:</p><table><tr><th>Pipeline</th><th>Type</th><th>Message</th><th>Duration (min)</th><th>LastRun</th><th>Status</th></tr>"



foreach($release in $Latestreleases){
    $body = @();$endtime= $null;$duration = $null
    $pipeline=$release.releasedefinition.name;$name = $release.name;$tempout.Pipeline = $pipeline;$tempout.Type=$name
	irm $release.logscontainerurl -method get -headers $h -outfile "$outpath\$name.zip"
	Expand-archive "$outpath\$name.zip" -DestinationPath $outpath\$name -Force ;
	Get-ChildItem "$outpath\$name" -Recurse -File | where{$(gc $_.fullname) -match 'Error:'} |%{$Body += gc $_.fullname |findstr ']Error:'}
	Get-ChildItem "$outpath\$name" -Recurse -File | %{$endtime = $((gc $_.fullname)[-1].split('#')[0])}
	$duration = if($endtime-ne$null){[math]::round(([datetime]$endtime-[datetime]$release.createdon).totalminutes,0);$status='Success'}else{0;$status = 'Pending for Approval'}
    If($body.count -ne 0){$emailhtml+="<tr><td>$pipeline</td><td>$name</td><td>$($body -join "`n")</td><td>$duration</td><td>$([datetime]$release.createdon)</td><td>Failed</td></tr>";$tempout.Status='Failed'}
    else{$emailhtml+="<tr><td>$pipeline</td><td>$name</td><td>Success</td><td>$duration</td><td>$([datetime]$release.createdon)</td><td>$status</td></tr>";$tempout.Status=$status}
    $tempout.Message=$($body -join "`n");$tempout.'Duration (min)'=$duration;$tempout.LastRun=$([datetime]$release.createdon);$tempout | Export-Csv $tempoutpath -NoTypeInformation -Append -Force
    Remove-Item $outpath\$name -Recurse -Force
    Remove-Item "$outpath\$name.zip" -Recurse -Force
}


foreach($run in $Latestruns){
    $pipeline = $run.definition.name;$name = $run.buildnumber;$tempout.Pipeline = $pipeline;$tempout.Type=$name
	$logs = Invoke-RestMethod $run.logs.url -headers $h
    $logurls = ($logs.value | where{$($(get-date)-[datetime]$_.createdon).totaldays-lt 1}).url
    $body=@();$temp=$null
    $logurls |%{$temp=irm $_ -headers $h;
        $temp1=@();$temp=$temp | findstr "##[error";
        foreach($t in $temp){if($temp1-notcontains$t){$temp1+=$t}};
        if($temp1.count-ne0){$body+=$temp1}
    }
    $duration = [math]::round(([datetime]$run.finishTime-[datetime]$run.startTime).TotalMinutes,0)
	If($body.count -ne 0){$emailhtml+="<tr><td>$pipeline</td><td>$name</td><td>$($body -join "`n")</td><td>$duration</td><td>$([datetime]$run.startTime)</td><td>Failed</td></tr>";$tempout.Status='Failed'}
    else{$emailhtml+="<tr><td>$pipeline</td><td>$name</td><td>Success</td><td>$duration</td><td>$([datetime]$run.startTime)</td><td>Success</td></tr>";$tempout.Status='Success'}
    $tempout.Message=$($body -join "`n");$tempout.'Duration (min)'=$duration;$tempout.LastRun=$([datetime]$run.startTime);$tempout | Export-Csv $tempoutpath -NoTypeInformation -Append -Force
}
$emailhtml+="</table>"
$emailhtml > $outpath\out.htm

$EmailTenantID = "72f988bf-86f1-41af-91ab-2d7cd011db47"
$EmailClientID = "b7dc09ee-128c-434c-beec-f4d5e476400f"
$EmailAuthenticationUser = "edibot@microsoft.com"
$EmailAuthenticationPass =(az keyvault secret show --name EdiAADKey --vault-name oexp-key-vault | ConvertFrom-Json).value | ConvertTo-SecureString -AsPlainText -Force
$recipient = "v-dsripuram@microsoft.com"
<#
$authparams = @{
    ClientId     = $EmailClientID
    TenantId     = $EmailTenantID
    ClientSecret = $EmailAuthenticationPass
}
$auth = Get-MsalToken @authParams
$authorizationHeader = @{Authorization = $auth.CreateAuthorizationHeader()}
Write-Host "Sending welcome email"
# Create message body and properties and send
$MessageParams = @{
          "URI"         = "https://graph.microsoft.com/v1.0/users/$MsgFrom/sendMail"
          "Headers"     = $authorizationHeader
          "Method"      = "POST"
          "ContentType" = 'application/json'
          "Body" = (@{
                "message" = @{
                    "subject" = $MsgSubject
                    "body"    = @{
                        "contentType" = 'HTML' 
                         "content"     = $emailhtml}  
                    "toRecipients" = @(
                        @{
                        "emailAddress" = @{"address" = $recipient }
                        })       
                }
          }) | ConvertTo-JSON -Depth 6
   }   # Send the message
Invoke-RestMethod @Messageparams 
#>
