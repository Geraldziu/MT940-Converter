<#
 .Synopsis
  Processes a bank statement in the MT940 format from the ING bank into the JSON format.

 .Description
  Processes a bank statement in the MT940 format from the ING bank into the JSON format.

 .Parameter filelocation
  The path to the file in MT940 format.

 .Example
   # Process a file with bank statement.
   Process-Ing-MT940 -filelocation "C:\users\xxxxx\Documents\statement.sta"
#>
Function Process_Ing_MT940 {
    Param(
        [Parameter(Mandatory=$true)][string]$fileLocation
    )
    Process
    {
     Try{

        $SR = New-Object System.IO.StreamReader($filelocation, [Text.Encoding]::GetEncoding("ibm852"))
        $lineNumber = 1
        $statementHead = @{}
        [System.Collections.ArrayList]$StatementItem = @()
        $Item = @{}
        While (($SRLine = $SR.ReadLine()) -ne $null)
        {
            IF(($SRLine -match ":61" -and $LastSRLine -match "~6[2-3]") -or ($SRLine -match ":61" -and $LastSRLine -match "~34") -or ($SRLine -match ":62F:"))
                {$StatementItem.Add($Item) | Out-Null}
            IF($SRLine -match "^:25:" ) {
                $statementHead.Add('Account number',$SRLine.Substring(5,24))
                continue
            }
            IF($SRLine -match "^:28C:" ) {
                $statementHead.Add("Statement number",$SRLine.Substring(5,$SRLine.Length-5))
                Continue
                }
            IF($SRLine.StartsWith(":60F")) {
                $statementHead.Add("Date", "20" + $SRLine.Substring(6,2) + "-" + $SRLine.Substring(8,2) + "-" + $SRLine.Substring(10,2))
                $statementHead.Add("Currency", $SRLine.Substring(12,3))
                $statementHead.Add("Balance", [float]$SRLine.Substring(15,$SRLine.Length-15).Replace(",","."))
                Continue
                }
            IF($SRLine -match ":61:"){
                $Item = @{
                "Date" = "20" + $SRLine.Substring(4,2) + "-" + $SRLine.Substring(6,2) + "-" + $SRLine.Substring(8,2);                "Accounting date" = "20" + $SRLine.Substring(4,2) + "-" + $SRLine.Substring(10,2) + "-" + $SRLine.Substring(12,2);                "Sign" = $sign = IF($SRLine.Substring(14,1) -eq "C"){1}Else{-1};                "Value" = [float]$SRLine.Substring(15,$SRLine.IndexOf(",")-12).Replace(",",".")}
                Continue
            }
            IF($SRLine -match "~2[0-5]"-and $SRLine.Length -gt 3 -and $SRLine.LastIndexOf("~") -lt 3){
                $Item = $Item += @{
                "Description" = $SRLine.Substring(3,$SRLine.Length-3)}
                Continue
            }
            IF($SRLine -match "~3[0-1]"-and $SRLine.Length -gt 6){
                $Item = $Item += @{
                "MyAccountNr" = $SRLine.Substring(3,$SRLine.LastIndexOf("~")-3) + $SRLine.Substring($SRLine.LastIndexOf("~")+3,$SRLine.Length-$SRLine.LastIndexOf("~")-3)
                }
                Continue
            }
            IF($SRLine -match "~3[2-3]"-and $SRLine.Length -gt 6){
                IF($SRLine -match "~33") {
                    $Item = $Item += @{
                    "Name" = $SRLine.Substring(3,$SRLine.LastIndexOf("~")-3) + $SRLine.Substring($SRLine.LastIndexOf("~")+3,$SRLine.Length-$SRLine.LastIndexOf("~")-3)}
                    }
                Else {
                    $Item = $Item += @{
                    "Name" = $SRLine.Substring(3,$SRLine.Length-3)}
                    }
                Continue
            }
            IF($SRLine -match "~38"-and $SRLine.Length -gt 6){
                $Item = $Item += @{
                "AccountNr" = $SRLine.Substring(3,$SRLine.Length-3)
                }
                Continue
            }
            IF($SRLine -match "~6[2-3]"-and $SRLine.Length -gt 6){
                $Item.Name = $Item.Name + $SRLine.Substring(3,$SRLine.Length-3)
                Continue
            }
            $LastSRLine = $SRLine
            $lineNumber++
        }
        $SR.Dispose()
        $statementJSON = @{}
        $statementJSON.Add("Transactions", $StatementItem)
        $statementJSON.Add("Account information", $statementHead)
        $statementJSON | sort | ConvertTo-Json -Depth 10
    }
    Catch{
    Write-Debug $SRLine
    }
  }
}

Export-ModuleMember -Function Process_Ing_MT940