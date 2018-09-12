<#
  .SYNOPSIS
  Runs DCDiag on all Domain Controllers and exports results to an Excel Spreadsheet.

  .DESCRIPTION
  Runs DCDiag on all Domain Controllers and exports a 'Passed' or 'Failed' value for each test on each DC to a nicely formatted and filtered Excel Spreadsheet.

  .PARAMETER
  None
  
  .EXAMPLE
  DCDiagReport.ps1

  .INPUTS
  None

  .OUTPUTS
  DCDiagReport.xlsx

  .NOTES
  Author:        Patrick Horne
  Creation Date: 12/09/18
  Requires:      ImportExcel Module

  Change Log:
  V1:         Initial Development
#>

Import-Module ImportExcel

$SB = { 

  $DCDIAG = dcdiag /v
  $DCDiagResults = New-Object System.Object
  $DCDiagResults | Add-Member -name Server -Value $env:COMPUTERNAME -Type NoteProperty -Force

    Foreach ($Entry in $DCDIAG)
      
      {
      
      Switch -Regex ($Entry) 
      
        {
        
        "Starting" {$Testname = ($Entry -replace ".*Starting test: ").Trim()}
        "passed|failed" {If ($Entry -match "passed") {$TestStatus = "Passed"} Else {$TestStatus = "failed"}}
        
        }
     
        If ($TestName -ne $null -and $TestStatus -ne $null) 

            {

            $DCDiagResults | Add-Member -Type NoteProperty -name $($TestName.Trim()) -Value $TestStatus -Force

            }

        }

    $DCDiagResults

}

$DCs = Get-ADDomainController -filter * | Select-Object Name

$Session = New-PSSession -ComputerName $DCs.Name

if($Session) { Invoke-Command -Session $Session -ScriptBlock $SB -ErrorAction SilentlyContinue | Export-Excel DCDiagReport.xlsx -FreezeTopRow -BoldTopRow -AutoFilter -WorkSheetname "DCDiag" -AutoSize  }

Remove-PSSession $Session
