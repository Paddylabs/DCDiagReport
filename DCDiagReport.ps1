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
  V1.0:         Initial Development
  V2.0:         Input from brettmillerb, fixed formatting, added Requires statement and splatting.
  V2.1:         Added a Table and formatting for improved readability of the data.
  V2.2:         Added some conditional formatting so any failures stand out.
#>

#Requires -Modules ImportExcel

$SB = {
    $DCDIAG = dcdiag /v
    $DCDiagResults = New-Object System.Object
    $DCDiagResults | Add-Member -name Server -Value $env:COMPUTERNAME -Type NoteProperty -Force

    Foreach ($Entry in $DCDIAG) {
        Switch -Regex ($Entry) {
            "Starting" {
                $Testname = ($Entry -replace ".*Starting test: ").Trim()
            }
            "passed|failed" {
                If ($Entry -match "passed") {
                    $TestStatus = "Passed"
                }
                Else {
                    $TestStatus = "failed"
                }
            }
        }

        If ($TestName -ne $null -and $TestStatus -ne $null) {
            $DCDiagResults | Add-Member -Type NoteProperty -name $($TestName.Trim()) -Value $TestStatus -Force
        }
    }

    $DCDiagResults
}

$DCs = Get-ADDomainController -filter * | Select-Object Name

$Session = New-PSSession -ComputerName $DCs.Name

if ($Session) {
    $invokeCommandSplat = @{
        ErrorAction = 'SilentlyContinue'
        Session     = $Session
        ScriptBlock = $SB
    }

    $exportExcelSplat = @{
        Path            = "DCDiagReport.xlsx"
        BoldTopRow      = $true
        AutoSize        = $true
        FreezeTopRow    = $true
        WorkSheetname   = "DCDiag"
        TableName       = "DCDiagTable"
        TableStyle      = "Medium6"
    }

    Invoke-Command @invokeCommandSplat | Export-Excel @exportExcelSplat  -ConditionalText @(
        New-ConditionalText -Text Failed -ConditionalTextColor Darkred -BackgroundColor LightPink
        )
}

Remove-PSSession $Session
