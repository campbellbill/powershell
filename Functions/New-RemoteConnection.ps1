<#
.Synopsis
   Creates a remote connection to server specified with provided credential
.DESCRIPTION
   Function has alias of NRC that can also be used to call. Additionally supports
   Pipeline input of ComputerName, and supports -Verbose and -WhatIf
   Tip:
   Add any additional helpful commands to output after remote connection is established. 
.EXAMPLE
   New-RemoteConnection
.EXAMPLE
   nrc -ComputerName "ServerName"
.EXAMPLE
   New-RemoteConnection -ComputerName "ServerName" -credential (get-credential)
#>
function New-RemoteConnection
{
    [CmdletBinding( SupportsShouldProcess=$true )]
    [Alias("nrc")]
    Param
    (
        # Computername of system to connect to
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $ComputerName,

        # Credential to use when connecting
        [Parameter(Mandatory=$false)]
        $Credential
    )

    Begin
    {
        if ($ComputerName -eq $null){
        #prompts for ComputerName to connect to if name is not supplied
            $ComputerName = Read-Host "Enter Computer Name to Connect to "
        }
        if ($credential -eq $null){
        #prompts for Credentials if none are supplied
            $credential = Get-Credential -Credential ""
        }
    }
    Process
    {
        #test to see if the ComputerName is responding
        if (Test-Connection -ComputerName $ComputerName -Count 1 -ErrorAction SilentlyContinue){
            Write-Verbose "Attemtping to connect to $ComputerName"
            try{
            
                    if ($pscmdlet.ShouldProcess("$ComputerName", "Opening Remote PowerShell Session"))
                    {
                        #open new session to ComputerName
                        Enter-PSSession -ComputerName $ComputerName -Credential $credential -ErrorAction Stop
                    }
                    #output information on a snapins that might be handy
                    Write-Host "Add-PsSnapin Microsoft.Exchange.Management.PowerShell.E2010" -ForegroundColor Green
                }

            catch{
                Write-Error "Error when connecting to $Computername Details Below : `n $_"
            }
        } 

        else {
            Write-Error "Server $ComputerName is unreachable. Check that server is responding"
        }
    }
    End
    {
    }
}



