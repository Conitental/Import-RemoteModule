Function Import-RemoteModule {
    <#
    .SYNOPSIS
    Function to remotely use powershell modules from different computers.

    Credit to https://github.com/ztrhgf/useful_powershell_functions/blob/master/SCCM/Connect-SCCM.ps1

    .DESCRIPTION
    This function creates a PSSession to a remote machine and imports a module as a whole or specific commands from it.

    .PARAMETER ComputerName
    The name of the remote computer.

    .PARAMETER Module
    The name of the module to be used.

    .PARAMETER CommandName
    (Optional)

    The command or commands that need to be imported.

    .PARAMETER Force
    (Optional)

    Recreate the session even if we have a correctly configured session already

    .EXAMPLE
    Import-RemoteModule -ComputerName 'Machine01' -Module 'ActiveDirectory'

    .EXAMPLE
    Import-RemoteModule -ComputerName 'Machine02' -Module 'ConfigurationManager' -CommandName 'Get-CMDistributionPointInfo', 'Get-CMContentDistribution'
    #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$ComputerName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$Module,

        [String[]]$CommandName,

        [Switch]$Force
    )

    Write-Verbose "Check for existing session with $ComputerName"
    $ExistingSession = Get-PSSession -ComputerName $ComputerName

    # Assume we need to create the session
    $NeedToCreateSession = $true

    If($ExistingSession) {
        Write-Verbose "Found existing session with $ComputerName with state $($ExistingSession.State)"

        If($ExistingSession.State -eq 'Opened') {
            Write-Verbose "Check session for module $Module"

            Try {
                Get-Module -Name $Module -ErrorAction Stop | Out-Null
                Write-Verbose "Module $Module found in session"
            } Catch {
                Write-Verbose "Could not find module $Module in session. Remove session and recreate"

                Remove-PSSession -Session $ExistingSession
                $NeedToCreateSession = $true
            }

            # Check commands if the module was found and we have command input
            If($CommandName) {
                Write-Verbose "Check session for required commands"

                # Check later if we could not find any command and thus need to reinitiate the session
                $CommandNotFound = $false

                # Using Foreach-Object in this case to keep singular naming of parameter
                $CommandName | ForEach-Object {
                    Try {
                        Write-Verbose "Check session for command $_"
                        Get-Command -Name $_ | Out-Null
                        Write-Verbose "Command $_ found in session"
                    } Catch {
                        Write-Verbose "Could not find command $_ in session"
                        $CommandNotFound = $true
                    }
                }

                If($CommandNotFound) {
                    Write-Verbose "At least one command could not be found. The session will be removed and reinitiated"
                    Remove-PSSession -Session $ExistingSession
                    $NeedToCreateSession = $true
                } Else {
                    Write-Verbose "All command have been found"
                }
            }
        } Else {
            Write-Verbose "Session is not in opened state and will be removed"
            Remove-PSSession -Session $ExistingSession
            $NeedToCreateSession = $true
        }

        # If the session is still open at this point we can assume correct configuration
        If($ExistingSession.State -eq 'Opened') {
            Write-Verbose "Session to $ComputerName is configured correctly"

            If($Force) {
                Write-Verbose "Force parameter is given. Session will be removed and recreated"
                Remove-PSSession -Session $ExistingSession

                $NeedToCreateSession = $true
            } Else {
                $NeedToCreateSession = $false
            }
        }
    }

    # Exit function if we do not need to create any new session
    If($NeedToCreateSession -eq $false) {
        Return
    }

    Write-Verbose "Try to create PSSession to $ComputerName and import module $Module"

    Write-Verbose "Try to resolve $ComputerName"
    If(Resolve-DnsName -Name $ComputerName -ErrorAction SilentlyContinue) {
        Write-Verbose "Success"
    } Else {
        Write-Error "Could not resolve target name: $ComputerName"
        Return
    }    

    Write-Verbose "Check network availability of $ComputerName"
    If(Test-Connection -ComputerName $ComputerName -Quiet) {
        Write-Verbose "Success"
    } Else {
        Write-Error "$ComputerName is offline"
        Return
    }


    $PSSession = New-PSSession -ComputerName $ComputerName

    Write-Verbose "Try to import module $Module"
    Try {
        Invoke-Command -Session $PSSession -ErrorAction 'Stop' -ScriptBlock {
            $ErrorActionPreference = 'Stop'

            Try {
                Import-Module $using:Module
            } Catch {
                Throw "Could not import module $($using:Module)"
            }
        }

        Write-Verbose "Success"
    } Catch {
        Write-Error "Could not successfully import module $Module from $ComputerName - Invokation failed"
        Return
    }

    Write-Verbose "Importing session"
    Try {
        Import-Module -PSSession $PSSession -Cmdlet $CommandName -Global -Force
        Write-Verbose "Successfully imported session"
    } Catch {
        Write-Error "Failed to import session"
    }
}

Export-ModuleMember -Function Import-RemoteModule
