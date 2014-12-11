#
# Hyper-V Scopes PowerShell module (for use with Azman authentication)
#
# Timothy Baldock <tb@entropy.me.uk>
#

Function Set-VMScope
{
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "Medium")]
    Param (
        [string]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true)]
        [Alias("ElementName")]
        [ValidateNotNullOrEmpty()]
            $VMName,
        [string]
            $Scope,
        [switch]
            $Force
    )
    Process {
        # Get ComputerSystem path
        $testsystem = Get-WmiObject -Namespace "root\virtualization" -Class MSVM_ComputerSystem -Filter "ElementName = '$($VMName)'"
        # Old scope for shouldprocess prompt
        $oldscope = (Get-WmiObject -Namespace "root\virtualization" -Class MSVM_VirtualSystemGlobalSettingData -Filter "ElementName = '$($VMName)'").ScopeOfResidence

        if ($Force -or $pscmdlet.ShouldProcess("VM: $($VMName), Old Scope: $($oldscope), New Scope: $($Scope)")) {
            # Get system global setting data
            $globalsettings = Get-WmiObject -Namespace "root\virtualization" -Class MSVM_VirtualSystemGlobalSettingData -Filter "ElementName = '$($VMName)'"
            $globalsettings.ScopeOfResidence = $Scope


            # Populate input parameters
            $methods = ([WMIClass] "root\virtualization:MSVM_VirtualSystemManagementService")
            $InParams = $methods.PSBase.GetMethodParameters("ModifyVirtualSystem")

            Write-Debug "InParams before modification:"
            $InParams | Write-Debug

            $InParams.ComputerSystem = $testsystem.Path.Path
            $InParams.SystemSettingData = $globalsettings.GetText(1)

            Write-Debug "InParams after modification:"
            $InParams | Write-Debug

            $MC = Get-WmiObject -Namespace "root\virtualization" -Class "MSVM_VirtualSystemManagementService"

            $returnval = $MC.PSBase.InvokeMethod("ModifyVirtualSystem", $InParams, $Null).ReturnValue | Write-Output

            New-Object Object |
                Add-Member NoteProperty VMName $VMName -PassThru |
                Add-Member NoteProperty Scope (Get-WmiObject -Namespace "root\virtualization" -Class MSVM_VirtualSystemGlobalSettingData -Filter "ElementName = '$($VMName)'").ScopeOfResidence -PassThru |
                Add-Member NoteProperty ReturnValue $returnval -PassThru |
                Write-Output
        }
    }
}

Function Get-VMScope
{
    [CmdletBinding()]
    Param (
        [string]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true)]
        [Alias("ElementName")]
        [ValidateNotNullOrEmpty()]
            $VMName
    )
    Process {
        # Old scope for shouldprocess prompt
        Get-WmiObject -Namespace "root\virtualization" -Class MSVM_VirtualSystemGlobalSettingData -Filter "ElementName like '$($VMName)'" | Select ElementName,ScopeOfResidence | Write-Output
    }
}
