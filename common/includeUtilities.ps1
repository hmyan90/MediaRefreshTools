# using module statement cannot include any variables. Its values must be static.
# So we have to create this script to support variable path

$scriptBody = "using module $PSScriptRoot\utilities.psm1";

$script = [ScriptBlock]::Create($scriptBody);
. $script