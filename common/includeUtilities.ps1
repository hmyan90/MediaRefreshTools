# -------------------------------------------------------------------------------------------
# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License. See License.txt in the project root for license information.
# -------------------------------------------------------------------------------------------

# Create this script to support variable path since using module statement cannot include any variables

$ScriptBody = "using module $PSScriptRoot\utilities.psm1";

$Script = [ScriptBlock]::Create($ScriptBody);
. $Script