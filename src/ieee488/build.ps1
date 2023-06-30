# ----------------------------------------------------------------------
# build 
#
# PURPOSE: build a release version of this top level workbook by
#          adding the referenced workbooks.
#
# CALLING SCRIPT:
#
#  ."build.ps1"
#
# ----------------------------------------------------------------------

# ----------------------------------------------------------------------
# VARIABLES

$CWD = (Resolve-Path .\).Path
$BUILD_DIRECTORY = [IO.Path]::Combine($CWD, "bin")
$XL_FILE_FORMAT_MACRO_ENABLED = 52

# END VARIABLES
# ----------------------------------------------------------------------

# ----------------------------------------------------------------------
# FUNCTIONS

Function LogInfo($message)
{
    Write-Host $message -ForegroundColor Gray
}

Function LogEmptyLine()
{
    echo ""
}

function AwaitKeyPress()
{
	do{ echo "Press any key";$x = [System.Console]::ReadKey() } while( $x.Key -ne "" )	
}

# -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  
# Summary:  Removes reference to the reference workbook from the primary work book 
#           and save the primary workbook.
#
# Parameters:
# 
# PRMYFILE - the file name of the primary workbook, e.g., "cc.isr.winsock.xlsm"
# REFFILE  - the file name of the referenced workbook, e.g., "cc.isr.core.xlsm"
# REFNAME  - The new of the referenced workbook, e.g., "cc_isr_Core"
# -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  
function RemoveWorkbookReference()
{
	$PRMYPATH = [IO.Path]::Combine($BUILD_DIRECTORY, $PRMYFILE)
	$PRMYBOOK = $excel.Workbooks.Open($PRMYPATH)
    echo "Opened " + $excel.Workbooks.Item($PRMYFILE).Name

	$REFPATH = [IO.Path]::Combine($BUILD_DIRECTORY, $REFFILE)

    $REFFILE = $REFFILE 
	RemoveReference( $PRMYBOOK )
	$PRMYBOOK.Save()

	$PRMYBOOK.SaveAs($PRMYPATH, $XL_FILE_FORMAT_MACRO_ENABLED)
    $PRMYBOOK.Close()
}

# END FUNCTIONS
# ----------------------------------------------------------------------


# ----------------------------------------------------------------------
# SCRIPT ENTRY POINT

# create the build directory if new 

MkDir -Force $BUILD_DIRECTORY > $null

# Copy all workbooks to the build directory

$SOURCE = [IO.Path]::Combine($CWD, "..\core\cc.isr.core.xlsm")
Try {
	echo coping $SOURCE to $BUILD_DIRECTORY
	copy-item $SOURCE -destination $BUILD_DIRECTORY
}
Catch {
    echo $_.Exception.Message
    return
}

$SOURCE = [IO.Path]::Combine($CWD, "..\winsock\cc.isr.winsock.xlsm")
Try {
	echo coping $SOURCE to $BUILD_DIRECTORY
	copy-item $SOURCE -destination $BUILD_DIRECTORY
}
Catch {
    echo $_.Exception.Message
	return
}

$SOURCE = [IO.Path]::Combine($CWD, "cc.isr.ieee488.xlsm")
Try {
	echo coping $SOURCE to $BUILD_DIRECTORY
	copy-item $SOURCE -destination $BUILD_DIRECTORY
}
Catch {
    echo $_.Exception.Message
	return
}

$SOURCE = [IO.Path]::Combine($CWD, "testing.md")
Try {
	echo coping $SOURCE to $BUILD_DIRECTORY
	copy-item $SOURCE -destination $BUILD_DIRECTORY
}
Catch {
    echo $_.Exception.Message
	return
}

# Open excel as hidden

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$AutoSecurity = $Excel.AutomationSecurity
$Excel.AutomationSecurity = 3
Try {
	
	$excel.DisplayAlerts = $false

# Remove the references from each workbook starting with the workbook that
# has the fewest references.

# Remove cc_isr_core references

    $REFNAME = "cc_isr_Core"
    $REFFILE = "cc.isr.core.xlsm"
	
    $PRMYFILE = "cc.isr.winsock.xlsm"
	RemoveWorkbookReference()
	
    $PRMYFILE = "cc.isr.ieee488.xlsm"
	RemoveWorkbookReference()

# Remove cc_isr_winsock references

    $REFFILE = "cc.isr.winsock.xlsm"
    $REFNAME = "cc_isr_Winsock"

    $PRMYFILE = "cc.isr.ieee488.xlsm"
	RemoveWorkbookReference()

}

Catch {
    echo $_.Exception.Message
    return
}
Finally{

	$excel.DisplayAlerts = $true
    $Excel.AutomationSecurity = $AutoSecurity
	$excel.Quit()
}

echo "project built"
LogInfo "Project built"
AwaitKeyPress()
