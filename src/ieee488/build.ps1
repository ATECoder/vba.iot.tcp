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
    # this does not work: getting an exception
    # Exception calling "ReadKey" with "1" argument(s): "Cannot read keys when either application does not have a console or when console input has been redirected from a 
    # file. Try Console.Read."
    loginfo( "Press any key" )
	do{ $x = [console]::ReadKey() } while( $x.Key -ne "" )	
}

# -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  
# Summary:  Removes reference from the specified workbook.
#
# Parameters:
# 
# $workbook - the workbook from which the reference REFNAME needs to be removed
# REFNAME   - The name of the referenced workbook that needs to be removed
# -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  
Function RemoveReference( $workbook )
{
	ForEach ($ref in $workbook.VBProject.References ) {
		if ( $ref.Name -eq $REFNAME )
		{
			$workbook.VBProject.References.Remove( $ref )
			break
		}
	}
}

# -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  
# Summary:  Adds a reference to the specified workbook.
#
# Parameters:
# 
# $workbook - the workbook to which the reference needs to be added
# REFFILE   - the file name of the referenced workbook to add, e.g., "cc.isr.core.xlsm"
# -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  
Function AddReference( $workbook )
{
	RemoveReference( $workbook ) 
	$workbook.VBProject.References.AddFromFile( $REFFILE )
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
	
    LogInfo( "Removing refernece " + $REFFILE )

    $PRMYPATH = [IO.Path]::Combine($BUILD_DIRECTORY, $PRMYFILE)
	$PRMYBOOK = $excel.Workbooks.Open($PRMYPATH)
    LogInfo( "   Opened " + $excel.Workbooks.Item($PRMYFILE).Name )

	$REFPATH = [IO.Path]::Combine($BUILD_DIRECTORY, $REFFILE)
    $REFFILE = $REFFILE 
	RemoveReference( $PRMYBOOK )

	$PRMYBOOK.Save()

	$PRMYBOOK.SaveAs($PRMYPATH, $XL_FILE_FORMAT_MACRO_ENABLED)
    $PRMYBOOK.Close()
    LogInfo( "   closed")
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

$SOURCE = [IO.Path]::Combine($CWD, "..\core\cc.isr.core.testing.md")
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

$SOURCE = [IO.Path]::Combine($CWD, "..\winsock\cc.isr.winsock.testing.md")
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

$SOURCE = [IO.Path]::Combine($CWD, "cc.isr.ieee488.testing.md")
Try {
	echo coping $SOURCE to $BUILD_DIRECTORY
	Copy-Item $SOURCE -Destination $BUILD_DIRECTORY
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
	RemoveWorkbookReference
	
    $PRMYFILE = "cc.isr.ieee488.xlsm"
	RemoveWorkbookReference

# Remove cc_isr_winsock references

    $REFFILE = "cc.isr.winsock.xlsm"
    $REFNAME = "cc_isr_Winsock"

    $PRMYFILE = "cc.isr.ieee488.xlsm"
	RemoveWorkbookReference

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

LogInfo( "project built" )
$z = Read-Host "Press enter to exit: "