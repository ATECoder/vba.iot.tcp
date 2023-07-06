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

Function LogError($message)
{
    Write-Host $message -ForegroundColor Red
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
	
    LogInfo( "Removing reference " + $REFFILE )

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

# -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  
# Summary:  Copies a specified file to the build directory.
#
# Parameters:
# 
# SOURCE           - the path of the source file
# BUILD_DIRECTORY  - the path of the build directory
# -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  
function CopyToBuildDirectory( $sourcePath )
{

	Try { 

        $path = (Resolve-Path $sourcePath).Path

		LogInfo( "coping " + $path + " to " + $BUILD_DIRECTORY )
		copy-item $path -destination $BUILD_DIRECTORY
		return $true

	}
	Catch {

		LogError( $_.Exception.Message )
        $z = Read-Host "Press enter to exit: "        
		return $false
	}

}

# END FUNCTIONS
# ----------------------------------------------------------------------


# ----------------------------------------------------------------------
# SCRIPT ENTRY POINT

# create the build directory if new 

MkDir -Force $BUILD_DIRECTORY > $null

# Copy all workbooks to the build directory

$src = [IO.Path]::Combine($CWD, "..\core\cc.isr.core.xlsm")
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit }

$src = [IO.Path]::Combine($CWD, "..\core\cc.isr.core.testing.md") 
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit }

$src = [IO.Path]::Combine($CWD, "..\winsock\cc.isr.winsock.xlsm") 
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit }

$src = [IO.Path]::Combine($CWD, "..\winsock\cc.isr.winsock.testing.md") 
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit }

$src = [IO.Path]::Combine($CWD, "..\ieee488\cc.isr.ieee488.xlsm") 
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit }

$src = [IO.Path]::Combine($CWD, "..\ieee488\cc.isr.ieee488.testing.md") 
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit }


$src = [IO.Path]::Combine($CWD, "cc.isr.vi.xlsm") 
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit }

$src =  [IO.Path]::Combine($CWD, "cc.isr.vi.testing.md") 
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit }

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

    $PRMYFILE = "cc.isr.vi.xlsm"
	RemoveWorkbookReference

# Remove cc_isr_winsock references

    $REFFILE = "cc.isr.winsock.xlsm"
    $REFNAME = "cc_isr_Winsock"

    $PRMYFILE = "cc.isr.ieee488.xlsm"
	RemoveWorkbookReference
	
    $PRMYFILE = "cc.isr.vi.xlsm"
	RemoveWorkbookReference

# Remove cc_isr_ieee488 references

    $REFFILE = "cc.isr.ieee488.xlsm"
    $REFNAME = "cc_isr_Ieee488"

    $PRMYFILE = "cc.isr.vi.xlsm"
	RemoveWorkbookReference
	
}
Catch {
    LogError( $_.Exception.Message )
}
Finally{

	$excel.DisplayAlerts = $true
    $Excel.AutomationSecurity = $AutoSecurity
	$excel.Quit()
}

LogInfo( "project built" )
$z = Read-Host "Press enter to exit"
