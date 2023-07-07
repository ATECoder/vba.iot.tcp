# ----------------------------------------------------------------------
# deploy
#
# PURPOSE: copy a release version of this top level workbook and its 
#          referenced workbooks to the bin folder for deployment.
#
# CALLING SCRIPT:
#
#  ."deploy.ps1"
#
# ----------------------------------------------------------------------

# ----------------------------------------------------------------------
# VARIABLES

$CWD = (Resolve-Path .\).Path
$BUILD_DIRECTORY = [IO.Path]::Combine($CWD, "..\..\bin\vi")
$BUILD_DIRECTORY = (Resolve-Path $BUILD_DIRECTORY).Path
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
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit 1 }

$src = [IO.Path]::Combine($CWD, "..\core\cc.isr.core.testing.md") 
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit 1 }

$src = [IO.Path]::Combine($CWD, "..\winsock\cc.isr.winsock.xlsm") 
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit 1 }

$src = [IO.Path]::Combine($CWD, "..\winsock\cc.isr.winsock.testing.md") 
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit 1 }

$src = [IO.Path]::Combine($CWD, "..\ieee488\cc.isr.ieee488.xlsm") 
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit 1 }

$src = [IO.Path]::Combine($CWD, "..\ieee488\cc.isr.ieee488.testing.md") 
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit 1 }


$src = [IO.Path]::Combine($CWD, "cc.isr.vi.xlsm") 
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit 1 }

$src =  [IO.Path]::Combine($CWD, "cc.isr.vi.testing.md") 
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit 1 }

LogInfo( "project deployed" )
$z = Read-Host "Press enter to exit"
exit 0

