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

Function AddReference( $workbook )
{
	RemoveReference( $workbook ) 
	$workbook.VBProject.References.AddFromFile( $REFFILE )
}


function AwaitKeyPress()
{
	do{ echo "Press any key";$x = [System.Console]::ReadKey() } while( $x.Key -ne "" )	
}

# END FUNCTIONS
# ----------------------------------------------------------------------


# ----------------------------------------------------------------------
# SCRIPT ENTRY POINT

MkDir -Force $BUILD_DIRECTORY > $null

$SOURCE = [IO.Path]::Combine($CWD, "..\core\cc.isr.core.xlsm")
Try {
	echo coping $SOURCE to $BUILD_DIRECTORY
	copy-item $SOURCE -destination $BUILD_DIRECTORY
}
Catch {
    echo $_.Exception.Message
    return
}

$SOURCE = [IO.Path]::Combine($CWD, "cc.isr.winsock.xlsm")
Try {
	echo coping $SOURCE to $BUILD_DIRECTORY
	copy-item $SOURCE -destination $BUILD_DIRECTORY
}
Catch {
    echo $_.Exception.Message
	return
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$AutoSecurity = $Excel.AutomationSecurity
$Excel.AutomationSecurity = 3
Try {
	
	$excel.DisplayAlerts = $false

    $FILE1 = "cc.isr.winsock.xlsm"
	$PATH1 = [IO.Path]::Combine($BUILD_DIRECTORY, $FILE1)
	$BOOK1 = $excel.Workbooks.Open($PATH1)
    echo "Opened " + $excel.Workbooks.Item($FILE1).Name

    $FILE2 = "cc.isr.core.xlsm"
	$PATH2 = [IO.Path]::Combine($BUILD_DIRECTORY, $FILE2)
	# $BOOK2 = $excel.Workbooks.Open($PATH2)
    # echo $excel.Workbooks.Item($FILE2).Name

    $REFFILE = $FILE2 
    $REFNAME = "cc_isr_Core"
	RemoveReference( $BOOK1 )
    # $excel.Workbooks.Item($FILE2).Close()
    # if looks like the reference was updated automatically at this point?!
	$BOOK1.Save()

	#$BOOK2 = $excel.Workbooks.Open($PATH2)
    #$excel.Workbooks.Add($PATH2)
    
    # $REFFILE = $FILE2 
    # $REFNAME = "cc_isr_Core"
	# AddReference( $BOOK1 )

	# $BOOK2.SaveAs($PATH2, $XL_FILE_FORMAT_MACRO_ENABLED)
	$BOOK1.SaveAs($PATH1, $XL_FILE_FORMAT_MACRO_ENABLED)
    $BOOK1.Close()

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
