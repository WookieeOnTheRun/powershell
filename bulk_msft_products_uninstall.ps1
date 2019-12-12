write-host "You are attempting to uninstall software by running this script - abandon all hope ye who enter here..."

# enter a wild card search string here - in PS, wildcards are "*" character
# script can easily be updated to also check for exact match of software name if so desired

$SearchArg = read-host "Please enter search string, surround with '*' if necessary: "

# build list of installed Microsoft applications
$apps = @( get-wmiobject -class win32_product -Filter ( "Vendor = 'Microsoft Corporation'" ) | select Name )

foreach( $app in $apps ) {

    if ( $app.Name -like $SearchArg ) {

        $Choice = Read-Host "Do you want to uninstall " + $app.Name + "? Y or N"

        if ( $Choice.ToUpper() -eq "Y" -or $Choice.ToUpper() -eq "YES" ) {

            try {

                $app.Uninstall()

            }

            catch {

                Write-Output "Error uninstalling " + $app.Name + " : see details below."

                Write-Output $.Exception

                Write-Output $.ErrorDetails

            }

        }

        else {

            Write-Output "Skipping uninstall of " + $app.Name

        }

    }    

}
