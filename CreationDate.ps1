$modifyfiles = Get-ChildItem -force | Where-Object {! $_.PSIsContainer}
foreach($object in $modifyfiles)
{
$object.CreationTime=("27/03/2023 18:44:12")

}