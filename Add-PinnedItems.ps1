# Add a folder to the pinned items list. The command does not return a value; 
# thus, it is not possible to know if the item was already present. This 
# function counts the number of items present before an after to determine if 
# the item  was removed and needs to be added. 
# Non-interactive function.

Param(
  [Parameter(Mandatory=$false)] [string] $folderNames,
  [Parameter(Mandatory=$false)] [switch] $chatty
)

# check for parameter, fail if not present.
if ($folderNames -ne ""){
  $regex = ';,'
  $paths=$folderNames.Split($regex)
  $shell = New-Object -ComObject Shell.Application
  foreach ($path in $paths){
    $pinnedItems = $shell.NameSpace("shell:::{679F85CB-0220-4080-B29B-5540CC05AAB6}").Items()
    $oldCount = $pinnedItems.Count()
    if ($chatty) {
      write-host "Found $oldCount items."
    }
    $shell.Namespace($path).Self.InvokeVerb("pintohome")
    $newCount = $shell.NameSpace("shell:::{679F85CB-0220-4080-B29B-5540CC05AAB6}").Items().Count()
    if ($chatty){
      write-host "There are now $newCount items."
    }
    if ($oldCount -gt $newCount){
      if ($chatty){
        write-host "Pinned item was removed. Putting it back." -foregroundcolor red
      }
      $shell.Namespace($path).Self.InvokeVerb("pintohome")
    }
  }
}