set-strictMode -version latest

function init {

   if ($psVersionTable.PsEdition -eq 'Core') {

      $src = get-content -raw "$psScriptRoot\COM.cs"
      add-type -typeDef $src

   }
}

function get-activeObject (
      [string] $progId
   ) {


  if ($psVersionTable.PsEdition -eq 'Core') {
     return [TQ84.COM]::getActiveObject($progId)
  }

  try {
     return [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
  }
  catch {

     if ($_.exception.hResult -eq -2146233087) {
        write-warning "get-activeObject: progId $progId not found"
        return $null
     }

     $_.exception | select *
     write-host ('HRESULT {0:x}' -f ([int64] $_.exception.hResult ))
  }
}

function get-COMPropertyValue {
 #
 # Returns the value of a collection that is indexed
 # by name or index.
 # Returns null if property does not exist.
 #
   param (
      [__ComObject] $obj,
      [string]      $property
   )

   try {
      return $obj.item($property)
   }
   catch [System.Runtime.InteropServices.COMException]  {

      if ($_.Exception.HResult -eq -2146825023) {
         return $null
      }

      write-host 'get-COMPropertyValue: System.Runtime.InteropServices.COMException'
      $_.exception | select *
   }
}

init
