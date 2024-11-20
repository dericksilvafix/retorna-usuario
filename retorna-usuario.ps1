function get-ouhierarchy {
    param (
        [string]$parentou,            
        [ref]$userlist,               
        [string]$parentouname        
    )
    $OUs = get-adorganizationalunit -filter * -searchbase $parentou -searchscope onelevel

    foreach ($OU in $OUs) {
        $ouname = "$parentouname\$($OU.Name)"
        write-host "$ouname" -foregroundcolor cyan

        $SubOUs = get-adorganizationalunit -filter * -searchbase $OU.distinguishedname -searchscope onelevel
        if ($SubOUs) {
            get-ouhierarchy -parentou $OU.distinguishedname -userlist $userlist -parentouname $ouname
        }
        
        $users = get-aduser -filter * -searchbase $OU.distinguishedname -properties givenname, surname, description, office, officephone, emailaddress
        foreach ($user in $users) {
            $userdetails = [PSCustomObject]@{
                Nome        = $user.givenName
                Sobrenome   = $user.surname
                Descrição   = $user.description
                Escritório  = $user.office
                Telefone    = $user.officephone
                Email       = $user.emailaddress
                OU          = $ouname              
                SubOU       = $ou.name              
            }

            $userlist.value += $userdetails
        }
    }
}

$rootou = "DC=fix,DC=local"

$allusers = @()

get-ouhierarchy -parentou $rootou -userlist ([ref]$allusers) -parentouname "ASPEC"

$excelpath = "C:\Users\derick.silva\Desktop\usuarios.xlsx"  # substitua pelo caminho que será salvo o arquivo

$allusers | export-excel -path $excelpath -worksheetname "Usuarios" -autosize -tablename "UsuariosAD"

write-host "exportação feita, o arquivo excel foi salvo em: $excelpath" -foregroundcolor white
