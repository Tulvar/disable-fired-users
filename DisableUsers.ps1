Import-Module ActiveDirectory

#����������� ������� ����
$curdate=Get-Date -Format dd.MM.yyyy

#����������� �������� ������
$curM=Get-Date -Format MM

#����������� �������� ����
$curY=get-date -Format yyyy

#����������� ������� ���� ���.�����
$desPath=Get-Date -Format yyyy.MM

#���������� ���� +7 ����
$Buckdate="_{0:dd.MM}" -f (get-date).AddDays(7)

$path = "\\fileserver\����� �� ���������\"
$jpath= "\\fileserver\����� �� ���������\������������\"

#������ ����� csv � ���������� �� �����
$spisokf=Get-Childitem -File -Path $path*.csv | Select-Object -ExpandProperty Name
#���������� ��� ��������������� ������� ����
$todayf=$spisokf | Select-String $curdate
#���������� ���� ��������������
$imppath=$path+$todayf
#OU ���� ����� ���������� ��������������� �������������
$TargetOU = "OU=���������,DC=domain,DC=local"

#��������� �������� � ������������ �����
$header = "N","NN","FIO","Company","Department","Department1","Position","Status","Date","Date2","No","Ready"

#������� ��������� ������
Function Get-Password
{
   param
   (
        # ����������� ����� � ��������� ������
        [int]$Minimum = 10,
        # ������������ ����� ������
        [int]$Maximum = 25
   )
    
   $Assembly = Add-Type -AssemblyName System.Web

   # ����� ������
   $PasswordLength = Get-Random -Minimum $Minimum -Maximum $Maximum

   [System.Web.Security.Membership]::GeneratePassword($PasswordLength, $Minimum)
}

#�������� ����� ���� ������������ �������� ����
If(!(test-path $path\$desPath))
    {
        #������� ����� ���� �� ���
        New-Item -ItemType Directory -Force -Path $path\$desPath
    }

#��������� ���� �� ����� ���� ��� ��������������� �����
If(!(test-path $jpath\$curY))
    {
        #������� ����� ���� �� ���
        New-Item -ItemType Directory -Force -Path $jpath$curY
    }
 #��������� ���� �� ����� ������
 If(!(test-path $jpath\$curY\$curM))
    {
        #������� ����� ���� �� ���
        New-Item -ItemType Directory -Force -Path $jpath$curY\$curM
    }

#���������� ����
$csv=Import-Csv -Delimiter (";") -Encoding Default -Header $header -Path $imppath

foreach ($user in $csv)
    {    
        #������� � ���������� 3 �������
        $fio="$($user.FIO)"
        #�������� ������� ��� � 11 ������ "���", � � 12 �����      
        if (($user.No -eq "���") -and ($user.Ready -eq $null))
            {
                #���� ������������ ������� � ������ ����� �������                                                          
                if ('(Get-ADGroup -Filter {name -eq "Domain Admins"} -prop members).members  -match $fio')
                    {
                        #���� ������������ ������� � ������ ����� ������� ������� ��� � ����������
                        $xyz=(Get-ADGroup -Filter {name -eq "Domain Admins"} -prop members).members  -match $fio
                          #���� ���������� �� ������
                          if ($xyz -ne $null)
                            {
                                #���������� ������ �� ���� �������
                                Send-MailMessage -To admins@domain.ru -From blockscript@domain.local -Subject "������� ���������� ����� ������" -Body "������������� ������� ���������� ����� ������ ($xyz) !!! ��������� ���� ���������� �������� �������������!" -Attachments $imppath -SmtpServer "mail.domain.local" -Encoding UTF8
                                #��������� ������
                                exit
                                
                            }                            
                    }                                               
            }            
    }

#���� ���� "���" � 11 �������
foreach ($user in $csv)
	{
      #����c�� � ���������� 3 �������
      $fio="$($user.FIO)"
      #������� ���������� ������� � �������� ������� "�� �� *
      $company="$($user.Company)" -replace '[''"��]','*'
      #���� � 11 ������� ���� "���" � 12 ������� ������
		if (($user.No -eq "���") -and ($user.Ready -eq $null))
			{
              
                #���� ������������ �� �����, �������� �� ������ � ���� �� �����(����� ������� ������ ����� ��� ���������� ����������� ������), � ��� �� ��������� �� ��������.
				$userdisable = Get-ADUser -Filter {(Name -eq $fio) -and (Enabled -eq "True") -and (mail -ne "null") -and (company -like $company)}

                #���� ������������ �� ����������
				if ($userdisable -eq $null)
					{
                         #����� � 12 ������� " "
						$user.Ready=" "
					}             
                   
				    else
					{                        
                        #��������� ������� � ����� � 12 ������� "������"
						Disable-ADAccount $userdisable
                        #�������� ���� ������������ �� �������� �����
                        Set-ADUser -Identity $userdisable -Replace @{'msExchHideFromAddressLists' = $true}
                        Set-ADAccountPassword -Identity $userdisable -NewPassword (ConvertTo-SecureString -AsPlainText -String Get-Password -force)
                        Move-ADObject -Identity $userdisable -TargetPath $TargetOU
						$user.Ready="������"
					}
			}
	}

#������� ��������� ����������� ��� ������������� ���������� ����� � �����������
if ($csv.ready -eq "������")
    {    
    #������������ ���� � ������� ���������
    $expcsv = $csv | Export-Csv -Delimiter (";") -Encoding Default -Path $jpath$curY\$curM\$curdate$Buckdate.csv  -NoTypeInformation
    #��������� ���� � �����
    Move-Item -Path $imppath -Destination $path$desPath\$todayf
    }
else
    {
    #��������� �������� ���� � �����   
    Move-Item -Path $imppath -Destination $path$desPath\$todayf
    exit
    }