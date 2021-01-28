Import-Module ActiveDirectory

#определение текущей даты
$curdate=Get-Date -Format dd.MM.yyyy

#определение текущего месяца
$curM=Get-Date -Format MM

#определение текущего года
$curY=get-date -Format yyyy

#определение формате даты год.месяц
$desPath=Get-Date -Format yyyy.MM

#вычисление даты +7 дней
$Buckdate="_{0:dd.MM}" -f (get-date).AddDays(7)

$path = "\\fileserver\Отчет по уволенным\"
$jpath= "\\fileserver\Отчет по уволенным\Обработанные\"

#парсим файлы csv и определяем их имена
$spisokf=Get-Childitem -File -Path $path*.csv | Select-Object -ExpandProperty Name
#Определяем имя соответствующее текущей дате
$todayf=$spisokf | Select-String $curdate
#определяем путь импортирования
$imppath=$path+$todayf
#OU куда будем переносить заблокированных пользователей
$TargetOU = "OU=Уволенные,DC=domain,DC=local"

#Заголовки столбцов в подключаемом файле
$header = "N","NN","FIO","Company","Department","Department1","Position","Status","Date","Date2","No","Ready"

#функция генерации пароля
Function Get-Password
{
   param
   (
        # Минимальная длина и сложность пароля
        [int]$Minimum = 10,
        # Максимальная длина пароля
        [int]$Maximum = 25
   )
    
   $Assembly = Add-Type -AssemblyName System.Web

   # Длина пароля
   $PasswordLength = Get-Random -Minimum $Minimum -Maximum $Maximum

   [System.Web.Security.Membership]::GeneratePassword($PasswordLength, $Minimum)
}

#проверка папки куда переноситься исходный файл
If(!(test-path $path\$desPath))
    {
        #создаем папку если ее нет
        New-Item -ItemType Directory -Force -Path $path\$desPath
    }

#проверяем есть ли папка года для експортируемого файла
If(!(test-path $jpath\$curY))
    {
        #создаем папку если ее нет
        New-Item -ItemType Directory -Force -Path $jpath$curY
    }
 #проверяем есть ли папка месяца
 If(!(test-path $jpath\$curY\$curM))
    {
        #создаем папку если ее нет
        New-Item -ItemType Directory -Force -Path $jpath$curY\$curM
    }

#Подключаем файл
$csv=Import-Csv -Delimiter (";") -Encoding Default -Header $header -Path $imppath

foreach ($user in $csv)
    {    
        #заносим в переменную 3 столбец
        $fio="$($user.FIO)"
        #выбираем столбцы где в 11 столбе "нет", а в 12 пусто      
        if (($user.No -eq "Нет") -and ($user.Ready -eq $null))
            {
                #если пользователь состоит в группе домен админов                                                          
                if ('(Get-ADGroup -Filter {name -eq "Domain Admins"} -prop members).members  -match $fio')
                    {
                        #если пользователь состоит в группе домен админов занести его в переменную
                        $xyz=(Get-ADGroup -Filter {name -eq "Domain Admins"} -prop members).members  -match $fio
                          #если переменная не пустая
                          if ($xyz -ne $null)
                            {
                                #отправляем письмо на ящик админов
                                Send-MailMessage -To admins@domain.ru -From blockscript@domain.local -Subject "Попытка блокировки домен админа" -Body "Зафиксирована попытка блокировки домен админа ($xyz) !!! Проверьте файл блокировки уволеных пользователей!" -Attachments $imppath -SmtpServer "mail.domain.local" -Encoding UTF8
                                #прерываем скрипт
                                exit
                                
                            }                            
                    }                                               
            }            
    }

#Цикл ищет "Нет" в 11 столбце
foreach ($user in $csv)
	{
      #Заноcим в переменную 3 Столбец
      $fio="$($user.FIO)"
      #Заносим переменную компани и заменяем символы "»« на *
      $company="$($user.Company)" -replace '[''"»«]','*'
      #Если в 11 Столбце есть "Нет" и 12 столбец пустой
		if (($user.No -eq "Нет") -and ($user.Ready -eq $null))
			{
              
                #Ищем пользователя по имени, включена ли учетка и есть ли почта(такие условия поиска нужны для исключения дублирующий учеток), а так же совпадает ли компания.
				$userdisable = Get-ADUser -Filter {(Name -eq $fio) -and (Enabled -eq "True") -and (mail -ne "null") -and (company -like $company)}

                #Если пользователь не существует
				if ($userdisable -eq $null)
					{
                         #Пишем в 12 столбце " "
						$user.Ready=" "
					}             
                   
				    else
					{                        
                        #Отключаем аккаунт и пишем в 12 столбце "Готово"
						Disable-ADAccount $userdisable
                        #скрываем ящик пользователя из адресной книги
                        Set-ADUser -Identity $userdisable -Replace @{'msExchHideFromAddressLists' = $true}
                        Set-ADAccountPassword -Identity $userdisable -NewPassword (ConvertTo-SecureString -AsPlainText -String Get-Password -force)
                        Move-ADObject -Identity $userdisable -TargetPath $TargetOU
						$user.Ready="Готово"
					}
			}
	}

#условие проверики содержимого для необходимости сохранения файла с изменениями
if ($csv.ready -eq "Готово")
    {    
    #Експортируем файл с записью изменений
    $expcsv = $csv | Export-Csv -Delimiter (";") -Encoding Default -Path $jpath$curY\$curM\$curdate$Buckdate.csv  -NoTypeInformation
    #переносим файл в папку
    Move-Item -Path $imppath -Destination $path$desPath\$todayf
    }
else
    {
    #Переносим исходный файл в папку   
    Move-Item -Path $imppath -Destination $path$desPath\$todayf
    exit
    }