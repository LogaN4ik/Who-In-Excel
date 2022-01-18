# Счетчик сколько времени проведено в каждом документе. (как часто он повторяется в документе за день)
Clear-Host
Remove-Variable * -ErrorAction SilentlyContinue
$excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')

#задаем переменные
$date = Get-Date -Format "d.MM.y HH:mm"
$person = $env:UserName
$emoji_person = [char]::ConvertFromUtf32(0x1F4A9)

#формируем массив с книгами и массив со статусами
ForEach ( $book in $excel.workbooks ) {$books_names = ,$book.Name + $books_names}
ForEach ( $book in $excel.workbooks ) {$books_reads = ,$book.ReadOnly + $books_reads}

#Отоброжаем все книги и статусы
#ForEach ($item in $books_names){$item}
#ForEach ($item in $books_reads){$item}

#Убираю пустую строку (беру всё без последнего элемента)
$books_names = $books_names[0..($books_names.Count-2)]
$books_reads = $books_reads[0..($books_reads.Count-2)]

#Замена елементов в массиве
$books_reads = $books_reads -replace "False" , "Редакт"
$books_reads = $books_reads -replace "True" , "Чтение"

# приводим 2 массива в табличный вид
$t = $books_names |%{$i=0}{[PSCustomObject]@{Пользователь= $person; Дата= $date;Статус=$books_reads[$i];Книга=$_};$i++}
#$t | ft
#$t | Out-GridView

# Запись в файл
Add-Content "G:\Plan\ОПБиК\2\Скрипты\wiExcel\Reports\Full\$person.txt" $t | ft
Set-Content "G:\Plan\ОПБиК\2\Скрипты\wiExcel\Reports\Last\$person.txt" $t | ft

#pause
