# PdfSamplesGenerator
<div style="text-align:center;">

**Техническое задание**

</div>

<div style="color:#000000;">

Программка например в PowerShell, которую запускаешь и она выполняет
следующие действия:

</div>

<div style="color:#000000;">

**Описание:**

</div>

Просматривает содержимое каталога и по заданному алгоритму открывает
имеющиеся файлы в соответствующих программах (которые установлены на
компьютере) и создает файлы \*.pdf объемом до 8 стр (при их наличии) с
защитными элементами на фоне (например слово “образец” которое нельзя
убрать при печати, можно предложить нам еще какие-нибудь элементы
защиты файла), созданные файлы копируются в соответствующую папку.

<div style="color:#000000;">

Возможный алгоритм

</div>

#

<div style="margin-left:1.27cm;margin-right:0cm;">

Перебираем файлы в текущем каталоге

</div>

1.   
        
        <div style="margin-left:2.461cm;margin-right:0cm;">
        
        <u>'''Алгоритм “А” '''</u>Если в каталоге файлы не **архивы**:
        
        </div>
        
  <div style="margin-left:2.461cm;margin-right:0cm;">

'''Файлы '''с расширениями **(\*.rtf, \*.cdr, \*.jpg, \*.tif, \*.doc,
\*.docx, \*.indd, \*.pdf)
<span style="background-color:#008080;">?</span><span style="background-color:#008080;color:#000000;">jpeg</span><span style="background-color:#008080;">
</span><span style="background-color:#008080;">tiff?**</span><span style="background-color:#008080;">
</span>открываем, сохраняем как **(\*.pdf)** с 1-8 страницу, с защитными
знаками, на каждой странице (если страниц меньше то с 1 стр. по
последнюю).Название файла тоже что и у Исходного только в конце
названия добавить “\_образец”. (Например: Исходный файл
“**(011-3-1-23665)домострой.docx”** - Конечный файл
“**(011-3-1-23665)домострой\_образец.pdf”**Файл сохранить **в папку
“Образцы”** которая находится в текущем каталоге (если такой папки
нет, остановить программу, **сообщить “Нет папки “Образцы””**
предложить **“Продолжить”** или **“Отменить”.** При нажатии
**“Продолжить”** проверяется есть ли папка “Образцы”, если есть -
обработка продолжается иначе в паузе.При нажатии **“Отменить”**
отменяется выполнение программы и она зарывается.

</div>

#

2  
        
        <div style="margin-left:2.54cm;margin-right:0cm;">
        
        <u>'''Алгоритм “В” *'</u>Если файл **Архив**, тогда заходим в
        архив:Если внутри есть папки тогда обязательно заходим в
        первую же до самого последнего вложения, и в последней
        папке обрабатываем файлы по следующему алгоритму:**:** Ищем
        файлы по приоритету, сначала (**1\_.pdf** если нет - **\*.pdf**
        если нет -**\*.jpg** если нет - **\*.tif** если нет -
        **\*.cdr** если нет - **файл с раширением (doc или docx)** **и
        названием начинающимся с “telo..”** если нет - **любой файл**
        с расширением **(doc или docx)** если нет, **выйти из архива,
        сделать запись в лог-файл, что такой-то архив не содержит
        искомых файлов и идти дальше.**Если находи искомый файл
        применяем*' <u>Алгоритм “A”. '''(в папку образец копируется
        только полученный файл \*.pdf).