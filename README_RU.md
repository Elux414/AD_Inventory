# AD_Inventory
PowerShell-скрипт для сбора базовых сведений о компьютерах в домене и вывода результатов в виде таблицы Excel.

Скрипт был протестирован на Windows 10 (PowerShell 5.1), но должен работать как в более новых, так и в более старых версиях.

Скрипт собирает следующую информацию о компьютерах: характеристики процессора (включая количество ядер и потоков); параметры оперативной памяти; сведения о текущем пользователе и всех пользователях, когда либо входивших в систему на данном компьютере; версию ОС и номер сборки; имя компьютера; IP-адрес и данные обо всех сетевых интерфейсах; информацию о производителе и модели; а также сведения о периферийных устройствах (за исключением мышей и клавиатур).

Все результаты экспортируются в таблицу Excel; предусмотрена возможность добавления данных в уже существующий файл.

При возникновении ошибок выполняется диагностика подключения, а результаты диагностики и подробности ошибки записываются на отдельный лист книги Excel.

# Скриншоты:

## Лист Computers
<div align="center">
  <img src="images/Computers.png" alt="Computers sheet" width="100%" />
</div>

## Лист Network
<div align="center">
  <img src="images/Network.png" alt="Network sheet" width="100%" />
</div>

## Лист Users
<div align="center">
  <img src="images/Users.png" alt="Users sheet" width="100%" />
</div>

## Лист Printers
<div align="center">
  <img src="images/Printers.png" alt="Printers sheet" width="100%" />
</div>

## Лист Devices
<div align="center">
  <img src="images/Devices.png" alt="Devies sheet" width="100%" />
</div>

## Лист Diagnostics
<div align="center">
  <img src="images/Diagnostics.png" alt="Diagnostics sheet" width="100%" />
</div>

## Лист Errors
<div align="center">
  <img src="images/Errors.png" alt="Errors sheet" width="100%" />
</div>
