# СПЖС5

## Обзор

СПЖС5 - робот созданный в ноябре 2023 года командой малой автоматизации
по заданию Валуйской Елены и Щеглова Ярослава.
Данный робот осуществляет проверку заявок со статусами "Ответы получены" и 
"Выписки получены" на наличие ФРОТ в данных заявках, а также проверку
кадастрового номера на дубли и назначение заявки на ответственного
сотрудника.

## Запуск робота
Данный робот написан без использования вертуального окружения (venv).
Версии библиотек и правило установки будет указано ниже.

### requests - библиотека предназначенная для общения запросами c сайтом МГПС
'''
pip install requests==2.28.2
'''

### pandas - библиотека для работы с данными

'''
pip install pandas==1.4.3
'''

### openpyxl - библиотека для работы с файлом эксель. В данном коде используется только для применения форрматирования
'''
pip install openpyxl==3.0.10
'''

### logger - библиотека для записи логов в файл. Позволяет отследить все, что происходит с пррограммой
'''
pip install logger==1.4
'''

### ConfigParser - библиотека для чтения данных из файла.
'''
pip install ConfigParser
'''

### exchangelib - библиотека для отправки сообщения в почте
'''
pip install exchangelib==4.9.0
'''

##  Разработка
### 1. config = ConfigParser()
### Создаем экземпляр с помощью которого в дальнейшем будет считывать всю информацию, которая касается логинов, паролей , ссылок и так далее.
### Вся информация данного типо хранится в файле конфига

### 2. class FunctionsStatic
### Класс функций которые переодически используются, но не описывают свойства какого либо объекта.

### 3. class MGPs(FunctionsStatic):
### Класс для работы с процессами связанными с МГПС. Наследуеюся от FunctionsStatic для того, чтобы в внутри можно было применить методы из FunctionsStatic.







