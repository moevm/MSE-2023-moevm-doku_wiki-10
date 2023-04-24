## Ручная установка плагина
1. Склонировать репозиторий по SSH или HTTPS:
    <br>SSH:    ```git clone git@github.com:moevm/MSE-2023-moevm-doku_wiki-10.git```
    <br>HTTPS:  ```git clone https://github.com/moevm/MSE-2023-moevm-doku_wiki-10.git```
2. Переместить все файлы плагина по пути:
    <br> ```dokuwiki-installation-directory/lib/plugins/xlsx2dw/```, где 
    <br> ```dokuwiki-installation-directory``` директория развернутой dokuwiki на вашем компьютере.
    <br> Сделать это можно через графический интерфейс. После перемещения всех файлов, каталог ```/xlsx2dw``` должен выглядеть следующим образом:
   ![Alt text](./screenshots/plugin_directory.png?raw=true "Содержимое каталога /xlsx2dw")
## Проверка работоспособности плагина
Для того чтобы убедиться в работоспособности плагина необходимо:
1. На главной странице dokuwiki выбрать раздел создания новой страницы или редактирования уже существующей. Далее будет показано использование плагина при создании новой страницы.
   ![Alt text](./screenshots/creating_page_section.png?raw=true "Создание страницы")
2. В панели инструментов для редактирования открытой страницы нажать на иконку импорта таблиц.
   ![Alt text](./screenshots/using_button.jpg?raw=true "Импорт таблиц")
3. Выбрать таблицу для импорта. Например, воспользовавшись тестовыми таблицами в ```/_test/test-tables/```.
   ![Alt text](./screenshots/selecting_tables.png?raw=true "Выбор таблицы")
4. Убедиться, что выбранная таблица преобразовалась в синтаксис dokuwiki.
   ![Alt text](./screenshots/table_in_dokusyntax.png?raw=true "Докувики синтаксис")
5. Итог работы плагина представлен на скриншоте ниже:
   ![Alt text](./screenshots/preview_table.png?raw=true "Итог")

   
