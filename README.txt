1. Настроить **config.cfg** (запустить единожды *.exe для создания шаблона файла)
   1. [host], IP или DNS имя RedMine (например: **http://192.168.1.1**)
   2. [apikey], *RedMine - Моя учетная запись - Ключ доступа к API*, RESTAPI должен быть глобально включен Администратором (например: **aldjfoeiwgj9348gn348**)
   3. [issuesid], список Issue ID или страниц Wiki разделенных *";"* (например: **1;2;114;9123** или **id100/wiki/Help1;id103/wiki/Help2**)
   4. [saveimg], сохранить картинки в той же директории (example: **true**)
   5. [combine], объединить результат в один файл **IssueCombine.docx** (example: **false**)
   6. [iswiki], issuesid содержит ссылки на Wiki страницы (example: **true**)

2. Запустить 