# LraDataExporter

Данная программа позволяет подключаться к .lra файлам и выгружать основные данные по тестам в формате csv, png, реализуя готовые методы HP Analysys API C#
За основую взят программа-пример "AnalysisApiSample2005", которая идет с HP Loadrunner      
![default](https://cloud.githubusercontent.com/assets/13558216/24837483/f97130f8-1d3d-11e7-8ebb-099b53139c58.png)

Как это выглядит:    
Папки с выгруженными данными   
![default](https://cloud.githubusercontent.com/assets/13558216/24837481/ebec670e-1d3d-11e7-9f8c-943939a2b561.png)

График png, сsv со статистикой, сsv с данными (Количество транзакций в секунду)    
![default](https://cloud.githubusercontent.com/assets/13558216/24837475/d51a17f6-1d3d-11e7-924a-fdf7a6ad41e7.png)    

Подробнее:    

После ознакомления с HP Analisys API была написана программа, которая умеет выгружать данные и строить по ним графики. Выгруженная информация сохраняется в папке Report, которая создается в директории с .lra файлом. Данная содержит только некоторые возможности, которые предоставляет HP Analisys API, а именно:   

1) подключиться к .lra файлу (файл HP Analysis)    
2) выгрузить данные по операциям в формате csv (количество транзакций в секунду, времена отклика и тд)    
3) построить по ним графики в png    
4) выгрузить время начала и конца, а также продолжительность теста    

________________________________
__English_

This program allows you to connect to .lra files and download the data for the tests in csv, png, implementing methods from HP Analysys API C#. The sample program "AnalysisApiSample2005" that comes with HP Loadrunner was used to write LraDataExporter.    

Read more:   
After studying HP Analisys API I've decided to write a program that can upload data and build graphics on them. The uploaded information is saved in the "Report" folder that is created in the directory with .lra file. This program contains only some of the possibilities offered by HP Analisys API, namely:     

1) to connect to .the lra file (HP Analysis)   
2) to upload transactions in csv format (number of transactions per second, response times, etc.)    
3) to build on them graphics to png     
4) unload start and end time and the duration of the test    
