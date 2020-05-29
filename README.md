# Stylechecker - программа для проверки форматирования документов в формате .docx

### Как установить и настроить среду для сборки программы:

##### Шаг 1:
Сначала необходимо [установить IDE Visual Studio Community 2019](visualstudio.microsoft.com/ru/vs/) с сайта Microsoft. 
![alt text](https://github.com/ctrlcandy/stylechecker/blob/dev/images/image1.png?raw=true)
 

##### Шаг 2:
Выберите .NET desktop development и нажмите Install
![alt text](https://github.com/ctrlcandy/stylechecker/blob/dev/images/image2.png?raw=true)
 

##### Шаг 3: 
Сохраните себе файлы проекта любым удобным способом с помощью кнопки “Clone or download” в этом репозитории.
![alt text](https://github.com/ctrlcandy/stylechecker/blob/dev/images/image3.png?raw=true)

##### Шаг 4:
Запустите файл stylechecker.sln и соберите проект.  
![alt text](https://github.com/ctrlcandy/stylechecker/blob/dev/images/image4.png?raw=true)


### *На случай, если не заработает какая-нибудь из библиотек…*

##### Шаг 1: 
Перейдите в stylechecker\packages и удалите папки, выделенные на скриншоте.
![alt text](https://github.com/ctrlcandy/stylechecker/blob/dev/images/image5.png?raw=true) 

##### Шаг 2:
Вернитесь обратно к проекту в Visual Studio и выбрать вкладку ПРОЕКТ -> Управление пакетами NuGet…
![alt text](https://github.com/ctrlcandy/stylechecker/blob/dev/images/image6.png?raw=true)

##### Шаг 3:
Установка первой библиотеки. Поисковый запрос – xceed words, необходимо выбрать самую первую.
![alt text](https://github.com/ctrlcandy/stylechecker/blob/dev/images/image7.png?raw=true) 

##### Шаг 4:
Установка второй библиотеки. Поисковый запрос – xml word, необходимо выбрать самую первую.  
![alt text](https://github.com/ctrlcandy/stylechecker/blob/dev/images/image8.png?raw=true)
