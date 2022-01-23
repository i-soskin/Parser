Парсер угроз безопасности из банка данных угроз ФСТЭК России.

Программа на языке С# с графическим интерфейсом WPF. Выполняет задачи парсера и вывода информации по нужным нам критериям, указанным в заданиях.


Для начала работы необходимо выбрать файл Excel с помощью первой кнопки. Пока файл не будет выбран и загружен в программу, при нажатии на другие кнопки будет уведомление о том, что необходимо загрузить файл. Причём при выборе файла в окне и на вход будут приниматься только файлы в формате Excel. Реализована поддержка сразу двух форматов: .xlsx и .xls. После загрузки файла и нажатия кнопок все данные будут отображаться в рамке.

Вторая кнопка выводит на экран краткий общий перечень угроз по критериям из задания №3 в виде таблицы отображая идентификатор угрозы и её название. 

Третья кнопка отвечает за вывод всей общей информации как указано в задании №4. Отображается вся таблица со всеми данными для каждой угрозы. Единственное что не до конца реализовано, так это разделение на страницы (пагинация).

Четвёртая кнопка позволяет нам вывести всю информацию конкретно о интересующей нас угрозе по заданию №5. Для вывода информации нужно в поле для ввода вписать идентификатор нужной нам угрозы. Программа не примет на вход некорректное значение, так как реализованная поддержка регулярных выражений. Значение буквами, меньше 0 или больше числа записей в таблицы ввести не получится. Вывод всей информации выводится в отдельное окно каждый параметр с новой строки.

Что касается задания №6, в коде программы я постарался применить знания с дисциплины “Управление данными” и реализовал подключение базы данных SQL. При обработке таблицы данные записываются в отдельный локальный файл базы данных. Обновление локальной базы по заданию №2 также реализовано в программе. При работе программы и выборе нового файла Excel, изменённые данные будут перезаписаны в базу, которая хранится на жёстком диске пользователя.
