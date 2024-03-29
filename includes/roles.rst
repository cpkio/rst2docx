.. vim: textwidth=72:tabstop=3:shiftwidth=3

.. role:: xml(raw)
   :format: openxml

.. role:: area

.. Области экрана

.. role:: button

.. Экранные кнопки.

.. role:: command

.. Команды, системные сервисы (pip, systemctrld и т.д.)

.. role:: field

.. Экранные поля ввода.

.. role:: file(literal)

.. Файл. Моноширинный, без кавычек

.. role:: flag

.. Флаги (бинарные переключатели)

.. role:: folder(literal)

.. Папка. Моноширинный, без кавычек

.. role:: icon

.. Кнопка или иконка

.. role:: key

.. Клавиша. Каждая клавиша в угловых скобках отдельно через плюс,
.. отбиваемый неразрывными пробелами.

.. role:: menu

.. Элементы меню.

.. role:: page

.. Страница. Использование под вопросом.

.. role:: parameter

.. Имена параметров. Моноширинный, полужирный, без кавычек.

.. role:: path

.. role:: screen

.. Эта роль отмечает любой текст, который читатель может видеть на
.. экране (цитата экрана). К таковым относятся приглашения командной строки; строки,
.. которые нужно найти в файле (пользователь должен их увидеть на
.. экране), и аналогичные. Внешний вид: нормальный шрифт, полужирный,
.. в кавычках в виде знаков дюйма.

.. role:: section

.. Раздел.

.. role:: tab

.. Вкладка браузера, закладка в интерфейсе.

.. role:: url

.. URL-адрес

.. role:: user

.. Пользователь

.. role:: userole

.. Пользовательская роль или группа

.. role:: value

.. Значения параметров. Моноширинный, полужирный, без кавычек.

.. role:: window

.. Окно в той или иной форме.

.. role:: json(code)

.. Такая роль позволяет делать подсвечиваемые инлайн-вставки кода вида
.. :json:`{"parameter": value}`

.. role:: input

.. Вместо этой роли вставляется пустое поле заполняемой формы
   с закладкой, на которую потом можно ссылаться. Закладка, как обычно,
   с хэш-суммой вложенного в роль текста; имя поля ввода = "TFтекст".

.. role:: link

.. role:: linkpage

.. Эта роль заменяется на поле REF, ссылающееся на закладку с идентичным
   текстом со вставкой номера страницы, на которой такая закладка
   расположена. Поскольку целью ссылки служит хэш, то таким способом
   можно ссылаться на любой элемент, для которого программно вставляется
   соответствующая закладка, будь то таблица или картинка. По сути эта
   роль дублирует подмену ссылок вида `Text`_, но со вставкой страницы.
   Возможно стоит отказаться от подмены прямых ссылок, заменив их на эту
   роль.

.. role:: prop

.. Эта роль вставляет текст из поля метаданных документа
   с соответствующим именем. Кавычки для обрамления наименования поля не
   нужны; имя поля не чувствительно к регистру. Стандартные поля
   заголовка и описания документа называются Title (title) и Comments
   (comments) соответственно. Эта роль позволяет ссылаться и на
   метаданные, определяемые в document-meta.json

.. ЦВЕТА

.. role:: yellow
.. role:: fuchsia
.. role:: green
.. role:: red

.. Эта роль вставляет «см.…»

.. role:: view
