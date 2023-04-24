# Act_of_balance
Разбор акта сверки

Предпосылки:
---------
В работе сбыта регулярно возникает потребность найти неоплаченные заказы, либо провести квартальную сверку с партнером. Даже нетрудозатратные запросы направлялись в бухгалтерию и становились в очередь на обработку в среднем на 3 дня. Появилась необходимость упростить процесс.

Описание:
---------
Цель программы - быстро выявить неоплаты и расхождения по акту сверки.
Обработка происходит в 4 этапа:
1. Сравнение сальдо по номеру документа. Таблица последовательно фильтруется по номеру документа и сравнивает сумму всех строк Дебета и Кредита. Если совпадение найдено - фиксирует в столбце Комментарии.

2. Сравнение совпадающих сумм. Программа исключает раннее найденные совпадения и ищет строгое совпадение сумм Дебета и Кредита. При нахождении одинаковых сумм, вносит в комментарии номер платежного документа.

3. Поиск неопознанных платежей, включающих оплату одной суммой по нескольким документам. Программа берет до 25 неопознанных сумм по полю Дебет, и на каждое оставшееся значение в поле Кредит подбирает все возможные комбинации сумм, пока не найдет соответствие. Если массив сумм по полю Дебет больше 25 значений, программа берет следующую группу сумм и проводит очередной поиск комбинаций, пока не завершит весь список. В комментарии фиксируются суммы, по которым была найдена комбинация оплаты.

4. Повторный поиск по номеру документа оставшихся неопознанных счетов. На этот раз ищем расхождения, чтобы подсветить недоплаты и переплаты.

Готовый разбор возвращается ответом бота в таблице "Разбор_сальдо.xlsx"

Быстродействие:
Время выполнения от 2 сек для простой сверки, до 300 сек для сложной сверки более чем в 10 тыс. строк. Скорость обработки регулируется настройками по поиску в пункте 3. Увеличение быстродействия достигается за счет снижения количества возможных комбинаций неопознанных сумм.
