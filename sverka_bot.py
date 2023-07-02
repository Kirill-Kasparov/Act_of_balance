import pandas as pd    # by Kirill Kasparov, 2023, v1.03
import itertools
import telebot
import time
import os
import datetime
import calendar
import numpy as np

with open('bot_token.TXT', 'r') as file:
    token = file.readline().strip()
bot = telebot.TeleBot(token)
# t.me/komus_sverka_bot
count_logs = 0

# Обработчик команды /start
@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.send_message(message.from_user.id, 'Привет! Отправь мне акт сверки из ХД: new_compare.xls')
    bot.send_message(message.from_user.id, 'Ссылка на отчет: http://kdw.komus.net/kdw/reps/paramreport.jsp?numrep=4177')
    bot.send_message(message.from_user.id, 'В оригинальном файле ничего не меняй!')


@bot.message_handler(content_types=['document'])
def handle_file(message):

    def act_of_balance():
        def search_for_orders(sverka_df):
            if len(sverka_df['№ заказа']) > 0:
                orders_list = sverka_df['№ заказа'].to_list()
                orders_list = list(set(orders_list))
                count_total = 0
                count_true = 0
                for order in orders_list:
                    if len(str(order)) > 8:
                        order_df = sverka_df[sverka_df['№ заказа'] == order]
                        if round(order_df['Дебет'].sum(), 2) == round(order_df['Кредит'].sum(), 2):
                            sverka_df.loc[sverka_df['№ заказа'] == order, 'Комментарий'] = 'Оплачено по заказу'
                            count_true += 1
                    count_total += 1
                # print('Опознаны оплаты по заказам:', count_true, '/ Всего заказов:', count_total)
            return sverka_df

        def search_by_amount(sverka_df):
            count_true = 0
            debet_db = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Дебет'] != 0)]
            if len(debet_db['Дебет']) > 0:
                debet_list = debet_db['Дебет'].to_list()
                for debet in debet_list:
                    result = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Кредит'] == debet)]
                    if len(result['Кредит']) > 0:
                        result2 = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Дебет'] == debet)]
                        sverka_df.at[result.iloc[0].name, 'Комментарий'] = 'Найдено по сумме к платежу: ' + str(
                            sverka_df.at[result.iloc[0].name, '№ документа'])
                        sverka_df.at[result2.iloc[0].name, 'Комментарий'] = 'Найдено по сумме к платежу: ' + str(
                            sverka_df.at[result.iloc[0].name, '№ документа'])
                        count_true += 1
                # print('Опознаны оплаты по сумме:', count_true, '/ Всего заказов:', len(debet_list))
            return sverka_df

        def search_by_combo(sverka_df, x=0):
            debet_db = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Дебет'] != 0.0)]
            credit_db = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Кредит'] != 0.0)]
            start_for_debet_list = x
            end_for_debet_list = 15 + x
            # print('Дебет:', len(debet_db['Дебет']), 'Кредит:', len(credit_db['Кредит']))
            if len(debet_db['Дебет']) > 0 and len(credit_db['Кредит']) > 0:
                # собираем словари
                debet_list = []
                for value in debet_db['Дебет'].to_list():
                    if str(value).isdigit():
                        debet_list.append(float(value))
                if len(debet_list) > 20:
                    debet_list = debet_list[
                                 start_for_debet_list:end_for_debet_list]  # диапазон будет смещаться на значение х
                credit_list = []
                for value in credit_db['Кредит'].to_list():
                    if str(value).isdigit():
                        credit_list.append(float(value))

                # ищем комбинации
                count_true = 0
                count_false = 0
                for target_sum in credit_list:
                    combo = []
                    for debet in range(len(debet_list)):
                        for combination in itertools.combinations(debet_list, debet + 1):
                            if sum(combination) == target_sum:
                                combo = list(combination)

                    # записываем результат
                    if len(combo) > 0:
                        # print('Для оплаты:', target_sum, 'найдена комбинация:', combo)
                        result = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Кредит'] == target_sum)]
                        combo_str = ''
                        for combo_value in combo:
                            combo_str = combo_str + str(combo_value) + '; '
                        for combo_value in combo:
                            result2 = sverka_df.loc[
                                (sverka_df['Комментарий'] == '-') & (sverka_df['Дебет'] == combo_value)]
                            sverka_df.at[result2.iloc[0].name, 'Комментарий'] = 'Для оплаты: ' + str(
                                target_sum) + ' найдена комбинация: ' + combo_str
                        sverka_df.at[result.iloc[0].name, 'Комментарий'] = 'Для оплаты: ' + str(
                            target_sum) + ' найдена комбинация: ' + combo_str

                        # пересматриваем лист с учетом вычеркнутых строк
                        debet_db = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Дебет'] != 0.0)]
                        debet_list = []
                        for value in debet_db['Дебет'].to_list():
                            if str(value).isdigit():
                                debet_list.append(float(value))
                        if len(debet_list) > 23:
                            debet_list = debet_list[start_for_debet_list:end_for_debet_list]
                        count_true += 1
                    count_false += 1
                # print('Опознано комбинаций:', count_true, 'из', len(credit_list), 'платежей. Обработано:', count_false)
            return sverka_df

        def search_for_non_payments(sverka_df):
            debet_db = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Дебет'] != 0.0)]
            if len(debet_db['№ заказа']) > 0:
                orders_list = debet_db['№ заказа'].to_list()
                orders_list = list(set(orders_list))
                count_true = 0
                for order in orders_list:
                    if len(str(order)) > 8:
                        order_df = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['№ заказа'] == order)]
                        if order_df['Дебет'].sum() > order_df['Кредит'].sum() and order_df['Кредит'].sum() != 0.0:
                            sverka_df.loc[sverka_df['№ заказа'] == order, 'Комментарий'] = 'Недоплата: ' + str(
                                round(order_df['Дебет'].sum() - order_df['Кредит'].sum(), 2)) + ' руб.'
                            count_true += 1
                        elif order_df['Дебет'].sum() < order_df['Кредит'].sum() and order_df['Кредит'].sum() != 0.0:
                            sverka_df.loc[sverka_df['№ заказа'] == order, 'Комментарий'] = 'Переплата: ' + str(
                                round(order_df['Кредит'].sum() - order_df['Дебет'].sum(), 2)) + ' руб.'
                            count_true += 1
                # print('Выявлено недоплат:', count_true)
            return sverka_df

        def logs():
            # логи
            debet_db = sverka_df.loc[(sverka_df['Комментарий'] != '-') & (sverka_df['Дебет'] != 0)]
            end_debet = len(debet_db['Дебет'])
            user_logs = 'date: ' + str(datetime.datetime.today()) + ';' + 'user:' + str(message.from_user.id) + ';' + \
                        'name:' + str(message.from_user.first_name) + ' ' + str(message.from_user.last_name) + ';' + \
                        'username:' + str(message.from_user.username) + ';' + \
                        'message:' + str(message.text) + ';' + 'file:' + str(
                message.document.file_name) + ';' + 'order:' + str(start_debet) + ';' + 'order_detect:' + str(
                end_debet) + ';' + 'time:' + str(int(time.time() - start))
            print(user_logs)

            while True:  # проверка, если файл открыт
                try:
                    save_user_logs = pd.Series(user_logs)
                    save_user_logs.to_csv('users_logs.csv', sep=';', encoding='windows-1251', index=False, mode='a',
                                          header=False)
                    break
                except UnicodeEncodeError:
                    break

        # Загрузка базы
        sverka_df = pd.read_excel(downloaded_file, sheet_name='сокращённый по ЮрЛицу', header=11, nrows=200000)

        # Чистка базы
        sverka_df = sverka_df.dropna(subset=['Дебет'])
        sverka_df = sverka_df.dropna(subset=['№ п/п'])
        sverka_df = sverka_df.iloc[:, 0: 8]
        sverka_df['Комментарий'] = '-'
        debet_db = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Дебет'] != 0)]
        start_debet = len(debet_db['Дебет'])

        # запускаем проверку по заказам
        sverka_df = search_for_orders(sverka_df)

        # запускаем проверку по сумме
        sverka_df = search_by_amount(sverka_df)

        # запускаем проверку по комбинациям
        total_order = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Дебет'] != 0.0)]
        total_order = len(total_order['Дебет'])
        if total_order > 500:
            bot.send_message(message.from_user.id, "Выявлено " + str(
                total_order) + " неопознанных платежей, обработка может занять несколько минут.")
        sverka_df = search_by_combo(sverka_df)

        # перезапускаем проверку на комбинации со смещением диапазона поиска
        count = 0
        step = 7
        while count != 200:
            debet_db = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Дебет'] != 0.0)]
            credit_db = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Кредит'] != 0.0)]
            if len(debet_db['Дебет']) > 30 and len(credit_db['Кредит']) > 0:
                sverka_df = search_by_combo(sverka_df, step)
                count += 1
                step += 7
                if step > len(debet_db['Дебет']):
                    break
            else:
                break
        end_order = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Дебет'] != 0.0)]
        end_order = len(end_order['Дебет'])
        # print('Опознаны оплаты по комбинациям заказам:', total_order - end_order, '/ Всего заказов:', total_order, '/ Циклов поиска:', count)

        # отмечаем недоплаты
        sverka_df = search_for_non_payments(sverka_df)

        # Завершаем обработку и сохраняем в файл
        logs()
        end = time.time()
        # print("Время выполнения: ", int(end - start), ' сек.')
        sverka_df.to_excel('Разбор_сальдо.xlsx', index=False)

        bot.send_message(message.from_user.id, 'Время выполнения: ' + str(round(end - start)) + ' сек.')

        with open('Разбор_сальдо.xlsx', 'rb') as f:
            bot.send_document(message.chat.id, f)
    def data_listclient_main_partners():
        df_partners_dir = pd.read_excel(downloaded_file, sheet_name='Клиенты', header=12)
        mask = df_partners_dir['Глав.код'].isna()
        # Завершаем обработку и сохраняем в файл
        end = time.time()
        # print("Время выполнения: ", int(end - start), ' сек.')
        df_partners_dir[mask].to_excel('ListClients_main_partners.xlsx', index=False)
        bot.send_message(message.from_user.id, "Время выполнения: " + str(int(end - start)) + ' сек.')
        with open('ListClients_main_partners.xlsx', 'rb') as f:
            bot.send_document(message.chat.id, f)
    def data_month_otgr():
        def workind_days():
            # Получаем дату из файла
            creation_time = pd.read_excel(downloaded_file, sheet_name='Лист ТРП', header=1, nrows=1)
            creation_time = creation_time.columns[0].replace('Отчётная дата: ', '').split()
            creation_time = creation_time[0].split('.')
            ddf = datetime.date(int(creation_time[2]), int(creation_time[1]), int(creation_time[0]))

            # Список праздничных дней
            holidays = pd.read_excel('holidays.xlsx')
            # Конвертируем дату из Excel формата в datetime формат
            for col in holidays.columns:
                holidays[col] = pd.to_datetime(holidays[col]).dt.date

            # Начальная и конечная дата текущего месяца
            start_date = datetime.date(int(ddf.strftime('%Y')), int(ddf.strftime('%m')), 1)
            now_date = datetime.date(int(ddf.strftime('%Y')), int(ddf.strftime('%m')), int(ddf.strftime('%d')))
            end_date = datetime.date(int(ddf.strftime('%Y')), int(ddf.strftime('%m')),
                                     calendar.monthrange(now_date.year, now_date.month)[1])

            # Вычисляем количество рабочих дней
            now_working_days = 0
            end_working_days = 0

            current_date = start_date
            while current_date <= end_date:
                # Если текущий день не является выходным или праздничным, увеличиваем счетчик рабочих дней
                if current_date in list(holidays['working_days']):  # рабочие дни исключения
                    now_working_days += 1
                    end_working_days += 1
                elif current_date.weekday() < 5 and current_date not in list(holidays[int(ddf.strftime('%Y'))]):
                    if current_date <= now_date:
                        now_working_days += 1
                    end_working_days += 1
                current_date += datetime.timedelta(days=1)

            # Получаем данные прошлого года
            start_date_2 = datetime.date(int(ddf.strftime('%Y')) - 1, int(ddf.strftime('%m')), 1)
            now_date_2 = datetime.date(int(ddf.strftime('%Y')) - 1, int(ddf.strftime('%m')), int(ddf.strftime('%d')))
            end_date_2 = datetime.date(int(ddf.strftime('%Y')) - 1, int(ddf.strftime('%m')),
                                       calendar.monthrange(now_date_2.year, now_date_2.month)[1])
            now_working_days_2 = 0
            end_working_days_2 = 0

            current_date = start_date_2
            while current_date <= end_date_2:
                # Если текущий день не является выходным или праздничным, увеличиваем счетчик рабочих дней
                if current_date in list(holidays['working_days']):  # рабочие дни исключения
                    now_working_days_2 += 1
                    end_working_days_2 += 1
                elif current_date.weekday() < 5 and current_date not in list(holidays[int(ddf.strftime('%Y')) - 1]):
                    if current_date < now_date_2:
                        now_working_days_2 += 1
                    end_working_days_2 += 1
                current_date += datetime.timedelta(days=1)
            date_lst = [start_date, start_date_2, now_date, now_date_2, end_date, end_date_2, now_working_days,
                        now_working_days_2, end_working_days, end_working_days_2]
            # print('Начало месяца:', start_date, 'в прошлом году', start_date_2)
            # print('Отчетная дата:', now_date, 'в прошлом году', now_date_2)
            # print('Конец месяца', end_date, 'в прошлом году', end_date_2)
            # print('Рабочих дней сейчас', now_working_days, 'в прошлом году', now_working_days_2)
            # print('Всего рабочих дней', end_working_days, 'в прошлом году', end_working_days_2)
            return date_lst
        def data_month_otgr_list_trp():
            # Загружаем базу 'Лист ТРП'
            df_trp = pd.read_excel(downloaded_file, sheet_name='Лист ТРП', header=14)
            df_trp.drop(df_trp.tail(1).index, inplace=True)  # удаляем строку Итого
            # получаем список месяцев от даты отчета
            col_for_total_result = [now_date.strftime('%d.%m.%Y')]  # вместо list(df.columns[13:0:-1])
            for i in range(1, 13):
                prev_month = now_date - datetime.timedelta(days=28 * i)
                while prev_month.strftime('%m.%Y') in col_for_total_result or prev_month.strftime(
                        '%m.%Y') == now_date.strftime(
                    '%m.%Y'):  # поправка на дни
                    prev_month = prev_month - datetime.timedelta(days=1)
                col_for_total_result.append(prev_month.strftime('%m.%Y'))
            # добавляем месяца, вместо столбцов Т, -1, -2...
            count = 0
            for i in df_trp.columns[4:29:2]:
                df_trp[col_for_total_result[count]] = df_trp[i]
                count += 1
            # добавляем ВП, вместо столбцов КТН
            count = 0
            for i in df_trp.columns[5:30:2]:
                df_trp['ВП ' + str(col_for_total_result[count])] = df_trp.iloc[:, 32 + count] - (
                            df_trp.iloc[:, 32 + count] / df_trp[i])
                df_trp['ВП ' + str(col_for_total_result[count])] = df_trp[
                    'ВП ' + str(col_for_total_result[count])].fillna(0)
                df_trp['ВП ' + str(col_for_total_result[count])] = df_trp[
                    'ВП ' + str(col_for_total_result[count])].replace(-np.inf, 0)
                df_trp['ВП ' + str(col_for_total_result[count])] = df_trp[
                    'ВП ' + str(col_for_total_result[count])].replace(np.inf, 0)
                count += 1
            # Удаляем лишние столбцы с 4 по 31 включительно
            df_trp = df_trp.iloc[:, :4].join(df_trp.iloc[:, 32:], how='outer')
            # Группируем по столбцу "Название ТС", отсекаем столбцы до него
            df_trp = df_trp.groupby('Название ТС')[df_trp.columns[4:]].sum().reset_index()
            # Добавляем строку суммы всех колонок
            sums = df_trp.select_dtypes(include=['number']).sum()  # Считаем сумму
            totals = pd.DataFrame([['Итого'] + sums.tolist()], columns=df_trp.columns)  # собираем ДФ к строке Итого
            # df_trp = pd.concat([df_trp, totals], ignore_index=True)    # добавляем строку в общий ДФ
            # Сортируем список по сумме текущего месяца
            df_trp = df_trp.sort_values(by=now_date.strftime('%d.%m.%Y'), ascending=False)
            # Добавляем КТН
            count = 0
            for i in df_trp.columns[1:14]:
                df_trp['КТН ' + str(col_for_total_result[count])] = round(
                    df_trp[i] / (df_trp[i] - df_trp.iloc[:, 14 + count]), 3)
                df_trp['КТН ' + str(col_for_total_result[count])] = df_trp[
                    'КТН ' + str(col_for_total_result[count])].fillna(0)
                df_trp['КТН ' + str(col_for_total_result[count])] = df_trp[
                    'КТН ' + str(col_for_total_result[count])].replace(-np.inf, 0)
                df_trp['КТН ' + str(col_for_total_result[count])] = df_trp[
                    'КТН ' + str(col_for_total_result[count])].replace(np.inf, 0)
                count += 1
            # Добавляем доли отгрузок
            df_trp['Доля на ' + df_trp.columns[1]] = round(df_trp[df_trp.columns[1]] / df_trp.iloc[0, 1], 4)
            df_trp['Доля на ' + df_trp.columns[13]] = round(df_trp[df_trp.columns[13]] / df_trp.iloc[0, 13], 4)
            df_trp['Отклонение долей'] = df_trp['Доля на ' + df_trp.columns[1]] - df_trp[
                'Доля на ' + df_trp.columns[13]]
            # Добавляем прогноз прироста
            df_trp['Прогноз выполнения'] = round(df_trp[df_trp.columns[1]] / now_working_days * end_working_days, 2)
            df_trp['Прогноз прироста в руб.'] = df_trp['Прогноз выполнения'] - round(
                (df_trp[df_trp.columns[13]] / end_working_days_2 * end_working_days), 2)
            df_trp['Прогноз прироста в %'] = round((df_trp['Прогноз выполнения'] / (
                        df_trp[df_trp.columns[13]] / end_working_days_2 * end_working_days) - 1) * 100, 2)
            df_trp['Прогноз прироста в %'] = df_trp['Прогноз прироста в %'].replace(-np.inf, 0)
            df_trp['Прогноз прироста в %'] = df_trp['Прогноз прироста в %'].replace(np.inf, 0)
            df_trp['Отклонение отгрузок по ССП(3 мес)'] = round(
                df_trp['Прогноз выполнения'] - (df_trp.iloc[:, 2] + df_trp.iloc[:, 3] + df_trp.iloc[:, 4]) / 3)
            df_trp['Прогноз выполнения ВП'] = round(df_trp[df_trp.columns[14]] / now_working_days * end_working_days, 2)
            df_trp['Прогноз прироста ВП в руб.'] = df_trp['Прогноз выполнения ВП'] - round(
                (df_trp[df_trp.columns[26]] / end_working_days_2 * end_working_days), 2)
            df_trp['Прогноз прироста ВП в %'] = round((df_trp['Прогноз выполнения ВП'] / (
                    df_trp[df_trp.columns[26]] / end_working_days_2 * end_working_days) - 1) * 100, 2)
            df_trp['Прогноз прироста ВП в %'] = df_trp['Прогноз прироста ВП в %'].replace(-np.inf, 0)
            df_trp['Прогноз прироста ВП в %'] = df_trp['Прогноз прироста ВП в %'].replace(np.inf, 0)
            df_trp['Прирост КТН'] = df_trp.iloc[:, 27] - df_trp.iloc[:, 39]
            # Завершаем обработку и сохраняем в файл
            end = time.time()
            # print("Время выполнения: ", int(end - start), ' сек.')
            df_trp.to_excel('month_otgr_list_trp.xlsx', index=False)
            bot.send_message(message.from_user.id, "Время выполнения: " + str(int(end - start)) + ' сек.')
            with open('month_otgr_list_trp.xlsx', 'rb') as f:
                bot.send_document(message.chat.id, f)
        def data_month_otgr_list_partners():
            # Загружаем базу 'Лист Партнер'
            df_trp = pd.read_excel(downloaded_file, sheet_name='Лист Партнер', header=14)
            df_trp.drop(df_trp.tail(1).index, inplace=True)  # удаляем строку Итого
            # получаем список месяцев от даты отчета
            col_for_total_result = [now_date.strftime('%d.%m.%Y')]  # вместо list(df.columns[13:0:-1])
            for i in range(1, 13):
                prev_month = now_date - datetime.timedelta(days=28 * i)
                while prev_month.strftime('%m.%Y') in col_for_total_result or prev_month.strftime(
                        '%m.%Y') == now_date.strftime(
                    '%m.%Y'):  # поправка на дни
                    prev_month = prev_month - datetime.timedelta(days=1)
                col_for_total_result.append(prev_month.strftime('%m.%Y'))
            # добавляем месяца, вместо столбцов Т, -1, -2...
            count = 0
            for i in df_trp.columns[15:53:3]:
                df_trp[col_for_total_result[count]] = df_trp[i]
                count += 1
            # добавляем ВП, вместо столбцов КТН
            count = 0
            for i in df_trp.columns[16:53:3]:
                df_trp['ВП ' + str(col_for_total_result[count])] = df_trp.iloc[:, 57 + count] - (
                            df_trp.iloc[:, 57 + count] / df_trp[i])
                df_trp['ВП ' + str(col_for_total_result[count])] = df_trp[
                    'ВП ' + str(col_for_total_result[count])].fillna(0)
                df_trp['ВП ' + str(col_for_total_result[count])] = df_trp[
                    'ВП ' + str(col_for_total_result[count])].replace(-np.inf, 0)
                df_trp['ВП ' + str(col_for_total_result[count])] = df_trp[
                    'ВП ' + str(col_for_total_result[count])].replace(np.inf, 0)
                count += 1
            # добавляем кол-во ТР
            count = 0
            for i in df_trp.columns[17:54:3]:
                df_trp['ТР ' + str(col_for_total_result[count])] = df_trp[i]
                df_trp['ТР ' + str(col_for_total_result[count])] = df_trp[
                    'ТР ' + str(col_for_total_result[count])].fillna(
                    0)
                df_trp['ТР ' + str(col_for_total_result[count])] = df_trp[
                    'ТР ' + str(col_for_total_result[count])].replace(
                    -np.inf, 0)
                count += 1
            # Удаляем лишние столбцы с 11 по 56 включительно
            df_trp = df_trp.iloc[:, :11].join(df_trp.iloc[:, 57:], how='outer')
            # Сортируем список по сумме текущего месяца
            df_trp = df_trp.sort_values(by=now_date.strftime('%d.%m.%Y'), ascending=False)
            # Убираем промежуточные итоги
            mask = df_trp['ГлавКод партнера'].notna()
            df_trp = df_trp[mask]
            # Добавляем КТН
            count = 0
            for i in df_trp.columns[11:24]:
                df_trp['КТН ' + str(col_for_total_result[count])] = round(
                    df_trp[i] / (df_trp[i] - df_trp.iloc[:, 24 + count]), 3)
                df_trp['КТН ' + str(col_for_total_result[count])] = df_trp[
                    'КТН ' + str(col_for_total_result[count])].fillna(0)
                df_trp['КТН ' + str(col_for_total_result[count])] = df_trp[
                    'КТН ' + str(col_for_total_result[count])].replace(-np.inf, 0)
                count += 1
            # Добавляем прогноз прироста
            df_trp['Прогноз выполнения'] = round(df_trp[df_trp.columns[11]] / now_working_days * end_working_days, 2)
            df_trp['Прогноз прироста в руб.'] = df_trp['Прогноз выполнения'] - round(
                (df_trp[df_trp.columns[23]] / end_working_days_2 * end_working_days), 2)
            df_trp['Прогноз прироста в %'] = round((df_trp['Прогноз выполнения'] / (
                        df_trp[df_trp.columns[23]] / end_working_days_2 * end_working_days) - 1) * 100, 2)
            df_trp['Прогноз прироста в %'] = df_trp['Прогноз прироста в %'].replace(-np.inf, 0)
            df_trp['Прогноз прироста в %'] = df_trp['Прогноз прироста в %'].replace(np.inf, 0)
            df_trp['Отклонение отгрузок по ССП(3 мес)'] = round(
                df_trp['Прогноз выполнения'] - (df_trp.iloc[:, 12] + df_trp.iloc[:, 13] + df_trp.iloc[:, 14]) / 3)
            df_trp['Прогноз выполнения ВП'] = round(df_trp[df_trp.columns[24]] / now_working_days * end_working_days, 2)
            df_trp['Прогноз прироста ВП в руб.'] = df_trp['Прогноз выполнения ВП'] - round(
                (df_trp[df_trp.columns[36]] / end_working_days_2 * end_working_days), 2)
            df_trp['Прогноз прироста ВП в %'] = round((df_trp['Прогноз выполнения ВП'] / (
                    df_trp[df_trp.columns[36]] / end_working_days_2 * end_working_days) - 1) * 100, 2)
            df_trp['Прогноз прироста ВП в %'] = df_trp['Прогноз прироста ВП в %'].replace(-np.inf, 0)
            df_trp['Прогноз прироста ВП в %'] = df_trp['Прогноз прироста ВП в %'].replace(np.inf, 0)
            df_trp['Прирост КТН'] = df_trp.iloc[:, 50] - df_trp.iloc[:, 62]
            # Завершаем обработку и сохраняем в файл
            end = time.time()
            # print("Время выполнения: ", int(end - start), ' сек.')
            df_trp.to_excel('month_otgr_list_partners.xlsx', index=False)
            bot.send_message(message.from_user.id, "Время выполнения: " + str(int(end - start)) + ' сек.')
            with open('month_otgr_list_partners.xlsx', 'rb') as f:
                bot.send_document(message.chat.id, f)
        def data_month_otgr_list_ts():
            # Загружаем базу 'Лист ТС'
            df_trp = pd.read_excel(downloaded_file, sheet_name='Лист ТС', header=14)
            df_trp.drop(df_trp.tail(1).index, inplace=True)  # удаляем строку Итого
            # получаем список месяцев от даты отчета
            col_for_total_result = [now_date.strftime('%d.%m.%Y')]  # вместо list(df.columns[13:0:-1])
            for i in range(1, 13):
                prev_month = now_date - datetime.timedelta(days=28 * i)
                while prev_month.strftime('%m.%Y') in col_for_total_result or prev_month.strftime(
                        '%m.%Y') == now_date.strftime(
                    '%m.%Y'):  # поправка на дни
                    prev_month = prev_month - datetime.timedelta(days=1)
                col_for_total_result.append(prev_month.strftime('%m.%Y'))
            # добавляем месяца, вместо столбцов Т, -1, -2...
            count = 0
            for i in df_trp.columns[13:38:2]:
                df_trp[col_for_total_result[count]] = df_trp[i]
                count += 1
            # добавляем ВП, вместо столбцов КТН
            count = 0
            for i in df_trp.columns[14:39:2]:
                df_trp['ВП ' + str(col_for_total_result[count])] = df_trp.iloc[:, 41 + count] - (
                            df_trp.iloc[:, 41 + count] / df_trp[i])
                df_trp['ВП ' + str(col_for_total_result[count])] = df_trp[
                    'ВП ' + str(col_for_total_result[count])].fillna(0)
                df_trp['ВП ' + str(col_for_total_result[count])] = df_trp[
                    'ВП ' + str(col_for_total_result[count])].replace(-np.inf, 0)
                df_trp['ВП ' + str(col_for_total_result[count])] = df_trp[
                    'ВП ' + str(col_for_total_result[count])].replace(np.inf, 0)
                count += 1
            # Удаляем лишние столбцы с 13 по 40 включительно
            df_trp = df_trp.iloc[:, :13].join(df_trp.iloc[:, 41:], how='outer')
            # Сортируем список по сумме текущего месяца
            df_trp = df_trp.sort_values(by=now_date.strftime('%d.%m.%Y'), ascending=False)
            # Убираем промежуточные итоги
            mask = df_trp['ГлавКод партнера'].notna()
            df_trp = df_trp[mask]
            # Добавляем КТН
            count = 0
            for i in df_trp.columns[13:26]:
                df_trp['КТН ' + str(col_for_total_result[count])] = round(
                    df_trp[i] / (df_trp[i] - df_trp.iloc[:, 26 + count]), 3)
                df_trp['КТН ' + str(col_for_total_result[count])] = df_trp[
                    'КТН ' + str(col_for_total_result[count])].fillna(0)
                df_trp['КТН ' + str(col_for_total_result[count])] = df_trp[
                    'КТН ' + str(col_for_total_result[count])].replace(-np.inf, 0)
                count += 1
            # Добавляем прогноз прироста
            df_trp['Прогноз выполнения'] = round(df_trp[df_trp.columns[13]] / now_working_days * end_working_days, 2)
            df_trp['Прогноз прироста в руб.'] = df_trp['Прогноз выполнения'] - round(
                (df_trp[df_trp.columns[25]] / end_working_days_2 * end_working_days), 2)
            df_trp['Прогноз прироста в %'] = round((df_trp['Прогноз выполнения'] / (
                        df_trp[df_trp.columns[25]] / end_working_days_2 * end_working_days) - 1) * 100, 2)
            df_trp['Прогноз прироста в %'] = df_trp['Прогноз прироста в %'].replace(-np.inf, 0)
            df_trp['Прогноз прироста в %'] = df_trp['Прогноз прироста в %'].replace(np.inf, 0)
            df_trp['Отклонение отгрузок по ССП(3 мес)'] = round(
                df_trp['Прогноз выполнения'] - (df_trp.iloc[:, 14] + df_trp.iloc[:, 15] + df_trp.iloc[:, 16]) / 3)
            df_trp['Прогноз выполнения ВП'] = round(df_trp[df_trp.columns[26]] / now_working_days * end_working_days, 2)
            df_trp['Прогноз прироста ВП в руб.'] = df_trp['Прогноз выполнения ВП'] - round(
                (df_trp[df_trp.columns[38]] / end_working_days_2 * end_working_days), 2)
            df_trp['Прогноз прироста ВП в %'] = round((df_trp['Прогноз выполнения ВП'] / (
                    df_trp[df_trp.columns[38]] / end_working_days_2 * end_working_days) - 1) * 100, 2)
            df_trp['Прогноз прироста ВП в %'] = df_trp['Прогноз прироста ВП в %'].replace(-np.inf, 0)
            df_trp['Прогноз прироста ВП в %'] = df_trp['Прогноз прироста ВП в %'].replace(np.inf, 0)
            df_trp['Прирост КТН'] = df_trp.iloc[:, 39] - df_trp.iloc[:, 51]
            df_trp['Отгрузка факт'] = round(df_trp.iloc[:, 13], 2)
            df_trp['Отгрузка -12 мес (пересчет)'] = round(df_trp.iloc[:, 25] / end_working_days_2 * end_working_days, 2)
            df_trp['ВП факт'] = round(df_trp.iloc[:, 26], 2)
            df_trp['ВП -12 мес (пересчет)'] = round(df_trp.iloc[:, 38] / end_working_days_2 * end_working_days, 2)
            # Завершаем обработку и сохраняем в файл
            end = time.time()
            # print("Время выполнения: ", int(end - start), ' сек.')
            df_trp.to_excel('month_otgr_list_ts.xlsx', index=False)
            bot.send_message(message.from_user.id, "Время выполнения: " + str(int(end - start)) + ' сек.')
            with open('month_otgr_list_ts.xlsx', 'rb') as f:
                bot.send_document(message.chat.id, f)

        start_date, start_date_2, now_date, now_date_2, end_date, end_date_2, now_working_days, now_working_days_2, end_working_days, end_working_days_2 = workind_days()
        data_month_otgr_list_trp()
        data_month_otgr_list_partners()
        data_month_otgr_list_ts()
    def data_network_partners():
        def workind_days():
            # Получаем дату из файла
            creation_time = pd.read_excel(downloaded_file, sheet_name='Лист Клиент', header=1, nrows=1)
            creation_time = creation_time.columns[0].replace('Отчётная дата: ', '').split()
            creation_time = creation_time[0].split('.')
            ddf = datetime.date(int(creation_time[2]), int(creation_time[1]), int(creation_time[0]))
            # Список праздничных дней
            holidays = pd.read_excel('holidays.xlsx')
            # Конвертируем дату из Excel формата в datetime формат
            for col in holidays.columns:
                holidays[col] = pd.to_datetime(holidays[col]).dt.date

            # Начальная и конечная дата текущего месяца
            start_date = datetime.date(int(ddf.strftime('%Y')), int(ddf.strftime('%m')), 1)
            now_date = datetime.date(int(ddf.strftime('%Y')), int(ddf.strftime('%m')), int(ddf.strftime('%d')))
            end_date = datetime.date(int(ddf.strftime('%Y')), int(ddf.strftime('%m')),
                                     calendar.monthrange(now_date.year, now_date.month)[1])

            # Вычисляем количество рабочих дней
            now_working_days = 0
            end_working_days = 0

            current_date = start_date
            while current_date <= end_date:
                # Если текущий день не является выходным или праздничным, увеличиваем счетчик рабочих дней
                if current_date in list(holidays['working_days']):  # рабочие дни исключения
                    now_working_days += 1
                    end_working_days += 1
                elif current_date.weekday() < 5 and current_date not in list(holidays[int(ddf.strftime('%Y'))]):
                    if current_date <= now_date:
                        now_working_days += 1
                    end_working_days += 1
                current_date += datetime.timedelta(days=1)

            # Получаем данные прошлого года
            start_date_2 = datetime.date(int(ddf.strftime('%Y')) - 1, int(ddf.strftime('%m')), 1)
            now_date_2 = datetime.date(int(ddf.strftime('%Y')) - 1, int(ddf.strftime('%m')), int(ddf.strftime('%d')))
            end_date_2 = datetime.date(int(ddf.strftime('%Y')) - 1, int(ddf.strftime('%m')),
                                       calendar.monthrange(now_date_2.year, now_date_2.month)[1])
            now_working_days_2 = 0
            end_working_days_2 = 0

            current_date = start_date_2
            while current_date <= end_date_2:
                # Если текущий день не является выходным или праздничным, увеличиваем счетчик рабочих дней
                if current_date in list(holidays['working_days']):  # рабочие дни исключения
                    now_working_days_2 += 1
                    end_working_days_2 += 1
                elif current_date.weekday() < 5 and current_date not in list(holidays[int(ddf.strftime('%Y')) - 1]):
                    if current_date < now_date_2:
                        now_working_days_2 += 1
                    end_working_days_2 += 1
                current_date += datetime.timedelta(days=1)
            date_lst = [start_date, start_date_2, now_date, now_date_2, end_date, end_date_2, now_working_days,
                        now_working_days_2, end_working_days, end_working_days_2]
            # print('Начало месяца:', start_date, 'в прошлом году', start_date_2)
            # print('Отчетная дата:', now_date, 'в прошлом году', now_date_2)
            # print('Конец месяца', end_date, 'в прошлом году', end_date_2)
            # print('Рабочих дней сейчас', now_working_days, 'в прошлом году', now_working_days_2)
            # print('Всего рабочих дней', end_working_days, 'в прошлом году', end_working_days_2)
            return date_lst
        start_date, start_date_2, now_date, now_date_2, end_date, end_date_2, now_working_days, now_working_days_2, end_working_days, end_working_days_2 = workind_days()
        df_trp = pd.read_excel(downloaded_file, sheet_name='Лист Клиент', header=10, nrows=200000)
        # получаем список месяцев от даты отчета
        col_for_total_result = [now_date.strftime('%d.%m.%Y')]
        for i in range(1, 13):
            prev_month = now_date - datetime.timedelta(days=28 * i)
            while prev_month.strftime('%m.%Y') in col_for_total_result or prev_month.strftime(
                    '%m.%Y') == now_date.strftime('%m.%Y'):  # поправка на дни
                prev_month = prev_month - datetime.timedelta(days=1)
            col_for_total_result.append(prev_month.strftime('%m.%Y'))
        # добавляем месяца, вместо столбцов Т, -1, -2...
        count = 0
        for i in df_trp.columns[8:34:2]:
            df_trp[col_for_total_result[count]] = df_trp[i]
            count += 1
        # добавляем ВП, вместо столбцов КТН
        count = 0
        for i in df_trp.columns[9:34:2]:
            df_trp['ВП ' + str(col_for_total_result[count])] = df_trp.iloc[:, 37 + count] - (
                    df_trp.iloc[:, 37 + count] / df_trp[i])
            df_trp['ВП ' + str(col_for_total_result[count])] = df_trp['ВП ' + str(col_for_total_result[count])].fillna(
                0)
            df_trp['ВП ' + str(col_for_total_result[count])] = df_trp['ВП ' + str(col_for_total_result[count])].replace(
                -np.inf, 0)
            df_trp['ВП ' + str(col_for_total_result[count])] = df_trp['ВП ' + str(col_for_total_result[count])].replace(
                np.inf, 0)
            count += 1
        # Удаляем лишние столбцы с 8 по 36 включительно
        df_trp = df_trp.iloc[:, :8].join(df_trp.iloc[:, 37:], how='outer')
        # Сортируем список по сумме текущего месяца
        df_trp = df_trp.sort_values(by=now_date.strftime('%d.%m.%Y'), ascending=False)
        # Убираем промежуточные итоги
        mask = df_trp['Код партнера'].notna()
        df_trp = df_trp[mask]
        # Добавляем прогноз прироста
        df_trp['Прогноз выполнения'] = round(df_trp[df_trp.columns[8]] / now_working_days * end_working_days, 2)
        df_trp['Прогноз прироста в руб.'] = df_trp['Прогноз выполнения'] - round(
            (df_trp[df_trp.columns[20]] / end_working_days_2 * end_working_days), 2)
        df_trp['Отклонение отгрузок по ССП(3 мес)'] = round(
            df_trp['Прогноз выполнения'] - (df_trp.iloc[:, 9] + df_trp.iloc[:, 10] + df_trp.iloc[:, 11]) / 3)
        df_trp['Прогноз выполнения ВП'] = round(df_trp[df_trp.columns[21]] / now_working_days * end_working_days, 2)
        df_trp['Прогноз прироста ВП в руб.'] = df_trp['Прогноз выполнения ВП'] - round(
            (df_trp[df_trp.columns[33]] / end_working_days_2 * end_working_days), 2)

        df_trp['Отгрузка факт'] = round(df_trp.iloc[:, 8], 2)
        df_trp['Отгрузка -12 мес (пересчет)'] = round(df_trp.iloc[:, 20] / end_working_days_2 * end_working_days, 2)
        df_trp['ВП факт'] = round(df_trp.iloc[:, 21], 2)
        df_trp['ВП -12 мес (пересчет)'] = round(df_trp.iloc[:, 33] / end_working_days_2 * end_working_days, 2)
        # Завершаем обработку и сохраняем в файл
        end = time.time()
        # print("Время выполнения: ", int(end - start), ' сек.')
        df_trp.to_excel('month_net_for_bi.xlsx', index=False)
        bot.send_message(message.from_user.id, "Время выполнения: " + str(int(end - start)) + ' сек.')
        with open('month_net_for_bi.xlsx', 'rb') as f:
            bot.send_document(message.chat.id, f)

    start = time.time()  # для таймера выполнения
    # Скачиваем файл
    file_info = bot.get_file(message.document.file_id)
    if '.xls' in str(message.document.file_name) and 'new_compare' in str(message.document.file_name):
        downloaded_file = bot.download_file(file_info.file_path)
        try:
            act_of_balance()
        except:
            bot.send_message(message.from_user.id, 'Что-то пошло не так: 1. попробуйте сформировать сверку за другой период 2. убедитесь в наличии листа "сокращённый по ЮрЛицу"')

    elif '.xls' in str(message.document.file_name) and 'ListClients' in str(message.document.file_name):
        downloaded_file = bot.download_file(file_info.file_path)
        data_listclient_main_partners()
    elif '.xls' in str(message.document.file_name) and 'month_otgr' in str(message.document.file_name):
        downloaded_file = bot.download_file(file_info.file_path)
        data_month_otgr()
    elif '.xls' in str(message.document.file_name) and 'month_net' in str(message.document.file_name):
        downloaded_file = bot.download_file(file_info.file_path)
        data_network_partners()
    else:
        bot.send_message(message.from_user.id, 'Неверный файл. Необходимо отправить исходник new_compare.xls')


try:  # перезапуск при достижении лимита и дисконекте
    count_logs += 1
    if count_logs == 200:  # перезапуск при достижении лимита
        bot.stop_polling()
        crash_logs = pd.Series(
            str(pd.to_datetime(int(time.time() + 10800), unit='s')) + ';' + 'Достигнут лимит 200 запросов')
        crash_logs.to_csv('crash_logs.csv', sep=';', encoding='windows-1251', index=False, mode='a', header=False)
        os.startfile("Sverka (server for bot).exe")

    bot.polling(none_stop=True, interval=0)
except:  # перезапуск при дисконекте
    crash_logs = pd.Series(str(pd.to_datetime(int(time.time() + 10800), unit='s')) + ';' + 'Дисконнект')
    crash_logs.to_csv('crash_logs.csv', sep=';', encoding='windows-1251', index=False, mode='a', header=False)
    time.sleep(10)
    os.startfile("Sverka (server for bot).exe")
