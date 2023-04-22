import pandas as pd
import itertools
import time


def search_for_orders(sverka_df):
    if len(sverka_df['№ заказа']) > 0:
        orders_list = sverka_df['№ заказа'].to_list()
        orders_list = list(set(orders_list))
        sverka_df['Комментарий'] = '-'
        count_total = 0
        count_true = 0
        for order in orders_list:
            if len(str(order)) > 8:
                order_df = sverka_df[sverka_df['№ заказа'] == order]
                if order_df['Дебет'].sum() == order_df['Кредит'].sum():
                    sverka_df.loc[sverka_df['№ заказа'] == order, 'Комментарий'] = 'Оплачено по заказу'
                    count_true += 1
            count_total += 1
        print('Опознаны оплаты по заказам:', count_true, '/ Всего заказов:', count_total)
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
        print('Опознаны оплаты по сумме:', count_true, '/ Всего заказов:', len(debet_list))
    return sverka_df
def search_by_combo(sverka_df, x = 0):
    debet_db = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Дебет'] != 0.0)]
    credit_db = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Кредит'] != 0.0)]
    start_for_debet_list = x
    end_for_debet_list = 20 + x
    # print('Дебет:', len(debet_db['Дебет']), 'Кредит:', len(credit_db['Кредит']))
    if len(debet_db['Дебет']) > 0 and len(credit_db['Кредит']) > 0:
        # собираем словари
        debet_list = []
        for value in debet_db['Дебет'].to_list():
            if str(value).isdigit():
                debet_list.append(float(value))
        if len(debet_list) > 23:
            debet_list = debet_list[start_for_debet_list:end_for_debet_list] # диапазон будет смещаться на значение х
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
                    result2 = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Дебет'] == combo_value)]
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
                if order_df['Дебет'].sum() != order_df['Кредит'].sum():
                    sverka_df.loc[sverka_df['№ заказа'] == order, 'Комментарий'] = 'Недоплата: ' + str(order_df['Дебет'].sum() - order_df['Кредит'].sum()) + 'руб.'
                    count_true += 1
        print('Выявлено недоплат:', count_true)
    return sverka_df

start = time.time()    # для таймера выполнения

# Загрузка базы
sverka_df = pd.read_excel('new_compare.xls', sheet_name='сокращённый по ЮрЛицу', header=11, nrows=200000)

# Чистка базы
sverka_df = sverka_df.dropna(subset=['Дебет'])
sverka_df = sverka_df.dropna(subset=['№ п/п'])
sverka_df = sverka_df.iloc[:, 0: 8]

# запускаем проверку по заказам
sverka_df = search_for_orders(sverka_df)

# запускаем проверку по сумме
sverka_df = search_by_amount(sverka_df)

# запускаем проверку по комбинациям
total_order = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Дебет'] != 0.0)]
total_order = len(total_order['Дебет'])
sverka_df = search_by_combo(sverka_df)

# перезапускаем проверку на комбинации со смещением диапазона поиска
count = 0
step = 10
while count != 200:
    debet_db = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Дебет'] != 0.0)]
    credit_db = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Кредит'] != 0.0)]
    if len(debet_db['Дебет']) > 30 and len(credit_db['Кредит']) > 0:
        sverka_df = search_by_combo(sverka_df, step)
        count +=1
        step += 10
        if step > len(debet_db['Дебет']):
            break
    else:
        break
end_order = sverka_df.loc[(sverka_df['Комментарий'] == '-') & (sverka_df['Дебет'] != 0.0)]
end_order = len(end_order['Дебет'])
print('Опознаны оплаты по комбинациям заказам:', total_order - end_order, '/ Всего заказов:', total_order, '/ Циклов поиска:', count)

# отмечаем недоплаты
# sverka_df = search_for_non_payments(sverka_df)

# Завершаем обработку и сохраняем в файл
end = time.time()
print("Время выполнения:", int(end-start), 'сек.')
sverka_df.to_excel('Разбор_сальдо.xlsx', index=False)