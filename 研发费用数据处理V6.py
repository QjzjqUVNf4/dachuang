# -*- coding: utf-8 -*-
import os
import pandas
import time
import decimal
from dachuang import categoriesV2


# main()用于收纳主程序、计算程序运行时间
def main():
    fileName = 'FN_Fn060.xlsx'

    # 根据各项科目中的关键字对科目进行分类
    def classify(text):
        for (category, matches) in categoriesV2.categories:
            if any(match in text for match in matches):
                return category
        if text not in classify.warned:
            # print('Warning: Not matched text ' + text)
            classify.warned.append(text)
        return None

    classify.warned = []
    # 导入科目excel
    df = pandas.read_excel(os.path.join(os.path.dirname(__file__), fileName),
                           sheet_name='sheet1', usecols='A:F',
                           dtype={'Stkcd': str},
                           # Warning: Values are in cents
                           converters={'FN_Fn06002':
                                           lambda s: decimal.Decimal.from_float(100 * round(float(s), 2)) if s else 0
                                       # print(type(s))
                                       },
                           skiprows=[1, 2])  # add “nrows = n” parameter to read partly
    # 构造分层索引（以三个列表：股票代码、年度、自行分类为包含多层级索引的Dataframe的索引）
    levels = [df['Stkcd'].drop_duplicates().tolist(),
              df['Accper'].drop_duplicates().tolist(),
              list(map(lambda tuple: tuple[0], categoriesV2.categories))]

    print(levels)

    # 构造包含多层级索引的Dataframe的框架（用分层索引）
    index = pandas.MultiIndex.from_product(levels, names=['Stkcd', 'Accper', 'Type'])

    series = pandas.Series(0, index, dtype=int)

    for index, row in df.iterrows():
        classifier = classify(row['FN_Fn06001'])
        classifier = 'Oth' if classifier is None else classifier
        series[(row['Stkcd'], row['Accper'], classifier)] = series[(row['Stkcd'], row['Accper'], classifier)] + row[
            'FN_Fn06002']
    # print(classify.warned)
    print(len(classify.warned))
    print(series)
    # print(series[603879, '2020-12-31', 'HC'])
    df = pandas.DataFrame(series[(0 != series.unstack().drop(columns=['All'])).any(1)].unstack())
    # print(df)
    df.to_excel("F:\python\dachuang\数据V6.xlsx", sheet_name='Sheet1')

    # df1=df[['Type', 'Accper']]
    # print(df1)
    # dataframe_warned = pandas.DataFrame(classify.warned)
    # dataframe_warned.to_excel("F:\python\dachuang\未归类项目.xlsx")


# 计算程序的执行时间
start_time = time.time()
main()
print("--- %s ms ---" % ((time.time() - start_time) * 1000))
