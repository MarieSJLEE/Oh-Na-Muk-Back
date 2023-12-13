import pandas as pd
import os

def calculate_representative_nutrient(row):
    nutrition_columns = ['단백질(g)', '지방(g)', '탄수화물(g)', '당류(g)', '나트륨(mg)']
    nutrition_values = []

    for col in nutrition_columns:
        if pd.notna(row[col]) and row[col] not in ('', '-', '0'):
            value_str = str(row[col]).replace('(', '').replace(')', '').replace(',', '').replace('Tr', '0').replace('g 미만', '0')
            nutrition_values.append(float(value_str))
        else:
            nutrition_values.append(0)

    total_weight = sum(nutrition_values)

    if total_weight == 0:
        return ''

    # 각 영양소의 비중 계산
    proportions = [value / total_weight for value in nutrition_values]

    # 가장 비중이 큰 영양소의 인덱스 찾기
    max_index = proportions.index(max(proportions))

    # 대표 영양소 선택
    representative_nutrient = nutrition_columns[max_index]

    return representative_nutrient

input_file = '음식.xlsx'
output_file = '식품데이터.xlsx'
sheet_name = 'Sheet0'

nutrition_columns = ['단백질(g)', '지방(g)', '탄수화물(g)', '당류(g)', '나트륨(mg)']

if not os.path.exists(output_file):
    df = pd.DataFrame(columns=['DB군', '식품명', '식품대분류', '에너지(㎉)', '대표영양소'])
else:
    df = pd.read_excel(output_file, sheet_name=sheet_name)

df_original = pd.read_excel(input_file, sheet_name=sheet_name)

df_original = df_original[['DB군', '식품명', '식품대분류', '에너지(㎉)', '단백질(g)', '지방(g)', '탄수화물(g)', '당류(g)', '나트륨(mg)']]

df_original.fillna(0, inplace=True)
df_original.replace('-', 0, inplace=True)  # '-'을 0으로 대체

df_original['대표영양소'] = df_original.apply(calculate_representative_nutrient, axis=1)

df = pd.concat([df, df_original], ignore_index=True)

# 5개 컬럼 삭제
df.drop(['단백질(g)', '지방(g)', '탄수화물(g)', '당류(g)', '나트륨(mg)'], axis=1, inplace=True)

df.to_excel(output_file, index=False, sheet_name=sheet_name)
