import pandas as pd
import os
import csv

folder_path = 'C:/Users/권수민/Desktop/python_test'
files = os.listdir(folder_path)
csv_files = [os.path.join(folder_path, file) for file in files if file.endswith('.csv')]

print(f"csv_files에 포함된 csv 파일 개수: {len(csv_files)}")

# 출력 자동화 함수 작성
def get_mean_by_foot(data, foot_dir: str, target_column: str):
    if foot_dir in ['L', 'R']: # param 유효성 검사
        return data.loc[data['L/R'] == foot_dir, target_column].mean()
    else: print('입력이 잘못되었습니다.')  # error alert

# 목표 행 Title
column_to_check = [
    'Step time',
    'Stride Time\Cycle',
    'Single support',
    'Single support%',
    'Total double support',
    'Total double support%',
    'Stance phase',
    'Stance phase%',
    'Swing phase%',
    'Load response',
    'Load response%',
    'Pre-Swing',
    'Pre-Swing%',
    'Foot flat',
    'Foot flat%',
    'Propulsive phase',
    'Propulsive phase%',
]

results = []

for file in csv_files:
    data = pd.read_csv(file, encoding='cp1252')
    
    for target_title in column_to_check:
        left_mean = get_mean_by_foot(data, "L", target_title)
        right_mean = get_mean_by_foot(data, "R", target_title)
        results.append({'File': file, 'Column': target_title, 'Left': left_mean, 'Right': right_mean})

file_path = os.path.join('C:', os.sep, 'Users', '권수민', 'Desktop', 'python_results', 'results.csv')

#csv 파일 생성 
with open(file_path, mode='w', newline='') as file:
    fieldnames = ['File', 'Column', 'Left', 'Right']
    writer = csv.DictWriter(file, fieldnames=fieldnames)
        
    writer.writeheader()
    writer.writerows(results)