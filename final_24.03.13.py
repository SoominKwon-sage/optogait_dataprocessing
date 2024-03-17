import pandas as pd
import os
import csv
import aspose.cells
from aspose.cells import Workbook

# XML 파일이 있는 디렉토리 경로
xml_dir = 'C:\\Users\\권수민\\Desktop\\xml2'

# CSV 파일을 저장할 디렉토리 경로
csv_dir = 'C:\\Users\\권수민\\Desktop\\python_test'

# XML 파일이 있는 디렉토리의 모든 파일 목록 가져오기
file_list = os.listdir(xml_dir)

# 각 XML 파일에 대해 반복
for file_name in file_list:
    # 파일 확장자가 .xml인 경우에만 처리
    if file_name.endswith('.xml'):
        # XML 파일 경로
        xml_path = os.path.join(xml_dir, file_name)
        
        # Workbook 객체 생성 및 XML 파일 로드
        workbook = Workbook(xml_path)
        
        # CSV 파일 경로 생성
        csv_path = os.path.join(csv_dir, file_name[:-4] + '.csv')
        
        # CSV 파일로 저장
        workbook.save(csv_path)
        
        print(f"{file_name}을 {csv_path}로 변환하여 저장했습니다.")

folder_path = 'C:/Users/권수민/Desktop/python_test'
files = os.listdir(folder_path)
csv_files = [os.path.join(folder_path, file) for file in files if file.endswith('.csv')]

# 출력 자동화 함수 작성
def get_mean_by_foot(data, foot_dir: str, target_column: str):
    if foot_dir in ['L', 'R']:
        return round(data.loc[data['L/R'] == foot_dir, target_column].mean(), 2)
    else:
        print('입력이 잘못되었습니다.')

# 표준편차 계산 함수
def get_std_by_foot(data, foot_dir: str, target_column: str):
    if foot_dir in ['L', 'R']:
        return round(data.loc[data['L/R'] == foot_dir, target_column].std(), 2)
    else:
        print('입력이 잘못되었습니다.')

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
    'Swing phase',
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

results = {}

for file in csv_files:
    data = pd.read_csv(file, encoding='cp1252', dtype={'L/R': str})
    
    print(data.columns)
    print(data.info())
    print(data['L/R'].head())
    
    for target_title in column_to_check:
        left_mean = get_mean_by_foot(data, "L", target_title)
        right_mean = get_mean_by_foot(data, "R", target_title)
        left_std = get_std_by_foot(data, "L", target_title)
        right_std = get_std_by_foot(data, "R", target_title)
        
        if target_title not in results:
            results[target_title] = {
                'Left_Mean': [],
                'Right_Mean': [],
                'Left_Std': [],
                'Right_Std': []
            }
        
        results[target_title]['Left_Mean'].append(left_mean)
        results[target_title]['Right_Mean'].append(right_mean)
        results[target_title]['Left_Std'].append(left_std)
        results[target_title]['Right_Std'].append(right_std)

file_path = os.path.join('C:', os.sep, 'Users', '권수민', 'Desktop', 'python_results2', 'results.csv')

# CSV 파일 생성
with open(file_path, mode='w', newline='') as file:
    writer = csv.writer(file)
    
    # 첫 번째 행에 column을 원하는 순서대로 작성하고 3칸씩 비우기
    header_row1 = ['', ''] + [column if i % 4 == 0 else '' for i, column in enumerate([col for col in column_to_check for _ in range(4)])]
    writer.writerow(header_row1)
    
    # 두 번째 행에 'Left_Mean', 'Right_Mean', 'Left_Std', 'Right_Std' 반복 작성
    header_row2 = ['', ''] + ['Left_Mean', 'Right_Mean', 'Left_Std', 'Right_Std'] * len(column_to_check)
    writer.writerow(header_row2)
    
    # 세 번째 행부터 데이터 작성
    for i in range(len(csv_files)):
        # 파일명에서 앞 3글자만 추출하여 사용
        row = [os.path.basename(csv_files[i])[:3], 'trial' + str((i % 2) + 1)]
        for column in column_to_check:
            row.extend([
                results[column]['Left_Mean'][i],
                results[column]['Right_Mean'][i],
                results[column]['Left_Std'][i],
                results[column]['Right_Std'][i]
            ])
        writer.writerow(row)