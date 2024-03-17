import pandas as pd
import os
import csv
import aspose.cells
from aspose.cells import Workbook

# XML 파일이 있는 디렉토리 경로
xml_dir = 'C:\\Users\\권수민\\Desktop\\xml1'

# CSV 파일을 저장할 디렉토리 경로
csv_dir = 'C:\\Users\\권수민\\Desktop\\xml2'

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

#print(f"csv_files에 포함된 csv 파일 개수: {len(csv_files)}")

# 출력 자동화 함수 작성
def get_mean_by_foot(data, foot_dir: str, target_column: str):
    if foot_dir in ['L', 'R']: # param 유효성 검사
        return data.loc[data['L/R'] == foot_dir, target_column].mean()
    else: print('입력이 잘못되었습니다.')  # error alert

#표준편차 계산 함수
def get_std_by_foot(data, foot_dir: str, target_column: str):
    if foot_dir in ['L', 'R']:
        return data.loc[data['L/R'] == foot_dir, target_column].std()
    else: print('입력이 잘못되었습니다.')

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

results = []

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
        results.append({'File': file, 'Column': target_title, 
                        'Left_Mean': left_mean, 'Right_Mean': right_mean, 
                        'Left_Std': left_std, 'Right_Std': right_std})
        
file_path = os.path.join('C:', os.sep, 'Users', '권수민', 'Desktop', 'python_results', 'results.csv')

#csv 파일 생성 
with open(file_path, mode='w', newline='') as file:
    fieldnames = ['File', 'Column', 'Left_Mean', 'Right_Mean', 'Left_Std', 'Right_Std']
    writer = csv.DictWriter(file, fieldnames=fieldnames)
        
    writer.writeheader()
    writer.writerows(results)