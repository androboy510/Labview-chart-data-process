import pandas as pd
import numpy as np

def create_sample_data():
    """
    테스트용 샘플 데이터를 생성합니다.
    """
    # 샘플 데이터 생성
    np.random.seed(42)  # 재현 가능한 결과를 위해 시드 설정
    
    # 1000개의 행 생성
    n_rows = 1000
    
    # sample 수 열 (0부터 시작)
    sample_count = np.arange(n_rows)
    
    # DOF1, DOF2 데이터 생성 (랜덤 값)
    dof1 = np.random.normal(0, 1, n_rows)
    dof2 = np.random.normal(0, 1, n_rows)
    
    # sample 1, sample 2 열 (sample 수와 동일한 값)
    sample1 = sample_count
    sample2 = sample_count
    
    # DataFrame 생성
    data = {
        ' sample 수 ': sample_count,  # 의도적으로 공백 포함
        'DOF1': dof1,
        ' sample 1 ': sample1,  # 의도적으로 공백 포함
        'DOF2': dof2,
        ' sample 2 ': sample2   # 의도적으로 공백 포함
    }
    
    df = pd.DataFrame(data)
    
    # 엑셀 파일로 저장
    output_file = 'sample_data.xlsx'
    df.to_excel(output_file, index=False)
    
    print(f"샘플 데이터가 생성되었습니다: {output_file}")
    print(f"데이터 형태: {df.shape[0]}행 x {df.shape[1]}열")
    print("\n처음 5행:")
    print(df.head())
    
    return output_file

if __name__ == "__main__":
    create_sample_data() 