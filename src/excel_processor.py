import pandas as pd
import os
import sys

def load_file(file_path):
    """
    엑셀 또는 CSV 파일을 pandas DataFrame으로 로드합니다.
    
    Args:
        file_path (str): 파일 경로
        
    Returns:
        pandas.DataFrame: 로드된 데이터
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
    
    file_extension = os.path.splitext(file_path)[1].lower()
    
    try:
        if file_extension == '.xlsx':
            df = pd.read_excel(file_path)
        elif file_extension == '.csv':
            df = pd.read_csv(file_path)
        else:
            raise ValueError(f"지원하지 않는 파일 형식입니다: {file_extension}")
        
        print(f"파일이 성공적으로 로드되었습니다: {file_path}")
        print(f"데이터 형태: {df.shape[0]}행 x {df.shape[1]}열")
        return df
        
    except Exception as e:
        print(f"파일 로드 중 오류가 발생했습니다: {e}")
        raise

def clean_headers(df):
    """
    DataFrame의 헤더(열 이름)에서 앞뒤 공백을 제거합니다.
    
    Args:
        df (pandas.DataFrame): 처리할 DataFrame
        
    Returns:
        pandas.DataFrame: 헤더가 정리된 DataFrame
    """
    df.columns = df.columns.str.strip()
    print("헤더 정리가 완료되었습니다.")
    return df

def process_sample_columns(df):
    """
    'sample' 관련 열들을 처리하여 첫 번째 'sample' 열만 남깁니다.
    
    Args:
        df (pandas.DataFrame): 처리할 DataFrame
        
    Returns:
        tuple: (처리된 DataFrame, 남겨진 sample 열 이름)
    """
    # 'sample'을 포함하는 모든 열 찾기
    sample_columns = [col for col in df.columns if 'sample' in col.lower()]
    
    if not sample_columns:
        print("경고: 'sample'을 포함하는 열을 찾을 수 없습니다.")
        return df, None
    
    print(f"발견된 'sample' 관련 열들: {sample_columns}")
    
    # 첫 번째 'sample' 열만 남기고 나머지 삭제
    first_sample_col = sample_columns[0]
    columns_to_drop = sample_columns[1:]
    
    if columns_to_drop:
        df = df.drop(columns=columns_to_drop)
        print(f"삭제된 열들: {columns_to_drop}")
    
    print(f"남겨진 'sample' 열: {first_sample_col}")
    return df, first_sample_col

def add_time_column(df, sample_column_name):
    """
    DataFrame의 첫 번째 위치에 'Time' 열을 추가하고 sample 열 값에 0.01을 곱합니다.
    
    Args:
        df (pandas.DataFrame): 처리할 DataFrame
        sample_column_name (str): sample 열 이름
        
    Returns:
        pandas.DataFrame: Time 열이 추가된 DataFrame
    """
    if sample_column_name is None:
        print("경고: sample 열이 없어 Time 열을 추가할 수 없습니다.")
        return df
    
    # Time 열 계산 (sample 값 * 0.01)
    time_values = df[sample_column_name] * 0.01
    
    # 첫 번째 위치에 Time 열 삽입
    df.insert(0, 'Time', time_values)
    
    print("'Time' 열이 성공적으로 추가되었습니다.")
    return df

def save_processed_file(df, output_path):
    """
    처리된 DataFrame을 파일로 저장합니다.
    
    Args:
        df (pandas.DataFrame): 저장할 DataFrame
        output_path (str): 출력 파일 경로
    """
    try:
        file_extension = os.path.splitext(output_path)[1].lower()
        
        if file_extension == '.xlsx':
            df.to_excel(output_path, index=False)
        elif file_extension == '.csv':
            df.to_csv(output_path, index=False)
        else:
            raise ValueError(f"지원하지 않는 출력 파일 형식입니다: {file_extension}")
        
        print(f"처리된 파일이 성공적으로 저장되었습니다: {output_path}")
        
    except Exception as e:
        print(f"파일 저장 중 오류가 발생했습니다: {e}")
        raise

def main():
    """
    메인 함수: 전체 프로세스를 실행합니다.
    """
    print("=" * 50)
    print("엑셀 파일 프로세싱 프로그램")
    print("=" * 50)
    
    try:
        # 1. 파일 로드
        input_file_path = input("처리할 파일 경로를 입력하세요: ").strip()
        df = load_file(input_file_path)
        
        # 2. 헤더 정리
        df = clean_headers(df)
        
        # 3. 'sample' 열 처리
        df, sample_column_name = process_sample_columns(df)
        
        # 4. 'Time' 열 삽입 및 계산
        df = add_time_column(df, sample_column_name)
        
        # 5. 결과 출력
        print("\n" + "=" * 50)
        print("처리된 데이터 (처음 5행)")
        print("=" * 50)
        print(df.head())
        
        # 6. 파일 저장 (선택사항)
        while True:
            save_choice = input("\n처리된 파일을 저장하시겠습니까? (y/n): ").strip().lower()
            if save_choice in ['y', 'yes', '예']:
                output_file_path = input("저장할 파일 경로를 입력하세요 (엔터 시 자동 저장): ").strip()
                if not output_file_path:
                    # 기존 파일명 뒤에 _processed 붙이기
                    base_name, ext = os.path.splitext(os.path.basename(input_file_path))
                    output_file_path = f"{base_name}_processed{ext}"
                save_processed_file(df, output_file_path)
                break
            elif save_choice in ['n', 'no', '아니오']:
                print("파일 저장을 건너뜁니다.")
                break
            else:
                print("y(예) 또는 n(아니오)로 입력해 주세요.")
        
        print("\n프로그램이 성공적으로 완료되었습니다!")
        
    except KeyboardInterrupt:
        print("\n\n프로그램이 사용자에 의해 중단되었습니다.")
    except Exception as e:
        print(f"\n오류가 발생했습니다: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 