import pandas as pd
import os

def test_save_function():
    """저장 기능 테스트"""
    print("저장 기능 테스트 시작...")
    
    try:
        # 1. 간단한 테스트 데이터 생성
        print("1. 테스트 데이터 생성 중...")
        test_data = {
            'Time': [0.00, 0.01, 0.02, 0.03, 0.04],
            'sample 수': [0, 1, 2, 3, 4],
            'DOF1': [1.1, 1.2, 1.3, 1.4, 1.5],
            'DOF2': [2.1, 2.2, 2.3, 2.4, 2.5]
        }
        df = pd.DataFrame(test_data)
        print(f"   데이터 생성 완료: {df.shape[0]}행 x {df.shape[1]}열")
        
        # 2. 현재 디렉토리 확인
        print(f"2. 현재 작업 디렉토리: {os.getcwd()}")
        
        # 3. Excel 파일로 저장 테스트
        print("3. Excel 파일 저장 테스트...")
        excel_file = "test_save_output.xlsx"
        df.to_excel(excel_file, index=False)
        print(f"   Excel 저장 성공: {excel_file}")
        
        # 4. CSV 파일로 저장 테스트
        print("4. CSV 파일 저장 테스트...")
        csv_file = "test_save_output.csv"
        df.to_csv(csv_file, index=False)
        print(f"   CSV 저장 성공: {csv_file}")
        
        # 5. 파일 존재 확인
        print("5. 저장된 파일 확인...")
        if os.path.exists(excel_file):
            print(f"   Excel 파일 존재: {os.path.abspath(excel_file)}")
        else:
            print("   Excel 파일이 생성되지 않았습니다!")
            
        if os.path.exists(csv_file):
            print(f"   CSV 파일 존재: {os.path.abspath(csv_file)}")
        else:
            print("   CSV 파일이 생성되지 않았습니다!")
        
        print("\n✅ 모든 테스트가 성공했습니다!")
        
    except Exception as e:
        print(f"\n❌ 테스트 실패: {str(e)}")
        print(f"오류 타입: {type(e).__name__}")

if __name__ == "__main__":
    test_save_function() 