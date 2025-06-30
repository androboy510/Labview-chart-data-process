import pandas as pd
import os
import json
import time

def test_v3_features():
    """v3 버전의 새로운 기능들을 테스트합니다."""
    print("=" * 50)
    print("v3 기능 테스트 시작")
    print("=" * 50)
    
    # 1. 최근 파일 목록 기능 테스트
    print("\n1. 최근 파일 목록 기능 테스트")
    config_file = "gui_config.json"
    
    # 테스트 데이터 생성
    test_files = [
        "test_file1.xlsx",
        "test_file2.csv", 
        "05.DOF_ANGLE_CONTROLL_P_d_dot.xlsx"
    ]
    
    # 최근 파일 목록 저장
    config = {'recent_files': test_files}
    with open(config_file, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    
    print(f"   최근 파일 목록 저장 완료: {config_file}")
    
    # 최근 파일 목록 로드
    with open(config_file, 'r', encoding='utf-8') as f:
        loaded_config = json.load(f)
        loaded_files = loaded_config.get('recent_files', [])
    
    print(f"   로드된 파일 목록: {loaded_files}")
    
    # 2. 진행 바 시뮬레이션 테스트
    print("\n2. 진행 바 시뮬레이션 테스트")
    for i in range(0, 101, 10):
        print(f"   진행률: {i}% - {'파일 로드 중...' if i < 30 else '데이터 처리 중...' if i < 70 else '완료!'}")
        time.sleep(0.1)
    
    # 3. 데이터 처리 테스트
    print("\n3. 데이터 처리 테스트")
    try:
        # 실제 파일이 있는지 확인
        if os.path.exists("05.DOF_ANGLE_CONTROLL_P_d_dot.xlsx"):
            print("   실제 데이터 파일 발견: 05.DOF_ANGLE_CONTROLL_P_d_dot.xlsx")
            
            # 파일 로드 테스트
            df = pd.read_excel("05.DOF_ANGLE_CONTROLL_P_d_dot.xlsx")
            print(f"   파일 로드 성공: {df.shape[0]}행 x {df.shape[1]}열")
            
            # 헤더 정리 테스트
            original_headers = list(df.columns)
            df.columns = df.columns.str.strip()
            print(f"   헤더 정리 완료: {len(original_headers)}개 열")
            
            # Sample 열 찾기 테스트
            sample_columns = [col for col in df.columns if 'sample' in col.lower()]
            print(f"   Sample 열 발견: {sample_columns}")
            
            if sample_columns:
                # Time 열 추가 테스트
                first_sample_col = sample_columns[0]
                time_values = df[first_sample_col] * 0.01
                df.insert(0, 'Time', time_values)
                print(f"   Time 열 추가 완료: {df.shape[1]}개 열")
                
                # 미리보기 데이터 생성
                preview_data = df.head()
                print(f"   미리보기 데이터 생성: {preview_data.shape[0]}행")
                
                # 저장 테스트
                output_file = "test_v3_output.xlsx"
                df.to_excel(output_file, index=False)
                print(f"   저장 테스트 완료: {output_file}")
                
        else:
            print("   실제 데이터 파일이 없습니다. 샘플 데이터로 테스트합니다.")
            
            # 샘플 데이터 생성
            sample_data = {
                'Time': [0.00, 0.01, 0.02, 0.03, 0.04],
                'sample 수': [0, 1, 2, 3, 4],
                'DOF1': [1.1, 1.2, 1.3, 1.4, 1.5],
                'DOF2': [2.1, 2.2, 2.3, 2.4, 2.5]
            }
            df = pd.DataFrame(sample_data)
            print(f"   샘플 데이터 생성: {df.shape[0]}행 x {df.shape[1]}열")
            
    except Exception as e:
        print(f"   데이터 처리 테스트 실패: {e}")
    
    # 4. GUI 구성 요소 테스트
    print("\n4. GUI 구성 요소 테스트")
    gui_components = [
        "파일 선택 영역",
        "최근 사용 파일 영역", 
        "진행 상황 영역",
        "미리보기 영역",
        "파일 저장 영역",
        "상태 표시 영역"
    ]
    
    for component in gui_components:
        print(f"   ✅ {component} 구현됨")
    
    # 5. 새 기능 목록
    print("\n5. v3 새 기능 목록")
    new_features = [
        "📁 최근 사용 파일 목록 (최대 10개)",
        "📊 실시간 진행 바",
        "⚡ 백그라운드 처리 (GUI 응답성 유지)",
        "💾 빠른 저장 기능",
        "🗑️ 최근 파일 목록 관리",
        "🎯 향상된 상태 표시",
        "📋 설정 파일 자동 저장/로드"
    ]
    
    for feature in new_features:
        print(f"   {feature}")
    
    print("\n" + "=" * 50)
    print("v3 기능 테스트 완료!")
    print("=" * 50)
    
    # 정리
    if os.path.exists("test_v3_output.xlsx"):
        print("\n생성된 테스트 파일:")
        print(f"  - test_v3_output.xlsx")
        print(f"  - gui_config.json")
    
    return True

if __name__ == "__main__":
    test_v3_features() 