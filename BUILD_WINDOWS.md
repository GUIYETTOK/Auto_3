윈도우 EXE 빌드/배포 가이드

구성 목표
- exe + DB 폴더만 복사해서 바로 사용

폴더 구조 예시
Auto_3/
  견적서자동화.exe
  DB/
    Templet.xlsx
    (견적서/견적의뢰서 파일들)
  output/

윈도우에서 빌드 방법
1) 윈도우 PC에 프로젝트 폴더 복사
2) cmd 또는 PowerShell에서 프로젝트 폴더로 이동
3) 의존성 설치
   python -m pip install -r requirements.txt
4) PyInstaller 설치
   python -m pip install pyinstaller
5) exe 생성
   python -m pyinstaller --onefile --name 견적서자동화 run_app.py

생성 결과
- dist/견적서자동화.exe

배포 방법
- dist/견적서자동화.exe 를 프로젝트 루트로 복사
- DB 폴더와 output 폴더를 같은 위치에 둠

실행 방법
- 견적서자동화.exe 더블클릭
- 브라우저가 자동으로 열리며 콘솔 UI 표시

주의
- 처음 실행 시 Windows Defender 경고가 나올 수 있음
- 템플릿은 .xlsx 형식이어야 서식/이미지가 유지됨

GitHub Actions 자동 빌드(윈도우 PC 없이)
1) GitHub에 저장소를 올림
2) Actions 탭에서 \"Build Windows EXE\" 실행 (또는 main/master에 push)
3) 완료 후 Artifacts에서 \"견적서자동화-windows\" 다운로드
4) 압축 해제 후 dist_bundle 폴더 사용

#abc@abcui-iMac Auto_3 % python3 -m uvicorn src.web_app:app --reload
