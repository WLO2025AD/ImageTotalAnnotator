# 🧰 ImageTotalAnnotator — Starter Pack

아주 & 대박의 **화면 어디서든 주석** 프로젝트 시작 세트입니다.

## 포함 파일
- `ImageTotalAnnotator.ahk` — 단축키 허브 (AutoHotkey v2)
- `AnnotateHelpers.bas` — Word 매크로(이미지 위 주석 캔버스 추가)
- `README.md` — 사용 설명서

## 단축키 (실행 중)
- **Ctrl+Alt+I** : ImageAnnotator(HTML) 열기
- **Ctrl+Alt+Z** : ZoomIt 드로잉 모드(또는 실행)
- **Ctrl+Alt+S** : Windows 스니핑(Win+Shift+S)
- **Ctrl+Alt+E** : Edge Web Capture (활성 탭에서 Ctrl+Shift+S)
- **Ctrl+Alt+W** : Word 주석 매크로 실행
- **Ctrl+Alt+H** : 도움말
- **Ctrl+Alt+Q** : 종료

## 설치
1) **AutoHotkey v2** 설치 (https://www.autohotkey.com/)
2) `ImageAnnotator.ahk` 경로의 `ImageAnnotatorPath` 와 `ZoomItPath` 수정
3) 더블클릭으로 스크립트 실행 (트레이 아이콘 표시)
4) (선택) 시작프로그램에 등록하여 자동 실행

## Word 매크로 설치
1) Word 열기 → **Alt+F11** (VBA 에디터)
2) **Insert → Module** → `AnnotateHelpers.bas` 내용 붙여넣기 → 저장
3) Word에서 이미지 선택 → **Ctrl+Alt+W** (또는 Alt+F8 → AnnotateSelectedImage 실행)
   - 도형/화살표/텍스트는 `삽입 → 도형`으로 계속 추가

## 참고
- ZoomIt 기본 드로잉 단축키: **Ctrl+2**
- Edge Web Capture: **Ctrl+Shift+S** (Edge 활성 시), 그 외에는 스니핑 도구 호출
- **Zoomig v9.01** 설치 (https://learn.microsoft.com/ko-kr/sysinternals/downloads/zoomit)

