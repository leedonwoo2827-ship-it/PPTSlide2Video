# PPTSlide2Video

> PowerPoint 파일을 **애니메이션이 살아있는 MP4 영상**으로 자동 변환하는 도구

PowerPoint의 모든 애니메이션, 전환효과, 레이아웃을 그대로 보존하여
슬라이드별 음성(WAV/MP3)과 자막을 합성한 발표 영상을 만들어줍니다.

---

## 빠른 시작

### 1단계. 사전 준비 (PC에 한 번만)

| 필수 프로그램 | 다운로드 | 비고 |
|---|---|---|
| Microsoft PowerPoint | Office 2013 이상 (2013 / 2016 / 2019 / 365) | 데스크톱 버전 필요 |
| Python 3.11 이상 | [python.org/downloads](https://www.python.org/downloads/) | 설치 시 "Add to PATH" 체크 |
| FFmpeg | [ffmpeg.org/download](https://ffmpeg.org/download.html) | Windows: PATH에 추가 / Mac: `brew install ffmpeg` |

### 2단계. PPTSlide2Video 설치 (PC에 한 번만)

1. ZIP 파일을 받아 원하는 위치에 압축 해제
2. **`setup.bat`**(Windows) 또는 **`setup.sh`**(Mac/Linux) 실행

```
  ============================================================
    PPTSlide2Video — 초기 설치
  ============================================================

  [1/3] Python 패키지 설치 중...
  [2/3] FFmpeg 확인 중...    FFmpeg OK
  [3/3] PowerPoint 확인 중... PowerPoint OK

  ============================================================
    설치 완료!
  ============================================================
```

> 이 과정은 최초 1회만 필요합니다.

### 3단계. 작업 폴더 준비

PPTX 파일과 음성 파일을 같은 폴더에 넣습니다:

```
📁 작업 폴더/
  ├── 발표자료.pptx
  ├── 01.wav          ← 슬라이드 1번 음성
  ├── 02.wav          ← 슬라이드 2번 음성
  ├── ...
  └── 26.wav          ← 슬라이드 26번 음성
```

### 4단계. 변환 실행 (3가지 방법)

---

#### A. Windows에서 실행

**드래그 앤 드롭 (가장 간편)**

탐색기에서 PPTX 파일을 `convert.bat` 위에 끌어다 놓으면
같은 폴더의 WAV를 자동 탐색하여 변환을 시작합니다.

**명령 프롬프트**

```bat
python run_local.py "D:\작업폴더\발표자료.pptx"
```

음성 파일이 다른 폴더에 있으면:

```bat
python run_local.py "D:\발표자료.pptx" "D:\음성폴더"
```

---

#### B. Mac에서 실행

터미널에서 실행합니다:

```bash
python3 run_local.py "/Users/me/작업폴더/발표자료.pptx"
```

음성 파일이 다른 폴더에 있으면:

```bash
python3 run_local.py "/Users/me/발표자료.pptx" "/Users/me/음성폴더"
```

> Mac에서는 PowerPoint for Mac이 설치되어 있어야 합니다.
> 변환 중 PowerPoint가 자동으로 열리고 닫힙니다.

---

#### C. Claude Desktop에서 실행 (3분 이하 / 슬라이드 10장 이하)

Claude Desktop MCP로 연결하면 대화로 변환할 수 있습니다.
짧은 영상에 적합합니다.

```
  발표자료.pptx를 MP4로 변환해줘. 음성 파일은 D:\음성폴더에 있어.
```

<details>
<summary>Claude Desktop 설정 방법</summary>

`%APPDATA%\Claude\claude_desktop_config.json` (Windows) 또는
`~/Library/Application Support/Claude/claude_desktop_config.json` (Mac)에 추가:

```json
{
  "mcpServers": {
    "pptslide2video": {
      "command": "python",
      "args": ["/path/to/PPTSlide2Video/server.py"]
    }
  }
}
```

경로를 실제 설치 위치로 수정한 뒤 Claude Desktop을 재시작합니다.

</details>

---

## 실행 화면

```
  ============================================================
    PPTSlide2Video — PPTX → MP4 변환
  ============================================================
    PPTX : D:\작업폴더\발표자료.pptx
    음성 : D:\작업폴더
    출력 : D:\작업폴더\발표자료_output.mp4
    자막 : hard
  ============================================================

    PowerPoint 네이티브 내보내기 중... (20~30분 소요될 수 있습니다)

    [  5%] Parsing PPTX…
    [ 25%] Rendering slides (this may take a while)…
    [PowerPoint export] 30s elapsed...
    [PowerPoint export] 60s elapsed...
    ...
    [ 75%] 음성 합치는 중…
    [ 90%] Adding subtitles (from slide notes)…
    [100%] Done!

    [완료] D:\작업폴더\발표자료_output.mp4
    파일 크기: 30.6 MB
    소요 시간: 1080초 (18.0분)
```

---

## 폴더 구조

```
📁 PPTSlide2Video/                ← ZIP 압축 푼 위치 (도구)
  ├── setup.bat / setup.sh       ← 최초 1회 실행 (설치)
  ├── convert.bat                ← PPTX 드래그 앤 드롭 (Windows)
  ├── run_local.py               ← 명령줄 실행용 (Windows / Mac)
  ├── server.py                  ← Claude Desktop MCP 서버
  └── slidecast/                 ← 내부 코드
```

```
📁 작업 폴더/                     ← 어디든 상관없음
  ├── 발표자료.pptx
  ├── 01.wav ~ 26.wav
  └── 발표자료_output.mp4        ← 결과 파일 (자동 생성)
```

---

## 음성 파일 형식

### 슬라이드별 번호 파일 (추천)

| 형식 | 예시 |
|------|------|
| 두 자리 번호 | `01.wav`, `02.wav` |
| 세 자리 번호 | `001.wav`, `002.wav` |
| 한 자리 번호 | `1.wav`, `2.wav` |
| 프리픽스-번호 | `chapter_01.mp3`, `2-01.wav` |

> 지원 형식: wav, mp3, m4a, aac

### 단일 음성 파일

전체 나레이션이 하나의 파일이면 그것도 됩니다.

```
📁 작업 폴더/
  ├── 발표자료.pptx
  └── narration.mp3
```

> 음성 파일이 없어도 됩니다. 영상만 만들어집니다.

---

## 자막

별도 자막 파일이 필요 없습니다.
PowerPoint **노트 패널**에 입력한 텍스트가 자동으로 자막이 됩니다.

```
┌──────────────────────────────────┐
│         슬라이드 화면             │
└──────────────────────────────────┘
  노트:
  안녕하세요, 오늘 발표를 시작하겠습니다.
```

- 각 슬라이드 노트 → 해당 슬라이드 재생 시간 동안 자막 표시
- 음성(WAV) 길이에 맞춰 자막 타이밍 자동 동기화

---

## 동작 원리

```
PPTX ──→ PowerPoint 네이티브 내보내기 ──→ MP4 (애니메이션 보존)
            (Windows: COM API)                │
            (Mac: AppleScript)                │
                                              │
WAV 파일들 ──→ FFmpeg 오디오 합성 ──────────→ 영상 + 음성
                                              │
슬라이드 노트 ──→ SRT 자막 생성 ──→ FFmpeg 자막 합성 ──→ 최종 MP4
```

---

## 자주 묻는 질문

**음성 파일이 없어도 되나요?**
네. 기본 5초씩 각 슬라이드가 표시되는 영상이 만들어집니다.

**슬라이드 수와 WAV 파일 수가 다르면요?**
WAV가 있는 슬라이드까지만 음성이 적용되고, 나머지는 기본 시간으로 표시됩니다.

**변환이 너무 오래 걸려요.**
PowerPoint가 애니메이션을 프레임 단위로 렌더링하기 때문입니다.
26슬라이드 기준 약 20~30분 소요됩니다. 변환 중 창을 닫지 마세요.

---

## 라이선스

MIT
