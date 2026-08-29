# OutlookAutomation DRM 복호화 구현 계획

사용자님의 설명을 종합해볼 때, 기존 `decodezip.py` 및 `Tension.slnx`에서 적용했던 **"COM API로 메모리(바이트)를 직접 읽어와 C# 환경에서 저장하여 SaveAs 훅(Hook)을 우회하는 기법"**을 Outlook에도 똑같이 적용할 수 있습니다.

## 🔍 원인 분석
현재 코드가 복호화에 실패하는 이유는 다음과 같습니다:
1. `_mailItem.Attachments.Add(SourceFile)`를 통해 파일을 첨부할 때, 대상 PC에 설치된 Outlook은 신뢰할 수 있는 프로세스(Trusted Process)이므로 Fasoo DRM이 투명하게 복호화된 데이터를 Outlook 메모리에 제공합니다.
2. 하지만 이후 `attachment.SaveAsFile(outputPath)`를 호출할 때, Fasoo DRM은 Outlook의 `SaveAs` 동작을 감지(Hooking)하여 저장되는 파일을 다시 암호화해버립니다.
3. 사용자가 마우스로 드래그 앤 드롭하거나 수동으로 '다른 이름으로 저장'할 때는 윈도우 쉘(Shell) 단의 기능이 개입되어 이 Hooking이 우회되거나 예외 처리되기 때문에 복호화가 되는 것입니다.

## 🛠 제안하는 해결 방안
Outlook COM의 기본 저장 기능인 `SaveAsFile()`을 **사용하지 않고**, Tension.slnx에서 Excel `UsedRange.Value`를 읽어냈던 것처럼 **Outlook 첨부 파일의 바이너리 데이터를 직접 메모리로 추출하여 C#으로 저장**합니다.

1. **PropertyAccessor 우회 기법 도입**: 
   Outlook 2007부터 지원하는 `PropertyAccessor.GetProperty`를 이용해 MAPI 속성인 `PR_ATTACH_DATA_BIN` (0x37010102)를 직접 조회합니다.
2. **C# FileStream으로 저장**: 
   조회해온 `byte[]` 배열(복호화된 원본 데이터)을 C#의 기본 `File.WriteAllBytes`를 통해 저장합니다. C# 앱 자체는 Fasoo의 Hooking 대상이 아니므로 안전하게 복호화된 상태로 디스크에 저장됩니다.
3. **불필요한 UI 노출 최적화 (선택)**: 
   현재 코드에 있는 `DisplayMail()`과 `Thread.Sleep(3000)` 대기열은 실제로 창을 띄워야만 DRM이 해제되는 것이 아니라면 생략하여 앱 동작 속도를 3초 이상 최적화할 수 있습니다.

---

> [!IMPORTANT]
> **사용자 확인 요청 사항 (Open Questions)**
> 1. 메일 창을 화면에 잠깐 띄우는 `DisplayMail()` 과정이 생략되어도 괜찮을까요? (생략 시 깜빡임 없이 백그라운드에서 3초 이상 더 빠르게 처리 가능합니다. 단, 간혹 DRM 정책상 창이 띄워져야만 풀리는 경우도 있어 확인이 필요합니다.)
> 2. 위 계획대로 `PropertyAccessor`를 이용한 메모리 추출 방식으로 코드를 수정해도 될까요?

위 계획과 질문에 대해 피드백을 주시면 즉시 코드 수정을 진행하겠습니다!
