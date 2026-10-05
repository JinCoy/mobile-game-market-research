MOBILE GAME MARKET RESEARCH WEBSITE — PROJECT RULES
이 문서는 프로젝트 전체에서 가장 우선되는 개발 및 디자인 지침이다. 아래 규칙은 기능 구현, UI/UX 디자인, 데이터 구조 설계, 리팩터링, 테스트 과정 전체에 적용한다.
1. 프로젝트 목적
이 프로젝트는 모바일 게임을 개발하기 전에 시장 상황을 빠르고 정확하게 이해하기 위한 Mobile Game Market Research Dashboard이다.
현재 보유하고 있는 2025년 모바일 게임 시장조사 자료를 기반으로 시작하지만, 시장조사는 완료된 것이 아니며 앞으로 지속적으로 새로운 게임, 순위, 장르, BM, 마케팅, 아트 스타일, 지역별 데이터가 추가된다.
따라서 이 프로젝트를 일회성 정적 페이지로 만들지 말고, 시장 데이터가 지속적으로 업데이트되는 것을 전제로 확장 가능한 구조로 설계한다.
주요 데이터 출처는 기존 Excel/XLSX 시장조사 파일이며, 해당 데이터를 웹 환경에서 더 빠르게 검색, 필터링, 비교, 분석할 수 있도록 구현한다.

2. 주요 사용자
이 웹사이트의 주요 사용자는 다음과 같다.
* 프로젝트 오너
* 공동 게임 개발자
* 추가로 합류하는 개발자
* 게임 디자이너
* 아트 디자이너
이 사이트는 외부 소비자를 위한 홍보 사이트가 아니라 게임 개발팀 내부에서 실제 시장 분석과 의사결정에 사용하는 리서치 도구이다.
따라서 장식적인 디자인보다 정보의 구조, 가독성, 비교 가능성, 탐색 속도를 우선한다.

3. 사용자가 사이트에서 얻어야 하는 결과
페이지를 처음 본 사람도 모바일 게임 시장에 대한 사전 지식이 없어도 다음 내용을 빠르게 이해할 수 있어야 한다.
* 현재 어떤 모바일 게임이 상위권에 있는가
* 어떤 장르가 시장에서 강한가
* App Store와 Google Play의 차이는 무엇인가
* Top 10 / Top 20 / Top 100에서 장르 분포가 어떻게 달라지는가
* 어떤 게임들이 두 스토어에서 동시에 성공하고 있는가
* 성공한 게임들의 핵심 Gameplay Loop는 무엇인가
* 어떤 Monetization / BM 전략을 사용하는가
* 어떤 Marketing 전략을 사용하는가
* 어떤 Art Direction과 Visual Style이 반복적으로 등장하는가
* 시장에서 발견되는 주요 패턴과 기회는 무엇인가
* 앞으로 개발할 게임에서 참고할 수 있는 시장의 빈 공간이나 기회는 무엇인가
* 마케팅의 방법과 후킹 요소에 대해 알려줄 것.
* 마케팅으로 인한 다운로드 수 인과관계가 어떤가
최종 목표는 페이지를 본 사람이 여러 개의 Excel 파일을 직접 읽지 않아도 모바일 게임 시장의 구조를 한눈에 이해할 수 있게 만드는 것이다.

4. 기술
기본 프레임워크는 다음을 사용한다.
* Next.js
* TypeScript
* GitHub
프로젝트는 GitHub Repository에 업로드하여 다른 개발자와 공유할 수 있어야 한다.
가능하면 Next.js App Router 기반으로 구축한다.
추가 라이브러리가 필요한 경우 프로젝트 목적에 맞는 안정적이고 유지보수가 쉬운 라이브러리를 스스로 판단하여 선택한다.
아이콘은 반드시 Phosphor Icons를 사용한다.
아이콘을 직접 SVG로 새로 만들거나 다른 아이콘 라이브러리를 혼용하지 않는다.

5. 데이터 설계 원칙
현재 시장조사는 계속 진행 중이다.
따라서 데이터를 페이지 안에 직접 hard-code하지 않는다.
게임 데이터와 UI를 가능한 한 분리한다.
향후 다음 데이터가 쉽게 추가될 수 있는 구조를 사용한다.
* 연도
* 날짜
* 국가 / 지역
* 플랫폼
* Store Rank
* 게임 이름
* Publisher
* Developer
* Genre
* Subgenre
* Core Gameplay
* Gameplay Loop
* Monetization / BM
* Marketing Strategy
* Art Style
* Target Audience
* Player Age
* Download / Revenue 관련 데이터
* App Store / Google Play 중복 여부
* 분석 메모
* 데이터 출처
* 마지막 업데이트 날짜
Excel/XLSX 데이터가 변경되더라도 웹사이트 전체 구조를 다시 작성하지 않아도 되도록 설계한다.
가능하면 하나의 공통 Game Data Schema를 정의한다.

6. 디자인 방향
이 사이트는 단순한 SaaS 관리자 페이지처럼 보이면 안 된다.
Editorial Research Report + Data Dashboard + Game Market Intelligence Tool이 결합된 형태를 목표로 한다.
정보량이 많아도 복잡하게 느껴지지 않아야 한다.
시각적 우선순위가 명확해야 하며 숫자, 차트, 게임 목록, 핵심 인사이트가 자연스럽게 연결되어야 한다.

7. 전체 화면 규격
Desktop 기준:
* 전체 기준 폭: 1920px
* Main Content Container: 최대 1280px
* 기본 좌우 여백: 80px
* 지나치게 넓은 full-width text section을 사용하지 않는다.
Tablet과 Mobile 화면도 반드시 대응한다.
Responsive Design은 개발 마지막에 추가하는 것이 아니라 처음부터 고려한다.
최소한 다음 화면 크기를 직접 확인한다.
* Desktop 1920px
* Laptop 1440px
* Tablet
* Mobile

8. 디자인 금지사항
다음 디자인 패턴을 사용하지 않는다.
* Emoji 사용 금지
* 의미 없이 과도하게 큰 Hero Typography 금지
* 거의 모든 Section을 반복적인 Left / Right 2-column 구조로 만드는 것 금지
* 불필요하게 많은 Card 사용 금지
* 모든 정보를 Rounded Rectangle 안에 넣는 Dashboard 디자인 금지
* 과도한 Gradient 사용 금지
* 불필요한 Glow 효과 금지
* 의미 없는 Decorative Graphic 금지
* 같은 정보를 여러 위치에서 반복하는 것 금지
* AI가 만든 Landing Page처럼 보이는 구성 금지
정보 자체가 디자인의 중심이 되어야 한다.

9. 글쓰기 규칙
사이트 안의 모든 설명 문구는 다음 원칙을 따른다.
* Emoji 사용 금지
* AI 특유의 과장된 표현 금지
* 불필요한 수식어 금지
* 같은 의미를 반복하지 않는다.
* 지나치게 많은 줄바꿈을 사용하지 않는다.
* 실제 시장조사 보고서를 읽는 것처럼 자연스럽고 전문적인 문장을 사용한다.
* 숫자와 사실을 먼저 보여주고 필요한 경우에만 설명을 추가한다.
"Unlock", "Discover", "Powerful insights", "Revolutionize", "Dive into"처럼 일반적인 AI/SaaS 마케팅 문구를 사용하지 않는다.

10. 개발 작업 방식
개발 과정에서 사용자가 직접 할 필요가 없는 작업을 사용자에게 요청하지 않는다.
Codex가 스스로 수행할 수 있는 작업은 직접 수행한다.
예:
* 패키지 설치
* 프로젝트 구조 생성
* 코드 작성
* 파일 생성
* 리팩터링
* 서버 실행
* Build 확인
* Error 확인
* 브라우저 실행
* 페이지 직접 확인
* Responsive 확인
* 콘솔 에러 확인
* TypeScript 오류 확인
* Layout 문제 수정
"직접 실행해서 확인해 주세요"라고 사용자에게 넘기지 않는다.
실제로 실행하고 확인한 뒤 결과만 보고한다.

11. 브라우저 검증
UI 구현 후 반드시 실제 브라우저에서 확인한다.
코드만 작성하고 디자인 작업이 끝났다고 판단하지 않는다.
다음 내용을 직접 확인한다.
* Layout
* Typography
* Spacing
* Navigation
* Chart readability
* Table readability
* Filter behavior
* Overflow
* Horizontal scrolling
* Mobile layout
* Tablet layout
* Desktop layout
* Console error
* Broken component
* Missing asset
* Loading behavior
문제가 발견되면 사용자에게 수정 여부를 묻지 말고 직접 수정한 후 다시 확인한다.

12. Design Skill 관련 규칙
이 프로젝트의 디자인 작업에서는 site skill을 절대 사용하지 않는다.
페이지 구조와 디자인 시스템은 프로젝트의 시장조사 목적, 제공된 Reference Image, 데이터 구조를 기반으로 직접 설계한다.

13. Reference Image
사용자가 제공하는 Reference Image는 그대로 복제하기 위한 것이 아니다.
다음 요소를 분석하여 프로젝트에 맞게 재구성한다.
* Information hierarchy
* Grid
* Typography
* Spacing
* Navigation
* Data visualization
* Table structure
* Card usage
* Section rhythm
* Density
* Interaction pattern
Reference의 좋은 구조는 활용하되 브랜드와 UI는 이 프로젝트만의 디자인으로 만든다.

14. 확장성
현재는 2025년 미국 모바일 게임 Top 100 분석을 중심으로 시작하더라도 구조적으로 다음 확장을 지원해야 한다.
United States → Latin America → Asia → 국가별 시장 비교
2025 → 2026 → 이후 연도 데이터
App Store → Google Play → Store Comparison
Ranking → Genre → Gameplay → Monetization → Marketing → Art Direction → Audience → Opportunity Analysis
새로운 시장조사 데이터가 추가될 때 페이지를 새로 만드는 방식이 아니라 동일한 시스템 안에서 확장되어야 한다.

15. 프로젝트 판단 원칙
구현 과정에서 사소한 선택마다 사용자에게 확인을 요청하지 않는다.
다음 조건에 해당하면 스스로 가장 합리적인 방법을 선택한다.
* 일반적인 UI 구현 방식
* Component 구조
* Folder 구조
* Naming
* Responsive breakpoint
* Utility library
* Chart library
* State management 방식
* Data transformation 방식
* Loading / Empty / Error state
* Accessibility 개선
* Performance 개선
사용자의 결정이 반드시 필요한 경우에만 질문한다.
프로젝트의 최우선 판단 기준은 다음 순서다.
1. 시장 데이터를 정확하게 이해할 수 있는가
2. 정보를 빠르게 찾을 수 있는가
3. 서로 다른 데이터를 쉽게 비교할 수 있는가
4. 데이터가 추가되어도 구조가 유지되는가
5. 실제 개발팀이 반복해서 사용할 수 있는가
6. 시각적으로 전문적이고 정돈되어 있는가
이 기준보다 장식적인 디자인을 우선하지 않는다.
