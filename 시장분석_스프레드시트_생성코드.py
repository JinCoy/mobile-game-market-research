# =====================================================================
# 모바일 게임 시장 분석 — 스프레드시트 생성 코드 (6단계-4: 남은 빈칸 채우기까지 반영)
#  [1부] 기존 시트 28개를 만드는 코드 (데이터·수식·서식 그대로, 변경 없음)
#  [2부] 만들어진 시트를 파일 2개로 나눠 저장
#        A) 시장조사.xlsx               : 시장·지역·플랫폼·벤치마크 '숫자' 자료
#        B) 게임전략_MVP_마케팅.xlsx   : 성공 비결·MVP·마케팅 '판단' 자료
#  실행: python3 시장분석_스프레드시트_생성코드.py  → 두 파일이 OUT_DIR에 생김
#  시트를 고치거나 추가할 때: [1부]에서 시트를 만들고, [2부]의 FILE_A / FILE_B 표에 한 줄 추가
# =====================================================================
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

F="Arial"
H_FILL=PatternFill("solid",fgColor="1F3864")
H_FONT=Font(name=F,bold=True,color="FFFFFF")
B=Font(name=F); BB=Font(name=F,bold=True)
thin=Side(style="thin",color="BFBFBF"); BD=Border(left=thin,right=thin,top=thin,bottom=thin)
CHK=PatternFill("solid",fgColor="FFF2CC")

ios = """1|Meowdoku!|Oakever Games|퍼즐|로직·두뇌 트릭|스도쿠류
2|Block Out! - Color Sort Puzzle|Grand Games|퍼즐|정렬·잼|
3|Rotate Rings|Nebula Studio|퍼즐|정렬·잼|
4|Magic Sort!|Grand Games|퍼즐|정렬·잼|
5|Colony Flow!|ABI Global|퍼즐|정렬·잼|확인 필요
6|Smash Fest!|Flow Games|하이퍼·하이브리드 캐주얼|물리 파괴|
7|CubeAway - 3DPuzzle|GOODROID|퍼즐|탭 탈출|
8|Roblox|Roblox Corporation|기타(UGC·리워드)|UGC 플랫폼|
9|Vita Mahjong|Vita Studio|퍼즐|타일·마작·트리플 매치|
10|Brain Puzzle 3: Crazy Mind|SmallStep Dev Team|퍼즐|로직·두뇌 트릭|
11|Woodoku Blast: Block Puzzle|Tripledot Studios|퍼즐|블록|
12|Bus Fever Party!|Lumi Games|퍼즐|정렬·잼|버스잼
13|Block Blast!|Hungry Studio|퍼즐|블록|
14|Bus Traffic Fever!|GOODROID|퍼즐|정렬·잼|버스잼
15|Amaze GO!|Oakever Games|퍼즐|탭 탈출|
16|Colorever: Gem Art Puzzle|Oakever Games|퍼즐|픽셀·컬러 아트|
17|Township|Playrix|시뮬레이션|농장·타운 빌딩|
18|Fortnite|Epic Games|액션·슈팅|배틀로얄|
19|Royal Match|Dream Games|퍼즐|매치3|
20|Mahjong Blast|Nebula Studio|퍼즐|타일·마작·트리플 매치|
21|Royal Smash! - Physics Puzzle|Cypher Games|하이퍼·하이브리드 캐주얼|물리 파괴|
22|MONOPOLY GO!|Scopely|카드·보드·카지노|보드·주사위 소셜캐주얼|
23|Whiteout Survival|Century Games|전략|4X·SLG|
24|Gossip Harbor: Merge & Story|Microfun|퍼즐|머지|스토리 결합
25|Solitaire Associations Journey|Hitapps Games|퍼즐|워드|연상 퍼즐
26|Magic Tiles 3: Piano Game|Amanotes|파티·음악|리듬|
27|CrownCards - Trading Cards|CrownCards|카드·보드·카지노|트레이딩 카드|확인 필요
28|Food Hunt: Pixel Puzzle|Funfinity|퍼즐|픽셀·컬러 아트|
29|Tasty Travels: Merge Game|Century Games|퍼즐|머지|
30|Subway Surfers|SYBO Games|하이퍼·하이브리드 캐주얼|러너|
31|Castle Busters|Voodoo|하이퍼·하이브리드 캐주얼|하이브리드 캐주얼|
32|Sled Surfers|CrazyLabs|하이퍼·하이브리드 캐주얼|러너|
33|Tang Luck: Casino Slots|SpinShift|카드·보드·카지노|소셜 카지노|
34|Triumph Arcade|Triumph Arcade|카드·보드·카지노|리얼머니 스킬게임|
35|Royal Kingdom|Dream Games|퍼즐|매치3|
36|Arrows – Puzzle Escape|Lessmore|퍼즐|탭 탈출|
37|Car Evolve|Heseri Games|하이퍼·하이브리드 캐주얼|러너|진화 러너
38|Disney Solitaire|SuperPlay|카드·보드·카지노|솔리테어|IP 활용
39|PlayCash Rewards|확인 불가|기타(UGC·리워드)|리워드 앱|퍼블리셔 확인 필요
40|Kingshot|Century Games|전략|4X·SLG|
41|Warline: Sniper Strike|Powerfantasy|액션·슈팅|스나이퍼|스토어 분류는 Strategy
42|Paper.io 2|Voodoo|하이퍼·하이브리드 캐주얼|io|
43|Candy Crush Saga|King|퍼즐|매치3|
44|Call of Duty: Mobile|Activision|액션·슈팅|FPS|
45|Jackpot Nights|Lucky Link|카드·보드·카지노|소셜 카지노|
46|NoomiClone|Matthew Jones|하이퍼·하이브리드 캐주얼|기타|확인 필요
47|Word Search Explorer|PlaySimple|퍼즐|워드|
48|Brain Puzzle 2: Logic Twist|SmallStep Dev Team|퍼즐|로직·두뇌 트릭|
49|Sudoku.com - Number Games|Easybrain|퍼즐|로직·두뇌 트릭|스도쿠
50|Chess.com - Play and Learn|Chess.com|카드·보드·카지노|클래식 보드|체스
51|Among Us!|InnerSloth|파티·음악|소셜 디덕션|
52|8 Ball Pool|Miniclip|스포츠|당구 PvP|
53|Mahjong Rumble: Win Real Cash|Aviagames|카드·보드·카지노|리얼머니 스킬게임|
54|Mob Control|Voodoo|하이퍼·하이브리드 캐주얼|군중 러너·디펜스|
55|Songless: Guess the Song!|Calico Studio|파티·음악|음악 퀴즈|
56|Yarn Loop: Knit Puzzle|Combo Games|퍼즐|정렬·잼|
57|Match Factory!|Peak Games|퍼즐|타일·마작·트리플 매치|3D 트리플 매치
58|Jigsaw Drop: Solitaire Puzzle|Runyou Game|퍼즐|직소·솔리테어 하이브리드|
59|Pixel Flow!|Loom Games|퍼즐|픽셀·컬러 아트|
60|Hotel Legacy: Merge Game|Century Games|퍼즐|머지|
61|Fish Sort Puzzle - Bubble Jam|Shycheese|퍼즐|정렬·잼|
62|Block Crush!|Wonderful Studio|퍼즐|블록|
63|Sand Blocks: Drop Puzzle|Rollic Games|퍼즐|블록|샌드 블록
64|EA SPORTS FC Soccer Mobile 27|Electronic Arts|스포츠|축구|
65|Solitaire Cash: Win Real Money|Papaya Gaming|카드·보드·카지노|리얼머니 스킬게임|
66|Brawl Stars|Supercell|액션·슈팅|팀 대전 슈터|
67|Coin Master|Moon Active|카드·보드·카지노|보드·주사위 소셜캐주얼|
68|82-0.com|Vaulty Studios|스포츠|기타|확인 필요
69|Z Route: Redemption|Building-Blocks Network|전략|4X·SLG|좀비 테마
70|Glow Fashion Idol|CrazyLabs|시뮬레이션|패션·꾸미기|
71|Bloons Blitz|Ninja Kiwi|하이퍼·하이브리드 캐주얼|아케이드|확인 필요
72|Hole Stars: Puzzle Game|Moon Active|퍼즐|기타 퍼즐|홀 수집
73|Dice Dreams|SuperPlay|카드·보드·카지노|보드·주사위 소셜캐주얼|
74|Word Tiles - Relaxing Puzzle|Agave Games|퍼즐|워드|
75|Jam Bus: Sandfall|Nova Games|퍼즐|정렬·잼|버스잼
76|Pokémon GO|Scopely|RPG|위치기반 수집|
77|Toon Blast|Peak Games|퍼즐|매치3|블라스트
78|NYT Games: Wordle & Crossword|The New York Times|퍼즐|워드|데일리 퍼즐
79|Grand Fortune Club|EverPrime Futures|카드·보드·카지노|소셜 카지노|
80|Clash Royale|Supercell|전략|실시간 카드 배틀|
81|Jigsawcard Solitaire Puzzle|Oakever Games|퍼즐|직소·솔리테어 하이브리드|
82|Playtest Pro by BestPlay|Bestplay Systems|기타(UGC·리워드)|리워드 앱|
83|Mahjong Clash: Win Real Cash|Aviagames|카드·보드·카지노|리얼머니 스킬게임|
84|Find Hidden Objects - Spot It!|Yolo Game Studios|퍼즐|숨은그림|
85|Double Tile!|Hungry Studio|퍼즐|타일·마작·트리플 매치|
86|Cube Land Puzzle Game|Rotatelab|퍼즐|블록|
87|Goods Puzzle: Sort Challenge|Falcon Games|퍼즐|정렬·잼|3매치 정렬
88|Cider Casino - Real Money|Mystic Mirror Studio|카드·보드·카지노|리얼머니 스킬게임|리얼머니 카지노
89|Hole.io|Voodoo|하이퍼·하이브리드 캐주얼|io|
90|Color Slide - Hexa Puzzle|SayGames|퍼즐|정렬·잼|헥사
91|Geometry Dash Lite|RobTop Games|하이퍼·하이브리드 캐주얼|아케이드|플랫포머
92|Bingo Cash - Win Real Money|Papaya Gaming|카드·보드·카지노|리얼머니 스킬게임|
93|EA College Football Mobile 27|Electronic Arts|스포츠|미식축구|
94|Victory Blast Casino|Luma Field|카드·보드·카지노|소셜 카지노|
95|Retro Bowl|New Star Games|스포츠|미식축구|
96|Aniimo|Pawprint|RPG|수집형 오픈월드|
97|Foodstars: Merge & Cook|Happibits|퍼즐|머지|요리 테마
98|Wordscapes - Word Game|PeopleFun|퍼즐|워드|
99|Imposter Game - Party Edition|확인 불가|파티·음악|소셜 디덕션|퍼블리셔 확인 필요
100|Offline Games - No Wifi Games|Moreno Maio|퍼즐|기타 퍼즐|미니게임 모음"""

gp = """1|Rotate Rings|Nebula Studio|퍼즐|정렬·잼|
2|Meowdoku: Brain Puzzle Games|Oakever Games|퍼즐|로직·두뇌 트릭|스도쿠류
3|Angry Birds 2|Rovio|하이퍼·하이브리드 캐주얼|물리 파괴|
4|Roblox|Roblox Corporation|기타(UGC·리워드)|UGC 플랫폼|
5|Vita Mahjong|Vita Studio|퍼즐|타일·마작·트리플 매치|
6|Smash Land!|Flow Games|하이퍼·하이브리드 캐주얼|물리 파괴|
7|Block Blast!|Hungry Studio|퍼즐|블록|
8|Bus Fever Party!|Lumi Games|퍼즐|정렬·잼|버스잼
9|Royal Smash! - Physics Puzzle|Cypher Games|하이퍼·하이브리드 캐주얼|물리 파괴|
10|MemeMix Challenge: Funny Party|Amobear VN|파티·음악|밈 챌린지|확인 필요
11|Magic Sort!|Grand Games|퍼즐|정렬·잼|
12|Brain Puzzle 3: Crazy Mind|JoyGame Studio|퍼즐|로직·두뇌 트릭|
13|Woodoku Blast|Tripledot Studios|퍼즐|블록|
14|Amaze GO!|Oakever Games|퍼즐|탭 탈출|
15|Block Craft 3D: Building Game|Wildlife Studios|시뮬레이션|샌드박스|건설
16|Word Search Explorer|PlaySimple|퍼즐|워드|
17|Gossip Harbor: Merge & Story|Microfun|퍼즐|머지|스토리 결합
18|Fish Sort Puzzle|Shycheese|퍼즐|정렬·잼|
19|Mahjong Blast|Nebula Studio|퍼즐|타일·마작·트리플 매치|
20|Pocket Sort: Coin Merge Puzzle|SayGames|퍼즐|정렬·잼|머지 요소
21|Colony Flow!|ABI Games|퍼즐|정렬·잼|확인 필요
22|Township|Playrix|시뮬레이션|농장·타운 빌딩|
23|Arrow Puzzle: Tap Puzzle Games|Easybrain|퍼즐|탭 탈출|
24|Loop Sort|Voodoo|퍼즐|정렬·잼|
25|Fortnite|Epic Games|액션·슈팅|배틀로얄|
26|MONOPOLY GO!|Scopely|카드·보드·카지노|보드·주사위 소셜캐주얼|
27|Cube Land Puzzle Game|Rotatelab|퍼즐|블록|
28|Castle Busters|Voodoo|하이퍼·하이브리드 캐주얼|하이브리드 캐주얼|
29|Bulldozer Master: Crush & Mine|Supercent|하이퍼·하이브리드 캐주얼|물리 파괴|파괴·채굴
30|Mob Control|Voodoo|하이퍼·하이브리드 캐주얼|군중 러너·디펜스|
31|CubeAway - 3DPuzzle|GOODROID|퍼즐|탭 탈출|
32|Paper.io 2|Voodoo|하이퍼·하이브리드 캐주얼|io|
33|Magic Tiles 3 - Piano Game|Amanotes|파티·음악|리듬|
34|Arrows – Puzzle Escape|Lessmore|퍼즐|탭 탈출|
35|Color Block Jam|Rollic Games|퍼즐|정렬·잼|
36|8 Ball Pool|Miniclip|스포츠|당구 PvP|
37|Sled Surfers|CrazyLabs|하이퍼·하이브리드 캐주얼|러너|
38|Jigsawcard Solitaire Puzzle|Oakever Games|퍼즐|직소·솔리테어 하이브리드|
39|Color by Number: Coloring Games|Wildlife Studios|퍼즐|픽셀·컬러 아트|
40|Yarn Loop|Combo Game|퍼즐|정렬·잼|
41|Warline: Sniper Strike|Powerfantasy|액션·슈팅|스나이퍼|
42|Royal Match|Dream Games|퍼즐|매치3|
43|Color Block: Combo Blast|Ivy|퍼즐|블록|
44|Brainy Escape Quest|Metagame|퍼즐|로직·두뇌 트릭|
45|CarLoop|UUTEAM|퍼즐|정렬·잼|확인 필요
46|Solitaire - Classic Card Games|Guru Puzzle Game|카드·보드·카지노|솔리테어|
47|Royal Kingdom|Dream Games|퍼즐|매치3|
48|Candy Crush Saga|King|퍼즐|매치3|
49|Pixel Flow!|Loom Games|퍼즐|픽셀·컬러 아트|
50|Slow Mo Strike: Knives in Time|KAYAC|하이퍼·하이브리드 캐주얼|아케이드|액션
51|Idle Cat Gunner: Shooter RPG|Neptune(Treeplla)|RPG|방치형 슈터|
52|Subway Surfers|SYBO Games|하이퍼·하이브리드 캐주얼|러너|
53|Word Tiles - Relaxing Puzzle|Agave Games|퍼즐|워드|
54|Geometry Dash Lite|RobTop Games|하이퍼·하이브리드 캐주얼|아케이드|플랫포머
55|Mahjong Master: Daily Match|Mindful Daily Puzzles|퍼즐|타일·마작·트리플 매치|
56|Search It - Hidden Objects|Leap Game Studios|퍼즐|숨은그림|
57|Happy Color: Color by Number|X-Flow|퍼즐|픽셀·컬러 아트|
58|Apex Sniper: Animal Hunt|CozyCraft|액션·슈팅|스나이퍼|헌팅
59|Toca Boca World|Toca Boca|시뮬레이션|샌드박스|키즈 라이프
60|Among Us|Innersloth|파티·음악|소셜 디덕션|
61|Ball Sort Puzzle - Color Game|Guru Puzzle Game|퍼즐|정렬·잼|
62|Disney Solitaire|SuperPlay|카드·보드·카지노|솔리테어|IP 활용
63|Whiteout Survival|Century Games|전략|4X·SLG|
64|EA SPORTS FC Soccer Mobile 27|Electronic Arts|스포츠|축구|
65|Hidden Object Games: Seek It|NimbleMind|퍼즐|숨은그림|
66|Aniimo|Pawprint Studio|RPG|수집형 오픈월드|
67|Z Route: Redemption|37GAMES|전략|4X·SLG|좀비 테마
68|Pokémon TCG Pocket|The Pokémon Company|카드·보드·카지노|트레이딩 카드|디지털 TCG
69|Tiki Smash!|PeopleFun|하이퍼·하이브리드 캐주얼|물리 파괴|
70|Block Crush!|Wonderful Studio|퍼즐|블록|
71|Word Connect Association|PlaySimple|퍼즐|워드|연상 퍼즐
72|Melon Sandbox|playducky.com|시뮬레이션|샌드박스|물리 샌드박스
73|Hole Stars: Puzzle Game|Moon Active|퍼즐|기타 퍼즐|홀 수집
74|Double Tile!|Hungry Studio|퍼즐|타일·마작·트리플 매치|
75|Brain Puzzle 2: Logic Twist|JoyGame Studio|퍼즐|로직·두뇌 트릭|
76|Frost Valley: Merge & Story|Tuyoo Games|퍼즐|머지|스토리 결합
77|Solitaire|Mouse Games|카드·보드·카지노|솔리테어|
78|Food Hunt: Pixel Puzzle|EVERFUN|퍼즐|픽셀·컬러 아트|
79|Angry Smash|Tripledot Studios|하이퍼·하이브리드 캐주얼|물리 파괴|
80|확인 불가|-|확인 불가|-|출처 페이지에서 해당 순위 누락
81|Screw Out 3D: Nut Sort Jam|Unico Studio|퍼즐|정렬·잼|
82|Crossword Go!|PlaySimple|퍼즐|워드|
83|Match Factory!|Peak|퍼즐|타일·마작·트리플 매치|3D 트리플 매치
84|Treasure Master|Gimica|하이퍼·하이브리드 캐주얼|아케이드|채굴
85|Car Evolve|Heseri Games|하이퍼·하이브리드 캐주얼|러너|진화 러너
86|Offline Games - No Wifi Games|JindoBlu|퍼즐|기타 퍼즐|미니게임 모음
87|Lordrush|Century Games|전략|4X·SLG|확인 필요
88|Loop Master: Color Jam Sort|Pleasure City|퍼즐|정렬·잼|
89|Freddy Playground Sandbox|Big Host Africa|시뮬레이션|샌드박스|
90|Wordscapes Search: Word Games|PeopleFun|퍼즐|워드|
91|Last Extract|Voodoo|액션·슈팅|익스트랙션 슈터|
92|Vigor Mahjong|CanaryDroid|퍼즐|타일·마작·트리플 매치|
93|Solar Smash|Paradyme Games|시뮬레이션|샌드박스|파괴 샌드박스
94|GOGO! Blast|Nebula Studio|퍼즐|탭 탈출|
95|Annoying Uncle Punch Game|Game District|하이퍼·하이브리드 캐주얼|기타|스트레스 해소
96|Food Sort: Puzzle Game|Playdayy|퍼즐|정렬·잼|
97|Free Fire x NARUTO SHIPPUDEN|Garena|액션·슈팅|배틀로얄|
98|Dice Dreams|SuperPlay|카드·보드·카지노|보드·주사위 소셜캐주얼|
99|Fast Cash Farkle|GreatPlay|카드·보드·카지노|리얼머니 스킬게임|주사위
100|Bloons Blitz|Ninja Kiwi|하이퍼·하이브리드 캐주얼|아케이드|확인 필요"""

def parse(s): return [l.split("|") for l in s.strip().split("\n")]
ios=parse(ios); gp=parse(gp)
assert len(ios)==100 and len(gp)==100

wb=Workbook()

# Guide sheet
g=wb.active; g.title="개요"
rows=[("모바일 게임 시장 분석 — 장르 분류 · 지역 비교 · 성공 비결 & MVP · 벤치마크",),
("",),
("항목","내용"),
("분석 범위","미국 App Store / Google Play 게임 카테고리 Top Free 1~100위"),
("App Store 기준일","2026-10-03"),
("App Store 출처","GameDropDaily (Apple 공개 iTunes RSS 차트 기반) https://gamedropdaily.com/mobile/ · 퍼블리셔·국가 교차 확인: AppBrain https://www.appbrain.com/stats/appstore-rankings"),
("Google Play 기준일","2026-10-04"),
("Google Play 출처","AppBrain 미국 Top Free Games https://www.appbrain.com/stats/google-play-rankings/top_free/game/us · 교차 확인: GameDropDaily https://gamedropdaily.com/android/"),
("참고","수집 시점 차이로 출처 간 순위가 1~5계단 다를 수 있음. Google Play 80위는 출처 페이지에서 누락되어 '확인 불가'로 표시."),
("장르 분류","게임명, 스토어 카테고리, 공개 게임 정보 기반의 분석자 판단. 확신이 낮은 항목은 '비고'에 '확인 필요'로 표시하고 노란색으로 강조."),
("수정 방법","App Store / Google Play 시트의 '대분류', '세부 장르' 열을 수정하면 '장르 분포' 시트 집계가 자동으로 갱신됨(COUNTIF 수식)."),
("시트 구성","[미국] App Store Top100 · Google Play Top100 · 장르 분포 · 세부 장르 · 인사이트 · 심층 분석 후보 / [지역 비교] 지역 비교 요약 · 국가별 Top20 · 국가별 장르 분포 · 공통 인기 게임 · 지역별 특징 / [2단계] 점수 기준 · 성공 점수표 · 성적표 카드 · 성공 비결 해부 · MVP 명세 · 성공 공식 TOP5 · 매출 순위 원자료 / [3단계] 벤치마크 근거 · 지역 벤치마크 · MVP 검증 목표 · 제외한 자료 / [4단계] 성공 비결 확장 · 공통점 분석 / [5단계] 장르 차트 판독값 · 장르 판독 요약 · 장르 연결(추정)"),
("3단계 출처","GameAnalytics 2026·2025 벤치마크, AppsFlyer State of App Monetization 2026·Gaming App Marketing 2024, Appodeal Mobile Casual Benchmarks 2025. 출처 없는 수치는 '제외한 자료'에 기록."),
("2단계 출처","AppBrain 구글 플레이 매출 순위: 미국·한국 2026-10-03, 일본 2026-10-04, 브라질 2026-10-01. 외부 자료: Wikipedia(Free Fire). 별점·MVP 규모·인원·기간·게임 방식 일부는 분석자 추정."),
("지역 비교 출처","AppBrain 국가별 Google Play Top Free Games: 캐나다·브라질·일본 2026-10-04, 한국 2026-10-03, 멕시코 2026-09-25. 외부 자료: Wikipedia(Free Fire), Sensor Tower & Adjust 일본 시장 리포트(2024)"),
("지역 비교 한계","캐나다·남미·아시아는 구글 플레이만 조사(App Store 미포함). 국가별 장르 분류는 Top 20 기준. 동남아(인도네시아·태국 등)는 미포함. 원인 설명 일부는 추정."),
]
for r in rows: g.append(r)
g["A1"].font=Font(name=F,bold=True,size=14)
for c in g[3]: c.fill=H_FILL; c.font=H_FONT
for row in g.iter_rows(min_row=4,max_row=len(rows)):
    row[0].font=BB; row[1].font=B; row[1].alignment=Alignment(wrap_text=True,vertical="top"); row[0].alignment=Alignment(vertical="top")
g.column_dimensions["A"].width=20; g.column_dimensions["B"].width=110

def chart_sheet(name,data,tname):
    ws=wb.create_sheet(name)
    hdr=["순위","게임명","퍼블리셔","대분류","세부 장르","비고"]
    ws.append(hdr)
    for r in data:
        ws.append([int(r[0])]+r[1:])
    for c in ws[1]: c.fill=H_FILL; c.font=H_FONT; c.alignment=Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=2):
        for c in row: c.font=B; c.border=BD
        row[0].alignment=Alignment(horizontal="center")
        if "확인" in str(row[5].value) or "확인 불가" in str(row[1].value) or "확인 불가" in str(row[2].value):
            for c in row: c.fill=CHK
    for col,w in zip("ABCDEF",[7,38,26,24,26,30]): ws.column_dimensions[col].width=w
    ws.freeze_panes="A2"
    ws.auto_filter.ref=f"A1:F{ws.max_row}"
    return ws
chart_sheet("App Store Top100",ios,"ios")
chart_sheet("Google Play Top100",gp,"gp")

# Genre distribution with formulas
majors=["퍼즐","하이퍼·하이브리드 캐주얼","카드·보드·카지노","시뮬레이션","전략","액션·슈팅","스포츠","RPG","파티·음악","기타(UGC·리워드)","확인 불가"]
reps={"퍼즐":"Meowdoku, Block Blast, Royal Match","하이퍼·하이브리드 캐주얼":"Royal Smash, Mob Control, Subway Surfers","카드·보드·카지노":"MONOPOLY GO!, Solitaire Cash, Disney Solitaire","시뮬레이션":"Township, Melon Sandbox, Block Craft 3D","전략":"Whiteout Survival, Kingshot, Clash Royale","액션·슈팅":"Fortnite, Call of Duty: Mobile, Free Fire","스포츠":"EA FC Mobile 27, 8 Ball Pool, Retro Bowl","RPG":"Pokémon GO, Aniimo, Idle Cat Gunner","파티·음악":"Magic Tiles 3, Among Us","기타(UGC·리워드)":"Roblox, PlayCash Rewards","확인 불가":"-"}
ws=wb.create_sheet("장르 분포")
ws.append(["대분류","App Store 수","App Store 비율","Google Play 수","Google Play 비율","합계","대표 게임"])
for i,m in enumerate(majors,start=2):
    ws.append([m,f"=COUNTIF('App Store Top100'!$D$2:$D$101,A{i})",f"=B{i}/$B${len(majors)+2}",
               f"=COUNTIF('Google Play Top100'!$D$2:$D$101,A{i})",f"=D{i}/$D${len(majors)+2}",f"=B{i}+D{i}",reps[m]])
t=len(majors)+2
ws.append(["합계",f"=SUM(B2:B{t-1})",f"=SUM(C2:C{t-1})",f"=SUM(D2:D{t-1})",f"=SUM(E2:E{t-1})",f"=SUM(F2:F{t-1})",""])
for c in ws[1]: c.fill=H_FILL; c.font=H_FONT; c.alignment=Alignment(horizontal="center",wrap_text=True)
for row in ws.iter_rows(min_row=2):
    for c in row: c.font=B; c.border=BD
    row[2].number_format="0.0%"; row[4].number_format="0.0%"
for c in ws[t]: c.font=BB
for col,w in zip("ABCDEFG",[26,13,14,15,16,10,44]): ws.column_dimensions[col].width=w
ws.freeze_panes="A2"
from openpyxl.chart import BarChart, Reference
ch=BarChart(); ch.type="bar"; ch.title="대분류별 게임 수 (미국 Top Free 100)"; ch.y_axis.title="게임 수"
ch.add_data(Reference(ws,min_col=2,min_row=1,max_row=t-2),titles_from_data=True)
ch.add_data(Reference(ws,min_col=4,min_row=1,max_row=t-2),titles_from_data=True)
ch.set_categories(Reference(ws,min_col=1,min_row=2,max_row=t-2))
ch.height=10; ch.width=20
ws.add_chart(ch,"I2")

# Sub-genre sheet
subs=sorted(set((r[3],r[4]) for r in ios+gp if r[3]!="확인 불가"))
ws=wb.create_sheet("세부 장르")
ws.append(["대분류","세부 장르","App Store 수","Google Play 수","합계"])
# sort by Python-computed total for ordering only; values via formula
def cnt(d,m,s): return sum(1 for r in d if r[3]==m and r[4]==s)
subs.sort(key=lambda x:-(cnt(ios,*x)+cnt(gp,*x)))
for i,(m,s) in enumerate(subs,start=2):
    ws.append([m,s,f"=COUNTIFS('App Store Top100'!$D$2:$D$101,A{i},'App Store Top100'!$E$2:$E$101,B{i})",
               f"=COUNTIFS('Google Play Top100'!$D$2:$D$101,A{i},'Google Play Top100'!$E$2:$E$101,B{i})",f"=C{i}+D{i}"])
n=ws.max_row
ws.append(["합계","",f"=SUM(C2:C{n})",f"=SUM(D2:D{n})",f"=SUM(E2:E{n})"])
for c in ws[1]: c.fill=H_FILL; c.font=H_FONT; c.alignment=Alignment(horizontal="center")
for row in ws.iter_rows(min_row=2):
    for c in row: c.font=B; c.border=BD
for c in ws[ws.max_row]: c.font=BB
for col,w in zip("ABCDE",[26,28,14,15,10]): ws.column_dimensions[col].width=w
ws.freeze_panes="A2"; ws.auto_filter.ref=f"A1:E{n}"

# Insights
ins=[("1. 미국 무료 차트는 사실상 퍼즐 시장","두 스토어 모두 절반 안팎이 퍼즐. App Store 1~20위 중 16개가 퍼즐. 매치3보다 정렬·잼(버스잼, 실·물고기·링 정렬)과 탭 탈출(Arrows류)이 압도적이며, 여러 회사가 같은 메커닉을 동시에 출시하는 패스트팔로우 경쟁이 치열함."),
("2. App Store에만 리얼머니·카지노가 두드러짐","카드·보드·카지노가 App Store 16개 vs Google Play 7개. 리얼머니 스킬게임은 App Store 6개 vs Google Play 1개. Google Play의 리얼머니 게임 정책과 iOS 고결제 유저 대상 UA 집중이 원인으로 추정(검증 필요)."),
("3. Google Play는 더 어리고 가벼운 취향","하이퍼캐주얼(17 vs 12)과 샌드박스 시뮬레이션(Melon Sandbox, Block Craft 3D, Toca Boca World, Solar Smash)이 Google Play에 더 많음. 안드로이드 유저층이 상대적으로 어리다는 추정과 일치."),
("4. 두 스토어 간 중복이 큼","약 54개 게임이 양쪽 Top 100에 동시 등장. 동일 퍼블리셔가 양 플랫폼에 동시 UA하는 구조로, 이후 단계는 게임별 전략 중심 분석이 효율적."),
("5. 상위권은 신작, 중위권은 장수작","1~20위는 앱 ID로 보아 2025~2026년 출시로 추정되는 신작 위주(Meowdoku, Rotate Rings, CubeAway 등). Candy Crush, Subway Surfers, Royal Match 등 장수작은 20~50위권 유지. 신작 초기 대규모 UA vs 장수작 안정 유입 구조."),
("6. 미드코어는 소수지만 4X는 꾸준","전략·액션·RPG 합계 약 10%. Century Games(Whiteout Survival, Kingshot, Lordrush) 등 4X는 공격적 UA로 무료 차트에도 등장. BM 분석 시 Top Grossing 차트 병행 필요."),
("7. 주요 퍼블리셔·튀르키예 스튜디오 강세","다수 진입: Oakever Games, Voodoo, Century Games, Nebula Studio, PlaySimple, Dream Games, Hungry Studio, SuperPlay, Papaya·Aviagames. Dream Games, Peak, Grand Games, Flow Games, Loom Games, Rotatelab, Falcon 등 튀르키예 스튜디오가 퍼즐·캐주얼 상위권 다수 점유."),
]
ws=wb.create_sheet("인사이트")
ws.append(["인사이트","내용"])
for r in ins: ws.append(r)
for c in ws[1]: c.fill=H_FILL; c.font=H_FONT
for row in ws.iter_rows(min_row=2):
    row[0].font=BB; row[1].font=B
    for c in row: c.alignment=Alignment(wrap_text=True,vertical="top"); c.border=BD
ws.column_dimensions["A"].width=36; ws.column_dimensions["B"].width=100

cand=[("Meowdoku!","퍼즐 / 로직","양쪽 스토어 1~2위의 신작 로직 퍼즐, 신작 UA 성공 사례"),
("Rotate Rings / Arrows – Puzzle Escape","퍼즐 / 정렬·잼, 탭 탈출","현재 최대 트렌드 메커닉의 대표작"),
("Block Blast!","퍼즐 / 블록","장수 히트 블록 퍼즐, 광고 BM 중심 모델의 표본"),
("Royal Match / Royal Kingdom","퍼즐 / 매치3","매치3 프리미엄 라이브옵스와 IAP의 정석"),
("Gossip Harbor","퍼즐 / 머지","머지+스토리 결합 성공 공식, 미국 여성 유저 타깃"),
("MONOPOLY GO! / Dice Dreams","카드·보드·카지노 / 보드·주사위 소셜캐주얼","IP·주사위 소셜캐주얼, 이벤트 기반 고매출 BM"),
("Whiteout Survival / Kingshot","전략 / 4X·SLG","4X UA 전략(광고 소재와 실제 게임의 괴리)과 고과금 구조"),
("Royal Smash! / Smash Fest!","하이퍼·하이브리드 캐주얼 / 물리 파괴","하이브리드 캐주얼 물리 파괴 장르의 급부상"),
("Solitaire Cash / Bingo Cash","카드·보드·카지노 / 리얼머니 스킬게임","미국 특화 리얼머니 스킬게임 모델"),
("Vita Mahjong","퍼즐 / 타일·마작","시니어 타깃 퍼즐로 양쪽 상위권 유지"),]
ws=wb.create_sheet("심층 분석 후보")
ws.append(["후보","장르","선정 이유"])
for r in cand: ws.append(r)
for c in ws[1]: c.fill=H_FILL; c.font=H_FONT
for row in ws.iter_rows(min_row=2):
    for c in row: c.font=B; c.border=BD; c.alignment=Alignment(wrap_text=True,vertical="top")
for col,w in zip("ABC",[36,38,70]): ws.column_dimensions[col].width=w


# ===================== 지역 비교 시트 =====================
YEL=PatternFill("solid",fgColor="FFF2CC")
def hdr(ws,row=1):
    for c in ws[row]: c.fill=H_FILL; c.font=H_FONT; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
def body(ws,start=2,wrap=True):
    for row in ws.iter_rows(min_row=start):
        for c in row:
            c.font=B; c.border=BD
            c.alignment=Alignment(wrap_text=wrap,vertical="top")

# 1) 지역 비교 요약
ws=wb.create_sheet("지역 비교 요약")
rows=[("항목","미국·캐나다","남미(브라질·멕시코)","아시아(일본·한국)"),
("한 줄 비유","조용한 퍼즐 도서관","시끌벅적한 운동장","캐릭터 문구점"),
("가장 좋아하는 것","짧게 머리 쓰는 퍼즐(정렬·잼, 블록, 탭 탈출, 단어)","축구, 총싸움, 오토바이, 블록 쌓기, 가상 펫","캐릭터 RPG, 뽑기(가챠), 유명 IP 게임"),
("구글 플레이 1위","캐나다 Meowdoku / 미국 Rotate Rings","브라질·멕시코 모두 Roblox","일본 Meowdoku / 한국 FORTBLOX"),
("우리만의 특징","미국 iOS의 리얼머니 게임, 캐나다 온라인 카지노 앱","브라질 오토바이 '그라우' 게임 6개·축구 게임 6개, 멕시코 축구팀(Chivas) 게임·베팅 앱","일본: 만화·캐릭터 IP, 리듬 게임, 포인트 적립(포이카츠) 게임, 고양이 / 한국: '뽑기 증정' 방치형 RPG, 고스톱·포커"),
("주로 하는 사람(추정)","쉬는 시간에 하는 어른","10대·친구 그룹","캐릭터 팬·수집가"),
("휴대폰 환경","아이폰 비중 높음","가벼운 안드로이드 폰 많음(추정)","일본은 게임 다운로드의 55%가 iOS"),
("퍼즐 비중","매우 높음(미국 구글 플레이 53%)","낮음(브라질 Top 20 중 8개)","상위권은 높음, 중하위권은 IP·RPG"),
("게임 개발 시사점","말이 필요 없는 단순 퍼즐, 빠른 신작 출시와 대규모 UA","가벼운 용량, 친구와 함께하는 재미, 축구·현지 문화 결합","귀여운 캐릭터·유명 IP 협업(일본), 방치형 성장과 초반 대량 뽑기 이벤트(한국)"),
]
for r in rows: ws.append(r)
hdr(ws); body(ws)
for r in range(2,ws.max_row+1): ws.cell(r,1).font=BB
for col,w in zip("ABCD",[20,42,48,52]): ws.column_dimensions[col].width=w
ws.freeze_panes="B2"

# 2) 국가별 Top 20
top20={
"캐나다":("2026-10-04","AppBrain Google Play Top Free Games Canada https://www.appbrain.com/stats/google-play-rankings/top_free/game/ca",[
("Meowdoku: Brain Puzzle Games","Oakever Games","퍼즐"),("Rotate Rings","Nebula Studio","퍼즐"),("Roblox","Roblox Corporation","기타(UGC·리워드)"),("Magic Sort!","Grand Games","퍼즐"),("Amaze GO!","Oakever Games","퍼즐"),
("Bus Fever Party!","Lumi Games","퍼즐"),("Block Blast!","Hungry Studio","퍼즐"),("Brain Puzzle 3: Crazy Mind","JoyGame Studio","퍼즐"),("Vita Mahjong","Vita Studio","퍼즐"),("Search It - Hidden Objects","Leap Game Studios","퍼즐"),
("MemeMix Challenge: Funny Party","Amobear VN","파티·음악"),("Fish Sort Puzzle","Shycheese","퍼즐"),("Gossip Harbor: Merge & Story","Microfun","퍼즐"),("Light Ignite - Laser Puzzle","Branching Factor","퍼즐"),("Colony Flow!","ABI Games","퍼즐"),
("Mahjong Blast","Nebula Studio","퍼즐"),("Royal Smash! - Physics Puzzle","Cypher Games","하이퍼·하이브리드 캐주얼"),("Castle Busters","Voodoo","하이퍼·하이브리드 캐주얼"),("Pixel Flow!","Loom Games","퍼즐"),("Word Tiles - Relaxing Puzzle","Agave Games","퍼즐")]),
"브라질":("2026-10-04","AppBrain Google Play Top Free Games Brazil https://www.appbrain.com/stats/google-play-rankings/top_free/game/br",[
("Roblox","Roblox Corporation","기타(UGC·리워드)"),("Free Fire x NARUTO SHIPPUDEN","Garena","액션·슈팅"),("Bus Fever Party!","Lumi Games","퍼즐"),("Meowdoku: Brain Puzzle Games","Oakever Games","퍼즐"),("Kumquat Hollow Fable","Azam hussam","확인 필요"),
("Subway Surfers","SYBO Games","하이퍼·하이브리드 캐주얼"),("BLACK RUSSIA","BLACKHUB GAMES","시뮬레이션"),("Rotate Rings","Nebula Studio","퍼즐"),("Block Blast!","Hungry Studio","퍼즐"),("Football League 2026","MOBILE SOCCER","스포츠"),
("Twisty Snake Out","NivaLogic","퍼즐"),("Minecraft Trial","Mojang","시뮬레이션"),("Brain Puzzle 3: Crazy Mind","JoyGame Studio","퍼즐"),("Snake.io","Kooapps","하이퍼·하이브리드 캐주얼"),("Block World - Craft City","Patane Global","시뮬레이션"),
("Police Chief","VoyagerOne IE","전략"),("Amaze GO!","Oakever Games","퍼즐"),("Color Block: Combo Blast","Ivy","퍼즐"),("Craft World Block Building 3D","Patane Global","시뮬레이션"),("Paper.io 2","Voodoo","하이퍼·하이브리드 캐주얼")]),
"멕시코":("2026-09-25","AppBrain Google Play Top Free Games Mexico https://www.appbrain.com/stats/google-play-rankings/top_free/game/mx (조회 시점 최신이 9/25)",[
("Roblox","Roblox Corporation","기타(UGC·리워드)"),("Bus Fever Party!","Lumi Games","퍼즐"),("Free Fire x NARUTO SHIPPUDEN","Garena","액션·슈팅"),("Block Blast!","Hungry Studio","퍼즐"),("Jigsawcard Solitaire Puzzle","Oakever Games","퍼즐"),
("Chivas: Pa' La Cancha","Grupo Omnilife - Chivas","스포츠"),("Brain Puzzle 2: Logic Twist","JoyGame Studio","퍼즐"),("Subway Surfers","SYBO Games","하이퍼·하이브리드 캐주얼"),("Block Craft 3D","Wildlife Studios","시뮬레이션"),("Brain Puzzle 3: Crazy Mind","JoyGame Studio","퍼즐"),
("Mob Control","Voodoo","하이퍼·하이브리드 캐주얼"),("EA SPORTS FC Soccer Mobile 27","Electronic Arts","스포츠"),("Color Block: Combo Blast","Ivy","퍼즐"),("X-Clash","Glaciers Game","전략"),("Vita Mahjong","Vita Studio","퍼즐"),
("Football League 2026","MOBILE SOCCER","스포츠"),("2 3 4 Player Mini Games","Better World Games","파티·음악"),("I Am Cat","Estoty","시뮬레이션"),("Minecraft Trial","Mojang","시뮬레이션"),("School Party Craft","Candy Room Games","시뮬레이션")]),
"일본":("2026-10-04","AppBrain Google Play Top Free Games Japan https://www.appbrain.com/stats/google-play-rankings/top_free/game/jp",[
("Meowdoku: Brain Puzzle Games","Oakever Games","퍼즐"),("Rotate Rings","Nebula Studio","퍼즐"),("Block Blast!","Hungry Studio","퍼즐"),("Fish Sort Puzzle","Shycheese","퍼즐"),("Royal Kingdom","Dream Games","퍼즐"),
("Idle Cat Gunner: Shooter RPG","Neptune(Treeplla)","RPG"),("Colony Flow!","ABI Games","퍼즐"),("Gossip Harbor: Merge & Story","Microfun","퍼즐"),("Angry Birds 2","Rovio","하이퍼·하이브리드 캐주얼"),("Bus Fever Party!","Lumi Games","퍼즐"),
("Brain Puzzle 3: Crazy Mind","JoyGame Studio","퍼즐"),("Amaze GO!","Oakever Games","퍼즐"),("Roblox","Roblox Corporation","기타(UGC·리워드)"),("Top Lords","GAME SPARK","전략"),("Royal Smash! - Physics Puzzle","Cypher Games","하이퍼·하이브리드 캐주얼"),
("Royal Match","Dream Games","퍼즐"),("Fluffy Drop - No Timer","UniRise Studio","퍼즐"),("Color Block: Combo Blast","Ivy","퍼즐"),("Food Hunt: Pixel Puzzle","EVERFUN","퍼즐"),("ちいかわぽけっと (치이카와 포켓)","Applibot","시뮬레이션")]),
"한국":("2026-10-03","AppBrain Google Play Top Free Games South Korea https://www.appbrain.com/stats/google-play-rankings/top_free/game/kr",[
("FORTBLOX: Mobile Fortress GO!","하이퍼라이즈","확인 필요"),("후더덕 서바이벌-777뽑기 증정","Joy Nice Games","RPG"),("Pokémon GO","Scopely","RPG"),("고스톱M","Noriworks","카드·보드·카지노"),("Rotate Rings","Nebula Studio","퍼즐"),
("Sudoku","Moca","퍼즐"),("Magic Sort!","Grand Games","퍼즐"),("Block Blast!","Hungry Studio","퍼즐"),("Tile King - Triple Match","Moca","퍼즐"),("Roblox","Roblox Corporation","기타(UGC·리워드)"),
("국민 고스톱","Noriworks","카드·보드·카지노"),("Idle Cat Gunner: Shooter RPG","Neptune(Treeplla)","RPG"),("Wild Water World","HK Just Game Technology","확인 필요"),("Top Lords","GAME SPARK","전략"),("로그W: 리메인즈 오브 갓","Efun Company","RPG"),
("Colony Flow!","ABI Games","퍼즐"),("Royal Kingdom","Dream Games","퍼즐"),("Meowdoku: Brain Puzzle Games","Oakever Games","퍼즐"),("Mahjong Blast","Nebula Studio","퍼즐"),("Gossip Harbor: Merge & Story","Microfun","퍼즐")]),
}
region={"캐나다":"미국·캐나다","브라질":"남미","멕시코":"남미","일본":"아시아","한국":"아시아"}
ws=wb.create_sheet("국가별 Top20")
ws.append(["지역","국가","기준일","순위","게임명","퍼블리셔","대분류","출처"])
for k,(d,src,lst) in top20.items():
    for i,(n,p,g) in enumerate(lst,1):
        ws.append([region[k],k,d,i,n,p,g,src])
hdr(ws); body(ws,wrap=False)
for row in ws.iter_rows(min_row=2):
    row[3].alignment=Alignment(horizontal="center")
    if row[6].value=="확인 필요":
        for c in row: c.fill=YEL
for col,w in zip("ABCDEFGH",[12,9,12,7,36,26,24,60]): ws.column_dimensions[col].width=w
ws.freeze_panes="A2"; ws.auto_filter.ref=f"A1:H{ws.max_row}"
last=ws.max_row

# 3) 국가별 Top20 장르 집계 (수식)
ws=wb.create_sheet("국가별 장르 분포")
countries=list(top20.keys())
ws.append(["대분류 (Top 20 기준)"]+countries)
for i,m in enumerate(majors,start=2):
    ws.append([m]+[f"=COUNTIFS('국가별 Top20'!$B$2:$B${last},{get_column_letter(j+2)}$1,'국가별 Top20'!$G$2:$G${last},$A{i})" for j in range(len(countries))])
n=ws.max_row+1
ws.append(["확인 필요"]+[f"=COUNTIFS('국가별 Top20'!$B$2:$B${last},{get_column_letter(j+2)}$1,'국가별 Top20'!$G$2:$G${last},$A{n})" for j in range(len(countries))])
ws.append(["합계"]+[f"=SUM({get_column_letter(j+2)}2:{get_column_letter(j+2)}{n})" for j in range(len(countries))])
hdr(ws); body(ws)
for c in ws[ws.max_row]: c.font=BB
ws.column_dimensions["A"].width=28
for j in range(len(countries)): ws.column_dimensions[get_column_letter(j+2)].width=11
ws.cell(ws.max_row+2,1,"참고: 각 국가 구글 플레이 Top 20만 분류한 값입니다. 미국 전체 Top 100 분포는 '장르 분포' 시트를 보세요.").font=Font(name=F,italic=True,color="595959")
ch=BarChart(); ch.type="col"; ch.grouping="stacked"; ch.overlap=100
ch.title="국가별 Top 20 장르 구성"; ch.y_axis.title="게임 수"
ch.add_data(Reference(ws,min_col=1,max_col=len(countries)+1,min_row=2,max_row=n),from_rows=True,titles_from_data=True)
ch.set_categories(Reference(ws,min_col=2,max_col=len(countries)+1,min_row=1))
ch.height=10; ch.width=20
ws.add_chart(ch,"I2")

# 4) 공통 인기 게임 순위 매트릭스
ws=wb.create_sheet("공통 인기 게임")
ws.append(["게임명","미국 GP","캐나다","브라질","멕시코","일본","한국","등장 국가 수 (6개국 중)"])
common=[("Block Blast!",7,7,9,4,3,8),("Meowdoku",2,1,4,39,1,18),("Magic Sort!",11,4,50,33,36,7),
("Bus Fever Party!",8,6,3,2,10,38),("Roblox",4,3,1,1,13,10),("Vita Mahjong",5,9,29,15,34,26),
("Rotate Rings",1,2,8,"-",2,5),("Royal Match",42,25,"-","-",16,21)]
for i,r in enumerate(common,start=2):
    ws.append(list(r)+[f"=COUNT(B{i}:G{i})"])
hdr(ws); body(ws,wrap=False)
for row in ws.iter_rows(min_row=2):
    for c in row[1:]: c.alignment=Alignment(horizontal="center")
ws.column_dimensions["A"].width=22
for col in "BCDEFG": ws.column_dimensions[col].width=10
ws.column_dimensions["H"].width=22
nr=ws.max_row+2
notes=["숫자는 각 국가 구글 플레이 Top Free Games 순위, '-'는 Top 100 밖.",
"Rotate Rings가 멕시코에 없는 것은 멕시코 자료가 9/25 기준으로 더 오래됐기 때문일 수 있음.",
"정정: 이전 답변에서 Royal Match를 '모든 지역 공통'으로 소개했으나, 브라질·멕시코 Top 100에는 없음."]
for k,t in enumerate(notes): ws.cell(nr+k,1,t).font=Font(name=F,italic=True,color="595959")

# 5) 지역별 특징 상세
ws=wb.create_sheet("지역별 특징")
ws.append(["지역","국가","특징","근거 게임 (순위)","분류"])
det=[
("미국·캐나다","캐나다","미국과 거의 동일한 퍼즐 중심 차트","Meowdoku(1), Rotate Rings(2), Magic Sort(4), Amaze GO!(5), Block Blast(7)","데이터"),
("미국·캐나다","캐나다","온라인 카지노 앱 진입","Jackpot City: Online Casino(33)","데이터"),
("미국·캐나다","캐나다","단어 퍼즐 인기","Word Tiles(20), Crossword Go!(44), Words of Wonders(60, 74)","데이터"),
("남미","브라질","Roblox·Free Fire가 1·2위","Roblox(1), Free Fire x NARUTO(2)","데이터"),
("남미","브라질","Free Fire는 저사양 폰에서도 잘 돌아간다는 평가, 브라질 투표 인기 1위 수상 이력","Wikipedia Free Fire 문서","외부 자료"),
("남미","브라질","축구 게임 6개","Football League 2026(10), eFootball(27), EA FC 27(34), Soccer Superstar(40), Mini Soccer(60), Dream League Soccer(94)","데이터"),
("남미","브라질","오토바이 앞바퀴 들기(그라우) 게임 6개","Moto Wheelie 3D(28), Elite Auto Brazil(31), Wheelie Master(43), Wheelie Life(46), Moto Wheelie Traffic(55), Projeto Grau(100)","데이터"),
("남미","브라질","블록 쌓기·샌드박스","Minecraft Trial(12), Block World(15), Craft World(19), School Party Craft(21), Block Craft 3D(26)","데이터"),
("남미","브라질","가상 펫·키즈 게임","Pou(24), My Talking Tom 2(37), My Talking Angela 2(42), Piano Kids(32), Bluey(76), BabyBus(82)","데이터"),
("남미","브라질","온라인 롤플레이 라이프 시뮬","BLACK RUSSIA(7), Rio Rise(83)","데이터"),
("남미","멕시코","미국식 퍼즐 + 남미식 축구·Free Fire 혼합","Bus Fever Party(2), Block Blast(4), Free Fire(3), Chivas(6), EA FC 27(12)","데이터"),
("남미","멕시코","현지 축구팀 공식 게임","Chivas: Pa' La Cancha(6)","데이터"),
("남미","멕시코","리얼머니 카지노·베팅 앱","AFUN Lite(26), Afun.MX(44), Betmaster(63), betmexico.mx(71)","데이터"),
("남미","멕시코","여럿이 하는 미니게임·오프라인 게임","2 3 4 Player Mini Games(17), Offline Games FunOn(45)","데이터"),
("아시아","일본","상위권은 글로벌 퍼즐","Meowdoku(1), Rotate Rings(2), Block Blast(3), Fish Sort(4), Royal Kingdom(5)","데이터"),
("아시아","일본","만화·캐릭터 IP 게임 다수","치이카와 포켓(20), 도검난무 퍼즐(24), Pokémon TCG Pocket(30), Pokémon Sleep(41), 몬스터 스트라이크(44), 냥코 대전쟁(51), 뿌요뿌요 퀘스트(67), 디즈니 츠무츠무(71), Pokémon Champions(87)","데이터"),
("아시아","일본","리듬 게임","뱅드림 아워노츠(48), 프로젝트 세카이(85)","데이터"),
("아시아","일본","포인트 적립(포이카츠) 게임","ポイ活にゃんこのつり日和(61), ポイムのはちうえ(64), ねこ旅(74)","데이터"),
("아시아","일본","파칭코·메달 푸셔","[7R]ビッグドリーム THE GOLDEN PUSHER(42)","데이터"),
("아시아","일본","매출의 48.8%가 RPG, 게임 다운로드의 55%가 iOS, 게이머 63.6%가 남성","Sensor Tower & Adjust 일본 시장 리포트(2024)","외부 자료"),
("아시아","한국","게임명에 '뽑기 증정'을 넣는 방치형 RPG","후더덕 서바이벌-777뽑기 증정(2), 트릭컬 리바이브-600뽑기 증정(44), 토이 신선-모든 영웅 증정(59)","데이터"),
("아시아","한국","방치형·키우기 RPG 다수","아가타 키우기(22), 나만 SSS급 퇴마사(34), MapleStory Idle RPG(54), 오늘도 환생 2(84)","데이터"),
("아시아","한국","고스톱·포커","고스톱M(4), 국민 고스톱(11), 한게임 포커 클래식(86)","데이터"),
("아시아","한국","과일 합치기(수박 게임류)","Watermelon Maker 2048(57), Melon Maker(81)","데이터"),
("아시아","한국","국산 주사위 디펜스","Random Dice 2(63)","데이터"),
("아시아","한국","해외 SLG/RPG의 한국어 현지화 출시","X 클래시(66), 라스트 에코(93), 삼국지: 제후의 전쟁(42)","데이터"),
]
for r in det: ws.append(r)
hdr(ws); body(ws)
for row in ws.iter_rows(min_row=2):
    if row[4].value=="외부 자료":
        row[4].font=Font(name=F,color="7F6000",bold=True)
for col,w in zip("ABCDE",[12,9,46,80,11]): ws.column_dimensions[col].width=w
ws.freeze_panes="A2"; ws.auto_filter.ref=f"A1:E{ws.max_row}"


# ===================== 2단계: 성공 비결 & MVP =====================
BLUE=Font(name=F,color="0000FF")
def hdr2(ws,row=1):
    for c in ws[row]:
        if c.value is not None:
            c.fill=H_FILL; c.font=H_FONT; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)

# --- 점수 기준 (가정값) ---
ws=wb.create_sheet("점수 기준")
crit=[("항목","값","설명"),
("인기(무료 순위) 배점",25,"6개국(미국·캐나다·브라질·멕시코·일본·한국) 구글 플레이 무료 순위"),
("돈 버는 힘(매출 순위) 배점",25,"4개국(미국·브라질·일본·한국) 구글 플레이 매출 순위"),
("규모(누적 설치) 배점",15,"10만 설치=0점, 10억 설치 이상=만점 (로그 비례)"),
("만족도(평점) 배점",15,"평점 3.5 이하=0점, 5.0=만점. 평점 미표시는 중간값(50%)"),
("오래가는 힘 배점",10,"출시 후 5년 이상 차트 유지 시 만점(1년당 20%)"),
("따라 만들기 쉬움 배점",10,"MVP 규모 점수표(아래)를 따름"),
("최고 시장 반영 비중",0.6,"순위 점수 = 최고 시장 점수×0.6 + 전체 평균×0.4 (지역 챔피언이 손해 보지 않도록)"),
("기준 연도",2026,"오래가는 힘 계산 기준"),
("A+ 전설 기준(이상)",80,"초기 가이드의 90점은 모든 지역을 석권한 게임이 없어 80점으로 조정"),
("A 강자 기준(이상)",70,"초기 가이드의 75점에서 조정"),
("B 유망 기준(이상)",60,"그 미만은 C 도전자"),
("MVP 규모 S 점수",10,"소규모: 2~4명, 2~3개월 내외(추정)"),
("MVP 규모 M 점수",7,"중규모: 5~10명, 4~8개월 내외(추정)"),
("MVP 규모 L 점수",4,"대규모: 10~25명, 8~15개월 내외 또는 IP 필요(추정)"),
("MVP 규모 XL 점수",1,"초대형: 30명 이상, 1.5년 이상·서버/라이브 운영 필수(추정)")]
for r in crit: ws.append(r)
hdr2(ws); body(ws)
for r in range(2,ws.max_row+1):
    ws.cell(r,2).font=BLUE; ws.cell(r,2).fill=YEL
ws.cell(8,2).number_format="0%"
for col,w in zip("ABC",[28,10,80]): ws.column_dimensions[col].width=w
ws.cell(ws.max_row+2,1,"파란 글씨·노란 칸은 바꿀 수 있는 가정값입니다. 바꾸면 '성공 점수표'가 자동으로 다시 계산됩니다.").font=Font(name=F,italic=True,color="595959")
C="'점수 기준'!"
W_POP,W_MON,W_SIZE,W_SAT,W_LONG,W_EASY,W_BEST,YR,GA,GB,GC=[f"{C}$B${i}" for i in range(2,13)]

# --- 성공 점수표 ---
ws=wb.create_sheet("성공 점수표")
ws.append(["","","","무료 순위 (구글 플레이, 빈칸=Top100 밖)","","","","","","매출 순위 (구글 플레이, 빈칸=Top100 밖)","","","","기본 정보","","","","점수 (자동 계산)","","","","","","","",""])
ws.append(["게임","퍼블리셔","장르","미국","캐나다","브라질","멕시코","일본","한국","미국","브라질","일본","한국",
           "누적 설치(GP)","평점","출시연도(추정)","MVP 규모","인기 /25","돈 버는 힘 /25","규모 /15","만족도 /15","오래가는 힘 /10","따라 만들기 /10","총점 /100","등급","순위"])
games=[
("Meowdoku","Oakever Games","퍼즐 · 로직",[2,1,4,39,1,18],[None]*4,21e6,4.7,2026,"S"),
("Rotate Rings","Nebula Studio","퍼즐 · 정렬",[1,2,8,None,2,5],[None]*4,1.9e6,4.8,2026,"S"),
("Block Blast!","Hungry Studio","퍼즐 · 블록",[7,7,9,4,3,8],[None]*4,1.1e9,4.8,2022,"S"),
("Magic Sort!","Grand Games","퍼즐 · 정렬",[11,4,50,33,36,7],[56,None,None,None],40e6,4.5,2024,"S"),
("Bus Fever Party!","Lumi Games","퍼즐 · 버스잼",[8,6,3,2,10,38],[None]*4,11e6,4.3,2025,"S"),
("Royal Match","Dream Games","퍼즐 · 매치3+꾸미기",[42,25,None,None,16,21],[4,7,23,14],370e6,4.6,2021,"M"),
("Gossip Harbor","Microfun","퍼즐 · 머지+스토리",[17,13,None,25,8,20],[5,4,16,12],91e6,4.3,2022,"L"),
("MONOPOLY GO!","Scopely","보드 · 주사위 소셜",[26,21,None,None,None,None],[1,73,None,None],130e6,4.6,2023,"L"),
("Whiteout Survival","Century Games","전략 · 4X",[63,None,None,None,28,48],[10,23,9,6],100e6,4.1,2023,"XL"),
("Free Fire","Garena","액션 · 배틀로얄",[97,None,2,3,None,None],[22,1,None,None],2.2e9,4.3,2017,"XL"),
("치이카와 포켓","Applibot","캐릭터 IP",[None,None,None,None,20,39],[None,None,81,None],1.1e6,None,2025,"L"),
("후더덕 서바이벌","Joy Nice Games","방치형 RPG",[None]*5+[2],[None,None,None,20],190e3,None,2026,"M"),
]
for i,(n,p,g,fr,gr,ins,rt,yr,sz) in enumerate(games,start=3):
    row=[n,p,g]+fr+gr+[ins,rt,yr,sz]
    pop=f"={W_POP}*({W_BEST}*IF(COUNT(D{i}:I{i})=0,0,(101-MIN(D{i}:I{i}))/100)+(1-{W_BEST})*(COUNT(D{i}:I{i})*101-SUM(D{i}:I{i}))/100/6)"
    mon=f"={W_MON}*({W_BEST}*IF(COUNT(J{i}:M{i})=0,0,(101-MIN(J{i}:M{i}))/100)+(1-{W_BEST})*(COUNT(J{i}:M{i})*101-SUM(J{i}:M{i}))/100/4)"
    siz=f"={W_SIZE}*MAX(0,MIN(1,LOG10(N{i}/100000)/4))"
    sat=f"={W_SAT}*IF(O{i}=\"\",0.5,MAX(0,MIN(1,(O{i}-3.5)/1.5)))"
    lng=f"={W_LONG}*MAX(0,MIN(1,({YR}-P{i})/5))"
    esy=f"={W_EASY}*INDEX({C}$B$13:$B$16,MATCH(\"MVP 규모 \"&Q{i}&\" 점수\",{C}$A$13:$A$16,0))/10"
    tot=f"=SUM(R{i}:W{i})"
    grd=f"=IF(X{i}>={GA},\"A+ 전설\",IF(X{i}>={GB},\"A 강자\",IF(X{i}>={GC},\"B 유망\",\"C 도전자\")))"
    rnk=f"=RANK(X{i},$X$3:$X$14)"
    ws.append(row+[pop,mon,siz,sat,lng,esy,tot,grd,rnk])
for rng in ["D1:I1","J1:M1","N1:Q1","R1:Z1"]: ws.merge_cells(rng)
hdr2(ws,1); hdr2(ws,2)
for row in ws.iter_rows(min_row=3):
    for c in row: c.font=B; c.border=BD; c.alignment=Alignment(horizontal="center",vertical="center")
    row[0].alignment=Alignment(horizontal="left"); row[0].font=BB
    for c in row[3:17]: c.font=BLUE
    row[13].number_format='#,##0'
    for c in row[17:24]: c.number_format="0.0"
    row[23].font=BB
for col,w in zip("ABC",[18,16,20]): ws.column_dimensions[col].width=w
for j in range(4,27): ws.column_dimensions[get_column_letter(j)].width=9
ws.column_dimensions["N"].width=14; ws.column_dimensions["X"].width=10; ws.column_dimensions["Y"].width=11
ws.row_dimensions[2].height=32
ws.freeze_panes="B3"
nr=ws.max_row+2
notes=["출처: AppBrain 국가별 Google Play 순위 — 무료: 미국·캐나다·브라질·일본 2026-10-04, 한국 10-03, 멕시코 09-25 / 매출: 미국·한국 10-03, 일본 10-04, 브라질 10-01.",
"누적 설치·평점: AppBrain 표기값(구글 플레이). 치이카와 포켓은 일본판 기준, 평점 미표시(빈칸)는 만족도 중간값 처리.",
"출시연도는 공개 정보·앱 ID 기반 추정치이며, MVP 규모는 분석자 판단(추정)입니다. 파란 글씨는 입력값입니다."]
for k,t in enumerate(notes): ws.cell(nr+k,1,t).font=Font(name=F,italic=True,color="595959")
ch=BarChart(); ch.type="bar"; ch.title="게임별 성공 점수 (100점 만점)"; ch.style=10
ch.add_data(Reference(ws,min_col=24,min_row=2,max_row=14),titles_from_data=True)
ch.set_categories(Reference(ws,min_col=1,min_row=3,max_row=14))
ch.height=11; ch.width=18; ch.legend=None
ws.add_chart(ch,f"A{nr+4}")

# --- 성적표 카드 ---
ws=wb.create_sheet("성적표 카드")
ws.append(["게임","한 줄 설명","이렇게 생각하면 쉬워요","🎮 재미","👶 쉬움","🔁 또 오기","👫 친구","💰 용돈 벌기","종합 등급","총점","이 게임의 비밀 한 줄","재미","쉬움","또오기","친구","용돈"])
cards=[
("Meowdoku","고양이 그림으로 칸을 채우는 스도쿠 같은 논리 퍼즐 (게임명·장르 기반 추정)","고양이 스티커로 하는 숫자 없는 스도쿠",4,4,4,1,2,"어려운 스도쿠를 귀여운 고양이로 포장해 문턱을 낮췄어요. 6개국 모두 상위권(일본·캐나다 1위)."),
("Rotate Rings","고리를 돌려 같은 색끼리 모으는 정렬 퍼즐 (게임명 기반 추정)","색깔 고리를 돌리는 큐브 맞추기",4,5,3,1,2,"손가락 하나로 돌리기만 하면 '착착 정리되는' 쾌감. 미국 구글 플레이 무료 1위."),
("Block Blast!","8×8 판에 블록 조각을 놓아 줄을 지우는 퍼즐","끝나지 않는 테트리스 조각 맞추기",5,5,4,1,3,"레벨도 이야기도 없이 '한 판 더'만으로 설치 11억. 대신 매출 순위에는 없어서 광고로 돈을 버는 것으로 보여요."),
("Magic Sort!","병 속 색깔 물을 옮겨 같은 색끼리 모으는 퍼즐","색깔 주스 나눠 담기",4,5,3,1,3,"검증된 '물 정렬'을 예쁜 연출로 업그레이드. 퍼즐 신작 중 미국 매출 Top 100(56위)에 들어간 유일한 게임."),
("Bus Fever Party!","승객을 같은 색 버스에 태워 보내는 버스잼 퍼즐","색깔별로 줄 세우는 반장 놀이",4,5,3,1,2,"말이 필요 없는 규칙이라 남미에서도 2~3위. 국경이 없는 게임이에요."),
("Royal Match","같은 그림 3개를 맞춰 별을 모으고, 별로 왕의 성을 꾸미는 게임","사탕 맞추기 + 내 성 꾸미기",5,4,5,3,5,"익숙한 퍼즐에 최고급 연출과 '꾸미는 재미'를 더해 미국 매출 4위, 4개국 매출 모두 Top 25."),
("Gossip Harbor","물건을 2개씩 합쳐 주문을 해결하면 드라마 같은 이야기가 이어지는 게임","레고 합체 + 다음 화가 궁금한 드라마",5,4,5,2,5,"'합치는 손맛'과 '다음 이야기 궁금증'의 조합. 4개국 매출 모두 Top 16(브라질 4위)."),
("MONOPOLY GO!","주사위를 굴려 보드판을 돌며 내 도시를 짓는 게임","온 가족이 아는 부루마불 + 스티커 앨범 모으기",5,5,5,5,5,"누구나 아는 이름(IP) + 친구와 함께하는 이벤트로 미국 매출 1위. 하지만 일본·한국에서는 순위 밖이에요."),
("Whiteout Survival","눈보라 속 마을을 키우고 동맹과 함께 싸우는 전략 게임","눈 오는 날 우리 마을 키우기 + 반 대항전",4,3,5,5,5,"다운로드는 많지 않아도 4개국 매출 모두 Top 25(한국 6위). 열심히 하는 소수가 큰돈을 써요."),
("Free Fire","여러 명이 섬에 떨어져 마지막까지 살아남는 짧은 총싸움 게임","10분짜리 서바이벌 술래잡기",5,4,5,5,4,"성능 낮은 휴대폰에서도 잘 돌아가게 만들어 브라질 무료 2위·매출 1위."),
("치이카와 포켓","인기 캐릭터 '치이카와'가 나오는 게임 (세부 게임 방식은 확인 필요)","좋아하는 캐릭터 스티커북",3,5,4,3,4,"캐릭터 팬이 곧 홍보대사. 일본 무료 20위·매출 81위, 한국에서도 39위."),
("후더덕 서바이벌","켜두기만 해도 캐릭터가 자라는 방치형 RPG로 추정 (확인 필요)","저절로 크는 다마고치 + 뽑기",3,5,4,2,5,"게임 이름에 '777뽑기 증정'을 넣어 이름 자체가 광고. 한국 무료 2위·매출 20위."),
]
for i,c in enumerate(cards,start=2):
    n,one,met,a,b,cc,d,e,sec=c
    ws.append([n,one,met]+[f'=REPT("★",{col}{i})&REPT("☆",5-{col}{i})' for col in "LMNOP"]+
              [f"=INDEX('성공 점수표'!$Y$3:$Y$14,MATCH(A{i},'성공 점수표'!$A$3:$A$14,0))",
               f"=INDEX('성공 점수표'!$X$3:$X$14,MATCH(A{i},'성공 점수표'!$A$3:$A$14,0))",sec,a,b,cc,d,e])
hdr2(ws); body(ws)
for row in ws.iter_rows(min_row=2):
    row[0].font=BB
    for c in row[3:10]: c.alignment=Alignment(horizontal="center",vertical="center")
    for c in row[3:8]: c.font=Font(name=F,color="C00000")
    row[9].number_format="0.0"
    for c in row[11:16]: c.font=BLUE
for col,w in zip("ABCDEFGHIJK",[18,42,30,10,10,10,10,11,11,8,60]): ws.column_dimensions[col].width=w
for col in "LMNOP": ws.column_dimensions[col].width=6
ws.cell(1,12).value="재미(입력)"; ws.cell(1,13).value="쉬움(입력)"; ws.cell(1,14).value="또오기(입력)"; ws.cell(1,15).value="친구(입력)"; ws.cell(1,16).value="용돈(입력)"
ws.freeze_panes="B2"
ws.cell(ws.max_row+2,1,"별점(1~5)은 분석자 평가이며 L~P열(파란 글씨)을 바꾸면 별 표시가 자동으로 바뀝니다. 종합 등급·총점은 '성공 점수표'에서 가져옵니다.").font=Font(name=F,italic=True,color="595959")

# --- 성공 비결 해부 ---
ws=wb.create_sheet("성공 비결 해부")
ws.append(["게임","① 핵심 재미 루프","② 첫 1분 경험","③ 다시 오게 하는 장치","④ 수익 모델(BM)","⑤ 마케팅","⑥ 아트 컨셉","근거 수준"])
anat=[
("Meowdoku","칸 고르기 → 규칙에 맞는 고양이 배치 → 판 완성","작은 판으로 규칙을 자연스럽게 익히는 튜토리얼(추정)","레벨 진행, 점점 커지는 판(추정)","광고 중심 추정 — 6개국 무료 상위권이지만 매출 Top 100 진입 없음","6개국 동시 무료 상위권 → 대규모 글로벌 UA 추정. 같은 회사(Oakever)가 Amaze GO!·Jigsawcard 등 다수 동시 운영","귀여운 고양이 캐릭터, 밝고 깔끔한 화면(추정)","차트 데이터 + 장르 추정"),
("Rotate Rings","고리 회전 → 같은 색 정렬 → 한 줄 정리되는 쾌감","한 번 돌려보면 바로 이해되는 단순 조작(추정)","레벨 진행(추정)","광고 중심 추정 — 매출 Top 100 진입 없음","미국 무료 1위 등 신작 집중 UA 추정. 같은 회사(Nebula)의 Mahjong Blast·GOGO! Blast와 포트폴리오 운영","단색 고리 위주의 미니멀 그래픽(추정)","차트 데이터 + 게임명 추정"),
("Block Blast!","블록 3개 받기 → 판에 배치 → 줄 지우기·콤보","설명 없이 바로 시작 가능한 한 판 구조","최고 점수 갱신, 콤보 연출의 손맛","광고 중심(매출 Top 100 진입 없음)으로 추정","설치 11억(구글 플레이) — 오래 쌓인 입소문과 꾸준한 UA","나무·보석 느낌의 블록, 단순한 판","차트·설치 데이터 + 공개 게임 방식"),
("Magic Sort!","병 고르기 → 색 옮기기 → 한 병 완성","워터소트 규칙(누구나 아는 방식)으로 즉시 이해","레벨 진행, 난이도 곡선","광고 + 인앱결제 혼합 추정 — 미국 매출 56위","같은 회사(Grand Games)의 Block Out!과 함께 미국 상위권","반짝이는 액체·마법 연출","차트 데이터 + 공개 게임 방식"),
("Bus Fever Party!","승객 탭 → 같은 색 버스 탑승 → 버스 출발","첫 판에서 바로 이해되는 색 맞추기","레벨 진행(추정)","광고 중심 추정 — 매출 Top 100 진입 없음","미국·캐나다·브라질·멕시코·일본 Top 10 — 언어 없이 통하는 광고 소재 추정","장난감 같은 3D 버스와 승객(추정)","차트 데이터 + 장르 추정"),
("Royal Match","3개 맞추기 → 레벨 클리어로 별 획득 → 성 꾸미기","짧고 쉬운 초반 레벨로 바로 성공 경험","성 꾸미기(메타), 이벤트, 팀(공개 정보 기반)","목숨·부스터·코인 인앱결제, 시즌 패스 계열(공개 정보 기반)","왕 캐릭터를 활용한 광고 — 미국 매출 4위","고급스러운 3D 카툰, 밝은 왕국 테마","매출·차트 데이터 + 공개 게임 방식"),
("Gossip Harbor","아이템 2개 합치기 → 주문 해결 → 이야기 진행","합치기 조작과 드라마 첫 장면을 함께 보여줌(추정)","다음 화가 궁금한 스토리, 장소 복원","에너지·아이템 인앱결제(추정) — 미국 5·브라질 4·일본 16·한국 12위 매출","스토리 장면을 앞세운 광고(추정). 같은 회사의 Seaside Escape·Flambé도 매출 Top 100","드라마 같은 인물 일러스트 + 아기자기한 아이템","매출·차트 데이터 + 장르 추정"),
("MONOPOLY GO!","주사위 굴리기 → 칸 이동·보상 → 도시 짓기","누구나 아는 모노폴리 규칙","스티커 앨범 수집, 친구와 함께하는 이벤트(공개 정보 기반)","주사위·이벤트 상품 인앱결제 — 미국 매출 1위","유명 IP + 대규모 광고, 친구 초대·스티커 교환","밝은 보드게임 카툰, 미스터 모노폴리 캐릭터","매출·차트 데이터 + 공개 게임 방식"),
("Whiteout Survival","건물 짓기 → 자원 모으기 → 영웅 성장 → 동맹 전투","광고와 다른 미니게임·생존 연출로 시작하는 경우 많음(추정)","동맹, 서버 경쟁, 매일 하는 일과","영웅·가속 아이템 인앱결제 — 미국 10·브라질 23·일본 9·한국 6위 매출","같은 회사(Century Games)의 Kingshot·Tasty Travels와 대규모 UA","얼어붙은 세계의 3D 리얼 카툰","매출·차트 데이터 + 장르 일반 지식"),
("Free Fire","낙하 → 장비 줍기 → 전투 → 마지막 생존","짧은 한 판, 빠른 매칭(공개 정보 기반)","랭크 시스템, 캐릭터 수집, 친구와 스쿼드","캐릭터·스킨 인앱결제, 콜라보(예: NARUTO) — 브라질 매출 1위","저사양 폰 최적화, 브라질·태국에서 투표 인기 1위 수상 이력","캐릭터 중심 3D, 화려한 콜라보 스킨","매출·차트 데이터 + 외부 자료(Wikipedia)"),
("치이카와 포켓","캐릭터와 함께 미니게임·수집(추정, 확인 필요)","좋아하는 캐릭터를 바로 만나는 경험(추정)","캐릭터 수집·꾸미기(추정)","인앱결제(추정) — 일본 매출 81위","원작 캐릭터 팬덤","원작 그림체 그대로","차트 데이터 + 추정(게임 방식 미확인)"),
("후더덕 서바이벌","켜두기 → 자동 전투·성장 → 뽑기로 강화(추정)","초반에 대량 뽑기 지급으로 강해지는 느낌(추정)","방치 보상, 매일 뽑기(추정)","뽑기·성장 패키지 인앱결제(추정) — 한국 매출 20위","앱 이름에 '777뽑기 증정'을 넣어 스토어 자체를 광고판으로 사용. 같은 회사(Joy Nice Games)가 Legend of Mushroom·갓깨비 키우기 운영","귀여운 2D 캐릭터(추정)","차트 데이터 + 추정"),
]
for r in anat: ws.append(r)
hdr2(ws); body(ws)
for row in ws.iter_rows(min_row=2):
    row[0].font=BB
    if "추정" in str(row[7].value): row[7].fill=YEL
for col,w in zip("ABCDEFGH",[16,32,30,30,36,40,28,20]): ws.column_dimensions[col].width=w
ws.freeze_panes="B2"
ws.cell(ws.max_row+2,1,"'추정'은 차트 데이터와 게임명·장르, 일반적으로 알려진 게임 방식을 바탕으로 한 판단입니다. 직접 플레이로 확인이 필요합니다.").font=Font(name=F,italic=True,color="595959")

# --- MVP 명세 ---
ws=wb.create_sheet("MVP 명세")
ws.append(["게임","꼭 필요 (Must)","있으면 좋음 (Should)","나중에 (Later)","MVP 규모","예상 인원·기간(추정)","출시 전후 확인할 숫자(검증 지표)","만들 때 가장 어려운 점"])
mvp=[
("Meowdoku","논리 퍼즐 규칙 엔진 + 정답 검증기, 고양이 아트 세트, 레벨 100개+, 힌트·되돌리기","보상형 광고·전면 광고, 일일 퍼즐","고양이 도감 수집, 시즌 이벤트","S","2~4명 / 2~3개월","D1·D7 리텐션, 레벨 이탈 구간, 광고 시청 수","풀 수 있으면서 너무 쉽지 않은 레벨 대량 생성"),
("Rotate Rings","고리 회전 조작 + 색 정렬 판정, 레벨 100개+, 손맛 연출(사운드·진동)","광고, 힌트 아이템","테마 스킨, 이벤트 레벨","S","2~4명 / 2~3개월","D1 리텐션, 첫 10레벨 완료율, 광고 시청 수","조작감(회전·스냅)의 기분 좋은 손맛"),
("Block Blast!","8×8 판 + 블록 생성 규칙, 줄 지우기·콤보, 최고 점수","광고(이어하기 보상 등)","모드 추가(레벨형), 테마","S","2~3명 / 1~2개월","세션당 판 수, D1·D30 리텐션, 광고 수익/유저","블록이 '공정하게' 나오는 생성 규칙"),
("Magic Sort!","병·액체 정렬 규칙, 레벨 150개+, 액체 연출","광고 + 소액 인앱(힌트·빈 병)","시즌 이벤트, 꾸미기 메타","S","3~4명 / 2~3개월","레벨 이탈 구간, 광고·결제 비율","레벨 난이도 곡선 설계"),
("Bus Fever Party!","버스·승객 색 매칭 규칙, 주차 공간 제한, 레벨 100개+","광고, 부스터(대기열 늘리기 등)","이벤트·수집","S","2~4명 / 2~3개월","D1 리텐션, 레벨당 실패율","막히는 순간과 풀리는 순간의 균형"),
("Royal Match","매치3 엔진, 장애물 5~8종, 레벨 100개+, 성 꾸미기 1구역","코인 상점·부스터, 목숨 시스템","팀(길드), 경쟁 이벤트, 시즌 패스","M","5~10명 / 4~8개월","D1·D7 리텐션, 레벨당 시도 횟수, 첫 결제 전환율","매치3 '손맛'과 연출 품질 (이미 강자 다수)"),
("Gossip Harbor","머지 보드 + 아이템 체인, 주문 시스템, 스토리 1~2장, 에너지","에너지·아이템 상점, 스토리 연출 강화","시즌 이벤트, 새 지역·장 추가","L","10~20명 / 8~12개월","D7·D30 리텐션, 스토리 장 완료율, 결제 전환율","스토리·아트 콘텐츠를 계속 만들어야 하는 제작량"),
("MONOPOLY GO!","주사위 보드 루프, 도시 짓기, 친구 연동(최소)","이벤트·스티커 앨범, 주사위 상점","파트너 이벤트, 앨범 교환, 대형 IP 계약","L","15~25명 / 10~15개월 (IP 제외)","D1·D30 리텐션, 친구 초대율, 이벤트 참여율, ARPDAU","IP 없이는 같은 효과를 내기 어려움 + 실시간 이벤트 운영"),
("Whiteout Survival","도시 건설·자원 루프, 영웅 1차 세트, 서버·동맹 기본 기능","영웅 뽑기·가속 상점, 일일 퀘스트","동맹 전쟁, 시즌 서버, 대형 이벤트","XL","30명 이상 / 1.5년 이상","D30·D90 리텐션, 상위 결제자 비중, 동맹 가입률","서버·밸런스·라이브 운영 + 매우 큰 마케팅 비용"),
("Free Fire","실시간 멀티플레이 서버, 매칭, 맵 1개, 무기·캐릭터 기본 세트, 저사양 최적화","랭크, 스킨 상점","콜라보, e스포츠, 신규 모드","XL","40명 이상 / 2년 이상","동시 접속자, 매칭 대기시간, D7 리텐션","실시간 서버 + 핵·부정행위 대응 + 저사양 최적화"),
("치이카와 포켓","IP 라이선스 확보, 원작 그림체 캐릭터, 미니게임·수집 루프(추정)","꾸미기 아이템 상점","IP 이벤트, 굿즈 연동","L","10~20명 / 8~12개월 + IP 계약","팬 유입 대비 D7 리텐션, 결제 전환율","IP 확보 비용과 원작사 승인 절차"),
("후더덕 서바이벌","방치 성장 루프, 자동 전투, 뽑기·장비 시스템, 초반 대량 뽑기 지급(추정)","성장 패키지 상점, 출석 보상","길드·랭킹, 시즌 콘텐츠","M","5~10명 / 4~8개월","D1·D7 리텐션, 첫 결제 전환율, 결제자당 매출","뽑기 확률·성장 밸런스와 확률 공시 등 규제 대응"),
]
for r in mvp: ws.append(r)
hdr2(ws); body(ws)
for row in ws.iter_rows(min_row=2):
    row[0].font=BB; row[4].alignment=Alignment(horizontal="center",vertical="top"); row[4].font=BB
for col,w in zip("ABCDEFGH",[16,44,30,28,9,22,36,36]): ws.column_dimensions[col].width=w
ws.freeze_panes="B2"
ws.cell(ws.max_row+2,1,"인원·기간은 분석자 추정입니다. 검증 지표의 목표 수치는 출처가 확인된 값만 'MVP 검증 목표' 시트에 정리했습니다(장르별 리텐션은 공식 숫자 자료가 없어, 2025 보고서 이미지 차트를 읽은 '판독값'을 '장르 판독 요약' 시트에 따로 정리).").font=Font(name=F,italic=True,color="595959")

# --- 성공 공식 ---
ws=wb.create_sheet("성공 공식 TOP5")
ws.append(["순위","성공 공식","쉽게 말하면","근거"])
form=[
(1,"3초 안에 이해되는 규칙","설명서 없이 보자마자 할 수 있어야 해요","6개국 무료 Top 10에 공통으로 오른 Block Blast·Meowdoku·Bus Fever Party·Rotate Rings 모두 한 손가락·한 규칙 게임"),
(2,"'많이 받는 게임'과 '돈 버는 게임'은 다르다","인기 1등이 꼭 부자는 아니에요","무료 상위 퍼즐 5개 중 매출 Top 100은 Magic Sort(미국 56위) 1개뿐. 반대로 Whiteout은 무료 순위가 낮아도 4개국 매출 Top 25"),
(3,"퍼즐 + 꾸미기/이야기 = 돈","퍼즐만 있으면 심심, 내 성·드라마가 있으면 계속 와요","Royal Match(성 꾸미기)·Gossip Harbor(스토리)가 퍼즐 장르로 4개국 매출 모두 Top 25"),
(4,"같이 하면 오래 간다","친구·동맹과 함께하면 그만두기 어려워요","MONOPOLY GO!(친구 이벤트) 미국 매출 1위, Whiteout(동맹) 한국 매출 6위, Free Fire(스쿼드) 브라질 매출 1위"),
(5,"현지 사람이 아는 얼굴(IP)을 빌린다","유명한 캐릭터가 나오면 처음부터 믿고 받아요","Free Fire x NARUTO, MONOPOLY, 치이카와(일본 무료 20위), Disney Solitaire"),
]
for r in form: ws.append(r)
nr=ws.max_row+2
ws.cell(nr,1,"가장 작게 시작할 수 있는 MVP 순위").font=Font(name=F,bold=True,size=12)
ws.append([])
ws.append(["순위","게임","MVP 규모","이유"])
hdrrow=ws.max_row
easy=[(1,"Block Blast!","S","판 하나·규칙 하나, 레벨 제작도 필요 없음"),
(2,"Rotate Rings","S","조작 하나 + 레벨 데이터"),(3,"Bus Fever Party!","S","색 매칭 규칙 + 레벨"),
(4,"Meowdoku","S","논리 퍼즐 엔진 + 레벨 생성기"),(5,"Magic Sort!","S","검증된 규칙, 연출 품질이 관건"),
(6,"Royal Match","M","매치3 엔진 + 꾸미기 메타"),(7,"후더덕 서바이벌","M","방치 루프 + 뽑기, 밸런스 작업 많음"),
(8,"Gossip Harbor","L","스토리·아트 콘텐츠 양이 많음"),(9,"MONOPOLY GO!","L","IP + 실시간 이벤트"),(10,"치이카와 포켓","L","IP 라이선스 필수"),
(11,"Whiteout Survival","XL","서버·라이브 운영·마케팅 비용"),(12,"Free Fire","XL","실시간 멀티플레이 서버")]
for r in easy: ws.append(r)
hdr2(ws,1); hdr2(ws,hdrrow)
for row in ws.iter_rows(min_row=2):
    for c in row:
        if c.value is not None and c.row!=nr:
            c.font=B; c.border=BD; c.alignment=Alignment(wrap_text=True,vertical="top")
    if row[0].value is not None and row[0].row not in (nr,hdrrow): row[0].alignment=Alignment(horizontal="center",vertical="top")
for r in range(2,7): ws.cell(r,2).font=BB
ws.cell(nr,1).font=Font(name=F,bold=True,size=12)
for col,w in zip("ABCD",[8,34,46,80]): ws.column_dimensions[col].width=w

# --- 매출 순위 원자료 ---
ws=wb.create_sheet("매출 순위 원자료")
ws.append(["국가","기준일","출처","Top 10 (구글 플레이 매출)"])
raw=[("미국","2026-10-03","https://www.appbrain.com/stats/google-play-rankings/top_grossing/game/us","1 MONOPOLY GO!, 2 Roblox, 3 Candy Crush Saga, 4 Royal Match, 5 Gossip Harbor, 6 Township, 7 Coin Master, 8 Royal Kingdom, 9 Kingshot, 10 Whiteout Survival"),
("브라질","2026-10-01","https://www.appbrain.com/stats/google-play-rankings/top_grossing/game/br","1 Free Fire, 2 Roblox, 3 Coin Master, 4 Gossip Harbor, 5 Candy Crush Saga, 6 eFootball, 7 Royal Match, 8 Minecraft, 9 Brawl Stars, 11 Gardenscapes"),
("일본","2026-10-04","https://www.appbrain.com/stats/google-play-rankings/top_grossing/game/jp","1 SD건담 G제네레이션 이터널, 2 Wuthering Waves, 3 Honkai: Star Rail, 4 홀로라이브 드림즈, 5 몬스터 스트라이크, 6 드래곤볼Z 돗칸배틀, 7 우마무스메, 8 원신, 9 Whiteout Survival, 10 디즈니 츠무츠무"),
("한국","2026-10-03","https://www.appbrain.com/stats/google-play-rankings/top_grossing/game/kr","1 제우스: 오만의 신, 2 리니지M, 3 Wuthering Waves, 4 이클립스: 더 어웨이크닝, 5 오딘, 6 Whiteout Survival, 7 트릭컬 리바이브, 8 메이플스토리 방치형 RPG, 9 Roblox, 10 SOL: enchant")]
for r in raw: ws.append(r)
hdr2(ws); body(ws)
for col,w in zip("ABCD",[8,12,62,110]): ws.column_dimensions[col].width=w


# ===================== 3단계: 벤치마크 & MVP 검증 목표 =====================
GA26="GameAnalytics, 2026 Mobile & PC Gaming Benchmarks (2025년 데이터, 모바일 16,000+개, MAU 1,000+, iOS+Android, 전 장르 합산) https://www.gameanalytics.com/reports/2026-mobile-pc-gaming-benchmarks"
GA25="GameAnalytics, 2025 Mobile Gaming Benchmarks (2024년 데이터, 11,600개 게임) https://files.gameindustrylibrary.com/documents/mobile-gaming-benchmarks-2025.pdf"
AF26="AppsFlyer, State of App Monetization 2026 (IAP $9억·광고 $72억, 2025.1~2026.3) — 요약: https://gamedevreports.substack.com/p/appsflyer-app-monetization-in-2026 / 원문: https://www.appsflyer.com/resources/reports/app-marketing-monetization-report/"
AF24="AppsFlyer, State of Gaming App Marketing 2024 https://www.appsflyer.com/resources/reports/gaming-app-marketing-2024-report/"
APD="Appodeal, Mobile Casual Benchmarks Report 2025 (미국 Android 캐주얼 게임 10,000+개, 2024.6~2025.1) — App Developer Magazine 보도 https://appdevelopermagazine.com/mobile-casual-benchmarks-report-2025/"

ws=wb.create_sheet("벤치마크 근거")
ws.append(["분류","지표","구간","값","조건·범위","출처","비고"])
bm=[
("리텐션(전체)","D1","하위 25% (P25)","13.42% → 12.45%","2025년 연초→연말","GA26",""),
("리텐션(전체)","D1","중앙값 (P50)","약 22%","2025년 연초","GA26",""),
("리텐션(전체)","D1","상위 25% (P75)","30% 약간 상회","2025년","GA26",""),
("리텐션(전체)","D1","상위 10% (P90)","약 40%","2025년 연중","GA26",""),
("리텐션(전체)","D1","상위 1% (P99)","64~68%","2025년","GA26",""),
("리텐션(전체)","D7","하위 25% (P25)","1.67~1.94%","2025년","GA26",""),
("리텐션(전체)","D7","중앙값 (P50)","4% 약간 미만","2025년","GA26",""),
("리텐션(전체)","D7","상위 25% (P75)","6~7%","2025년","GA26",""),
("리텐션(전체)","D7","상위 10% (P90)","11~12%","2025년 대부분","GA26",""),
("리텐션(전체)","D7","상위 1% (P99)","25% 이상 (최고 28% 초과)","2025년","GA26",""),
("리텐션(전체)","D30","하위 25% (P25)","0.5% 미만","2025년","GA26",""),
("리텐션(전체)","D30","중앙값 (P50)","0.68~0.79%","2025년","GA26",""),
("리텐션(전체)","D30","상위 25% (P75)","1.6~1.8%","2025년","GA26",""),
("리텐션(전체)","D30","상위 10% (P90)","전체 기준 미공개","—","GA26","지역별 P90은 '지역 벤치마크' 시트 참고"),
("리텐션(전체)","D30","상위 1% (P99)","13~15%","2025년","GA26",""),
("리텐션(플랫폼)","D1","상위 25% (P75)","iOS 31~33% / Android 25~27%","2024년 데이터","GA25","iOS가 Android보다 높음"),
("참여도(전체)","하루 플레이 시간","P50 / P75 / P90 / P99","약 12분 / 22~24분 / 40~42분 / 94분 초과","2025년","GA26",""),
("참여도(전체)","세션 길이","P50 / P75 / P90 / P99","3.1~3.5분 / 약 5.2분 / 8분 약간 초과 / 22분 초과","2025년","GA26",""),
("참여도(전체)","하루 세션 수","P25 / P50 / P75 / P90 / P99","2.7 / 3.8~3.9 / 5.3~5.7 / 일부 월 9.6 / 12 초과","2025년","GA26",""),
("장르(정성)","리텐션","중앙값 기준","보드·카드·퍼즐·카지노가 D1·D7·D28 모두 강함, 아케이드는 D1만 강함, 멀티플레이는 리텐션 최저","2024년 데이터","GA25","장르별 숫자는 원문 PDF에 이미지 차트로만 있음 → 눈으로 읽은 판독값을 '장르 차트 판독값' 시트에 정리(공식 텍스트 수치 아님)"),
("장르(정성)","하루 세션 수","중앙값","미드코어 6~7회 (전 장르 중 최고)","2024년 데이터","GA25",""),
("장르","D1·D7·D30 수치","—","미공개","—","GA26","2026 보고서는 장르 분류 개편 중이라 장르별 수치를 제공하지 않는다고 명시"),
("수익(장르)","D90 IAP ARPU","카지노 / 미드코어 / 캐주얼","$2.43 / $2.13 / $1.34","글로벌","AF26","북미+유럽: 미드코어 $2.66, 카지노 $2.55"),
("수익(장르)","D90 광고(IAA) ARPU","캐주얼 / 카지노 / 미드코어 / 하이퍼캐주얼","$0.55 / $0.47 / $0.40 / $0.22","글로벌 평균","AF26","북미·유럽은 더 높다고 언급"),
("수익(장르)","D90 IAP ARPPU","카지노 / 미드코어 / 캐주얼","$11.40 / $9.80 / $7.26","글로벌","AF26",""),
("수익(장르)","결제 전환율","카지노","신규 설치의 4.95%가 결제, 3.01%가 2회 이상 결제","30일 이내로 추정(원문 확인 필요)","AF26","다른 장르의 전환율은 요약본에 없음"),
("수익(장르)","첫 결제 대비 재결제 비율","카지노 / 미드코어 / 캐주얼","1.65배 / 1.76배 / 1.89배","낮을수록 좋음","AF26",""),
("수익(장르)","IAP 매출 중 유료 유입 비중","하이퍼캐주얼 / 카지노 / 캐주얼 / 미드코어","79% / 64% / 61% / 49%","","AF26",""),
("수익(지역)","매출 중 유료 유입 비중","남미(LATAM)","전체 37%, 미드코어 43%","","AF26",""),
("수익(구조)","매출 구성","미드코어 / 카지노","IAP 90% / IAP 83%","","AF26",""),
("수익(시점)","D60 매출 중 D7까지 누적","광고 / IAP","89% / 60%","전체 앱","AF26","하이퍼캐주얼은 D60 광고 매출의 63%가 D1에 발생"),
("수익(일반)","결제 유저 비율","전체 게임","대개 5% 미만 (RPG는 더 높음)","","AF24",""),
("수익(일반)","첫 결제 시점","iOS","D2 전후 첫 결제가 전체 결제자의 1/4, D3에 17% 추가","","AF24",""),
("수익(일반)","상위 결제자 집중도","Android, 2024 Q4","상위 5% 결제자가 IAP 매출의 45.95%","앱당 평균","AF24",""),
("광고 수익(캐주얼)","광고 ARPU","하이퍼캐주얼 / 파티 / 매치","$0.86 / $4.90 / $2.99","미국 Android","APD","ARPU 측정 기간이 보도에 명시되지 않음"),
("광고 수익(캐주얼)","광고 ARPU (세부 장르)","머지3 / 럭배틀 / 러닝 / 슬라이싱","$14.83 / $12.23 / $2.34 / $2.19","미국 Android","APD","ARPU 측정 기간이 보도에 명시되지 않음"),
("광고 수익(캐주얼)","유저당 광고 노출 수","퍼즐: 전면 / 보상형 / 배너","72.5 / 23.4 / 241.5","미국 Android","APD",""),
("광고 수익(캐주얼)","유저당 광고 노출 수","매치: 전면 / 보상형 / 배너","36.5 / 39.1 / 114.3","미국 Android","APD",""),
("광고 수익(캐주얼)","유저당 보상형 광고 수","머지3 / 방치형","101.5 / 73.2","미국 Android","APD",""),
]
srcmap={"GA26":GA26,"GA25":GA25,"AF26":AF26,"AF24":AF24,"APD":APD}
for r in bm: ws.append(list(r[:5])+[srcmap[r[5]],r[6]])
hdr2(ws); body(ws)
for row in ws.iter_rows(min_row=2):
    row[3].font=BB
    if "미공개" in str(row[3].value) or "미확인" in str(row[6].value) or "명시되지" in str(row[6].value) or "확인 필요" in str(row[4].value):
        for c in row: c.fill=YEL
for col,w in zip("ABCDEFG",[16,22,30,40,22,70,42]): ws.column_dimensions[col].width=w
ws.freeze_panes="A2"; ws.auto_filter.ref=f"A1:G{ws.max_row}"

# 지역 벤치마크 (숫자, 수식 조회용)
ws=wb.create_sheet("지역 벤치마크")
ws.append(["지역","지표","P50 (중앙값)","P90 (상위 10%)","P99 (상위 1%)","단위","출처"])
reg={
"북미":[(0.2328,0.3690,0.5089),(0.0497,0.1258,0.2478),(0.0118,0.0478,0.1326),(14.45,45.84,89.43),(3.64,8.50,20.22),(4.23,7.59,12.85)],
"중미":[(0.2137,0.3424,0.4683),(0.0371,0.0960,0.1815),(0.0078,0.0285,0.0810),(12.41,36.42,66.26),(2.83,6.25,17.32),(4.90,7.62,11.06)],
"남미":[(0.2009,0.3406,0.4689),(0.0325,0.0914,0.1894),(0.0065,0.0254,0.0793),(11.14,35.77,67.52),(2.77,6.44,15.62),(4.49,7.32,10.49)],
"아시아":[(0.1750,0.3277,0.4558),(0.0268,0.0865,0.2008),(0.0053,0.0259,0.1005),(11.79,40.72,92.07),(3.12,7.15,17.69),(4.19,7.79,13.32)],
}
mets=["D1 리텐션","D7 리텐션","D30 리텐션","하루 플레이 시간","세션 길이","하루 세션 수"]
units=["%","%","%","분","분","회"]
for rg,vals in reg.items():
    for m,u,v in zip(mets,units,vals):
        ws.append([rg,m,v[0],v[1],v[2],u,"GameAnalytics 2026 보고서 지역별 표 (2025년 주간 데이터의 연평균)"])
hdr2(ws); body(ws,wrap=False)
for row in ws.iter_rows(min_row=2):
    for c in row[2:5]:
        c.font=BLUE; c.alignment=Alignment(horizontal="center")
        c.number_format="0.00%" if row[5].value=="%" else "0.00"
for col,w in zip("ABCDEFG",[10,18,14,15,14,6,60]): ws.column_dimensions[col].width=w
ws.freeze_panes="A2"; ws.auto_filter.ref=f"A1:G{ws.max_row}"
nr=ws.max_row+2
ws.cell(nr,1,"참고 (2024년 데이터, 중앙값): 북미 D1 20.48% / D7 4.64% / D28 1.46%, 남미 18.32% / 3.31% / 0.82%, 중미 18.95% / 3.68% / 0.98%, 아시아 17.57% / 3.14% / 0.75%, 동남아 18.28% / 3.33% / 0.82% — GameAnalytics 2025 보고서").font=Font(name=F,italic=True,color="595959")
ws.cell(nr+1,1,"멕시코가 GameAnalytics의 '중미'와 '북미' 중 어디에 속하는지는 보고서에서 확인되지 않았습니다.").font=Font(name=F,italic=True,color="595959")
REGLAST=ws.max_row

# MVP 검증 목표
ws=wb.create_sheet("MVP 검증 목표")
ws.append(["","","","D1 리텐션","","","D7 리텐션","","","D30 리텐션","","","수익 참고 (AppsFlyer, D90, 글로벌)","","",""])
ws.append(["게임","주 타깃 지역","수익 장르 분류","합격선: 전체 상위 25%","목표: 타깃 지역 상위 10%","지역 중앙값",
           "합격선: 전체 상위 25%","목표: 타깃 지역 상위 10%","지역 중앙값",
           "합격선: 전체 상위 25%","목표: 타깃 지역 상위 10%","지역 중앙값",
           "IAP ARPU","광고 ARPU","IAP ARPPU","장르별 리텐션"])
plan=[("Meowdoku","북미","캐주얼"),("Rotate Rings","북미","캐주얼"),("Block Blast!","북미","캐주얼"),("Magic Sort!","북미","캐주얼"),
("Bus Fever Party!","북미","캐주얼"),("Royal Match","북미","캐주얼"),("Gossip Harbor","북미","캐주얼"),("MONOPOLY GO!","북미","캐주얼"),
("Whiteout Survival","북미","미드코어"),("Free Fire","남미","미드코어"),("치이카와 포켓","아시아",""),("후더덕 서바이벌","아시아","미드코어")]
R="'지역 벤치마크'!"
def L(col,i,met): return f'=SUMIFS({R}${col}$2:${col}${REGLAST},{R}$A$2:$A${REGLAST},$B{i},{R}$B$2:$B${REGLAST},"{met}")'
arpu={"캐주얼":(1.34,0.55,7.26),"미드코어":(2.13,0.40,9.80)}
for i,(g,rg,bk) in enumerate(plan,start=3):
    a=arpu.get(bk,("","",""))
    ws.append([g,rg,bk if bk else "미정 (게임 방식 미확인)",
        "30% 약간 상회",L("D",i,"D1 리텐션"),L("C",i,"D1 리텐션"),
        "6~7%",L("D",i,"D7 리텐션"),L("C",i,"D7 리텐션"),
        "1.6~1.8%",L("D",i,"D30 리텐션"),L("C",i,"D30 리텐션"),
        a[0],a[1],a[2],"미공개"])
for rng in ["D1:F1","G1:I1","J1:L1","M1:O1"]: ws.merge_cells(rng)
hdr2(ws,1); hdr2(ws,2)
for row in ws.iter_rows(min_row=3):
    for c in row: c.font=B; c.border=BD; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    row[0].font=BB; row[0].alignment=Alignment(horizontal="left",vertical="center")
    for idx in (4,5,7,8,10,11): row[idx].number_format="0.00%"
    for idx in (12,13,14):
        row[idx].number_format='"$"0.00'; row[idx].font=BLUE
    row[1].font=BLUE; row[2].font=BLUE
    if "미정" in str(row[2].value):
        for c in row[2:3]+row[12:15]: c.fill=YEL
    row[15].fill=YEL
for col,w in zip("ABC",[18,12,20]): ws.column_dimensions[col].width=w
for j in range(4,17): ws.column_dimensions[get_column_letter(j)].width=13
ws.row_dimensions[2].height=45
ws.freeze_panes="B3"
nr=ws.max_row+2
notes=[
"모든 숫자는 '벤치마크 근거'·'지역 벤치마크' 시트의 출처값을 그대로 가져온 것이며, 새로 만든 수치는 없습니다.",
"'합격선'과 '목표'를 어느 구간(P75·P90)으로 둘지는 분석자 제안입니다. 기준을 바꾸려면 D·G·J열 또는 '지역 벤치마크' 조회 열을 바꾸면 됩니다.",
"주 타깃 지역과 수익 장르 분류(캐주얼·미드코어)는 분석자 판단입니다(파란 글씨). 지역을 바꾸면 지역 수치가 자동으로 바뀝니다.",
"리텐션 수치는 전 장르 합산값입니다. 장르별 수치는 GameAnalytics 2026 보고서에 없고, 2025 보고서 이미지 차트에서 읽은 판독값은 '장르 판독 요약'·'장르 연결(추정)' 시트에 따로 두었습니다(이 표에는 섞지 않음).",
"D30 리텐션은 2026 보고서 기준이며, 2025 보고서는 D28을 사용해 직접 비교에 주의가 필요합니다.",
"수익 수치(AppsFlyer)는 D90 기준 글로벌 평균이며, 지역·플랫폼에 따라 다릅니다(북미·유럽은 더 높다고 보고됨)."]
for k,t in enumerate(notes): ws.cell(nr+k,1,t).font=Font(name=F,italic=True,color="595959")

# 사용하지 않은 자료
ws=wb.create_sheet("제외한 자료")
ws.append(["자료 유형","예시","제외 이유"])
exc=[("블로그·용어집의 범위값","'모바일 D1 35~45%가 좋은 성과', 'ARPDAU 캐주얼 $0.05~0.15' 등","1차 데이터 출처가 제시되지 않음. 일부 분석 글은 이런 표가 오래된 AppsFlyer 데이터에서 반복 인용된 것이라고 지적"),
("오래된 장르별 표","AppsFlyer 2022년 3분기 장르 데이터(캐주얼·퍼즐 Android D1 28~32% 등)","4년 전 데이터이고 2차 인용으로만 확인됨"),
("지역 범위가 불명확한 수치","북미 첫 결제 전환율 11.14% vs 남미 6.51% (AppsFlyer 2026 요약)","게임만의 수치인지 비게임 포함인지 요약본에서 확인되지 않음"),
("비공개 보고서","유료 플랫폼(Sensor Tower 등)의 장르 벤치마크","접근에 계정·결제가 필요함 (GameAnalytics 2025 장르별 차트는 5단계에서 '판독값'으로 별도 사용)")]
for r in exc: ws.append(r)
hdr2(ws); body(ws)
for col,w in zip("ABC",[24,60,70]): ws.column_dimensions[col].width=w


# ===================== 4단계: 장르별 확장 + 공통점 =====================
ws=wb.create_sheet("성공 비결 확장")
ws.append(["장르 그룹","게임","퍼블리셔","차트 근거 (이번 조사에서 확인된 순위)","① 핵심 재미 루프","③ 다시 오게 하는 장치","④ 수익 모델(BM)","⑤ 마케팅","⑥ 아트","근거 수준"])
ext=[
("1. 글로벌 퍼즐","Block Out!","Grand Games","미국 App Store 무료 2위","블록을 같은 색 출구로 빼내는 정렬 퍼즐(게임명 기반 추정)","레벨 진행(추정)","미확인 (App Store 매출 데이터 미수집)","같은 회사 Magic Sort!와 함께 미국 상위권","색 블록 위주(추정)","차트 데이터 + 게임명 추정"),
("1. 글로벌 퍼즐","Vita Mahjong","Vita Studio","무료: 미국 5, 캐나다 9, 브라질 29, 멕시코 15, 일본 34, 한국 26","같은 타일 2개 짝 맞춰 없애기","레벨 진행(추정)","미확인","6개국 모두 Top 35 — 언어가 필요 없는 규칙","큰 타일·깔끔한 판(시니어 친화로 알려짐, 확인 권장)","차트 데이터 + 공개 정보(확인 권장)"),
("2. 퍼즐 + 메타","Candy Crush Saga","King","매출: 미국 3, 브라질 5 / 무료: 미국 GP 48","3개 맞추기 → 레벨 클리어 → 지도 따라 전진","긴 레벨 지도, 이벤트, 친구 점수 비교","목숨·부스터·이동 추가 인앱결제","오래된 브랜드 인지도","사탕 테마의 밝은 2D","매출 데이터 + 공개 게임 방식"),
("2. 퍼즐 + 메타","Royal Kingdom","Dream Games","매출: 미국 8 / 무료: 일본 5, 한국 17, 미국 GP 47","3개 맞추기 → 보상 → 왕국 건설(추정 세부)","왕국 건설 메타(추정)","인앱결제 (미국 매출 8위)","Royal Match 개발사의 후속작","3D 카툰 판타지","매출 데이터 + 일부 추정"),
("3. 소셜 캐주얼","Coin Master","Moon Active","매출: 미국 7, 브라질 3 / 무료: 미국 iOS 67","슬롯 돌리기 → 코인 → 마을 건설, 다른 유저 공격·습격","카드 수집, 친구와 공격·복수, 이벤트","스핀·코인 인앱결제","친구 초대·SNS 연동(공개 정보)","바이킹 테마 카툰","매출 데이터 + 공개 게임 방식"),
("3. 소셜 캐주얼","Dice Dreams","SuperPlay","무료: 미국 iOS 73, GP 98","주사위 굴리기 → 보드 칸 보상 → 왕국 건설, 다른 유저 공격","이벤트(공개 정보), 카드 수집 여부는 확인 필요","인앱결제 (매출 순위 미확인)","같은 회사가 Disney Solitaire도 운영","동화풍 카툰","차트 데이터 + 공개 게임 방식"),
("4. 전략(4X 등)","Kingshot","Century Games","매출: 미국 9 / 무료: 미국 iOS 40","건설 → 자원 → 영웅 성장 → 동맹 전투(장르 일반)","동맹, 서버 경쟁(장르 일반)","인앱결제 (미국 매출 9위)","같은 회사 Whiteout Survival과 같은 장르","3D 카툰(추정)","매출 데이터 + 장르 일반 지식"),
("4. 전략(4X 등)","Clash Royale","Supercell","무료: 미국 iOS 80 (매출 미확인)","카드 덱 → 3분 실시간 1:1 대전 → 트로피","아레나 승급, 클랜, 시즌","패스·보석 인앱결제(공개 정보)","Clash of Clans 세계관·캐릭터 재사용","클래시 시리즈 카툰","차트 데이터 + 공개 게임 방식"),
("5. 액션·슈팅","Fortnite","Epic Games","무료: 미국 iOS 18, GP 25 (매출 미확인)","낙하 → 장비 줍기 → 건축·전투 → 최후 생존","배틀패스 시즌, 라이브 이벤트, 유저 제작 모드","스킨·배틀패스 (외형 아이템)","대형 IP 콜라보","밝은 카툰 3D","차트 데이터 + 공개 게임 방식"),
("5. 액션·슈팅","Brawl Stars","Supercell","매출: 브라질 9 / 무료: 미국 iOS 66","3대3 짧은 대전 → 트로피 → 캐릭터(브롤러) 해금","브롤러 수집, 클럽, 패스","패스·스킨 인앱결제","Supercell 세계관","캐주얼 카툰","매출 데이터 + 공개 게임 방식"),
("6. 캐릭터 IP","Pokémon GO","Scopely","무료: 한국 3, 미국 iOS 76","걸어 다니며 포켓몬 발견 → 포획 → 수집","레이드, 커뮤니티 이벤트(공개 정보)","아이템 인앱결제","포켓몬 IP + 실외 활동","AR + 포켓몬 원작 그림","차트 데이터 + 공개 게임 방식"),
("6. 캐릭터 IP","Pokémon TCG Pocket","The Pokémon Company","무료: 일본 30, 미국 GP 68","매일 카드팩 열기 → 수집 → 대전","매일 무료 카드팩(공개 정보)","카드팩 인앱결제 (랜덤 보상)","포켓몬 IP + 실물 카드의 추억","원작 카드 일러스트","차트 데이터 + 공개 게임 방식"),
("7. 한국식 RPG","메이플스토리 방치형 RPG","Nexon","매출: 한국 8 / 무료: 한국 54","자동 사냥 → 장비·스탯 성장 → 보스","방치 보상, 길드(장르 일반)","성장 패키지·뽑기 인앱결제","메이플스토리 IP","원작 2D 그림체","매출 데이터 + 장르 일반 지식"),
("7. 한국식 RPG","트릭컬 리바이브","Epid Games","매출: 한국 7 / 무료: 한국 44","캐릭터 뽑기 → 육성 → 전투(세부 확인 필요)","캐릭터 수집, 이벤트","뽑기 인앱결제","게임명에 '600뽑기 증정' 표기","귀여운 2D 캐릭터(확인 권장)","매출 데이터 + 일부 추정"),
]
for r in ext: ws.append(r)
hdr2(ws); body(ws)
for row in ws.iter_rows(min_row=2):
    row[1].font=BB
    if "추정" in str(row[9].value) or "확인" in str(row[9].value): row[9].fill=YEL
for col,w in zip("ABCDEFGHIJ",[16,22,18,34,36,30,30,30,24,24]): ws.column_dimensions[col].width=w
ws.freeze_panes="C2"; ws.auto_filter.ref=f"A1:J{ws.max_row}"
ws.cell(ws.max_row+2,1,"기존 12개 게임은 '성공 비결 해부' 시트에 있습니다. 확장 게임은 차트 데이터를 다시 모으지 않아 '성공 점수표'에는 넣지 않았습니다.").font=Font(name=F,italic=True,color="595959")

# 공통점 분석 매트릭스
tags=["한 손가락 단순 조작","짧은 한 판(몇 분)","메타 진행(꾸미기·건설·스토리·성장)","수집 요소","친구·동맹·대전(소셜)","유명 IP·콜라보","랜덤 보상(뽑기·슬롯·카드팩)","광고 수익 중심","인앱결제 중심","라이브 이벤트·시즌"]
Q="?"
rows=[
("1. 글로벌 퍼즐","Meowdoku","아니오",[1,1,Q,Q,Q,0,0,1,0,Q]),
("1. 글로벌 퍼즐","Rotate Rings","아니오",[1,1,Q,Q,Q,0,0,1,0,Q]),
("1. 글로벌 퍼즐","Block Blast!","아니오",[1,1,Q,0,0,0,0,1,0,Q]),
("1. 글로벌 퍼즐","Magic Sort!","예",[1,1,Q,Q,Q,0,0,1,1,Q]),
("1. 글로벌 퍼즐","Bus Fever Party!","아니오",[1,1,Q,Q,Q,0,0,1,0,Q]),
("1. 글로벌 퍼즐","Block Out!","미확인",[1,1,Q,Q,Q,0,0,Q,Q,Q]),
("1. 글로벌 퍼즐","Vita Mahjong","미확인",[1,1,Q,Q,Q,0,0,Q,Q,Q]),
("2. 퍼즐 + 메타","Royal Match","예",[1,1,1,Q,1,0,0,0,1,1]),
("2. 퍼즐 + 메타","Gossip Harbor","예",[1,1,1,Q,Q,0,0,0,1,Q]),
("2. 퍼즐 + 메타","Candy Crush Saga","예",[1,1,1,0,1,0,0,0,1,1]),
("2. 퍼즐 + 메타","Royal Kingdom","예",[1,1,1,Q,Q,0,0,0,1,Q]),
("3. 소셜 캐주얼","MONOPOLY GO!","예",[1,1,1,1,1,1,1,0,1,1]),
("3. 소셜 캐주얼","Coin Master","예",[1,1,1,1,1,Q,1,0,1,1]),
("3. 소셜 캐주얼","Dice Dreams","미확인",[1,1,1,Q,1,0,1,0,1,1]),
("4. 전략(4X 등)","Whiteout Survival","예",[0,0,1,1,1,0,1,0,1,1]),
("4. 전략(4X 등)","Kingshot","예",[0,0,1,1,1,0,Q,0,1,Q]),
("4. 전략(4X 등)","Clash Royale","미확인",[0,1,1,1,1,0,1,0,1,1]),
("5. 액션·슈팅","Free Fire","예",[0,1,1,1,1,1,Q,0,1,1]),
("5. 액션·슈팅","Fortnite","미확인",[0,0,1,1,1,1,0,0,1,1]),
("5. 액션·슈팅","Brawl Stars","예",[0,1,1,1,1,Q,Q,0,1,1]),
("6. 캐릭터 IP","치이카와 포켓","예",[Q,Q,Q,Q,Q,1,Q,0,Q,Q]),
("6. 캐릭터 IP","Pokémon GO","미확인",[1,1,1,1,1,1,0,0,1,1]),
("6. 캐릭터 IP","Pokémon TCG Pocket","미확인",[1,1,Q,1,1,1,1,0,1,Q]),
("7. 한국식 RPG","후더덕 서바이벌","예",[Q,Q,1,Q,Q,0,1,0,1,Q]),
("7. 한국식 RPG","메이플스토리 방치형 RPG","예",[1,Q,1,Q,1,1,1,0,1,Q]),
("7. 한국식 RPG","트릭컬 리바이브","예",[Q,Q,1,1,Q,0,1,0,1,1]),
]
ws=wb.create_sheet("공통점 분석")
ws.append(["장르 그룹","게임","매출 Top100 진입 (4개국 중 1곳 이상)"]+tags)
for g,n,gr,v in rows: ws.append([g,n,gr]+v)
N=ws.max_row
hdr2(ws); ws.row_dimensions[1].height=60
GREEN=PatternFill("solid",fgColor="C6EFCE"); GREY=PatternFill("solid",fgColor="EDEDED")
for row in ws.iter_rows(min_row=2,max_row=N):
    for c in row: c.font=B; c.border=BD; c.alignment=Alignment(horizontal="center",vertical="center")
    row[0].alignment=Alignment(horizontal="left"); row[1].alignment=Alignment(horizontal="left"); row[1].font=BB
    for c in row[3:]:
        c.font=BLUE
        if c.value==1: c.fill=GREEN
        elif c.value==Q: c.fill=GREY
    if row[2].value=="미확인": row[2].fill=YEL
for col,w in zip("ABC",[16,24,16]): ws.column_dimensions[col].width=w
for j in range(4,14): ws.column_dimensions[get_column_letter(j)].width=13
ws.freeze_panes="D2"
# 요약표
s0=N+2
ws.cell(s0,1,"공통점 요약 (자동 계산)").font=Font(name=F,bold=True,size=12)
hdr_row=s0+1
heads=["요소","'예' 게임 수","확인된 게임 중 비율","매출 진입 게임 중 비율","매출 미진입 게임 중 비율","차이(매출 진입 − 미진입)"]
for j,h in enumerate(heads,1): ws.cell(hdr_row,j,h)
for k,t in enumerate(tags):
    r=hdr_row+1+k; col=get_column_letter(4+k); rng=f"${col}$2:${col}${N}"; grr=f"$C$2:$C${N}"
    ws.cell(r,1,t)
    ws.cell(r,2,f"=COUNTIF({rng},1)")
    ws.cell(r,3,f"=IFERROR(COUNTIF({rng},1)/(COUNTIF({rng},1)+COUNTIF({rng},0)),\"-\")")
    ws.cell(r,4,f"=IFERROR(COUNTIFS({rng},1,{grr},\"예\")/(COUNTIFS({rng},1,{grr},\"예\")+COUNTIFS({rng},0,{grr},\"예\")),\"-\")")
    ws.cell(r,5,f"=IFERROR(COUNTIFS({rng},1,{grr},\"아니오\")/(COUNTIFS({rng},1,{grr},\"아니오\")+COUNTIFS({rng},0,{grr},\"아니오\")),\"-\")")
    ws.cell(r,6,f"=IFERROR(D{r}-E{r},\"-\")")
for c in ws[hdr_row]:
    if c.value: c.fill=H_FILL; c.font=H_FONT; c.alignment=Alignment(horizontal="center",wrap_text=True)
for r in range(hdr_row+1,hdr_row+1+len(tags)):
    for j in range(1,7):
        c=ws.cell(r,j); c.font=B; c.border=BD
        if j>=3: c.number_format="0%"; c.alignment=Alignment(horizontal="center")
    ws.cell(r,1).font=BB
ws.row_dimensions[hdr_row].height=32
# 그룹별
g0=hdr_row+len(tags)+2
ws.cell(g0,1,"장르 그룹별 '예' 개수 (자동 계산)").font=Font(name=F,bold=True,size=12)
groups=sorted(set(r[0] for r in rows))
ws.cell(g0+1,1,"장르 그룹"); ws.cell(g0+1,2,"게임 수")
for k,t in enumerate(tags): ws.cell(g0+1,3+k,t)
for i,gname in enumerate(groups):
    r=g0+2+i
    ws.cell(r,1,gname); ws.cell(r,2,f"=COUNTIF($A$2:$A${N},A{r})")
    for k in range(len(tags)):
        col=get_column_letter(4+k)
        ws.cell(r,3+k,f"=COUNTIFS($A$2:$A${N},$A{r},${col}$2:${col}${N},1)")
for c in ws[g0+1]:
    if c.value: c.fill=H_FILL; c.font=H_FONT; c.alignment=Alignment(horizontal="center",wrap_text=True)
ws.row_dimensions[g0+1].height=60
for r in range(g0+2,g0+2+len(groups)):
    for j in range(1,3+len(tags)):
        c=ws.cell(r,j); c.font=B; c.border=BD
        if j>=2: c.alignment=Alignment(horizontal="center")
nr=g0+3+len(groups)
notes=["1=해당, 0=해당 안 함, ?=확인 안 됨(회색). 분류는 공개된 게임 방식과 이번 차트 조사에 근거한 분석자 판단이며, '?'는 계산에서 제외됩니다.",
"'매출 Top100 진입'은 이번에 조사한 구글 플레이 매출 순위(미국·브라질·일본·한국)에서 확인된 경우만 '예'입니다. 조사 범위 밖이면 '미확인'.",
"게임 수가 적어(매출 진입 15개, 미진입 4개, 미확인 7개) 비율은 경향을 보는 참고용이며 통계적 결론이 아닙니다."]
for k,t in enumerate(notes): ws.cell(nr+k,1,t).font=Font(name=F,italic=True,color="595959")


# ===================== 5단계: GameAnalytics 2025 장르별 차트 판독값 =====================
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation
GENRES=[("Action","액션"),("Adventure","어드벤처"),("Arcade","아케이드"),("Board","보드"),("Card","카드"),("Casino","카지노"),
("Casual","캐주얼"),("Multiplayer","멀티플레이"),("Puzzle","퍼즐"),("Racing","레이싱"),("Role Playing","롤플레잉(RPG)"),
("Simulation","시뮬레이션"),("Sports","스포츠"),("Strategy","전략"),("Trivia","퀴즈(트리비아)"),("Word","단어")]
# 판독값: PDF 이미지(1056x547 JPEG)를 확대해 눈으로 읽은 값. %는 숫자로 입력(18.50 = 18.50%), 아래에서 /100.
RD={
"D1":"""18.50 18.51 18.78 18.76 18.70 18.40 18.06 18.33 18.21 18.27 18.00 17.72
16.50 16.53 16.64 16.53 16.14 15.66 15.47 15.56 16.07 15.97 15.85 15.70
22.33 21.64 22.53 22.66 22.63 22.29 22.03 22.14 22.35 22.47 22.06 21.79
19.81 20.06 19.91 19.45 19.50 18.94 18.85 18.62 18.67 18.74 18.64 18.00
19.05 19.39 19.16 18.86 19.34 18.74 19.10 18.68 18.43 18.63 18.46 18.22
19.91 20.14 20.38 19.77 19.31 18.83 19.18 18.74 18.72 18.76 19.39 19.46
20.18 20.02 20.91 20.92 21.11 20.98 20.64 20.65 21.02 20.92 20.42 20.03
12.80 12.17 12.11 11.47 11.27 11.18 11.07 11.09 11.11 11.22 11.18 10.86
20.53 20.47 20.74 20.51 20.45 20.28 19.90 20.05 20.11 19.99 19.77 19.66
17.92 18.00 18.36 18.36 18.21 18.34 18.06 18.19 18.33 18.17 17.60 17.65
15.27 14.96 15.13 15.32 15.13 15.03 14.69 15.04 15.42 15.26 14.81 14.80
17.09 17.02 17.28 17.12 17.14 17.10 16.77 16.88 17.04 17.24 16.89 16.85
19.47 19.67 19.73 19.36 19.27 18.99 18.82 18.78 18.93 19.51 19.45 19.17
19.85 19.98 20.65 20.39 20.21 20.01 19.82 20.02 19.84 19.79 19.00 19.23
17.45 16.97 16.62 16.65 16.63 16.04 16.04 16.09 16.22 16.03 15.72 15.57
19.73 19.51 19.28 18.58 18.58 18.32 18.44 17.77 17.60 17.73 17.67 17.55""",
"D7":"""2.87 2.87 3.06 2.93 2.96 2.85 2.72 2.83 2.98 3.01 2.98 2.89
2.26 2.26 2.40 2.29 2.29 2.19 2.09 2.13 2.32 2.42 2.39 2.34
4.07 3.94 4.31 4.17 4.28 4.11 3.93 3.98 4.28 4.40 4.38 4.23
5.96 6.16 6.17 5.68 5.84 5.63 5.53 5.60 5.85 5.97 5.87 5.68
6.83 7.14 7.01 6.75 6.92 6.65 6.62 6.54 6.61 6.56 6.56 6.45
5.53 5.59 5.76 5.45 5.52 5.34 5.38 5.19 5.43 5.40 5.71 5.76
3.14 3.18 3.46 3.34 3.44 3.35 3.17 3.19 3.45 3.43 3.41 3.26
1.87 1.82 1.86 1.72 1.59 1.56 1.47 1.49 1.70 1.67 1.69 1.57
4.51 4.53 4.79 4.54 4.65 4.50 4.27 4.34 4.65 4.66 4.56 4.43
2.86 2.93 3.13 3.07 3.09 2.98 2.86 2.84 3.10 3.19 3.06 3.02
2.05 2.02 2.08 2.01 2.07 2.03 1.96 2.00 2.15 2.23 2.15 2.12
2.30 2.31 2.51 2.37 2.40 2.33 2.22 2.27 2.47 2.56 2.48 2.45
3.48 3.47 3.68 3.60 3.64 3.46 3.25 3.31 3.65 3.83 3.94 3.68
3.18 3.23 3.49 3.34 3.38 3.35 3.28 3.29 3.48 3.58 3.40 3.38
3.56 3.42 3.47 3.30 3.37 3.15 3.03 3.08 3.29 3.35 3.22 3.14
4.67 4.59 4.44 4.04 4.10 3.94 3.95 3.77 4.10 4.21 4.22 4.17""",
"D28":"""0.61 0.65 0.66 0.64 0.65 0.62 0.59 0.61 0.60 0.66 0.66 0.65
0.47 0.49 0.50 0.48 0.50 0.46 0.43 0.44 0.46 0.51 0.51 0.53
0.93 0.95 0.99 0.96 1.02 0.98 0.89 0.92 0.96 1.03 1.04 1.06
2.23 2.30 2.49 2.09 2.11 1.97 2.01 2.16 2.11 2.28 2.13 2.18
2.85 3.23 3.24 2.95 3.06 2.94 2.85 2.87 3.08 3.06 2.96 2.91
1.95 2.11 2.19 1.91 1.98 1.82 1.74 1.76 1.96 1.79 1.90 1.94
0.65 0.69 0.73 0.71 0.75 0.73 0.67 0.67 0.70 0.74 0.73 0.71
0.51 0.55 0.52 0.49 0.46 0.45 0.41 0.40 0.45 0.48 0.47 0.47
1.19 1.21 1.26 1.21 1.26 1.18 1.09 1.11 1.17 1.22 1.19 1.19
0.60 0.63 0.65 0.69 0.69 0.65 0.62 0.61 0.64 0.71 0.67 0.70
0.42 0.43 0.44 0.44 0.44 0.42 0.39 0.41 0.43 0.47 0.46 0.44
0.47 0.50 0.52 0.50 0.52 0.50 0.45 0.47 0.49 0.53 0.53 0.53
0.77 0.83 0.83 0.85 0.86 0.80 0.72 0.73 0.77 0.85 0.89 0.88
0.68 0.68 0.72 0.70 0.73 0.71 0.66 0.71 0.73 0.76 0.77 0.72
0.84 0.93 0.89 0.81 0.83 0.80 0.80 0.74 0.77 0.78 0.82 0.83
1.23 1.28 1.23 1.09 1.10 1.04 1.01 1.00 1.06 1.07 1.16 1.20""",
"PT":"""14.0 13.9 13.7 13.8 13.9 14.0 14.3 14.4 14.4 14.4 14.1 14.1
16.3 15.8 15.9 16.2 16.2 16.3 16.6 16.7 16.6 16.6 16.8 16.8
12.4 12.2 12.3 12.5 12.8 12.9 13.3 13.4 13.3 13.1 13.0 12.9
40.5 40.4 40.0 40.4 41.6 41.4 41.5 41.7 41.8 41.4 41.2 40.8
47.7 48.4 48.4 49.0 48.5 48.6 48.3 48.2 48.3 48.9 49.1 48.9
35.1 35.6 35.9 34.4 34.3 35.3 34.5 34.5 34.8 35.5 37.1 37.2
13.0 12.7 13.1 13.5 13.8 13.9 14.1 14.1 14.3 14.1 13.7 13.4
25.8 25.3 25.2 24.2 22.3 22.5 22.7 22.6 24.2 24.6 25.6 25.8
29.5 29.6 29.7 30.1 30.4 29.6 29.5 29.7 31.1 31.1 31.1 30.9
13.6 13.5 13.5 13.8 13.9 14.0 14.3 14.5 14.6 14.4 14.5 14.8
15.3 14.9 14.7 15.0 15.1 15.1 15.3 15.8 16.3 16.8 16.9 16.5
14.1 14.1 14.3 14.6 14.6 14.8 14.9 15.0 15.0 15.0 14.9 15.0
15.4 15.3 15.4 15.5 15.5 15.3 15.8 16.0 16.2 16.1 16.0 16.0
22.3 22.2 22.2 22.3 23.5 23.9 24.3 24.8 24.8 24.5 23.7 23.5
13.7 13.9 13.8 13.7 13.8 13.9 14.3 14.2 13.7 13.7 14.0 13.8
24.7 24.3 23.7 24.0 23.8 23.9 23.8 23.9 23.9 23.6 23.1 23.0""",
"SL":"""3.8 3.8 3.8 3.7 3.8 3.8 3.9 3.9 3.9 3.9 4.0 4.0
4.4 4.3 4.5 4.5 4.5 4.4 4.6 4.7 4.8 4.8 4.9 4.9
3.2 3.3 3.3 3.3 3.2 3.3 3.4 3.4 3.3 3.3 3.4 3.4
7.3 7.2 7.2 7.2 7.4 7.2 7.3 7.3 7.3 7.4 7.4 7.3
8.2 8.2 8.2 8.3 8.3 8.2 8.2 8.2 8.4 8.5 8.5 8.5
7.4 7.3 7.3 7.4 7.3 7.3 7.2 7.1 7.2 7.2 7.4 7.5
3.5 3.5 3.5 3.6 3.6 3.6 3.7 3.7 3.7 3.8 3.9 4.0
9.3 9.3 9.3 9.5 9.2 9.3 9.3 9.4 9.4 9.7 9.8 9.7
5.6 5.6 5.6 5.7 5.6 5.6 5.6 5.6 5.7 5.7 5.7 5.7
3.6 3.7 3.7 3.7 3.8 3.8 3.9 3.9 3.9 3.9 3.9 4.0
4.2 4.2 4.1 4.1 4.1 4.1 4.3 4.5 4.7 4.7 4.8 4.7
3.9 4.0 4.0 4.0 4.1 4.1 4.1 4.2 4.2 4.2 4.3 4.3
4.2 4.2 4.2 4.3 4.3 4.3 4.3 4.3 4.4 4.4 4.5 4.4
5.3 5.3 5.3 5.4 5.4 5.5 5.5 5.6 5.6 5.5 5.5 5.5
3.5 3.6 3.6 3.6 3.7 3.7 3.8 3.7 3.6 3.6 3.7 3.7
5.0 5.0 4.9 5.0 5.0 5.1 5.1 5.1 5.1 5.3 5.2 5.2""",
"SC":"""3.54 3.56 3.55 3.59 3.58 3.58 3.59 3.63 3.6 3.55 3.47 3.47
3.65 3.68 3.67 3.76 3.76 3.76 3.75 3.73 3.74 3.66 3.6 3.62
3.82 3.78 3.86 3.97 4.09 4.1 4.1 4.08 4.14 4.08 3.97 3.95
5.08 5.07 5 5.02 5.11 5.08 5.05 5.1 5.16 5.17 5.14 5.08
4.78 4.89 4.86 4.83 4.83 4.83 4.87 4.89 4.88 4.9 4.85 4.91
4.09 4.2 4.28 4.35 4.26 4.3 4.35 4.37 4.26 4.34 4.24 4.23
3.87 3.86 3.95 4.06 4.11 4.1 4.07 4.04 4.13 4.01 3.89 3.87
2.36 2.38 2.41 2.35 2.26 2.31 2.3 2.29 2.33 2.28 2.31 2.26
4.59 4.69 4.74 4.78 4.82 4.77 4.76 4.77 4.9 4.9 4.79 4.8
3.84 3.85 3.89 3.99 3.99 3.94 3.96 4.01 4.02 3.95 3.9 3.97
3.97 3.97 4.02 4.09 4.06 4.08 3.99 3.99 4.04 4.02 3.95 4
3.58 3.61 3.67 3.73 3.72 3.73 3.7 3.71 3.72 3.68 3.62 3.64
3.79 3.79 3.81 3.84 3.79 3.77 3.85 3.85 3.86 3.83 3.74 3.78
4.24 4.21 4.23 4.27 4.32 4.3 4.31 4.35 4.34 4.35 4.28 4.36
3.78 3.79 3.72 3.8 3.8 3.81 3.92 3.94 3.93 3.9 3.84 3.86
4.36 4.36 4.38 4.42 4.37 4.35 4.31 4.22 4.26 4.24 4.27 4.33""",
}
METRICS=[("D1","D1 리텐션 (설치 다음날 다시 온 비율)","17쪽","%"),("D7","D7 리텐션 (7일째 다시 온 비율)","18쪽 위","%"),
("D28","D28 리텐션 (28일째 다시 온 비율)","18쪽 아래","%"),("PT","하루 플레이 시간 (분)","19쪽 위","분"),
("SL","세션 길이 (한 번 켰을 때 노는 시간, 분)","19쪽 아래","분"),("SC","하루 세션 수 (하루에 켜는 횟수, 회)","20쪽","회")]
DATA={}
for k,txt in RD.items():
    rows_=[[float(x) for x in l.split()] for l in txt.strip().split("\n")]
    assert len(rows_)==16 and all(len(r)==12 for r in rows_),k
    DATA[k]=[[v/100 for v in r] for r in rows_] if k in ("D1","D7","D28") else rows_
FMT={"%":"0.00%","분":"0.0","회":"0.00"}
SCALE=ColorScaleRule(start_type="min",start_color="F2F2F2",end_type="max",end_color="8EA9DB")
NOTE=Font(name=F,italic=True,color="595959")
MON=[f"{m}월" for m in range(1,13)]

ws=wb.create_sheet("장르 차트 판독값")
ws["A1"]="GameAnalytics 2025 보고서 · 장르별 차트 판독값 (2024년 데이터, 장르별 중앙값 게임, 월별 코호트)"
ws["A1"].font=Font(name=F,bold=True,size=13)
ws["A2"]="⚠ 차트 판독값: PDF 17~20쪽은 글자가 없는 그림(히트맵)이라, 그림을 확대해 칸 안에 적힌 숫자를 눈으로 옮겨 적었습니다. 공식 표·API 값이 아닙니다. 파란 글씨=판독값, 검은 글씨=판독값으로 계산한 값."
ws["A2"].font=Font(name=F,bold=True,color="C00000")
ws["A3"]="출처: "+GA25+" — 각 표 제목 옆에 PDF 쪽수 표시. 판독 방법: PDF 안 원본 그림(1056×547픽셀)을 2배 확대해 전부 두 번 읽고 대조."
ws["A3"].font=NOTE
BLOCK={}  # metric -> first data row
r=5
for key,title,page,unit in METRICS:
    ws.cell(r,1,f"{title}  —  PDF {page}").font=Font(name=F,bold=True,size=12,color="1F3864")
    r+=1
    heads=["장르(보고서 표기)","장르(한글)"]+MON+["연평균*","최저 달","최고 달","변동폭(최고−최저)","16개 장르 중 순위 (1=가장 높음)"]
    for j,h in enumerate(heads,1): ws.cell(r,j,h)
    for c in ws[r]:
        if c.value: c.fill=H_FILL; c.font=H_FONT; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.row_dimensions[r].height=46
    r+=1; first=r; BLOCK[key]=first; last=first+15
    for gi,(en,ko) in enumerate(GENRES):
        rr=first+gi
        ws.cell(rr,1,en).font=BB; ws.cell(rr,2,ko).font=B
        for m in range(12):
            c=ws.cell(rr,3+m,DATA[key][gi][m]); c.font=BLUE; c.number_format=FMT[unit]
        ws.cell(rr,15,f"=AVERAGE(C{rr}:N{rr})"); ws.cell(rr,16,f"=MIN(C{rr}:N{rr})"); ws.cell(rr,17,f"=MAX(C{rr}:N{rr})")
        ws.cell(rr,18,f"=Q{rr}-P{rr}"); ws.cell(rr,19,f"=RANK(O{rr},$O${first}:$O${last})")
        for j in range(15,19): ws.cell(rr,j).number_format=FMT[unit]
        for j in range(1,20):
            c=ws.cell(rr,j); c.border=BD
            if j>=3: c.alignment=Alignment(horizontal="center")
            if j>=15 and j!=19: c.font=B
        ws.cell(rr,15).font=BB; ws.cell(rr,19).font=BB
    ws.conditional_formatting.add(f"C{first}:N{last}",SCALE)
    ws.conditional_formatting.add(f"O{first}:O{last}",SCALE)
    r=last+2
nr=r
notes5=["* 연평균·최저·최고·변동폭·순위는 보고서에 있는 값이 아니라, 12개월 판독값으로 이 시트가 계산한 값입니다(단순 평균).",
"각 칸은 '그 달에 처음 설치한 사람들(월별 코호트)' 중 장르의 중앙값(가운데 순서) 게임의 수치입니다. 보고서 방법론(5쪽): 장르 장은 중앙값 50% 구간만 사용.",
"차트 숫자는 반올림되어 표시됩니다(리텐션 소수 둘째 자리 %, 시간 소수 첫째 자리 분). 판독값의 정밀도도 그 이상은 아닙니다.",
"D28은 28일째 기준입니다. 2026 보고서(다른 시트)의 D30과 하루 차이가 있어 직접 비교에 주의하세요.",
"판독을 두 번 하면서 처음 읽은 값 5칸을 고쳤습니다(액션 D1 11월, 카드 D1 11월, 퀴즈 D7 7월, 카드 D28 3월, 카지노 플레이 시간 7월). 그 밖의 칸은 두 번 모두 같은 값으로 읽혔습니다."]
for k,t in enumerate(notes5): ws.cell(nr+k,1,t).font=NOTE
ws.column_dimensions["A"].width=18; ws.column_dimensions["B"].width=15
for j in range(3,15): ws.column_dimensions[get_column_letter(j)].width=8.5
for j,w in zip(range(15,20),[10,9,9,11,14]): ws.column_dimensions[get_column_letter(j)].width=w
ws.freeze_panes="C5"
RS="'장르 차트 판독값'!"

# ---------- 장르 판독 요약 ----------
ws=wb.create_sheet("장르 판독 요약")
ws["A1"]="장르별 한눈에 보기 (차트 판독값의 12개월 평균)"; ws["A1"].font=Font(name=F,bold=True,size=13)
ws["A2"]="⚠ 모든 숫자는 '장르 차트 판독값' 시트에서 자동으로 가져온 판독값 기반 계산입니다. 순위: 1위=16개 장르 중 가장 높음."
ws["A2"].font=Font(name=F,bold=True,color="C00000")
H=4
heads=["장르","장르(한글)","D1 리텐션","순위","D7 리텐션","순위","D28 리텐션","순위","하루 플레이(분)","순위","세션 길이(분)","순위","하루 세션 수","순위",
       "쉬운 해설 ① 100명이 설치하면","쉬운 해설 ② 하루 노는 모습"]
for j,h in enumerate(heads,1): ws.cell(H,j,h)
hdr2(ws,H); ws.row_dimensions[H].height=36
S1=H+1; SL_=S1+15
for gi,(en,ko) in enumerate(GENRES):
    rr=S1+gi
    ws.cell(rr,1,en).font=BB; ws.cell(rr,2,ko).font=B
    for k,(key,_,_,unit) in enumerate(METRICS):
        src=BLOCK[key]+gi
        v=ws.cell(rr,3+2*k,f"={RS}O{src}"); v.number_format=FMT[unit]; v.font=B
        ws.cell(rr,4+2*k,f"={RS}S{src}").font=BB
    ws.cell(rr,15,f'="다음날 "&TEXT(C{rr}*100,"0.0")&"명 → 1주 뒤 "&TEXT(E{rr}*100,"0.0")&"명 → 4주 뒤 "&TEXT(G{rr}*100,"0.0")&"명이 다시 옴"')
    ws.cell(rr,16,f'="하루 "&TEXT(M{rr},"0.0")&"번 켜고, 한 번에 "&TEXT(K{rr},"0.0")&"분, 하루 합계 약 "&TEXT(I{rr},"0")&"분"')
    for j in range(1,17):
        c=ws.cell(rr,j); c.border=BD
        if 3<=j<=14: c.alignment=Alignment(horizontal="center")
        if j>=15: c.font=B
for k in range(6):
    col=get_column_letter(3+2*k); ws.conditional_formatting.add(f"{col}{S1}:{col}{SL_}",SCALE)
for col,w in zip("ABCDEFGHIJKLMNOP",[14,14,10,6,10,6,10,6,11,6,11,6,10,6,52,44]): ws.column_dimensions[col].width=w
ws.freeze_panes=f"C{S1}"

# 보고서 문장 vs 판독값 교차 확인
X0=SL_+3
ws.cell(X0,1,"보고서 글(17쪽·4쪽·12쪽)과 판독값이 맞는지 확인 (자동 계산)").font=Font(name=F,bold=True,size=12)
xh=["보고서가 글로 쓴 내용 (요약)","판독값으로 확인한 결과","판정"]
for a_,b_ in [(1,8),(9,14)]: ws.merge_cells(start_row=X0+1,start_column=a_,end_row=X0+1,end_column=b_)
for col_,h in zip((1,9,15),xh):
    c=ws.cell(X0+1,col_,h); c.fill=H_FILL; c.font=H_FONT; c.alignment=Alignment(horizontal="center")
A_=f"$A${S1}:$A${SL_}"
def col(c): return f"${c}${S1}:${c}${SL_}"
def rk(g,c): return f'INDEX({col(c)},MATCH("{g}",{A_},0))'
chk=[
("아케이드는 D1(다음날) 리텐션이 뛰어나다 (17쪽)",
 f'="D1 1위: "&INDEX({A_},MATCH(MAX({col("C")}),{col("C")},0))&" ("&TEXT(MAX({col("C")}),"0.00%")&")"',
 f'=IF(INDEX({A_},MATCH(MAX({col("C")}),{col("C")},0))="Arcade","일치","다름")'),
("아케이드는 오래 붙잡는 힘이 약하다 (17쪽)",
 f'="Arcade 순위: D1 "&{rk("Arcade","D")}&"위 → D7 "&{rk("Arcade","F")}&"위 → D28 "&{rk("Arcade","H")}&"위"',
 f'=IF({rk("Arcade","H")}>{rk("Arcade","D")},"일치 (뒤로 갈수록 순위 하락)","다름")'),
("보드·카드·퍼즐·카지노는 D1·D7·D28 모두 강하다 (17쪽) — D7·D28 확인",
 f'="D7 순위 보드 "&{rk("Board","F")}&"·카드 "&{rk("Card","F")}&"·퍼즐 "&{rk("Puzzle","F")}&"·카지노 "&{rk("Casino","F")}&" / D28 순위 "&{rk("Board","H")}&"·"&{rk("Card","H")}&"·"&{rk("Puzzle","H")}&"·"&{rk("Casino","H")}',
 f'=IF(MAX({rk("Board","F")},{rk("Card","F")},{rk("Puzzle","F")},{rk("Casino","F")},{rk("Board","H")},{rk("Card","H")},{rk("Puzzle","H")},{rk("Casino","H")})<=5,"일치 (모두 5위 안)","부분 일치")'),
("같은 문장 — D1 확인",
 f'="D1 순위 보드 "&{rk("Board","D")}&"·카드 "&{rk("Card","D")}&"·퍼즐 "&{rk("Puzzle","D")}&"·카지노 "&{rk("Casino","D")}',
 f'=IF(MAX({rk("Board","D")},{rk("Card","D")},{rk("Puzzle","D")},{rk("Casino","D")})<=5,"일치 (모두 5위 안)","부분 일치 (D1은 중간권 포함)")'),
("멀티플레이는 리텐션이 가장 낮다 (17쪽)",
 f'="최저 장르: D1 "&INDEX({A_},MATCH(MIN({col("C")}),{col("C")},0))&" / D7 "&INDEX({A_},MATCH(MIN({col("E")}),{col("E")},0))&" / D28 "&INDEX({A_},MATCH(MIN({col("G")}),{col("G")},0))',
 f'=IF(AND(INDEX({A_},MATCH(MIN({col("C")}),{col("C")},0))="Multiplayer",INDEX({A_},MATCH(MIN({col("E")}),{col("E")},0))="Multiplayer",INDEX({A_},MATCH(MIN({col("G")}),{col("G")},0))="Multiplayer"),"일치","부분 일치 (D28 최저는 다른 장르)")'),
("멀티플레이는 세션 길이가 가장 길다 (17쪽)",
 f'="세션 길이 1위: "&INDEX({A_},MATCH(MAX({col("K")}),{col("K")},0))&" ("&TEXT(MAX({col("K")}),"0.0")&"분)"',
 f'=IF(INDEX({A_},MATCH(MAX({col("K")}),{col("K")},0))="Multiplayer","일치","다름")'),
("보드·카드·퍼즐·카지노는 플레이 시간·세션 수도 높다 (17쪽)",
 f'="플레이 시간 순위 "&{rk("Board","J")}&"·"&{rk("Card","J")}&"·"&{rk("Puzzle","J")}&"·"&{rk("Casino","J")}&" / 세션 수 순위 "&{rk("Board","N")}&"·"&{rk("Card","N")}&"·"&{rk("Puzzle","N")}&"·"&{rk("Casino","N")}&" (보드·카드·퍼즐·카지노 순)"',
 f'=IF(MAX({rk("Board","J")},{rk("Card","J")},{rk("Puzzle","J")},{rk("Casino","J")},{rk("Board","N")},{rk("Card","N")},{rk("Puzzle","N")},{rk("Casino","N")})<=4,"일치 (모두 4위 안)","부분 일치")'),
("미드코어 게임은 하루 세션 수 중앙값이 6~7회 (12쪽)",
 f'="16개 장르 중 하루 세션 수 최고: "&INDEX({A_},MATCH(MAX({col("M")}),{col("M")},0))&" "&TEXT(MAX({col("M")}),"0.00")&"회"',
 f'=IF(MAX({col("M")})>=6,"일치","판독값으로는 확인 안 됨 (분류 기준이 다를 가능성, 추정)")'),
]
for k,(a,b,c) in enumerate(chk):
    rr=X0+2+k
    ws.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=8)
    ws.merge_cells(start_row=rr,start_column=9,end_row=rr,end_column=14)
    for col_,v in zip((1,9,15),(a,b,c)):
        cc=ws.cell(rr,col_,v); cc.font=B; cc.border=BD; cc.alignment=Alignment(wrap_text=True,vertical="top")
    ws.cell(rr,15).font=BB
    ws.row_dimensions[rr].height=50
nr=X0+3+len(chk)
notes6=["보고서 본문(4쪽·11~12쪽)의 '전체 게임' 기준값: 플레이 시간 중앙값 약 22분, 세션 길이 중앙값 5~6분, 하루 세션 수 중앙값 4회, D7 중앙값 3.42~3.94%. 장르 표는 '월별 코호트'라 계산 방식이 달라 위 장르 값과 바로 비교하지 마세요.",
"'미드코어 6~7회'는 보고서 12쪽 문장입니다. 21쪽 화면에는 캐주얼·클래식·미드코어라는 별도 장르 묶음이 보이므로, 16개 장르 표와 분류 기준이 다를 수 있습니다(추정).",
"쉬운 해설의 '명'은 판독값 %를 100명 기준으로 바꾼 것입니다(예: 6.72% → 100명 중 6.7명). 소수점 사람은 '1,000명이면 67명'처럼 이해하면 됩니다."]
for k,t in enumerate(notes6): ws.cell(nr+k,1,t).font=NOTE

# ---------- 장르 연결(추정) ----------
ws=wb.create_sheet("장르 연결(추정)")
ws["A1"]="우리 프로젝트 장르 그룹 ↔ GameAnalytics 장르 연결 (분석자 추정)"; ws["A1"].font=Font(name=F,bold=True,size=13)
ws["A2"]="⚠ 연결(C열)은 분석자 추정입니다. 각 게임이 스토어에 어떤 장르로 등록됐는지는 이번에 확인하지 않았습니다. C열(노란 칸)을 목록에서 바꾸면 오른쪽 숫자가 자동으로 바뀝니다."
ws["A2"].font=Font(name=F,bold=True,color="C00000")
H=4
heads=["우리 장르 그룹","대표 게임","연결한 GA 장르 (바꿀 수 있음)","연결 근거","신뢰도","다른 후보 장르",
       "D1","D7","D28","하루 플레이(분)","세션 길이(분)","하루 세션 수","쉬운 해설 (100명 설치 기준)"]
for j,h in enumerate(heads,1): ws.cell(H,j,h)
hdr2(ws,H); ws.row_dimensions[H].height=36
link=[
("1. 글로벌 퍼즐","Block Blast!, Magic Sort!, Meowdoku 등","Puzzle","정렬·블록·스도쿠류 퍼즐이라 이름이 가장 가까움","중간","Board (Vita Mahjong 같은 마작류), Casual"),
("2. 퍼즐 + 메타","Royal Match, Candy Crush Saga 등","Puzzle","3개 맞추기 퍼즐이 핵심 놀이","중간","Casual"),
("3. 소셜 캐주얼","MONOPOLY GO!, Coin Master, Dice Dreams","Board","주사위·보드판 칸 이동이 핵심인 게임이 2개","낮음","Casual, Casino (Coin Master의 슬롯)"),
("4. 전략(4X 등)","Whiteout Survival, Kingshot, Clash Royale","Strategy","장르 이름이 같음","중간","Multiplayer (Clash Royale 실시간 대전)"),
("5. 액션·슈팅","Free Fire, Fortnite, Brawl Stars","Action","장르 이름이 같음","중간","Multiplayer (실시간 대전)"),
("6. 캐릭터 IP (걷기·포획)","Pokémon GO","Adventure","밖을 돌아다니며 탐험·수집","낮음","Simulation"),
("6. 캐릭터 IP (카드)","Pokémon TCG Pocket","Card","카드 수집·대전","중간","—"),
("7. 한국식 RPG","메이플스토리 방치형 RPG, 트릭컬 리바이브, 후더덕 서바이벌","Role Playing","캐릭터 성장·전투가 핵심","중간","Simulation (방치형), Action (서바이벌류)"),
]
L1=H+1
dv=DataValidation(type="list",formula1=f"='장르 판독 요약'!$A${S1}:$A${SL_}",allow_blank=False)
ws.add_data_validation(dv)
SM="'장르 판독 요약'!"
def lk(c,rr): return f'=INDEX({SM}${c}${S1}:${c}${SL_},MATCH($C{rr},{SM}$A${S1}:$A${SL_},0))'
for k,(g,ex,ga,why,conf,alt) in enumerate(link):
    rr=L1+k
    for j,v in enumerate([g,ex,ga,why,conf,alt],1): ws.cell(rr,j,v)
    dv.add(ws.cell(rr,3))
    for j,(c,fmt) in enumerate([("C","0.00%"),("E","0.00%"),("G","0.00%"),("I","0.0"),("K","0.0"),("M","0.00")],7):
        cc=ws.cell(rr,j,lk(c,rr)); cc.number_format=fmt
    ws.cell(rr,13,lk("O",rr))
    for j in range(1,14):
        c=ws.cell(rr,j); c.font=B; c.border=BD; c.alignment=Alignment(wrap_text=True,vertical="center",horizontal="center" if 7<=j<=12 else "left")
    ws.cell(rr,1).font=BB; ws.cell(rr,3).font=BLUE; ws.cell(rr,3).fill=YEL
    if conf=="낮음": ws.cell(rr,5).fill=YEL
    ws.row_dimensions[rr].height=34
LL=L1+len(link)-1
for c in "GHIJKL": ws.conditional_formatting.add(f"{c}{L1}:{c}{LL}",SCALE)
nr=LL+2
notes7=["치이카와 포켓은 게임 방식을 확인하지 못해 연결하지 않았습니다.",
"GameAnalytics 숫자는 이 회사 분석 도구를 쓰는 게임 11,600개(작은 개발사 게임 포함)의 '장르 가운데 게임' 값입니다. Top 차트 게임의 실제 리텐션이 아니므로 'MVP가 최소한 넘어야 할 장르 보통선'으로만 쓰세요.",
"MVP 합격선(전 장르 합산 상위 25%)은 'MVP 검증 목표' 시트에 있고, 이 시트는 그와 별개인 '같은 장르 보통 게임' 참고값입니다.",
"신뢰도: 중간 = 놀이 방식이 장르 이름과 맞음 / 낮음 = 그룹 안 게임들이 서로 다른 장르에 걸쳐 있음."]
for k,t in enumerate(notes7): ws.cell(nr+k,1,t).font=NOTE
for col,w in zip("ABCDEFGHIJKLM",[22,34,16,32,8,30,9,9,9,10,10,10,52]): ws.column_dimensions[col].width=w
ws.freeze_panes=f"C{L1}"

# ===================== 6단계-2: 시장조사 파일 추가 조사 (2026-10-05) =====================
# 새 시트 4개(대륙별 OS 분포 · 동남아 분석 · 매출 Top100 · 매출-다운로드 비교)와
# '지역 비교 요약' E열(동남아), '개요' 맨 아래 2줄을 추가합니다. 기존 칸은 바꾸지 않습니다.
BLUE_IN = Font(name=F, color="0000FF", bold=True)
NOTE = Font(name=F, italic=True, color="595959")
TITLE = Font(name=F, bold=True, size=12)
def M(s):
    """AppBrain 표기(1.7 B, 130 M, 740 K)를 숫자로 바꿈 (AppBrain 추정치 그대로 옮김)"""
    s = s.strip()
    if s in ("0", ""): return 0
    n, u = s[:-1].strip(), s[-1]
    return int(round(float(n) * {"B": 1e9, "M": 1e6, "K": 1e3}[u]))
def head_row(ws, r, vals, c0=1):
    for i, v in enumerate(vals):
        c = ws.cell(r, c0 + i, v); c.fill = H_FILL; c.font = H_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BD
def put(ws, r, vals, c0=1, wrap=False):
    for i, v in enumerate(vals):
        c = ws.cell(r, c0 + i, v); c.font = B; c.border = BD
        c.alignment = Alignment(vertical="top", wrap_text=wrap)
    return r
def yel_row(ws, r, c0, c1):
    for c in range(c0, c1 + 1): ws.cell(r, c).fill = YEL
CATS = ["퍼즐","하이퍼·하이브리드 캐주얼","카드·보드·카지노","시뮬레이션","전략","액션·슈팅","스포츠","RPG","파티·음악","기타(UGC·리워드)","확인 불가"]
SRC_AB_GROSS = "AppBrain Google Play Top Grossing Games United States https://www.appbrain.com/stats/google-play-rankings/top_grossing/game/us"
SRC_AB_FREE = "AppBrain Google Play Top Free Games United States https://www.appbrain.com/stats/google-play-rankings/top_free/game/us"
SRC_AB_ID = "AppBrain Google Play Top Free Games Indonesia https://www.appbrain.com/stats/google-play-rankings/top_free/game/id"
SC = "https://gs.statcounter.com/os-market-share/mobile/"

# ---------- (1) 매출 Top100 : 미국 구글 플레이 Top Grossing (AppBrain, 2026-10-04) ----------
# 순위|게임명|퍼블리셔|대분류|세부 장르|비고|AppBrain 추정 누적 설치
gross = """1|MONOPOLY GO!|Scopely|카드·보드·카지노|보드·주사위 소셜캐주얼||130 M
2|Roblox|Roblox Corporation|기타(UGC·리워드)|UGC 플랫폼||1.7 B
3|Candy Crush Saga|King|퍼즐|매치3||2.3 B
4|Royal Match|Dream Games|퍼즐|매치3||370 M
5|Gossip Harbor: Merge & Story|Microfun|퍼즐|머지|스토리 결합|91 M
6|Township|Playrix|시뮬레이션|농장·타운 빌딩||520 M
7|Coin Master|Moon Active|카드·보드·카지노|보드·주사위 소셜캐주얼||350 M
8|Kingshot|Century Games|전략|4X·SLG||59 M
9|Royal Kingdom|Dream Games|퍼즐|매치3||110 M
10|Whiteout Survival|Century Games|전략|4X·SLG||100 M
11|Pokémon GO|Scopely Explore|RPG|위치기반 수집||560 M
12|Last War:Survival Game|FUNFLY|전략|4X·SLG||110 M
13|DRAGON BALL Z DOKKAN BATTLE|Bandai Namco|RPG|수집형(가챠) RPG|IP 활용|93 M
14|Toon Blast|Peak|퍼즐|매치3|블라스트|130 M
15|Total Battle: War Strategy|Scorewarrior|전략|4X·SLG||72 M
16|Last Z: Survival Shooter|Omnilojo|전략|4X·SLG||41 M
17|Evony: The King's Return|TG Inc.|전략|4X·SLG||220 M
18|RAID: Shadow Legends|Plarium|RPG|수집형(가챠) RPG||77 M
19|Gardenscapes|Playrix|퍼즐|매치3||610 M
20|Dice Dreams|SuperPlay|카드·보드·카지노|보드·주사위 소셜캐주얼||67 M
21|Bingo Blitz - Bingo Games|Playtika Santa Monica|카드·보드·카지노|빙고||83 M
22|Jackpot Party Casino Slots|SciPlay|카드·보드·카지노|소셜 카지노||36 M
23|Last Asylum: Plague|37GAMES GLOBAL|전략|4X·SLG|확인 필요|18 M
24|Fishdom|Playrix|퍼즐|매치3||470 M
25|Cashman Casino Slots Games|Product Madness|카드·보드·카지노|소셜 카지노||22 M
26|확인 불가 (AppBrain 목록에 26위 없음)|-|확인 불가|-|출처 페이지 누락|
27|Match Factory!|Peak|퍼즐|타일·마작·트리플 매치|3D 트리플 매치|29 M
28|Free Fire x NARUTO SHIPPUDEN|Garena|액션·슈팅|배틀로얄||2.2 B
29|Pokémon TCG Pocket|The Pokémon Company|카드·보드·카지노|트레이딩 카드|디지털 TCG|69 M
30|Tasty Travels: Merge Game|Century Games|퍼즐|머지||27 M
31|Clash of Clans|Supercell|전략|4X·SLG|기지 건설·공성|880 M
32|Candy Crush Soda Saga|King|퍼즐|매치3||630 M
33|Lightning Link Casino Slots|Product Madness|카드·보드·카지노|소셜 카지노||15 M
34|Sword x Staff|BOLTRAY GAMES|RPG|방치형 RPG|확인 필요|2.6 M
35|Pixel Flow!|Loom Games|퍼즐|픽셀·컬러 아트||13 M
36|Lotsa Slots - Casino Games|SpinX Games|카드·보드·카지노|소셜 카지노||34 M
37|Homescapes|Playrix|퍼즐|매치3||520 M
38|Dark War Survival|Omnilojo|전략|4X·SLG||33 M
39|Yarn Loop|Combo Game|퍼즐|정렬·잼|확인 필요|3.4 M
40|Honkai: Star Rail|COGNOSPHERE|RPG|수집형(가챠) RPG||37 M
41|Solitaire Grand Harvest|Supertreat (Playtika)|카드·보드·카지노|솔리테어||69 M
42|Wuthering Waves|Kuro Games|RPG|수집형 오픈월드||18 M
43|Disney Solitaire|SuperPlay|카드·보드·카지노|솔리테어|IP 활용|19 M
44|Coin Master - Board Adventure|Moon Active|카드·보드·카지노|보드·주사위 소셜캐주얼|확인 필요|7.3 M
45|Mystery Town: Merge Games|Cedar Games Studio|퍼즐|머지||5.8 M
46|Cash Frenzy - Casino Slots|SpinX Games|카드·보드·카지노|소셜 카지노||32 M
47|Puzzles & Survival|37GAMES|전략|4X·SLG|매치3 결합|41 M
48|All in Hole: Black Hole Games|Homa|하이퍼·하이브리드 캐주얼|하이브리드 캐주얼|확인 필요|12 M
49|Z Route: Redemption|37GAMES GLOBAL|전략|4X·SLG|확인 필요|9.3 M
50|Slots: Heart of Vegas Casino|Product Madness|카드·보드·카지노|소셜 카지노||25 M
51|Flambé: Merge & Cook|Microfun|퍼즐|머지|요리 테마|7.6 M
52|Merge Cooking|Happibits Game|퍼즐|머지|요리 테마|23 M
53|Matching Story - Puzzle Games|VERTEX GAMES|퍼즐|타일·마작·트리플 매치|확인 필요|11 M
54|Hay Day|Supercell|시뮬레이션|농장·타운 빌딩||370 M
55|June's Journey: Hidden Objects|Wooga|퍼즐|숨은그림|스토리 결합|78 M
56|Magic: The Gathering Arena|Wizards of the Coast|카드·보드·카지노|트레이딩 카드||8.4 M
57|Magic Sort!|Grand Games|퍼즐|정렬·잼||40 M
58|Quick Hit Casino Slots Games|SciPlay|카드·보드·카지노|소셜 카지노||23 M
59|Travel Town - Merge Adventure|Magmatic Games|퍼즐|머지||50 M
60|PUBG MOBILE|Level Infinite|액션·슈팅|배틀로얄||1.1 B
61|Puzzles & Chaos: Frozen Castle|37GAMES|전략|4X·SLG|매치3 결합|16 M
62|GODDESS OF VICTORY: NIKKE|Level Infinite|RPG|수집형(가챠) RPG|슈팅|11 M
63|Bingo Voyage - Live Bingo Game|VERTEX GAMES|카드·보드·카지노|빙고||6.6 M
64|Cube Land Puzzle Game|Rotatelab|퍼즐|블록||3 M
65|확인 불가 (AppBrain 목록에 65위 없음)|-|확인 불가|-|출처 페이지 누락|
66|Colony Flow!|ABI GAMES|퍼즐|정렬·잼|확인 필요|3.5 M
67|Brawl Stars|Supercell|액션·슈팅|팀 대전 슈터||550 M
68|Lordrush|Century Games Innovation|전략|4X·SLG|확인 필요|2.7 M
69|eFootball|KONAMI|스포츠|축구||400 M
70|Royal Smash! - Physics Puzzle|Cypher Games|하이퍼·하이브리드 캐주얼|물리 파괴||10 M
71|Minecraft: Dream it, Build it!|Mojang|시뮬레이션|샌드박스|유료 앱|98 M
72|Hollywood Merge|VoyagerOne IE|퍼즐|머지|확인 필요|8.1 M
73|Piggy Kingdom|olleyo|확인 불가|-|게임 방식 확인 안 됨|4.6 M
74|MARVEL Strike Force: Squad RPG|Scopely|RPG|수집형(가챠) RPG|IP 활용|65 M
75|Golf Clash - Golfing Simulator|Electronic Arts|스포츠|골프||45 M
76|Bingo Frenzy-Live Bingo Games|VERTEX GAMES|카드·보드·카지노|빙고||13 M
77|Clash of Critters|FARLIGHT|확인 불가|-|게임 방식 확인 안 됨|6 M
78|WSOP Poker: Texas Holdem Game|Playtika|카드·보드·카지노|포커||83 M
79|Toy Blast|Peak|퍼즐|매치3|블라스트|110 M
80|Jackpot World - Slots Casino|SpinX Games|카드·보드·카지노|소셜 카지노||38 M
81|Yalla Ludo - Ludo&Jackaroo|Aviva Sun|카드·보드·카지노|클래식 보드|루도|300 M
82|Hole Stars: Puzzle Game|Moon Active|퍼즐|기타 퍼즐|홀 수집|5.4 M
83|Hunting Sniper|Sparks Info|액션·슈팅|스나이퍼|헌팅|6.8 M
84|Color Block Jam|Rollic Games|퍼즐|정렬·잼||39 M
85|Star Wars: Galaxy of Heroes|Electronic Arts|RPG|수집형(가챠) RPG|IP 활용|75 M
86|Tiki Solitaire TriPeaks|Scopely|카드·보드·카지노|솔리테어||29 M
87|Fate/Grand Order (English)|Aniplex|RPG|수집형(가챠) RPG|IP 활용|5.4 M
88|Yahtzee With Buddies: Dice|Scopely|카드·보드·카지노|보드·주사위 소셜캐주얼||23 M
89|Hero Wars: Alliance RPG Legend|NEXTERS GLOBAL|RPG|수집형(가챠) RPG||170 M
90|Police Chief|VoyagerOne IE|전략|4X·SLG|확인 필요|3.6 M
91|Bus Traffic Fever!|GOODROID|퍼즐|정렬·잼|버스잼|25 M
92|Rise of Kingdoms: Lost Crusade|LilithGames|전략|4X·SLG||95 M
93|Star Trek Fleet Command|Scopely|전략|4X·SLG|IP 활용|15 M
94|Solitaire Associations Journey|Hitapps Games|퍼즐|워드|연상 퍼즐|35 M
95|Seaside Escape: Merge & Story|Microfun|퍼즐|머지|스토리 결합|18 M
96|Clash Royale|Supercell|전략|실시간 카드 배틀||630 M
97|Loop Sort|VOODOO|퍼즐|정렬·잼||5.5 M
98|Angry Birds 2|Rovio|하이퍼·하이브리드 캐주얼|물리 파괴||350 M
99|Match Masters|Candivore|퍼즐|매치3|PvP|75 M
100|Genshin Impact 6th Anniversary|COGNOSPHERE|RPG|수집형 오픈월드||140 M"""
gross = parse(gross)
assert len(gross) == 100 and [int(r[0]) for r in gross] == list(range(1, 101))

ws = wb.create_sheet("매출 Top100")
ws.append(["순위", "게임명", "퍼블리셔", "대분류", "세부 장르", "비고"])
for r in gross:
    ws.append([int(r[0]), r[1], r[2], r[3], r[4], r[5]])
hdr(ws); body(ws, wrap=False)
for row in ws.iter_rows(min_row=2, max_row=101, max_col=6):
    row[0].alignment = Alignment(horizontal="center")
    if "확인" in str(row[5].value) or row[3].value == "확인 불가":
        for c in row: c.fill = YEL
for col, w in zip("ABCDEF", [6, 38, 26, 24, 24, 22]): ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"
# 장르 집계 (수식) — H열부터
head_row(ws, 1, ["대분류", "매출 Top100 수", "매출 비율", "무료 Top100 수 (1-2)", "무료 비율", "차이 (매출-무료)"], c0=8)
for i, cat in enumerate(CATS, start=2):
    put(ws, i, [cat, f"=COUNTIF($D$2:$D$101,H{i})", f"=I{i}/$I$13",
                f"=COUNTIF('Google Play Top100'!$D$2:$D$101,H{i})", f"=K{i}/$K$13", f"=J{i}-L{i}"], c0=8)
put(ws, 13, ["합계", "=SUM(I2:I12)", "=SUM(J2:J12)", "=SUM(K2:K12)", "=SUM(L2:L12)", ""], c0=8)
for r in range(2, 14):
    for c in (10, 12, 13): ws.cell(r, c).number_format = "0%;-0%;0%"
    ws.cell(r, 8).font = BB
for col, w in zip("HIJKLM", [24, 13, 11, 16, 11, 15]): ws.column_dimensions[col].width = w
notes = ["출처: " + SRC_AB_GROSS + "  ·  기준일 2026-10-04 (AppBrain 페이지 'Last updated: October 4, 2026', 2026-10-05 조회)",
 "26위·65위는 AppBrain 목록에 줄이 없어 '확인 불가'로 둠 (1-2 Google Play Top100의 80위와 같은 현상).",
 "장르 분류: 1-1·1-2에 이미 있는 게임은 같은 분류를 그대로 씀. 새 게임은 게임명·스토어 카테고리 기반 분석자 판단. 확신이 낮으면 비고 '확인 필요' + 노란 칸.",
 "새로 생긴 세부 장르: 수집형(가챠) RPG · 방치형 RPG · 빙고 · 포커 · 골프 (1-4 세부 장르 시트에는 없음 — 1-4는 1-1·1-2만 집계).",
 "'3-1 매출 순위 원자료'의 미국 Top10(2026-10-03)과 순서가 조금 다름 — 하루 차이. 3-1은 그대로 둠.",
 "App Store 매출 순위: 이번 단계 미조사. 후보 공개 출처 — AppBrain App Store ranking https://www.appbrain.com/stats/appstore-rankings (다음 단계 확인 필요).",
 "무료 비율 열은 1-2(같은 날 2026-10-04 미국 무료 Top100)를 수식으로 세어 비교한 것."]
for k, t in enumerate(notes):
    c = ws.cell(15 + k, 8, t); c.font = NOTE
ws.cell(20, 8).fill = YEL

# ---------- (2) 매출-다운로드 비교 ----------
# 미국 무료 Top100 (AppBrain, 2026-10-04 — 1-2와 같은 날 같은 목록) : 순위|게임명|추정 누적 설치
free = """1|Rotate Rings|1.9 M
2|Meowdoku: Brain Puzzle Games|21 M
3|Angry Birds 2|350 M
4|Roblox|1.7 B
5|Vita Mahjong|410 M
6|Smash Land!|740 K
7|Block Blast!|1.1 B
8|Bus Fever Party!|11 M
9|Royal Smash! - Physics Puzzle|10 M
10|MemeMix Challenge: Funny Party|290 K
11|Magic Sort!|40 M
12|Brain Puzzle 3: Crazy Mind|22 M
13|Woodoku Blast|16 M
14|Amaze GO!|110 M
15|Block Craft 3D: Building Game|580 M
16|Word Search Explorer|110 M
17|Gossip Harbor: Merge & Story|91 M
18|Fish Sort Puzzle|2.2 M
19|Mahjong Blast|94 M
20|Pocket Sort: Coin Merge Puzzle|13 M
21|Colony Flow!|3.5 M
22|Township|520 M
23|Arrow Puzzle: Tap Puzzle Games|88 M
24|Loop Sort|5.5 M
25|Fortnite|37 M
26|MONOPOLY GO!|130 M
27|Cube Land Puzzle Game|3 M
28|Castle Busters|19 M
29|Bulldozer Master: Crush & Mine|3.8 M
30|Mob Control|290 M
31|CubeAway - 3DPuzzle|6.4 M
32|Paper.io 2|480 M
33|Magic Tiles 3 - Piano Game|840 M
34|Arrows – Puzzle Escape|120 M
35|Color Block Jam|39 M
36|8 Ball Pool|1.3 B
37|Sled Surfers|22 M
38|Jigsawcard Solitaire Puzzle|17 M
39|Color by Number: Coloring Games|110 M
40|Yarn Loop|3.4 M
41|Warline: Sniper Strike|7.1 M
42|Royal Match|370 M
43|Color Block: Combo Blast|70 M
44|Brainy Escape Quest|12 M
45|CarLoop|750 K
46|Solitaire - Classic Card Games|79 M
47|Royal Kingdom|110 M
48|Candy Crush Saga|2.3 B
49|Pixel Flow!|13 M
50|Slow Mo Strike: Knives in Time|7.5 M
51|Idle Cat Gunner: Shooter RPG|2.8 M
52|Subway Surfers|2.9 B
53|Word Tiles - Relaxing Puzzle|350 K
54|Geometry Dash Lite|620 M
55|Mahjong Master: Daily Match|11 M
56|Search It - Hidden Objects|8.9 M
57|Happy Color: Color by Number|350 M
58|Apex Sniper: Animal Hunt|1.1 M
59|Toca Boca World|460 M
60|Among Us|800 M
61|Ball Sort Puzzle - Color Game|67 M
62|Disney Solitaire|19 M
63|Whiteout Survival|100 M
64|EA SPORTS FC Soccer Mobile 27|800 M
65|Hidden Object Games: Seek It|11 M
66|Aniimo|0
67|Z Route: Redemption|9.3 M
68|Pokémon TCG Pocket|69 M
69|Tiki Smash!|99 K
70|Block Crush!|58 M
71|Word Connect Association|1.7 M
72|Melon Sandbox|120 M
73|Hole Stars: Puzzle Game|5.4 M
74|Double Tile!|13 M
75|Brain Puzzle 2: Logic Twist|38 M
76|Frost Valley: Merge & Story|11 M
77|Solitaire|41 M
78|Food Hunt: Pixel Puzzle|3.8 M
79|Angry Smash|680 K
81|Screw Out 3D: Nut Sort Jam|4.3 M
82|Crossword Go!|3.6 M
83|Match Factory!|29 M
84|Treasure Master|14 M
85|Car Evolve|3.8 M
86|Offline Games - No Wifi Games|210 M
87|Lordrush|2.7 M
88|Loop Master: Color Jam Sort|3.2 M
89|Freddy Playground Sandbox|150 K
90|Wordscapes Search: Word Games|34 M
91|Last Extract|1.5 M
92|Vigor Mahjong|5.2 M
93|Solar Smash|250 M
94|GOGO! Blast|1.8 M
95|Annoying Uncle Punch Game|110 M
96|Food Sort: Puzzle Game|4.4 M
97|Free Fire x NARUTO SHIPPUDEN|2.2 B
98|Dice Dreams|67 M
99|Fast Cash Farkle|77 K
100|Bloons Blitz|200 K"""
free = parse(free)
assert len(free) == 99
free_rank = {r[1]: int(r[0]) for r in free}
inst = {r[1]: M(r[2]) for r in free}
for r in gross:
    if r[6]: inst[r[1]] = M(r[6])
# Google Play 상세 페이지에서 직접 확인한 값 (그 외는 미조사)
gp_detail = {"Kingshot": ("50,000,000+", "2025-02 (Google Play 등록 월)",
             "AppBrain 앱 상세 https://www.appbrain.com/app/kingshot/com.run.tower.defense (페이지 갱신 2026-09-20)")}
rows33 = []
for r in gross:
    if r[3] == "확인 불가" and r[2] == "-": continue          # 순위 누락 줄
    rows33.append((r[1], int(r[0]), free_rank.get(r[1], "-")))
gnames = {r[1] for r in gross}
for r in free:
    if r[1] not in gnames: rows33.append((r[1], "-", int(r[0])))

ws = wb.create_sheet("매출-다운로드 비교")
ws["A1"] = "미국 구글 플레이: 매출 순위 vs 무료(다운로드) 순위 — 같은 날(2026-10-04) 두 차트를 합친 목록"; ws["A1"].font = TITLE
ws["A2"] = "매출 '높음' 기준 (매출 순위가 이 숫자 이내)"; ws["B2"] = 30
ws["A3"] = "다운로드 '높음' 기준 (무료 순위가 이 숫자 이내)"; ws["B3"] = 30
for a in ("A2", "A3"): ws[a].font = BB
for b in ("B2", "B3"):
    ws[b].font = BLUE_IN; ws[b].fill = YEL; ws[b].border = BD; ws[b].alignment = Alignment(horizontal="center")
ws["C2"] = "← 파란 글씨 = 바꿀 수 있는 가정값(분석자 제안). 바꾸면 H열 분류가 자동으로 바뀜"; ws["C2"].font = NOTE
H0 = 5
head_row(ws, H0, ["게임명", "매출 순위", "무료 순위", "AppBrain 추정 누적 설치 수 (추정)", "추정 설치 구간 (수식)",
                  "Google Play 공식 설치 구간", "출시일 (Google Play 등록)", "분류 (수식)", "비고"])
r = H0 + 1
for name, gr, fr in rows33:
    gp_b, rel, note = gp_detail.get(name, ("미조사", "미조사", ""))
    put(ws, r, [name, gr, fr, inst.get(name, ""), None, gp_b, rel, None, note])
    ws.cell(r, 5).value = (f'=IF(D{r}="","",IF(D{r}>=1E9,"1B+",IF(D{r}>=5E8,"500M+",IF(D{r}>=1E8,"100M+",'
                           f'IF(D{r}>=5E7,"50M+",IF(D{r}>=1E7,"10M+",IF(D{r}>=5E6,"5M+",IF(D{r}>=1E6,"1M+","1M 미만"))))))))')
    ws.cell(r, 8).value = (f'=IF(AND(ISNUMBER(B{r}),B{r}<=$B$2,ISNUMBER(C{r}),C{r}<=$B$3),"① 둘 다 높음",'
                           f'IF(AND(ISNUMBER(B{r}),B{r}<=$B$2),"② 매출만 높음",'
                           f'IF(AND(ISNUMBER(C{r}),C{r}<=$B$3),"③ 다운로드만 높음","④ 둘 다 기준 밖")))')
    ws.cell(r, 4).number_format = "#,##0"
    for c in (2, 3): ws.cell(r, c).alignment = Alignment(horizontal="center")
    for c in (4, 5): ws.cell(r, c).fill = YEL
    if gp_b == "미조사":
        for c in (6, 7): ws.cell(r, c).fill = YEL
    r += 1
L33 = r - 1
for col, w in zip("ABCDEFGHI", [38, 10, 10, 16, 13, 16, 18, 18, 40]): ws.column_dimensions[col].width = w
ws.freeze_panes = f"B{H0 + 1}"
ws.auto_filter.ref = f"A{H0}:I{L33}"
# 요약 (수식)
head_row(ws, H0, ["분류", "게임 수", "쉬운 뜻"], c0=11)
expl = {"① 둘 다 높음": "많이 받고 돈도 많이 버는 게임",
        "② 매출만 높음": "받는 사람은 적어도 결제를 많이 하는 게임 (인앱 결제형)",
        "③ 다운로드만 높음": "많이 받지만 결제 순위는 낮은 게임 — 광고 수익형일 가능성 (추정)",
        "④ 둘 다 기준 밖": "두 차트 Top100 안이지만 기준 순위 밖"}
for i, (k, v) in enumerate(expl.items()):
    rr = H0 + 1 + i
    put(ws, rr, [k, f'=COUNTIF($H${H0+1}:$H${L33},K{rr})', v], c0=11, wrap=True)
put(ws, H0 + 5, ["합계", f"=SUM(L{H0+1}:L{H0+4})", ""], c0=11)
ws.cell(H0 + 3, 13).fill = YEL
for col, w in zip("KLM", [18, 9, 44]): ws.column_dimensions[col].width = w
n33 = ["출처: 매출 순위 — " + SRC_AB_GROSS + " / 무료 순위 — " + SRC_AB_FREE + " (둘 다 2026-10-04, 1-2와 같은 날 같은 목록)",
 "'-' = 그 차트 Top100 밖(또는 AppBrain 목록 누락: 매출 26·65위, 무료 80위). 순위 숫자만 비교함.",
 "D열 '추정 누적 설치 수'는 AppBrain이 공개한 추정치(Installs 열)이며 Google Play 공식 숫자가 아님. E열은 그 추정치를 Google Play식 구간으로 나눈 것(역시 추정).",
 "F·G열(공식 설치 구간·출시일)은 게임별 상세 페이지를 하나씩 열어야 해서 이번에는 Kingshot 1개만 확인. 나머지는 '미조사'(다음 단계).",
 "실제 매출액·실제 다운로드 수: 비공개 (공개 출처 없음). 이 표는 순위와 공개 추정치로만 비교함.",
 "③ '광고 수익형'은 순위 패턴에서 나온 추정이며, 실제 수익 구조는 게임별 확인이 필요함."]
for k, t in enumerate(n33):
    c = ws.cell(H0 + 7 + k, 11, t); c.font = NOTE

# ---------- (3) 동남아 분석 ----------
ws = wb.create_sheet("동남아 분석")
ws["A1"] = "① 인도네시아 구글 플레이 무료 게임 Top20"; ws["A1"].font = TITLE
head_row(ws, 2, ["지역", "국가", "기준일", "순위", "게임명", "퍼블리셔", "대분류", "출처", "비고"])
idn = [("Roblox","Roblox Corporation","기타(UGC·리워드)",""),
("Mobile Legends: Bang Bang","MOONTON","액션·슈팅","MOBA — 기존 분류표에 MOBA 칸이 없어 '액션·슈팅'에 넣음(분석자 판단)"),
("Block Blast!","HungryStudio","퍼즐",""),("Free Fire x NARUTO SHIPPUDEN","Garena","액션·슈팅","배틀로얄"),
("Easy Drum - Real Music Drum","AVA GAMES","파티·음악","확인 필요 — 악기 연주 앱"),
("Brain Puzzle 2: Logic Twist","JoyGame Studio","퍼즐",""),("Color Block: Combo Blast","Ivy","퍼즐",""),
("Fruit Wins","Buddyplay Limited","확인 필요","확인 필요 — 게임 방식 확인 안 됨"),
("Brain Puzzle 3: Crazy Mind","JoyGame Studio","퍼즐",""),("Gossip Harbor: Merge & Story","Microfun","퍼즐","머지"),
("Water Go","MicroSurge","확인 필요","확인 필요 — 게임 방식 확인 안 됨"),
("Stickman Party 234 MiniGames","PlayMax Game Studio","파티·음악","로컬 멀티 미니게임"),
("Bus Fever Party!","Lumi Games","퍼즐","버스잼"),("School Party Craft","Candy Room Games & RabbitCo","시뮬레이션","샌드박스"),
("MemeMix Challenge: Funny Party","Amobear VN Game Global","파티·음악",""),
("My Talking Tom 2: Pet Game","Outfit7","시뮬레이션","가상 펫"),("Football League 2026","MOBILE SOCCER","스포츠","축구"),
("Honor of Kings","Level Infinite","액션·슈팅","MOBA — 위 Mobile Legends와 같은 기준"),
("Super Bear Adventure","Earthkwak Games","하이퍼·하이브리드 캐주얼","3D 플랫포머 (1-2의 플랫포머 분류 기준을 따름)"),
("8 Ball Pool","Miniclip.com","스포츠","당구 PvP")]
for i, (n, p, g, note) in enumerate(idn, 1):
    put(ws, 2 + i, ["동남아", "인도네시아", "2026-10-05", i, n, p, g, SRC_AB_ID, note])
    ws.cell(2 + i, 4).alignment = Alignment(horizontal="center")
    if g == "확인 필요" or note.startswith("확인 필요"):
        yel_row(ws, 2 + i, 1, 9)
ws["A24"] = ("태국·베트남·필리핀·말레이시아 Top20: 확인 불가 — AppBrain 국가 목록(30개국)에 이 4개국이 없음(2026-10-05 확인). "
             "42matters 공개 페이지는 'Last update 16 Feb 2025'에 Top10만 있음, appfigures·FoxData는 로그인/스크립트가 필요해 읽지 못함. 싱가포르는 제외(분석자 제안대로).")
ws["A24"].font = NOTE; ws["A24"].fill = YEL
# ② 장르 집계
ws["A26"] = "② 장르 집계 (Top 20 기준, 수식)"; ws["A26"].font = TITLE
SEA = ["인도네시아", "태국", "베트남", "필리핀", "말레이시아"]
head_row(ws, 27, ["대분류 (Top 20 기준)"] + SEA)
CATS2 = CATS[:-1] + ["확인 불가", "확인 필요"]
for i, cat in enumerate(CATS2, start=28):
    vals = [cat, f"=COUNTIFS($B$3:$B$22,B$27,$G$3:$G$22,$A{i})"] + ["확인 불가"] * 4
    put(ws, i, vals)
    for c in range(3, 7): ws.cell(i, c).fill = YEL
RT = 28 + len(CATS2)
put(ws, RT, ["합계", f"=SUM(B28:B{RT-1})"] + ["확인 불가"] * 4)
for r_ in range(28, RT + 1):
    ws.cell(r_, 1).font = BB
    for c in range(2, 7): ws.cell(r_, c).alignment = Alignment(horizontal="center")
PUZ_ROW = 28
# ③ 공통 인기 게임
R3 = RT + 2
ws.cell(R3, 1, "③ 공통 인기 게임 — 기존 6개국 공통 게임이 인도네시아에서는 몇 위인가").font = TITLE
head_row(ws, R3 + 1, ["게임명", "기존 6개국 등장 수 (2-4에서 가져옴)", "인도네시아 순위 (Top100)", "7개국 등장 수 (수식)"])
id_rank = {"Block Blast!": 3, "Meowdoku": "-", "Magic Sort!": "-", "Bus Fever Party!": 13, "Roblox": 1,
           "Vita Mahjong": "-", "Rotate Rings": "-", "Royal Match": "-"}
for k, (gname, rk) in enumerate(id_rank.items()):
    rr = R3 + 2 + k
    put(ws, rr, [gname, f"='공통 인기 게임'!H{2 + k}", rk, f"=B{rr}+COUNT(C{rr})"])
    for c in (2, 3, 4): ws.cell(rr, c).alignment = Alignment(horizontal="center")
ws.cell(R3 + 10, 1, "인도네시아 순위: " + SRC_AB_ID + " (2026-10-05, Top100 확인). '-' = Top100 밖. 동남아 다른 4개국은 순위 자료가 없어 비교 불가.").font = NOTE
# ④ 지역 특징
R4 = R3 + 12
ws.cell(R4, 1, "④ 지역 특징 — 인도네시아 (근거는 2026-10-05 무료 Top100 순위)").font = TITLE
head_row(ws, R4 + 1, ["특징 (쉬운 말)", "근거 게임 (순위)", "분류"])
feat = [("5대5 팀 대전(MOBA)이 아주 인기", "Mobile Legends: Bang Bang (2), Honor of Kings (18)", "사실(순위)"),
("총싸움 배틀로얄도 상위", "Free Fire x NARUTO (4), Free Fire MAX x NARUTO (23)", "사실(순위)"),
("우리 동네 소재 시뮬레이터가 인기", "Bus Simulator Indonesia (31), Ojol Life: Food Delivery Game (65), Herex Simulator: Bike Racing (87)", "사실(순위)"),
("드럼·피아노 같은 악기·리듬 앱이 많음", "Easy Drum (5), Beat Piano (22), Magic Tiles 3 (38), Real Drum (45), Tiles Hop (55)", "사실(순위)"),
("글로벌 두뇌·블록 퍼즐도 상위권", "Block Blast! (3), Brain Puzzle 2 (6), Color Block (7), Brain Puzzle 3 (9)", "사실(순위)"),
("미국형 결제 퍼즐(매치3·머지)은 Top20에 1개뿐", "Gossip Harbor (10)", "사실(순위)"),
("이름이 '888XH…', '77RT…'인 앱 여러 개가 Top100에 있음", "77RTPabu (49), 888XHSukari (51), 888XHPearl (57), 888XHCrown (58)", "확인 필요 — 무엇을 하는 앱인지 확인 안 됨"),
("시사점: 가벼운 안드로이드 대응 + 친구와 짧게 겨루는 재미 + 현지 소재(오토바이·버스·배달)", "위 특징들", "분석자 제안")]
for k, row_ in enumerate(feat):
    rr = R4 + 2 + k
    put(ws, rr, list(row_), wrap=True)
    if row_[2].startswith("확인") or row_[2] == "분석자 제안":
        yel_row(ws, rr, 1, 3)
# ⑤ 동남아 벤치마크
R5 = R4 + 2 + len(feat) + 1
ws.cell(R5, 1, "⑤ 동남아 벤치마크 — GameAnalytics 2025 보고서 (2024년 데이터, 지역별 중앙값)").font = TITLE
head_row(ws, R5 + 1, ["지표", "동남아", "아시아", "북미", "단위", "출처 (2025 보고서 쪽수)"])
bm = [("D1 리텐션", .1828, .1757, .2048, "%", "p.13 지역별 리텐션 표"),
      ("D7 리텐션", .0333, .0314, .0464, "%", "p.13 지역별 리텐션 표"),
      ("D28 리텐션", .0082, .0075, .0146, "%", "p.13 지역별 리텐션 표 (D30 아님)"),
      ("하루 플레이 시간", 20.85, 22.15, 21.7, "분", "p.14 Playtime 표"),
      ("세션 길이", 5.5, 5.05, 5.9, "분", "p.15 Session length 표"),
      ("하루 세션 수", 4.09, 4.31, 3.67, "회", "p.16 Session count 표")]
for k, (m, a, b_, c_, u, s) in enumerate(bm):
    rr = R5 + 2 + k
    put(ws, rr, [m, a, b_, c_, u, "GameAnalytics 2025 Mobile Gaming Benchmarks " + s])
    for c in (2, 3, 4):
        ws.cell(rr, c).font = Font(name=F, color="0000FF")
        ws.cell(rr, c).number_format = "0.00%" if u == "%" else "0.00"
        ws.cell(rr, c).alignment = Alignment(horizontal="center")
bmn = ["2026 보고서(2025년 데이터): 지역 표에 동남아 없음 — 아프리카·아시아·유럽·중동·북미·중미·남미·오세아니아 8개 지역만 있음 (2026 Mobile & PC Gaming Benchmarks p.15~19, investgame.net 공개 PDF로 확인).",
       "주의: 2025 보고서와 2026 보고서는 같은 지표도 값 차이가 큼(예: 아시아 하루 플레이 시간 2025 보고서 22.15분 vs 2026 보고서 11.79분). 연도끼리 직접 비교하지 말 것.",
       "그래서 '4-2 지역 벤치마크'(2026 보고서 표)와 게임전략 파일 2-2 MVP 검증 목표 수식에는 이 값을 넣지 않고 여기에 따로 둠.",
       "출처: 사용자가 첨부한 GameAnalytics 2025 Mobile Gaming Benchmarks PDF (files.gameindustrylibrary.com) — 값은 표의 글자를 그대로 옮김(판독값 아님)."]
for k, t in enumerate(bmn):
    c = ws.cell(R5 + 9 + k, 1, t); c.font = NOTE
ws.cell(R5 + 9, 1).fill = YEL
for col, w in zip("ABCDEFGHI", [44, 34, 22, 20, 26, 30, 18, 60, 48]): ws.column_dimensions[col].width = w
ws.freeze_panes = "A3"

# ---------- (4) 대륙별 OS 분포 ----------
ws = wb.create_sheet("대륙별 OS 분포")
ws["A1"] = "표 A. 모바일 OS 사용 점유율 — StatCounter (웹사이트 방문(페이지뷰) 기준이라 '기기 대수'와 다를 수 있음)"; ws["A1"].font = TITLE
head_row(ws, 2, ["구분", "지역", "Android", "iOS", "기타 (수식)", "더 많은 OS (수식)", "기준 월", "출처", "비고"])
osrows = [("세계","세계",.6917,.3081,"2026-09","worldwide",""),
("대륙","북미",.4695,.5303,"2026-09","north-america","StatCounter는 멕시코를 북미에 포함"),
("대륙","남미",.8353,.1646,"2026-09","south-america",""),
("대륙","유럽",.6468,.3531,"2026-09","europe",""),
("대륙","아시아",.7847,.2149,"2026-09","asia",""),
("대륙","오세아니아",.3825,.6173,"2026-08","oceania","조회 시점 페이지에 9월 값이 아직 없어 8월 값 사용"),
("대륙","아프리카",.8338,.1658,"2026-09","africa",""),
("국가","미국",.4639,.5359,"2026-09","united-states-of-america",""),
("국가","캐나다",.3513,.6487,"2026-09","canada",""),
("국가","브라질",.7878,.2121,"2026-09","brazil",""),
("국가","멕시코",.7204,.2796,"2026-09","mexico",""),
("국가","일본",.3949,.6051,"2026-09","japan","아시아 평균(안드로이드 78%)과 반대로 아이폰이 더 많음"),
("국가","한국",.6218,.3782,"2026-09","south-korea",""),
("국가","인도네시아",.7916,.2079,"2026-09","indonesia",""),
("국가","태국",.7396,.2601,"2026-09","thailand",""),
("국가","베트남",.5979,.4019,"2026-09","viet-nam",""),
("국가","필리핀",.8653,.1347,"2026-09","philippines",""),
("국가","말레이시아",.5523,.4475,"2026-09","malaysia",""),
("국가(참고)","싱가포르",.5625,.4369,"2026-09","singapore","Top20 조사에서는 제외, OS만 참고")]
todo = []
r = 3
for g_, n, a, i_, mth, slug, note in osrows:
    put(ws, r, [g_, n, a, i_, None, None, mth, SC + slug, note])
    r += 1
for g_, n, slug in todo:
    put(ws, r, [g_, n, "미조사", "미조사", None, None, "", SC + slug, "이번 단계 미조사 — 출처 주소 확인됨, 다음 단계에서 채움"])
    yel_row(ws, r, 3, 9)
    r += 1
LOS = r - 1
for rr in range(3, LOS + 1):
    ws.cell(rr, 5).value = f'=IF(ISNUMBER(C{rr}),1-C{rr}-D{rr},"")'
    ws.cell(rr, 6).value = f'=IF(ISNUMBER(C{rr}),IF(C{rr}>=D{rr},"Android","iOS"),"미조사")'
    for c in (3, 4, 5):
        ws.cell(rr, c).number_format = "0.00%"; ws.cell(rr, c).alignment = Alignment(horizontal="center")
    for c in (3, 4):
        if ws.cell(rr, c).value != "미조사": ws.cell(rr, c).font = Font(name=F, color="0000FF")
ws.cell(8, 9).fill = YEL
# 표 B
RB = LOS + 2
ws.cell(RB, 1, "표 B. 게임 이용자 중 OS별 비율").font = TITLE
head_row(ws, RB + 1, ["구분", "지역", "Android", "iOS", "비고"])
for k, n in enumerate(["세계", "북미", "남미", "유럽", "아시아", "오세아니아", "아프리카"]):
    put(ws, RB + 2 + k, ["", n, "확인 불가", "확인 불가", "공개 자료를 찾지 못함 (표 A의 웹 점유율과 다른 숫자라 대신 쓰지 않음)"])
    yel_row(ws, RB + 2 + k, 3, 5)
# 표 C
RC = RB + 10
ws.cell(RC, 1, "표 C. OS(스토어)별 모바일 게임 매출 — 2025년, 세계").font = TITLE
head_row(ws, RC + 1, ["지역", "App Store 매출 ($B)", "Google Play 매출 ($B)", "App Store 비율 (수식)", "Google Play 비율 (수식)", "출처"])
put(ws, RC + 2, ["세계", 52.5, 30, f"=B{RC+2}/(B{RC+2}+C{RC+2})", f"=C{RC+2}/(B{RC+2}+C{RC+2})",
     "Sensor Tower State of Gaming 2026 (2025년 데이터) — 9to5Mac 보도 https://9to5mac.com/?p=1040750"])
put(ws, RC + 3, ["대륙·국가별", "확인 불가", "확인 불가", "", "", "공개 자료를 찾지 못함"])
yel_row(ws, RC + 3, 2, 3)
for c in (4, 5): ws.cell(RC + 2, c).number_format = "0.0%"
for c in (2, 3): ws.cell(RC + 2, c).font = Font(name=F, color="0000FF")
ws.cell(RC + 4, 1, "Google Play 매출에는 중국 안드로이드 마켓이 포함되지 않음(Google Play만 집계). App Store 매출은 중국 포함.").font = NOTE
# 표 D
RD = RC + 6
ws.cell(RD, 1, "표 D (참고). OS(스토어)별 모바일 게임 다운로드 — 2025년, 세계 (표 A·B·C와 또 다른 숫자)").font = TITLE
head_row(ws, RD + 1, ["지역", "App Store 다운로드 (억 건)", "Google Play 다운로드 (억 건)", "App Store 비율 (수식)", "Google Play 비율 (수식)", "출처"])
put(ws, RD + 2, ["세계", 78, 424, f"=B{RD+2}/(B{RD+2}+C{RD+2})", f"=C{RD+2}/(B{RD+2}+C{RD+2})",
     "Sensor Tower State of Gaming 2026 — 9to5Mac 보도 (App Store 7.8B, Google Play 42.4B)"])
for c in (4, 5): ws.cell(RD + 2, c).number_format = "0.0%"
for c in (2, 3): ws.cell(RD + 2, c).font = Font(name=F, color="0000FF")
# 해설
RE = RD + 4
ws.cell(RE, 1, "OS 출시 우선순위 — 분석자 제안 (사실이 아닌 판단)").font = TITLE
ws.cell(RE + 1, 1, "iOS 먼저(또는 동시) 기준: iOS 점유율이 이 값 이상").font = BB
ws.cell(RE + 1, 3, 0.5); ws.cell(RE + 1, 3).font = BLUE_IN; ws.cell(RE + 1, 3).fill = YEL; ws.cell(RE + 1, 3).number_format = "0%"
ws.cell(RE + 1, 4, "← 가정값(분석자 제안). 바꾸면 아래 판단이 자동으로 바뀜").font = NOTE
head_row(ws, RE + 2, ["지역", "iOS 점유율 (표 A)", "먼저 낼 OS (수식)", "쉬운 이유"])
why = {"북미": "아이폰 쓰는 사람이 절반 넘음, 세계 게임 매출도 App Store가 더 큼(표 C)",
       "오세아니아": "아이폰이 10명 중 6명",
       "유럽": "안드로이드가 많지만 아이폰도 3명 중 1명 — 둘 다 준비 권장",
       "남미": "10명 중 8명 넘게 안드로이드", "아시아": "10명 중 8명 가까이 안드로이드(단, 일본은 아이폰이 더 많음 — 아래 일본 줄)",
       "아프리카": "10명 중 8명 넘게 안드로이드", "인도네시아": "10명 중 8명 가까이 안드로이드"}
src_row = {n: 3 + k for k, (g_, n, *_ ) in enumerate(osrows)}
DEC = ["북미", "오세아니아", "유럽", "남미", "아시아", "아프리카", "미국", "캐나다", "브라질", "멕시코", "일본", "한국",
       "인도네시아", "태국", "베트남", "필리핀", "말레이시아"]
for k, n in enumerate(DEC):
    rr = RE + 3 + k
    w_ = why.get(n) or f'="안드로이드 "&TEXT(C{src_row[n]},"0.0%")&" · 아이폰 "&TEXT(D{src_row[n]},"0.0%")'
    put(ws, rr, [n, f"=D{src_row[n]}", f'=IF(B{rr}>=$C${RE+1},"iOS 먼저(또는 동시)","Android 먼저")', w_], wrap=True)
    ws.cell(rr, 2).number_format = "0.0%"
    yel_row(ws, rr, 4, 4)
summ = ["쉬운 요약(분석자 제안): 2인 팀이라면 ① 안드로이드로 먼저 만들어 동남아·남미에서 작게 시험 출시 → ② 숫자가 좋으면 iOS를 붙여 북미·오세아니아로 넓히는 순서가 부담이 적음.",
        "이유: 안드로이드 사용자가 많은 지역은 '많이 받아 보기'에 좋고, 아이폰이 많은 북미는 '돈을 버는 곳'(세계 게임 매출 중 App Store 약 64%, 표 C).",
        "나라별로 보면: 미국·캐나다·일본은 아이폰이 더 많음(iOS 먼저 또는 동시). 필리핀(안드로이드 86.5%)·브라질·인도네시아·태국은 안드로이드가 압도적. 베트남·말레이시아는 아이폰도 40% 이상이라 둘 다 준비 권장(분석자 제안).",
        "주의: 표 A는 웹 방문 기준, 표 C·D는 세계 합계라서 '대륙·나라별 게임 매출 비율'은 아직 모름(확인 불가)."]
for k, t in enumerate(summ):
    c = ws.cell(RE + 3 + len(DEC) + 1 + k, 1, t); c.font = NOTE; c.fill = YEL
for col, w in zip("ABCDEFGHI", [14, 20, 20, 22, 18, 30, 10, 58, 52]): ws.column_dimensions[col].width = w
ws.freeze_panes = "A3"

# ---------- (5) 지역 비교 요약 E열 (동남아) ----------
ws = wb["지역 비교 요약"]
e_vals = ["동남아 (인도네시아 기준)",
          "모바일 PC방 (분석자 비유)",
          "5대5 대전(MOBA)·배틀로얄, 블록·두뇌 퍼즐, 악기·파티 미니게임",
          "인도네시아 Roblox (2026-10-05)",
          "MOBA 2개(Mobile Legends 2위, Honor of Kings 18위), 현지 시뮬레이터(Bus Simulator Indonesia 31위, Ojol Life 65위), 드럼·피아노 앱",
          "확인 불가 (연령 자료 없음)",
          "안드로이드 79.16% (StatCounter 2026-09, 웹 방문 기준)",
          None,
          "가벼운 안드로이드 대응, 친구와 짧게 겨루는 재미, 현지 소재(오토바이·버스·배달) 결합 — 분석자 제안"]
for k, v in enumerate(e_vals, start=1):
    c = ws.cell(k, 5, v)
    if k == 1:
        c.fill = H_FILL; c.font = H_FONT; c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    else:
        c.font = B; c.border = BD; c.alignment = Alignment(wrap_text=True, vertical="top")
ws.cell(8, 5).value = f'="Top 20 중 "&\'동남아 분석\'!B{PUZ_ROW}&"개 (인도네시아, 수식)"'
for k in (2, 6, 9): ws.cell(k, 5).fill = YEL
ws.column_dimensions["E"].width = 50

# ---------- (6) 개요 맨 아래 2줄 추가 ----------
g = wb["개요"]
nr = g.max_row + 1
for k, (a, b_) in enumerate([
    ("6단계-2 출처 (2026-10-05 조회)", "매출 Top100·매출-다운로드 비교: AppBrain 미국 Top Grossing / Top Free Games (둘 다 2026-10-04). 동남아: AppBrain 인도네시아 Top Free Games (2026-10-05). OS: StatCounter Global Stats 모바일 OS 점유율 (2026-09, 오세아니아 2026-08). 게임 매출·다운로드 스토어 비율: Sensor Tower State of Gaming 2026 (9to5Mac 보도). 동남아 벤치마크: GameAnalytics 2025 보고서 p.13~16, 2026 보고서 p.15~19."),
    ("6단계-2 한계", "태국·베트남·필리핀·말레이시아 Top20은 공개 순위 출처가 없어 확인 불가. OS 값은 대륙 6곳·나라 12곳(싱가포르 참고 포함) 조사. 매출-다운로드 비교의 공식 설치 구간·출시일은 Kingshot만 확인. 위 '지역 비교 한계' 줄의 '동남아 미포함' 문구는 기존 칸이라 그대로 둠.")]):
    g.cell(nr + k, 1, a).font = BB; g.cell(nr + k, 1).border = BD
    c = g.cell(nr + k, 2, b_); c.font = B; c.border = BD; c.alignment = Alignment(wrap_text=True, vertical="top")

# ===================== 6단계-3: 게임전략 파일 — 2인 개발 MVP 판단 · 광고 대비 유입 (2026-10-05) =====================
# 새 시트 2개('2인 개발 MVP 판단', '광고 대비 유입')와
# '세부 장르' 맨 아래 별도 표(매출 Top100의 새 세부 장르 5개), '개요' 맨 아래 2줄을 추가합니다. 기존 칸은 바꾸지 않습니다.
# 새 시트 수식은 [1부]의 옛 시트 이름으로 씀 → [2부]가 게임전략 파일의 '(참조)' 복사본 이름으로 자동 변환
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

OK_FILL = PatternFill("solid", fgColor="E2EFDA")    # 2인 가능
MID_FILL = PatternFill("solid", fgColor="FFF2CC")   # 조건부
NO_FILL = PatternFill("solid", fgColor="FCE4D6")    # 2인 어려움
def blue_in(c, fmt=None):
    c.font = BLUE_IN; c.fill = YEL; c.border = BD; c.alignment = Alignment(horizontal="center", vertical="center")
    if fmt: c.number_format = fmt
def judge_colors(ws, rng, col_letter, first_row):
    for txt, fill in (("2인 가능", OK_FILL), ("조건부 (범위를 줄이면)", MID_FILL), ("2인 어려움", NO_FILL)):
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'${col_letter}{first_row}="{txt}"'], fill=fill))

# ---------- (1) 1-4 세부 장르: 매출 Top100에서 새로 생긴 세부 장르 5개 (사용자 승인, 아래쪽 별도 표) ----------
ws = wb["세부 장르"]
base = ws.max_row
assert ws.cell(base, 1).value == "합계"
r0 = base + 2
ws.cell(r0, 1, "6단계-3 추가 (2026-10-05): '3-2 매출 Top100'에서 새로 생긴 세부 장르 5개. 위 표는 1-1·1-2 무료 차트만 세므로 C~E열은 0이 정상이고, F열은 매출 Top100에서 센 수(수식)입니다.").font = NOTE
head_row(ws, r0 + 1, ["대분류", "세부 장르", "App Store 수", "Google Play 수", "합계", "매출 Top100 수 (3-2)"])
new_subs = [("카드·보드·카지노", "빙고"), ("카드·보드·카지노", "포커"), ("스포츠", "골프"),
            ("RPG", "수집형(가챠) RPG"), ("RPG", "방치형 RPG")]
for m_, s_ in new_subs:
    assert any(r[3] == m_ and r[4] == s_ for r in gross), (m_, s_)
    assert not any(r[3] == m_ and r[4] == s_ for r in ios + gp), (m_, s_)
for k, (m_, s_) in enumerate(new_subs):
    i = r0 + 2 + k
    put(ws, i, [m_, s_,
                f"=COUNTIFS('App Store Top100'!$D$2:$D$101,A{i},'App Store Top100'!$E$2:$E$101,B{i})",
                f"=COUNTIFS('Google Play Top100'!$D$2:$D$101,A{i},'Google Play Top100'!$E$2:$E$101,B{i})",
                f"=C{i}+D{i}",
                f"=COUNTIFS('매출 Top100'!$D$2:$D$101,A{i},'매출 Top100'!$E$2:$E$101,B{i})"])
nl = r0 + 1 + len(new_subs)
put(ws, nl + 1, ["합계", "", f"=SUM(C{r0+2}:C{nl})", f"=SUM(D{r0+2}:D{nl})", f"=SUM(E{r0+2}:E{nl})", f"=SUM(F{r0+2}:F{nl})"])
for c in range(1, 7): ws.cell(nl + 1, c).font = BB
ws.column_dimensions["F"].width = 18

# ---------- (2) 2인 개발 MVP 판단 ----------
ws = wb.create_sheet("2인 개발 MVP 판단")
ws["A1"] = "2인 개발 MVP 판단 — 미국 구글 플레이 매출 Top100 + '③ 다운로드만 높음' 게임 중 2명이 작게 만들 수 있는 후보 찾기"
ws["A1"].font = TITLE
ws["A2"] = ("읽는 법: ① 세부 장르마다 5가지 기준을 1~5점(5점 = 2명이 하기 쉬움)으로 매김 → ② 가중치로 100점 만점 점수 계산 → "
            "③ 게임은 자기 세부 장르의 점수·판정을 받음.  파란 글씨 = 바꿀 수 있는 가정값,  노란 칸 = 분석자 판단·추정")
ws["A2"].font = NOTE
# 가정값
head_row(ws, 4, ["기준", "쉬운 뜻", "5점 (2명이 하기 쉬움)", "1점 (2명이 하기 어려움)", "가중치"])
crit = [("콘텐츠 양", "레벨·맵·캐릭터·이야기를 얼마나 많이 만들어야 하나", "규칙 하나로 레벨을 자동으로 만들 수 있음 (블록·정렬·솔리테어)", "오픈월드·수십 명 캐릭터·긴 이야기", 0.25),
        ("서버 필요", "게임 서버를 따로 만들고 돌려야 하나", "인터넷 없이 혼자 플레이 (저장만 기기에)", "실시간 대전·연맹전·거래 (서버가 꼭 필요)", 0.25),
        ("아트 양", "그림·3D 모델·애니메이션이 얼마나 많이 필요한가", "2D 도형·아이콘 위주", "3D 캐릭터·애니메이션 대량", 0.20),
        ("라이브 운영 부담", "출시 뒤 이벤트·시즌·밸런스 패치를 얼마나 자주 해야 하나", "업데이트가 거의 없어도 됨", "매주 이벤트·시즌 패스·밸런스 조정 필수", 0.20),
        ("심사·규제 부담", "스토어 정책·나이 제한·확률 표시 같은 규칙이 얼마나 까다로운가", "특별한 규제 없음", "도박류(소셜 카지노·리얼머니) 정책·나라별 규제", 0.10)]
for k, (a, b_, c5, c1, w) in enumerate(crit):
    i = 5 + k
    put(ws, i, [a, b_, c5, c1, w], wrap=True)
    ws.cell(i, 1).font = BB
    blue_in(ws.cell(i, 5), "0%")
put(ws, 10, ["가중치 합계 (100%가 되게)", "", "", "", "=SUM(E5:E9)"])
ws.cell(10, 1).font = BB; ws.cell(10, 5).number_format = "0%"
for i, (lab, val) in enumerate([("'2인 가능' 기준 점수 (이 점수 이상)", 70), ("'조건부' 기준 점수 (이 점수 이상)", 50),
                                 ("1점 항목이 하나라도 있으면 '2인 어려움'으로 판정 (예/아니오)", "예")], start=12):
    ws.cell(i, 1, lab).font = BB
    blue_in(ws.cell(i, 5, val))
dv_yn = DataValidation(type="list", formula1='"예,아니오"', allow_blank=False); ws.add_data_validation(dv_yn); dv_yn.add("E14")
ws.cell(12, 6, "← 가중치·기준 점수는 분석자 제안. 바꾸면 아래 점수·판정이 모두 자동으로 바뀜").font = NOTE

# 세부 장르 점수표 (분석자 판단) — (콘텐츠, 서버, 아트, 라이브 운영, 심사·규제), None = 판단 보류
SUB_SCORE = {
 ("퍼즐", "정렬·잼"): ((5, 5, 4, 4, 5), "색·물건을 옮겨 정렬하는 규칙 하나로 레벨을 자동 생성하기 쉬움. 3D 잼류는 아트가 조금 더 듦"),
 ("퍼즐", "블록"): ((5, 5, 5, 4, 5), "끝없는 모드 중심이라 레벨 제작이 거의 없음. 2D 블록 그래픽"),
 ("퍼즐", "타일·마작·트리플 매치"): ((4, 5, 4, 4, 5), "타일(물건) 그림 세트가 필요하지만 배치는 반자동 생성 가능. 3D 물건형은 아트 부담↑"),
 ("퍼즐", "로직·두뇌 트릭"): ((3, 5, 4, 4, 5), "스도쿠류는 자동 생성, 두뇌 트릭류는 판마다 아이디어를 손으로 만들어야 함"),
 ("퍼즐", "탭 탈출"): ((5, 5, 5, 4, 5), "화살표·블록을 탭해 빼내는 규칙. 레벨 자동 생성 쉬움, 2D 위주"),
 ("퍼즐", "워드"): ((4, 5, 5, 4, 5), "단어 데이터가 콘텐츠. 나라·언어마다 단어 목록을 새로 만들어야 함(현지화 부담)"),
 ("퍼즐", "매치3"): ((2, 4, 3, 2, 5), "수백~수천 판을 손으로 맞춰야 하고, 매출 상위작은 꾸미기 메타·주간 이벤트가 큼"),
 ("퍼즐", "머지"): ((2, 4, 2, 2, 5), "합칠 물건 그림이 아주 많고 이야기·이벤트 운영 부담이 큼"),
 ("퍼즐", "픽셀·컬러 아트"): ((4, 5, 3, 4, 5), "그림 자료가 곧 콘텐츠. 그림을 자동 변환하면 양은 줄일 수 있음"),
 ("퍼즐", "숨은그림"): ((2, 5, 1, 3, 5), "판마다 큰 그림을 새로 그려야 해서 아트 부담이 가장 큼"),
 ("퍼즐", "기타 퍼즐"): (None, "'기타'로 묶인 장르라 한 번에 판단 불가 — 게임별 확인 필요"),
 ("하이퍼·하이브리드 캐주얼", "물리 파괴"): ((4, 5, 3, 3, 5), "물리 엔진 손맛 조정이 핵심. 저폴리 3D, 레벨 수는 적어도 됨"),
 ("하이퍼·하이브리드 캐주얼", "하이브리드 캐주얼"): ((3, 5, 3, 3, 5), "간단한 핵심 놀이 + 성장(업그레이드) 메타. 메타 설계 분량이 중간"),
 ("하이퍼·하이브리드 캐주얼", "군중 러너·디펜스"): ((3, 5, 3, 3, 5), "3D 군중 표현·성능 최적화가 필요. 레벨은 반복 구조"),
 ("카드·보드·카지노", "보드·주사위 소셜캐주얼"): ((2, 2, 2, 1, 3), "친구 이벤트·경제를 서버가 관리. 매주 이벤트가 매출의 핵심(라이브 운영 1점)"),
 ("카드·보드·카지노", "소셜 카지노"): ((3, 2, 2, 1, 1), "도박 모사 게임 정책·나이 제한, 매일 이벤트"),
 ("카드·보드·카지노", "빙고"): ((3, 2, 3, 1, 2), "여럿이 함께하는 방·이벤트 운영. 도박 모사로 분류될 수 있음"),
 ("카드·보드·카지노", "포커"): ((4, 1, 4, 2, 1), "실시간 대전 서버 필수, 도박 모사 규제"),
 ("카드·보드·카지노", "솔리테어"): ((5, 5, 4, 4, 5), "카드 배치 자동 생성, 오프라인 가능. 매출형은 맵·수집 메타가 붙음"),
 ("카드·보드·카지노", "트레이딩 카드"): ((1, 1, 2, 1, 3), "카드 그림 수백 장, 대전·거래·확률형 뽑기 서버"),
 ("카드·보드·카지노", "클래식 보드"): ((5, 2, 5, 4, 5), "규칙은 이미 있음. 온라인 대전이 핵심이면 서버 필요(2점) — AI 상대로 줄이면 MVP 가능"),
 ("전략", "4X·SLG"): ((1, 1, 1, 1, 3), "연맹전·서버 시즌 운영, 대규모 콘텐츠"),
 ("전략", "실시간 카드 배틀"): ((2, 1, 2, 1, 3), "실시간 대전 서버와 밸런스 패치가 필수"),
 ("RPG", "수집형(가챠) RPG"): ((1, 2, 1, 1, 3), "캐릭터 대량 제작·확률형 아이템 표시 의무·주간 운영"),
 ("RPG", "방치형 RPG"): ((2, 3, 2, 2, 3), "오프라인 진행은 가능하지만 캐릭터·성장 경제 설계 분량이 큼"),
 ("RPG", "수집형 오픈월드"): ((1, 1, 1, 1, 3), "대형 3D 세계·캐릭터 — 대형 스튜디오 규모"),
 ("RPG", "위치기반 수집"): ((1, 1, 2, 1, 3), "지도 데이터·위치 서버·안전 정책"),
 ("액션·슈팅", "배틀로얄"): ((1, 1, 1, 1, 4), "실시간 대규모 대전 서버, 3D 맵·무기"),
 ("액션·슈팅", "팀 대전 슈터"): ((1, 1, 1, 1, 4), "실시간 대전 서버, 3D 캐릭터·맵"),
 ("액션·슈팅", "스나이퍼"): ((3, 4, 2, 3, 4), "혼자 하는 미션형이면 서버 부담은 적지만 3D 아트가 필요"),
 ("스포츠", "축구"): ((1, 1, 1, 2, 4), "선수·리그 라이선스와 3D 애니메이션"),
 ("스포츠", "골프"): ((3, 2, 2, 2, 5), "물리 손맛 + 코스 3D. 상위작은 1:1 대전(서버)"),
 ("시뮬레이션", "농장·타운 빌딩"): ((1, 3, 1, 1, 5), "건물·아이템 수백 개와 꾸준한 이벤트"),
 ("시뮬레이션", "샌드박스"): ((3, 4, 3, 3, 4), "블록(복셀) 그래픽은 단순하지만 만들기 도구 범위가 넓음"),
 ("기타(UGC·리워드)", "UGC 플랫폼"): ((1, 1, 1, 1, 3), "플랫폼 자체(사용자 제작 도구·서버·심사)라 2인 범위 밖"),
 ("파티·음악", "밈 챌린지"): ((4, 5, 3, 2, 2), "유행 밈을 계속 바꿔야 함(운영), 밈 저작권·초상권 위험"),
 ("확인 불가", "-"): (None, "장르를 확인하지 못한 게임 — 판단 보류"),
}
cand_keys = []
def _gen(name):
    for r in gross:
        if r[1] == name: return (r[3], r[4])
    for r in gp:
        if r[1] == name: return (r[3], r[4])
    return ("확인 불가", "-")
cands = [n for n, g_, f_ in rows33 if isinstance(g_, int) or
         (isinstance(f_, int) and f_ <= 30 and not (isinstance(g_, int) and g_ <= 30))]
from collections import Counter as _C
key_cnt = _C(_gen(n) for n in cands)
missing = [k for k in key_cnt if k not in SUB_SCORE]
assert not missing, missing
sub_rows = sorted(key_cnt, key=lambda k: (-key_cnt[k], k))

BT = 17   # 세부 장르 점수표 머리 줄
ws.cell(BT - 1, 1, "① 세부 장르 점수표 — C~G열 점수는 분석자 판단(노란 칸). 장르의 일반적인 구조 기준이며 게임마다 다를 수 있음").font = TITLE
head_row(ws, BT, ["대분류", "세부 장르", "콘텐츠 양", "서버 필요", "아트 양", "라이브 운영 부담", "심사·규제 부담",
                  "점수 (100점, 수식)", "판정 (수식)", "판단 근거 (분석자)", "찾기용 키 (수식)", "후보 게임 수 (수식)"])
B1, BL = BT + 1, BT + len(sub_rows)
for k, key in enumerate(sub_rows):
    j = B1 + k
    sc, why = SUB_SCORE[key]
    put(ws, j, [key[0], key[1]] + (list(sc) if sc else ["", "", "", "", ""]) + [None, None, why, None, None], wrap=True)
    for c in range(3, 8):
        ws.cell(j, c).fill = YEL; ws.cell(j, c).alignment = Alignment(horizontal="center", vertical="top")
    ws.cell(j, 10).fill = YEL
    ws.cell(j, 8).value = (f'=IF(COUNT(C{j}:G{j})<5,"",ROUND((C{j}*$E$5+D{j}*$E$6+E{j}*$E$7+F{j}*$E$8+G{j}*$E$9)'
                           f'/SUM($E$5:$E$9)/5*100,0))')
    ws.cell(j, 9).value = (f'=IF(H{j}="","판단 보류",IF(AND($E$14="예",MIN(C{j}:G{j})=1),"2인 어려움",'
                           f'IF(H{j}>=$E$12,"2인 가능",IF(H{j}>=$E$13,"조건부 (범위를 줄이면)","2인 어려움"))))')
    ws.cell(j, 11).value = f'=A{j}&"|"&B{j}'
    for c in (8, 9, 12): ws.cell(j, c).alignment = Alignment(horizontal="center", vertical="top")
judge_colors(ws, f"I{B1}:I{BL}", "I", B1)

# 추천 후보 (분석자 제안)
DT = BL + 3
ws.cell(DT - 1, 1, "② 2인 개발 추천 후보 — 게임 선택·역할·기간·꼭 넣을 것·주의점은 모두 분석자 제안/추정 (점수·순위는 수식으로 가져옴)").font = TITLE
head_row(ws, DT, ["게임명", "세부 장르 (수식)", "매출 순위 (수식)", "무료 순위 (수식)", "점수 (수식)", "3-3 유형 (수식)",
                  "판정 (수식)", "2인 역할 (분석자 가정)", "MVP 기간 (분석자 추정)", "MVP에 꼭 넣을 것 (분석자 제안)", "주의점"])
picks = [
 ("Magic Sort!", "개발 1 + 기획·아트 1", "약 2~3개월",
  "물 색 정렬 규칙, 레벨 자동 생성기와 난이도 곡선, 보상형 광고(되돌리기·병 추가)",
  "Liftoff 2025 보고서가 '물 정렬 퍼즐 첫 하이브리드 캐주얼 성공작'으로 소개. 같은 규칙 경쟁작이 많아 차별점 1개가 필요"),
 ("Match Factory!", "개발 1 + 3D 아트·기획 1", "약 3~4개월",
  "3D 물건 3개 모으기 규칙, 물건 모델 세트(에셋 구매 검토), 시간 제한·부스터",
  "3D 물건 아트가 가장 큰 일 — 에셋 구매 여부가 기간을 좌우"),
 ("Cube Land Puzzle Game", "개발 1 + 기획·아트 1", "약 2~3개월",
  "블록 놓기·줄 지우기 규칙, 점수·콤보, 끝없는 모드와 광고",
  "Block Blast!·Woodoku Blast 같은 강자가 무료 상위에 있음(1-2) — 테마·연출 차별화 필요"),
 ("Tiki Solitaire TriPeaks", "개발 1 + 기획·아트 1", "약 2~3개월",
  "트라이픽스 규칙, 카드 배치 자동 생성, 테마 배경 몇 개, 부스터",
  "매출 상위 솔리테어는 맵·수집 메타가 있음 — MVP 뒤에 메타를 붙이는 순서"),
 ("Royal Smash! - Physics Puzzle", "개발 1 + 3D 아트·기획 1", "약 3~4개월",
  "물리 파괴 손맛, 짧은 레벨 묶음, 보상형 광고",
  "물리 튜닝과 3D 저폴리 아트가 개발 부담의 중심"),
 ("Arrow Puzzle: Tap Puzzle Games", "개발 1 + 기획·아트 1", "약 1.5~2.5개월",
  "화살표 방향대로 탭해서 빼내기 규칙, 레벨 생성기, 힌트(광고 보고 받기)",
  "매출 Top100 밖(③ 유형) → 광고 수익형일 가능성(추정). 결제 매출 기대는 낮게"),
 ("Solitaire Associations Journey", "개발 1 + 기획·아트 1", "약 2~3개월",
  "단어 묶음 카드 규칙, 영어 단어 데이터, 하루 한 판 퍼즐",
  "언어마다 단어 데이터를 새로 만들어야 함 — 처음엔 영어(미국)만"),
]
names33 = [n for n, g_, f_ in rows33]
D1_, DL_ = DT + 1, DT + len(picks)
CT = DL_ + 14          # 전체 게임 목록 머리 줄 (아래 C1_~CL_)
C1_, CL_ = CT + 1, CT + len(rows33)
def cl(col): return f"${col}${C1_}:${col}${CL_}"
for k, (nm, role, dur, must, warn) in enumerate(picks):
    assert nm in names33 and nm in cands, nm
    i = D1_ + k
    put(ws, i, [nm, None, None, None, None, None, None, role, dur, must, warn], wrap=True)
    ws.cell(i, 1).font = BB
    for c, src in zip(range(2, 8), ["C", "D", "E", "H", "F", "I"]):
        ws.cell(i, c).value = f'=INDEX({cl(src)},MATCH($A{i},{cl("A")},0))'
        ws.cell(i, c).alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
    for c in (8, 9, 10, 11): ws.cell(i, c).fill = YEL
judge_colors(ws, f"G{D1_}:G{DL_}", "G", D1_)

# 메모
notes24 = [
 "이 표는 '2명이 MVP(최소 기능 제품)를 만들 수 있나'만 봅니다. 돈을 벌 수 있는지는 매출 순위(4-5)·광고 대비 유입(3-1)·마케팅 실행안(3-4, 다음 단계)에서 따로 봅니다.",
 "후보 = 매출 Top100 안 게임 + 3-3 분류가 '③ 다운로드만 높음'인 게임 (G열 수식). '4-6 매출-다운로드 비교(참조)'의 B2·B3 기준값을 바꾸면 ③ 후보도 함께 바뀝니다.",
 "세부 장르는 '4-5 매출 Top100(참조)' → 없으면 '4-4 Google Play Top100(참조)'에서 찾습니다(수식). 세 시트 모두 시장조사.xlsx의 참조용 복사본입니다.",
 "세부 장르 점수(C~G열)·역할·MVP 기간은 분석자 판단/추정이며 근거 수치가 없습니다. 실제 게임별 개발 인원·기간·개발비·매출액: 비공개 (공개 자료로 확인 불가).",
 "공개 자료로 대신 쓸 수 있는 지표: 매출 순위·무료 순위(AppBrain, 2026-10-04), AppBrain 추정 누적 설치 수(3-3 D열, 추정), 스토어 출시일(3-3 G열, 대부분 미조사).",
 "'1점 항목이 있으면 어려움' 규칙(E14)은 실시간 서버·대형 콘텐츠처럼 하나만 있어도 2명이 감당하기 어려운 경우를 거르려는 분석자 제안입니다. '아니오'로 바꾸면 점수만으로 판정합니다.",
 "아래 ③ 전체 목록은 3-3의 172개 게임 전부이며, '후보 아님' 줄은 점수를 매기지 않습니다. 맨 위 머리 줄의 필터로 '후보'만 골라 볼 수 있습니다.",
]
for k, t in enumerate(notes24):
    ws.cell(DL_ + 2 + k, 1, t).font = NOTE

# 요약 (G4~I9)
head_row(ws, 4, ["판정", "후보 게임 수 (수식)", "세부 장르 수 (수식)"], c0=7)
for k, lab in enumerate(["2인 가능", "조건부 (범위를 줄이면)", "2인 어려움", "판단 보류"]):
    i = 5 + k
    put(ws, i, [lab, f'=COUNTIFS({cl("I")},G{i},{cl("G")},"후보")', f'=COUNTIF($I${B1}:$I${BL},G{i})'], c0=7)
put(ws, 9, ["합계", "=SUM(H5:H8)", "=SUM(I5:I8)"], c0=7)
for c in (7, 8, 9): ws.cell(9, c).font = BB
judge_colors(ws, "G5:G8", "G", 5)

# ③ 전체 목록
ws.cell(CT - 1, 1, "③ 전체 게임 목록 (3-3의 172개) — 장르·순위·유형은 참조 시트에서 수식으로 가져옴").font = TITLE
head_row(ws, CT, ["게임명", "대분류 (수식)", "세부 장르 (수식)", "매출 순위 (수식)", "무료 순위 (수식)", "3-3 유형 (수식)",
                  "후보 여부 (수식)", "점수 (수식)", "판정 (수식)", "", "찾기용 키 (수식)"])
MD_ = "'매출-다운로드 비교'!"
m1, mL = H0 + 1, H0 + len(rows33)
GT_, GP_ = "'매출 Top100'!", "'Google Play Top100'!"
for k, (nm, g_, f_) in enumerate(rows33):
    i = C1_ + k
    put(ws, i, [nm])
    ws.cell(i, 2).value = (f'=IFERROR(INDEX({GT_}$D$2:$D$101,MATCH($A{i},{GT_}$B$2:$B$101,0)),'
                           f'IFERROR(INDEX({GP_}$D$2:$D$101,MATCH($A{i},{GP_}$B$2:$B$101,0)),"확인 불가"))')
    ws.cell(i, 3).value = (f'=IFERROR(INDEX({GT_}$E$2:$E$101,MATCH($A{i},{GT_}$B$2:$B$101,0)),'
                           f'IFERROR(INDEX({GP_}$E$2:$E$101,MATCH($A{i},{GP_}$B$2:$B$101,0)),"-"))')
    for c, src in ((4, "B"), (5, "C"), (6, "H")):
        ws.cell(i, c).value = f'=INDEX({MD_}${src}${m1}:${src}${mL},MATCH($A{i},{MD_}$A${m1}:$A${mL},0))'
    ws.cell(i, 7).value = f'=IF(OR(ISNUMBER(D{i}),F{i}="③ 다운로드만 높음"),"후보","후보 아님")'
    ws.cell(i, 8).value = f'=IF(G{i}<>"후보","",IFERROR(INDEX($H${B1}:$H${BL},MATCH(K{i},$K${B1}:$K${BL},0)),""))'
    ws.cell(i, 9).value = f'=IF(G{i}<>"후보","-",IFERROR(INDEX($I${B1}:$I${BL},MATCH(K{i},$K${B1}:$K${BL},0)),"판단 보류"))'
    ws.cell(i, 11).value = f'=B{i}&"|"&C{i}'
    for c in range(2, 12):
        ws.cell(i, c).font = B; ws.cell(i, c).border = BD
    for c in (4, 5, 7, 8, 9): ws.cell(i, c).alignment = Alignment(horizontal="center")
judge_colors(ws, f"I{C1_}:I{CL_}", "I", C1_)
# 후보 게임 수 (세부 장르 표 L열)
for j in range(B1, BL + 1):
    ws.cell(j, 12).value = f'=COUNTIFS({cl("K")},K{j},{cl("G")},"후보")'
ws.auto_filter.ref = f"A{CT}:K{CL_}"
for col, w in zip("ABCDEFGHIJKL", [34, 24, 22, 22, 12, 22, 18, 14, 22, 52, 30, 12]):
    ws.column_dimensions[col].width = w
for i in range(5, 10): ws.row_dimensions[i].height = 32
ws.freeze_panes = "B4"
assert len(cands) == 118, len(cands)

# ---------- (3) 광고 대비 유입 ----------
ws = wb.create_sheet("광고 대비 유입")
ws["A1"] = "광고 대비 유입 — 광고비 → 설치 수 → D1·D7·D30에 다시 오는 사람 수"; ws["A1"].font = TITLE
ws["A2"] = ("읽는 법: 광고비 ÷ CPI(설치 1번에 드는 광고비) = 설치 수 → 설치 수 × 리텐션(며칠 뒤 다시 오는 비율) = 그날 남는 사람 수.  "
            "파란 글씨 = 바꿀 수 있는 가정값,  노란 칸 = 2차 인용·섞인 가정·확인 불가")
ws["A2"].font = NOTE
ws["A4"] = "꼭 읽을 주의점 3가지"; ws["A4"].font = BB
caveats = [
 "① 출처마다 CPI가 10배 넘게 달라요. 예) 같은 '캐주얼 게임 Android'가 Liftoff 원문은 $0.14인데, 한 2026년 블로그(digitalapplied.com)는 $1.12로 적음 — 그 블로그는 여러 보고서를 섞어 원문 확인이 안 돼 쓰지 않음. 이 표는 원문을 확인한 값만 '원문', 블로그 등을 거친 값은 '2차 인용'(노란 칸)으로 나눔.",
 "② 시장 전체 값이 아니에요. Liftoff 보고서의 CPI는 측정 회사 Singular의 데이터(노출 1.1조·설치 24억·광고비 $119억)이고, Adjust 값은 Adjust가 추적하는 앱만의 값이에요(Adjust 스스로 '전체 시장을 반영하지 않을 수 있다'고 밝힘, p.8).",
 "③ '지역 × OS × 장르'를 모두 나눈 공개 값은 없어요. 지역별 값은 Adjust(iOS+Android 합산, 게임 전체), OS·장르별 값은 Liftoff(전 세계)뿐이라, 둘을 섞은 시나리오(S5~S8)는 '섞인 가정'이에요. 블로그에 도는 '미드코어 북미 $5.45·남미 $0.27'은 원문을 찾지 못해 쓰지 않음.",
]
for k, t in enumerate(caveats):
    c = ws.cell(5 + k, 1, t); c.font = B; c.fill = YEL

# CPI 출처표
AT = 10
ws.cell(AT - 1, 1, "① CPI 출처표 (설치 1번에 드는 광고비, 미국 달러)").font = TITLE
head_row(ws, AT, ["ID", "지역", "OS", "장르", "CPI ($)", "데이터 기간", "자료 등급", "비고", "출처 (이름 · URL)"])
ADJ = "Adjust, The gaming app insights report: 2026 edition, p.28~29 (Adjust 추적 앱 집계, 2024-01~2026-01) — 원문 PDF 게시본 https://investgame.net/wp-content/uploads/2026/03/2026-03-26-gaming-app-insights-report-2026_wp.pdf (2026-10-05 조회)"
LFT = "Liftoff·Singular, 2025 Casual Gaming Apps Report — 'CPI by genre by platform' (Singular 데이터 2024-02~2025-02) https://liftoff.ai/2025-casual-gaming-apps-report/ (2026-10-05 조회)"
STA = "Statista, Cost per install (CPI) of gaming apps worldwide Feb 2024–Feb 2025, by genre and platform (출처 표기: Liftoff, 2025-04-29) https://www.statista.com/statistics/1241651/global-cpi-gaming-apps-genre-platform/"
MIS = "Mistplay 블로그, How much does mobile user acquisition cost in 2026? (Liftoff 2025 보고서 인용) https://business.mistplay.com/resources/user-acquisition-cost"
GDR = "GameDev Reports, Liftoff & Singular: Casual Games in 2025 (요약) https://gamedevreports.substack.com/p/liftoff-and-singular-casual-games"
cpi_rows = [
 ("C01", "전 세계", "iOS+Android 합산", "게임 전체", 0.56, "2025년", "원문", "2024년 대비 +30%", ADJ),
 ("C02", "북미", "iOS+Android 합산", "게임 전체", 1.68, "2025년", "원문", "2024년 $1.28", ADJ),
 ("C03", "미국", "iOS+Android 합산", "게임 전체", 1.71, "2025년", "원문", "2024년 $1.31", ADJ),
 ("C04", "남미 (LATAM)", "iOS+Android 합산", "게임 전체", 0.14, "2025년", "원문", "2024년 대비 +40%", ADJ),
 ("C05", "아시아태평양 (APAC)", "iOS+Android 합산", "게임 전체", 0.27, "2025년", "원문", "2024년 $0.21. 일본·한국 등 나라별 값은 그래프에만 있어 확인 불가", ADJ),
 ("C06", "유럽", "iOS+Android 합산", "게임 전체", 0.53, "2025년", "원문", "2024년 $0.36", ADJ),
 ("C07", "전 세계", "iOS+Android 합산", "슬롯", 4.47, "2025년", "원문", "2025년 장르 중 가장 비쌈", ADJ),
 ("C08", "전 세계", "iOS+Android 합산", "방치형 RPG", 3.19, "2025년", "원문", "", ADJ),
 ("C09", "전 세계", "iOS+Android 합산", "전략", 1.03, "2025년", "원문", "", ADJ),
 ("C10", "전 세계", "iOS", "캐주얼 (하이퍼캐주얼 포함)", 1.41, "2024-02~2025-02", "원문", "", LFT),
 ("C11", "전 세계", "Android", "캐주얼 (하이퍼캐주얼 포함)", 0.14, "2024-02~2025-02", "원문", "", LFT),
 ("C12", "전 세계", "iOS", "카지노", 21.03, "2024-02~2025-02", "원문", "", LFT),
 ("C13", "전 세계", "Android", "카지노", 4.10, "2024-02~2025-02", "2차 인용", "원문은 그래프 이미지라 직접 확인 못 함", MIS),
 ("C14", "전 세계", "iOS", "퍼즐", 2.80, "2024-02~2025-02", "2차 인용", "Statista 표의 열 이름이 안 보여, 같은 표 캐주얼 줄(0.14·1.41)이 원문 Android·iOS 값과 같은 것으로 열 순서를 판단", STA),
 ("C15", "전 세계", "Android", "퍼즐", 0.64, "2024-02~2025-02", "2차 인용", "C14와 같은 방법으로 열 순서 판단", STA),
 ("C16", "전 세계", "Android", "RPG", 4.29, "2024-02~2025-02", "2차 인용", "요약 글에만 있음 (원문 그래프 이미지)", GDR),
 ("C17", "북미·남미·아시아 각각", "iOS / Android 각각", "퍼즐 등 장르 각각", "확인 불가", "—", "확인 불가", "세 가지를 모두 나눈 공개 원문 값을 찾지 못함", "—"),
]
A1_, AL_ = AT + 1, AT + len(cpi_rows)
for k, row in enumerate(cpi_rows):
    i = A1_ + k
    put(ws, i, list(row))
    ws.cell(i, 5).number_format = '"$"0.00'
    ws.cell(i, 5).font = BB
    for c in (1, 3, 5, 7): ws.cell(i, c).alignment = Alignment(horizontal="center", vertical="top")
    ws.cell(i, 8).alignment = Alignment(vertical="top", wrap_text=True)
    if row[6] != "원문":
        for c in range(1, 10): ws.cell(i, c).fill = YEL

# 계산표
BT2 = AL_ + 3
ws.cell(BT2 - 1, 1, "② 계산표 — 시나리오 조합·광고비는 분석자 가정(파란 글씨, 바꿔 쓰는 칸). 리텐션은 '4-1 지역 벤치마크(참조)'에서 수식으로 가져옴").font = TITLE
head_row(ws, BT2, ["시나리오", "설명", "광고비 ($)", "CPI 출처 ID (A열에서 선택)", "출처 CPI ($)", "직접 입력 CPI ($) (비우면 출처값)",
                   "쓰는 CPI ($)", "설치 수", "리텐션 지역", "리텐션 구간", "D1 리텐션", "D7 리텐션", "D30 리텐션",
                   "D1에 남는 사람", "D7에 남는 사람", "D30에 남는 사람", "D7에 100명 남기려면 필요한 광고비 ($)"])
scen = [
 ("S1", "미국 · 게임 전체 평균 CPI · 보통 게임(중앙값 P50) 리텐션", "C03", "북미", "P50", False),
 ("S2", "미국 · 게임 전체 평균 CPI · 상위 10%(P90) 리텐션 — 잘 만든 게임 목표", "C03", "북미", "P90", False),
 ("S3", "남미 · 게임 전체 평균 CPI · P50 리텐션", "C04", "남미", "P50", False),
 ("S4", "아시아태평양 CPI · 아시아 P50 리텐션 (APAC과 GameAnalytics '아시아'는 범위가 조금 다름)", "C05", "아시아", "P50", True),
 ("S5", "iOS 캐주얼 CPI(전 세계) + 북미 P50 리텐션 — 섞인 가정", "C10", "북미", "P50", True),
 ("S6", "Android 캐주얼 CPI(전 세계) + 북미 P50 리텐션 — 섞인 가정", "C11", "북미", "P50", True),
 ("S7", "iOS 퍼즐 CPI(전 세계, 2차 인용) + 북미 P50 리텐션 — 섞인 가정", "C14", "북미", "P50", True),
 ("S8", "Android 퍼즐 CPI(전 세계, 2차 인용) + 북미 P50 리텐션 — 섞인 가정", "C15", "북미", "P50", True),
]
RB = "'지역 벤치마크'!"
def ret(i, met):
    rng = lambda col: f"{RB}${col}$2:${col}${REGLAST}"
    s = lambda col: f'SUMIFS({rng(col)},{rng("A")},$I{i},{rng("B")},"{met}")'
    return f'=IF($J{i}="P50",{s("C")},IF($J{i}="P90",{s("D")},IF($J{i}="P99",{s("E")},"")))'
S1_, SL_ = BT2 + 1, BT2 + len(scen)
for k, (sid, desc, cid, reg_, q, mixed) in enumerate(scen):
    i = S1_ + k
    put(ws, i, [sid, desc, 1000, cid, None, None, None, None, reg_, q], wrap=True)
    ws.cell(i, 1).font = BB
    if mixed: ws.cell(i, 2).fill = YEL
    for c, fmt in ((3, '"$"#,##0'), (4, None), (6, '"$"0.00'), (9, None), (10, None)):
        blue_in(ws.cell(i, c), fmt)
    ws.cell(i, 5).value = f'=IFERROR(INDEX($E${A1_}:$E${AL_},MATCH($D{i},$A${A1_}:$A${AL_},0)),"")'
    ws.cell(i, 7).value = f'=IF($F{i}<>"",$F{i},$E{i})'
    ws.cell(i, 8).value = f'=IFERROR($C{i}/$G{i},"")'
    ws.cell(i, 11).value = ret(i, "D1 리텐션")
    ws.cell(i, 12).value = ret(i, "D7 리텐션")
    ws.cell(i, 13).value = ret(i, "D30 리텐션")
    for c, src in ((14, "K"), (15, "L"), (16, "M")):
        ws.cell(i, c).value = f'=IFERROR($H{i}*{src}{i},"")'
    ws.cell(i, 17).value = f'=IFERROR(100/$L{i}*$G{i},"")'
    for c, fmt in ((5, '"$"0.00'), (7, '"$"0.00'), (8, "#,##0"), (11, "0.00%"), (12, "0.00%"), (13, "0.00%"),
                   (14, "#,##0"), (15, "#,##0"), (16, "#,##0"), (17, '"$"#,##0')):
        ws.cell(i, c).number_format = fmt
        ws.cell(i, c).alignment = Alignment(horizontal="center", vertical="top")
dv_id = DataValidation(type="list", formula1=f"$A${A1_}:$A${AL_-1}", allow_blank=False)
dv_rg = DataValidation(type="list", formula1='"북미,중미,남미,아시아"', allow_blank=False)
dv_q = DataValidation(type="list", formula1='"P50,P90,P99"', allow_blank=False)
for dv, col in ((dv_id, "D"), (dv_rg, "I"), (dv_q, "J")):
    ws.add_data_validation(dv); dv.add(f"{col}{S1_}:{col}{SL_}")

# 메모
notes31 = [
 "리텐션: '4-1 지역 벤치마크(참조)' = GameAnalytics 2026 보고서 지역 표(2025년 주간 데이터의 연평균, 전 장르 합산). P50 = 중앙값(보통 게임), P90 = 상위 10%, P99 = 상위 1%. 동남아는 이 표에 없음(시장조사 2-7 참고).",
 "D30: 4-1(참조)의 GameAnalytics 2026 값은 D30이라 그대로 씀. 2025 보고서(첨부 PDF)는 D28을 써서 직접 비교에 주의.",
 "장르별 리텐션은 이 표에 섞지 않음 — '4-3 장르 판독 요약(참조)'의 판독값(차트를 눈으로 읽은 값) 참고.",
 "광고 없이 들어오는 사람(오가닉)은 계산에 넣지 않음. 참고: Adjust 2025 게임 유료:오가닉 설치 비율 전 세계 중앙값 3.33 (2024년 2.07, p.16) — 이 비율을 우리 게임에 그대로 쓸 근거는 없음.",
 "참고 (투자 회수): Liftoff 2025 — 캐주얼 게임의 30일 광고비 회수율(D30 ROAS) 평균 iOS 47%, Android 15%. 즉 보통은 30일 안에 광고비를 다 되찾지 못함.",
 "광고비 $1,000과 시나리오 조합은 '같은 돈으로 비교'하려는 분석자 가정. 실제 CPI는 우리 게임 광고를 소액으로 돌려 보면 가장 정확하게 알 수 있음.",
 "실제 게임별 광고비·CPI·설치 수: 비공개. 대신 쓸 수 있는 공개 지표: 위 CPI 출처표, 4-1(참조) 리텐션 구간, 3-3 AppBrain 추정 설치 수(추정).",
]
for k, t in enumerate(notes31):
    ws.cell(SL_ + 2 + k, 1, t).font = NOTE
for col, w in zip("ABCDEFGHIJKLMNOPQ", [8, 44, 18, 24, 11, 16, 12, 30, 16, 11, 11, 11, 11, 11, 11, 11, 16]):
    ws.column_dimensions[col].width = w
ws.row_dimensions[BT2].height = 48
ws.freeze_panes = "A4"

# ---------- (5) '매출 Top100' H18 메모 고치기 (사용자 승인, 2026-10-05) ----------
ws = wb["매출 Top100"]
_old = "(1-4 세부 장르 시트에는 없음 — 1-4는 1-1·1-2만 집계)"
_new = "(6단계-3에서 '1-4 세부 장르' 맨 아래 별도 표에 추가함 — 1-4 위쪽 표는 1-1·1-2만 집계)"
assert isinstance(ws["H18"].value, str) and ws["H18"].value.count(_old) == 1
ws["H18"].value = ws["H18"].value.replace(_old, _new)

# ---------- (6) 광고 채널 분석 ----------
ws = wb.create_sheet("광고 채널 분석")
ws["A1"] = "광고 채널 분석 — 유튜브·인스타그램·페이스북·틱톡·인플루언서 (+ 분석자가 추가한 2개 채널)"; ws["A1"].font = TITLE
ws["A2"] = ("읽는 법: ① 공개 자료로 확인한 사실(도달·성과 순위)과 ② 2인 팀 기준 점수(분석자 판단)를 나눠 둠. "
            "파란 글씨 = 바꿀 수 있는 가중치, 노란 칸 = 2차 인용·분석자 판단·확인 불가")
ws["A2"].font = NOTE
DR_P = "DataReportal, Digital 2026 Mid-Year Global Update Report (18세 이상 광고 도달, 인도 제외) https://datareportal.com/reports/digital-2026-mid-year-global-update-report"
DR_S = "SQ Magazine, Social Media Advertising Statistics 2026 (DataReportal 2026-04 Global Statshot 인용) https://sqmagazine.co.uk/social-media-advertising-statistics/"
AF25 = "AppsFlyer, 2025 Performance Index 보도자료 (2025-12-03, 88개 매체·설치 162억 건) https://www.appsflyer.com/company/newsroom/pr/performance-index-2025/"
PG25 = "PocketGamer.biz, AppsFlyer Performance Index 2025 장르별 요약 (2025-12-04) https://www.pocketgamer.biz/apple-and-google-continue-to-lead-in-mobile-game-ads-but-applovin-wins-over-select-genres/"
AF24 = "AppsFlyer, 17th Performance Index 보도자료 (2024-10-14, 2024년 상반기 데이터) https://www.appsflyer.com/company/newsroom/pr/performance-index-17-released/"
IMH = "Influencer Marketing Hub, Influencer Marketing Benchmark Report 2026 (전 업종 브랜드 설문) https://influencermarketinghub.com/influencer-marketing-benchmark-report/"
IMH25 = "archive.com 블로그 (Influencer Marketing Hub Benchmark Report 2025 인용) https://archive.com/blog/social-media-platform-specific-influencer"
HUB = "Hubfluence, 2026 Influencer Marketing Benchmark Report (여러 출처 집계, 2026-06 기준) https://www.hubfluence.io/resources/influencer-marketing-report-2026"
LFT2 = "Liftoff·Singular, 2025 Casual Gaming Apps Report — 'Where Casual Games Find Their Players' (Liftoff 데이터 2024-02~2025-02) https://liftoff.ai/2025-casual-gaming-apps-report/"

T1 = 5
ws.cell(T1 - 1, 1, "① 채널별 공개 자료 (사실) — 게임 광고 비용(CPI)은 채널별 공개 원문이 없어 '확인 불가'. 지역·OS별 CPI는 3-1 시트 참고").font = TITLE
head_row(ws, T1, ["채널", "광고를 사는 곳", "광고 도달 (사람 수)", "도달 자료 등급", "게임 광고 성과 근거 (공개 순위)", "장르 근거 (공개 자료)",
                  "게임 광고 비용 지표", "출처"])
chan = [
 ("유튜브", "Google Ads 앱 캠페인 (유튜브·검색·Play 등에 함께 노출)", "약 26.5억 명", "2차 인용",
  "AppsFlyer 2025: Android 게임 매체 1위는 Google Ads (유튜브는 Google Ads 안에 묶여 따로 순위 없음)",
  "Google Ads는 Android에서 카지노·테이블탑·퍼즐·레이싱·액션 장르에서 특히 강함 (PocketGamer 요약)",
  "확인 불가 (채널별 게임 CPI 원문 없음)", DR_S + " / " + AF25 + " / " + PG25),
 ("인스타그램", "Meta Ads (페이스북과 같은 광고 도구로 함께 집행)", "14.4억 명 (인도 제외, 원문) / 약 19.9억 명 (전체, 2차)", "원문 + 2차",
  "AppsFlyer는 Meta Ads로 묶어 순위를 매김 → 인스타그램 단독 순위는 확인 불가",
  "Meta Ads: 2024년 상반기 Android 게임에서 카지노·미드코어·스포츠&레이싱 상위 (2024 데이터)",
  "확인 불가", DR_P + " / " + DR_S + " / " + AF24),
 ("페이스북", "Meta Ads (인스타그램과 같은 광고 도구)", "약 23.9억 명", "2차 인용",
  "인스타그램과 같음 (Meta Ads 한 묶음). 2025년에는 '격차를 좁힌 매체'로 언급",
  "인스타그램과 같음", "확인 불가", DR_S + " / " + AF25 + " / " + AF24),
 ("틱톡", "TikTok for Business (TikTok Ads Manager)", "22.1억 명 (18세 이상, 인도 제외)", "원문",
  "AppsFlyer 2025: Google·Apple과의 격차를 좁힌 매체로 언급. 게임 부문 구체 순위는 보도자료에 없음 → 확인 불가",
  "확인 불가 (게임 장르별 공개 원문을 찾지 못함)", "확인 불가", DR_P + " / " + AF25),
 ("인플루언서", "크리에이터에게 직접 의뢰 (유튜브·틱톡·인스타 채널)", "크리에이터마다 다름 (확인 불가)", "—",
  "게임 설치 성과에 대한 공개 순위 없음 (확인 불가). 참고: 2026년 브랜드 투자 계획 1순위 플랫폼은 틱톡(선택 31%, 전 업종)",
  "전 업종 기준: 인스타 57.1% · 틱톡 51.6% · 유튜브 36.7% · 페이스북 28.4%가 인플루언서 마케팅에 사용 (2025, 2차 인용)",
  "참고(2차 집계, 전 업종 크리에이터 콘텐츠): 1,000회 노출당 틱톡 $4.63 · 인스타 릴스 $8.50 · 유튜브 $15", IMH + " / " + IMH25 + " / " + HUB),
 ("[분석자 추가] Apple Ads", "App Store 검색 광고 (iOS 전용)", "확인 불가 (App Store 사용자 수 공개 기준 다름)", "—",
  "AppsFlyer 2025: iOS 게임 매체 1위 (단, 북미·서유럽에서는 AppLovin이 앞섬)",
  "iOS 게임 장르 중 하이퍼캐주얼·테이블탑을 뺀 모든 장르에서 1위 (PocketGamer 요약)", "확인 불가", AF25 + " / " + PG25),
 ("[분석자 추가] 광고 네트워크 (AppLovin·Mintegral·Unity·Liftoff 등, 다른 앱 안 광고)", "각 네트워크 광고 도구", "확인 불가", "—",
  "AppsFlyer 2025: AppLovin이 iOS 게임에서 북미·서유럽 선두, Android에서 AppLovin·Mintegral 상승",
  "캐주얼 게임 설치의 약 절반은 하이퍼캐주얼·퍼즐 게임 안 광고에서 옴 (Liftoff, 게임 퍼블리셔 기준)", "확인 불가", AF25 + " / " + LFT2),
]
C1, CL = T1 + 1, T1 + len(chan)
for k, row in enumerate(chan):
    i = C1 + k
    put(ws, i, list(row), wrap=True)
    ws.cell(i, 1).font = BB
    if row[3] in ("2차 인용", "원문 + 2차"): ws.cell(i, 3).fill = YEL; ws.cell(i, 4).fill = YEL
    for c in (3, 5, 6, 7):
        if "확인 불가" in str(ws.cell(i, c).value) or "2차" in str(ws.cell(i, c).value): ws.cell(i, c).fill = YEL
    ws.cell(i, 8).alignment = Alignment(vertical="top", wrap_text=False)

# 점수표
T2 = CL + 3
ws.cell(T2 - 1, 1, "② 2인 팀 채널 점수 — 점수(1~5, 5 = 좋음)는 위 ① 사실을 바탕으로 한 분석자 판단(노란 칸). 가중치는 파란 글씨로 바꿀 수 있음").font = TITLE
head_row(ws, T2, ["채널", "광고 도달 크기", "게임 광고 성과 근거", "우리 후보 장르(퍼즐·캐주얼) 적합", "2인 팀 운영 쉬움",
                  "작은 예산으로 시작", "점수 (100점, 수식)", "순위 (수식)", "점수 이유 (분석자)"])
W2 = T2 + 1
put(ws, W2, ["가중치 (합계 100%) →", 0.20, 0.25, 0.25, 0.20, 0.10, f"=SUM(B{W2}:F{W2})", "", "← 가중치는 분석자 제안"])
for c in range(2, 7): blue_in(ws.cell(W2, c), "0%")
ws.cell(W2, 7).number_format = "0%"; ws.cell(W2, 1).font = BB
sc = [
 ("유튜브 (Google Ads)", 5, 5, 5, 4, 4, "도달 최대(2차), Android 게임 1위 매체, 퍼즐 장르 강세 언급. 앱 캠페인은 소재만 넣으면 자동 최적화라 운영이 비교적 쉬움(분석자 판단)"),
 ("인스타그램 (Meta Ads)", 4, 3, 3, 3, 4, "도달 큼. 게임 성과 근거는 카지노·미드코어 쪽(2024). 소재를 자주 바꿔야 해서 운영 중간(분석자 판단)"),
 ("페이스북 (Meta Ads)", 5, 3, 3, 3, 4, "인스타그램과 같은 광고 도구. 도달은 더 큼(2차)"),
 ("틱톡", 5, 2, 3, 2, 3, "도달 큼(원문). 게임 장르별 성과 원문은 확인 불가. 세로 영상 소재를 계속 새로 만들어야 해 2인 팀 부담이 큼(분석자 판단)"),
 ("인플루언서", 3, 2, 3, 2, 3, "게임 설치 성과 공개 자료 없음. 섭외·계약·확인에 사람 손이 많이 듦(분석자 판단)"),
 ("[분석자 추가] Apple Ads", 3, 5, 4, 4, 4, "iOS 게임 1위 매체. 검색 키워드 광고라 영상 소재 부담이 적음. iOS에만 해당"),
 ("[분석자 추가] 광고 네트워크", 4, 4, 4, 3, 2, "퍼즐·캐주얼 설치가 많이 나오는 곳(Liftoff). 플레이어블 광고 소재가 필요하고, 최소 예산 조건은 확인 불가"),
]
S1, SL = T2 + 2, T2 + 1 + len(sc)
for k, row in enumerate(sc):
    i = S1 + k
    put(ws, i, [row[0]] + list(row[1:6]) + [None, None, row[6]], wrap=True)
    ws.cell(i, 1).font = BB
    for c in range(2, 7):
        ws.cell(i, c).fill = YEL; ws.cell(i, c).alignment = Alignment(horizontal="center", vertical="top")
    ws.cell(i, 9).fill = YEL
    w = W2
    ws.cell(i, 7).value = (f"=ROUND((B{i}*$B${w}+C{i}*$C${w}+D{i}*$D${w}+E{i}*$E${w}+F{i}*$F${w})/SUM($B${w}:$F${w})/5*100,0)")
    ws.cell(i, 8).value = f"=RANK(G{i},$G${S1}:$G${SL},0)"
    for c in (7, 8): ws.cell(i, c).alignment = Alignment(horizontal="center", vertical="top")

notes32 = [
 "인스타그램·페이스북은 같은 Meta Ads로 사고, AppsFlyer도 'Meta Ads' 하나로 순위를 매겨 두 채널을 따로 비교한 공개 성과 자료는 확인 불가.",
 "유튜브는 Google Ads 앱 캠페인 안에 함께 묶여 집행·순위가 나옴 → 유튜브만의 게임 성과 순위는 확인 불가.",
 "도달 숫자는 각 플랫폼 광고 도구가 알려 주는 '광고가 닿을 수 있는 사람 수'라 실제 사용자 수와 다를 수 있음(DataReportal도 같은 주의를 붙임). 출처 기준이 달라 직접 비교는 대략적으로만.",
 "인플루언서 수치는 게임이 아닌 전 업종 브랜드 설문·집계. 게임 설치 성과·비용은 비공개 → 대신 쓸 수 있는 공개 지표: 크리에이터 영상 조회수·참여율(vidIQ 등 도구로 확인 가능).",
 "채널별 실제 게임 광고비·CPI·설치 수: 비공개. 3-1 시트의 지역·OS·장르별 CPI를 대신 쓰고, 실제 값은 소액 테스트로 확인.",
 "Apple Ads·광고 네트워크 2줄은 요청 목록 밖이지만, 게임 광고 매체 순위 상위라 비교를 위해 분석자가 추가함.",
 "3-3 후킹 영상 분석과 3-4 마케팅 실행안은 이 점수표를 근거로 씀(다음 단계).",
]
for k, t in enumerate(notes32):
    ws.cell(SL + 2 + k, 1, t).font = NOTE
for col, w in zip("ABCDEFGHI", [30, 30, 26, 16, 42, 40, 30, 12, 60]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "B4"

# ---------- (7) 후킹 영상 분석 (AdWhispr 메타 광고 라이브러리 데이터, 2026-10-05 조회) ----------
ws = wb.create_sheet("후킹 영상 분석")
ws["A1"] = "후킹 영상 분석 — 광고 첫 마디(후킹) 유형별 빈도: Royal Match · MONOPOLY GO! · Magic Sort!"; ws["A1"].font = TITLE
ws["A2"] = ("읽는 법: 후킹 유형·형식은 AdWhispr가 AI로 분류한 값('AdWhispr 분류'). 숫자는 2026-10-05에 조회한 그대로이며, 비율·합계는 수식. "
            "처음 3초 화면은 메타 광고 라이브러리에서 영상을 직접 재생해 본 '분석자 판독'(⑤, 노란 칸)")
ws["A2"].font = NOTE
ADW = "AdWhispr (Meta 광고 라이브러리 수집·AI 분류), 브랜드 통계·광고 목록 조회 2026-10-05"
GAMES3 = ["Royal Match", "MONOPOLY GO!", "Magic Sort!"]
# ① 게임별 요약
R0 = 5
ws.cell(R0 - 1, 1, "① 게임별 광고 현황 (AdWhispr 브랜드 통계)").font = TITLE
head_row(ws, R0, ["게임", "3-3 유형 (수식)", "지금 집행 중인 광고 수", "AdWhispr가 추적하는 광고 수", "AI 분류가 끝난 표본 수",
                  "가장 오래 집행한 광고 (일)", "AdWhispr 추정 광고비 범위 ($)", "비고"])
stat = [("Royal Match", 1988, 116, 25, 49, "66,599 ~ 123,703"),
        ("MONOPOLY GO!", 1451, 120, 25, 203, "292,532 ~ 543,274"),
        ("Magic Sort!", 119, 119, 25, 68, "66,143 ~ 122,820")]
MDR = "'매출-다운로드 비교'!"
for k, (g, act, trk, smp, lng, spend) in enumerate(stat):
    i = R0 + 1 + k
    put(ws, i, [g, f'=IFERROR(INDEX({MDR}$H${H0+1}:$H${H0+len(rows33)},MATCH(A{i},{MDR}$A${H0+1}:$A${H0+len(rows33)},0)),"확인 불가")',
                act, trk, smp, lng, spend, "추정 광고비의 기간·계산 방법은 확인 불가 (AdWhispr 추정)"], wrap=True)
    ws.cell(i, 1).font = BB
    for c in (2, 3, 4, 5, 6): ws.cell(i, c).alignment = Alignment(horizontal="center", vertical="top")
    ws.cell(i, 3).number_format = "#,##0"
    ws.cell(i, 7).fill = YEL; ws.cell(i, 8).fill = YEL
# ② 후킹 유형 빈도
HD = R0 + 6
ws.cell(HD - 1, 1, "② 후킹 유형별 빈도 — AdWhispr 분류 표본(게임당 25개) 기준").font = TITLE
head_row(ws, HD, ["후킹 유형 (AdWhispr 라벨)", "쉬운 뜻 (분석자 번역)"] + GAMES3 + ["합계 (수식)", "비율 (수식)"])
hooks = [("direct_offer", "직접 제안형 — '무료·광고 없음·지금 해 봐'처럼 혜택을 바로 말함", 24, 21, 8),
         ("curiosity_gap", "호기심형 — '1%만 풀 수 있다'처럼 궁금하게 만듦", 0, 4, 14),
         ("problem_agitation", "문제 자극형 — 답답함·불편함을 먼저 보여 줌", 1, 0, 2),
         ("story", "이야기형 — 짧은 이야기로 시작", 0, 0, 1)]
H1, HL = HD + 1, HD + len(hooks)
for k, (lab, mean, a, b_, c_) in enumerate(hooks):
    i = H1 + k
    put(ws, i, [lab, mean, a, b_, c_, f"=SUM(C{i}:E{i})", f"=F{i}/$F${HL+1}"], wrap=True)
    ws.cell(i, 2).fill = YEL
    ws.cell(i, 7).number_format = "0%"
put(ws, HL + 1, ["합계", "", f"=SUM(C{H1}:C{HL})", f"=SUM(D{H1}:D{HL})", f"=SUM(E{H1}:E{HL})", f"=SUM(F{H1}:F{HL})", f"=SUM(G{H1}:G{HL})"])
ws.cell(HL + 1, 7).number_format = "0%"
for c in range(1, 8): ws.cell(HL + 1, c).font = BB
# ③ 형식 빈도
FD = HL + 4
ws.cell(FD - 1, 1, "③ 광고 형식별 빈도 — AdWhispr 분류 표본(게임당 25개) 기준").font = TITLE
head_row(ws, FD, ["형식 (AdWhispr 라벨)", "쉬운 뜻 (분석자 번역)"] + GAMES3 + ["합계 (수식)", "비율 (수식)"])
fmts = [("animation", "애니메이션 (게임 장면을 만든 영상)", 20, 19, 5),
        ("ugc_talking_head", "일반인이 화면에 나와 말함", 0, 0, 13),
        ("ugc_lifestyle", "일반인이 일상에서 게임하는 모습", 1, 0, 4),
        ("text_overlay", "큰 글자가 중심", 0, 6, 2),
        ("screenshot", "게임 화면 캡처", 3, 0, 1),
        ("studio_product", "제품 촬영형", 1, 0, 0)]
F1, FL = FD + 1, FD + len(fmts)
for k, (lab, mean, a, b_, c_) in enumerate(fmts):
    i = F1 + k
    put(ws, i, [lab, mean, a, b_, c_, f"=SUM(C{i}:E{i})", f"=F{i}/$F${FL+1}"], wrap=True)
    ws.cell(i, 2).fill = YEL
    ws.cell(i, 7).number_format = "0%"
put(ws, FL + 1, ["합계", "", f"=SUM(C{F1}:C{FL})", f"=SUM(D{F1}:D{FL})", f"=SUM(E{F1}:E{FL})", f"=SUM(F{F1}:F{FL})", f"=SUM(G{F1}:G{FL})"])
ws.cell(FL + 1, 7).number_format = "0%"
for c in range(1, 8): ws.cell(FL + 1, c).font = BB
# ④ 개별 광고 30개
AD = FL + 4
ws.cell(AD - 1, 1, "④ 게임별 영상 광고 10개씩 (AdWhispr 광고 목록, 오래 집행한 순, 집행 중) — 문구는 광고 원문 그대로").font = TITLE
head_row(ws, AD, ["게임", "Meta 광고 라이브러리 ID", "집행 일수", "후킹 유형 (AdWhispr)", "형식 (AdWhispr)", "톤 (AdWhispr)",
                  "오퍼 (AdWhispr)", "광고 문구 (첫 줄)"])
ads30 = [
 ("Royal Match", "1726316855083037", 50, "direct_offer", "animation", "excitement", "no_explicit_offer", "NO ADS. NO WIFI. PURE FUN."),
 ("Royal Match", "1089702293564432", 50, "direct_offer", "animation", "excitement", "no_explicit_offer", "NO ADS. NO WIFI. PURE FUN."),
 ("Royal Match", "1411165037532310", 50, "direct_offer", "studio_product", "curiosity", "no_explicit_offer", "No ADS"),
 ("Royal Match", "4630787260485910", 50, "problem_agitation", "screenshot", "excitement", "no_explicit_offer", "NO ADS. NO WIFI. PURE FUN."),
 ("Royal Match", "28734790306122605", 50, "direct_offer", "animation", "excitement", "no_explicit_offer", "No ADS and FREE"),
 ("Royal Match", "1360608269119654", 50, "direct_offer", "animation", "excitement", "no_explicit_offer", "NO ADS. NO WIFI. PURE FUN."),
 ("Royal Match", "1098005362662932", 50, "direct_offer", "animation", "excitement", "no_explicit_offer", "No ADS and FREE"),
 ("Royal Match", "1309645668892427", 50, "direct_offer", "ugc_lifestyle", "curiosity", "no_explicit_offer", "No ADS"),
 ("Royal Match", "2192760164630774", 50, "direct_offer", "animation", "excitement", "no_explicit_offer", "NO ADS. NO WIFI. PURE FUN."),
 ("Royal Match", "1392560629687289", 50, "direct_offer", "animation", "excitement", "no_explicit_offer", "NO ADS. NO WIFI. PURE FUN."),
 ("MONOPOLY GO!", "1674823853701845", 204, "curiosity_gap", "animation", "excitement", "no_explicit_offer", "Libera gli animaletti!"),
 ("MONOPOLY GO!", "796330853023057", 204, "direct_offer", "animation", "excitement", "free_trial", "Gioca gratis a Monopoly GO!"),
 ("MONOPOLY GO!", "1309942491068984", 177, "direct_offer", "animation", "excitement", "free_trial", "Gioca gratis a Monopoly GO!"),
 ("MONOPOLY GO!", "2424531614685734", 131, "direct_offer", "animation", "excitement", "no_explicit_offer", "Don't miss The Simpsons GO!"),
 ("MONOPOLY GO!", "1559166695787060", 131, "direct_offer", "animation", "excitement", "new_launch", "¡No te pierdas The Simpson GO!"),
 ("MONOPOLY GO!", "1316376480460608", 128, "direct_offer", "animation", "excitement", "new_launch", "¡No te pierdas The Simpson GO!"),
 ("MONOPOLY GO!", "1012642531160291", 128, "direct_offer", "animation", "excitement", "no_explicit_offer", "The Simpson GO! está en curso"),
 ("MONOPOLY GO!", "1357761026259464", 127, "direct_offer", "animation", "excitement", "new_launch", "The Simpsons GO! is live!"),
 ("MONOPOLY GO!", "1501392118334017", 127, "curiosity_gap", "animation", "excitement", "no_explicit_offer", "The Simpson GO! está en curso"),
 ("MONOPOLY GO!", "983924064252944", 126, "direct_offer", "animation", "excitement", "new_launch", "Non perderti The Simpsons GO!"),
 ("Magic Sort!", "3985548428413528", 69, "curiosity_gap", "animation", "curiosity", "no_explicit_offer", "Magical journey of sorting!"),
 ("Magic Sort!", "1573783297425746", 69, "curiosity_gap", "animation", "curiosity", "no_explicit_offer", "Magical journey of sorting!"),
 ("Magic Sort!", "1896160021343579", 69, "direct_offer", "ugc_lifestyle", "excitement", "no_explicit_offer", "(문구 없음)"),
 ("Magic Sort!", "1065385889771960", 69, "curiosity_gap", "animation", "curiosity", "no_explicit_offer", "Magic Sort"),
 ("Magic Sort!", "900245332631697", 69, "problem_agitation", "ugc_talking_head", "excitement", "no_explicit_offer", "(문구 없음)"),
 ("Magic Sort!", "1060580849761773", 69, "direct_offer", "ugc_lifestyle", "curiosity", "no_explicit_offer", "(문구 없음)"),
 ("Magic Sort!", "2477486662759105", 69, "curiosity_gap", "ugc_talking_head", "curiosity", "no_explicit_offer", "(문구 없음)"),
 ("Magic Sort!", "1032627419485915", 69, "curiosity_gap", "ugc_lifestyle", "curiosity", "no_explicit_offer", "(문구 없음)"),
 ("Magic Sort!", "1556811082584487", 55, "curiosity_gap", "text_overlay", "curiosity", "no_explicit_offer", "Only 1% can solve this puzzle."),
 ("Magic Sort!", "1834079867955631", 55, "curiosity_gap", "text_overlay", "curiosity", "no_explicit_offer", "Only 1% can solve this puzzle."),
]
A1x, ALx = AD + 1, AD + len(ads30)
for k, row in enumerate(ads30):
    i = A1x + k
    put(ws, i, list(row))
    ws.cell(i, 2).number_format = "@"
    ws.cell(i, 3).alignment = Alignment(horizontal="center")
# ④-2 30개 기준 후킹 빈도 (수식)
head_row(ws, AD, ["후킹 유형", "30개 중 개수 (수식)", "비율 (수식)"], c0=10)
for k, (lab, *_r) in enumerate(hooks):
    i = AD + 1 + k
    put(ws, i, [lab, f'=COUNTIF($D${A1x}:$D${ALx},J{i})', f"=K{i}/COUNTA($D${A1x}:$D${ALx})"], c0=10)
    ws.cell(i, 12).number_format = "0%"
# ⑤ 처음 3초 화면 유형 — 분석자 판독 (메타 광고 라이브러리, 2026-10-05)
P5 = ALx + 3
ws.cell(P5 - 1, 1, "⑤ 처음 3초 화면 유형 — 분석자 판독 (메타 광고 라이브러리에서 영상을 재생해 0초·약 2초 화면을 캡처해 봄, 2026-10-05). 노란 칸 = 분석자 판독").font = TITLE
head_row(ws, P5, ["게임", "Meta 광고 라이브러리 ID", "처음 3초 화면 유형 (분석자 판독)", "요청 유형으로 묶기", "장면 설명 (분석자 판독)", "같은 영상 묶음"])
T_CRISIS, T_CHOICE, T_HAND, T_PLAY, T_REAL, T_3D, T_IP, T_ACTOR = (
    "위기 장면 (캐릭터가 위험)", "선택지 (어느 핀을 뽑을까)", "실사 손 연출 (녹이기·씻기)", "게임 플레이 화면",
    "실사 도전 장면 (사람·팻말)", "3D 연출 장면", "IP·캐릭터 등장", "실사 배우 연기")
GROUP = {T_CRISIS: "실패·위기 장면", T_CHOICE: "선택지"}
look = [
 ("Royal Match", "1726316855083037", T_CRISIS, "왕이 둥근 방에 갇혀 있고 핀 옆으로 물이 흘러듦", ""),
 ("Royal Match", "1089702293564432", T_CRISIS, "왕이 물살에 휩쓸려 미끄러진 뒤 물이 찬 칸에 갇힘, 아래는 매치3 판", ""),
 ("Royal Match", "1411165037532310", T_HAND, "실제 사람 손이 토치로 얼음을 녹임", ""),
 ("Royal Match", "4630787260485910", T_CRISIS, "왕이 자갈 아래 칸에 갇힘, 실제 사람 손이 도구를 듦, 아래는 매치3 판", ""),
 ("Royal Match", "28734790306122605", T_CRISIS, "지그재그 통로 위에서 자갈이 쏟아질 듯한 칸에 왕이 있음, 매치3 블록", ""),
 ("Royal Match", "1360608269119654", T_CHOICE, "핀 여러 개 중 하나를 실제 손가락이 당김, 왕 위에 블록", ""),
 ("Royal Match", "1098005362662932", T_CRISIS, "물이 찬 구불구불한 수로에 상어 같은 물체와 왕, 아래는 매치3 판", ""),
 ("Royal Match", "1309645668892427", T_HAND, "진흙 속 왕을 실제 손이 호스로 씻어 냄", ""),
 ("Royal Match", "2192760164630774", T_CRISIS, "왕 아래 자갈 더미, 옆에 매치3 판, 실제 손 등장", ""),
 ("Royal Match", "1392560629687289", T_3D, "블록이 움직이다 무너지며 자갈이 쏟아짐 (판독 애매)", ""),
 ("MONOPOLY GO!", "1674823853701845", T_ACTOR, "동물 조각상 클로즈업 → 철창 앞 실제 배우가 놀라는 장면", ""),
 ("MONOPOLY GO!", "796330853023057", T_IP, "모노폴리 판 위 3D 도시 → 미스터 모노폴리 캐릭터 클로즈업", "MG-A"),
 ("MONOPOLY GO!", "1309942491068984", T_IP, "796330853023057과 같은 장면", "MG-A"),
 ("MONOPOLY GO!", "2424531614685734", T_IP, "심슨 집 거실을 위에서 비추고 바트·개가 뛰어다님", "MG-B"),
 ("MONOPOLY GO!", "1559166695787060", T_IP, "2424531614685734와 같은 장면", "MG-B"),
 ("MONOPOLY GO!", "1316376480460608", T_IP, "2424531614685734와 같은 장면", "MG-B"),
 ("MONOPOLY GO!", "1012642531160291", T_IP, "미스터 모노폴리와 호머 심슨이 함께 등장", ""),
 ("MONOPOLY GO!", "1357761026259464", T_IP, "만화식 싸움 구름 + 협업 로고", ""),
 ("MONOPOLY GO!", "1501392118334017", T_IP, "심슨 그림체의 배 두 척 장면", ""),
 ("MONOPOLY GO!", "983924064252944", T_IP, "스프링필드 마을 전경 + 협업 로고", ""),
 ("Magic Sort!", "3985548428413528", T_PLAY, "병이 가득한 정렬 화면에서 손가락 아이콘이 병을 옮김", ""),
 ("Magic Sort!", "1573783297425746", T_PLAY, "병 정렬 플레이 (색 액체를 옮김)", ""),
 ("Magic Sort!", "1896160021343579", T_PLAY, "병 정렬 플레이 (어두운 배경)", ""),
 ("Magic Sort!", "1065385889771960", T_3D, "실사풍 병 3개 → 뒤쪽 병들이 깨지며 흩어짐", ""),
 ("Magic Sort!", "900245332631697", T_PLAY, "거의 정리된 병 화면에서 하나씩 옮김", ""),
 ("Magic Sort!", "1060580849761773", T_PLAY, "병 정렬 플레이 (나무 선반)", ""),
 ("Magic Sort!", "2477486662759105", T_PLAY, "빈 병을 손가락 아이콘이 옮기는 안내형 화면", ""),
 ("Magic Sort!", "1032627419485915", T_PLAY, "색 블록이 꽉 찬 정렬 화면", ""),
 ("Magic Sort!", "1556811082584487", T_REAL, "길거리에서 'Only 1% can solve this puzzle' 팻말 앞 여성이 실물 퍼즐에 도전, 사람들이 구경", "MS-A"),
 ("Magic Sort!", "1834079867955631", T_REAL, "1556811082584487과 같은 장면", "MS-A"),
]
assert len(look) == 30 and set(x[1] for x in look) == set(a[1] for a in ads30)
L1, LL = P5 + 1, P5 + len(look)
for k, (g, aid, typ, desc, grp) in enumerate(look):
    i = L1 + k
    put(ws, i, [g, aid, typ, GROUP.get(typ, "기타"), desc, grp], wrap=True)
    ws.cell(i, 2).number_format = "@"
    for c in (3, 4, 5): ws.cell(i, c).fill = YEL
# 집계 (수식)
types_ = [T_CRISIS, T_CHOICE, T_HAND, T_PLAY, T_REAL, T_3D, T_IP, T_ACTOR]
head_row(ws, P5, ["처음 3초 화면 유형", "Royal Match", "MONOPOLY GO!", "Magic Sort!", "합계"], c0=8)
for k, t in enumerate(types_):
    i = P5 + 1 + k
    put(ws, i, [t] + [f'=COUNTIFS($A${L1}:$A${LL},{col}${P5},$C${L1}:$C${LL},$H{i})' for col in ("I", "J", "K")]
            + [f"=SUM(I{i}:K{i})"], c0=8)
tr = P5 + len(types_) + 1
put(ws, tr, ["합계", f"=SUM(I{P5+1}:I{tr-1})", f"=SUM(J{P5+1}:J{tr-1})", f"=SUM(K{P5+1}:K{tr-1})", f"=SUM(L{P5+1}:L{tr-1})"], c0=8)
for c in range(8, 13): ws.cell(tr, c).font = BB
g0 = tr + 2
head_row(ws, g0, ["요청 유형", "Royal Match", "MONOPOLY GO!", "Magic Sort!", "합계"], c0=8)
for k, t in enumerate(["실패·위기 장면", "선택지", "전후 비교", "기타"]):
    i = g0 + 1 + k
    put(ws, i, [t] + [f'=COUNTIFS($A${L1}:$A${LL},{col}${g0},$D${L1}:$D${LL},$H{i})' for col in ("I", "J", "K")]
            + [f"=SUM(I{i}:K{i})"], c0=8)
ws.cell(g0 + 6, 8, "같은 영상 묶음(F열)을 하나로 치면 서로 다른 영상은 26개 (분석자 판독)").font = NOTE
# ⑥ 발견
P6 = LL + 3
ws.cell(P6 - 1, 1, "⑥ 발견 (광고 문구·라벨 기준 — 영상 화면을 본 판독 아님)").font = TITLE
finds = [
 "Royal Match: 영상 광고 10개 중 10개 문구가 '광고 없음(No ADS)'·'와이파이 없어도 됨'·'무료'를 앞세움 (위 ④ H열). 분류 표본 25개 중 24개가 직접 제안형.",
 "MONOPOLY GO!: 영상 광고 10개 중 7개가 'The Simpsons GO!' 협업 이벤트 알림 (④ H열). 가장 오래 집행한 광고는 204일 — 같은 소재를 오래 씀.",
 "Magic Sort!: 분류 표본 25개 중 호기심형 14개, 형식은 '일반인이 말하기'(UGC) 13개로 가장 많음. 대표 문구 'Only 1% can solve this puzzle.'",
 "처음 3초 (⑤, 분석자 판독): Royal Match는 10개 중 6개가 '왕이 위험에 빠진 장면'(실패·위기), 2개는 실제 사람 손이 녹이기·씻기를 하는 장면. MONOPOLY GO!는 10개 중 9개가 IP·캐릭터 등장(심슨 협업 7개). Magic Sort!는 10개 중 7개가 게임 플레이 화면 그대로.",
 "흥미로운 점(분석자 해석): AdWhispr 형식 라벨로는 Magic Sort!가 '일반인 말하기'가 많았지만, 실제 처음 3초는 대부분 게임 화면이었음 — 사람이 나오는 부분은 3초 뒤에 있거나 라벨이 영상 전체 기준일 수 있음.",
 "'전후 비교'형 첫 장면은 30개 중 0개. Royal Match의 위기 장면은 실제 매치3 판과 함께 나오는 경우가 많음 (보여 준 장면이 실제 게임에 있는지는 확인 불가).",
 "주의: 수집된 광고 일부는 유럽 대상(이탈리아어·스페인어 문구)이라 미국 광고만 따로 거르지 못함. AdWhispr 분류는 게임당 25개 표본이며, 나머지 광고 분류는 조회 시점에 진행 중이었음.",
 "실제 광고 성과(조회수·설치 수·광고비): 비공개. AdWhispr '추정 광고비'는 계산 방법이 공개되지 않은 추정치.",
 "출처: " + ADW,
]
for k, t in enumerate(finds):
    c = ws.cell(P6 + k, 1, t); c.font = NOTE if k >= 4 else B
for col, w in zip("ABCDEFGHIJKL", [26, 24, 26, 18, 48, 14, 20, 30, 14, 16, 14, 10]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A4"

# ---------- (8) 마케팅 실행안 ----------
ws = wb.create_sheet("마케팅 실행안")
ws["A1"] = "마케팅 실행안 — 우리 MVP(2인 개발)용. 이 시트의 판단은 모두 '분석자 제안'이며, 숫자는 근거 시트에서 수식으로 가져옴"; ws["A1"].font = TITLE
ws["A2"] = "파란 글씨 = 바꿀 수 있는 가정값(예산·후보·지역). 노란 칸 = 분석자 제안. 근거 열에 시트·칸을 적어 둠"; ws["A2"].font = NOTE
MV = "'2인 개발 MVP 판단'!"
put(ws, 4, ["MVP 후보 (분석자 제안, 바꿀 수 있음)", "Magic Sort!",
            f'=IFERROR("세부 장르: "&INDEX({MV}$C${C1_}:$C${CL_},MATCH(B4,{MV}$A${C1_}:$A${CL_},0))&" · 2인 점수: "&INDEX({MV}$H${C1_}:$H${CL_},MATCH(B4,{MV}$A${C1_}:$A${CL_},0))&" · 판정: "&INDEX({MV}$I${C1_}:$I${CL_},MATCH(B4,{MV}$A${C1_}:$A${CL_},0)),"2-4 목록에 없는 게임")',
            "근거: 2-4 ②·③ (점수), 3-3 (이 게임만 광고 데이터가 있음)"])
ws.cell(4, 1).font = BB; blue_in(ws.cell(4, 2))
ws.cell(4, 3).fill = YEL
# ① 단계 계획
T4 = 7
ws.cell(T4 - 1, 1, "① 단계별 실행 계획 — 지역·OS·채널·예산·예상 결과 (예상치는 3-1과 같은 계산: 예산 ÷ CPI × 리텐션)").font = TITLE
head_row(ws, T4, ["단계", "언제 (분석자 제안)", "지역", "OS (수식: 4-7 OS 우선순위)", "채널 (분석자 제안)", "채널 순위 (수식: 3-2)",
                  "예산 ($)", "CPI ID (3-1)", "CPI ($, 수식)", "예상 설치 수", "리텐션 지역", "D7 남는 사람 (P50)",
                  "합격 기준 (2-2)", "근거"])
OSR = "'대륙별 OS 분포'!"
CH = "'광고 채널 분석'!"
CP = "'광고 대비 유입'!"
RB2 = "'지역 벤치마크'!"
stages = [
 ("0 준비", "출시 4주 전", "—", "", "스토어 페이지 + 광고 소재 준비 (② 중 D·A·B 먼저)", "", 0, "", "",
  "광고 없음 — 소재·스토어 페이지만 준비", "3-3 ⑥"),
 ("1 소프트런치", "출시 1~4주", "브라질", "브라질", "유튜브 (Google Ads)", "유튜브 (Google Ads)", 1000, "C04", "남미",
  "D1·D7이 남미 중앙값 이상이면 2단계로", "CPI가 가장 싼 남미에서 리텐션부터 확인 — 3-1 S3, 2-6/4-7 브라질 Android 먼저"),
 ("2 아시아 테스트", "출시 5~8주", "인도네시아", "인도네시아", "유튜브 (Google Ads)", "유튜브 (Google Ads)", 1000, "C05", "아시아",
  "D1·D7이 아시아 중앙값 이상", "아시아태평양 CPI $0.27 (3-1 C05). 인도네시아는 2-7 동남아 분석 지역"),
 ("3 미국 출시", "출시 9~16주", "미국", "미국", "유튜브 (Google Ads) + Apple Ads", "[분석자 추가] Apple Ads", 3000, "C03", "북미",
  "D1 '30% 약간 상회'(전체 상위 25%)에 가까우면 확대", "미국은 iOS 먼저(또는 동시) — 4-7. Apple Ads는 iOS 게임 1위 매체 (3-2)"),
 ("4 확대", "출시 17주~", "미국", "미국", "Meta(인스타·페이스북) + 광고 네트워크 추가", "[분석자 추가] 광고 네트워크", 5000, "C03", "북미",
  "D30 광고비 회수율이 오르면 계속 (참고: 캐주얼 D30 ROAS iOS 47%·Android 15%, 3-1 메모)", "3-2 점수 3~5위 채널. 틱톡·인플루언서는 소재 만들 여력이 생긴 뒤"),
]
T41, T4L = T4 + 1, T4 + len(stages)
for k, s in enumerate(stages):
    i = T41 + k
    stage, when, reg_, osk, chan_, chkey, bud, cid, rreg, passc, why = s
    put(ws, i, [stage, when, reg_, None, chan_, None, bud, cid, None, None, rreg, None, passc, why], wrap=True)
    ws.cell(i, 1).font = BB
    for c in (2, 5, 13, 14): ws.cell(i, c).fill = YEL
    if osk:
        ws.cell(i, 4).value = f'=IFERROR(INDEX({OSR}$C$1:$C$200,MATCH("{osk}",{OSR}$A$1:$A$200,0)),"확인 불가")'
        ws.cell(i, 6).value = f'=IFERROR(INDEX({CH}$H${S1}:$H${SL},MATCH("{chkey}",{CH}$A${S1}:$A${SL},0))&"위","확인 불가")'
        blue_in(ws.cell(i, 7), '"$"#,##0'); blue_in(ws.cell(i, 8)); blue_in(ws.cell(i, 11))
        ws.cell(i, 9).value = f'=IFERROR(INDEX({CP}$E${A1_}:$E${AL_},MATCH(H{i},{CP}$A${A1_}:$A${AL_},0)),"")'
        ws.cell(i, 10).value = f'=IFERROR(G{i}/I{i},"")'
        ws.cell(i, 12).value = (f'=IFERROR(J{i}*SUMIFS({RB2}$C$2:$C${REGLAST},{RB2}$A$2:$A${REGLAST},K{i},'
                                f'{RB2}$B$2:$B${REGLAST},"D7 리텐션"),"")')
        ws.cell(i, 9).number_format = '"$"0.00'
        for c in (10, 12): ws.cell(i, c).number_format = "#,##0"
    for c in (3, 4, 6, 8, 9, 10, 11, 12): ws.cell(i, c).alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
put(ws, T4L + 1, ["합계", "", "", "", "", "", f"=SUM(G{T41}:G{T4L})", "", "", f"=SUM(J{T41}:J{T4L})", "", f"=SUM(L{T41}:L{T4L})", "", ""])
ws.cell(T4L + 1, 7).number_format = '"$"#,##0'
for c in (10, 12): ws.cell(T4L + 1, c).number_format = "#,##0"
for c in range(1, 15): ws.cell(T4L + 1, c).font = BB
# ② 소재 계획
T5 = T4L + 4
ws.cell(T5 - 1, 1, "② 광고 소재 5종 (분석자 제안 — 3-3 발견을 2인 팀이 만들 수 있게 바꾼 것). 쉬운 D·A·B부터 시작").font = TITLE
head_row(ws, T5, ["소재", "처음 3초에 보여 줄 것 (제안)", "참고한 3-3 근거", "2인 팀 제작 난이도 (분석자 판단)"])
mats = [("A 호기심형", "'1%만 풀 수 있다' 같은 글자 + 거의 다 푼 퍼즐 화면", "Magic Sort! 호기심형 14/25 (3-3 ②)", "쉬움 — 게임 화면 녹화 + 글자"),
        ("B 직접 제안형", "'무료 · 와이파이 없이' 같은 혜택 문구 + 게임 화면", "Royal Match 직접 제안형 24/25 (3-3 ②)", "쉬움 — 우리 게임에 실제로 맞는 혜택만 써야 함"),
        ("C 일반인 반응형 (UGC)", "사람이 퍼즐을 보고 반응하는 장면", "Magic Sort! 형식 '일반인 말하기' 13/25 (3-3 ③), 실물 퍼즐 길거리 도전 2/10 (3-3 ⑤)", "보통 — 출연자 섭외·촬영 필요"),
        ("D 게임 화면 그대로", "처음부터 실제 정렬 플레이 화면 (손가락 아이콘으로 옮기는 모습)", "Magic Sort! 처음 3초 7/10이 게임 화면 (3-3 ⑤)", "가장 쉬움 — 화면 녹화만"),
        ("E 위기 장면형", "캐릭터가 위험에 빠진 장면 → 퍼즐로 구함", "Royal Match 처음 3초 6/10 (3-3 ⑤)", "어려움 — 캐릭터·연출 필요. 실제 게임에 있는 장면만 쓰는 것이 안전(분석자 판단)")]
for k, m in enumerate(mats):
    i = T5 + 1 + k
    put(ws, i, list(m), wrap=True)
    ws.cell(i, 1).font = BB
    for c in (2, 4): ws.cell(i, c).fill = YEL
# ③ 측정 지표
T6 = T5 + 9
ws.cell(T6 - 1, 1, "③ 측정 지표 — 매주 볼 숫자와 비교 기준 (기준값은 수식 또는 근거 시트 값)").font = TITLE
head_row(ws, T6, ["지표", "쉬운 뜻", "비교 기준", "근거"])
NA_ = lambda met, col: f'SUMIFS({RB2}${col}$2:${col}${REGLAST},{RB2}$A$2:$A${REGLAST},"북미",{RB2}$B$2:$B${REGLAST},"{met}")'
mets = [
 ("CPI", "설치 1번에 쓴 광고비", f'="미국 게임 전체 평균 $"&TEXT(INDEX({CP}$E${A1_}:$E${AL_},MATCH("C03",{CP}$A${A1_}:$A${AL_},0)),"0.00")', "3-1 C03 (Adjust 2026)"),
 ("D1 리텐션", "설치 다음 날 다시 온 비율", f'="북미 중앙값 "&TEXT({NA_("D1 리텐션","C")},"0.0%")&" / 상위 10% "&TEXT({NA_("D1 리텐션","D")},"0.0%")', "4-1 지역 벤치마크(참조), 2-2"),
 ("D7 리텐션", "7일 뒤 다시 온 비율", f'="북미 중앙값 "&TEXT({NA_("D7 리텐션","C")},"0.0%")&" / 상위 10% "&TEXT({NA_("D7 리텐션","D")},"0.0%")', "4-1 지역 벤치마크(참조), 2-2"),
 ("D30 리텐션", "30일 뒤 다시 온 비율", f'="북미 중앙값 "&TEXT({NA_("D30 리텐션","C")},"0.0%")&" / 상위 10% "&TEXT({NA_("D30 리텐션","D")},"0.0%")', "4-1 지역 벤치마크(참조), 2-2"),
 ("D30 광고비 회수율 (ROAS)", "30일 동안 번 돈 ÷ 쓴 광고비", "캐주얼 평균 iOS 47% · Android 15%", "Liftoff 2025 (3-1 메모)"),
 ("소재별 설치 효율", "같은 돈으로 어느 소재가 설치를 더 많이 냈나", "공개 기준값 확인 불가 — 우리 소재 A~E끼리 비교", "② 소재 계획"),
]
for k, m in enumerate(mets):
    i = T6 + 1 + k
    put(ws, i, list(m), wrap=True)
    ws.cell(i, 1).font = BB
    if "확인 불가" in str(m[2]): ws.cell(i, 3).fill = YEL
# 메모
T7 = T6 + len(mets) + 3
for k, t in enumerate([
 "이 시트의 단계·지역·채널·예산·소재·합격 기준은 모두 분석자 제안이며, 실제 성과를 보장하지 않음. 예상 설치·D7 숫자는 공개 평균값으로 계산한 '참고용 추정'.",
 "OS 열은 '4-7 대륙별 OS 분포(참조)'의 'OS 출시 우선순위'(분석자 제안, 기준값 50%)를 수식으로 가져옴. 채널 순위는 '3-2 광고 채널 분석' 점수표 순위.",
 "CPI는 3-1의 지역별 값(iOS+Android 합산, 게임 전체)을 씀 — 퍼즐·OS별로 나눈 지역 값은 확인 불가.",
 "1단계를 남미로 둔 이유(분석자 제안): CPI가 가장 싸서(3-1) 같은 돈으로 리텐션 표본을 많이 모을 수 있음. 단, 남미 리텐션 중앙값은 북미보다 낮음(4-1) — 비교는 같은 지역 기준으로.",
 "MVP 후보(B4)를 바꾸면 위 요약 칸이 2-4에서 새로 계산됨. 3-3 광고 데이터는 Magic Sort!·Royal Match·MONOPOLY GO! 3개만 있음.",
]):
    ws.cell(T7 + k, 1, t).font = NOTE
for col, w in zip("ABCDEFGHIJKLMN", [26, 34, 22, 16, 30, 16, 12, 10, 10, 12, 10, 12, 34, 44]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "B4"

# ---------- (9) 광고 소재 해부 (메타 광고 라이브러리 영상 12개, 2026-10-05) ----------
ws = wb.create_sheet("광고 소재 해부")
ws["A1"] = "광고 소재 해부 — 게임별 4개씩 12개 영상: 도파민(보상·긴장) 요소 · 영상 효과 · 광고 구성 · 사운드(사용자 기입)"; ws["A1"].font = TITLE
ws["A2"] = ("보는 법: 메타 광고 라이브러리에서 영상을 열고 0·1·2·3·4·6·8·10·12·15·20초와 마지막 장면(12장)을 캡처해 판독. "
            "길이·크기는 영상 파일에서 잰 측정값, 나머지(노란 칸)는 분석자 판독 — 캡처 사이 장면은 보지 못함. "
            "T~Y열(사운드)은 사용자가 직접 듣고 고르는 칸(파란 글씨)")
ws["A2"].font = NOTE
HR = 4
head_row(ws, HR, ["게임", "Meta 광고 라이브러리 ID", "길이 (초, 측정)", "가로 px (측정)", "세로 px (측정)", "화면 비율 (수식)",
                  "처음 3초 (판독)", "위기·실패 장면", "위기 시작 → 해소 (초, 판독)", "일부러 틀리기·실패 연출",
                  "첫 보상 연출 (판독)", "만족감 장면", "진행 표시", "실사 사람·손", "큰 장면 전환 수 (12장 기준)",
                  "눈에 띄는 영상 효과", "엔드카드", "엔드카드 내용", "메모",
                  "[사운드] 배경음악", "[사운드] 음악 분위기", "[사운드] 효과음 종류", "[사운드] 목소리·나레이션",
                  "[사운드] 장면 전환이 박자와 맞음", "[사운드] 소리 없이도 이해됨"])
NS = "샘플에서 안 보임"
deep = [
 ("Royal Match", "1360608269119654", 59.1, 360, 450, "실사 손가락이 핀을 당기는 퍼즐, 왕 아래는 용암", "있음",
  "0초 위기 → 10초 실패(블록이 왕을 맞힘) → 12초 매치3 판으로 전환", "있음", NS, "—", NS, "있음", 2, "불꽃(6초)", "예",
  "로고 + 게임 화면 모음 + Play Now", ""),
 ("Royal Match", "1309645668892427", 59.7, 360, 540, "진흙 속 왕 클로즈업 → 실사 손이 호스로 씻김", "있음",
  "0초 진흙에 묻힘 → 6초 깨끗해짐(해소) → 8초 물 찬 방(새 위기)", NS, NS, "씻기(청소)", "목표 아이콘 (10초~)", "있음", 2,
  "흐림·카메라 이동(1초)", "예", "로고 + 다음 레벨 고르기 화면 ('Choose the next level!')", ""),
 ("Royal Match", "1392560629687289", 59.7, 360, 640, "3D 상자가 갈리며 자갈이 쏟아짐", "있음",
  "6초 자갈이 왕 쪽으로 쏟아짐 → 20초까지 점점 차오름", NS, NS, "부수기·쏟아지기", "자갈 높이 (긴장 표시)", "있음", 2,
  "파편·입자가 쏟아짐", "예", "로고 + 게임 화면 모음 + Play Now", ""),
 ("Royal Match", "1089702293564432", 59.7, 360, 360, "왕이 물살에 미끄러져 물 찬 방에 갇힘", "있음",
  "0초 미끄러짐 → 1초 갇힘 → 20초까지 매치3로 구출 진행", NS, NS, "—", NS, "있음", 2, "물 흐름", "예",
  "로고 + 게임 화면 모음 + Play Now", "메타에 '여러 버전이 있는 광고'로 표시 — 본 영상은 그중 1개"),
 ("MONOPOLY GO!", "1674823853701845", 17.0, 640, 360, "동물 조각상 클로즈업 → 실사 손에 든 휴대폰 게임 화면", "있음",
  "2초 철창 속 배우(코믹하게 갇힘)", NS, NS, "—", "—", "있음", 7, "큰 자막(8초), 실사 → 3D 게임 화면 전환", "예",
  "로고 + 이벤트 캐릭터 + 이벤트 문구 + 버튼 (14초~, 약 3초)", "이탈리아어 문구 — 유럽 대상 광고로 보임"),
 ("MONOPOLY GO!", "796330853023057", 29.5, 640, 360, "3D 보드 도시 → 미스터 모노폴리 → 'GO TO JAIL' 칸", "있음",
  "3초 감옥 + '도와줘(AIUTO!)' 말풍선 → 4~20초 블록 밀기 퍼즐로 열쇠 꺼내기", NS, NS, "블록 밀기 퍼즐", "열쇠 위치",
  "없음", 6, "말풍선 글자 팝", "예", "보드 + 로고 + 'GIOCA ORA'(지금 플레이) 버튼", "광고 속 퍼즐이 실제 게임에 있는지는 확인 불가"),
 ("MONOPOLY GO!", "2424531614685734", 21.6, 360, 360, "심슨 집 거실을 위에서 비추고 바트·개가 뜀", "있음",
  "8초 바트가 감옥(철창)에 들어감", NS, "3초 '7,500' 돈 숫자", "—", "보드 칸 이동", "없음", 6, "숫자 팝업", "아니오",
  "게임 장면으로 끝남", ""),
 ("MONOPOLY GO!", "1012642531160291", 40.5, 360, 360, "미스터 모노폴리와 호머가 빨간 버튼 앞에서 놀람", "있음",
  "4초 노란 폭발 플래시, 15초 경고 아이콘, 20초 감옥", NS, "12초 '+150K' 숫자", "—", "보드 칸 이동", "없음", 7,
  "화면 플래시(4초), 초록 연기 입자(6~10초)", "예", "로고 + 캐릭터 2명 + 버튼", ""),
 ("Magic Sort!", "1556811082584487", 56.2, 360, 640, "길거리 실물 퍼즐 도전 + 'Only 1%' 팻말 + 구경꾼", "없음",
  "—", NS, NS, "액체 붓기·채우기", "완성된 병이 늘어남", "있음", 2, "실사 → 게임 화면 전환(약 7초)", "예", "로고 + Play Now", ""),
 ("Magic Sort!", "3985548428413528", 60.0, 360, 640, "병이 가득한 정렬 화면 + 손가락 아이콘", "없음",
  "—", NS, "단색 병 완성 (작은 보상으로 판독)", "정렬(정리)", "완성된 병이 늘어남", "없음", 0, "손가락 아이콘·붓기 애니메이션",
  "아니오", "플레이 화면으로 끝남", "60초 내내 한 장면"),
 ("Magic Sort!", "1065385889771960", 26.8, 360, 360, "실사풍 3D 병에 액체 붓기 → 2초 유리 깨짐", "있음",
  "2·4초 병이 깨짐 → 21초 휴대폰 속 실제 게임 화면", "있음 (깨짐을 실패 연출로 판독)", NS, "붓기·부수기", "—", "없음", 2,
  "유리 파편", "예", "휴대폰 여러 대에 게임 화면 + 손가락", "3D 연출은 실제 게임 화면(2D)과 다름"),
 ("Magic Sort!", "1032627419485915", 28.3, 360, 640, "가운데 긴 튜브 + 꽉 찬 색 블록 정렬 화면", "없음",
  "—", NS, "긴 튜브가 차오름 (진행 보상으로 판독)", "채우기", "긴 튜브 높이", "없음", 1, "붓기 애니메이션", "예", "로고 + Play Now", ""),
]
assert all(d[1] in {a[1] for a in ads30} for d in deep)
X1, XL = HR + 1, HR + len(deep)
for k, d in enumerate(deep):
    i = X1 + k
    g, aid, ln, w_, h_ = d[:5]
    rest = list(d[5:])
    put(ws, i, [g, aid, ln, w_, h_, f'=IF(D{i}>E{i},"가로형 ",IF(D{i}=E{i},"정사각형 ","세로형 "))&TEXT(D{i}/E{i},"0.00")'] + rest, wrap=True)
    ws.cell(i, 1).font = BB
    ws.cell(i, 2).number_format = "@"
    for c in range(7, 20): ws.cell(i, c).fill = YEL
    for c in range(20, 26):
        cc = ws.cell(i, c); cc.font = BLUE_IN; cc.border = BD
    for c in (3, 4, 5, 6, 8, 10, 14, 15, 17): ws.cell(i, c).alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
for col, opts in (("T", '"있음,없음,모름"'), ("W", '"있음,없음,모름"'), ("X", '"맞음,안 맞음,모름"'), ("Y", '"예,아니오,모름"')):
    dv = DataValidation(type="list", formula1=opts, allow_blank=True); ws.add_data_validation(dv); dv.add(f"{col}{X1}:{col}{XL}")
# 요약 (수식)
SM = XL + 3
ws.cell(SM - 1, 1, "요약 — 위 12개 영상 기준 (수식)").font = TITLE
head_row(ws, SM, ["게임", "영상 수", "평균 길이 (초)", "세로형 수", "정사각형 수", "가로형 수", "위기·실패 장면 있음",
                  "실사 사람·손 있음", "평균 큰 장면 전환 수", "엔드카드 있음"])
rg = lambda col: f"${col}${X1}:${col}${XL}"
for k, g in enumerate(GAMES3 + ["전체"]):
    i = SM + 1 + k
    crit = f'$A{i}' if g != "전체" else '"*"'
    put(ws, i, [g, f"=COUNTIF({rg('A')},{crit})", f"=ROUND(AVERAGEIF({rg('A')},{crit},{rg('C')}),1)",
                f'=COUNTIFS({rg("A")},{crit},{rg("F")},"세로형*")', f'=COUNTIFS({rg("A")},{crit},{rg("F")},"정사각형*")',
                f'=COUNTIFS({rg("A")},{crit},{rg("F")},"가로형*")', f'=COUNTIFS({rg("A")},{crit},{rg("H")},"있음")',
                f'=COUNTIFS({rg("A")},{crit},{rg("N")},"있음*")', f"=ROUND(AVERAGEIF({rg('A')},{crit},{rg('O')}),1)",
                f'=COUNTIFS({rg("A")},{crit},{rg("Q")},"예")'])
    ws.cell(i, 1).font = BB
    for c in range(2, 11): ws.cell(i, c).alignment = Alignment(horizontal="center")
notes35 = [
 "고른 기준: 3-3 ⑤의 처음 3초 유형이 겹치지 않게 게임별 4개 (Royal Match: 선택지·실사 손·3D 연출·위기 / MONOPOLY GO!: 실사 배우·캐릭터 2개·IP / Magic Sort!: 실사 도전·플레이 2개·3D 연출).",
 "'도파민 요소'는 뇌 반응을 잰 것이 아니라 화면에 보이는 보상·긴장·만족감 연출을 센 것(분석자 판독). '샘플에서 안 보임'은 없다는 뜻이 아니라 캡처한 12장에 없었다는 뜻.",
 "큰 장면 전환 수: 연속한 캡처 2장 사이에 장소·화면이 확 바뀐 횟수(엔드카드로 바뀜 포함). 캡처 간격이 1~5초라 실제 컷 수보다 적게 셀 수 있음.",
 "사운드: 이 판독 방식(화면 캡처)으로는 소리를 들을 수 없음 → T~Y열은 사용자가 메타 광고 라이브러리에서 B열 ID로 영상을 열어 듣고 고르는 칸.",
 "실제 광고 성과(조회·설치·광고비): 비공개. 대신 쓸 수 있는 공개 지표: 집행 일수(3-3 ④), 같은 영상 재사용(3-3 ⑤ F열).",
 "출처: Meta 광고 라이브러리 (https://www.facebook.com/ads/library/?id=광고 ID), 2026-10-05 조회. 일부 광고는 유럽 대상(이탈리아어·독일어 문구).",
]
for k, t in enumerate(notes35):
    ws.cell(SM + 6 + k, 1, t).font = NOTE
widths35 = [14, 18, 9, 8, 8, 12, 26, 9, 30, 14, 18, 14, 16, 10, 10, 22, 9, 26, 24, 11, 14, 16, 12, 12, 12]
for idx, w in enumerate(widths35, start=1):
    ws.column_dimensions[get_column_letter(idx)].width = w
ws.row_dimensions[HR].height = 48
ws.freeze_panes = "C5"

# ---------- (10) 광고 MVP 구성표 (분석자 제안) ----------
ws = wb.create_sheet("광고 MVP 구성표")
ws["A1"] = "광고 MVP 구성표 — 우리 Magic Sort!형 정렬 퍼즐용 15초·30초 광고 (모두 분석자 제안, 근거는 3-3·3-5)"; ws["A1"].font = TITLE
ws["A2"] = "광고 MVP = 광고 한 편에 꼭 있어야 할 최소 구성. 노란 칸 = 분석자 제안. 숫자 근거는 3-5에서 수식으로 가져옴"; ws["A2"].font = NOTE
SH = "'광고 소재 해부'!"
facts = [
 ("12개 중 엔드카드가 있는 영상", f'=COUNTIF({SH}$Q${X1}:$Q${XL},"예")&" / "&COUNTA({SH}$A${X1}:$A${XL})', "3-5 Q열"),
 ("12개 중 세로형 영상", f'=COUNTIF({SH}$F${X1}:$F${XL},"세로형*")&" / "&COUNTA({SH}$A${X1}:$A${XL})', "3-5 F열"),
 ("Magic Sort! 4개 평균 길이 (초)", f'=ROUND(AVERAGEIF({SH}$A${X1}:$A${XL},"Magic Sort!",{SH}$C${X1}:$C${XL}),1)', "3-5 C열"),
 ("Royal Match 4개 중 위기·실패 장면", f'=COUNTIFS({SH}$A${X1}:$A${XL},"Royal Match",{SH}$H${X1}:$H${XL},"있음")&" / 4"', "3-5 H열"),
 ("Magic Sort! 4개 평균 큰 장면 전환 수", f'=ROUND(AVERAGEIF({SH}$A${X1}:$A${XL},"Magic Sort!",{SH}$O${X1}:$O${XL}),1)', "3-5 O열"),
]
head_row(ws, 4, ["근거 숫자", "값 (수식)", "어디서"])
for k, (a, f_, src) in enumerate(facts):
    put(ws, 5 + k, [a, f_, src]); ws.cell(5 + k, 2).alignment = Alignment(horizontal="center")
TB = 5 + len(facts) + 3
ws.cell(TB - 1, 1, "① 구간별 구성 (분석자 제안)").font = TITLE
head_row(ws, TB, ["구간", "15초 버전", "30초 버전", "화면에 보여 줄 것", "도파민 장치 (보상·긴장)", "영상 효과",
                  "사운드 (제안 — 3-5 사운드 기입 뒤 확정)", "2인 팀 제작 난이도", "근거"])
segs = [
 ("① 후킹", "0~3초", "0~3초", "거의 다 푼 판 + 'Only 1%만 풀 수 있다' 같은 도전 문구 (또는 실물 퍼즐 도전 장면)",
  "도전 문구로 궁금증 + '나도 할 수 있나?'", "첫 장면부터 게임 화면, 글자는 크게 한 줄", "첫 0.5초 안에 소리 시작 (제안)",
  "쉬움 (도전 장면은 보통)", "3-3 ②·⑤ Magic Sort! 호기심형, 3-5 Magic Sort! 실사 도전"),
 ("② 핵심 놀이", "3~8초", "3~15초", "손가락 아이콘이 병을 옮기는 실제 플레이, 한 장면 유지",
  "단색 병이 하나씩 완성되는 진행감", "손가락 아이콘, 붓는 애니메이션", "붓는 소리 (제안)",
  "쉬움 (화면 녹화)", "3-5 Magic Sort! 플레이 영상 2개 (장면 전환 0~1)"),
 ("③ 긴장 → 해소", "8~12초", "15~25초", "마지막 한 칸 남기고 막힘 → 다시 풀어 성공 (일부러 틀리기 1번)",
  "아슬아슬함 + 해소 / '내가 하면 더 잘할 텐데'", "막혔을 때 화면 살짝 흔들기 (제안)", "실패 '삐' + 성공 '딩' (제안)",
  "보통 (연출 편집 필요)", "3-5 Royal Match 핀 퍼즐 실패 장면(10초), Magic Sort! 3D 깨짐 연출"),
 ("④ 보상", "12초 무렵", "25초 무렵", "판 완성 + 반짝임 + 별·숫자", "숫자 보상 팝업", "반짝이는 입자, 숫자 팝업",
  "완성 효과음 (제안)", "쉬움", "3-5 MONOPOLY GO! 숫자 보상('7,500', '+150K')"),
 ("⑤ 엔드카드", "12~15초", "25~30초", "게임 로고 + 실제 게임 화면 + 'Play Now' 버튼", "—", "버튼이 살짝 커졌다 작아짐 (제안)",
  "로고 소리 (제안)", "쉬움", "3-5 엔드카드 있음 (위 근거 숫자)"),
]
for k, s in enumerate(segs):
    i = TB + 1 + k
    put(ws, i, list(s), wrap=True)
    ws.cell(i, 1).font = BB
    for c in range(2, 9): ws.cell(i, c).fill = YEL
RB_ = TB + len(segs) + 3
ws.cell(RB_ - 1, 1, "② 기본 규칙 (분석자 제안)").font = TITLE
rules = [
 "화면 비율: 세로형(9:16)을 기본으로, 정사각형(1:1)을 하나 더 — 3-5에서 세로형·정사각형이 대부분.",
 "길이: 큰 회사는 60초 가까이 쓰지만(3-5 Royal Match), 2인 팀은 15초·30초 두 가지로 시작. 길이별 광고 성과 공개 자료는 확인 불가.",
 "실제 게임에 없는 장면은 쓰지 않기: 3-5에서 실제 게임과 다른 3D 연출이 있었음 — 우리 광고는 실제 화면 위주로(스토어·광고 정책 위험을 피하려는 분석자 판단).",
 "소재 테스트: 3-4 ② 소재 A~E 중 D(게임 화면)·A(호기심)·B(직접 제안)를 위 구성으로 각각 만들어 같은 예산으로 비교.",
 "사운드 칸은 3-5 T~Y열을 직접 듣고 채운 뒤 확정 — 지금은 제안만.",
]
for k, t in enumerate(rules):
    c = ws.cell(RB_ + k, 1, t); c.font = B; c.fill = YEL
for col, w in zip("ABCDEFGHI", [34, 14, 14, 40, 30, 28, 30, 18, 40]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A4"

# ---------- (4) 개요 맨 아래 2줄 추가 ----------
g = wb["개요"]
nr = g.max_row + 1
for k, (a, b_) in enumerate([
    ("6단계-3 출처 (2026-10-05 조회)", "2인 개발 MVP 판단: 3-2 매출 Top100·3-3 매출-다운로드 비교·1-2 Google Play Top100(모두 2026-10-04)을 게임전략 파일에 참조용 복사본으로 두고 수식으로 사용 + 세부 장르 점수는 분석자 판단. 광고 대비 유입: CPI — Adjust 'The gaming app insights report: 2026 edition' p.28~29(2024-01~2026-01), Liftoff·Singular '2025 Casual Gaming Apps Report'(2024-02~2025-02), 2차 인용 3건(Statista·Mistplay·GameDev Reports, 노란 칸). 리텐션 — GameAnalytics 2026 지역 표(4-2 / 게임전략 4-1 참조)."),
    ("6단계-3 한계", "세부 장르 점수·2인 역할·MVP 기간은 분석자 판단/추정(근거 수치 없음). '지역×OS×장르'를 모두 나눈 CPI는 확인 불가. 광고 채널 분석(3-2)의 채널별 게임 CPI는 확인 불가. 후킹 영상(3-3)은 AdWhispr 분류(게임당 25개 표본) + 영상 30개 처음 3초 분석자 판독(화면 캡처 2~3장 기준). 마케팅 실행안(3-4)·광고 MVP 구성표(3-6)는 전부 분석자 제안. 광고 소재 해부(3-5)는 영상 12개를 화면 캡처 12장씩으로 판독했고 사운드는 사용자 기입 칸.")]):
    g.cell(nr + k, 1, a).font = BB; g.cell(nr + k, 1).border = BD
    c = g.cell(nr + k, 2, b_); c.font = B; c.border = BD; c.alignment = Alignment(wrap_text=True, vertical="top")

# ===================== 6단계-4: 남은 빈칸 채우기 (2026-10-07 조회) =====================
# 사용자 승인: ① App Store 매출 Top100 새 시트 ② 3-3 ①·②·③ 유형 55개의 공식 설치 구간·출시일을 '미조사' 칸에 채움
# ③ '확인 필요' 칸은 재검토 표만 만들고 원래 칸은 그대로 ④ 2-6 표 B·C·동남아 4개국은 찾지 못해 그대로
SRC_AB_IOS_GROSS = "AppBrain App Store Top Grossing Games United States https://www.appbrain.com/stats/appstore-rankings/top_grossing/games/us"
GREEN = PatternFill("solid", fgColor="E2EFDA")

# ---------- (1) 3-3 매출-다운로드 비교: F·G열 채우기 ----------
inst64 = """MONOPOLY GO!|100,000,000+|2022-12|com.scopely.monopolygo
Roblox|1,000,000,000+|2014-06|com.roblox.client
Candy Crush Saga|1,000,000,000+|2012-11|com.king.candycrushsaga
Royal Match|100,000,000+|2020-07|com.dreamgames.royalmatch
Gossip Harbor: Merge & Story|50,000,000+|2022-09|com.mergegames.gossipharbor
Township|500,000,000+|2013-11|com.playrix.township
Coin Master|100,000,000+|2016-01|com.moonactive.coinmaster
Royal Kingdom|100,000,000+|2023-03|com.dreamgames.royalkingdom
Whiteout Survival|100,000,000+|2023-01|com.gof.global
Pokémon GO|500,000,000+|2016-07|com.nianticlabs.pokemongo
Last War:Survival Game|100,000,000+|2023-06|com.fun.lastwar.gp
DRAGON BALL Z DOKKAN BATTLE|50,000,000+|2015-07|com.bandainamcogames.dbzdokkanww
Toon Blast|100,000,000+|2017-03|net.peakgames.toonblast
Total Battle: War Strategy|50,000,000+|2017-10|com.totalbattle
Last Z: Survival Shooter|10,000,000+|2023-12|com.readygo.barrel.gp
Evony: The King's Return|100,000,000+|2016-04|com.topgamesinc.evony
RAID: Shadow Legends|50,000,000+|2018-07|com.plarium.raidlegends
Gardenscapes|500,000,000+|2016-07|com.playrix.gardenscapes
Dice Dreams|50,000,000+|2019-09|com.superplaystudios.dicedreams
Jackpot Party Casino Slots|10,000,000+|2013-06|com.williamsinteractive.jackpotparty
Last Asylum: Plague|10,000,000+|2026-02|com.phs.global
Fishdom|100,000,000+|2016-02|com.playrix.fishdomdd.gplay
Cashman Casino Slots Games|10,000,000+|2016-10|com.productmadness.cashmancasino
Match Factory!|10,000,000+|2023-12|net.peakgames.match
Bingo Blitz - Bingo Games|50,000,000+|2012-09|air.com.buffalo_studios.newflashbingo
Free Fire x NARUTO SHIPPUDEN|1,000,000,000+|2017-09|com.dts.freefireth
Pokémon TCG Pocket|50,000,000+|2024-09|jp.pokemon.pokemontcgp
Tasty Travels: Merge Game|10,000,000+|2023-12|com.fatmerge.global
Magic Sort!|10,000,000+|2024-06|com.grandgames.magicsort
Cube Land Puzzle Game|1,000,000+|2026-01|com.rotatelab.cubeblast
Colony Flow!|1,000,000+|2026-06|com.abi.colony.flow
Royal Smash! - Physics Puzzle|10,000,000+|2026-06|com.cyphergames.royalsmash
Loop Sort|1,000,000+|2025-09|com.Garawell.LoopSort
Angry Birds 2|100,000,000+|2015-07|com.rovio.baba
Rotate Rings|1,000,000+|2026-07|com.nebula.rotaterings
Meowdoku: Brain Puzzle Games|10,000,000+|2026-04|com.oakever.meowdoku
Vita Mahjong|100,000,000+|2024-01|com.vitastudio.mahjong
Smash Land!|1,000,000+|2026-08|com.flow.smashparty
Block Blast!|1,000,000,000+|2022-09|com.block.juggle
Bus Fever Party!|10,000,000+|2025-04|gridplus.busjam.carpuzzle
MemeMix Challenge: Funny Party|1,000,000+ (GP)|2026-09|com.mimik.challenge.fun
Brain Puzzle 3: Crazy Mind|10,000,000+|2025-12|com.brain.crazy.joygame
Woodoku Blast|10,000,000+|2024-10|com.tripledot.blockbash
Amaze GO!|100,000,000+|2026-02|com.oakever.arrows
Block Craft 3D: Building Game|500,000,000+|2016-08|com.fungames.blockcraft
Word Search Explorer|100,000,000+|2021-11|in.playsimple.wordsearch
Fish Sort Puzzle|1,000,000+|2026-03|triple.sorting.bubble.fish.match
Mahjong Blast|50,000,000+|2025-06|com.nebula.mahjongtile
Pocket Sort: Coin Merge Puzzle|10,000,000+|2025-09|com.pocket.sort.coin.puzzle.game
Arrow Puzzle: Tap Puzzle Games|50,000,000+|2025-11|com.easybrain.arrow.puzzle.game
Fortnite|10,000,000+|2020-04|com.epicgames.fortnite
Castle Busters|10,000,000+|2025-07|com.epicoro.castleclashers
Mob Control|100,000,000+|2021-05|com.vincentb.MobControl
Bulldozer Master: Crush & Mine|1,000,000+|2025-04|io.supercent.bulldozermasters
Clash of Clans|500,000,000+|2013-09|com.supercell.clashofclans
Candy Crush Soda Saga|500,000,000+|2014-06|com.king.candycrushsodasaga
Lightning Link Casino Slots|10,000,000+|2018-08|com.productmadness.lightninglink
Sword x Staff|1,000,000+|2026-05|com.zjcs.android.us
Pixel Flow!|10,000,000+|2025-10|com.loomgames.pixelflow
Lotsa Slots - Casino Games|10,000,000+|2017-10|com.diamondlife.slots.vegas.free
Homescapes|500,000,000+|2017-08|com.playrix.homescapes
Dark War Survival|10,000,000+|2023-12|com.readygo.dark.gp
Yarn Loop|1,000,000+|2025-12|com.combo.yarnflow
Honkai: Star Rail|10,000,000+|2023-04|com.HoYoverse.hkrpgoversea
Solitaire Grand Harvest|50,000,000+|2017-06|net.supertreat.solitaire
Wuthering Waves|10,000,000+|2024-05|com.kurogame.wutheringwaves.global
Disney Solitaire|10,000,000+|2025-02|com.superplaystudios.disneysolitairedreams
Coin Master - Board Adventure|5,000,000+|2025-08|com.moonactive.cmboard
Mystery Town: Merge Games|5,000,000+|2024-10|com.cedargames.mysterytown
Cash Frenzy - Casino Slots|10,000,000+|2019-01|slots.pcg.casino.games.free.android
Puzzles & Survival|10,000,000+|2020-09|com.global.ztmslg
All in Hole: Black Hole Games|10,000,000+|2023-06|com.homagames.studio.allinhole
Z Route: Redemption|5,000,000+|2026-02|com.zroute.global
Slots: Heart of Vegas Casino|10,000,000+|2015-07|com.productmadness.hovmobile
Flambé: Merge & Cook|5,000,000+|2024-12|com.cola.game
Merge Cooking|10,000,000+|2022-09|com.merge.cooking.theme.restaurant.food
Matching Story - Puzzle Games|10,000,000+|2021-05|com.joycastle.mergematch
Hay Day|100,000,000+|2013-11|com.supercell.hayday
June's Journey: Hidden Objects|50,000,000+|2017-03|net.wooga.junes_journey_hidden_object_mystery_game
Magic: The Gathering Arena|5,000,000+|2021-03|com.wizards.mtga
Quick Hit Casino Slots Games|10,000,000+|2015-03|com.ballytechnologies.quickhitslots
Travel Town - Merge Adventure|10,000,000+|2021-03|io.randomco.travel
PUBG MOBILE|1,000,000,000+|2018-03|com.tencent.ig
Puzzles & Chaos: Frozen Castle|10,000,000+|2022-12|com.global.pnck
GODDESS OF VICTORY: NIKKE|10,000,000+|2022-10|com.proximabeta.nikke
Bingo Voyage - Live Bingo Game|5,000,000+|2023-06|com.bingo.cruise.free.best.top.game
Brawl Stars|500,000,000+|2018-06|com.supercell.brawlstars
Lordrush|1,000,000+|2026-07|com.s01.global
eFootball|100,000,000+|2017-05|jp.konami.pesam
Minecraft: Dream it, Build it!|50,000,000+|2011-10|com.mojang.minecraftpe
Hollywood Merge|5,000,000+|2026-01|com.hollywood.merge
Piggy Kingdom|1,000,000+|2023-07|com.olleyo.piggy.king.free
MARVEL Strike Force: Squad RPG|50,000,000+|2018-01|com.foxnextgames.m3
Golf Clash - Golfing Simulator|10,000,000+|2016-10|com.playdemic.golf.android
Bingo Frenzy-Live Bingo Games|10,000,000+|2018-12|com.cooking.bingo.ldkwh
Clash of Critters|5,000,000+|2024-12|com.farlightgames.pgame.gp
WSOP Poker: Texas Holdem Game|50,000,000+|2013-11|com.playtika.wsop.gp
Toy Blast|100,000,000+|2014-10|net.peakgames.amy
Jackpot World - Slots Casino|10,000,000+|2018-05|com.grandegames.slots.dafu.casino
Hole Stars: Puzzle Game|5,000,000+|2025-09|com.moonactive.holestars
Hunting Sniper|5,000,000+|2023-07|com.online.huntingsniper
Color Block Jam|10,000,000+|2024-11|com.GybeGames.ColorBlockJam
Tiki Solitaire TriPeaks|10,000,000+|2013-08|com.gsn.android.tripeaks
Yahtzee With Buddies: Dice|10,000,000+|2017-06|com.scopely.yux
Hero Wars: Alliance RPG Legend|100,000,000+|2017-02|com.nexters.herowars
Police Chief|1,000,000+|2026-01|com.slg.policewar
Bus Traffic Fever!|10,000,000+|2026-02|jp.co.goodroid.hyper.busflow
Rise of Kingdoms: Lost Crusade|50,000,000+|2018-04|com.lilithgame.roc.gp
Solitaire Associations Journey|10,000,000+|2025-08|com.hitappsgames.wordsolitaire
Seaside Escape: Merge & Story|10,000,000+|2022-09|com.gamedots.seasideescape
Clash Royale|500,000,000+|2016-02|com.supercell.clashroyale
Match Masters|50,000,000+|2017-05|com.funtomic.matchmasters
CubeAway - 3DPuzzle|5,000,000+|2026-02|jp.co.goodroid.hyper.puzzle.arrows3dcubepuzzle
Paper.io 2|100,000,000+|2018-09|io.voodoo.paper2
Magic Tiles 3 - Piano Game|500,000,000+|2017-02|com.youmusic.magictiles
Arrows – Puzzle Escape|100,000,000+|2025-07|com.ecffri.arrows
8 Ball Pool|1,000,000,000+|2013-01|com.miniclip.eightballpool
Sled Surfers|10,000,000+|2025-05|com.sled.surfers.game
Jigsawcard Solitaire Puzzle|10,000,000+|2025-11|com.oakever.jigsawcard
Color by Number: Coloring Games|100,000,000+|2017-12|com.tfgco.apps.coloring.free.color.by.number
Warline: Sniper Strike|5,000,000+|2025-07|com.farlightgames.warline.gp.global
Color Block: Combo Blast|50,000,000+|2019-04|com.puzzlegames.puzzlebrickslegend
Brainy Escape Quest|10,000,000+|2025-05|com.metagame.annoying.trick.brainypuzzle
CarLoop|500,000+|2026-04|com.blast.puzzlegame.bear
Solitaire - Classic Card Games|50,000,000+|2021-04|solitaire.patience.card.games.klondike.free
Slow Mo Strike: Knives in Time|5,000,000+|2026-05|com.kayac.slow_mo_strike
Idle Cat Gunner: Shooter RPG|1,000,000+|2026-04|com.Chodun.CatGunner
Subway Surfers|1,000,000,000+|2012-09|com.kiloo.subwaysurf
Word Tiles - Relaxing Puzzle|100,000+|2026-07|com.agavegames.wordtiles
Geometry Dash Lite|500,000,000+|2013-09|com.robtopx.geometryjumplite
Mahjong Master: Daily Match|10,000,000+|2026-01|com.mahjong.master.solitaire.match.puzzle
Search It - Hidden Objects|5,000,000+|2026-01|com.leap.searchit
Happy Color: Color by Number|100,000,000+|2018-01|com.pixel.art.coloring.color.number
Apex Sniper: Animal Hunt|1,000,000+|2026-07|com.animal.hunter.shooting.apex
Toca Boca World|100,000,000+|2018-09|com.tocaboca.tocalifeworld
Among Us|500,000,000+|2018-06|com.innersloth.spacemafia
Ball Sort Puzzle - Color Game|50,000,000+|2022-05|ball.sort.puzzle.color.sorting.bubble.games
EA SPORTS FC Soccer Mobile 27|500,000,000+|2016-09|com.ea.gp.fifamobile
Hidden Object Games: Seek It|10,000,000+|2025-07|hidden.object.find.seek.search.mystery.puzzle.games.free
Aniimo|1,000,000+ (GP)|2026-09|com.x.aniimos
Tiki Smash!|100,000+ (GP)|?|com.thrivegames.bashbest
Word Connect Association|1,000,000+|2025-07|in.playsimple.wordtrip.association
Melon Sandbox|100,000,000+|2022-12|com.studio27.MelonPlayground
Double Tile!|10,000,000+|2025-10|com.hungrystudio.mahjong
Brain Puzzle 2: Logic Twist|10,000,000+|2025-05|com.brain.logic.joygame
Frost Valley: Merge & Story|10,000,000+|2025-11|com.liteorange.hometopia
Solitaire|10,000,000+|2017-01|com.smilerlee.klondike
Food Hunt: Pixel Puzzle|1,000,000+|2026-05|com.vigafun.funfinity.foodhunt
Angry Smash|500,000+|2026-07|com.tripledot.angrysmash
Car Evolve|1,000,000+|2026-03|io.heseri.car
Freddy Playground Sandbox|100,000+|2026-08|com.freddys.playground.sandbox
Solar Smash|100,000,000+|2020-01|com.paradyme.solarsmash
GOGO! Blast|1,000,000+|2026-04|com.nebula.arrows
Annoying Uncle Punch Game|100,000,000+|2024-08|com.gtsy.annoying.punch.game
Food Sort: Puzzle Game|1,000,000+|2025-09|com.triple.food.sort.puzzle.game
Fast Cash Farkle|50,000+|2026-08|com.greatplay.farklecash
Star Wars: Galaxy of Heroes|50,000,000+|2015-09|com.ea.game.starwarscapital_row
Fate/Grand Order (English)|5,000,000+|2017-06|com.aniplex.fategrandorder.en
Star Trek Fleet Command|10,000,000+|2018-09|com.scopely.startrek
Genshin Impact 6th Anniversary|100,000,000+|2020-08|com.miHoYo.GenshinImpact
Yalla Ludo - Ludo&Jackaroo|100,000,000+|2018-09|com.yalla.yallagames
Bloons Blitz|100,000+|2026-08|com.ninjakiwi.bloonsblitz
Screw Out 3D: Nut Sort Jam|1,000,000+|2025-08|net.happywit.screw.out.jam
Loop Master: Color Jam Sort|1,000,000+|2026-05|com.loop.clear.cube.free.card.puzzle.sort.match
Wordscapes Search: Word Games|10,000,000+|2019-08|com.peoplefun.wordsearch
Block Crush!|50,000,000+|2021-12|com.wood.block.sudoku.puzzle.bm
Crossword Go!|1,000,000+|2025-04|in.playsimple.crossword.go
Offline Games - No Wifi Games|100,000,000+|2023-08|com.JindoBlu.OfflineGames
Last Extract|1,000,000+|2026-05|com.xiaomo.lastextract
Vigor Mahjong|5,000,000+|2024-06|com.godlike.vigormahjong
Treasure Master|10,000,000+|2021-11|com.gimica.treasuremaster"""
ws = wb["매출-다운로드 비교"]
filled = {}
for line in inst64.split("\n"):
    name, rng, since, pkg = line.split("|")
    filled[name] = (rng, since, pkg)
CONFLICT = {  # 2026년 출시 게임 10개를 Google Play 상세 페이지에서 직접 대조 → 다른 3개 (AppBrain 값을 쓰고 노란 칸)
    "Amaze GO!": "Google Play 페이지에서 직접 읽으면 '50M+' → 출처끼리 다름 (노란 칸)",
    "Royal Smash! - Physics Puzzle": "Google Play 페이지에서 직접 읽으면 '5M+' → 출처끼리 다름 (노란 칸)",
    "Smash Land!": "Google Play 페이지에서 직접 읽으면 '100K+'~'500K+'(읽을 때마다 다름) → 출처끼리 다름 (노란 칸)",
}
done = 0
for r in range(6, ws.max_row + 1):
    name = ws.cell(r, 1).value
    if name in filled and ws.cell(r, 6).value == "미조사":
        rng, since, pkg = filled[name]
        gp_only = rng.endswith(" (GP)")
        rng = rng.replace(" (GP)", "")
        ws.cell(r, 6).value = rng
        ws.cell(r, 7).value = f"{since} (Google Play 등록 월)" if since != "?" else "확인 불가"
        for c in (6, 7): ws.cell(r, c).fill = PatternFill(fill_type=None)   # Kingshot 줄과 같은 모양
        if since == "?": ws.cell(r, 7).fill = YEL
        assert ws.cell(r, 9).value in (None, "")
        if gp_only:
            ws.cell(r, 9).value = (f"공식 설치 구간: Google Play 상세 https://play.google.com/store/apps/details?id={pkg} (2026-10-07 조회, AppBrain 페이지에 표시 없음/0+) · "
                                   f"등록 월: AppBrain 앱 상세 https://www.appbrain.com/app/{pkg}")
        else:
            ws.cell(r, 9).value = f"AppBrain 앱 상세 https://www.appbrain.com/app/{pkg} (2026-10-07 조회)"
        ws.cell(r, 9).font = B
        if name in CONFLICT:
            ws.cell(r, 6).fill = YEL
            ws.cell(r, 9).value += " — " + CONFLICT[name]
        done += 1
assert done == 171, done
n33b = ["6단계-4 (2026-10-07 조회): 172개 전부의 F·G열을 채움(Kingshot 포함). F열 = AppBrain 앱 상세 페이지에 옮겨 적힌 Google Play 다운로드 표시(예: '100,000,000+ Downloads'), "
        "G열 = 같은 페이지의 'available on Google Play since <월>' 문장. 링크는 I열. AppBrain에 표시가 없거나 '0+'인 3개(MemeMix·Aniimo·Tiki Smash!)는 Google Play 페이지에서 직접 확인.",
        "위 'Kingshot 1개만 확인' 줄은 6단계-2 당시 메모라 그대로 둠. 공식 구간은 '이 숫자 이상'이라는 뜻이라 실제 다운로드 수는 여전히 비공개. 2026년 출시 게임 10개를 Google Play에서 직접 대조해 3개가 달라 노란 칸으로 표시(AppBrain 옮겨 적기가 늦을 수 있음)."]
for k, t in enumerate(n33b):
    c = ws.cell(H0 + 7 + 6 + k, 11, t); c.font = NOTE

# ---------- (1-2) '확인 필요' 칸 반영 — 같은 부류끼리 묶어서 (사용자 승인, 2026-10-07) ----------
# 묶음 A 확인됨→노란색 해제 / B 대분류 확정 / C 퍼블리셔 / D 같은 방식·같은 회사 게임과 기준 맞춤 / E 스토어 분류만 확인(노란색 유지) / F 사실만 덧붙임
NOFILL = PatternFill(fill_type=None)
def fix(sn, r, vals, clear=None):
    ws_ = wb[sn]
    for col, v in vals.items():
        ws_[f"{col}{r}"].value = v
    if clear:
        for c in range(clear[0], clear[1] + 1): ws_.cell(r, c).fill = NOFILL
T = " (6단계-4)"
APPLY = [  # (묶음, 시트, 행, {열: 새 값}, 노란색 해제 범위 또는 None)
 ("A", "Google Play Top100", 11, {"F": "AI 음성 파티 게임 확인" + T}, (1, 6)),
 ("A", "Google Play Top100", 46, {"F": "색 블록 정렬 퍼즐 확인" + T}, (1, 6)),
 ("A", "매출 Top100", 40, {"F": "색 정렬 퍼즐 확인" + T}, (1, 6)),
 ("A", "매출 Top100", 45, {"F": "주사위 보드 게임 확인" + T}, (1, 6)),
 ("A", "매출 Top100", 73, {"F": "머지 + 스토리 확인" + T}, (1, 6)),
 ("A", "매출 Top100", 49, {"F": "블랙홀 io 게임 확인 (스토어 분류는 Puzzle)" + T}, (1, 6)),
 ("A", "App Store Top100", 69, {"F": "농구 라인업 짜기 게임 확인 (스토어 분류 Sports)" + T}, (1, 6)),
 ("A", "동남아 분석", 7, {"I": "리듬 게임 확인 (스토어 분류 음악)" + T}, (1, 9)),
 ("B", "국가별 Top20", 82, {"G": "전략"}, (1, 8)),
 ("B", "국가별 Top20", 94, {"G": "전략"}, (1, 8)),
 ("B", "동남아 분석", 13, {"G": "퍼즐", "I": "물 정렬 퍼즐 확인 (세부: 정렬·잼)" + T}, (1, 9)),
 ("C", "App Store Top100", 100, {"C": "Sven Vucak", "F": "퍼블리셔 확인 (App Store 판매자)" + T}, (1, 6)),
 ("D", "App Store Top100", 6, {"E": "픽셀·컬러 아트", "F": "픽셀 그림 + 색 슬롯 방식 확인 — Pixel Flow!와 같은 부류로 맞춤" + T}, (1, 6)),
 ("D", "Google Play Top100", 22, {"E": "픽셀·컬러 아트", "F": "픽셀 그림 + 색 슬롯 방식 확인 — Pixel Flow!와 같은 부류로 맞춤" + T}, (1, 6)),
 ("D", "매출 Top100", 67, {"E": "픽셀·컬러 아트", "F": "픽셀 그림 + 색 슬롯 방식 확인 — Pixel Flow!와 같은 부류로 맞춤" + T}, (1, 6)),
 ("D", "Google Play Top100", 88, {"F": "스토어 분류 Strategy 확인 — 같은 회사 Kingshot과 같은 기준으로 4X·SLG (세부는 분석자 판단)" + T}, None),
 ("D", "매출 Top100", 69, {"F": "스토어 분류 Strategy 확인 — 같은 회사 Kingshot과 같은 기준으로 4X·SLG (세부는 분석자 판단)" + T}, None),
 ("D", "App Store Top100", 72, {"F": "탄막 생존 게임 확인 — 아케이드 유지 (분석자 판단)" + T}, None),
 ("D", "Google Play Top100", 101, {"F": "탄막 생존 게임 확인 — 아케이드 유지 (분석자 판단)" + T}, None),
 ("E", "매출 Top100", 24, {"F": "스토어 분류 Strategy 확인, 세부(4X·SLG)는 분석자 판단" + T}, None),
 ("E", "매출 Top100", 50, {"F": "스토어 분류 Strategy 확인, 세부(4X·SLG)는 분석자 판단" + T}, None),
 ("E", "매출 Top100", 91, {"F": "스토어 분류 Strategy 확인, 세부(4X·SLG)는 분석자 판단 · 경영 성격도 있음" + T}, None),
 ("E", "매출 Top100", 35, {"F": "스토어 분류 RPG 확인, '방치형'은 확인 안 됨" + T}, None),
 ("E", "매출 Top100", 54, {"F": "매치3·타일·머지 혼합형 (스토어 분류 Casual)" + T}, None),
 ("E", "App Store Top100", 28, {"F": "실물 카드 뽑기·출금 가능(2차 출처) — 리얼머니 성격" + T}, None),
 ("F", "동남아 분석", 10, {"I": "확인 필요 — 게임 방식 확인 안 됨 · 리뷰에 '현금 보상' 주장 있음(검증 불가)" + T}, None),
 ("F", "동남아 분석", 62, {"C": "부분 확인 — 888XHSukari는 4x4 스도쿠, 888XHPearl은 타이밍 탭 게임(스토어 설명). 이름 규칙의 이유는 확인 안 됨" + T}, None),
]
for grp, sn, r, vals, clr in APPLY:
    fix(sn, r, vals, clr)

# ---------- (2) App Store 매출 Top100 (새 시트) ----------
ios_g = """1|Roblox|Roblox Corporation|기타(UGC·리워드)|UGC 플랫폼|1-1에 있는 분류 그대로|Roblox|
2|Royal Match|Dream Games|퍼즐|매치3|1-1에 있는 분류 그대로|Royal Match|
3|MONOPOLY GO!|Scopely, Inc.|카드·보드·카지노|보드·주사위 소셜캐주얼|1-1에 있는 분류 그대로|MONOPOLY GO!|
4|Gossip Harbor®: Merge & Story|Microfun Limited|퍼즐|머지|1-1에 있는 분류 그대로|Gossip Harbor: Merge & Story|
5|Candy Crush Saga|King|퍼즐|매치3|1-1에 있는 분류 그대로|Candy Crush Saga|
6|Kingshot|Century Games Pte. Ltd.|전략|4X·SLG|1-1에 있는 분류 그대로|Kingshot|
7|Whiteout Survival|Century Games Pte. Ltd.|전략|4X·SLG|1-1에 있는 분류 그대로|Whiteout Survival|
8|Pokémon GO|Scopely Explore, Inc.|RPG|위치기반 수집|1-1에 있는 분류 그대로|Pokémon GO|
9|Toon Blast|Peak Games|퍼즐|매치3|1-1에 있는 분류 그대로|Toon Blast|
10|Last War:Survival|FUNFLY PTE. LTD.|전략|4X·SLG|3-2에 있는 분류 그대로|Last War:Survival Game|
11|Clash Royale|Supercell|전략|실시간 카드 배틀|1-1에 있는 분류 그대로|Clash Royale|
12|Tasty Travels: Merge Game|Century Games Pte. Ltd.|퍼즐|머지|1-1에 있는 분류 그대로|Tasty Travels: Merge Game|
13|Township|Playrix|시뮬레이션|농장·타운 빌딩|1-1에 있는 분류 그대로|Township|
14|Match Factory!|Peak Games|퍼즐|타일·마작·트리플 매치|1-1에 있는 분류 그대로|Match Factory!|
15|Gardenscapes|Playrix|퍼즐|매치3|3-2에 있는 분류 그대로|Gardenscapes|
16|Free Fire x NARUTO SHIPPUDEN|GARENA INTERNATIONAL I PRIVATE LIMITED|액션·슈팅|배틀로얄|1-2에 있는 분류 그대로|Free Fire x NARUTO SHIPPUDEN|
17|Coin Master|Moon Active|카드·보드·카지노|보드·주사위 소셜캐주얼|1-1에 있는 분류 그대로|Coin Master|
18|Pixel Flow!|Loom Games Oyun Yazilim ve Pazarlama Anonim Sirketi|퍼즐|픽셀·컬러 아트|1-1에 있는 분류 그대로|Pixel Flow!|
19|Colony Flow!|ABI GLOBAL LTD.|퍼즐|픽셀·컬러 아트|1-1에 있는 분류 그대로 (6단계-4에서 확인 후 픽셀·컬러 아트로 맞춤)|Colony Flow!|
20|Fishdom|Playrix|퍼즐|매치3|3-2에 있는 분류 그대로|Fishdom|
21|Homescapes: Match 3 Games|Playrix|퍼즐|매치3|3-2에 있는 분류 그대로|Homescapes|
22|Evony|TOP GAMES INC.|전략|4X·SLG|3-2에 있는 분류 그대로|Evony: The King's Return|
23|Clash of Clans|Supercell|전략|4X·SLG|3-2에 있는 분류 그대로|Clash of Clans|
24|Merge Cooking®|Happibits|퍼즐|머지|3-2에 있는 분류 그대로|Merge Cooking|
25|Screwdom|Zego Global Pte Ltd|퍼즐|정렬·잼|새 게임 — 게임명 기반 분석자 판단 (나사 빼기 정렬 퍼즐로 봄)||Y
26|Total Battle: War Strategy|Scorewarrior|전략|4X·SLG|3-2에 있는 분류 그대로|Total Battle: War Strategy|
27|Lightning Link Casino Slots|Product Madness|카드·보드·카지노|소셜 카지노|3-2에 있는 분류 그대로|Lightning Link Casino Slots|
28|Pokémon TCG Pocket|The Pokemon Company|카드·보드·카지노|트레이딩 카드|1-2에 있는 분류 그대로|Pokémon TCG Pocket|
29|Jackpot Party - Casino Slots|Phantom EFX, Inc.|카드·보드·카지노|소셜 카지노|3-2에 있는 분류 그대로|Jackpot Party Casino Slots|
30|EA SPORTS FC Soccer Mobile 27|Electronic Arts|스포츠|축구|1-1에 있는 분류 그대로||
31|NYT Games: Wordle & Crossword|The New York Times Company|퍼즐|워드|1-1에 있는 분류 그대로||
32|Bus Traffic Fever!|GOODROID,Inc.|퍼즐|정렬·잼|1-1에 있는 분류 그대로|Bus Traffic Fever!|
33|Call of Duty®: Mobile|Activision Publishing, Inc.|액션·슈팅|FPS|1-1에 있는 분류 그대로||
34|Cashman Casino Slots Games|Product Madness|카드·보드·카지노|소셜 카지노|3-2에 있는 분류 그대로|Cashman Casino Slots Games|
35|Solitaire Associations Journey|Hitapps Games LTD|퍼즐|워드|1-1에 있는 분류 그대로|Solitaire Associations Journey|
36|DRAGON BALL Z DOKKAN BATTLE|Bandai Namco Entertainment Inc.|RPG|수집형(가챠) RPG|3-2에 있는 분류 그대로|DRAGON BALL Z DOKKAN BATTLE|
37|Dark War:Survival|Omnilojo Pte Ltd|전략|4X·SLG|3-2에 있는 분류 그대로|Dark War Survival|
38|Candy Crush Soda Saga|King|퍼즐|매치3|1-1에 있는 분류 그대로|Candy Crush Soda Saga|
39|Yarn Loop: Knit Puzzle|Combo Games|퍼즐|정렬·잼|1-1에 있는 분류 그대로|Yarn Loop|
40|All in Hole: Black Hole Games|HOMA GAMES|하이퍼·하이브리드 캐주얼|하이브리드 캐주얼|3-2에 있는 분류 그대로 (6단계-4에서 확인됨)|All in Hole: Black Hole Games|
41|Color Block Jam|Rollic Games|퍼즐|정렬·잼|1-2에 있는 분류 그대로|Color Block Jam|
42|Sword x Staff|Boltray Games|RPG|방치형 RPG|3-2에 있는 분류 그대로 (원래 칸이 '확인 필요' — 6단계-4 재검토 후에도 세부는 분석자 판단)|Sword x Staff|Y
43|Fortnite|Epic Games Inc.|액션·슈팅|배틀로얄|1-1에 있는 분류 그대로||
44|Travel Town - Merge Adventure|Moon Active|퍼즐|머지|3-2에 있는 분류 그대로|Travel Town - Merge Adventure|
45|Hay Day|Supercell|시뮬레이션|농장·타운 빌딩|3-2에 있는 분류 그대로|Hay Day|
46|eFootball™|KONAMI|스포츠|축구|3-2에 있는 분류 그대로|eFootball|
47|Quick Hit Casino - Vegas Slots|Appchi Media Ltd|카드·보드·카지노|소셜 카지노|3-2에 있는 분류 그대로|Quick Hit Casino Slots Games|
48|Magic: The Gathering Arena|Wizards of the Coast|카드·보드·카지노|트레이딩 카드|3-2에 있는 분류 그대로|Magic: The Gathering Arena|
49|Bingo Voyage: Live Bingo Games|VERTEX GAMES PTE. LTD.|카드·보드·카지노|빙고|3-2에 있는 분류 그대로|Bingo Voyage - Live Bingo Game|
50|Toy Blast|Peak Games|퍼즐|매치3|3-2에 있는 분류 그대로|Toy Blast|
51|Last Asylum: Plague|WISENESS GAME ONLINE INTERNATIONAL LTD|전략|4X·SLG|3-2에 있는 분류 그대로 (원래 칸이 '확인 필요' — 6단계-4 재검토 후에도 세부는 분석자 판단)|Last Asylum: Plague|Y
52|Puzzles & Survival|BUILDING-BLOCKS NETWORK TECHNOLOGY CO.,LIMITED|전략|4X·SLG|3-2에 있는 분류 그대로|Puzzles & Survival|
53|Bingo Blitz™ - BINGO Games|Playtika Santa Monica, LLC|카드·보드·카지노|빙고|3-2에 있는 분류 그대로|Bingo Blitz - Bingo Games|
54|Dice Dreams™|SuperPlay|카드·보드·카지노|보드·주사위 소셜캐주얼|1-1에 있는 분류 그대로|Dice Dreams|
55|Food Hunt: Pixel Puzzle|FUNFINITY PTE. LTD|퍼즐|픽셀·컬러 아트|1-1에 있는 분류 그대로||
56|Matching Story - Merge & Match|VERTEX GAMES PTE. LTD.|퍼즐|타일·마작·트리플 매치|3-2에 있는 분류 그대로 (원래 칸이 '확인 필요' — 6단계-4 재검토 후에도 세부는 분석자 판단)|Matching Story - Puzzle Games|Y
57|Lordrush|Century Games Innovation Pte. Ltd.|전략|4X·SLG|1-2에 있는 분류 그대로 (원래 칸이 '확인 필요' — 6단계-4 재검토 후에도 세부는 분석자 판단)|Lordrush|Y
58|Cube Land Puzzle Game|Rotatelab|퍼즐|블록|1-1에 있는 분류 그대로|Cube Land Puzzle Game|
59|Minecraft: Play with Friends!|Mojang|시뮬레이션|샌드박스|3-2의 Minecraft와 같은 분류 (스토어별 이름 다름)|Minecraft: Dream it, Build it!|
60|Chess.com - Play and Learn|Chess.com|카드·보드·카지노|클래식 보드|1-1에 있는 분류 그대로||
61|Free Fire MAX x NARUTO|GARENA INTERNATIONAL I PRIVATE LIMITED|액션·슈팅|배틀로얄|Free Fire와 같은 분류 (2-7에서 배틀로얄로 적음)||
62|Merge Mansion: Puzzles & Story|Metacore Games Oy|퍼즐|머지|새 게임 — 게임명 기반 분석자 판단||
63|Brawl Stars|Supercell|액션·슈팅|팀 대전 슈터|1-1에 있는 분류 그대로|Brawl Stars|
64|Hotel Legacy: Merge Game|Century Games Innovation Pte. Ltd.|퍼즐|머지|1-1에 있는 분류 그대로||
65|Big Fish Casino: Slots Games|Product Madness|카드·보드·카지노|소셜 카지노|새 게임 — 게임명 기반 분석자 판단||
66|Triple Match 3D|Boombox Games LTD|퍼즐|타일·마작·트리플 매치|새 게임 — 게임명 기반 분석자 판단||
67|Foodstars: Merge & Cook|Happibits|퍼즐|머지|1-1에 있는 분류 그대로||
68|Seaside Escape®: Merge & Story|Microfun Limited|퍼즐|머지|3-2에 있는 분류 그대로|Seaside Escape: Merge & Story|
69|Age of Origins:Tower Defense|Hong Kong Ke Mo software Co., Limited|전략|4X·SLG|새 게임 — 분석자 판단 (이름은 타워 디펜스지만 SLG로 봄)||Y
70|Hunting Sniper|SPARKS INFORMATION (SINGAPORE) PTE. LTD.|액션·슈팅|스나이퍼|3-2에 있는 분류 그대로|Hunting Sniper|
71|X-Clash: Survival Challenge|9z Games(HK)|전략|4X·SLG|2-2 멕시코 Top20에서 전략으로 분류 · 세부는 분석자 판단||Y
72|Z Route: Redemption|BUILDING-BLOCKS NETWORK TECHNOLOGY CO.,LIMITED|전략|4X·SLG|1-1에 있는 분류 그대로|Z Route: Redemption|
73|Sand Blocks: Drop Puzzle|Rollic Games|퍼즐|블록|1-1에 있는 분류 그대로||
74|Golf Clash - Golfing Simulator|Electronic Arts|스포츠|골프|3-2에 있는 분류 그대로|Golf Clash - Golfing Simulator|
75|PUBG MOBILE|Tencent Mobile International Limited|액션·슈팅|배틀로얄|3-2에 있는 분류 그대로|PUBG MOBILE|
76|Heart of Vegas - Casino Slots|Product Madness|카드·보드·카지노|소셜 카지노|3-2의 Slots: Heart of Vegas Casino와 같은 분류|Slots: Heart of Vegas Casino|
77|Triple Match City|Breeze Games L.L.C-FZ|퍼즐|타일·마작·트리플 매치|새 게임 — 게임명 기반 분석자 판단||
78|Castle Busters|Voodoo|하이퍼·하이브리드 캐주얼|하이브리드 캐주얼|1-1에 있는 분류 그대로||
79|Cash Frenzy™ - Slots Casino|SpinX Games Limited|카드·보드·카지노|소셜 카지노|3-2에 있는 분류 그대로|Cash Frenzy - Casino Slots|
80|Top Lords|GAME SPARK PTE. LTD.|전략|4X·SLG|2-2 일본·한국 Top20에서 전략으로 분류 · 세부는 분석자 판단||Y
81|Project Makeover|Magic Tavern, Inc.|퍼즐|매치3|새 게임 — 분석자 판단 (꾸미기 + 매치3)||
82|Top Force: Commander|Century Games Innovation Pte. Ltd.|전략|4X·SLG|새 게임 — 게임명 기반 분석자 판단||Y
83|Car Sort: Color Puzzle|Rollic Games|퍼즐|정렬·잼|새 게임 — 게임명 기반 분석자 판단||
84|Lotsa Slots™ Casino Slots Game|SpinX Games Limited|카드·보드·카지노|소셜 카지노|3-2에 있는 분류 그대로|Lotsa Slots - Casino Games|
85|Smash Fest!|Flow Games Bilisim Yazilim ve Pazarlama Anonim Sirketi|하이퍼·하이브리드 캐주얼|물리 파괴|1-1에 있는 분류 그대로||
86|Marble Sort!|Voodoo|퍼즐|정렬·잼|새 게임 — 게임명 기반 분석자 판단||
87|Royal Smash! - Physics Puzzle|Cypher Games|하이퍼·하이브리드 캐주얼|물리 파괴|1-1에 있는 분류 그대로|Royal Smash! - Physics Puzzle|
88|Words With Friends Word Game|Zynga Inc.|퍼즐|워드|새 게임 — 게임명 기반 분석자 판단||"""
ws = wb.create_sheet("App Store 매출 Top100")
head_row(ws, 1, ["순위", "게임명", "퍼블리셔", "대분류", "세부 장르", "분류 근거", "구글 매출 Top100 속 같은 게임 (분석자 대조)",
                 "구글 플레이 매출 순위 (수식)", "비고"])
rows_i = [l.split("|") for l in ios_g.split("\n")]
for i, (rk, n, p, d, e, why, gp, y) in enumerate(rows_i, start=2):
    put(ws, i, [int(rk), n, p, d, e, why, gp or "-", None, ""])
    ws.cell(i, 8).value = f"=IF(G{i}=\"-\",\"-\",IFERROR(INDEX('매출 Top100'!$A$2:$A$101,MATCH(G{i},'매출 Top100'!$B$2:$B$101,0)),\"-\"))"
    for c in (1, 8): ws.cell(i, c).alignment = Alignment(horizontal="center")
    if y == "Y":
        ws.cell(i, 9).value = "확인 필요"
        yel_row(ws, i, 4, 6); ws.cell(i, 9).fill = YEL
LI = 1 + len(rows_i)
for k in range(LI + 1, 102):
    put(ws, k, [k - 1, "확인 불가", "", "", "", "", "-", "-", "AppBrain 페이지에 88위까지만 나옴"])
    ws.cell(k, 1).alignment = Alignment(horizontal="center"); yel_row(ws, k, 2, 9)
for col, w in zip("ABCDEFGHI", [6, 34, 30, 20, 20, 36, 30, 12, 30]): ws.column_dimensions[col].width = w
ws.freeze_panes = "C2"
ws.auto_filter.ref = "A1:I101"
# 장르 비교 (수식)
head_row(ws, 1, ["대분류", "App Store 매출 수", "App Store 비율", "구글 매출 Top100 수 (3-2)", "구글 비율", "차이 (App Store-구글)"], c0=11)
for i, cat in enumerate(CATS, start=2):
    put(ws, i, [cat, f"=COUNTIF($D$2:$D$101,K{i})", f"=IF($L$13=0,\"\",L{i}/$L$13)",
                f"=COUNTIF('매출 Top100'!$D$2:$D$101,K{i})", f"=IF($N$13=0,\"\",N{i}/$N$13)", f"=IF(M{i}=\"\",\"\",M{i}-O{i})"], c0=11)
    for c in (13, 15, 16): ws.cell(i, c).number_format = "0.0%"
put(ws, 13, ["합계", "=SUM(L2:L12)", "", "=SUM(N2:N12)", "", ""], c0=11)
for c in range(11, 17): ws.cell(13, c).font = BB
put(ws, 15, ["두 스토어 매출 순위에 모두 있는 게임 수", "=COUNT(H2:H101)"], c0=11)
put(ws, 16, ["App Store 매출에만 있는 게임 수 (구글 Top100 밖)", "=COUNTIF(G2:G101,\"-\")-COUNTIF(B2:B101,\"확인 불가\")"], c0=11)
for col, w in zip("KLMNOP", [36, 12, 12, 14, 10, 14]): ws.column_dimensions[col].width = w
notes_i = [
 "출처: " + SRC_AB_IOS_GROSS + " (페이지 'Last updated: October 6, 2026', 2026-10-07 조회). 구글 매출 Top100(3-2)은 2026-10-04 기준이라 날짜가 2일 다름.",
 "AppBrain 페이지에는 88위까지만 나옴 → 89~100위는 '확인 불가'(노란 줄). 그래서 비율(M열)은 88개 기준, 구글은 100개 기준.",
 "장르 분류: 1-1·1-2·3-2에 이미 있는 게임은 같은 분류를 그대로 씀(F열에 어느 시트인지 적음). 새 게임 16개는 게임명·공개 정보 기반 분석자 판단 — 확신이 낮으면 노란 칸 + '확인 필요'.",
 "G열 '같은 게임' 대조는 스토어마다 이름이 조금 달라 분석자가 맞춘 것(예: 'Homescapes: Match 3 Games' = 'Homescapes'). H열은 그 이름으로 3-2에서 순위를 찾는 수식.",
 "실제 매출액은 비공개 — 순위만 비교함. 매출 순위가 높다 = 그날 그 스토어에서 결제 금액이 많았다는 뜻(정확한 금액은 모름).",
 "쉬운 요약: 아이폰(App Store)과 안드로이드(구글) 매출 순위에 같은 게임이 많이 겹침(L15). 나머지 비교는 K~P열 표를 보면 됨.",
]
for k, t in enumerate(notes_i):
    c = ws.cell(18 + k, 11, t); c.font = NOTE; c.alignment = Alignment(wrap_text=True, vertical="top")
ws.cell(18 + 1, 11).fill = YEL

# ---------- (3) 확인 필요 재검토 (새 시트) ----------
ws = wb.create_sheet("확인 필요 재검토")
ws["A1"] = "'확인 필요' 노란 칸 재검토 — 확인한 사실·고칠 제안과, 같은 부류끼리 원래 칸에 반영한 결과 (2026-10-07)"; ws["A1"].font = TITLE
ws["A2"] = ("C열 '지금 값'은 원래 칸을 수식으로 보여 줌(원래 칸을 고치면 여기도 바뀜). G열 '고칠 제안'은 분석자 제안이며, 사용자 승인으로 I열에 적은 대로 원래 칸에 반영함(묶음 A~F). "
            "출처가 공식 스토어가 아닌 2차 사이트면 결과를 '부분 확인'으로 낮춰 적음.")
ws["A2"].font = NOTE
head_row(ws, 4, ["번호", "시트 · 칸", "지금 값 (수식)", "게임 / 항목", "확인한 사실 (쉬운 말)", "확인 결과", "고칠 제안 (분석자 제안)", "출처 (조회일 2026-10-07)", "원래 칸 반영 결과 (6단계-4)"])
AS, GP, CT, SEA2, GR, BM = "App Store Top100", "Google Play Top100", "국가별 Top20", "동남아 분석", "매출 Top100", "벤치마크 근거"
def cur(sn, cells):
    return "=" + "&\" / \"&".join(f"'{sn}'!{a}" for a in cells)
AB = "https://www.appbrain.com/app/"
rv = [
 ("1-1 App Store Top100 · 6행", cur(AS, ["D6", "E6", "F6"]), "Colony Flow!",
  "상자를 5칸 자리에 놓으면 같은 색 개미가 나와 픽셀 그림의 같은 색 블록을 나름(스토어 태그: Puzzle·Logic).", "부분 확인",
  "'정렬·잼' 유지 가능. 같은 방식인 Pixel Flow!는 '픽셀·컬러 아트'로 분류돼 있어 둘을 한 세부 장르로 맞출지 결정 필요.",
  "Google Play https://play.google.com/store/apps/details?id=com.abi.colony.flow"),
 ("1-1 App Store Top100 · 28행", cur(AS, ["D28", "E28", "F28"]), "CrownCards - Trading Cards",
  "돈을 넣고 디지털 카드 팩을 열면 진짜 트레이딩 카드가 나오고, 실물 배송이나 출금도 가능하다고 설명됨. 개발사 Crown Cards LLC.", "부분 확인",
  "대분류 그대로. 세부 장르는 '트레이딩 카드'보다 실물 경품·현금이 걸린 '리얼머니' 성격 — 비고에 적을지 검토.",
  "game-solver.com 정리 페이지 https://game-solver.com/crowncards-trading-cards/ (2차 출처)"),
 ("1-1 App Store Top100 · 40행", cur(AS, ["C40", "F40"]), "PlayCash Rewards (퍼블리셔)",
  "'PlayCash'라는 리워드 앱 기사(2022-06-21, Good Gamer Entertainment)는 있으나 차트 속 앱과 같은 앱인지 확인 안 됨.", "확인 불가",
  "그대로 둠.", "mugglehead.com 기사 https://mugglehead.com/1365226-2-playcash/"),
 ("1-1 App Store Top100 · 47행", cur(AS, ["D47", "E47", "F47"]), "NoomiClone",
  "공개 검색으로 앱 정보를 찾지 못함.", "확인 불가", "그대로 둠.", "웹 검색 (결과 없음)"),
 ("1-1 App Store Top100 · 69행", cur(AS, ["D69", "E69", "F69"]), "82-0.com",
  "스토어 분류 '게임: 스포츠'. 무작위 팀·시대에서 농구 선수를 뽑아 5명 라인업으로 82전 전승에 도전하는 게임.", "확인됨",
  "'스포츠 / 기타' 유지, 비고에 '농구 라인업 짜기' 추가 제안.", "AppGoblin https://appgoblin.info/apps/6780195090 (App Store 정보 정리)"),
 ("1-1 App Store Top100 · 72행", cur(AS, ["D72", "E72", "F72"]), "Bloons Blitz",
  "사방에서 몰려오는 풍선(적)을 영웅 하나로 버티는 '탄막 생존(bullet heaven)' 게임. 2026-05-19 필리핀 선출시.", "부분 확인",
  "'아케이드' 유지 가능. 탄막 생존 장르를 액션·슈팅으로 볼지 기준을 정하면 1-2 101행과 함께 맞춤.",
  "Bloons Wiki https://www.bloonswiki.com/Bloons_Blitz (팬 위키, 2차 출처)"),
 ("1-1 App Store Top100 · 100행", cur(AS, ["C100", "F100"]), "Imposter Game - Party Edition (퍼블리셔)",
  "App Store 판매자 'Sven Vucak', 카테고리 '단어'. 3명 이상이 한 기기로 하는 오프라인 '거짓말쟁이 찾기' 파티 게임.", "확인됨",
  "C100 퍼블리셔 '확인 불가' → 'Sven Vucak'. 세부 장르 '소셜 디덕션' 유지.", "App Store https://apps.apple.com/us/app/-/id6745120053"),
 ("1-2 Google Play Top100 · 11행", cur(GP, ["D11", "E11", "F11"]), "MemeMix Challenge: Funny Party",
  "마이크를 켜고 하는 'AI 음성 파티 게임'. 개발사 Amobear VN Game Global.", "확인됨",
  "'파티·음악 / 밈 챌린지' 유지 — 노란색 해제 가능.", AB + "mememix-challenge-funny-party/com.mimik.challenge.fun"),
 ("1-2 Google Play Top100 · 22행", cur(GP, ["D22", "E22", "F22"]), "Colony Flow!",
  "1번과 같은 게임.", "부분 확인", "1번과 함께 결정.", "1번과 같음"),
 ("1-2 Google Play Top100 · 46행", cur(GP, ["D46", "E46", "F46"]), "CarLoop",
  "색 블록을 트랙에 올리면 자동으로 차에 나뉘어 담기는 '정렬 퍼즐'(스토어 분류 Puzzle). 개발사 UUTEAM Technology.", "확인됨",
  "'퍼즐 / 정렬·잼' 유지 — 노란색 해제 가능.", AB + "carloop/com.blast.puzzlegame.bear"),
 ("1-2 Google Play Top100 · 88행", cur(GP, ["D88", "E88", "F88"]), "Lordrush",
  "스토어 분류 Strategy. '중세 타워 디펜스 전략 게임' — 요새 재건·자원·영웅·라이벌 영주.", "부분 확인",
  "'4X·SLG' 유지 가능(같은 회사 Kingshot도 4X·SLG로 분류). 다만 설명에 연맹·영토 전쟁 문구는 없음.",
  "game-solver.com https://game-solver.com/lordrush/ (2차 출처)"),
 ("1-2 Google Play Top100 · 101행", cur(GP, ["D101", "E101", "F101"]), "Bloons Blitz",
  "6번과 같은 게임.", "부분 확인", "6번과 함께 결정.", "6번과 같음"),
 ("2-2 국가별 Top20 · 26행", cur(CT, ["G26"]), "Kumquat Hollow Fable (브라질)",
  "지금 브라질 무료 Top100에 없고, 공개 검색으로 정보를 찾지 못함.", "확인 불가", "그대로 둠.",
  "AppBrain 브라질 무료 순위 https://www.appbrain.com/stats/google-play-rankings/top_free/game/br"),
 ("2-2 국가별 Top20 · 82행", cur(CT, ["G82"]), "FORTBLOX: Mobile Fortress GO! (한국 1위)",
  "'모바일 오토배틀러'(HYPERRISE). 2026-09-22 글로벌 출시.", "확인됨",
  "대분류 '전략' 제안(자동 전투 + 배치 전략). 기존 분류표에 오토배틀러 칸이 없어 '전략'에 넣는 것이 가장 가까움.",
  "Inven Global 2026-09-21 https://www.invenglobal.com/articles/26352/hyperrise-launches-fortblox-mobile-fortress-go-globally"),
 ("2-2 국가별 Top20 · 94행", cur(CT, ["G94"]), "Wild Water World (한국)",
  "SLG(전략) 방식을 섞은 '뗏목 생존·건설 게임'. 개발 EWORLD(광저우).", "확인됨",
  "대분류 '전략' 제안.", "FoxData 블로그 2026-10-01 https://foxdata.com/cn/blogs/wild-water-world-how-eworld-hit-1-in-japan-korea-taiwan/"),
 ("2-7 동남아 분석 · 7행", cur(SEA2, ["G7", "I7"]), "Easy Drum - Real Music Drum",
  "스토어 분류 '음악(게임)'. 색에 맞춰 가상 드럼을 치는 리듬 게임.", "확인됨",
  "'파티·음악' 유지, 비고 '확인 필요 — 악기 연주 앱' → '리듬 게임(스토어 분류 음악)'.", AB + "easy-drum-real-music-drum/easy.drum.real.music.drum.set"),
 ("2-7 동남아 분석 · 10행", cur(SEA2, ["G10", "I10"]), "Fruit Wins",
  "스토어 분류 Casual, 설명은 'Blast-and-Win' 한 줄뿐. 이용자 리뷰에 '돈을 번다/출금이 안 된다'는 주장이 있음(리뷰는 검증 불가).", "부분 확인",
  "게임 방식은 여전히 확인 불가 → 그대로. 비고에 '현금 보상 주장이 있는 앱(리뷰 기준)' 메모 추가 검토.", AB + "fruit-wins/com.fruitwins"),
 ("2-7 동남아 분석 · 13행", cur(SEA2, ["G13", "I13"]), "Water Go",
  "색깔 물을 관에 옮겨 한 관에 한 색만 남기는 '물 정렬' 퍼즐.", "확인됨",
  "대분류 '퍼즐' 제안(세부: 정렬·잼).", "Industry.co.id 2026-09-04 https://www.industry.co.id/read/155381/cara-menyelesaikan-level-susah-di-water-go-tanpa-hint"),
 ("2-7 동남아 분석 · 62행", cur(SEA2, ["C62"]), "888XH… / 77RT… 이름의 앱들",
  "888XHSukari = 4x4 스도쿠(개발 KESAR SINGH), 888XHPearl = 타이밍 맞춰 누르기(개발 PRIME CONSUMER CORPORATION), 77RTPabu 개발 KK INNOVATION SOLUTION SDN. BHD. 스토어 설명상 단순 캐주얼 게임.",
  "부분 확인", "설명상 하는 일은 확인됨. 개발사가 서로 다른데 이름 규칙이 같은 이유는 확인 불가 → 그대로.",
  AB + "888xhsukari/com.qhezf.sukari · " + AB + "888xhpearl/com.northel.orbit"),
 ("3-2 매출 Top100 · 24행", cur(GR, ["D24", "E24", "F24"]), "Last Asylum: Plague",
  "스토어 분류 Strategy. 역병이 퍼진 세상에서 마을을 지키는 생존 게임(37GAMES).", "부분 확인",
  "'4X·SLG' 유지 가능. 같은 회사 Z Route도 같은 분류.", AB + "last-asylum-plague/com.phs.global"),
 ("3-2 매출 Top100 · 35행", cur(GR, ["D35", "E35", "F35"]), "Sword x Staff",
  "스토어 분류 Role Playing. 'Third Way RPG'라고 소개.", "부분 확인",
  "대분류 RPG는 맞음. '방치형'인지는 설명에서 확인 안 됨 → 그대로.", AB + "sword-x-staff/com.zjcs.android.us"),
 ("3-2 매출 Top100 · 40행", cur(GR, ["D40", "E40", "F40"]), "Yarn Loop",
  "실패(보빈)를 컨베이어에 올려 같은 색 뜨개 칸을 모으고, 남으면 선반에 보관 — 색 정렬 퍼즐.", "확인됨",
  "'정렬·잼' 유지 — 노란색 해제 가능.", AB + "yarn-loop/com.combo.yarnflow"),
 ("3-2 매출 Top100 · 45행", cur(GR, ["D45", "E45", "F45"]), "Coin Master - Board Adventure",
  "주사위를 굴려 보드를 돌며 코인·습격·방어 칸을 밟는 게임.", "확인됨",
  "'보드·주사위 소셜캐주얼' 유지 — 노란색 해제 가능.", AB + "coin-master-board-adventure/com.moonactive.cmboard"),
 ("3-2 매출 Top100 · 49행", cur(GR, ["D49", "E49", "F49"]), "All in Hole: Black Hole Games",
  "블랙홀을 움직여 물건을 삼키는 io 스타일 게임. 스토어 분류는 Puzzle.", "확인됨",
  "'하이브리드 캐주얼' 유지 가능(스토어 분류 Puzzle이라는 점만 비고에).", AB + "all-in-hole-black-hole-games/com.homagames.studio.allinhole"),
 ("3-2 매출 Top100 · 50행", cur(GR, ["D50", "E50", "F50"]), "Z Route: Redemption",
  "스토어 분류 Strategy. 좀비 종말 세계의 생존 전략(37GAMES).", "부분 확인", "'4X·SLG' 유지 가능.", AB + "com.zroute.global"),
 ("3-2 매출 Top100 · 54행", cur(GR, ["D54", "E54", "F54"]), "Matching Story - Puzzle Games",
  "스토어 분류 Casual. 매치3·타일 매치·머지·스토리를 한 게임에 섞었다고 소개.", "부분 확인",
  "지금 분류 유지 가능, 비고에 '혼합형(매치3·타일·머지)' 추가 제안.", AB + "com.joycastle.mergematch"),
 ("3-2 매출 Top100 · 67행", cur(GR, ["D67", "E67", "F67"]), "Colony Flow!", "1번과 같은 게임.", "부분 확인", "1번과 함께 결정.", "1번과 같음"),
 ("3-2 매출 Top100 · 69행", cur(GR, ["D69", "E69", "F69"]), "Lordrush", "11번과 같은 게임.", "부분 확인", "11번과 함께 결정.", "11번과 같음"),
 ("3-2 매출 Top100 · 73행", cur(GR, ["D73", "E73", "F73"]), "Hollywood Merge",
  "스토어 분류 Casual. 'Merge! Fashion! Drama!' — 머지 + 스토리.", "확인됨", "'머지' 유지 — 노란색 해제 가능.", AB + "com.hollywood.merge"),
 ("3-2 매출 Top100 · 91행", cur(GR, ["D91", "E91", "F91"]), "Police Chief",
  "스토어 분류 Strategy. 작은 경찰서를 넓히고 부서·경찰을 키우는 게임(패키지 이름 com.slg.policewar).", "부분 확인",
  "'4X·SLG' 유지 가능. 경영(타이쿤) 성격도 있음.", AB + "com.slg.policewar"),
 ("3-2 매출 Top100 · H20 메모", cur(GR, ["H20"]), "App Store 매출 순위 (다음 단계 확인 필요)",
  "AppBrain App Store 미국 매출 게임 순위로 새 시트를 만듦.", "해결됨",
  "'3-4 App Store 매출 Top100' 시트 참고. 원래 메모 칸은 그대로 둠.", SRC_AB_IOS_GROSS),
 ("4-1 벤치마크 근거 · E27", cur(BM, ["E27"]), "카지노 결제 전환율의 기간 (AppsFlyer)",
  "원문·요약 페이지를 다시 열려 했으나 이번에는 열리지 않음.", "확인 불가", "그대로 둠.", "https://www.appsflyer.com/resources/reports/app-marketing-monetization-report/"),
]
A_ = "반영: 비고를 확인 내용으로 바꾸고 노란색 해제 (묶음 A)"
E_ = "반영: 비고만 확인 내용으로 바꿈, 세부 장르는 분석자 판단이라 노란색 유지 (묶음 E)"
F_ = "그대로 둠 — 확인 못 함 (묶음 F)"
FA = "반영: 확인한 사실만 비고에 덧붙임, 노란색 유지 (묶음 F)"
CF = "반영: 세부 장르 정렬·잼 → 픽셀·컬러 아트, 노란색 해제 (묶음 D — Pixel Flow!와 같은 부류)"
LR_ = "반영: 비고에 'Kingshot과 같은 기준' 적음, 노란색 유지 (묶음 D)"
BL = "반영: 비고에 '탄막 생존, 아케이드 유지' 적음, 노란색 유지 (묶음 D)"
applied = [CF, E_, F_, F_, A_, BL, "반영: 퍼블리셔 '확인 불가' → 'Sven Vucak', 노란색 해제 (묶음 C)", A_, CF, A_, LR_, BL, F_,
           "반영: 대분류 '확인 필요' → '전략', 노란색 해제 (묶음 B)", "반영: 대분류 '확인 필요' → '전략', 노란색 해제 (묶음 B)", A_, FA,
           "반영: 대분류 '확인 필요' → '퍼즐', 노란색 해제 (묶음 B)", FA, E_, E_, A_, A_, A_, E_, E_, CF, LR_, A_, E_,
           "그대로 둠 — 새 시트 3-4로 해결, 원래 메모 칸은 유지", F_]
assert len(applied) == len(rv), (len(applied), len(rv))
for k, row in enumerate(rv, start=1):
    r = 4 + k
    put(ws, r, [k] + list(row) + [applied[k - 1]], wrap=True)
    res = row[4]
    ws.cell(r, 6).fill = GREEN if res in ("확인됨", "해결됨") else YEL
    ws.cell(r, 7).fill = YEL
    ws.cell(r, 9).fill = GREEN if applied[k - 1].startswith("반영") else YEL
LR = 4 + len(rv)
head_row(ws, 4, ["확인 결과", "개수 (수식)"], c0=11)
for i, k in enumerate(["확인됨", "부분 확인", "확인 불가", "해결됨"], start=5):
    put(ws, i, [k, f"=COUNTIF($F$5:$F${LR},K{i})"], c0=11)
put(ws, 9, ["합계", "=SUM(L5:L8)"], c0=11)
put(ws, 11, ["원래 칸에 반영한 곳", f"=COUNTIF($I$5:$I${LR},\"반영*\")"], c0=11)
put(ws, 12, ["그대로 둔 곳", f"=COUNTIF($I$5:$I${LR},\"그대로*\")"], c0=11)
put(ws, 13, ["노란색을 해제한 곳", f"=COUNTIF($I$5:$I${LR},\"*노란색 해제*\")"], c0=11)
ws.cell(15, 11, "초록 = 확인됨/해결됨 또는 반영함, 노란 = 부분 확인/확인 불가, 분석자 제안 또는 그대로 둠").font = NOTE
for col, w in zip("ABCDEFGHIJKL", [6, 24, 30, 26, 52, 11, 48, 46, 40, 2, 20, 10]): ws.column_dimensions[col].width = w
ws.freeze_panes = "A5"

# ---------- (4) 개요 맨 아래 2줄 추가 ----------
g = wb["개요"]
nr = g.max_row + 1
for k, (a, b_) in enumerate([
    ("6단계-4 출처 (2026-10-07 조회)", "App Store 매출 순위: AppBrain App Store Top Grossing Games 미국 (페이지 갱신 2026-10-06, 88위까지 표시). 매출-다운로드 비교 F·G열: AppBrain 앱 상세 페이지(Google Play 다운로드 표시·등록 월) 172개 전부, 그중 3개는 Google Play 상세 페이지 직접 확인. 확인 필요 재검토: Google Play·App Store·AppBrain 앱 상세 + 2차 출처(게임 위키·정리 사이트·기사, 시트에 표시)."),
    ("6단계-4 한계", "App Store 매출 89~100위는 확인 불가(페이지에 없음). App Store 매출(10-06)과 구글 매출(10-04)은 날짜가 2일 다름. AppBrain이 옮겨 적은 설치 구간은 Google Play보다 늦을 수 있음(2026년 출시 10개 대조 중 3개 다름, 노란 칸). Tiki Smash!의 등록 월은 확인 불가. 2-6 표 B·C(대륙별 게이머 OS 비율·스토어별 매출)와 동남아 4개국 Top20은 다시 찾아봤지만 공개 자료가 없어 그대로 '확인 불가'. '확인 필요' 칸은 같은 부류끼리 묶어 원래 칸에 반영(0-2 I열) — 세부 장르가 분석자 판단인 칸은 노란색 유지.")]):
    g.cell(nr + k, 1, a).font = BB; g.cell(nr + k, 1).border = BD
    c = g.cell(nr + k, 2, b_); c.font = B; c.border = BD; c.alignment = Alignment(wrap_text=True, vertical="top")

# (예전에는 여기서 시트 28개짜리 파일 1개를 저장했음 → 이제 아래 [2부]에서 파일 2개로 나눠 저장)


# ===== [2부] 파일 나누기 =====
# 아래부터는 [1부]에서 만든 시트를 파일 2개로 나눕니다. 시트 내용은 건드리지 않고
# ① 이름 앞에 그룹 번호 붙이기 ② 바뀐 이름에 맞춰 수식·차트·조건부서식·드롭다운의 시트 이름 고치기
# ③ 탭 색 ④ 목차 시트 ⑤ '개요'의 '시트 구성' 한 줄 갱신 만 합니다.
import os, re
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.hyperlink import Hyperlink

OUT_DIR = "/mnt/user-data/outputs"
FILE_A_NAME = "시장조사.xlsx"
FILE_B_NAME = "게임전략_MVP_마케팅.xlsx"

# 그룹: (번호, 그룹 이름, 탭 색)
GROUPS_A = {0: ("안내", "404040"), 1: ("미국 차트", "2F5597"), 2: ("지역", "548235"),
            3: ("매출", "C55A11"), 4: ("벤치마크", "7030A0")}
GROUPS_B = {0: ("안내", "404040"), 1: ("성공 비결", "C00000"), 2: ("MVP", "1F7A7A"),
            3: ("마케팅", "BF9000"), 4: ("참조용 복사본", "A6A6A6")}

# (그룹, 기존 시트 이름 또는 None=예정, 새 시트 이름, 한 줄 설명, 비고)
FILE_A = [
    (0, "개요", "0-1 개요", "분석 범위·기준일·출처 모음", ""),
    (0, "확인 필요 재검토", "0-2 확인 필요 재검토", "노란 '확인 필요' 칸 32곳을 다시 확인한 결과·고칠 제안과, 같은 부류끼리 원래 칸에 반영한 결과", "6단계-4에서 만듦"),
    (1, "App Store Top100", "1-1 App Store Top100", "미국 App Store 무료 게임 1~100위와 장르 분류 (기준일 2026-10-03)", ""),
    (1, "Google Play Top100", "1-2 Google Play Top100", "미국 구글 플레이 무료 게임 1~100위와 장르 분류 (기준일 2026-10-04)", "게임전략 파일 4-4에 참조용 복사본 있음"),
    (1, "장르 분포", "1-3 장르 분포", "두 스토어 Top100의 대분류별 게임 수·비율 (수식 자동 집계)", ""),
    (1, "세부 장르", "1-4 세부 장르", "대분류 안 세부 장르별 게임 수 (수식 자동 집계)", "6단계-3: 매출 Top100의 새 세부 장르 5개를 아래쪽 별도 표로 추가"),
    (1, "인사이트", "1-5 인사이트", "미국 차트에서 찾은 핵심 발견", ""),
    (1, "심층 분석 후보", "1-6 심층 분석 후보", "더 깊게 볼 게임 후보와 고른 이유", ""),
    (2, "지역 비교 요약", "2-1 지역 비교 요약", "미국·캐나다 / 남미 / 아시아 / 동남아(인도네시아) 시장 한눈에 비교", "6단계-2에서 동남아 열(E) 추가함"),
    (2, "국가별 Top20", "2-2 국가별 Top20", "캐나다·브라질·멕시코·일본·한국 구글 플레이 무료 Top20", ""),
    (2, "국가별 장르 분포", "2-3 국가별 장르 분포", "국가별 Top20의 장르 집계 (수식 자동 집계)", ""),
    (2, "공통 인기 게임", "2-4 공통 인기 게임", "여러 나라에서 함께 순위권에 오른 게임", ""),
    (2, "지역별 특징", "2-5 지역별 특징", "지역·국가별 특징과 근거 게임", ""),
    (2, "대륙별 OS 분포", "2-6 대륙별 OS 분포", "대륙·국가별 Android/iOS 점유율, 스토어별 게임 매출·다운로드, OS 출시 우선순위(분석자 제안)", "국가 12곳 포함 (StatCounter 2026-09) · 게임전략 파일 4-7에 참조용 복사본 있음"),
    (2, "동남아 분석", "2-7 동남아 분석", "인도네시아 Top20·장르 집계·공통 게임·특징 + 동남아 벤치마크(2025 보고서)", "태국·베트남·필리핀·말레이시아 순위는 확인 불가"),
    (3, "매출 순위 원자료", "3-1 매출 순위 원자료", "미국·브라질·일본·한국 구글 플레이 매출 Top10 원자료", ""),
    (3, "매출 Top100", "3-2 매출 Top100", "미국 구글 플레이 매출 Top100 장르 분류 (기준일 2026-10-04) + 무료 Top100과 장르 비교", "게임전략 파일 4-5에 참조용 복사본 있음"),
    (3, "매출-다운로드 비교", "3-3 매출-다운로드 비교", "매출 순위 vs 무료 순위·추정 설치 수 구간, 4가지 유형 자동 분류", "게임전략 파일 4-6에 참조용 복사본 있음"),
    (3, "App Store 매출 Top100", "3-4 App Store 매출 Top100", "미국 App Store 매출 게임 순위(88위까지)·장르 분류 + 구글 매출 Top100과 장르·순위 비교 (수식)", "6단계-4에서 만듦 · 89~100위 확인 불가"),
    (4, "벤치마크 근거", "4-1 벤치마크 근거", "리텐션·수익 등 공개 벤치마크 수치와 출처", ""),
    (4, "지역 벤치마크", "4-2 지역 벤치마크", "지역별 지표 중앙값·상위 구간 표", "게임전략 파일 4-1에 참조용 복사본 있음"),
    (4, "장르 차트 판독값", "4-3 장르 차트 판독값", "GameAnalytics 2025 장르별 차트를 눈으로 읽은 값 (판독값)", "게임전략 파일 4-2에 참조용 복사본 있음"),
    (4, "장르 판독 요약", "4-4 장르 판독 요약", "판독값의 12개월 평균과 장르 순위 (수식 자동 계산)", "게임전략 파일 4-3에 참조용 복사본 있음"),
    (4, "제외한 자료", "4-5 제외한 자료", "출처가 약해 쓰지 않은 자료와 이유", ""),
]
FILE_B = [
    (0, "개요", "0-1 개요", "분석 범위·기준일·출처 모음 (시장조사 파일과 같은 내용)", ""),
    (1, "점수 기준", "1-1 점수 기준", "성공 점수 배점 가정값 (바꾸면 점수표가 자동 재계산)", ""),
    (1, "성공 점수표", "1-2 성공 점수표", "12개 게임의 순위 기반 성공 점수 (수식 자동 계산)", ""),
    (1, "성적표 카드", "1-3 성적표 카드", "게임별 쉬운 설명과 별점 카드", ""),
    (1, "성공 비결 해부", "1-4 성공 비결 해부", "12개 게임의 재미 루프·첫 1분·다시 오게 하는 장치", ""),
    (1, "성공 비결 확장", "1-5 성공 비결 확장", "장르 그룹별로 넓힌 게임 목록과 차트 근거", ""),
    (1, "공통점 분석", "1-6 공통점 분석", "26개 게임의 공통 요소 표와 요약", ""),
    (1, "성공 공식 TOP5", "1-7 성공 공식 TOP5", "공통점에서 뽑은 성공 공식 5가지", ""),
    (2, "MVP 명세", "2-1 MVP 명세", "게임별 꼭 필요/있으면 좋음/나중에 기능", ""),
    (2, "MVP 검증 목표", "2-2 MVP 검증 목표", "리텐션 합격선·목표 (4-1 지역 벤치마크(참조)에서 가져옴)", ""),
    (2, "장르 연결(추정)", "2-3 장르 연결(추정)", "우리 장르 그룹 ↔ GameAnalytics 장르 연결 (4-3에서 가져옴)", ""),
    (2, "2인 개발 MVP 판단", "2-4 2인 개발 MVP 판단", "매출 Top100 + ③ 다운로드만 높은 게임 중 2명이 만들 수 있는 MVP 후보 (세부 장르 점수 → 게임 판정, 수식)", "6단계-3에서 만듦"),
    (3, "광고 대비 유입", "3-1 광고 대비 유입", "광고비 → 설치 수 → D1/D7/D30에 남는 사람 수 계산표 (CPI 출처표 + 4-1 리텐션)", "6단계-3에서 만듦"),
    (3, "광고 채널 분석", "3-2 광고 채널 분석", "유튜브·인스타·페이스북·틱톡·인플루언서 비교 (공개 자료 + 2인 팀 점수표)", "6단계-3에서 만듦"),
    (3, "후킹 영상 분석", "3-3 후킹 영상 분석", "Royal Match·MONOPOLY GO!·Magic Sort! 메타 광고의 후킹 유형·형식 빈도 (AdWhispr 분류)", "6단계-3에서 만듦 · 처음 3초는 메타 광고 라이브러리 화면 판독(분석자)"),
    (3, "마케팅 실행안", "3-4 마케팅 실행안", "우리 MVP용 단계별 지역·OS·채널·예산·소재·측정 지표 (분석자 제안)", "6단계-3에서 만듦"),
    (3, "광고 소재 해부", "3-5 광고 소재 해부", "영상 12개의 도파민(보상·긴장) 요소·영상 효과·구성 판독 + 사운드 체크리스트(사용자 기입)", "6단계-3에서 만듦"),
    (3, "광고 MVP 구성표", "3-6 광고 MVP 구성표", "우리 정렬 퍼즐용 15초·30초 광고 구간별 구성안 (분석자 제안)", "6단계-3에서 만듦"),
    (4, "지역 벤치마크", "4-1 지역 벤치마크(참조)", "2-2 MVP 검증 목표가 쓰는 표의 복사본",
        "원본: 시장조사.xlsx › 4-2 지역 벤치마크"),
    (4, "장르 차트 판독값", "4-2 장르 차트 판독값(참조)", "4-3이 계산에 쓰는 판독값의 복사본",
        "원본: 시장조사.xlsx › 4-3 장르 차트 판독값"),
    (4, "장르 판독 요약", "4-3 장르 판독 요약(참조)", "2-3 장르 연결(추정)이 쓰는 표의 복사본",
        "원본: 시장조사.xlsx › 4-4 장르 판독 요약"),
    (4, "Google Play Top100", "4-4 Google Play Top100(참조)", "2-4가 세부 장르를 찾는 무료 Top100 표의 복사본",
        "원본: 시장조사.xlsx › 1-2 Google Play Top100"),
    (4, "매출 Top100", "4-5 매출 Top100(참조)", "2-4가 쓰는 매출 Top100 장르 표의 복사본",
        "원본: 시장조사.xlsx › 3-2 매출 Top100"),
    (4, "매출-다운로드 비교", "4-6 매출-다운로드 비교(참조)", "2-4가 쓰는 매출·무료 순위와 4가지 유형 분류의 복사본",
        "원본: 시장조사.xlsx › 3-3 매출-다운로드 비교"),
    (4, "대륙별 OS 분포", "4-7 대륙별 OS 분포(참조)", "3-4가 쓰는 OS 점유율·OS 출시 우선순위 표의 복사본",
        "원본: 시장조사.xlsx › 2-6 대륙별 OS 분포"),
]

def sheet_list_text(spec, groups, other_file, other_desc):
    parts = []
    for gno in sorted(groups):
        names = [n for g, o, n, d, memo in spec if g == gno and o is not None]
        if names:
            parts.append(f"[{gno} {groups[gno][0]}] " + " · ".join(names))
    return " / ".join(parts) + f"  ※ {other_desc}은(는) '{other_file}' 파일에 있음 (파일 사이 수식 연결 없음)"

def fix_refs(text, rename):
    """수식 문자열 안의 '시트 이름'! 을 새 이름으로 바꿈 (가리키는 칸은 그대로)."""
    if not isinstance(text, str):
        return text
    def rep(m):
        name = m.group(1)
        return f"'{rename[name]}'!" if name in rename else m.group(0)
    return re.sub(r"'([^']+)'!", rep, text)

def split_workbook(wb, spec, groups, file_label, other_file, other_desc):
    keep = {o: n for g, o, n, d, memo in spec if o is not None}
    # 1) 이 파일에 안 들어가는 시트 지우기
    for ws in list(wb.worksheets):
        if ws.title not in keep:
            wb.remove(ws)
    # 2) 수식·조건부서식·드롭다운·차트 안의 시트 이름 고치기
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    c.value = fix_refs(c.value, keep)
        for cf in ws.conditional_formatting:
            for rule in cf.rules:
                if rule.formula:
                    rule.formula = [fix_refs(f, keep) for f in rule.formula]
        for dv in ws.data_validations.dataValidation:
            dv.formula1 = fix_refs(dv.formula1, keep)
            dv.formula2 = fix_refs(dv.formula2, keep)
        for ch in ws._charts:
            for s in ch.series:
                for part in (s.val, s.cat, s.xVal, s.yVal):
                    if part is None: continue
                    for ref in (getattr(part, "numRef", None), getattr(part, "strRef", None)):
                        if ref is not None and ref.f: ref.f = fix_refs(ref.f, keep)
                if s.tx is not None and s.tx.strRef is not None and s.tx.strRef.f:
                    s.tx.strRef.f = fix_refs(s.tx.strRef.f, keep)
    # 3) 시트 이름 바꾸기 + 탭 색
    gmap = {o: g for g, o, n, d, memo in spec if o is not None}
    for ws in wb.worksheets:
        old = ws.title
        ws.title = keep[old]
        ws.sheet_properties.tabColor = groups[gmap[old]][1]
    # 4) 개요의 '시트 구성' 한 줄만 새 이름으로 갱신
    g = wb["0-1 개요"]
    for row in g.iter_rows(min_col=1, max_col=2):
        if row[0].value == "시트 구성":
            row[1].value = sheet_list_text(spec, groups, other_file, other_desc)
    # 5) 참조용 복사본 표시 (빈 칸에 원본 위치 메모, 기존 칸은 건드리지 않음)
    for g_, o, n, d, memo in spec:
        if o is not None and memo.startswith("원본:"):
            ws = wb[n]
            col = ws.max_column + 2
            cell = ws.cell(row=1, column=col)
            assert cell.value is None
            cell.value = f"참조용 복사본 · {memo} · 원본을 고치면 코드를 다시 실행해 두 파일을 함께 갱신"
            cell.font = Font(name=F, bold=True, color="7F7F7F")
            cell.fill = PatternFill("solid", fgColor="EDEDED")
    # 6) 목차 시트
    toc = wb.create_sheet("0 목차", 0)
    toc.sheet_properties.tabColor = groups[0][1]
    toc["A1"] = f"{file_label} — 목차"
    toc["A1"].font = Font(name=F, bold=True, size=14)
    toc["A2"] = (f"이 파일: {file_label}  /  짝 파일: {other_file} ({other_desc})  ·  "
                 "시트 이름을 누르면 그 시트로 이동합니다. 회색 줄은 다음 단계에서 만들 '예정' 시트입니다.")
    toc["A2"].font = Font(name=F, color="595959")
    heads = ["그룹", "시트 이름", "한 줄 설명", "상태", "비고"]
    for i, h in enumerate(heads, 1):
        c = toc.cell(row=4, column=i, value=h); c.fill = H_FILL; c.font = H_FONT
    r = 5
    for gno, o, n, d, memo in [(0, "_", "0 목차", "이 시트", "")] + spec:
        gname, color = groups[gno]
        planned = o is None
        vals = [f"{gno} {gname}", n, d, "예정" if planned else "완료", memo]
        for i, v in enumerate(vals, 1):
            c = toc.cell(row=r, column=i, value=v)
            c.border = BD
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.font = Font(name=F, italic=planned, color="808080" if planned else "000000")
        toc.cell(row=r, column=1).fill = PatternFill("solid", fgColor=color)
        toc.cell(row=r, column=1).font = Font(name=F, bold=True, color="FFFFFF")
        if not planned and n != "0 목차":
            link = toc.cell(row=r, column=2)
            link.hyperlink = Hyperlink(ref=link.coordinate, location=f"'{n}'!A1")  # 파일 안 시트로 이동
            link.font = Font(name=F, color="0563C1", underline="single")
        r += 1
    for col, w in zip("ABCDE", [16, 30, 58, 8, 44]):
        toc.column_dimensions[col].width = w
    toc.freeze_panes = "A5"
    # 7) 순서 맞추기 (목차 → 개요 → 그룹 순)
    order = ["0 목차"] + [n for g_, o, n, d, memo in spec if o is not None]
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = (ws.title == "0 목차")
    return wb

# [1부] 코드를 한 번 더 실행해서 파일 B용 통합 문서를 따로 만듦
# (openpyxl은 통합 문서 복사 기능이 없어서, 같은 코드로 똑같은 시트를 한 번 더 만든 뒤 나눔)
def build_again():
    MARK = "# ===== [2부]" + " 파일 나누기 ====="
    with open(__file__, encoding="utf-8") as fh:
        part1 = fh.read().split(MARK)[0]
    ns = {}
    exec(compile(part1, __file__ + " [1부]", "exec"), ns)
    return ns["wb"]

wb_a = wb
wb_b = build_again()
split_workbook(wb_a, FILE_A, GROUPS_A, FILE_A_NAME, FILE_B_NAME, "성공 비결·MVP·마케팅 판단 자료")
split_workbook(wb_b, FILE_B, GROUPS_B, FILE_B_NAME, FILE_A_NAME, "시장·지역·매출·벤치마크 숫자 자료")
# 메모 문장 속 옛 시트 이름 고치기 (6단계-2, 사용자 승인) — 칸마다 (옛 글자, 새 글자)를 정확히 1번 바꿈
MEMO_FIX_A = {
    ("0-1 개요", "B11"): [("App Store / Google Play 시트", "'1-1 App Store Top100' / '1-2 Google Play Top100' 시트"), ("'장르 분포' 시트", "'1-3 장르 분포' 시트")],
    ("2-3 국가별 장르 분포", "A16"): [("'장르 분포' 시트", "'1-3 장르 분포' 시트")],
    ("4-1 벤치마크 근거", "G15"): [("'지역 벤치마크' 시트", "'4-2 지역 벤치마크' 시트")],
    ("4-1 벤치마크 근거", "G21"): [("'장르 차트 판독값' 시트", "'4-3 장르 차트 판독값' 시트")],
    ("4-4 장르 판독 요약", "A2"): [("'장르 차트 판독값' 시트", "'4-3 장르 차트 판독값' 시트")],
}
MEMO_FIX_B = {
    ("0-1 개요", "B11"): [("App Store / Google Play 시트", "시장조사.xlsx의 '1-1 App Store Top100' / '1-2 Google Play Top100' 시트"), ("'장르 분포' 시트", "'1-3 장르 분포' 시트(시장조사.xlsx)")],
    ("1-5 성공 비결 확장", "A17"): [("'성공 비결 해부'", "'1-4 성공 비결 해부'"), ("'성공 점수표'", "'1-2 성공 점수표'")],
    ("2-1 MVP 명세", "A15"): [("'MVP 검증 목표'", "'2-2 MVP 검증 목표'"), ("'장르 판독 요약'", "'4-3 장르 판독 요약(참조)'")],
    ("2-2 MVP 검증 목표", "A16"): [("'벤치마크 근거'·'지역 벤치마크' 시트", "시장조사.xlsx › '4-1 벤치마크 근거' 시트와 이 파일의 '4-1 지역 벤치마크(참조)' 시트")],
    ("2-2 MVP 검증 목표", "A17"): [("'지역 벤치마크' 조회 열", "'4-1 지역 벤치마크(참조)' 조회 열")],
    ("2-2 MVP 검증 목표", "A19"): [("'장르 판독 요약'·'장르 연결(추정)'", "'4-3 장르 판독 요약(참조)'·'2-3 장르 연결(추정)'")],
    ("2-3 장르 연결(추정)", "A16"): [("'MVP 검증 목표'", "'2-2 MVP 검증 목표'")],
    ("4-3 장르 판독 요약(참조)", "A2"): [("'장르 차트 판독값'", "'4-2 장르 차트 판독값(참조)'")],
}
def apply_memo_fix(wb, fixes):
    for (sn, addr), pairs in fixes.items():
        c = wb[sn][addr]
        t = c.value
        for old, new in pairs:
            assert isinstance(t, str) and t.count(old) == 1, (sn, addr, old)
            t = t.replace(old, new)
        c.value = t
apply_memo_fix(wb_a, MEMO_FIX_A)
apply_memo_fix(wb_b, MEMO_FIX_B)
os.makedirs(OUT_DIR, exist_ok=True)
wb_a.save(os.path.join(OUT_DIR, FILE_A_NAME))
wb_b.save(os.path.join(OUT_DIR, FILE_B_NAME))
print("저장 완료:", FILE_A_NAME, len(wb_a.worksheets), "시트 /", FILE_B_NAME, len(wb_b.worksheets), "시트")
