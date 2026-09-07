// FGI 결과 보고서 템플릿 생성 스크립트
// 기준 사례: 2025 더불어민주당 서울시당 서울시장 후보 인식 FGI (코리아스픽스)
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, WidthType, ShadingType, BorderStyle,
  LevelFormat, PageBreak, PageNumber, Footer, Header, TableOfContents,
  VerticalAlign,
} = require("docx");

// ───────────────────────────── 스타일 상수
const FONT = "맑은 고딕";
const C = {
  navy: "1F3864", blue: "2F5496", lightBlue: "EAF1FB", gray: "F2F2F2",
  midGray: "808080", line: "BFBFBF", quote: "FBF7EA", white: "FFFFFF",
  placeholder: "7F7F7F", example: "5B7DB1",
};
const PAGE_W = 11906, PAGE_H = 16838; // A4 (DXA)
const MARGIN = 1134; // 2cm
const CONTENT_W = PAGE_W - MARGIN * 2; // 9638

// ───────────────────────────── 텍스트 헬퍼
// 규칙: "[ ]"로 감싼 문장 = 채워 넣을 자리(회색 이탤릭)
//       "예) "로 시작 = 서울시장 FGI 사례(파란색 작은 글씨)
function runsOf(text, opt = {}) {
  const base = { font: FONT, size: opt.size || 20, bold: opt.bold, color: opt.color };
  if (/^예\)/.test(text)) {
    return [new TextRun({ ...base, text, size: (opt.size || 20) - 2, color: C.example, italics: true })];
  }
  // 대괄호 플레이스홀더를 회색 이탤릭으로 분리 렌더링
  const parts = text.split(/(\[[^\]]*\])/g).filter(Boolean);
  return parts.map((p) =>
    /^\[[^\]]*\]$/.test(p)
      ? new TextRun({ ...base, text: p, italics: true, color: C.placeholder, bold: false })
      : new TextRun({ ...base, text: p })
  );
}

function P(text, opt = {}) {
  return new Paragraph({
    alignment: opt.align || AlignmentType.LEFT,
    spacing: { before: opt.before ?? 40, after: opt.after ?? 40, line: opt.line ?? 300 },
    numbering: opt.bullet ? { reference: "bul", level: opt.level || 0 } : undefined,
    indent: opt.indent,
    shading: opt.shading,
    children: Array.isArray(text) ? text : runsOf(text, opt),
  });
}
const B = (text, opt = {}) => P(text, { ...opt, bullet: true });
const blank = (n = 1) => Array.from({ length: n }, () => new Paragraph({ spacing: { before: 0, after: 0 }, children: [] }));

function H1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 160 },
    children: [new TextRun({ text, font: FONT, size: 32, bold: true, color: C.navy })],
  });
}
function H2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2, spacing: { before: 280, after: 120 },
    children: [new TextRun({ text, font: FONT, size: 26, bold: true, color: C.blue })],
  });
}
function H3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, font: FONT, size: 22, bold: true, color: "000000" })],
  });
}
// 질문 제목(Q번호) — 굵게, 본문 크기
function Q(text) {
  return new Paragraph({
    spacing: { before: 200, after: 80 }, keepNext: true,
    children: runsOf(text, { bold: true, size: 21 }),
  });
}
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

// ───────────────────────────── 표 헬퍼
const border = (color = C.line) => ({ style: BorderStyle.SINGLE, size: 4, color });
const borders = (color) => ({ top: border(color), bottom: border(color), left: border(color), right: border(color) });

function cellParas(content, opt = {}) {
  const lines = Array.isArray(content) ? content : String(content).split("\n");
  return lines.map((l) => {
    const bullet = l.startsWith("- ");
    const t = bullet ? l.slice(2) : l;
    return new Paragraph({
      alignment: opt.align || AlignmentType.LEFT,
      spacing: { before: 20, after: 20, line: 276 },
      numbering: bullet ? { reference: "bul", level: 0 } : undefined,
      children: runsOf(t, { size: opt.size || 19, bold: opt.bold, color: opt.color }),
    });
  });
}

function cell(content, width, opt = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: opt.fill ? { type: ShadingType.CLEAR, fill: opt.fill, color: "auto" } : undefined,
    borders: borders(opt.borderColor),
    verticalAlign: opt.vAlign || VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    columnSpan: opt.span,
    children: cellParas(content, opt),
  });
}

// 표: headers(문자열 배열) + rows(문자열 배열의 배열) + widths(비율 배열)
function tbl(headers, rows, ratios, opt = {}) {
  const total = ratios.reduce((a, b) => a + b, 0);
  const widths = ratios.map((r) => Math.round((CONTENT_W * r) / total));
  widths[widths.length - 1] = CONTENT_W - widths.slice(0, -1).reduce((a, b) => a + b, 0); // 합계를 정확히 맞춤
  const headRow = headers
    ? new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => cell(h, widths[i], { fill: opt.headFill || C.gray, bold: true, align: AlignmentType.CENTER })),
      })
    : null;
  const bodyRows = rows.map(
    (r) =>
      new TableRow({
        children: r.map((c, i) =>
          cell(c, widths[i], {
            fill: i === 0 && opt.firstColFill !== false ? "F7F7F7" : undefined,
            bold: i === 0 && opt.firstColBold !== false,
            align: i === 0 && opt.firstColCenter !== false ? AlignmentType.CENTER : AlignmentType.LEFT,
          })
        ),
      })
  );
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    rows: [headRow, ...bodyRows].filter(Boolean),
  });
}

// 요약 박스: 연한 파랑 단일 셀, 각 줄은 "→"로 시작
function summaryBox(lines, opt = {}) {
  const content = lines.map((l) => (l.startsWith("→") || l.startsWith("예)") || l.startsWith("- ") ? l : "→ " + l));
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [cell(content, CONTENT_W, { fill: opt.fill || C.lightBlue, borderColor: "9DB7E0", size: 19 })] })],
  });
}

// 인용 박스: 패널 주요 발언
function quoteBox(quotes) {
  const content = quotes.map((q) => (q.startsWith("예)") ? q : q.startsWith("※") ? q : `“${q}”`));
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [cell(content, CONTENT_W, { fill: C.quote, borderColor: "E0D3A6", size: 19 })] })],
  });
}

// 섹션 배너: 진한 파랑 바탕 + 흰 굵은 글씨 (III장 주제 구분용)
function banner(text) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [cell(text, CONTENT_W, { fill: C.blue, color: C.white, bold: true, size: 22, borderColor: C.blue })] })],
  });
}

// 표지 상단 띠
function titleBand(text) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [cell(text, CONTENT_W, { fill: C.navy, color: C.white, bold: true, size: 26, align: AlignmentType.CENTER, borderColor: C.navy })] })],
  });
}

// 안내문(회색 상자): 서식 사용 지침
function note(lines) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [cell(lines, CONTENT_W, { fill: C.gray, borderColor: C.line, size: 18, color: "404040" })] })],
  });
}

// 그룹 축 (서울시장 FGI 기준 4그룹). 다른 사업에서는 그룹명만 바꿔 쓴다.
const G = ["[그룹 1]", "[그룹 2]", "[그룹 3]", "[그룹 4]"];
const groupRows = (fill = "[ ]") => G.map((g) => [g, fill]);

// ───────────────────────────── 문서 본문
const body = [];
const push = (...items) => items.forEach((i) => body.push(i));

// ========== 템플릿 표지
push(
  ...blank(6),
  titleBand("FGI 결과 보고서 작성 서식"),
  ...blank(2),
  P("집단심층면접(FGI) 종합보고서 · 그룹별 보고서 · 작성 가이드", { align: AlignmentType.CENTER, size: 28, bold: true, color: C.navy }),
  ...blank(1),
  P("기준 사례 : 2025 더불어민주당 서울시당 4050 권리당원 서울시장 후보 인식 FGI (4그룹 · 29명)", { align: AlignmentType.CENTER, size: 20, color: "404040" }),
  ...blank(12),
  P("코리아스픽스 (주) K-speaks Inc. · 숙의공론컨설팅본부", { align: AlignmentType.CENTER, size: 20 }),
  P("서식 v1.0 · 2026. 09.", { align: AlignmentType.CENTER, size: 18, color: C.midGray }),
  pageBreak()
);

// ========== 서식 사용 안내
push(
  H1("서식 사용 안내"),
  H3("1. 구성"),
  tbl(
    ["구분", "문서", "쓰임"],
    [
      ["Part A", "FGI 종합보고서 서식", "전체 그룹을 가로로 묶어 발주처에 제출하는 최종 보고서. 그룹 간 공통점과 차이를 한 표에서 비교한다."],
      ["Part B", "그룹(회차)별 보고서 서식", "회차 하나를 정리한 보고서. 종합보고서의 재료가 되며, 발주처가 요청하면 별도 제출한다."],
      ["Part C", "작성 가이드 · 체크리스트", "블록별 작성 규칙, 문체, 파일명, 제출 전 점검 항목."],
    ],
    [1, 2.2, 6.8]
  ),
  H3("2. 작성 순서"),
  tbl(
    ["단계", "산출물", "기준 일정(서울시장 FGI 사례)"],
    [
      ["① 녹취 로데이터", "회차별 녹취록(엑셀·한글). 발언자 불명확 구간은 후처리, 추임새·잡음은 (.) 처리", "실사 D+2 ~ D+7"],
      ["② 그룹별 보고서", "Part B 서식으로 회차마다 1부. 요약 박스 + 발언 인용 중심", "실사 종료 후 1주 이내 초안"],
      ["③ 종합보고서", "Part A 서식. 그룹별 보고서의 표를 가로로 이어 비교", "그룹별 초안 익일"],
      ["④ 발주처 포맷 변환", "요청 시 한글(.hwp) 변환, PPT 요약본 제작", "종합보고서 제출 당일~익일"],
    ],
    [1.6, 5.4, 3]
  ),
  H3("3. 표기 규칙"),
  tbl(
    ["표기", "의미"],
    [
      ["[ 대괄호 회색 글씨 ]", "채워 넣을 자리. 작성 후 반드시 삭제한다."],
      ["예) 파란 글씨", "서울시장 후보 FGI 보고서에서 가져온 실제 작성 예시. 참고 후 삭제한다."],
      ["→ 로 시작하는 연한 파랑 상자", "요약 박스. 해당 질문의 결론을 2~4줄로 먼저 적는다."],
      ["노란 상자 “ ”", "패널 주요 발언. 녹취 원문을 그대로 인용한다(개인정보·비속어만 정리)."],
      ["진한 파랑 띠", "III장의 주제(파트) 구분. 질문지의 파트 번호와 맞춘다."],
    ],
    [3.2, 6.8]
  ),
  pageBreak()
);

// ========== PART A. 종합보고서
push(
  titleBand("Part A. FGI 종합보고서 서식"),
  ...blank(1),
  note([
    "이 파트 전체가 종합보고서 한 부의 골격이다. 표지부터 부록까지 순서를 유지하고, 그룹 수(기본 4그룹)와 주제 수(기본 5개 + 마무리)는 사업에 맞게 늘리거나 줄인다.",
    "그룹 순서는 모든 표에서 동일하게(실사 순서 또는 발주처 관례 순서) 고정한다. 예) 50대 남성 → 50대 여성 → 40대 남성 → 40대 여성",
  ]),
  ...blank(1)
);

// A-0 표지
push(
  H2("A-0. 표지"),
  titleBand("[조사 주제] 관련 FGI 보고서"),
  ...blank(1),
  P("[발주처·대상] 심층 좌담회 (FGI) 종합보고서", { align: AlignmentType.CENTER, size: 28, bold: true }),
  P("예) 더불어민주당 서울시당 4050 권리당원 심층 좌담회 (FGI) 종합보고서", { align: AlignmentType.CENTER }),
  ...blank(1),
  B("패널 : [대상 정의] 총 [N]명   예) 민주당 4050 남녀 권리당원 총 29명"),
  tbl(["구분", "일시", "인원"], G.map((g) => [g, "[MM/DD 요일 HH:MM ~ HH:MM]", "[n]"]), [3, 5, 2]),
  P("예) 50대 남성 | 11/12 PM 2:00 ~ 5:00 | 6"),
  B("장소 : [장소명·호실]   예) 패스트파이브 시청1,2호점"),
  B("모더레이터 : [성명·직위·자격]   예) 이병덕 코리아스픽스 대표이사, fKF 마스터 퍼실리테이터"),
  B("모니터룸 : [참관·기록 인력]   예) 이소연 팀장, 정지혜 팀장, 이소영 대리, 공정호 fKF 1급 퍼실리테이터"),
  B("보고일 : [YYYY. MM. DD.]   |   작성 : 코리아스픽스 숙의공론컨설팅본부"),
  pageBreak()
);

// A-0 목차
push(
  H2("A-0. 목차"),
  note(["목차는 III장의 주제·질문 번호와 1:1로 맞춘다. 질문지 파트 번호를 그대로 쓰면 회차 보고서와 종합보고서의 번호가 어긋나지 않는다."]),
  ...blank(1),
  P("I. FGI 개요", { bold: true }),
  P("II. 요약 (Summary)", { bold: true }),
  P("III. 주요 논의 내용 (Main Discussion Points)", { bold: true }),
  P("   1. [주제 1]        1.1 [소항목]  /  1.2 [소항목]"),
  P("   2. [주제 2]        2.1 [소항목]  /  2.2 [소항목]  /  2.3 [소항목]"),
  P("   3. [주제 3 : 비교 대상 인식]   3.1 이미지·강점·약점  /  3.2 선호·비선호 이유·조언  /  3.3 기억에 남는 정책·행동  /  3.4 경쟁력·역량·기여도  /  3.5 한 단어 정의"),
  P("   4. [주제 4]        4.1 ~ 4.4"),
  P("   5. [주제 5 : 바라는 점]   5.1 ~ 5.2"),
  P("   6. 마무리"),
  P("IV. 시사점 및 제언 (선택)", { bold: true }),
  P("부록. 참가자 명단 · 질문지 · 녹취 로데이터 목록", { bold: true }),
  P("예) 1. 대통령 국정운영에 대한 인식 / 2. 내년 지방선거 전망 / 3. 후보 이미지(박주민, 전현희, 서영교) / 4. 서울시 현안 / 5. 차기 서울시장에게 바라는 점 / 6. 마무리"),
  pageBreak()
);

// I. 개요
push(
  H1("I. FGI 개요"),
  H3("1. 조사 개요"),
  tbl(
    ["항목", "내용"],
    [
      ["조사명", "[ ]   예) 민주당 서울시당 권리당원의 서울시장 후보군 인식 FGI"],
      ["목적", "[무엇을 누구에게 확인해 어디에 쓰는가]   예) 4050 서울 민주당 권리당원의 서울시장 후보 인식 및 서울시정 평가"],
      ["발주", "[발주처]"],
      ["수행", "코리아스픽스 숙의공론컨설팅본부 (설계·진행·분석·보고)"],
      ["기간", "[실사 기간 / 보고서 제출일]   예) 실사 2025.11.12 ~ 11.22 (4회) / 보고서 11.28"],
    ],
    [2, 8]
  ),
  H3("2. 아젠다"),
  summaryBox(["[한 문장으로 적는 조사의 핵심 질문]", "예) 민주당 서울시당 권리당원의 서울시장 후보군 인식"]),
  H3("3. 조사 설계"),
  tbl(
    ["구분", "내용"],
    [
      ["대상", "[모집 조건]   예) 4050 서울시 민주당 권리당원, 나이·성별·권역(강남북) 반영"],
      ["그룹 구성", "[그룹 수 × 그룹당 인원]   예) 총 4회, 각 7~8명 확정(실제 6~8명 참석)"],
      ["회당 시간", "[분]   예) 3시간 (도입 10분 · 인터뷰 160분 · 마무리 10분)"],
      ["장소", "[인터뷰룸 + 모니터룸]   예) 패스트파이브 시청1호점 11C 회의실 + 12A 모니터룸"],
      ["사례비", "[1인당 금액·지급 방식]   예) 1인당 10만원(현금)"],
      ["기록", "[녹음·녹화·기록 방식]   예) 녹음 2대, 모니터룸 실시간 기록, 녹취록 후처리"],
    ],
    [2, 8]
  ),
  H3("4. 주요 질문"),
  note(["질문지의 파트 번호와 순서를 그대로 옮긴다. III장의 Q번호(Q2-1, Q3-4 …)는 여기서 정한 파트 번호를 따른다."]),
  tbl(
    ["파트", "질문 요지"],
    [
      ["①", "[도입 : 개인 관심사·고민] & [대주제 1 인식과 근거]"],
      ["②", "[대주제 2] : [세부 질문 A] / [세부 질문 B] / [세부 질문 C]"],
      ["③", "[비교 대상 인식] : 강점·약점 / 외적 이미지·조언 / 좋아하는 점·싫어하는 점 / 기억에 남는 정책·행동 / 경쟁력·역량 / 떠오르는 단어"],
      ["④", "[현안 평가] : 불편한 점·개선점 / 비교 평가와 이유 / Good & Bad 정책 / 특정 정책에 대한 주변 평가와 본인 인식"],
      ["⑤", "[바라는 점] : 거주지 민원 / 바라는 덕목 / 바라는 정책"],
      ["⑥", "마무리 감사 인사"],
    ],
    [1, 9]
  ),
  P("예) ③ 후보 3인(박주민, 전현희, 서영교) 인식 — 강점과 약점 / 외적 이미지 & 조언 / 좋아하는 점, 싫어하는 점 / 경선 조언 / 기억에 남는 정책, 행동 / 12.3 내란 이후 인상적 활약상 / 서울시장 후보로서 경쟁력, 이후 업무역량 / 이재명 대통령 집권 안정에 가장 많은 도움을 줄 후보 / 세 후보하면 떠오르는 단어"),
  H3("5. 수행 체계"),
  tbl(
    ["역할", "담당", "하는 일"],
    [
      ["총괄(PM)", "[ ]", "일정·품질 총괄, 발주처 소통"],
      ["모더레이터", "[ ]", "설계 참여, 진행, 인사이트 도출"],
      ["모니터룸", "[ ]", "실시간 기록, 거수 집계, 참관 응대"],
      ["녹취·분석", "[ ]", "녹취록 정리, 회차 보고서 초안"],
      ["등록·운영", "[ ]", "참석 확인, 동의서·사례비, 다과"],
    ],
    [2, 3, 5]
  ),
  H3("6. 해석 시 유의사항"),
  B("집단심층면접은 표본을 대표하지 않는다. 빈도가 아니라 이유를 읽는 자료다."),
  B("그룹별 참석자가 6~8명이므로, 소수 의견이 그룹 전체의 인식으로 보이지 않게 쓴다. 한 명의 발언은 “(소수의견)”으로 표기한다."),
  B("거수·지목 결과는 “N표 중 M표”로 적고 인원 산정의 근거(참석 인원, 기권)를 함께 적는다."),
  pageBreak()
);

// II. 요약
push(
  H1("II. 요약 (Summary)"),
  H3("1. 패널 특성"),
  summaryBox([
    "[전체 그룹을 관통하는 특성 3줄 : ① 생활·관심사 ② 판단 기준·성향 ③ 조사 주제에 대한 기본 태도]",
    "예) 4050 권리당원 전반은 경제·주거·돌봄·일자리·노후 등 ‘생활 압박’이 가장 큰 세대로 현실적 체감 문제 해결을 최우선으로 인식.",
    "예) 정치적으로는 중도 실용주의에 가깝고 ‘말의 정치’보다 실제 성과·정책 실행 능력을 중시.",
  ]),
  ...blank(1),
  tbl(
    ["구분", "패널 특성"],
    G.map((g) => [g, "- [핵심 관심사·생활 압박]\n- [성향·판단 기준]\n- [조사 주제에 대한 태도, 선택 기준]"]),
    [2, 8]
  ),
  P("예) 50대 남성 | 경제·생활 압박이 매우 커 생활경제·일자리·주거 안정 중시 / 정치 성향: 합리성과 실용성을 중시하는 중도 이동형 / 후보 선택 기준: 확장성과 능력 그리고 소통력"),
  H3("2. 주요 아젠다 인식 특성"),
  summaryBox([
    "[조사 전체의 결론 2줄 : 무엇이 판단 기준으로 작동했고, 어떤 결과를 기대하는가]",
    "예) 대통령에 대한 높은 신뢰가 서울시장 후보 선택의 기준으로 전이되어 행정 경험 이력과 아울러 소통력의 후보를 선호.",
  ]),
  ...blank(1),
  P("■ [주제 1]", { bold: true, size: 21 }),
  B("[전 그룹 공통 인식 — 한 줄]   예) 전원 고평가: 문제 대응·위기관리·정책 추진력에서 “실용적이고 안정적” 평가 지배적."),
  B("[근거가 된 대표 사유]"),
  B("[단서·유보 의견]   예) 다만, 내란 척결 과정은 “더디고 답답하다”는 평가 존재."),
  P("■ [주제 2]", { bold: true, size: 21 }),
  B("[전망·평가의 방향]   예) 서울시장 선거는 박빙 전망 : 오세훈 출마 시 불리, 불출마 시 민주당 우위 가능성 언급."),
  B("[그룹 간 차이가 있으면 한 줄]"),
  P("- 비교 대상이 둘일 때는 좌우 비교표를 쓴다.", { color: C.midGray, size: 18 }),
  tbl(["[비교 대상 A]", "[비교 대상 B]"], [["[한 줄 총평]\n[근거 1~2개]", "[한 줄 총평]\n[근거 1~2개]"]], [5, 5], { firstColFill: false, firstColBold: false, firstColCenter: false }),
  P("예) 정청래 대표 | 신뢰감 있으나 추진력·조율능력·확장성에서 미흡  ↔  장동혁 대표 | 비전 없고 정체성이 자주 바뀜"),
  P("■ [주제 3 : 비교 대상 인식]", { bold: true, size: 21 }),
  P("1. 공통 기준", { bold: true }),
  B("[패널이 비교 대상을 평가할 때 쓴 잣대 3개]   예) 생활 밀착형 정책 이해도 / 전달력·소통력 / 중도 확장성"),
  P("2. 대상별 요약", { bold: true }),
  tbl(
    ["대상", "요약"],
    [
      ["[대상 1]", "- 이미지: [ ]\n- 강점: [ ]\n- 약점: [ ]\n- 조언: [ ]"],
      ["[대상 2]", "- 이미지: [ ]\n- 강점: [ ]\n- 약점: [ ]\n- 조언: [ ]"],
      ["[대상 3]", "- 이미지: [ ]\n- 강점: [ ]\n- 약점: [ ]\n- 조언: [ ]"],
    ],
    [2, 8]
  ),
  P("예) 박주민 | 이미지: 친근함, 솔직함, ‘발로 뛰는 사람’ / 강점: 진정성·세월호 상징성·실제 생활문제 해결에 관심 / 약점: 만만해 보임, 서울시장다운 ‘품격·무게감’ 부족 / 조언: 고급스러운 이미지·행정형 포지셔닝 강화 필요"),
  P("■ [주제 4 : 현안 인식]", { bold: true, size: 21 }),
  B("[가장 뚜렷한 대비 한 줄]   예) 오세훈 vs 박원순 시정 비교 매우 명확"),
  tbl(["[비교 대상 A]", "[비교 대상 B]"], [["[한 줄 총평]\n[근거]", "[한 줄 총평]\n[근거]"]], [5, 5], { firstColFill: false, firstColBold: false, firstColCenter: false }),
  P("■ [주제 5 : 바라는 점]", { bold: true, size: 21 }),
  B("[요구 사항을 굵게 한 줄씩, 5~7개]   예) 주거·안전·교육·복지 등 40대가 겪는 실질적 어려움에 대한 즉각적 해결 능력"),
  B("[ ]"),
  B("[ ]"),
  pageBreak()
);

// III. 주요 논의 내용
push(
  H1("III. 주요 논의 내용 (Main Discussion Points)"),
  note([
    "질문마다 같은 순서로 쓴다 : ① 질문 제목(Q번호 + 질문 요지) → ② 요약 박스(→ 2~4줄, 결론 먼저) → ③ 그룹별 세부내용 표 → ④ 필요 시 비교표·거수 결과·한 단어 표.",
    "종합보고서에서는 발언 인용을 최소화하고, 인용은 그룹별 보고서(Part B)에 둔다. 다만 결론을 바꾸는 결정적 발언은 종합에도 싣는다.",
  ]),
  ...blank(1),

  // 주제 1
  banner("1. [주제 1]        예) 대통령 국정 운영"),
  Q("Q1-1. [질문 요지]        예) Q2-1. 개인적 걱정거리 혹은 관심 사항"),
  summaryBox([
    "[전 그룹 공통 결론]   예) 대체로 불경기와 일자리 우려가 매우 강하며 단기 정책보다 생태계 구축 선호",
    "[그룹 간 공통점 또는 차이]   예) 4개 그룹 모두 중도층 유권자에 대한 관심 높게 나타남",
  ]),
  ...blank(1),
  tbl(["구분", "세부내용"], groupRows("- [그룹의 핵심 발언 요지 1~3개]"), [2, 8]),
  P("예) 40대 남성 | 먹고 사는 문제에 대한 큰 고민, 생활비 압박, 대출 및 금리 부담 심화 / AI 영향으로 인한 직업 파괴와 빠른 은퇴 압박 / 서울 유권자 지형 변화(중도/보수화)와 청년층의 경기도 이주 현상 언급"),
  Q("Q1-2. [질문 요지]        예) Q2-2. 이재명 대통령의 국정운영 평가"),
  summaryBox(["[결론]", "[유보·불만 지점]"]),
  ...blank(1),
  tbl(["구분", "세부내용"], groupRows("- [ ]"), [2, 8]),
  ...blank(1),

  // 주제 2
  banner("2. [주제 2]        예) 내년 지방선거 전망"),
  Q("Q2-1. [평가 대상이 둘 이상인 질문]        예) Q3-1. 여야 지도부 평가"),
  summaryBox(["[대상 A에 대한 결론]", "[대상 B에 대한 결론]"]),
  ...blank(1),
  B("[대상 A]   예) 정청래 더불어민주당 대표"),
  tbl(["구분", "세부내용"], groupRows("- [ ]"), [2, 8]),
  B("[대상 B]   예) 장동혁 국민의힘 대표"),
  tbl(["구분", "세부내용"], groupRows("- [ ]"), [2, 8]),
  Q("Q2-2. [전망을 둘로 나눠 묻는 질문]        예) Q3-3. 서울시장 선거 및 전국 지방선거 판세 예측"),
  summaryBox(["[전망 A]", "[전망 B와 그 근거]"]),
  ...blank(1),
  tbl(["구분", "[전망 A]", "[전망 B]"], G.map((g) => [g, "- [ ]", "[ ]"]), [2, 5, 3]),
  P("예) 40대 여성 | 오세훈 시장 피로감에도 민주당 승리 불투명, 김민석 총리를 유력 대항마로 소구 | 대통령 국정 운영 안정화에 힘입어 민주당 압승 예상"),
  Q("Q2-3. [주관식 인지·언급 질문]        예) Q3-4. 서울시장 출마 후보 인지도"),
  summaryBox(["[가장 많이 언급된 항목 순서]   예) 민주당 후보군 중에서는 김민석 총리, 정원오 구청장, 전현희 의원, 박주민 의원 주로 언급"]),
  ...blank(1),
  tbl(["구분", "언급 내용(언급 순)"], groupRows("[이름·항목을 언급 순으로 나열, 소속·유형이 다르면 줄 바꿈]"), [2, 8]),
  ...blank(1),

  // 주제 3 비교 대상
  banner("3. [주제 3 : 비교 대상 인식]        예) 후보 이미지 (박주민·전현희·서영교)"),
  Q("Q3-1. [이미지·키워드 / 강점·약점 / 외적 이미지]"),
  note([
    "관찰 메모(패널이 질문을 어떻게 받아들였는지)를 요약 박스 앞에 * 표시로 1~3줄 적는다.",
    "예) * 남성 패널은 이미지와 외적 이미지, 키워드, 화법 등을 총칭해 통합적 메시지로 인지했음 / * 여성 패널은 외모, 화장법, 의상 등을 분리해 인지했음",
  ]),
  ...blank(1),
  P("[대상 1]", { bold: true }),
  summaryBox([
    "[대상 1] → 이미지 : [키워드 3~4개]",
    "→ 강점 1. [ ]  2. [ ]  3. [ ]",
    "→ 약점 1. [ ]  2. [ ]  3. [ ]",
    "예) 전현희 의원 → 이미지 : 힐러리 같은 똑똑함, 엘리트, 정의로움 → 강점 1. 압도적 스펙 2. 강남 승리 경험 3. 권익위 시절 저항과 버팀 → 약점 1. 서울시장 후보로 이력 부족해 보임 2. 혼자 문제 해결할 듯 보임 3. 전달력 낮고 불분명한 발성",
  ]),
  ...blank(1),
  P("- 그룹별 키워드", { size: 19 }),
  tbl(["구분", "키워드"], groupRows("[키워드 3~5개, 쉼표 구분]"), [2, 8]),
  P("- 그룹별 이미지·강점·약점", { size: 19 }),
  tbl(["구분", "이미지", "강점", "약점"], G.map((g) => [g, "- [ ]", "- [ ]", "- [ ]"]), [1.6, 2.6, 2.9, 2.9]),
  P("예) 50대 여성 | 강남에 강한, 모범생 이미지 / 권익위 | 이과·문과 융합형 스펙 / 진정성 있어 의외로 강한 전투력 | 두꺼운 아이라인, 과한 화장으로 고지식해 보임 / 발성이 불명확하고 문장이 길어 전달력이 떨어짐"),
  P("[대상 2], [대상 3] … 같은 구성을 반복", { color: C.midGray }),
  Q("Q3-2. [좋아하는 이유 / 싫어하는 이유 / 참모로서 조언]"),
  P("[대상 1]", { bold: true }),
  summaryBox(["[대상 1]", "- 선호 이유 : [ ]", "- 비선호 이유 : [ ]", "- 참모로서 조언 : [ ]"]),
  ...blank(1),
  tbl(["구분", "좋은 점", "싫은 점", "조언"], G.map((g) => [g, "- [ ]", "- [ ]", "- [ ]"]), [1.6, 2.8, 2.8, 2.8]),
  P("[대상 2], [대상 3] … 반복", { color: C.midGray }),
  Q("Q3-3. [기억에 남는 정책·행동 / 특정 시점 이후 활약상]"),
  P("[대상 1]", { bold: true }),
  summaryBox(["[대상 1]", "- 기억에 남는 정책·행동 : [ ]", "- 최근 활동 : [ ]"]),
  ...blank(1),
  tbl(["구분", "기억에 남는 정책이나 행동", "[기준 시점] 이후 행동"], G.map((g) => [g, "- [ ]", "- [ ]"]), [2, 4, 4]),
  P("[대상 2], [대상 3] … 반복", { color: C.midGray }),
  Q("Q3-4. [경쟁력 / 역량 / 기여도 — 거수·지목 질문]"),
  note(["거수·지목 결과는 “N표 중 M표”로 적고, 이유가 있으면 같은 칸에 줄을 바꿔 적는다. 기권·복수응답은 괄호로 표기한다. 예) 8표 중 6표(기권 1표)"]),
  ...blank(1),
  P("[대상 1]", { bold: true }),
  tbl(["구분", "[지표 A : 후보 경쟁력]", "[지표 B : 업무 역량]", "[지표 C : 기여도]"], G.map((g) => [g, "- [n]표 중 [m]표\n- [이유]", "- [n]표 중 [m]표", "- [n]표 중 [m]표"]), [1.6, 2.8, 2.8, 2.8]),
  P("예) 50대 남성 | 6표 중 2표 | 6표 중 3표 — 입법 활동을 활발히 한 점, 공무원과 토론이 가능한 점으로 제2의 박원순 시장으로 서민 정책 잘 할 것이라는 의견 | 언론 이간질 우려, 그러나 대체로 누가 되든 정부와 호흡 맞출 것으로 예상"),
  P("[대상 2], [대상 3] … 반복", { color: C.midGray }),
  Q("Q3-5. [한 단어 정의]"),
  summaryBox(["[그룹 간 공통 키워드와 결이 다른 키워드 한 줄]"]),
  ...blank(1),
  B("[그룹 1]"),
  tbl(["대상", "한 단어 정의"], [["[대상 1]", "[원문 그대로, 쉼표 구분]"], ["[대상 2]", "[ ]"], ["[대상 3]", "[ ]"]], [2, 8]),
  P("예) 박주민 | 세월호, 서민, 스마트, 대통령감      전현희 | 똑똑한 엘리트, 행정부 경력을 쌓았으면 좋겠다"),
  P("[그룹 2] ~ [그룹 4] 같은 표를 반복", { color: C.midGray }),
  ...blank(1),

  // 주제 4 현안
  banner("4. [주제 4 : 현안 평가]        예) 서울시 현안"),
  Q("Q4-1. [불편한 점 · 개선해야 할 점]"),
  summaryBox(["[핵심 결론 — 굵게 한 줄]", "- [세부 근거 1]", "- [세부 근거 2]", "- [세부 근거 3]"]),
  ...blank(1),
  tbl(["구분", "세부내용"], groupRows("- [ ]"), [2, 8]),
  Q("Q4-2. [두 대상 비교 평가]        예) Q5-2,3. 오세훈 시장과 박원순 시장 비교"),
  summaryBox(["[대상 A 총평]   예) 박원순 시장은 ‘시민의 삶’과 ‘디테일’에 초점을 맞춘 시장으로 긍정적 평가", "[대상 B 총평]   예) 오세훈 시장은 ‘도시 외형’과 ‘실적주의’에 치중한다는 시장으로 부정적 평가"]),
  ...blank(1),
  tbl(["구분", "세부내용"], groupRows("- ([대상 A]) [평가와 사례]\n- ([대상 B]) [평가와 사례]"), [2, 8]),
  Q("Q4-3. [특정 정책에 대한 주변 여론과 본인 인식]"),
  summaryBox(["[주변 여론의 방향]", "[본인 인식과의 간극]"]),
  ...blank(1),
  tbl(["구분", "세부내용"], groupRows("- [주변 여론]\n- [본인 인식]"), [2, 8]),
  ...blank(1),

  // 주제 5 바라는 점
  banner("5. [주제 5 : 바라는 점]        예) 차기 서울시장에게 바라는 점"),
  Q("Q5-1. [거주지·현장에서 해결해야 할 과제]"),
  note(["요구 사항이 많을 때는 요약 박스 안에서 ‘→ 굵은 소제목 + - 세부 항목’으로 묶는다. 예) → 도시 안전 및 인프라 신뢰 회복 필요 / - 싱크홀 우려 해소 및 노후 지하 시설물 전면 점검 …"]),
  ...blank(1),
  summaryBox(["[범주 1]", "- [세부 요구]", "- [세부 요구]", "[범주 2]", "- [세부 요구]", "[범주 3]", "- [세부 요구]"]),
  ...blank(1),
  tbl(["구분", "세부내용"], groupRows("- [ ]"), [2, 8]),
  Q("Q5-2. [바라는 덕목과 정책]"),
  summaryBox(["[덕목 — 한 줄]   예) 행정을 깊이 이해하고 직전 사업을 지속할 수 있는 유능하고 검증된 전문가", "[시정·정책 방향 — 한 줄]", "[기대하는 리더십 — 한 줄]"]),
  ...blank(1),
  tbl(["구분", "세부내용"], groupRows("[덕목 / 정책 / 전략 순으로 줄 바꿈]"), [2, 8]),
  ...blank(1),

  banner("6. 마무리"),
  P("오늘 수고 많으셨다. 마지막으로 하실 말씀 부탁드린다.", { bold: true }),
  summaryBox(["[마무리 발언에서 나온 요청 사항 한 줄]   예) 당원 간담회 및 교육 개최 활성화"]),
  pageBreak()
);

// IV. 시사점 (선택)
push(
  H1("IV. 시사점 및 제언 (선택)"),
  note(["발주처가 ‘요약, 비교, 기대 정책’ 이상의 제언을 요구할 때 넣는다. 서울시장 FGI에서는 II장 요약이 이 역할을 대신했으므로 생략했다. 넣을 때는 근거(그룹·질문 번호)를 반드시 붙인다."]),
  ...blank(1),
  H3("1. 핵심 발견"),
  tbl(["순", "발견", "근거(그룹 · Q번호)", "시사점"], [["1", "[ ]", "[ ]", "[ ]"], ["2", "[ ]", "[ ]", "[ ]"], ["3", "[ ]", "[ ]", "[ ]"], ["4", "[ ]", "[ ]", "[ ]"], ["5", "[ ]", "[ ]", "[ ]"]], [0.8, 4, 2.2, 3]),
  H3("2. 대상별 제언"),
  tbl(["대상", "유지할 것", "보완할 것", "우선 과제"], [["[대상 1]", "[ ]", "[ ]", "[ ]"], ["[대상 2]", "[ ]", "[ ]", "[ ]"], ["[대상 3]", "[ ]", "[ ]", "[ ]"]], [1.6, 2.8, 2.8, 2.8]),
  H3("3. 정책·메시지 제언"),
  tbl(["영역", "제언", "근거", "시기"], [["[ ]", "[ ]", "[ ]", "단기"], ["[ ]", "[ ]", "[ ]", "단기"], ["[ ]", "[ ]", "[ ]", "중기"]], [1.6, 4.4, 2.6, 1.4]),
  pageBreak()
);

// 부록
push(
  H1("부록"),
  H3("부록 1. 참가자 명단"),
  note(["종합보고서 본문에는 실명을 싣지 않는다. 명단은 발주처 요청이 있을 때만 부록으로 붙이고, 없으면 그룹·번호·성별·연령대·거주 권역만 남긴다."]),
  ...blank(1),
  tbl(["그룹", "번호", "성별·연령", "거주지(구)", "직업·특성", "비고(초반/후반, 불참)"], [
    ["1", "P1", "[ ]", "[ ]", "[ ]", "[ ]"], ["1", "P2", "[ ]", "[ ]", "[ ]", "[ ]"], ["1", "…", "", "", "", ""],
    ["2", "P1", "[ ]", "[ ]", "[ ]", "[ ]"], ["2", "…", "", "", "", ""],
    ["3", "P1", "[ ]", "[ ]", "[ ]", "[ ]"], ["3", "…", "", "", "", ""],
    ["4", "P1", "[ ]", "[ ]", "[ ]", "[ ]"], ["4", "…", "", "", "", ""],
  ], [1, 1, 1.6, 1.8, 2.4, 2.2]),
  P("예) 강봉비(초반, 은평), 장훈(후반, 마포) … (사전 확정 8명 중 2명 불참)"),
  H3("부록 2. 질문지(진행 스크립트)"),
  P("[질문지 원문 또는 파일명·경로]"),
  H3("부록 3. 녹취 로데이터 목록"),
  tbl(["그룹", "파일명", "형식", "비고"], G.map((g) => [g, "[YYYY_그룹_MMDD_FGI_로데이터]", "[xlsx / hwp]", "[후처리 여부]"]), [2, 4.5, 1.5, 2]),
  pageBreak()
);

// ========== PART B. 그룹별 보고서
push(
  titleBand("Part B. 그룹(회차)별 FGI 보고서 서식"),
  ...blank(1),
  note([
    "회차 하나를 정리한다. 종합보고서(Part A)와 장 구성을 같게 하되, ‘그룹별 표’ 대신 ‘패널 주요 발언내용’ 인용 박스를 질문마다 붙인다.",
    "이 문서의 요약 박스와 비교표가 그대로 종합보고서의 한 행이 된다. 그러므로 질문 번호·표 열 구성을 종합보고서와 반드시 맞춘다.",
  ]),
  ...blank(1),
  H2("B-0. 표지"),
  titleBand("[조사 주제] 관련 FGI 보고서"),
  ...blank(1),
  P("[발주처·대상] 심층 좌담회 (FGI)", { align: AlignmentType.CENTER, size: 28, bold: true }),
  P("[그룹 1], [그룹 2], [그룹 3], [그룹 4]   ← 해당 회차 그룹만 굵게 표시", { align: AlignmentType.CENTER, size: 20 }),
  P("예) 50대 남성, 50대 여성, 40대 남성, 40대 여성", { align: AlignmentType.CENTER }),
  ...blank(1),
  P("개요", { bold: true, size: 22 }),
  summaryBox([
    "- 일시 : [YYYY년 MM월 DD일(요일) 오전/오후 HH:MM ~ HH:MM]",
    "- 장소 : [인터뷰룸] + [모니터룸]",
    "- 패널 : [그룹명] [n]인",
    "   [성명(초반/후반, 거주지)] … (사전 확정 [n]명 중 [n]명 불참)",
    "- 퍼실리테이터 : [성명·직위·자격]",
    "- 모니터룸 : [성명·직위]",
    "예) 일시 : 2025년 11월 12일(수) 오후 2:00 ~ 5:00 / 장소 : 패스트파이브 시청1호점 11C 회의실 + 12A 모니터룸 / 패널 : 50대 남성 6인 / 퍼실리테이터 : 이병덕 코리아스픽스 대표이사, fKF 마스터 퍼실리테이터 / 모니터룸 : 이소연 팀장, 이소영 대리, 공정호 fKF 1급 퍼실리테이터",
  ], { fill: C.gray }),
  ...blank(1),
  P("목차 : Part A-0 목차와 동일 (I. FGI 개요 / II. 요약 / III. 주요 논의 내용 / IV. 이 회차의 시사점)", { color: C.midGray }),
  pageBreak(),

  H1("I. FGI 개요"),
  P("Part A의 I장(아젠다·주요 질문)을 그대로 쓴다. 회차 보고서에서는 ‘조사 개요·조사 설계·수행 체계’ 표를 생략하고 아젠다와 주요 질문만 남겨도 된다.", { color: C.midGray }),
  H3("아젠다"),
  B("[한 문장]"),
  H3("주요 질문"),
  P("1. [파트 ①] …  2. [파트 ②] …  3. [파트 ③] …  4. [파트 ④] …  5. [파트 ⑤] …  6. 마무리"),
  ...blank(1),

  H1("II. 요약 (Summary)"),
  H3("1. 패널 특성"),
  note(["회차 보고서의 패널 특성은 ‘굵은 소제목 + 설명 1~2줄’을 4~5개 쌓는 방식으로 쓴다. 종합보고서의 그룹별 특성 표 한 칸이 여기서 나온다."]),
  ...blank(1),
  P("[특성 1 : 생활·경제 상황]", { bold: true }),
  P("[설명 1~2줄]   예) 경제·생활 압박이 매우 큰 연령대 — 교육비·생활비·주거비·노후 대비가 동시에 부담되는 세대."),
  P("[특성 2 : 핵심 관심사]", { bold: true }),
  P("[설명]   예) 관심사: 생활경제·일자리·주거 안정 — 단기 정책보다 지속 가능한 생태계 중시."),
  P("[특성 3 : 조사 주제에 대한 기본 태도]", { bold: true }),
  P("[설명]   예) 권리당원이지만 서울시장 선거에 대한 우려 높음."),
  P("[특성 4 : 성향·판단 기준]", { bold: true }),
  P("[설명]   예) 합리성과 실용성을 중시하는 중도 이동형 — “일자리·경제·자산·시정 체감 성과”를 최우선 판단 기준으로 삼음."),
  P("[특성 5 : 선택 기준]", { bold: true }),
  P("[설명]   예) 후보 선택 기준: 확장성과 능력 그리고 소통력."),
  H3("2. 주요 아젠다 인식 특성"),
  P("■ [주제 1]", { bold: true, size: 21 }),
  B("[결론 — 굵게]   예) 전원 고평가: 5개월여 짧은 집권 기간에도 “내공 있는 실용주의 정부”로 인식."),
  B("[근거·특징]"),
  P("■ [주제 2]", { bold: true, size: 21 }),
  B("[전체 기류 한 줄]   예) 전체적 기류: “서울 선거는 쉽지 않다” → 후보 경쟁력·스타일·확장성 매우 중요."),
  B("[거수 결과와 토론 중 변화]   예) 1차 예상은 민주당 우세(4:2)였으나 토론 과정에서 “서울의 보수화” 우려로 의견 분화"),
  tbl(["[비교 대상 A]", "[비교 대상 B]"], [["[한 줄 총평]", "[한 줄 총평]"]], [5, 5], { firstColFill: false, firstColBold: false, firstColCenter: false }),
  P("■ [주제 3 : 비교 대상 인식]", { bold: true, size: 21 }),
  P("1. 공통 기준", { bold: true }),
  B("[이 그룹이 쓴 판단 기준]   예) 후보 경쟁력 판단기준 ① 확장성 ② 전달력 ③ 생활정책 역량 ④ 강남·중도 공략력"),
  P("2. 대상별 요약", { bold: true }),
  tbl(["대상", "요약"], [["[대상 1]", "- 이미지: [ ]\n- 강점: [ ]\n- 약점: [ ]\n- 조언: [ ]"], ["[대상 2]", "- 이미지: [ ]\n- 강점: [ ]\n- 약점: [ ]\n- 조언: [ ]"], ["[대상 3]", "- 이미지: [ ]\n- 강점: [ ]\n- 약점: [ ]\n- 조언: [ ]"]], [2, 8]),
  P("■ [주제 4 : 현안 인식]", { bold: true, size: 21 }),
  B("[가장 뚜렷한 인식 한 줄]"),
  B("[대표 사례]   예) 사례: 난폭해진 버스, 공사 중심 정책, 생활불편 증가"),
  tbl(["[비교 대상 A]", "[비교 대상 B]"], [["[ ]", "[ ]"]], [5, 5], { firstColFill: false, firstColBold: false, firstColCenter: false }),
  P("■ [주제 5 : 바라는 점]", { bold: true, size: 21 }),
  B("[ ]"),
  B("[ ]"),
  B("[ ]"),
  pageBreak(),

  H1("III. 주요 논의 내용 (Main Discussion Points)"),
  note([
    "질문마다 : ① 질문(질문지 원문 그대로) → ② 요약 박스(→ 2~5줄) → ③ 패널 주요 발언내용(인용 박스) → ④ 비교 질문이면 대상별 표.",
    "인용은 녹취 원문을 유지한다. 실명·지역 등 식별 정보와 비속어만 정리하고, 문장을 매끄럽게 고치지 않는다. 지지 대상이 갈리는 질문은 ‘[대상] 지지 의견’ 소제목으로 인용을 묶는다.",
  ]),
  ...blank(1),
  banner("1. [주제 1]"),
  Q("Q1-1. [질문지 원문]        예) Q2-1. 요즘 개인적으로 가장 큰 걱정거리 혹은 관심 사항이 있다면 무엇인지요?"),
  summaryBox(["[결론]", "[근거·특이점]", "[정책적 함의 또는 요청]"]),
  ...blank(1),
  B("패널 주요 발언내용", { bold: true }),
  quoteBox(["[발언 원문 1]", "[발언 원문 2]", "[발언 원문 3]", "예) “일자리 문제가 심각하다. 단기 일자리 정책보다 생태계를 구축하는 게 필요하다.”"]),
  Q("Q1-2. [질문지 원문]"),
  summaryBox(["[결론]", "[근거]"]),
  ...blank(1),
  B("패널 주요 발언내용", { bold: true }),
  quoteBox(["[ ]", "[ ]"]),
  ...blank(1),

  banner("2. [주제 2]"),
  Q("Q2-1. [평가 대상이 둘 이상인 질문]"),
  B("[대상 A]", { bold: true }),
  summaryBox(["[결론]", "[유보 의견]", "[소수의견은 ‘(소수의견)’ 표기]"]),
  ...blank(1),
  B("패널 주요 발언내용", { bold: true }),
  quoteBox(["[ ]", "예) “예 저는 10점 중에 9점이구요. 일단 법사위원장으로서 윤석열 탄핵을 마무리하셨고 …”"]),
  B("[대상 B]", { bold: true }),
  summaryBox(["[결론]"]),
  ...blank(1),
  B("패널 주요 발언내용", { bold: true }),
  quoteBox(["[ ]"]),
  Q("Q2-2. [거수로 시작하는 전망 질문]"),
  summaryBox([
    "[초반 거수 결과와 근거]   예) 초반엔 4:2 정도로 민주당 승리 예상, 그 근거로 …",
    "[재질문 후 변화]   예) 유권자 지형이 불리하다는 한 패널의 주장에 근거해 재질문했을 때 … 우려함",
    "[소수 의견]",
  ]),
  ...blank(1),
  B("패널 주요 발언내용", { bold: true }),
  quoteBox(["[ ]", "[ ]"]),
  Q("Q2-3. [주관식 인지 질문]"),
  summaryBox(["[언급된 항목을 소속·유형별로 나열]   예) 민주당 : 박주민 의원, 전현희 의원, 정원오 성동구청장 … / 국민의힘 : 오세훈 시장, 나경원 의원 / * 서영교 의원 지목자 없음"]),
  ...blank(1),
  Q("Q2-4. [경선·선택 예상 질문]"),
  summaryBox(["[지목 결과 : 대상별 표 수(이유)]   예) 박주민 3(높은 인지도, 합리적, 적극성), 정원오 1(업무역량), 전현희 1(강남에 통하는 후보), 서영교 1(국감 이미지)", "[관심도·분위기]"]),
  ...blank(1),
  B("패널 주요 발언내용", { bold: true }),
  quoteBox(["※ [대상 1] 지지 의견", "[ ]", "※ [대상 2] 지지 의견", "[ ]"]),
  ...blank(1),

  banner("3. [주제 3 : 비교 대상 인식]"),
  Q("Q3-1. [이미지·키워드 / 강점·약점 / 외적 이미지]"),
  summaryBox([
    "[이 그룹이 질문을 받아들인 방식]   예) 50대 남성층은 외모와 이미지 강점과 약점을 분리해 인식하지 않는 경향을 보임",
    "[공통 판단 기준]",
    "[대상 1] : [한 줄 총평]",
    "[대상 2] : [한 줄 총평]",
    "[대상 3] : [한 줄 총평]",
  ]),
  ...blank(1),
  B("후보 이미지와 강점 약점 정리", { bold: true }),
  tbl(["대상", "이미지", "강점", "약점"], [["[대상 1]", "- [ ]", "- [ ]", "- [ ]"], ["[대상 2]", "- [ ]", "- [ ]", "- [ ]"], ["[대상 3]", "- [ ]", "- [ ]", "- [ ]"]], [1.6, 2.6, 2.9, 2.9]),
  Q("Q3-2. [좋은 점·싫은 점 / 조언 / 기억에 남는 점 / 최근 활동]"),
  B("좋은 점 / 싫은 점 / 조언", { bold: true }),
  tbl(["대상", "좋은 점", "싫은 점", "참모로서 조언"], [["[대상 1]", "- [ ]", "- [ ]", "- [ ]"], ["[대상 2]", "- [ ]", "- [ ]", "- [ ]"], ["[대상 3]", "- [ ]", "- [ ]", "- [ ]"]], [1.6, 2.8, 2.8, 2.8]),
  B("기억에 남는 정책과 [기준 시점] 이후 기억에 남는 행동", { bold: true }),
  tbl(["대상", "기억에 남는 정책이나 행동", "[기준 시점] 이후 행동"], [["[대상 1]", "- [ ]", "- [ ]"], ["[대상 2]", "- [ ]", "- [ ]"], ["[대상 3]", "- [ ]", "- [ ]"]], [2, 4, 4]),
  Q("Q3-3. [경쟁력 / 역량 / 기여도 — 거수]"),
  summaryBox([
    "[지표 A] : [대상 순위와 표 수]   예) 후보 경쟁력 : 서영교 의원 3 > 박주민 의원 2 > 전현희 의원 1",
    "[대상별 이유 한 줄씩]",
    "[지표 B] : [ ]",
    "[지표 C] : [ ]",
  ]),
  ...blank(1),
  Q("Q3-4. [떠오르는 단어]"),
  B("대상별 한 단어 정의", { bold: true }),
  tbl(["대상", "한 단어 정의"], [["[대상 1]", "[원문 그대로, 쉼표 구분]"], ["[대상 2]", "[ ]"], ["[대상 3]", "[ ]"]], [2, 8]),
  ...blank(1),

  banner("4. [주제 4 : 현안 평가]"),
  Q("Q4-1. [불편한 점 · 개선점]"),
  summaryBox(["[결론 — 굵게]", "- [사례 1]", "- [사례 2]", "[긍정 평가가 있으면 별도 줄]"]),
  ...blank(1),
  Q("Q4-2. [두 대상 비교]"),
  summaryBox(["[비교의 핵심]", "[유보 의견]"]),
  ...blank(1),
  B("[대상 A] vs [대상 B] 비교", { bold: true }),
  tbl(["[대상 A]", "[대상 B]"], [["[굵은 총평]\n[사례 나열]", "[굵은 총평]\n[사례 나열]"]], [5, 5], { firstColFill: false, firstColBold: false, firstColCenter: false }),
  P("예) 박원순 | 서울 시민의 삶이 목적인 도시 — 치수, 보행친화도시, 따릉이 등  ↔  오세훈 | 시민의 삶이 도시의 도구로 전락 — 도시외형 치중, 실적주의, 난폭해진 버스"),
  Q("Q4-3. [특정 정책에 대한 주변 인식과 판단]"),
  summaryBox(["[주변 여론]", "[본인 인식]", "[선거·사업에 미칠 영향]"]),
  ...blank(1),
  B("패널 주요 발언내용", { bold: true }),
  quoteBox(["[ ]", "[ ]"]),
  ...blank(1),

  banner("5. [주제 5 : 바라는 점]"),
  Q("Q5-1. [거주지에서 꼭 해결해야 할 점]"),
  summaryBox(["[요구 1]", "[요구 2]", "[제안]"]),
  ...blank(1),
  B("패널 주요 발언내용", { bold: true }),
  quoteBox(["[ ]", "[ ]", "[ ]"]),
  Q("Q5-2. [바라는 덕목과 정책]"),
  summaryBox(["[덕목]", "[정책]", "[전략·기대]"]),
  ...blank(1),

  banner("6. 마무리"),
  summaryBox(["[마무리 발언 요지]   예) 당원 간담회 자주 개최 및 더 많은 당원 교육 개최 희망"]),
  pageBreak(),

  H1("IV. 이 회차의 시사점 (선택)"),
  note(["종합보고서를 쓰는 사람에게 넘기는 메모다. 발주처 제출본에서는 뺀다."]),
  ...blank(1),
  tbl(["구분", "내용"], [
    ["종합보고서에 올릴 것", "[이 그룹에서만 나온 뚜렷한 발견]"],
    ["다른 그룹과 대조할 것", "[다음 회차에서 재확인할 질문·가설]"],
    ["판단을 보류한 것", "[근거가 한 명뿐이거나 상충하는 발언]"],
    ["진행상 특이사항", "[지각·불참, 시간 배분, 질문 순서 변경 등]"],
  ], [2.6, 7.4]),
  pageBreak()
);

// ========== PART C. 작성 가이드
push(
  titleBand("Part C. 작성 가이드 · 체크리스트"),
  ...blank(1),
  H2("C-1. 블록별 작성 규칙"),
  tbl(
    ["블록", "규칙", "서울시장 FGI 예시"],
    [
      ["요약 박스", "- 각 줄을 → 로 시작하고 결론을 먼저 쓴다\n- 2~4줄. 길어지면 ‘→ 굵은 소제목 + - 세부’로 묶는다\n- 한 명의 의견은 “(소수의견)” 표기", "→ 대체로 불경기와 일자리 우려가 매우 강하며 단기 정책보다 생태계 구축 선호\n→ 4개 그룹 모두 중도층 유권자에 대한 관심 높게 나타남"],
      ["그룹별 세부내용 표", "- 그룹 순서를 문서 전체에서 고정\n- 셀 안은 ‘- ’ 항목 1~3개, 명사형 종결\n- 언급이 없으면 “언급 없음” 또는 “특정 행동 기억 없음”", "50대 남성 | - 불경기 일자리 감소, 부동산 안정책에 따른 민심 이반 우려 나타남"],
      ["비교표(이미지·강점·약점·조언)", "- 4축을 유지한다. 축이 비어도 열을 지우지 않는다\n- 대상 순서는 질문지 순서(조사 순서)로 고정", "구분 | 이미지 | 강점 | 약점  /  구분 | 좋은 점 | 싫은 점 | 조언"],
      ["거수·지목 결과", "- “N표 중 M표” 형식. N은 그 시점 참석 인원\n- 기권·복수응답은 괄호  예) 8표 중 6표(기권 1표)\n- 이유가 있으면 같은 칸에서 줄 바꿈", "40대 남성 | 7표 중 7표 | 7표 중 5표 | 대통령과 삐그덕거릴 수도 있겠다는 의견"],
      ["패널 주요 발언", "- 큰따옴표, 녹취 원문 유지(문법 교정 금지)\n- 실명·직장·동네 등 식별 정보는 지우거나 일반화\n- 지지 대상이 갈리면 ‘※ [대상] 지지 의견’으로 묶음", "“서울시장 선거는 대선과 다르다.”"],
      ["한 단어 정의", "- 패널이 말한 단어를 그대로, 쉼표로 나열\n- 문장형 답변도 자르지 않고 싣는다", "전현희 | 출발선에 서있지만 심판인지 러너인지 헷갈리는 사람"],
      ["관찰 메모", "- 패널이 질문을 받아들인 방식은 * 로 시작해 요약 박스 앞에 둔다", "* 남성 패널은 이미지·키워드·화법을 통합적으로 인지 / * 여성 패널은 외모·화장·의상을 분리해 인지"],
    ],
    [2, 4.4, 3.6]
  ),
  H2("C-2. 문체"),
  B("명사형·서술형 종결로 통일한다. 예) “~로 인식”, “~ 우려 표명”, “~ 선호”, “~ 지적”. ‘~습니다’체를 쓰지 않는다."),
  B("인물 호칭은 문서 전체에서 하나로 고정한다. 예) “박주민 의원”, “오세훈 시장”, “정청래 대표”. 표 안에서는 성명만 써도 된다."),
  B("평가 표현은 패널의 말에서 나온 것만 쓴다. 작성자의 판단은 IV장(시사점)에만 쓴다."),
  B("숫자는 반각 아라비아 숫자, 날짜는 “11/12 PM 2:00 ~ 5:00”(표지) 또는 “2025년 11월 12일(수) 오후 2:00 ~ 5:00”(개요)로 쓴다."),
  H2("C-3. 파일명 규칙"),
  tbl(
    ["문서", "파일명", "예시"],
    [
      ["그룹별 보고서", "YYYY_[그룹약칭]_MMDD_FGI_보고서", "2025_50남_1112_FGI_보고서 / 2025_40대여_1115_FGI_보고서"],
      ["종합보고서", "YYYY_종합_[주제약칭]_FGI_보고서", "2025_종합_서울시장_M_FGI_보고서"],
      ["녹취 로데이터", "[발주처] FGI 로데이터 (시트 = 그룹)", "민주당 서울시당 FGI 로데이터"],
      ["PPT 요약본", "[발주처] FGI 보고서", "민주당 FGI 보고서"],
      ["계획서", "FGI 계획 (대외비)", "FGI 계획 (대외비) 초안"],
    ],
    [2.2, 3.8, 4]
  ),
  H2("C-4. 제출 전 체크리스트"),
  tbl(
    ["구분", "점검 항목", "확인"],
    [
      ["표지", "그룹별 인원 합계 = ‘총 N명’ / 일시·장소가 그룹 보고서와 일치", "□"],
      ["목차", "목차의 주제·소항목 번호가 III장 배너·Q번호와 1:1로 일치", "□"],
      ["요약", "II장 요약의 결론이 III장 본문에서 근거를 찾을 수 있음", "□"],
      ["그룹 순서", "모든 표에서 그룹 순서가 동일", "□"],
      ["거수", "표 수 합계 ≤ 참석 인원, 기권 표기, N 값이 그룹 인원과 일치", "□"],
      ["인용", "실명·직장·동네 등 식별 정보 제거, 원문 유지", "□"],
      ["개인정보", "종합보고서 본문에 참가자 실명 없음 / 명단은 발주처 요청 시에만 부록", "□"],
      ["플레이스홀더", "[ ] 회색 글씨와 예) 파란 글씨가 모두 삭제됨", "□"],
      ["포맷", "발주처 요청 포맷(한글·PPT) 변환본 준비, 변환 후 표 깨짐 확인", "□"],
      ["보고 경로", "그룹별 초안 → 종합 초안 → 발주처 담당자 순으로 메일 발송, 참조자 지정", "□"],
    ],
    [1.8, 7.2, 1]
  ),
  H2("C-5. 서울시장 FGI 산출물 참조"),
  P("서식의 기준이 된 실제 문서들. 구조를 확인할 때 참고한다(구글 드라이브).", { color: C.midGray }),
  tbl(
    ["문서", "성격"],
    [
      ["FGI 계획 (대외비) 초안", "설계·질문 파트·예산. I장 ‘조사 설계’와 ‘주요 질문’의 원천"],
      ["2025_50남_1112_FGI_보고서 외 3부", "그룹별 보고서. Part B의 원형"],
      ["2025_종합_서울시장_M_FGI_보고서", "종합보고서. Part A의 원형"],
      ["민주당 서울시당 FGI 로데이터", "녹취 로데이터(엑셀). 부록 3의 원천"],
      ["민주당 FGI 보고서 (슬라이드)", "발주처 제출용 PPT 요약본"],
    ],
    [4, 6]
  )
);

// ───────────────────────────── 문서 조립
const doc = new Document({
  creator: "코리아스픽스 숙의공론컨설팅본부",
  title: "FGI 결과 보고서 작성 서식",
  styles: {
    default: { document: { run: { font: FONT, size: 20 }, paragraph: { spacing: { line: 300 } } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 32, bold: true, color: C.navy }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 26, bold: true, color: C.blue }, paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 22, bold: true }, paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bul", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "-", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 200 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "·", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 200 } } } },
      ] },
    ],
  },
  sections: [
    {
      properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
      headers: {
        default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "FGI 결과 보고서 작성 서식 v1.0 · 코리아스픽스", font: FONT, size: 16, color: C.midGray })] })] }),
      },
      footers: {
        default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: C.midGray })] })] }),
      },
      children: body,
    },
  ],
});

const out = process.argv[2] || "FGI_결과보고서_템플릿.docx";
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log("written:", out, buf.length, "bytes");
});
