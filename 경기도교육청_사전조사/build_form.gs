/**
 * 경기형 학교시민교육 100인 토의 — 참여단 사전조사 구글폼 생성 스크립트
 *
 * 근거: 실행계획서(안) v0.5.26(2026. 10. 16.)
 *   - Ⅲ-2. 사전조사 [표 2] 조사 문항(안): 개인정보 동의·배경 문항·추가 사안
 *   - Ⅴ-2. 사전조사 개정 문항(안): Q1~Q12 (최신안 우선 적용)
 *
 * 사용법
 *   1. https://script.google.com 에서 새 프로젝트를 만들고 이 파일 내용을 붙여넣는다.
 *   2. 함수 목록에서 buildForm 을 선택해 실행하고 권한을 승인한다.
 *   3. 실행 로그에 출력되는 편집 URL·응답 URL·응답 시트 URL·개인별 링크 형식을 확인한다.
 */

const FORM_TITLE = '경기형 학교시민교육 100인 토의 참여단 사전조사';

const FORM_DESCRIPTION = [
  '안녕하십니까.',
  '경기도교육청은 「경기형 보이텔스바흐 합의」 도출을 위해 2026년 10월 29일(목) 조원청사에서 「경기형 학교시민교육 100인 토의」를 개최합니다.',
  '',
  '이 조사는 토의에 앞서 참여단 여러분의 인식·경험·제안을 듣고, 논쟁적 사회 현안 수업의 원칙과 교사·학생·학교의 실천 기준에 관한 토론 쟁점을 정리하기 위해 실시합니다. 응답 결과는 토의자료와 당일 결과 브리핑에 반영됩니다.',
  '',
  '· 응답 소요 시간: 약 10분',
  '· 응답 기한: 2026년 10월 22일(목) 18시',
  '· 응답 내용은 통계 목적으로만 사용되며, 개인을 식별할 수 없는 형태로 집계됩니다.',
  '',
  '주최: 경기도교육청 민주시민교육과 / 수행: 코리아스픽스·지혜의영토(주)·(사)한국퍼실리테이터연합회',
].join('\n');

const CONSENT_TEXT = [
  '■ 수집·이용 목적: 100인 토의 사전조사 결과 분석 및 토의자료·결과보고서 작성',
  '■ 수집 항목: 응답자 확인번호, 성별, 연령대, 참여 유형, 설문 응답 내용',
  '■ 보유·이용 기간: 사업 종료(결과보고서 제출) 후 지체 없이 파기',
  '■ 동의 거부권: 개인정보 수집·이용에 동의하지 않을 권리가 있습니다. 다만 동의하지 않으시면 사전조사에 참여하실 수 없으며, 100인 토의 참여에는 영향이 없습니다.',
].join('\n');

// 국가교육위원회 「학교시민교육 기본원칙」 제3장 요약 (2026. 9. 확정, 보도참고자료 기준)
const PRINCIPLES_INTRO = [
  '국가교육위원회는 2026년 9월 「학교시민교육 기본원칙」을 확정했습니다. 그중 논쟁적 사회 현안 수업과 관련된 원칙(제3장)은 다음과 같습니다.',
  '',
  '· 명확성(제12조 ①): 헌법 가치 등 기본 내용은 명확하게 가르친다.',
  '· 논쟁성(제12조 ②): 견해가 나뉘는 사안은 정치적 중립을 지키며 다양한 관점을 제시하고, 논쟁 자체를 학습 기회로 삼는다(편향 주입 금지).',
  '· 시의성(제13조): 시의성 있는 공적 사안을 교육과정·목표·학생 발달 수준에 맞게 적극적으로 다룬다.',
  '· 안전성(제14조): 학생은 불이익 없이 의견을 표현하고 자치·사회활동에 참여할 수 있으며, 교원은 부당한 압력·간섭 없이 정당한 교육활동을 할 수 있도록 보호받는다.',
  '',
  '※ 출처: 국가교육위원회, 「학교시민교육 기본원칙」 보도참고자료(2026. 9. 17.)',
].join('\n');

const PRINCIPLE_ROWS = [
  '명확성',
  '논쟁성',
  '시의성',
  '학생의 안전한 표현·참여',
  '교원 보호',
];

function buildForm() {
  const form = FormApp.create(FORM_TITLE);
  form
    .setDescription(FORM_DESCRIPTION)
    .setCollectEmail(false)
    .setAllowResponseEdits(false)
    .setShowLinkToRespondAgain(false)
    .setProgressBar(true)
    .setConfirmationMessage(
      '응답해 주셔서 감사합니다.\n' +
      '응답 결과는 10월 29일(목) 100인 토의 자료와 결과 브리핑에 반영됩니다.'
    );

  // ── 1쪽: 응답자 확인 · 개인정보 동의 ─────────────────────────────
  const idItem = form.addTextItem()
    .setTitle('응답자 확인번호')
    .setHelpText('문자로 받으신 개인별 링크로 접속하시면 자동으로 입력됩니다. 수정하지 마십시오.')
    .setRequired(true);

  form.addSectionHeaderItem()
    .setTitle('개인정보 수집·이용 안내')
    .setHelpText(CONSENT_TEXT);

  const consent = form.addMultipleChoiceItem()
    .setTitle('위 내용을 안내받았으며, 개인정보 수집·이용에 동의하십니까?')
    .setRequired(true);
  // 선택지 이동 설정은 아래 페이지가 생긴 뒤에 지정한다.

  // ── 2쪽: 응답자 정보 ─────────────────────────────────────────────
  form.addPageBreakItem().setTitle('응답자 정보');

  form.addMultipleChoiceItem()
    .setTitle('귀하의 성별은 무엇입니까?')
    .setChoiceValues(['남성', '여성', '응답 원치 않음'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('귀하의 연령대는 어디에 해당합니까?')
    .setChoiceValues(['10대', '20대', '30대', '40대', '50대', '60대', '70대 이상'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('이번 토의에 참여하는 주된 직업·활동 유형을 선택해 주십시오.')
    .setChoiceValues(['교육공동체 관련 전문가', '교사', '학부모', '시민단체'])
    .showOtherOption(true)
    .setRequired(true);

  // ── 3쪽: 기본원칙 인지도·관심도 (원칙 설명 전에 질문) ──────────────
  form.addPageBreakItem()
    .setTitle('학교시민교육 기본원칙에 대한 인식')
    .setHelpText('다음 쪽에서 원칙을 간략히 소개해 드립니다. 이 쪽은 현재 알고 계신 대로 응답해 주십시오.');

  form.addMultipleChoiceItem()
    .setTitle('Q1. 국가교육위원회가 제시한 학교시민교육 기본원칙을 이전부터 알고 있었습니까?')
    .setChoiceValues(['내용까지 알고 있었다', '들어본 적 있다', '몰랐다'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('Q2. 해당 원칙과 학교 적용 논의에 얼마나 관심이 있습니까?')
    .setChoiceValues([
      '① 전혀 관심 없다',
      '② 별로 관심 없다',
      '③ 보통이다',
      '④ 관심 있는 편이다',
      '⑤ 매우 관심 있다',
      '잘 모르겠음',
    ])
    .setRequired(true);

  // ── 4쪽: 경기형 공통 원칙의 필요성 ─────────────────────────────────
  form.addPageBreakItem()
    .setTitle('경기형 공통 원칙의 필요성')
    .setHelpText(PRINCIPLES_INTRO);

  form.addMultipleChoiceItem()
    .setTitle('Q3. 경기도 학교에서 논쟁적 사회 현안을 다루는 수업·교육활동에서 교사·학생·학교가 함께 지킬 ‘경기형 공통 원칙’을 마련하는 것이 얼마나 필요하다고 생각합니까?')
    .setHelpText('‘경기형 공통 원칙’은 국가 원칙을 경기 학교 현장에 맞게 구체화한 원칙을 말합니다.')
    .setChoiceValues([
      '① 전혀 필요하지 않다',
      '② 필요하지 않은 편이다',
      '③ 보통이다',
      '④ 필요한 편이다',
      '⑤ 매우 필요하다',
      '잘 모르겠음',
    ])
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('Q3-1. 그렇게 응답하신 이유는 무엇입니까?')
    .setHelpText('필요하다·필요하지 않다 어느 쪽이든 이유를 적어 주십시오.')
    .setRequired(true);

  // ── 5쪽: 학생의 현재 문제와 우려 ───────────────────────────────────
  form.addPageBreakItem()
    .setTitle('학생의 현재 문제와 우려')
    .setHelpText('아래 선택지는 예시이며 확정된 진단이 아닙니다. 해당하는 것이 없으면 ‘기타’ 또는 ‘판단 어려움’을 선택해 주십시오.');

  form.addCheckboxItem()
    .setTitle('Q4. 학생들이 현재 겪는 문제 중 학교시민교육으로 개선할 필요가 큰 것은 무엇입니까? (최대 2개)')
    .setChoiceValues([
      '편향된 정보·인식',
      'AI 시대 비판적 판단·정보 검증 역량 부족',
      '주입식 교육·토론 경험 부족',
      '다양한 관점 이해·갈등 해결 역량 부족',
      '해당 없음 / 판단 어려움',
    ])
    .showOtherOption(true)
    .setValidation(atMost_(2))
    .setRequired(true);

  form.addCheckboxItem()
    .setTitle('Q5. 학교에서 시민교육이 충분히 이루어지지 않을 경우 가장 우려되는 결과는 무엇입니까? (최대 2개)')
    .setChoiceValues([
      '타인·사회문제 이해 부족',
      '허위정보·편향적 주장에 대한 취약성',
      '민주적 의사결정·참여 경험 부족',
      '갈등·혐오 표현 대응 역량 부족',
      '판단 어려움',
    ])
    .showOtherOption(true)
    .setValidation(atMost_(2))
    .setRequired(true);

  // ── 6쪽: 원칙별 중요도와 우선 구체화 ───────────────────────────────
  form.addPageBreakItem()
    .setTitle('원칙별 중요도와 우선 구체화')
    .setHelpText(PRINCIPLES_INTRO);

  form.addGridItem()
    .setTitle('Q6. 학교시민교육에서 다음 원칙이 얼마나 중요하다고 생각합니까?')
    .setHelpText('‘학생의 안전한 표현·참여’와 ‘교원 보호’는 안전성(제14조)을 나누어 여쭙는 항목입니다.')
    .setRows(PRINCIPLE_ROWS)
    .setColumns([
      '① 전혀 중요하지 않다',
      '② 중요하지 않은 편',
      '③ 보통',
      '④ 중요한 편',
      '⑤ 매우 중요하다',
      '판단 어려움',
    ])
    .setRequired(true);

  form.addCheckboxItem()
    .setTitle('Q7. 위 원칙 가운데 경기 학교 현장에서 가장 우선 구체화할 항목은 무엇입니까? (최대 2개)')
    .setChoiceValues(PRINCIPLE_ROWS)
    .setValidation(atMost_(2))
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('Q7-1. 그 항목을 선택하신 이유는 무엇입니까? (선택)')
    .setRequired(false);

  form.addParagraphTextItem()
    .setTitle('Q8. 기존 원칙만으로 다루기 어렵거나 경기형으로 보완해야 할 기준은 무엇입니까?')
    .setHelpText('없으면 ‘없음’이라고 적어 주십시오.')
    .setRequired(true);

  // ── 7쪽: 현장 정착을 위한 조건 ─────────────────────────────────────
  form.addPageBreakItem()
    .setTitle('현장 정착을 위한 조건')
    .setHelpText('이 쪽의 응답은 2차 토의에서 운영 가이드·상황별 매뉴얼을 설계하는 데 활용됩니다.');

  form.addCheckboxItem()
    .setTitle('Q9. 원칙을 현장에 정착시키는 데 가장 큰 걸림돌은 무엇입니까? (최대 3개)')
    .setChoiceValues([
      '교사 역량·연수 부족',
      '토론 운영 가이드 부재',
      '상황별 대응 매뉴얼 부재',
      '외부 민원·간섭',
      '학생·보호자·교직원 간 갈등',
      '자료·사례 접근성 부족',
      '시간·업무 부담',
    ])
    .showOtherOption(true)
    .setValidation(atMost_(3))
    .setRequired(true);

  form.addCheckboxItem()
    .setTitle('Q10. 원칙 정착을 위해 가장 우선 필요한 지원은 무엇입니까? (최대 3개)')
    .setChoiceValues([
      '교사·퍼실리테이터 교육',
      '토론 운영 가이드',
      '민원·갈등 상황별 매뉴얼',
      '온라인 자료·사례 검색 체계',
      '수업 주제·안건 설계 지원',
      '학교·교육청 공동 대응 절차',
    ])
    .showOtherOption(true)
    .setValidation(atMost_(3))
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('Q11. 현장에서 원칙 적용이 어렵거나 갈등이 발생한 실제·가상 상황이 있다면 제시해 주십시오. (선택)')
    .setHelpText('※ 개인 이름, 학교명 등 개인·학교를 식별할 수 있는 정보는 입력하지 마십시오.')
    .setRequired(false);

  form.addMultipleChoiceItem()
    .setTitle('Q12. 성공적인 정착을 위해 가장 먼저 역할이 필요한 주체를 하나 선택해 주십시오. (선택)')
    .setChoiceValues(['학교', '교육청', '교사', '학생', '보호자'])
    .setRequired(false);

  form.addParagraphTextItem()
    .setTitle('Q12-1. 선택하신 주체에게 필요한 조건이나 역할을 한 가지 제안해 주십시오. (선택)')
    .setRequired(false);

  form.addParagraphTextItem()
    .setTitle('Q13. 이번 숙의공론에서 꼭 다뤄야 할 사안이 있다면 말씀해 주십시오. (선택)')
    .setRequired(false);

  // ── 동의 문항 이동: 동의함 → 다음 쪽, 동의하지 않음 → 바로 제출 ─────
  consent.setChoices([
    consent.createChoice('동의함', FormApp.PageNavigationType.CONTINUE),
    consent.createChoice('동의하지 않음', FormApp.PageNavigationType.SUBMIT),
  ]);

  // ── 응답 시트 연결 ─────────────────────────────────────────────────
  const sheet = SpreadsheetApp.create(FORM_TITLE + ' (응답)');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());

  // ── 개인별 링크 형식: {ID} 자리에 참여단별 확인번호를 넣어 문자 발송 ───
  const prefilled = form.createResponse()
    .withItemResponse(idItem.createResponse('ID_PLACEHOLDER'))
    .toPrefilledUrl()
    .replace('ID_PLACEHOLDER', '{ID}');

  Logger.log('편집 URL: ' + form.getEditUrl());
  Logger.log('응답 URL: ' + form.getPublishedUrl());
  Logger.log('응답 시트: ' + sheet.getUrl());
  Logger.log('개인별 링크 형식: ' + prefilled);
}

function atMost_(n) {
  return FormApp.createCheckboxValidation()
    .setHelpText('최대 ' + n + '개까지 선택해 주십시오.')
    .requireSelectAtMost(n)
    .build();
}
