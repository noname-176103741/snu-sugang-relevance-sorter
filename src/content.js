(function runSugangRelevanceSorter() {
  'use strict';

  const PAGE_SIZE = 10;
  const SEARCH_FIELD_NAMES = [
    'workType',
    'pageNo',
    'srchOpenSchyy',
    'srchOpenShtm',
    'srchSbjtNm',
    'srchSbjtCd',
    'seeMore',
    'srchCptnCorsFg',
    'srchOpenShyr',
    'srchOpenUpSbjtFldCd',
    'srchOpenSbjtFldCd',
    'srchOpenUpDeptCd',
    'srchOpenDeptCd',
    'srchOpenMjCd',
    'srchOpenSubmattCorsFg',
    'srchOpenSubmattFgCd1',
    'srchOpenSubmattFgCd2',
    'srchOpenSubmattFgCd3',
    'srchOpenSubmattFgCd4',
    'srchOpenSubmattFgCd5',
    'srchOpenSubmattFgCd6',
    'srchOpenSubmattFgCd7',
    'srchOpenSubmattFgCd8',
    'srchOpenSubmattFgCd9',
    'srchExcept',
    'srchOpenPntMin',
    'srchOpenPntMax',
    'srchCamp',
    'srchBdNo',
    'srchProfNm',
    'srchOpenSbjtTmNm',
    'srchOpenSbjtDayNm',
    'srchOpenSbjtTm',
    'srchOpenSbjtNm',
    'srchTlsnAplyCapaCntMin',
    'srchTlsnAplyCapaCntMax',
    'srchLsnProgType',
    'srchTlsnRcntMin',
    'srchTlsnRcntMax',
    'srchMrksGvMthd',
    'srchIsEngSbjt',
    'srchMrksApprMthdChgPosbYn',
    'srchIsPendingCourse',
    'srchGenrlRemoteLtYn',
    'srchLanguage',
    'srchCurrPage',
    'srchPageSize',
  ];

  const state = {
    courses: [],
    currentPage: 1,
    query: '',
  };

  function createElement(tagName, options = {}) {
    const element = document.createElement(tagName);
    if (options.className) element.className = options.className;
    if (options.text != null) element.textContent = String(options.text);
    if (options.attributes) {
      for (const [name, value] of Object.entries(options.attributes)) {
        element.setAttribute(name, String(value));
      }
    }
    return element;
  }

  function directNamedElement(form, name) {
    const candidate = form?.elements?.namedItem(name);
    if (!candidate) return null;
    if (typeof RadioNodeList !== 'undefined' && candidate instanceof RadioNodeList) {
      return candidate[0] || null;
    }
    return candidate;
  }

  function buildExcelParameters() {
    const resultForm = document.querySelector('form#CC100');
    const excelForm = document.querySelector('form#HD102');
    if (!resultForm) throw new Error('강좌검색 결과 폼(CC100)을 찾지 못했습니다.');

    const parameters = new URLSearchParams();
    for (const name of SEARCH_FIELD_NAMES) {
      const source = directNamedElement(resultForm, name) || directNamedElement(excelForm, name);
      if (source) parameters.set(name, source.value || '');
    }
    parameters.set('workType', 'EX');
    parameters.set('pageNo', '1');
    if (!parameters.has('srchCurrPage')) parameters.set('srchCurrPage', '1');
    if (!parameters.has('srchPageSize')) parameters.set('srchPageSize', '9999');
    return parameters;
  }

  function findQuery() {
    const form = document.querySelector('form#CC100');
    return directNamedElement(form, 'srchSbjtNm')?.value?.trim() || '';
  }

  async function fetchAllCourses() {
    const response = await fetch('/sugang/cc/cc100InterfaceExcel.action', {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8',
      },
      body: buildExcelParameters().toString(),
    });
    if (!response.ok) throw new Error(`Excel 응답 오류 (${response.status})`);

    const contentType = response.headers.get('content-type') || '';
    if (!contentType.toLocaleLowerCase().includes('excel')
      && !contentType.toLocaleLowerCase().includes('octet-stream')) {
      throw new Error('Excel 형식이 아닌 응답을 받았습니다. 선택한 학기의 Excel 저장 가능 여부를 확인하세요.');
    }

    const buffer = await response.arrayBuffer();
    if (buffer.byteLength < 8) throw new Error('Excel 응답이 비어 있습니다.');

    const workbook = XLSX.read(buffer, { type: 'array', cellDates: false, cellText: true });
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) throw new Error('Excel 워크시트를 찾지 못했습니다.');

    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], {
      header: 1,
      defval: '',
      raw: false,
      blankrows: false,
    });
    return SugangRelevanceCore.parseExcelRows(rows).courses;
  }

  function insertStatus() {
    const existing = document.querySelector('.srs-status');
    if (existing) return existing;

    const status = createElement('div', {
      className: 'srs-status srs-loading',
      text: 'Excel 전체 검색 결과를 불러와 관련도순으로 정렬하고 있습니다…',
      attributes: { role: 'status', 'aria-live': 'polite' },
    });
    const resultSummary = document.querySelector('.search-result-con');
    resultSummary?.insertAdjacentElement('afterend', status);
    return status;
  }

  function updateStatus(kind, message) {
    const status = insertStatus();
    status.className = `srs-status srs-${kind}`;
    status.textContent = message;
  }

  function hiddenInput(name, value, index) {
    return createElement('input', {
      attributes: {
        type: 'hidden',
        id: `${name}_${index}`,
        name,
        value: value || '',
      },
    });
  }

  function semesterParts() {
    const combined = directNamedElement(document.querySelector('form#CC100'), 'srchOpenShtm')?.value || '';
    const match = combined.match(/^(U\d{9})(U\d{9})$/);
    if (match) return { semester: match[1], detailSemester: match[2] };
    return { semester: combined.slice(0, 10), detailSemester: combined.slice(10) };
  }

  function appendCourseFlags(container, course) {
    const note = course.note.normalize('NFKC').toLocaleLowerCase('ko-KR');
    const flags = [];
    if (/ⓞ|원격/.test(note)) flags.push({ label: 'O', title: '원격수업강좌' });
    if (/ⓜ|군휴학/.test(note)) flags.push({ label: 'M', title: '군휴학생 원격수업 강좌' });
    if (/ⓒ|cross|크로스/.test(note)) flags.push({ label: 'C', title: 'Cross-Listing' });
    if (/ⓡ|수강반/.test(note)) flags.push({ label: 'R', title: '수강반 제한' });
    if (/ⓚ|거점/.test(note)) flags.push({ label: 'K', title: '거점국립대학 원격수업 강좌' });

    for (const flag of flags) {
      const wrapper = createElement('div', { attributes: { title: flag.title } });
      wrapper.append(createElement('span', { className: 'icon-1', text: flag.label }));
      container.append(wrapper);
    }

    const language = course.language.normalize('NFKC').toLocaleLowerCase('ko-KR');
    if (language && !/^(한국어|korean|ko)$/.test(language)) {
      const wrapper = createElement('div', { attributes: { title: course.language } });
      wrapper.append(createElement('span', { className: 'globe' }));
      container.append(wrapper);
    }
  }

  function labeledMetric(label, value) {
    const span = createElement('span', { attributes: { lang: 'ko' } });
    span.append(document.createTextNode(`${label} `), createElement('em', { text: value || '-' }));
    return span;
  }

  function createCourseItem(course, index) {
    const item = createElement('div', { className: 'course-info-item' });
    const left = createElement('div', { className: 'left' });
    const label = createElement('label', { className: 'cc-check-item round full' });
    const radio = createElement('input', {
      attributes: {
        type: 'radio',
        name: 'check',
        value: index,
        'aria-label': `${course.title} 선택`,
      },
    });
    const year = directNamedElement(document.querySelector('form#CC100'), 'srchOpenSchyy')?.value || '';
    const { semester, detailSemester } = semesterParts();
    label.append(
      radio,
      createElement('em'),
      hiddenInput('openSchyy', year, index),
      hiddenInput('openShtmFg', semester, index),
      hiddenInput('openDetaShtmFg', detailSemester, index),
      hiddenInput('sbjtNm', course.title, index),
      hiddenInput('sbjtCd', course.courseCode, index),
      hiddenInput('ltNo', course.lectureNumber, index),
      hiddenInput('ltTimeForTT', course.schedule, index),
      // Excel에는 부제 코드가 없으며, 사이트는 부제가 없는 강좌에 기본값 000을 사용한다.
      // 실제 수강신청 함수는 교과목번호와 강좌번호만 전송한다.
      hiddenInput('sbjtSubhCd', '000', index),
      hiddenInput('lsnTmtablFormaSmryCtnt_hidden', '', index),
    );
    left.append(label);

    const main = createElement('div');
    const detail = createElement('a', {
      className: 'course-info-detail',
      attributes: {
        href: 'javascript:void(0)',
        id: `course_info_detail_${index}`,
        name: 'course_info_detail',
        'data-srs-detail': 'true',
        'data-course-code': course.courseCode,
        'data-lecture-number': course.lectureNumber,
      },
    });
    const courseName = createElement('div', { className: 'course-name' });
    courseName.append(document.createTextNode(
      `${course.program ? `[${course.program}] ` : ''}${course.curriculum ? `[${course.curriculum}] ` : ''}`,
    ));
    const strong = createElement('strong');
    strong.append(document.createTextNode(course.title));
    if (course.subtitle) {
      strong.append(createElement('span', { className: 'srs-subtitle', text: `(${course.subtitle})` }));
    }
    courseName.append(strong);

    const info = createElement('ul', { className: 'course-info' });
    const identity = createElement('li', { className: 'txt' });
    identity.append(
      createElement('span', { text: `${course.professor || ''}\u00a0 ` }),
      createElement('span', { text: `${course.department || course.college || ''}\u00a0 ` }),
      createElement('span', { text: `${course.courseCode}(${course.lectureNumber})` }),
    );
    const metrics = createElement('li', { className: 'txt' });
    metrics.append(
      labeledMetric('수강신청인원/정원(재학생)', `${course.enrolled || '0'}/${course.capacity || '-'}`),
      labeledMetric('총수강인원', course.enrolled || '0'),
      labeledMetric('학점', course.credit || '-'),
      createElement('span', { text: course.schedule || '' }),
    );
    const stateItem = createElement('li', { className: 'state' });
    const flags = createElement('div', { className: 'icon-remo' });
    appendCourseFlags(flags, course);
    stateItem.append(flags);
    if (course.note) stateItem.append(createElement('div', { className: 'srs-note', text: course.note }));
    info.append(identity, metrics, stateItem);

    const icons = createElement('div', { className: 'course-icons' });
    if (course.basketCount && course.basketCount !== '0') {
      const baskets = createElement('span', {
        className: 'carts',
        text: course.basketCount,
        attributes: { style: 'color:#376DC8' },
      });
      baskets.prepend(createElement('em', { attributes: { title: '장바구니' } }));
      icons.append(baskets);
    }

    detail.append(courseName, info, icons);
    main.append(detail);
    item.append(left, main);
    return item;
  }

  function pageLink({ className, page, text = '', current = false, disabled = false, label = '' }) {
    const link = createElement('a', {
      className: `${className}${current ? ' on' : ''}`,
      text,
      attributes: {
        href: '#',
        'data-srs-page': page,
        'aria-label': label || `${page}페이지`,
      },
    });
    if (current) link.setAttribute('aria-current', 'page');
    if (disabled) link.setAttribute('aria-disabled', 'true');
    return link;
  }

  function renderPagination(model) {
    const paging = document.querySelector('.cc-paging');
    if (!paging) return;
    paging.replaceChildren();

    paging.append(
      pageLink({ className: 'arrow first', page: model.firstPage, disabled: !model.canGoPrevious, label: '첫 페이지' }),
      pageLink({ className: 'arrow prev', page: model.previousPage, disabled: !model.canGoPrevious, label: '이전 페이지 묶음' }),
    );
    const numberGroup = createElement('span');
    for (const page of model.pages) {
      numberGroup.append(pageLink({ className: 'num', page, text: page, current: page === model.currentPage }));
    }
    paging.append(
      numberGroup,
      pageLink({ className: 'arrow next', page: model.nextPage, disabled: !model.canGoNext, label: '다음 페이지 묶음' }),
      pageLink({ className: 'arrow last', page: model.lastPage, disabled: !model.canGoNext, label: '마지막 페이지' }),
    );
  }

  function updateNativeCount(count) {
    for (const element of document.querySelectorAll('.search-result-con .fc-orange, .total-list-count .num')) {
      element.textContent = String(count);
    }
  }

  function updateNativePageFields(page) {
    const form = document.querySelector('form#CC100');
    const pageNo = directNamedElement(form, 'pageNo');
    const currentPage = directNamedElement(form, 'srchCurrPage');
    if (pageNo) pageNo.value = String(page);
    if (currentPage) currentPage.value = String(page);
  }

  function renderPage(page, shouldScroll = false) {
    const list = document.querySelector('.course-info-list');
    if (!list) throw new Error('강좌 목록 영역을 찾지 못했습니다.');

    const model = SugangRelevanceCore.getPagination(page, state.courses.length, PAGE_SIZE);
    state.currentPage = model.currentPage;
    const courseActions = list.querySelector(':scope > #courseFabDiv');
    list.replaceChildren();

    const visibleCourses = state.courses.slice(model.sliceStart, model.sliceEnd);
    if (visibleCourses.length === 0) {
      list.append(createElement('div', { className: 'srs-empty', text: '검색된 교과목이 없습니다.' }));
    } else {
      visibleCourses.forEach((course, index) => list.append(createCourseItem(course, index)));
    }
    if (courseActions) list.append(courseActions);
    renderPagination(model);
    updateNativeCount(state.courses.length);
    updateNativePageFields(model.currentPage);

    if (shouldScroll) {
      document.querySelector('.search-result-con')?.scrollIntoView({ block: 'start', behavior: 'auto' });
    }
  }

  function openCourseDetail(link) {
    const index = link.id.match(/^course_info_detail_(\d+)$/)?.[1];
    if (index == null) throw new Error('강좌 상세 조회용 목록 번호를 찾지 못했습니다.');

    const copyValue = (targetSelector, sourceName) => {
      const target = document.querySelector(targetSelector);
      const source = document.querySelector(`#${sourceName}_${index}`);
      if (!target || !source) return false;
      target.value = source.value || '';
      return true;
    };

    const modal = document.querySelector('.course-info-detail-md');
    const firstTab = document.querySelector('#tab1');
    const requiredValuesCopied = [
      copyValue('#layer_openSchyy', 'openSchyy'),
      copyValue('#layer_openShtmFg', 'openShtmFg'),
      copyValue('#layer_openDetaShtmFg', 'openDetaShtmFg'),
      copyValue('#layer_sbjtCd', 'sbjtCd'),
      copyValue('#layer_ltNo', 'ltNo'),
      copyValue('#layer_sbjtSubhCd', 'sbjtSubhCd'),
    ].every(Boolean);

    const workType = document.querySelector('#layer_workType');
    if (workType) workType.value = ' ';

    if (!modal || !firstTab || !requiredValuesCopied) {
      throw new Error('사이트의 강좌 상세 모달 구조를 찾지 못했습니다.');
    }

    modal.classList.add('opened');
    const modalSection = modal.querySelector('.md-section');
    if (modalSection) modalSection.scrollTop = 0;

    // 사이트가 탭에 연결해 둔 원래 AJAX 조회와 렌더러를 그대로 사용한다.
    firstTab.click();
  }

  function installClientInteractions() {
    document.addEventListener('click', (event) => {
      const pageTarget = event.target.closest?.('[data-srs-page]');
      if (pageTarget) {
        event.preventDefault();
        if (pageTarget.getAttribute('aria-disabled') !== 'true') {
          renderPage(Number(pageTarget.dataset.srsPage), true);
        }
        return;
      }

      const detailTarget = event.target.closest?.('[data-srs-detail]');
      if (detailTarget) {
        event.preventDefault();
        openCourseDetail(detailTarget);
        return;
      }

      if (event.target.matches?.('.course-info-list input[name="check"]') && event.target.checked) {
        document.querySelector('#inputTextView')?.focus();
      }
    });
  }

  async function initialize() {
    await new Promise((resolve) => {
      if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', resolve, { once: true });
      else resolve();
    });

    if (!document.querySelector('.course-info-list') || !document.querySelector('form#CC100')) {
      document.documentElement.classList.add('srs-ready');
      return;
    }

    insertStatus();
    installClientInteractions();
    state.query = findQuery();
    const courses = await fetchAllCourses();
    state.courses = SugangRelevanceCore.sortCourses(courses, state.query);
    renderPage(1);
    updateStatus(
      'complete',
      `Excel 전체 결과 ${state.courses.length}건을 교과목명·부제명 관련도순으로 정렬했습니다. 페이지 이동은 서버 재조회 없이 처리됩니다.`,
    );
    document.documentElement.classList.add('srs-ready');
  }

  initialize().catch((error) => {
    console.error('[서울대 수강검색 관련도 정렬]', error);
    updateStatus('error', `관련도 정렬을 적용하지 못해 사이트의 기본 결과를 표시합니다: ${error.message}`);
    document.documentElement.classList.add('srs-ready');
  });
}());
