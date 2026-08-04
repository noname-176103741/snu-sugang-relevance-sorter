(function exposeSugangRelevanceCore(globalObject) {
  'use strict';

  const COLUMN_INDEX = Object.freeze({
    curriculum: 0,
    college: 1,
    department: 2,
    program: 3,
    grade: 4,
    courseCode: 5,
    lectureNumber: 6,
    title: 7,
    subtitle: 8,
    credit: 9,
    lectureHours: 10,
    practiceHours: 11,
    schedule: 12,
    classType: 13,
    classroom: 14,
    professor: 15,
    basketCount: 16,
    returningBasketCount: 17,
    freshmanBasketCount: 18,
    capacity: 19,
    enrolled: 20,
    note: 21,
    language: 22,
    status: 23,
  });

  function toText(value) {
    return value == null ? '' : String(value).trim();
  }

  function normalizeForSearch(value) {
    return toText(value)
      .normalize('NFKC')
      .toLocaleLowerCase('ko-KR')
      .replace(/[\p{P}\p{S}]+/gu, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function compact(value) {
    return normalizeForSearch(value).replace(/\s+/g, '');
  }

  function uniqueTokens(value) {
    const normalized = normalizeForSearch(value);
    const tokens = normalized.split(' ').filter(Boolean);
    return [...new Set(tokens)];
  }

  function ngrams(value, size) {
    const normalized = compact(value);
    if (!normalized) return [];
    if (normalized.length <= size) return [normalized];

    const grams = [];
    for (let index = 0; index <= normalized.length - size; index += 1) {
      grams.push(normalized.slice(index, index + size));
    }
    return grams;
  }

  function diceCoefficient(left, right) {
    const leftCompact = compact(left);
    const rightCompact = compact(right);
    if (!leftCompact || !rightCompact) return 0;
    if (leftCompact === rightCompact) return 1;

    const size = Math.min(leftCompact.length, rightCompact.length) < 3 ? 1 : 2;
    const leftGrams = ngrams(leftCompact, size);
    const rightGrams = ngrams(rightCompact, size);
    const counts = new Map();
    for (const gram of leftGrams) counts.set(gram, (counts.get(gram) || 0) + 1);

    let intersection = 0;
    for (const gram of rightGrams) {
      const count = counts.get(gram) || 0;
      if (count > 0) {
        intersection += 1;
        counts.set(gram, count - 1);
      }
    }
    return (2 * intersection) / (leftGrams.length + rightGrams.length);
  }

  function fieldScore(value, query, weights) {
    const normalized = normalizeForSearch(value);
    const queryNormalized = normalizeForSearch(query);
    const compactValue = compact(normalized);
    const compactQuery = compact(queryNormalized);
    if (!normalized || !queryNormalized) return 0;

    let score = 0;
    if (normalized === queryNormalized || compactValue === compactQuery) {
      score += weights.exact;
    } else if (normalized.startsWith(queryNormalized) || compactValue.startsWith(compactQuery)) {
      score += weights.prefix;
    } else if (normalized.includes(queryNormalized) || compactValue.includes(compactQuery)) {
      score += weights.contains;
    }

    const words = new Set(normalized.split(' ').filter(Boolean));
    for (const token of uniqueTokens(queryNormalized)) {
      const compactToken = compact(token);
      if (words.has(token)) score += weights.tokenExact;
      else if (compactValue.includes(compactToken)) score += weights.tokenContains;
    }

    score += Math.round(diceCoefficient(normalized, queryNormalized) * weights.fuzzy);
    return score;
  }

  function relevanceScore(course, query) {
    if (!normalizeForSearch(query)) return 0;

    return fieldScore(course.title, query, {
      exact: 12_000,
      prefix: 7_000,
      contains: 4_500,
      tokenExact: 800,
      tokenContains: 350,
      fuzzy: 450,
    }) + fieldScore(course.subtitle, query, {
      exact: 5_500,
      prefix: 3_200,
      contains: 2_000,
      tokenExact: 450,
      tokenContains: 200,
      fuzzy: 240,
    });
  }

  function sortCourses(courses, query) {
    return courses
      .map((course, index) => ({
        ...course,
        originalIndex: Number.isInteger(course.originalIndex) ? course.originalIndex : index,
        relevance: relevanceScore(course, query),
      }))
      .sort((left, right) => right.relevance - left.relevance || left.originalIndex - right.originalIndex);
  }

  function findHeaderIndex(rows) {
    const exact = rows.findIndex((row) => row.some((cell) => toText(cell) === '교과목명')
      && row.some((cell) => toText(cell) === '교과목번호'));
    if (exact >= 0) return exact;

    const likely = rows.findIndex((row) => row.length >= 20
      && row.filter((cell) => toText(cell)).length >= 12);
    return likely >= 0 ? likely : 2;
  }

  function courseFromRow(row, originalIndex) {
    const value = (key) => toText(row[COLUMN_INDEX[key]]);
    return {
      originalIndex,
      curriculum: value('curriculum'),
      college: value('college'),
      department: value('department'),
      program: value('program'),
      grade: value('grade'),
      courseCode: value('courseCode'),
      lectureNumber: value('lectureNumber'),
      title: value('title'),
      subtitle: value('subtitle'),
      credit: value('credit'),
      lectureHours: value('lectureHours'),
      practiceHours: value('practiceHours'),
      schedule: value('schedule'),
      classType: value('classType'),
      classroom: value('classroom'),
      professor: value('professor'),
      basketCount: value('basketCount'),
      returningBasketCount: value('returningBasketCount'),
      freshmanBasketCount: value('freshmanBasketCount'),
      capacity: value('capacity'),
      enrolled: value('enrolled'),
      note: value('note'),
      language: value('language'),
      status: value('status'),
    };
  }

  function parseExcelRows(rows) {
    if (!Array.isArray(rows) || rows.length === 0) {
      return { courses: [], headerIndex: -1 };
    }

    const headerIndex = findHeaderIndex(rows);
    const courses = rows
      .slice(headerIndex + 1)
      .map((row, index) => courseFromRow(Array.isArray(row) ? row : [], index))
      .filter((course) => course.courseCode && course.title);
    return { courses, headerIndex };
  }

  function getPagination(page, totalItems, pageSize = 10, blockSize = 5) {
    const totalPages = Math.max(1, Math.ceil(totalItems / pageSize));
    const currentPage = Math.min(totalPages, Math.max(1, Number(page) || 1));
    const blockStart = Math.floor((currentPage - 1) / blockSize) * blockSize + 1;
    const blockEnd = Math.min(totalPages, blockStart + blockSize - 1);
    const pages = [];
    for (let value = blockStart; value <= blockEnd; value += 1) pages.push(value);

    return {
      currentPage,
      totalPages,
      pages,
      firstPage: 1,
      previousPage: Math.max(1, blockStart - 1),
      nextPage: Math.min(totalPages, blockEnd + 1),
      lastPage: totalPages,
      canGoPrevious: currentPage > 1,
      canGoNext: currentPage < totalPages,
      sliceStart: (currentPage - 1) * pageSize,
      sliceEnd: Math.min(totalItems, currentPage * pageSize),
    };
  }

  const api = Object.freeze({
    COLUMN_INDEX,
    compact,
    diceCoefficient,
    getPagination,
    normalizeForSearch,
    parseExcelRows,
    relevanceScore,
    sortCourses,
  });

  globalObject.SugangRelevanceCore = api;
}(globalThis));
