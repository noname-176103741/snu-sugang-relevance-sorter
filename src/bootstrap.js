(function bootstrapSugangRelevanceSorter() {
  'use strict';

  const root = document.documentElement;
  root.classList.add('srs-active');

  // 사이트 변경이나 예기치 않은 스크립트 오류가 있어도 원래 결과를 영구히 가리지 않는다.
  window.setTimeout(() => root.classList.add('srs-ready'), 20_000);
}());
