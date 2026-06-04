/**
 * 移动端 H5 rem 适配（设计稿基准宽度 375px）
 * 1rem = 屏幕宽度 / 10，在 375px 宽屏下 1rem ≈ 37.5px
 * 大屏（平板/桌面预览）封顶 540px，保持手机视觉比例
 */
(function () {
  var docEl = document.documentElement;
  var MAX_WIDTH = 540;
  var MIN_WIDTH = 320;

  function setRootFontSize() {
    var width = docEl.clientWidth || window.innerWidth;
    if (width > MAX_WIDTH) width = MAX_WIDTH;
    if (width < MIN_WIDTH) width = MIN_WIDTH;
    docEl.style.fontSize = width / 10 + "px";
  }

  setRootFontSize();
  window.addEventListener("resize", setRootFontSize);
  window.addEventListener("orientationchange", setRootFontSize);
  window.addEventListener("pageshow", function (e) {
    if (e.persisted) setRootFontSize();
  });
})();
