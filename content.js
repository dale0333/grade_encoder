/* content.js — runs in Chrome's isolated world.
   Its ONLY job: inject xlsx + encoder into the PAGE's JS context
   so they can access window.jQuery from AIMS.
*/
(function () {
  function injectFile(url, onDone) {
    var s = document.createElement('script');
    s.src = url;
    s.onload = function () { s.remove(); if (onDone) onDone(); };
    s.onerror = function () { s.remove(); };
    (document.head || document.documentElement).appendChild(s);
  }

  var xlsxUrl    = chrome.runtime.getURL('xlsx.min.js');
  var encoderUrl = chrome.runtime.getURL('encoder.js');

  /* Inject SheetJS first, then encoder after it loads */
  injectFile(xlsxUrl, function () {
    injectFile(encoderUrl, null);
  });
})();
