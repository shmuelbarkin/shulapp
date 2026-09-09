(function () {
  var KEY = 'mms-notice-dismissed';
  var notice = document.querySelector('.notice');
  if (!notice) return;

  try {
    if (localStorage.getItem(KEY) === '1') notice.hidden = true;
  } catch (e) {
    /* storage unavailable (private mode, blocked cookies) - leave it visible */
  }

  var close = notice.querySelector('.notice__close');
  if (!close) return;

  close.addEventListener('click', function () {
    notice.hidden = true;
    try {
      localStorage.setItem(KEY, '1');
    } catch (e) {
      /* nothing to persist to; it returns on the next page load */
    }
  });
})();
