/* Same-origin navigation without a document swap.
 *
 * Every click used to be a full document load: the browser tore down the page
 * and rebuilt it, chrome and all, which is what read as a flicker. The notice,
 * masthead, sticky menu and footer are identical on every page and live outside
 * <main>, so only <main> has to change. Nothing else repaints.
 *
 * Progressive enhancement throughout: if anything here fails or is unsupported,
 * the link is left alone and the browser navigates normally. */
(function () {
  if (!window.fetch || !window.history || !history.pushState || !window.DOMParser) return;

  var main = document.querySelector('main');
  var nav = document.querySelector('.nav');
  if (!main) return;

  var cache = new Map();
  var LIMIT = 12;

  /* Which document is currently rendered, ignoring the fragment. Chrome fires
     popstate for same-document fragment navigation as well as for real history
     traversal, so without this the handler would re-render <main> and reset the
     scroll every time someone clicked an in-page anchor. */
  var current = location.pathname + location.search;

  function internal(a) {
    if (!a || !a.href || a.hasAttribute('download')) return null;
    if (a.target && a.target !== '_self') return null;
    var u;
    try { u = new URL(a.href, location.href); } catch (e) { return null; }
    if (u.origin !== location.origin) return null;
    // leave assets and anything that is not a page to the browser
    var last = u.pathname.split('/').pop();
    if (last && last.indexOf('.') !== -1 && !/\.html?$/i.test(last)) return null;
    return u;
  }

  function fetchPage(url) {
    if (cache.has(url)) return cache.get(url);
    var p = fetch(url, { credentials: 'same-origin' }).then(function (r) {
      if (!r.ok) throw new Error(r.status);
      return r.text();
    });
    if (cache.size > LIMIT) cache.clear();
    cache.set(url, p);
    return p;
  }

  function markCurrent(pathname) {
    if (!nav) return;
    var links = nav.querySelectorAll('a');
    for (var i = 0; i < links.length; i++) {
      var u = internal(links[i]);
      if (u && u.pathname === pathname) links[i].setAttribute('aria-current', 'page');
      else links[i].removeAttribute('aria-current');
    }
  }

  function render(html, url, scroll) {
    var doc = new DOMParser().parseFromString(html, 'text/html');
    var fresh = doc.querySelector('main');
    if (!fresh) { location.href = url; return; }

    main.replaceWith(fresh);
    main = fresh;
    if (doc.title) document.title = doc.title;

    var canonical = document.querySelector('link[rel=canonical]');
    var freshCanonical = doc.querySelector('link[rel=canonical]');
    if (canonical && freshCanonical) canonical.href = freshCanonical.href;

    var target = new URL(url, location.href);
    current = target.pathname + target.search;
    markCurrent(target.pathname);
    window.scrollTo(0, scroll || 0);

    /* Move focus into the new content so keyboard and screen reader users are
       not left on a link that no longer exists. */
    main.setAttribute('tabindex', '-1');
    main.focus({ preventScroll: true });
  }

  function go(url, push) {
    fetchPage(url).then(function (html) {
      if (push) {
        history.replaceState({ y: window.scrollY }, '');
        history.pushState({ y: 0 }, '', url);
      }
      render(html, url, 0);
    })['catch'](function () { location.href = url; });
  }

  document.addEventListener('click', function (e) {
    if (e.defaultPrevented || e.button !== 0) return;
    if (e.metaKey || e.ctrlKey || e.shiftKey || e.altKey) return;
    var a = e.target.closest ? e.target.closest('a[href]') : null;
    var u = internal(a);
    if (!u) return;
    // in-page anchors stay with the browser
    if (u.pathname === location.pathname && u.hash) return;
    if (u.href === location.href) { e.preventDefault(); return; }
    e.preventDefault();
    go(u.href, true);
  });

  document.addEventListener('mouseover', function (e) {
    var a = e.target.closest ? e.target.closest('a[href]') : null;
    var u = internal(a);
    if (u && u.href !== location.href) fetchPage(u.href)['catch'](function () {});
  });

  window.addEventListener('popstate', function (e) {
    // a fragment change on the page already shown: let the browser scroll
    if (location.pathname + location.search === current) return;
    var y = (e.state && e.state.y) || 0;
    fetchPage(location.href).then(function (html) {
      render(html, location.href, y);
    })['catch'](function () { location.reload(); });
  });

  /* Warm the menu once the page is idle: six documents of about 20KB, so the
     first click on any of them is already in memory. */
  var warm = function () {
    if (!nav) return;
    var links = nav.querySelectorAll('a');
    for (var i = 0; i < links.length; i++) {
      var u = internal(links[i]);
      if (u && u.href !== location.href) fetchPage(u.href)['catch'](function () {});
    }
  };
  if (window.requestIdleCallback) requestIdleCallback(warm, { timeout: 3000 });
  else setTimeout(warm, 1200);
})();
