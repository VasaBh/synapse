(function () {
  var style = document.createElement('style');
  style.textContent =
    '#go-top{' +
      'position:fixed;bottom:1.75rem;right:1.75rem;' +
      'width:2.75rem;height:2.75rem;border-radius:50%;' +
      'background:#4f46e5;color:#fff;border:none;cursor:pointer;' +
      'display:flex;align-items:center;justify-content:center;' +
      'font-size:1.25rem;line-height:1;' +
      'box-shadow:0 4px 14px rgba(79,70,229,.4);' +
      'opacity:0;transform:translateY(10px);' +
      'transition:opacity .22s,transform .22s;' +
      'z-index:9999;pointer-events:none;' +
    '}' +
    '#go-top.show{opacity:1;transform:translateY(0);pointer-events:auto;}' +
    '#go-top:hover{background:#4338ca;box-shadow:0 6px 20px rgba(79,70,229,.5);transform:translateY(-2px);}' +
    '#go-top:active{transform:translateY(0);}' +
    '@media(max-width:480px){#go-top{bottom:1rem;right:1rem;width:2.4rem;height:2.4rem;font-size:1.1rem;}}';
  document.head.appendChild(style);

  var btn = document.createElement('button');
  btn.id = 'go-top';
  btn.setAttribute('aria-label', 'Back to top');
  btn.textContent = '↑'; /* ↑ */
  document.body.appendChild(btn);

  /* Pages may scroll via window OR an internal container (.main-scroll / .content-area).
     We check all candidates and use whichever has actually scrolled. */
  function scrollY() {
    var selectors = ['.main-scroll', '.content-area'];
    for (var i = 0; i < selectors.length; i++) {
      var el = document.querySelector(selectors[i]);
      if (el && el.scrollTop > 0) return el.scrollTop;
    }
    return window.pageYOffset || document.documentElement.scrollTop || 0;
  }

  function update() {
    btn.classList.toggle('show', scrollY() > 280);
  }

  function toTop() {
    var selectors = ['.main-scroll', '.content-area'];
    for (var i = 0; i < selectors.length; i++) {
      var el = document.querySelector(selectors[i]);
      if (el && el.scrollTop > 0) {
        el.scrollTo({ top: 0, behavior: 'smooth' });
        return;
      }
    }
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  window.addEventListener('scroll', update, { passive: true });
  document.querySelectorAll('.main-scroll, .content-area').forEach(function (el) {
    el.addEventListener('scroll', update, { passive: true });
  });
  btn.addEventListener('click', toTop);
})();
