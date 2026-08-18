const reduceMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

/* ── PRELOADER ── */
window.addEventListener('load', () => {
  setTimeout(() => document.getElementById('loader').classList.add('done'), 500);
  document.querySelector('.hero')?.classList.add('ready');
});

/* ── SCROLL PROGRESS ── */
const progressEl = document.getElementById('scroll-progress');
function updateProgress(){
  const h = document.documentElement;
  const pct = (h.scrollTop) / (h.scrollHeight - h.clientHeight) * 100;
  progressEl.style.width = pct + '%';
}

/* ── NAV SCROLL STATE ── */
const navEl = document.getElementById('nav');
function updateNav(){ navEl.classList.toggle('scrolled', window.scrollY > 30); }

/* ── UNIFIED SCROLL LOOP (progress + nav + parallax) ── */
let ticking = false;
window.addEventListener('scroll', () => {
  if (!ticking){
    requestAnimationFrame(() => { updateProgress(); updateNav(); parallaxTick(); ticking = false; });
    ticking = true;
  }
});
updateProgress(); updateNav();

/* ── PARALLAX LAYERS ── */
const parallaxEls = [...document.querySelectorAll('[data-parallax]')];
function parallaxTick(){
  if (reduceMotion) return;
  const y = window.scrollY;
  parallaxEls.forEach(el => {
    const speed = parseFloat(el.dataset.parallax);
    el.style.transform = `translateY(${y * speed}px)`;
  });
}

/* ── CUSTOM CURSOR ── */
const curDot = document.getElementById('cur-dot');
const curRing = document.getElementById('cur-ring');
let mx=0,my=0,rx=0,ry=0;
if (!('ontouchstart' in window)) {
  document.addEventListener('mousemove', e => { mx=e.clientX; my=e.clientY; curDot.style.left=mx+'px'; curDot.style.top=my+'px'; });
  (function ringLoop(){ rx += (mx-rx)*0.18; ry += (my-ry)*0.18; curRing.style.left=rx+'px'; curRing.style.top=ry+'px'; requestAnimationFrame(ringLoop); })();
  document.addEventListener('mouseover', e => { if(e.target.closest('a,button,summary,.cursor-hover')) curRing.classList.add('hover'); });
  document.addEventListener('mouseout', e => { if(e.target.closest('a,button,summary,.cursor-hover')) curRing.classList.remove('hover'); });
}

/* ── MAGNETIC BUTTONS ── */
document.querySelectorAll('.magnetic').forEach(btn => {
  btn.addEventListener('mousemove', e => {
    const r = btn.getBoundingClientRect();
    const x = (e.clientX - r.left - r.width/2) * 0.3;
    const y = (e.clientY - r.top - r.height/2) * 0.3;
    btn.style.transform = `translate(${x}px, ${y}px)`;
  });
  btn.addEventListener('mouseleave', () => { btn.style.transform = ''; });
});

/* ── MOBILE NAV ── */
function toggleMNav(){ document.getElementById('mnav').classList.toggle('open'); }
document.querySelectorAll('#mnav a').forEach(a => a.addEventListener('click', () => document.getElementById('mnav').classList.remove('open')));

/* ── NAV ACTIVE STATE (per current page) ── */
const currentPage = document.body.dataset.page;
document.querySelectorAll('.nav-links a, #mnav a').forEach(a => {
  a.classList.toggle('active', a.dataset.page === currentPage);
});

/* ── SCROLL REVEAL ── */
const revealObs = new IntersectionObserver((entries) => {
  entries.forEach(en => { if (en.isIntersecting){ en.target.classList.add('on'); revealObs.unobserve(en.target); } });
}, { threshold: 0.15, rootMargin: '0px 0px -8% 0px' });
function initReveal(){
  document.querySelectorAll('.rv, .rv-scale, .rv-l, .rv-r').forEach((el,i) => {
    el.style.setProperty('--i', i % 8);
    revealObs.observe(el);
  });
}

/* ── COUNTERS ── */
function animateCounters(){
  document.querySelectorAll('.cnt[data-target]').forEach(el => {
    if (el.dataset.done) return;
    const io = new IntersectionObserver((entries) => {
      entries.forEach(en => {
        if (en.isIntersecting){
          el.dataset.done = '1';
          const target = parseFloat(el.dataset.target);
          const suffix = el.dataset.suffix || '';
          const decimals = el.dataset.decimals ? parseInt(el.dataset.decimals) : 0;
          let cur = 0; const step = target / 50;
          const t = setInterval(() => {
            cur = Math.min(cur + step, target);
            el.textContent = cur.toFixed(decimals) + suffix;
            if (cur >= target) clearInterval(t);
          }, 22);
          io.disconnect();
        }
      });
    }, { threshold: 0.4 });
    io.observe(el);
  });
}
animateCounters();

/* ── PROCESS SCROLLYTELLING ── */
const steps = [...document.querySelectorAll('.process-step')];
const figIcon = document.querySelector('.process-fig .pf-icon');
if (steps.length){
  const stepObs = new IntersectionObserver((entries) => {
    entries.forEach(en => {
      if (en.isIntersecting){
        steps.forEach(s => s.classList.remove('active'));
        en.target.classList.add('active');
        if (figIcon) { figIcon.style.opacity = 0; setTimeout(() => { figIcon.textContent = en.target.dataset.icon; figIcon.style.opacity = 1; }, 200); }
      }
    });
  }, { threshold: 0.6, rootMargin: '-30% 0px -30% 0px' });
  steps.forEach(s => stepObs.observe(s));
}

/* ── TESTIMONIAL CAROUSEL ── */
(function(){
  const slides = [...document.querySelectorAll('.tm-slide')];
  const dots = [...document.querySelectorAll('.tm-dot')];
  if (!slides.length) return;
  let idx = 0, timer;
  function go(i){
    slides[idx].classList.remove('active'); dots[idx].classList.remove('active');
    idx = (i + slides.length) % slides.length;
    slides[idx].classList.add('active'); dots[idx].classList.add('active');
  }
  function auto(){ clearInterval(timer); timer = setInterval(() => go(idx+1), 5500); }
  dots.forEach((d,i) => d.addEventListener('click', () => { go(i); auto(); }));
  document.getElementById('tm-prev')?.addEventListener('click', () => { go(idx-1); auto(); });
  document.getElementById('tm-next')?.addEventListener('click', () => { go(idx+1); auto(); });
  auto();
})();

/* ── PORTFOLIO FILTER ── */
function filterPF(cat, btn){
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  document.querySelectorAll('.pf-card').forEach(c => {
    c.style.display = (cat === 'all' || c.dataset.cat === cat) ? '' : 'none';
  });
}

/* ── FORM SUBMIT ── */
function handleSubmit(btn){
  const original = btn.innerHTML;
  btn.innerHTML = '✓ Message sent — we\'ll reply within 2 hours';
  btn.style.background = 'linear-gradient(120deg,#22c55e,#16a34a)';
  setTimeout(() => { btn.innerHTML = original; btn.style.background = ''; }, 4000);
}

/* ── AURORA BACKGROUND CANVAS ── */
(function(){
  const c = document.getElementById('aurora');
  const ctx = c.getContext('2d');
  let W,H;
  function resize(){ W = c.width = window.innerWidth; H = c.height = window.innerHeight; }
  resize(); window.addEventListener('resize', resize);
  const blobs = [
    { x:.18, y:.18, r:.42, hue:'123,47,190', s:0.00018 },
    { x:.82, y:.28, r:.38, hue:'34,197,240', s:0.00013 },
    { x:.5,  y:.78, r:.46, hue:'233,30,140', s:0.00021 },
  ];
  let t = 0;
  function draw(){
    t += 1;
    ctx.clearRect(0,0,W,H);
    ctx.fillStyle = '#05060f'; ctx.fillRect(0,0,W,H);
    blobs.forEach((b,i) => {
      const cx = (b.x + Math.sin(t*b.s + i)*0.06) * W;
      const cy = (b.y + Math.cos(t*b.s*1.3 + i)*0.06) * H;
      const r = b.r * Math.max(W,H);
      const g = ctx.createRadialGradient(cx,cy,0,cx,cy,r);
      g.addColorStop(0, `rgba(${b.hue},0.35)`);
      g.addColorStop(1, `rgba(${b.hue},0)`);
      ctx.fillStyle = g;
      ctx.beginPath(); ctx.arc(cx,cy,r,0,Math.PI*2); ctx.fill();
    });
    if (!reduceMotion) requestAnimationFrame(draw); else setTimeout(draw, 4000);
  }
  draw();
})();

/* ── INIT ── */
initReveal();
