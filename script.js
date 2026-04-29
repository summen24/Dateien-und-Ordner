'use strict';
const TOTAL = 13;
let current = 1, animating = false;
const counter   = document.getElementById('current-slide');
const btnPrev   = document.getElementById('btn-prev');
const btnNext   = document.getElementById('btn-next');
const dotsEl    = document.getElementById('nav-dots');
const progress  = document.getElementById('progress-bar');

const LABELS = ['Titel','Einstieg','Datei vs. Ordner','Datei-Explorer Aufbau',
  'Benennungsregeln','Dateierweiterungen','Shortcuts',
  'AB 1: Begriffe','AB 2: Spickzettel','AB 3: Ordnerstruktur',
  'AB 4: Explorer beschriften','AB 5: Erweiterung','AB 6: Merkregel'];

for (let i = 1; i <= TOTAL; i++) {
  const d = document.createElement('button');
  d.className = 'nav__dot'; d.dataset.target = i;
  d.setAttribute('aria-label', `Folie ${i}: ${LABELS[i-1]}`);
  d.addEventListener('click', () => goTo(i));
  dotsEl.appendChild(d);
}

function goTo(n) {
  if (animating || n < 1 || n > TOTAL || n === current) return;
  animating = true;
  document.querySelectorAll('.slide').forEach(s => { s.style.display='none'; s.classList.remove('is-active'); });
  const next = document.getElementById(`slide-${n}`);
  if (next) { next.style.display='flex'; next.classList.add('is-active'); }
  current = n; updateUI();
  setTimeout(() => { animating = false; }, 400);
}

function updateUI() {
  counter.textContent = current;
  btnPrev.disabled = current === 1;
  btnNext.disabled = current === TOTAL;
  document.querySelectorAll('.nav__dot').forEach((d,i) => d.classList.toggle('is-active', i+1===current));
  progress.style.setProperty('--progress', ((current-1)/(TOTAL-1)*100)+'%');
}

btnPrev.addEventListener('click', () => goTo(current-1));
btnNext.addEventListener('click', () => goTo(current+1));

document.addEventListener('keydown', e => {
  if (e.key==='ArrowRight'||e.key==='ArrowDown'||e.key===' ') { e.preventDefault(); goTo(current+1); }
  if (e.key==='ArrowLeft' ||e.key==='ArrowUp')                { e.preventDefault(); goTo(current-1); }
  if (e.key==='Home') goTo(1);
  if (e.key==='End')  goTo(TOTAL);
  if (e.key==='f'||e.key==='F') toggleFs();
});

let tx=0, ty=0;
document.addEventListener('touchstart', e => { tx=e.touches[0].clientX; ty=e.touches[0].clientY; }, {passive:true});
document.addEventListener('touchend',   e => {
  const dx=tx-e.changedTouches[0].clientX, dy=Math.abs(ty-e.changedTouches[0].clientY);
  if (Math.abs(dx)>50 && dy<60) dx>0 ? goTo(current+1) : goTo(current-1);
}, {passive:true});

function toggleFs() {
  if (!document.fullscreenElement) document.documentElement.requestFullscreen().catch(()=>{});
  else document.exitFullscreen().catch(()=>{});
}
document.getElementById('btn-fs').addEventListener('click', toggleFs);

// Init
(function(){
  document.querySelectorAll('.slide').forEach(s => { s.style.display='none'; s.classList.remove('is-active'); });
  const first = document.getElementById('slide-1');
  if (first) { first.style.display='flex'; first.classList.add('is-active'); }
  updateUI();

  // Spoiler: click to reveal
  document.querySelectorAll('.spoiler').forEach(el => {
    el.addEventListener('click', () => el.classList.toggle('revealed'));
  });

  // Screenshot placeholders: show image if it loaded successfully
  document.querySelectorAll('.screenshot-placeholder img').forEach(img => {
    const activate = () => img.closest('.screenshot-placeholder').classList.add('has-image');
    if (img.complete && img.naturalWidth > 0) {
      activate();
    } else {
      img.addEventListener('load', activate);
    }
  });

  // Aufgabe 1: click explanation row → number appears inside [ ]
  document.querySelectorAll('.zu-item--leer[data-answer]').forEach(el => {
    el.style.cursor = 'pointer';
    const box = el.querySelector('.zu-box');
    const answer = el.dataset.answer;
    el.addEventListener('click', () => {
      if (box.textContent === '[ ]') {
        box.textContent = '[' + answer + ']';
        box.style.color = '#2E8B57';
        box.style.fontWeight = '800';
        el.classList.add('correct');
      } else {
        box.textContent = '[ ]';
        box.style.color = '';
        box.style.fontWeight = '';
        el.classList.remove('correct');
      }
    });
  });
})();
