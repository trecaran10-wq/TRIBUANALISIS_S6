document.addEventListener("DOMContentLoaded", () => {
  const sidebar = document.getElementById("sidebar");
  document.querySelectorAll("[data-menu-toggle]").forEach(btn => btn.addEventListener("click", () => sidebar?.classList.toggle("open")));
  document.querySelectorAll("form[data-confirm]").forEach(form => form.addEventListener("submit", e => {
    if (!window.confirm(form.dataset.confirm || "¿Continuar?")) e.preventDefault();
  }));
  const select = document.getElementById("category-select");
  document.querySelectorAll("[data-category-choice]").forEach(card => {
    card.addEventListener("click", () => {
      if (select) { select.value = card.dataset.categoryChoice; select.dispatchEvent(new Event('change')); }
      document.querySelectorAll("[data-category-choice]").forEach(c => c.classList.remove("chosen"));
      card.classList.add("chosen");
    });
  });
  if (select && new URLSearchParams(location.search).get("category")) {
    const desired = new URLSearchParams(location.search).get("category");
    if ([...select.options].some(o => o.value === desired)) { select.value = desired; select.dispatchEvent(new Event('change')); }
    document.getElementById("formulario")?.scrollIntoView({behavior:"smooth", block:"start"});
  }
});

// Navegación interactiva de los módulos didácticos de Semana 6
(function(){
  const tabs = [...document.querySelectorAll('[data-lesson-tab]')];
  const panels = [...document.querySelectorAll('[data-lesson-panel]')];
  function activateLesson(name){
    tabs.forEach(tab=>{const active=tab.dataset.lessonTab===name;tab.classList.toggle('active',active);tab.setAttribute('aria-selected',active?'true':'false');});
    panels.forEach(panel=>{const active=panel.dataset.lessonPanel===name;panel.hidden=!active;panel.classList.toggle('active',active);});
    const progress=document.querySelector('.learning-progress i');
    const label=document.querySelector('.learning-progress b');
    const index=['cuarta','quinta','digital','campo'].indexOf(name);
    if(progress) progress.style.width=((index+1)*25)+'%';
    if(label) label.textContent=(index+1)+' de 4 módulos';
  }
  tabs.forEach(tab=>tab.addEventListener('click',()=>activateLesson(tab.dataset.lessonTab)));
})();

// Ajusta la ficha de entrada a la categoría seleccionada sin aplicar tasas genéricas.
(function(){
  const select = document.getElementById('category-select');
  if (!select) return;
  const labels = {
    income: document.getElementById('main-income-label'),
    frequency: document.getElementById('income-frequency-label'),
    otherFourth: document.getElementById('other-fourth-label'),
    otherFifth: document.getElementById('other-fifth-label'),
    payments: document.getElementById('payments-label'),
    paymentFrequency: document.getElementById('payments-frequency-label'),
    additional: document.getElementById('additional-deductions-label'),
    itf: document.getElementById('itf-label'),
    donations: document.getElementById('donations-label')
  };
  function syncFields(){
    const category = select.value;
    if(labels.income){
      const text = category==='cuarta' ? 'Ingresos brutos de cuarta categoría (S/) *' : category==='quinta' ? 'Remuneración bruta de quinta categoría (S/) *' : 'Ingresos digitales del caso (S/) *';
      labels.income.childNodes[0].textContent = text;
    }
    if(labels.otherFourth) labels.otherFourth.hidden = category==='cuarta' || category==='creador' || category==='plataforma';
    if(labels.otherFifth) labels.otherFifth.hidden = category==='quinta' || category==='creador' || category==='plataforma';
    const workIncome = category==='cuarta' || category==='quinta';
    [labels.payments,labels.paymentFrequency,labels.additional,labels.itf,labels.donations].forEach(el=>{if(el)el.hidden=!workIncome;});
    if(labels.frequency){
      const small = labels.frequency.querySelector('select');
      if(small) small.title = category==='quinta' ? 'Para una proyección completa, agrega gratificaciones y otros conceptos anuales gravados en otros ingresos de quinta.' : 'Los importes mensuales se multiplican por 12.';
    }
  }
  select.addEventListener('change', syncFields);
  syncFields();
})();


// Formularios por categoría: permite combinar cuarta y quinta en una misma ficha.
document.addEventListener('DOMContentLoaded', () => {
  const toggle = document.getElementById('has-other-category');
  const extra = document.getElementById('other-income-wrap');
  if (toggle && extra) {
    const syncExtra = () => { extra.hidden = !toggle.checked; };
    toggle.addEventListener('change', syncExtra);
    syncExtra();
  }
});

// Videos específicos para el módulo didáctico activo; no se repiten videos de otras categorías.
(function(){
  const tabs = [...document.querySelectorAll('[data-lesson-tab]')];
  const cards = [...document.querySelectorAll('[data-video-topic]')];
  const section = document.getElementById('topic-videos');
  function filterVideos(topic){
    cards.forEach(card => { card.hidden = card.dataset.videoTopic !== topic; });
    if(section) section.hidden = topic === 'campo';
  }
  tabs.forEach(tab => tab.addEventListener('click', () => filterVideos(tab.dataset.lessonTab)));
  filterVideos(document.querySelector('[data-lesson-tab].active')?.dataset.lessonTab || 'cuarta');
})();
