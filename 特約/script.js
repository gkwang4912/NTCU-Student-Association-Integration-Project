// 簡易錨點定位高亮
const links = document.querySelectorAll('.nav-link');
const setActive = () => {
  const hash = location.hash || '#home';
  links.forEach(a => a.classList.toggle('active', a.getAttribute('href')===hash));
};
window.addEventListener('hashchange', setActive);
setActive();