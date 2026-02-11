/**
 * 學生會特約店家 - 核心邏輯
 */

// --- 1. 定義與設定 ---
const SHEET_ID = '1NRrs9PXp_1XfSAx3bDxnpII0X_YMWHWBLEl5Kn_qTuM';
const JSONP_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=responseHandler:handleSheetData`;

const HERO_DATA = [
    {
        image: 'https://images.unsplash.com/photo-1541339907198-e08756ebafe3?q=80&w=1600',
        title: '專屬學生的優惠特約',
        sub: '探索校園周邊精選合作店家，享受學生專屬折扣與好康'
    },
    {
        image: 'https://images.unsplash.com/photo-1523050853064-8504f2f3905d?q=80&w=1600',
        title: '吃喝玩樂 一網打盡',
        sub: '從美食餐飲到休閒娛樂，學生會為您爭取最優質的店家合作'
    },
    {
        image: 'https://images.unsplash.com/photo-1525921429624-479b6a29d81d?q=80&w=1600',
        title: '校園合作 永續經營',
        sub: '建立在地商家與學生的良性連結，共創美好的校園生活圈'
    }
];

let allPartners = [];
let currentPage = 0;
const ITEMS_PER_PAGE = 6;
let currentHeroSlide = 0;

// --- 2. 初始化功能 ---

function init() {
    setupHeroCarousel();
    fetchSheetData();
    setupEventListeners();
}

// --- 3. Hero 輪播邏輯 ---

function setupHeroCarousel() {
    const container = document.getElementById('hero-slides-container');
    const dotsContainer = document.getElementById('hero-dots');

    HERO_DATA.forEach((data, index) => {
        // 建立 Slide
        const slide = document.createElement('div');
        slide.className = `hero-slide ${index === 0 ? 'active' : ''}`;
        slide.style.backgroundImage = `url('${data.image}')`;
        container.appendChild(slide);

        // 建立 Dot
        const dot = document.createElement('div');
        dot.className = `hero-dot ${index === 0 ? 'active' : ''}`;
        dot.onclick = () => goToHeroSlide(index);
        dotsContainer.appendChild(dot);
    });

    // 自動輪播
    setInterval(() => {
        nextHeroSlide();
    }, 5000);
}

function updateHeroDisplay() {
    const slides = document.querySelectorAll('.hero-slide');
    const dots = document.querySelectorAll('.hero-dot');
    const title = document.getElementById('hero-title');
    const sub = document.getElementById('hero-sub');

    slides.forEach((s, i) => s.classList.toggle('active', i === currentHeroSlide));
    dots.forEach((d, i) => d.classList.toggle('active', i === currentHeroSlide));

    title.textContent = HERO_DATA[currentHeroSlide].title;
    sub.textContent = HERO_DATA[currentHeroSlide].sub;
}

function nextHeroSlide() {
    currentHeroSlide = (currentHeroSlide + 1) % HERO_DATA.length;
    updateHeroDisplay();
}

function prevHeroSlide() {
    currentHeroSlide = (currentHeroSlide - 1 + HERO_DATA.length) % HERO_DATA.length;
    updateHeroDisplay();
}

function goToHeroSlide(index) {
    currentHeroSlide = index;
    updateHeroDisplay();
}

// --- 4. 資料抓取與處理 ---

function fetchSheetData() {
    const script = document.createElement('script');
    script.src = JSONP_URL;
    script.onerror = () => {
        const loading = document.getElementById('loading');
        if (loading) loading.textContent = '連線失敗，請檢查網路狀態。';
    };
    document.body.appendChild(script);
}

// 全域 JSONP 回呼
window.handleSheetData = function (data) {
    if (!data || !data.table || !data.table.rows) return;

    const rows = data.table.rows;
    allPartners = rows.map(row => {
        const cells = row.c;
        const getText = (idx) => (cells[idx] ? (cells[idx].f || cells[idx].v || '') : '');
        return {
            name: getText(0),
            deal: getText(1),
            address: getText(2),
            contact: getText(3),
            category: getText(4) || '精選商家'
        };
    }).filter(p => p.name);

    renderCurrentPage();
    renderLatestPartners();
};

// --- 5. 渲染店家功能 ---

function renderCurrentPage() {
    const grid = document.getElementById('partners-grid');
    grid.innerHTML = '';

    const start = currentPage * ITEMS_PER_PAGE;
    const end = start + ITEMS_PER_PAGE;
    const pageItems = allPartners.slice(start, end);

    pageItems.forEach(partner => {
        const card = document.createElement('div');
        card.className = 'partner-card';
        card.setAttribute('onclick', `updateMap('${partner.address}')`);
        card.innerHTML = `
      <div class="ticket-header">
        <div class="ticket-deal">${partner.deal.includes('折') ? partner.deal : '優惠'}</div>
        <div class="ticket-deal-sub">${partner.deal.includes('折') ? '憑證享折扣' : partner.deal}</div>
      </div>
      <div class="ticket-divider"></div>
      <div class="ticket-body">
        <div class="info-top">
          <div class="partner-name">${partner.name}</div>
          <div class="partner-cat">${partner.category}</div>
        </div>
        <div class="partner-loc">📍 ${partner.address.substring(0, 8)}...</div>
        <div class="ticket-hint">點擊查看地圖 〉</div>
      </div>
    `;
        grid.appendChild(card);
    });

    // 更新按鈕狀態
    document.getElementById('page-prev').style.opacity = currentPage === 0 ? '0.3' : '1';
    document.getElementById('page-next').style.opacity = end >= allPartners.length ? '0.3' : '1';
}

function renderLatestPartners() {
    const list = document.getElementById('latest-list');
    list.innerHTML = '';

    const latest = allPartners.slice(-3).reverse();
    latest.forEach(p => {
        const li = document.createElement('li');
        li.className = 'latest-item';
        li.innerHTML = `
      <div class="item-badge">NEW</div>
      <div class="item-content">${p.name} - ${p.deal}</div>
    `;
        list.appendChild(li);
    });
}

// --- 6. 互動與導覽 ---

function setupEventListeners() {
    // Hero 控制
    document.getElementById('hero-prev').onclick = prevHeroSlide;
    document.getElementById('hero-next').onclick = nextHeroSlide;

    // 分頁控制
    document.getElementById('page-prev').onclick = () => {
        if (currentPage > 0) {
            currentPage--;
            renderCurrentPage();
        }
    };
    document.getElementById('page-next').onclick = () => {
        if ((currentPage + 1) * ITEMS_PER_PAGE < allPartners.length) {
            currentPage++;
            renderCurrentPage();
        }
    };

    // 錨點平滑捲動
    document.querySelectorAll('.nav-link').forEach(link => {
        link.onclick = (e) => {
            const href = link.getAttribute('href');
            if (href.startsWith('#')) {
                e.preventDefault();
                const target = document.querySelector(href);
                if (target) {
                    target.scrollIntoView({ behavior: 'smooth' });
                    // 更新主動狀態
                    document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
                    link.classList.add('active');
                }
            }
        };
    });
}

function updateMap(address) {
    const mapFrame = document.getElementById('google-map');
    const query = encodeURIComponent(address);
    mapFrame.src = `https://maps.google.com/maps?q=${query}&t=&z=15&ie=UTF8&iwloc=&output=embed`;

    // 捲動到地圖
    mapFrame.parentElement.scrollIntoView({ behavior: 'smooth' });
}

// --- 啟動 ---
init();
