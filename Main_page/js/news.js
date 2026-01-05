/**
 * News System Controller
 * Fetches data from Google Sheets via Visualization API (JSONP)
 */

const SHEET_ID = '1uXsSujvzidIDc1qAeO-DCh4cUJpFJi71Ep-hZxRJ5Io';
// Query: Select A, B, C (Date, Title, Link). 
const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&tq=SELECT A, B, C`;

// Global store for news data
let textNewsCache = [];

/**
 * Parses the Google Viz API JSON response
 * The response comes wrapped in `google.visualization.Query.setResponse(...)` if using JSONP,
 * but here we fetch raw text and manually strip the wrapper to avoid loading the heavy Viz library.
 */
/**
 * Parses the Google Viz API JSON response using JSONP to bypass CORS.
 */
function fetchNewsData(callback) {
    // Define a unique global callback name
    const callbackName = 'googleVizCallback_' + Math.floor(Math.random() * 100000);

    // Create the script URL with the callback
    const scriptUrl = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=responseHandler:${callbackName}&tq=SELECT A, B, C`;

    // Define the global callback
    window[callbackName] = function (json) {
        // Cleanup
        delete window[callbackName];
        document.body.removeChild(scriptElement);

        try {
            // Extract rows
            const rows = json.table.rows.map(row => {
                // Formatting date/values safely
                const date = row.c[0] ? (row.c[0].f || row.c[0].v) : '';
                const title = row.c[1] ? (row.c[1].v) : '';
                const link = row.c[2] ? (row.c[2].v) : '#';
                return { date, title, link };
            });

            textNewsCache = rows;
            if (callback) callback(rows);
        } catch (err) {
            console.error('Error parsing news data:', err);
            showError();
        }
    };

    // Create and inject script tag
    const scriptElement = document.createElement('script');
    scriptElement.src = scriptUrl;
    scriptElement.onerror = function () {
        console.error('Error loading news script via JSONP');
        showError();
        delete window[callbackName];
        if (scriptElement.parentNode) {
            document.body.removeChild(scriptElement);
        }
    };
    document.body.appendChild(scriptElement);
}

function showError() {
    const container = document.getElementById('news-loading-state');
    if (container) container.innerHTML = '<p class="error-message">無法載入公告資料</p>';
}

/**
 * Render logic for Index Page (Latest 3)
 */
function renderIndexNews() {
    const container = document.getElementById('latest-news-container');
    if (!container) return;

    fetchNewsData((data) => {
        container.innerHTML = ''; // Clear loading state

        // Take latest 3 (assuming sheet is appended chronologically, we reverse to get latest first? 
        // Or if user enters latest at bottom. Usually sheets are top-down. 
        // Let's assume user adds new rows at bottom, so we reverse it.
        const latestNews = [...data].reverse().slice(0, 3);

        // Color classes cycler
        const colors = ['item-yellow', 'item-green', 'item-blue'];

        latestNews.forEach((item, index) => {
            const colorClass = colors[index % colors.length];
            const safeTitle = escapeHtml(item.title);
            const safeDate = escapeHtml(item.date);
            const safeLink = encodeURIComponent(item.link);
            const detailUrl = `news_detail.html?title=${encodeURIComponent(item.title)}&date=${encodeURIComponent(item.date)}&doc=${safeLink}`;

            const html = `
                <a href="${detailUrl}" class="news-item-link">
                    <div class="news-item ${colorClass}">
                        <h4>${safeTitle}</h4>
                        <span class="date">${safeDate}</span>
                    </div>
                </a>
            `;
            container.innerHTML += html;
        });
    });
}

/**
 * Render logic for All News Page
 */
function renderAllNews() {
    const container = document.getElementById('all-news-container');
    if (!container) return;

    fetchNewsData((data) => {
        container.innerHTML = '';
        const allNews = [...data].reverse();

        allNews.forEach((item) => {
            const safeTitle = escapeHtml(item.title);
            const safeDate = escapeHtml(item.date);
            const safeLink = encodeURIComponent(item.link);
            const detailUrl = `news_detail.html?title=${encodeURIComponent(item.title)}&date=${encodeURIComponent(item.date)}&doc=${safeLink}`;

            const html = `
                <a href="${detailUrl}" class="all-news-item">
                    <div class="news-date">${safeDate}</div>
                    <div class="news-title">${safeTitle}</div>
                    <div class="news-arrow"><i class="fas fa-chevron-right"></i></div>
                </a>
            `;
            container.innerHTML += html;
        });
    });
}

/**
 * Render logic for Detail Page
 */
function renderNewsDetail() {
    const params = new URLSearchParams(window.location.search);
    const title = params.get('title');
    const date = params.get('date');
    const docUrl = params.get('doc');

    if (!title || !docUrl) {
        document.body.innerHTML = '<div class="container" style="padding-top:100px;"><h1>無效的公告連結</h1><a href="index.html">返回首頁</a></div>';
        return;
    }

    document.getElementById('detail-title').textContent = title;
    document.getElementById('detail-date').textContent = date;

    // Convert standard Edit link to Preview/Embed link
    // Input: https://docs.google.com/document/d/DOC_ID/edit...
    // Output: https://docs.google.com/document/d/DOC_ID/preview
    let embedUrl = docUrl;
    const docIdMatch = docUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (docIdMatch && docIdMatch[1]) {
        embedUrl = `https://docs.google.com/document/d/${docIdMatch[1]}/preview?embedded=true`;
    }

    document.getElementById('doc-frame').src = embedUrl;
}

// Utility
function escapeHtml(text) {
    if (!text) return '';
    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

// Initializer
document.addEventListener('DOMContentLoaded', () => {
    if (document.getElementById('latest-news-container')) {
        renderIndexNews();
    }
    if (document.getElementById('all-news-container')) {
        renderAllNews();
    }
    if (document.getElementById('detail-title')) {
        renderNewsDetail();
    }
});
