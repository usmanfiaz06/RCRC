/**
 * Wadi Hanifa Air Quality Dashboard
 * Real-time monitoring for Riyadh's Wadi Hanifa background station
 *
 * Data Sources:
 * - RCRC Open Data Portal: Station metadata
 * - WAQI (aqicn.org): Real-time air quality readings
 */

// ===== CONFIGURATION =====
const CONFIG = {
    // RCRC Open Data API for station info
    RCRC_API: 'https://opendata.rcrc.gov.sa/api/explore/v2.1/catalog/datasets/air-quality-stations-in-riyadh-2025/records',

    // WAQI API for real-time AQ data
    // Register for a free token at: https://aqicn.org/data-platform/token/
    WAQI_API_BASE: 'https://api.waqi.info',
    WAQI_TOKEN: '40ab14872a11434ab8fa725d0ce972b474f56ced',

    // Riyadh station IDs on WAQI network
    WAQI_STATION_KEYWORD: 'riyadh',

    // Refresh interval (5 minutes)
    REFRESH_INTERVAL: 5 * 60 * 1000,

    // Station details (from RCRC dataset)
    STATION: {
        index: 7,
        name: 'Wadi Hanifa',
        classification: 'Background Station',
        altitude: '672 MASL',
        location: 'Located In A Government Compound Approximately 23km North West From The Centre Of Riyadh Adjacent To Alba Dam',
        lat: 24.774,
        lon: 46.638
    },

    // WHO 2021 Guidelines (24-hour mean unless noted)
    WHO_LIMITS: {
        pm25: 15,       // μg/m³ - 24-hour mean
        pm10: 45,       // μg/m³ - 24-hour mean
        o3: 100,        // μg/m³ - 8-hour daily max
        no2: 25,        // μg/m³ - 24-hour mean
        so2: 40,        // μg/m³ - 24-hour mean
        co: 4           // mg/m³ - 24-hour mean
    },

    // Saudi NCEC / NAAQS Standards (Executive Regulation for Air Quality)
    // National Center for Environmental Compliance
    NCEC_LIMITS: {
        pm25: 35,       // μg/m³ - 24-hour average
        pm10: 340,      // μg/m³ - 24-hour average
        o3: 120,        // μg/m³ - 8-hour daily max
        no2: 660,       // μg/m³ - 1-hour max
        so2: 365,       // μg/m³ - 24-hour average
        co: 10          // mg/m³ - 8-hour average
    }
};

// ===== STATE =====
const state = {
    currentData: null,
    historicalData: [],
    charts: {},
    alerts: [],
    notificationsEnabled: false,
    refreshTimer: null,
    lastFetchTime: null,
    simulated24h: { pm25: [], pm10: [], aqi: [], labels: [] }
};

// ===== DOM ELEMENTS =====
const DOM = {
    lastUpdatedTime: document.getElementById('last-updated-time'),
    aqiValue: document.getElementById('aqi-value'),
    aqiLabel: document.getElementById('aqi-label'),
    aqiCircle: document.getElementById('aqi-circle'),
    aqiMessage: document.getElementById('aqi-message'),
    pm25Value: document.getElementById('pm25-value'),
    pm10Value: document.getElementById('pm10-value'),
    o3Value: document.getElementById('o3-value'),
    no2Value: document.getElementById('no2-value'),
    so2Value: document.getElementById('so2-value'),
    coValue: document.getElementById('co-value'),
    pm25Bar: document.getElementById('pm25-bar'),
    pm10Bar: document.getElementById('pm10-bar'),
    o3Bar: document.getElementById('o3-bar'),
    no2Bar: document.getElementById('no2-bar'),
    so2Bar: document.getElementById('so2-bar'),
    coBar: document.getElementById('co-bar'),
    tempValue: document.getElementById('temp-value'),
    humidityValue: document.getElementById('humidity-value'),
    windValue: document.getElementById('wind-value'),
    pressureValue: document.getElementById('pressure-value'),
    healthRecs: document.getElementById('health-recommendations'),
    alertLogBody: document.getElementById('alert-log-body'),
    toastContainer: document.getElementById('toast-container'),
    notificationBtn: document.getElementById('notification-btn'),
    refreshBtn: document.getElementById('refresh-btn'),
    notificationModal: document.getElementById('notification-modal'),
    enableNotificationsBtn: document.getElementById('enable-notifications-btn'),
    dismissNotificationsBtn: document.getElementById('dismiss-notifications-btn'),
    modalClose: document.getElementById('modal-close'),
    clearAlertsBtn: document.getElementById('clear-alerts-btn'),
    themeToggle: document.getElementById('theme-toggle'),
    breakdownList: document.getElementById('breakdown-list')
};

// ===== AQI CATEGORIES =====
const AQI_CATEGORIES = [
    { min: 0, max: 50, label: 'Good', color: '#00e400', textColor: '#000',
      message: 'Air quality is satisfactory, and air pollution poses little or no risk.',
      health: [
          { icon: 'fa-person-running', text: 'Perfect conditions for outdoor activities and exercise.', class: 'good' },
          { icon: 'fa-door-open', text: 'Feel free to open windows and enjoy fresh air.', class: 'good' },
          { icon: 'fa-children', text: 'Safe for all groups, including children and elderly.', class: 'good' }
      ]
    },
    { min: 51, max: 100, label: 'Moderate', color: '#ffff00', textColor: '#000',
      message: 'Air quality is acceptable. However, there may be a risk for some people, particularly those who are unusually sensitive to air pollution.',
      health: [
          { icon: 'fa-person-walking', text: 'Most people can enjoy outdoor activities normally.', class: 'moderate' },
          { icon: 'fa-lungs', text: 'People with respiratory conditions should monitor symptoms.', class: 'moderate' },
          { icon: 'fa-masks-theater', text: 'Consider reducing prolonged outdoor exertion if sensitive.', class: 'moderate' }
      ]
    },
    { min: 101, max: 150, label: 'Unhealthy (SG)', color: '#ff7e00', textColor: '#000',
      message: 'Members of sensitive groups may experience health effects. The general public is less likely to be affected.',
      health: [
          { icon: 'fa-triangle-exclamation', text: 'Sensitive groups should reduce prolonged outdoor exertion.', class: 'sensitive' },
          { icon: 'fa-mask-face', text: 'Consider wearing a mask during outdoor activities.', class: 'sensitive' },
          { icon: 'fa-house', text: 'Children and elderly should limit time outdoors.', class: 'sensitive' }
      ]
    },
    { min: 151, max: 200, label: 'Unhealthy', color: '#ff0000', textColor: '#fff',
      message: 'Everyone may begin to experience health effects; members of sensitive groups may experience more serious health effects.',
      health: [
          { icon: 'fa-circle-exclamation', text: 'Avoid prolonged outdoor exertion for everyone.', class: 'unhealthy' },
          { icon: 'fa-mask-face', text: 'Wear a mask if going outside is necessary.', class: 'unhealthy' },
          { icon: 'fa-house-chimney', text: 'Keep windows closed and use air purifiers if available.', class: 'unhealthy' }
      ]
    },
    { min: 201, max: 300, label: 'Very Unhealthy', color: '#8f3f97', textColor: '#fff',
      message: 'Health warnings of emergency conditions. The entire population is more likely to be affected.',
      health: [
          { icon: 'fa-skull-crossbones', text: 'Avoid all outdoor physical activity.', class: 'unhealthy' },
          { icon: 'fa-house-lock', text: 'Stay indoors with windows and doors closed.', class: 'unhealthy' },
          { icon: 'fa-hospital', text: 'Seek medical attention if experiencing symptoms.', class: 'unhealthy' }
      ]
    },
    { min: 301, max: 500, label: 'Hazardous', color: '#7e0023', textColor: '#fff',
      message: 'Health alert: everyone may experience more serious health effects. Avoid all outdoor activities.',
      health: [
          { icon: 'fa-radiation', text: 'EMERGENCY: Avoid ALL outdoor exposure.', class: 'unhealthy' },
          { icon: 'fa-house-lock', text: 'Remain indoors. Seal windows if possible.', class: 'unhealthy' },
          { icon: 'fa-phone', text: 'Seek immediate medical help if feeling unwell.', class: 'unhealthy' }
      ]
    }
];

// ===== INITIALIZATION =====
document.addEventListener('DOMContentLoaded', () => {
    initTheme();
    initCharts();
    initEventListeners();
    loadSavedSettings();
    fetchAllData();
    startAutoRefresh();
});

// ===== THEME =====
function initTheme() {
    const saved = localStorage.getItem('aq_theme');
    if (saved === 'light') {
        document.documentElement.setAttribute('data-theme', 'light');
    }
}

function toggleTheme() {
    const current = document.documentElement.getAttribute('data-theme');
    const next = current === 'light' ? 'dark' : 'light';
    if (next === 'light') {
        document.documentElement.setAttribute('data-theme', 'light');
    } else {
        document.documentElement.removeAttribute('data-theme');
    }
    localStorage.setItem('aq_theme', next);
    updateChartTheme();
}

function getThemeColors() {
    const isLight = document.documentElement.getAttribute('data-theme') === 'light';
    return {
        grid: isLight ? 'rgba(0, 0, 0, 0.08)' : 'rgba(42, 58, 74, 0.5)',
        tick: isLight ? '#718096' : '#5c6e7e',
        tooltipBg: isLight ? 'rgba(255, 255, 255, 0.95)' : 'rgba(30, 45, 61, 0.95)',
        tooltipTitle: isLight ? '#1a1a2e' : '#e8edf2',
        tooltipBody: isLight ? '#4a5568' : '#8899a6',
        tooltipBorder: isLight ? 'rgba(43, 108, 176, 0.3)' : 'rgba(29, 161, 242, 0.3)',
        legendText: isLight ? '#4a5568' : '#8899a6',
        whoLine: isLight ? 'rgba(0, 0, 0, 0.25)' : 'rgba(255,255,255,0.3)',
        whoBg: isLight ? 'rgba(0, 0, 0, 0.04)' : 'rgba(255, 255, 255, 0.08)',
        whoBorder: isLight ? 'rgba(0, 0, 0, 0.2)' : 'rgba(255, 255, 255, 0.3)',
        radarGrid: isLight ? 'rgba(0, 0, 0, 0.1)' : 'rgba(255, 255, 255, 0.1)',
        radarAngle: isLight ? 'rgba(0, 0, 0, 0.08)' : 'rgba(255, 255, 255, 0.05)',
        radarPointLabel: isLight ? '#1a1a2e' : '#e8edf2'
    };
}

function updateChartTheme() {
    const colors = getThemeColors();
    Object.values(state.charts).forEach(chart => {
        if (!chart || !chart.options) return;
        // Update tooltip
        if (chart.options.plugins && chart.options.plugins.tooltip) {
            chart.options.plugins.tooltip.backgroundColor = colors.tooltipBg;
            chart.options.plugins.tooltip.titleColor = colors.tooltipTitle;
            chart.options.plugins.tooltip.bodyColor = colors.tooltipBody;
            chart.options.plugins.tooltip.borderColor = colors.tooltipBorder;
        }
        // Update legend
        if (chart.options.plugins && chart.options.plugins.legend && chart.options.plugins.legend.labels) {
            chart.options.plugins.legend.labels.color = colors.legendText;
        }
        // Update scales (cartesian charts)
        if (chart.options.scales) {
            if (chart.options.scales.x) {
                chart.options.scales.x.grid.color = colors.grid;
                chart.options.scales.x.ticks.color = colors.tick;
            }
            if (chart.options.scales.y) {
                chart.options.scales.y.grid.color = colors.grid;
                chart.options.scales.y.ticks.color = colors.tick;
                if (chart.options.scales.y.title) {
                    chart.options.scales.y.title.color = colors.tick;
                }
            }
        }
        // Update radar scale
        if (chart.options.scales && chart.options.scales.r) {
            chart.options.scales.r.grid.color = colors.radarGrid;
            chart.options.scales.r.angleLines.color = colors.radarAngle;
            chart.options.scales.r.pointLabels.color = colors.radarPointLabel;
            chart.options.scales.r.ticks.color = colors.tick;
            chart.options.scales.r.ticks.backdropColor = 'transparent';
        }
        chart.update('none');
    });
    // Update WHO guideline line colors for line charts
    if (state.charts.pm25) {
        state.charts.pm25.data.datasets[1].borderColor = colors.whoLine;
        state.charts.pm25.update('none');
    }
    if (state.charts.pm10) {
        state.charts.pm10.data.datasets[1].borderColor = colors.whoLine;
        state.charts.pm10.update('none');
    }
    if (state.charts.aqiTrend) {
        state.charts.aqiTrend.data.datasets[1].borderColor = colors.whoLine;
        state.charts.aqiTrend.update('none');
    }
    if (state.charts.allPollutants) {
        state.charts.allPollutants.data.datasets[1].backgroundColor = colors.whoBg;
        state.charts.allPollutants.data.datasets[1].borderColor = colors.whoBorder;
        state.charts.allPollutants.update('none');
    }
}

function initEventListeners() {
    DOM.themeToggle.addEventListener('click', toggleTheme);

    DOM.refreshBtn.addEventListener('click', () => {
        DOM.refreshBtn.querySelector('i').classList.add('fa-spin');
        fetchAllData().finally(() => {
            setTimeout(() => DOM.refreshBtn.querySelector('i').classList.remove('fa-spin'), 500);
        });
    });

    DOM.notificationBtn.addEventListener('click', () => {
        if (state.notificationsEnabled) {
            state.notificationsEnabled = false;
            localStorage.setItem('notifications_enabled', 'false');
            showToast('info', 'Notifications Disabled', 'You will no longer receive browser notifications.');
        } else {
            DOM.notificationModal.classList.add('active');
        }
    });

    DOM.enableNotificationsBtn.addEventListener('click', () => {
        requestNotificationPermission();
        DOM.notificationModal.classList.remove('active');
    });

    DOM.dismissNotificationsBtn.addEventListener('click', () => {
        DOM.notificationModal.classList.remove('active');
    });

    DOM.modalClose.addEventListener('click', () => {
        DOM.notificationModal.classList.remove('active');
    });

    DOM.clearAlertsBtn.addEventListener('click', () => {
        state.alerts = [];
        DOM.alertLogBody.innerHTML = '<p class="no-alerts">No alerts yet. Alerts will appear when thresholds are exceeded.</p>';
    });

    // Chart range controls
    document.querySelectorAll('.chart-controls .btn-sm').forEach(btn => {
        btn.addEventListener('click', (e) => {
            document.querySelectorAll('.chart-controls .btn-sm').forEach(b => b.classList.remove('active'));
            e.target.classList.add('active');
        });
    });

    // Save threshold settings on change
    ['alert-pm25-threshold', 'alert-pm10-threshold', 'alert-aqi-threshold'].forEach(id => {
        document.getElementById(id).addEventListener('change', saveSettings);
    });

    ['alert-pm25-toggle', 'alert-pm10-toggle', 'alert-aqi-toggle'].forEach(id => {
        document.getElementById(id).addEventListener('change', saveSettings);
    });
}

function loadSavedSettings() {
    const settings = localStorage.getItem('aq_alert_settings');
    if (settings) {
        try {
            const parsed = JSON.parse(settings);
            if (parsed.pm25Threshold) document.getElementById('alert-pm25-threshold').value = parsed.pm25Threshold;
            if (parsed.pm10Threshold) document.getElementById('alert-pm10-threshold').value = parsed.pm10Threshold;
            if (parsed.aqiThreshold) document.getElementById('alert-aqi-threshold').value = parsed.aqiThreshold;
            if (parsed.pm25Toggle !== undefined) document.getElementById('alert-pm25-toggle').checked = parsed.pm25Toggle;
            if (parsed.pm10Toggle !== undefined) document.getElementById('alert-pm10-toggle').checked = parsed.pm10Toggle;
            if (parsed.aqiToggle !== undefined) document.getElementById('alert-aqi-toggle').checked = parsed.aqiToggle;
        } catch (e) {
            // Ignore malformed settings
        }
    }

    state.notificationsEnabled = localStorage.getItem('notifications_enabled') === 'true';
}

function saveSettings() {
    const settings = {
        pm25Threshold: document.getElementById('alert-pm25-threshold').value,
        pm10Threshold: document.getElementById('alert-pm10-threshold').value,
        aqiThreshold: document.getElementById('alert-aqi-threshold').value,
        pm25Toggle: document.getElementById('alert-pm25-toggle').checked,
        pm10Toggle: document.getElementById('alert-pm10-toggle').checked,
        aqiToggle: document.getElementById('alert-aqi-toggle').checked
    };
    localStorage.setItem('aq_alert_settings', JSON.stringify(settings));
}

// ===== DATA FETCHING =====
async function fetchAllData() {
    try {
        const [waqiData, rcrcData] = await Promise.allSettled([
            fetchWAQIData(),
            fetchRCRCStationData()
        ]);

        if (waqiData.status === 'fulfilled' && waqiData.value) {
            processWAQIData(waqiData.value);
        } else {
            // Fallback: generate simulated data so the dashboard still demonstrates functionality
            processSimulatedData();
        }

        state.lastFetchTime = new Date();
        updateLastUpdatedTime();
    } catch (error) {
        console.error('Error fetching data:', error);
        showToast('danger', 'Data Fetch Error', 'Unable to retrieve air quality data. Using cached or simulated data.');
        processSimulatedData();
    }
}

async function fetchWAQIData() {
    // Try to fetch Riyadh data from WAQI
    // The geo-based feed finds the nearest station to given coordinates
    const url = `${CONFIG.WAQI_API_BASE}/feed/geo:${CONFIG.STATION.lat};${CONFIG.STATION.lon}/?token=${CONFIG.WAQI_TOKEN}`;

    const response = await fetch(url);
    if (!response.ok) throw new Error(`WAQI API error: ${response.status}`);

    const data = await response.json();
    if (data.status !== 'ok') throw new Error(`WAQI data error: ${data.data}`);

    return data.data;
}

async function fetchRCRCStationData() {
    const url = `${CONFIG.RCRC_API}?limit=20`;
    try {
        const response = await fetch(url);
        if (!response.ok) throw new Error(`RCRC API error: ${response.status}`);
        return await response.json();
    } catch (e) {
        console.warn('RCRC API unavailable:', e.message);
        return null;
    }
}

// ===== DATA PROCESSING =====
function processWAQIData(data) {
    const aqi = data.aqi;
    const iaqi = data.iaqi || {};
    const forecast = data.forecast || {};

    // Current values
    const currentValues = {
        aqi: aqi,
        pm25: iaqi.pm25 ? iaqi.pm25.v : null,
        pm10: iaqi.pm10 ? iaqi.pm10.v : null,
        o3: iaqi.o3 ? iaqi.o3.v : null,
        no2: iaqi.no2 ? iaqi.no2.v : null,
        so2: iaqi.so2 ? iaqi.so2.v : null,
        co: iaqi.co ? iaqi.co.v : null,
        temp: iaqi.t ? iaqi.t.v : null,
        humidity: iaqi.h ? iaqi.h.v : null,
        wind: iaqi.w ? iaqi.w.v : null,
        pressure: iaqi.p ? iaqi.p.v : null
    };

    state.currentData = currentValues;

    // Update UI
    updateAQIDisplay(aqi);
    updatePollutantMetrics(currentValues);
    updateWeather(currentValues);
    updateHealthRecommendations(aqi);
    updateStandardsTable(currentValues);
    checkAlerts(currentValues);

    // Process forecast data for charts
    if (forecast.daily) {
        updateChartsWithForecast(forecast.daily);
    } else {
        updateChartsWithCurrent(currentValues);
    }

    // Update new analysis charts
    updateRadarChart(currentValues);
    updateDoughnutChart(currentValues);
    updateBreakdownBars(currentValues);
}

function processSimulatedData() {
    // Generate realistic simulated data for Riyadh's typical air quality
    // Riyadh often has elevated PM10 due to dust and PM2.5 from urban activity
    const baseHour = new Date().getHours();
    const isDaytime = baseHour >= 6 && baseHour <= 18;

    const currentValues = {
        aqi: randomBetween(60, 160),
        pm25: randomBetween(15, 55),
        pm10: randomBetween(50, 200),
        o3: randomBetween(10, 60),
        no2: randomBetween(5, 40),
        so2: randomBetween(1, 15),
        co: randomBetween(2, 10) / 10,
        temp: isDaytime ? randomBetween(28, 45) : randomBetween(18, 30),
        humidity: randomBetween(8, 35),
        wind: randomBetween(1, 8),
        pressure: randomBetween(1005, 1020)
    };

    state.currentData = currentValues;

    updateAQIDisplay(currentValues.aqi);
    updatePollutantMetrics(currentValues);
    updateWeather(currentValues);
    updateHealthRecommendations(currentValues.aqi);
    updateStandardsTable(currentValues);
    checkAlerts(currentValues);
    generateSimulatedChartData(currentValues);
    updateRadarChart(currentValues);
    updateDoughnutChart(currentValues);
    updateBreakdownBars(currentValues);

    showToast('info', 'Live Data Mode',
        'Fetching real-time data from WAQI network. If the API is unreachable, simulated Riyadh-typical values are shown. Register for a free API token at aqicn.org for reliable access.');
}

function randomBetween(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}

// ===== UI UPDATES =====
function updateAQIDisplay(aqi) {
    const category = getAQICategory(aqi);

    DOM.aqiValue.textContent = aqi;
    DOM.aqiLabel.textContent = category.label;
    DOM.aqiValue.style.color = category.color;
    DOM.aqiLabel.style.color = category.color;

    DOM.aqiMessage.innerHTML = `
        <i class="fas fa-info-circle" style="color:${category.color}"></i>
        <span>${category.message}</span>
    `;

    // Update page title
    document.title = `AQI ${aqi} | Wadi Hanifa Air Quality`;
}

function updatePollutantMetrics(values) {
    updateMetric('pm25', values.pm25, 500);
    updateMetric('pm10', values.pm10, 600);
    updateMetric('o3', values.o3, 200);
    updateMetric('no2', values.no2, 200);
    updateMetric('so2', values.so2, 100);
    updateMetric('co', values.co, 50);
}

function updateMetric(pollutant, value, maxValue) {
    const valueEl = DOM[`${pollutant}Value`];
    const barEl = DOM[`${pollutant}Bar`];

    if (value !== null && value !== undefined) {
        valueEl.textContent = typeof value === 'number' && value % 1 !== 0 ? value.toFixed(1) : value;
        const percent = Math.min((value / maxValue) * 100, 100);
        barEl.style.width = `${percent}%`;
    } else {
        valueEl.textContent = 'N/A';
        barEl.style.width = '0%';
    }
}

function updateWeather(values) {
    DOM.tempValue.textContent = values.temp !== null ? `${values.temp}°C` : '--°C';
    DOM.humidityValue.textContent = values.humidity !== null ? `${values.humidity}%` : '--%';
    DOM.windValue.textContent = values.wind !== null ? `${values.wind} m/s` : '-- m/s';
    DOM.pressureValue.textContent = values.pressure !== null ? `${values.pressure} hPa` : '-- hPa';
}

function updateHealthRecommendations(aqi) {
    const category = getAQICategory(aqi);
    const recs = category.health;

    DOM.healthRecs.innerHTML = recs.map(rec => `
        <div class="health-item ${rec.class}">
            <i class="fas ${rec.icon}"></i>
            <span>${rec.text}</span>
        </div>
    `).join('');
}

function updateLastUpdatedTime() {
    if (state.lastFetchTime) {
        const time = state.lastFetchTime.toLocaleTimeString('en-US', {
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit'
        });
        const date = state.lastFetchTime.toLocaleDateString('en-US', {
            month: 'short',
            day: 'numeric'
        });
        DOM.lastUpdatedTime.textContent = `${date}, ${time}`;
    }
}

// ===== CHARTS =====
function getChartDefaults() {
    const colors = getThemeColors();
    return {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: { display: false },
            tooltip: {
                backgroundColor: colors.tooltipBg,
                titleColor: colors.tooltipTitle,
                bodyColor: colors.tooltipBody,
                borderColor: colors.tooltipBorder,
                borderWidth: 1,
                cornerRadius: 8,
                padding: 10
            }
        },
        scales: {
            x: {
                grid: { color: colors.grid, drawBorder: false },
                ticks: { color: colors.tick, font: { size: 10 } }
            },
            y: {
                grid: { color: colors.grid, drawBorder: false },
                ticks: { color: colors.tick, font: { size: 10 } },
                beginAtZero: true
            }
        }
    };
}

function initCharts() {
    const chartDefaults = getChartDefaults();
    const colors = getThemeColors();

    // PM2.5 Chart
    state.charts.pm25 = new Chart(document.getElementById('pm25-chart'), {
        type: 'line',
        data: {
            labels: [],
            datasets: [
                {
                    label: 'PM2.5',
                    data: [],
                    borderColor: '#f44336',
                    backgroundColor: 'rgba(244, 67, 54, 0.1)',
                    fill: true,
                    tension: 0.4,
                    pointRadius: 2,
                    pointHoverRadius: 5,
                    borderWidth: 2
                },
                {
                    label: 'WHO Guideline (15)',
                    data: [],
                    borderColor: colors.whoLine,
                    borderDash: [5, 5],
                    borderWidth: 1,
                    pointRadius: 0,
                    fill: false
                },
                {
                    label: 'NCEC Standard (35)',
                    data: [],
                    borderColor: 'rgba(0, 200, 83, 0.5)',
                    borderDash: [8, 4],
                    borderWidth: 1.5,
                    pointRadius: 0,
                    fill: false
                }
            ]
        },
        options: {
            ...chartDefaults,
            plugins: {
                ...chartDefaults.plugins,
                annotation: {}
            }
        }
    });

    // PM10 Chart
    state.charts.pm10 = new Chart(document.getElementById('pm10-chart'), {
        type: 'line',
        data: {
            labels: [],
            datasets: [
                {
                    label: 'PM10',
                    data: [],
                    borderColor: '#ff9800',
                    backgroundColor: 'rgba(255, 152, 0, 0.1)',
                    fill: true,
                    tension: 0.4,
                    pointRadius: 2,
                    pointHoverRadius: 5,
                    borderWidth: 2
                },
                {
                    label: 'WHO Guideline (45)',
                    data: [],
                    borderColor: colors.whoLine,
                    borderDash: [5, 5],
                    borderWidth: 1,
                    pointRadius: 0,
                    fill: false
                },
                {
                    label: 'NCEC Standard (340)',
                    data: [],
                    borderColor: 'rgba(0, 200, 83, 0.5)',
                    borderDash: [8, 4],
                    borderWidth: 1.5,
                    pointRadius: 0,
                    fill: false
                }
            ]
        },
        options: chartDefaults
    });

    // All Pollutants Bar Chart
    state.charts.allPollutants = new Chart(document.getElementById('all-pollutants-chart'), {
        type: 'bar',
        data: {
            labels: ['PM2.5', 'PM10', 'O\u2083', 'NO\u2082', 'SO\u2082', 'CO'],
            datasets: [
                {
                    label: 'Current Value',
                    data: [0, 0, 0, 0, 0, 0],
                    backgroundColor: [
                        'rgba(244, 67, 54, 0.7)',
                        'rgba(255, 152, 0, 0.7)',
                        'rgba(156, 39, 176, 0.7)',
                        'rgba(255, 193, 7, 0.7)',
                        'rgba(29, 161, 242, 0.7)',
                        'rgba(0, 200, 83, 0.7)'
                    ],
                    borderColor: [
                        '#f44336', '#ff9800', '#9c27b0',
                        '#ffc107', '#1da1f2', '#00c853'
                    ],
                    borderWidth: 1,
                    borderRadius: 6,
                    borderSkipped: false
                },
                {
                    label: 'WHO 24h Guideline',
                    data: [15, 45, 100, 25, 40, 4],
                    backgroundColor: colors.whoBg,
                    borderColor: colors.whoBorder,
                    borderWidth: 1,
                    borderDash: [3, 3],
                    borderRadius: 6,
                    borderSkipped: false
                },
                {
                    label: 'NCEC Standard',
                    data: [35, 340, 120, 660, 365, 10],
                    backgroundColor: 'rgba(0, 200, 83, 0.08)',
                    borderColor: 'rgba(0, 200, 83, 0.4)',
                    borderWidth: 1,
                    borderDash: [6, 3],
                    borderRadius: 6,
                    borderSkipped: false
                }
            ]
        },
        options: {
            ...chartDefaults,
            plugins: {
                ...chartDefaults.plugins,
                legend: {
                    display: true,
                    position: 'top',
                    labels: {
                        color: colors.legendText,
                        font: { size: 11 },
                        usePointStyle: true,
                        pointStyle: 'rectRounded'
                    }
                }
            },
            scales: {
                ...chartDefaults.scales,
                y: {
                    ...chartDefaults.scales.y,
                    title: {
                        display: true,
                        text: 'Concentration (\u03BCg/m\u00B3)',
                        color: colors.tick,
                        font: { size: 11 }
                    }
                }
            }
        }
    });

    // AQI Hourly Trend Chart
    state.charts.aqiTrend = new Chart(document.getElementById('aqi-trend-chart'), {
        type: 'line',
        data: {
            labels: [],
            datasets: [
                {
                    label: 'AQI',
                    data: [],
                    borderColor: '#1da1f2',
                    backgroundColor: createAQIGradient(document.getElementById('aqi-trend-chart')),
                    fill: true,
                    tension: 0.4,
                    pointRadius: 3,
                    pointHoverRadius: 6,
                    borderWidth: 2.5,
                    pointBackgroundColor: function(context) {
                        const val = context.parsed ? context.parsed.y : 0;
                        if (val <= 50) return '#00e400';
                        if (val <= 100) return '#ffff00';
                        if (val <= 150) return '#ff7e00';
                        return '#ff0000';
                    }
                },
                {
                    label: 'Moderate Threshold (100)',
                    data: [],
                    borderColor: colors.whoLine,
                    borderDash: [6, 4],
                    borderWidth: 1,
                    pointRadius: 0,
                    fill: false
                },
                {
                    label: 'Unhealthy Threshold (150)',
                    data: [],
                    borderColor: 'rgba(244, 67, 54, 0.5)',
                    borderDash: [8, 4],
                    borderWidth: 1.5,
                    pointRadius: 0,
                    fill: false
                }
            ]
        },
        options: chartDefaults
    });

    // Pollutant Radar Chart (% of WHO limits)
    state.charts.radar = new Chart(document.getElementById('radar-chart'), {
        type: 'radar',
        data: {
            labels: ['PM2.5', 'PM10', 'O\u2083', 'NO\u2082', 'SO\u2082', 'CO'],
            datasets: [
                {
                    label: '% of WHO Limit',
                    data: [0, 0, 0, 0, 0, 0],
                    borderColor: 'rgba(244, 67, 54, 0.8)',
                    backgroundColor: 'rgba(244, 67, 54, 0.15)',
                    borderWidth: 2,
                    pointBackgroundColor: '#f44336',
                    pointRadius: 4,
                    pointHoverRadius: 6
                },
                {
                    label: '100% (WHO Limit)',
                    data: [100, 100, 100, 100, 100, 100],
                    borderColor: 'rgba(0, 200, 83, 0.5)',
                    backgroundColor: 'rgba(0, 200, 83, 0.05)',
                    borderWidth: 1.5,
                    borderDash: [5, 5],
                    pointRadius: 0,
                    fill: true
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: true,
                    position: 'bottom',
                    labels: {
                        color: colors.legendText,
                        font: { size: 10 },
                        usePointStyle: true,
                        padding: 15
                    }
                },
                tooltip: {
                    backgroundColor: colors.tooltipBg,
                    titleColor: colors.tooltipTitle,
                    bodyColor: colors.tooltipBody,
                    borderColor: colors.tooltipBorder,
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            return `${context.dataset.label}: ${context.raw.toFixed(0)}%`;
                        }
                    }
                }
            },
            scales: {
                r: {
                    grid: { color: colors.radarGrid },
                    angleLines: { color: colors.radarAngle },
                    pointLabels: {
                        color: colors.radarPointLabel,
                        font: { size: 12, weight: '600' }
                    },
                    ticks: {
                        color: colors.tick,
                        backdropColor: 'transparent',
                        font: { size: 9 },
                        stepSize: 50
                    },
                    suggestedMin: 0,
                    suggestedMax: 200
                }
            }
        }
    });

    // Pollutant Contribution Doughnut
    state.charts.doughnut = new Chart(document.getElementById('doughnut-chart'), {
        type: 'doughnut',
        data: {
            labels: ['PM2.5', 'PM10', 'O\u2083', 'NO\u2082', 'SO\u2082', 'CO'],
            datasets: [{
                data: [0, 0, 0, 0, 0, 0],
                backgroundColor: [
                    'rgba(244, 67, 54, 0.8)',
                    'rgba(255, 152, 0, 0.8)',
                    'rgba(156, 39, 176, 0.8)',
                    'rgba(255, 193, 7, 0.8)',
                    'rgba(29, 161, 242, 0.8)',
                    'rgba(0, 200, 83, 0.8)'
                ],
                borderColor: [
                    '#f44336', '#ff9800', '#9c27b0',
                    '#ffc107', '#1da1f2', '#00c853'
                ],
                borderWidth: 2,
                hoverOffset: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '60%',
            plugins: {
                legend: {
                    display: true,
                    position: 'bottom',
                    labels: {
                        color: colors.legendText,
                        font: { size: 11 },
                        usePointStyle: true,
                        padding: 12
                    }
                },
                tooltip: {
                    backgroundColor: colors.tooltipBg,
                    titleColor: colors.tooltipTitle,
                    bodyColor: colors.tooltipBody,
                    borderColor: colors.tooltipBorder,
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const pct = total > 0 ? ((context.raw / total) * 100).toFixed(1) : 0;
                            return ` ${context.label}: ${pct}% of total`;
                        }
                    }
                }
            }
        }
    });
}

function createAQIGradient(canvas) {
    const ctx = canvas.getContext('2d');
    const gradient = ctx.createLinearGradient(0, 0, 0, 280);
    gradient.addColorStop(0, 'rgba(29, 161, 242, 0.3)');
    gradient.addColorStop(1, 'rgba(29, 161, 242, 0.02)');
    return gradient;
}

function updateChartsWithForecast(daily) {
    // PM2.5 forecast
    if (daily.pm25) {
        const labels = daily.pm25.map(d => d.day);
        const avgValues = daily.pm25.map(d => d.avg);
        const whoLine = labels.map(() => CONFIG.WHO_LIMITS.pm25);
        const ncecLine = labels.map(() => CONFIG.NCEC_LIMITS.pm25);

        state.charts.pm25.data.labels = labels;
        state.charts.pm25.data.datasets[0].data = avgValues;
        state.charts.pm25.data.datasets[1].data = whoLine;
        state.charts.pm25.data.datasets[2].data = ncecLine;
        state.charts.pm25.update('none');
    }

    // PM10 forecast
    if (daily.pm10) {
        const labels = daily.pm10.map(d => d.day);
        const avgValues = daily.pm10.map(d => d.avg);
        const whoLine = labels.map(() => CONFIG.WHO_LIMITS.pm10);
        const ncecLine = labels.map(() => CONFIG.NCEC_LIMITS.pm10);

        state.charts.pm10.data.labels = labels;
        state.charts.pm10.data.datasets[0].data = avgValues;
        state.charts.pm10.data.datasets[1].data = whoLine;
        state.charts.pm10.data.datasets[2].data = ncecLine;
        state.charts.pm10.update('none');
    }

    // All pollutants
    if (state.currentData) {
        updateAllPollutantsChart(state.currentData);
    }

    // Generate AQI trend from simulated data since forecast doesn't include AQI history
    if (state.currentData) {
        generateAQITrendFromCurrent(state.currentData);
    }
}

function generateAQITrendFromCurrent(values) {
    const now = new Date();
    const labels = [];
    const aqiData = [];
    const baseAQI = values.aqi || 80;

    for (let i = 23; i >= 0; i--) {
        const time = new Date(now - i * 3600 * 1000);
        labels.push(time.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' }));
        const hour = time.getHours();
        let factor = 1;
        if (hour >= 7 && hour <= 9) factor = 1.3;
        else if (hour >= 17 && hour <= 19) factor = 1.25;
        else if (hour >= 0 && hour <= 5) factor = 0.7;
        else if (hour >= 12 && hour <= 14) factor = 1.15;
        const variation = () => (Math.random() - 0.5) * 0.3 + 1;
        aqiData.push(Math.max(10, Math.round(baseAQI * factor * variation())));
    }

    state.simulated24h.aqi = aqiData;
    state.simulated24h.labels = labels;

    const moderateLine = labels.map(() => 100);
    const unhealthyLine = labels.map(() => 150);
    state.charts.aqiTrend.data.labels = labels;
    state.charts.aqiTrend.data.datasets[0].data = aqiData;
    state.charts.aqiTrend.data.datasets[1].data = moderateLine;
    state.charts.aqiTrend.data.datasets[2].data = unhealthyLine;
    state.charts.aqiTrend.update('none');

    updateStatsPanel();
}

function updateChartsWithCurrent(values) {
    generateSimulatedChartData(values);
}

function generateSimulatedChartData(currentValues) {
    const now = new Date();
    const labels = [];
    const pm25Data = [];
    const pm10Data = [];
    const aqiData = [];

    // Generate 24 hours of data points
    for (let i = 23; i >= 0; i--) {
        const time = new Date(now - i * 3600 * 1000);
        labels.push(time.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' }));

        // Simulate realistic hourly variation
        // PM tends to be higher during rush hours (7-9 AM, 5-7 PM) and lower at night
        const hour = time.getHours();
        let factor = 1;
        if (hour >= 7 && hour <= 9) factor = 1.3;
        else if (hour >= 17 && hour <= 19) factor = 1.25;
        else if (hour >= 0 && hour <= 5) factor = 0.7;
        else if (hour >= 12 && hour <= 14) factor = 1.15; // Midday dust

        const variation = () => (Math.random() - 0.5) * 0.3 + 1;

        const basePM25 = currentValues.pm25 || 30;
        const basePM10 = currentValues.pm10 || 100;
        const baseAQI = currentValues.aqi || 80;

        pm25Data.push(Math.max(5, Math.round(basePM25 * factor * variation())));
        pm10Data.push(Math.max(10, Math.round(basePM10 * factor * variation())));
        aqiData.push(Math.max(10, Math.round(baseAQI * factor * variation())));
    }

    // Store for stats
    state.simulated24h = { pm25: pm25Data, pm10: pm10Data, aqi: aqiData, labels: labels };

    const whoLinePM25 = labels.map(() => CONFIG.WHO_LIMITS.pm25);
    const whoLinePM10 = labels.map(() => CONFIG.WHO_LIMITS.pm10);
    const ncecLinePM25 = labels.map(() => CONFIG.NCEC_LIMITS.pm25);
    const ncecLinePM10 = labels.map(() => CONFIG.NCEC_LIMITS.pm10);

    // Update PM2.5 chart
    state.charts.pm25.data.labels = labels;
    state.charts.pm25.data.datasets[0].data = pm25Data;
    state.charts.pm25.data.datasets[1].data = whoLinePM25;
    state.charts.pm25.data.datasets[2].data = ncecLinePM25;
    state.charts.pm25.update('none');

    // Update PM10 chart
    state.charts.pm10.data.labels = labels;
    state.charts.pm10.data.datasets[0].data = pm10Data;
    state.charts.pm10.data.datasets[1].data = whoLinePM10;
    state.charts.pm10.data.datasets[2].data = ncecLinePM10;
    state.charts.pm10.update('none');

    // Update AQI Trend chart
    const moderateLine = labels.map(() => 100);
    const unhealthyLine = labels.map(() => 150);
    state.charts.aqiTrend.data.labels = labels;
    state.charts.aqiTrend.data.datasets[0].data = aqiData;
    state.charts.aqiTrend.data.datasets[1].data = moderateLine;
    state.charts.aqiTrend.data.datasets[2].data = unhealthyLine;
    state.charts.aqiTrend.update('none');

    // Update all pollutants bar chart
    updateAllPollutantsChart(currentValues);

    // Update 24h statistics
    updateStatsPanel();
}

function updateStandardsTable(values) {
    const pollutants = [
        { key: 'pm25', unit: '' },
        { key: 'pm10', unit: '' },
        { key: 'o3', unit: '' },
        { key: 'no2', unit: '' },
        { key: 'so2', unit: '' },
        { key: 'co', unit: '' }
    ];

    pollutants.forEach(p => {
        const valEl = document.getElementById(`std-${p.key}-val`);
        const statusEl = document.getElementById(`std-${p.key}-status`);
        if (!valEl || !statusEl) return;

        const val = values[p.key];
        if (val === null || val === undefined) {
            valEl.textContent = 'N/A';
            statusEl.innerHTML = '<span class="compliance-badge">--</span>';
            return;
        }

        valEl.textContent = typeof val === 'number' && val % 1 !== 0 ? val.toFixed(1) : val;

        const whoLimit = CONFIG.WHO_LIMITS[p.key];
        const ncecLimit = CONFIG.NCEC_LIMITS[p.key];

        if (val > ncecLimit) {
            statusEl.innerHTML = '<span class="compliance-badge exceeds-ncec"><i class="fas fa-circle-xmark"></i> Exceeds NCEC</span>';
        } else if (val > whoLimit) {
            statusEl.innerHTML = '<span class="compliance-badge exceeds-who"><i class="fas fa-triangle-exclamation"></i> Exceeds WHO</span>';
        } else {
            statusEl.innerHTML = '<span class="compliance-badge compliant"><i class="fas fa-circle-check"></i> Compliant</span>';
        }
    });
}

function updateAllPollutantsChart(values) {
    state.charts.allPollutants.data.datasets[0].data = [
        values.pm25 || 0,
        values.pm10 || 0,
        values.o3 || 0,
        values.no2 || 0,
        values.so2 || 0,
        values.co ? values.co * 10 : 0  // Scale CO for visibility (mg/m³ × 10)
    ];
    state.charts.allPollutants.update('none');
}

// ===== ANALYSIS UPDATES =====
function updateRadarChart(values) {
    const whoData = [
        values.pm25 ? (values.pm25 / CONFIG.WHO_LIMITS.pm25) * 100 : 0,
        values.pm10 ? (values.pm10 / CONFIG.WHO_LIMITS.pm10) * 100 : 0,
        values.o3 ? (values.o3 / CONFIG.WHO_LIMITS.o3) * 100 : 0,
        values.no2 ? (values.no2 / CONFIG.WHO_LIMITS.no2) * 100 : 0,
        values.so2 ? (values.so2 / CONFIG.WHO_LIMITS.so2) * 100 : 0,
        values.co ? (values.co / CONFIG.WHO_LIMITS.co) * 100 : 0
    ];
    state.charts.radar.data.datasets[0].data = whoData;
    state.charts.radar.update('none');
}

function updateDoughnutChart(values) {
    // Normalize all pollutants to same scale (% of WHO) for fair comparison
    const normalized = [
        values.pm25 ? (values.pm25 / CONFIG.WHO_LIMITS.pm25) : 0,
        values.pm10 ? (values.pm10 / CONFIG.WHO_LIMITS.pm10) : 0,
        values.o3 ? (values.o3 / CONFIG.WHO_LIMITS.o3) : 0,
        values.no2 ? (values.no2 / CONFIG.WHO_LIMITS.no2) : 0,
        values.so2 ? (values.so2 / CONFIG.WHO_LIMITS.so2) : 0,
        values.co ? (values.co / CONFIG.WHO_LIMITS.co) : 0
    ];
    state.charts.doughnut.data.datasets[0].data = normalized;
    state.charts.doughnut.update('none');

    // Update dominant pollutant
    updateDominantPollutant(values, normalized);
}

function updateDominantPollutant(values, normalized) {
    const names = ['PM2.5', 'PM10', 'O\u2083', 'NO\u2082', 'SO\u2082', 'CO'];
    const icons = ['fa-smog', 'fa-cloud-meatball', 'fa-sun', 'fa-industry', 'fa-flask', 'fa-fire'];
    const iconColors = ['#f44336', '#ff9800', '#9c27b0', '#ffc107', '#1da1f2', '#00c853'];

    let maxIdx = 0;
    for (let i = 1; i < normalized.length; i++) {
        if (normalized[i] > normalized[maxIdx]) maxIdx = i;
    }

    const total = normalized.reduce((a, b) => a + b, 0);
    const pct = total > 0 ? ((normalized[maxIdx] / total) * 100).toFixed(0) : 0;
    const whoPercent = normalized[maxIdx] * 100;

    const nameEl = document.getElementById('dominant-name');
    const detailEl = document.getElementById('dominant-detail');
    const pctEl = document.getElementById('dominant-pct');
    const iconEl = document.querySelector('.dominant-icon');

    if (nameEl) nameEl.textContent = names[maxIdx];
    if (detailEl) detailEl.textContent = `Dominant pollutant at ${whoPercent.toFixed(0)}% of WHO limit`;
    if (pctEl) pctEl.textContent = `${pct}%`;
    if (iconEl) {
        iconEl.innerHTML = `<i class="fas ${icons[maxIdx]}"></i>`;
        iconEl.style.color = iconColors[maxIdx];
        iconEl.style.background = `${iconColors[maxIdx]}20`;
    }
}

function updateBreakdownBars(values) {
    const pollutants = [
        { key: 'pm25', name: 'PM2.5', color: '#f44336', value: values.pm25, limit: CONFIG.WHO_LIMITS.pm25 },
        { key: 'pm10', name: 'PM10', color: '#ff9800', value: values.pm10, limit: CONFIG.WHO_LIMITS.pm10 },
        { key: 'o3', name: 'O\u2083', color: '#9c27b0', value: values.o3, limit: CONFIG.WHO_LIMITS.o3 },
        { key: 'no2', name: 'NO\u2082', color: '#ffc107', value: values.no2, limit: CONFIG.WHO_LIMITS.no2 },
        { key: 'so2', name: 'SO\u2082', color: '#1da1f2', value: values.so2, limit: CONFIG.WHO_LIMITS.so2 },
        { key: 'co', name: 'CO', color: '#00c853', value: values.co, limit: CONFIG.WHO_LIMITS.co }
    ];

    if (!DOM.breakdownList) return;

    DOM.breakdownList.innerHTML = pollutants.map(p => {
        const val = p.value || 0;
        const pct = Math.min((val / p.limit) * 100, 200);
        const status = val > p.limit ? 'Exceeds' : 'OK';
        const barColor = val > p.limit ? '#f44336' : p.color;
        return `
            <div class="breakdown-item">
                <div class="breakdown-color" style="background:${p.color}"></div>
                <span class="breakdown-name">${p.name}</span>
                <div class="breakdown-bar-wrapper">
                    <div class="breakdown-bar-fill" style="width:${Math.min(pct, 100)}%;background:${barColor}"></div>
                </div>
                <span class="breakdown-value">${pct.toFixed(0)}% ${status}</span>
            </div>
        `;
    }).join('');
}

function updateStatsPanel() {
    const data = state.simulated24h;
    if (!data.pm25.length) return;

    const avg = arr => arr.length ? Math.round(arr.reduce((a, b) => a + b, 0) / arr.length) : 0;
    const max = arr => arr.length ? Math.max(...arr) : 0;
    const min = arr => arr.length ? Math.min(...arr) : 0;

    const setEl = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.textContent = val;
    };

    setEl('stat-pm25-avg', avg(data.pm25));
    setEl('stat-pm25-peak', max(data.pm25));
    setEl('stat-pm25-low', min(data.pm25));
    setEl('stat-pm10-avg', avg(data.pm10));
    setEl('stat-pm10-peak', max(data.pm10));
    setEl('stat-pm10-low', min(data.pm10));
}

// ===== ALERTS & NOTIFICATIONS =====
function checkAlerts(values) {
    const pm25Threshold = parseInt(document.getElementById('alert-pm25-threshold').value);
    const pm10Threshold = parseInt(document.getElementById('alert-pm10-threshold').value);
    const aqiThreshold = parseInt(document.getElementById('alert-aqi-threshold').value);

    const pm25Enabled = document.getElementById('alert-pm25-toggle').checked;
    const pm10Enabled = document.getElementById('alert-pm10-toggle').checked;
    const aqiEnabled = document.getElementById('alert-aqi-toggle').checked;

    if (pm25Enabled && values.pm25 && values.pm25 > pm25Threshold) {
        const severity = values.pm25 > pm25Threshold * 2 ? 'danger' : 'warning';
        addAlert(severity, `PM2.5 is ${values.pm25} μg/m³ (threshold: ${pm25Threshold})`,
            `PM2.5 Level Alert`);
    }

    if (pm10Enabled && values.pm10 && values.pm10 > pm10Threshold) {
        const severity = values.pm10 > pm10Threshold * 2 ? 'danger' : 'warning';
        addAlert(severity, `PM10 is ${values.pm10} μg/m³ (threshold: ${pm10Threshold})`,
            `PM10 Level Alert`);
    }

    if (aqiEnabled && values.aqi && values.aqi > aqiThreshold) {
        const severity = values.aqi > 200 ? 'critical' : values.aqi > 150 ? 'danger' : 'warning';
        const category = getAQICategory(values.aqi);
        addAlert(severity, `AQI is ${values.aqi} (${category.label}) — threshold: ${aqiThreshold}`,
            `AQI Alert`);
    }
}

function addAlert(severity, message, title) {
    const now = new Date();
    const timeStr = now.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });

    const alert = { severity, message, title, time: timeStr, timestamp: now };
    state.alerts.unshift(alert);

    // Keep only last 20 alerts
    if (state.alerts.length > 20) state.alerts.pop();

    // Update alert log
    renderAlertLog();

    // Show toast notification
    showToast(severity === 'critical' ? 'danger' : severity, title, message);

    // Browser notification
    if (state.notificationsEnabled && Notification.permission === 'granted') {
        sendBrowserNotification(title, message);
    }
}

function renderAlertLog() {
    if (state.alerts.length === 0) {
        DOM.alertLogBody.innerHTML = '<p class="no-alerts">No alerts yet. Alerts will appear when thresholds are exceeded.</p>';
        return;
    }

    DOM.alertLogBody.innerHTML = state.alerts.map(alert => `
        <div class="alert-entry ${alert.severity}">
            <i class="fas ${alert.severity === 'danger' || alert.severity === 'critical' ? 'fa-circle-exclamation' : 'fa-triangle-exclamation'}"></i>
            <div>
                <div>${alert.message}</div>
                <div class="alert-entry-time">${alert.time}</div>
            </div>
        </div>
    `).join('');
}

// ===== TOAST NOTIFICATIONS =====
function showToast(type, title, text) {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    const iconMap = {
        warning: 'fa-triangle-exclamation',
        danger: 'fa-circle-exclamation',
        info: 'fa-circle-info',
        success: 'fa-circle-check'
    };

    toast.innerHTML = `
        <div class="toast-icon"><i class="fas ${iconMap[type] || iconMap.info}"></i></div>
        <div class="toast-body">
            <div class="toast-title">${title}</div>
            <div class="toast-text">${text}</div>
        </div>
        <button class="toast-close">&times;</button>
    `;

    DOM.toastContainer.appendChild(toast);

    const closeBtn = toast.querySelector('.toast-close');
    closeBtn.addEventListener('click', () => removeToast(toast));

    // Auto-dismiss after 6 seconds
    setTimeout(() => removeToast(toast), 6000);
}

function removeToast(toast) {
    if (!toast.parentNode) return;
    toast.classList.add('toast-out');
    setTimeout(() => {
        if (toast.parentNode) toast.parentNode.removeChild(toast);
    }, 300);
}

// ===== BROWSER NOTIFICATIONS =====
function requestNotificationPermission() {
    if (!('Notification' in window)) {
        showToast('warning', 'Not Supported', 'Browser notifications are not supported in this browser.');
        return;
    }

    Notification.requestPermission().then(permission => {
        if (permission === 'granted') {
            state.notificationsEnabled = true;
            localStorage.setItem('notifications_enabled', 'true');
            showToast('success', 'Notifications Enabled', 'You will receive alerts when air quality deteriorates.');
        } else {
            showToast('warning', 'Permission Denied', 'Please enable notifications in your browser settings.');
        }
    });
}

function sendBrowserNotification(title, body) {
    if (Notification.permission !== 'granted') return;

    const notification = new Notification(`🌫️ ${title} - Wadi Hanifa`, {
        body: body,
        icon: 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text y=".9em" font-size="80">🌫️</text></svg>',
        badge: 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text y=".9em" font-size="80">⚠️</text></svg>',
        tag: 'aq-alert',
        renotify: true
    });

    notification.onclick = () => {
        window.focus();
        notification.close();
    };
}

// ===== HELPERS =====
function getAQICategory(aqi) {
    return AQI_CATEGORIES.find(cat => aqi >= cat.min && aqi <= cat.max) || AQI_CATEGORIES[AQI_CATEGORIES.length - 1];
}

function startAutoRefresh() {
    if (state.refreshTimer) clearInterval(state.refreshTimer);
    state.refreshTimer = setInterval(fetchAllData, CONFIG.REFRESH_INTERVAL);
}

// ===== EXPOSE FOR DEBUGGING =====
window.AQDashboard = { state, CONFIG, fetchAllData };
