const {
    prefix, use_teams, watched,
    num_x: numX, num_y: numY,
    c0, c1, c2,

    json_players        : players,
    json_tour_stats     : tourStats,
    json_teams          : teamStats,
    json_team_hl_rules  : teamHlRules,
    json_tier_merged    : tierStats,
    json_songs          : songData,
    json_matrix_songs   : matrixSongs,
    json_scatter        : scatterData,
    json_arrows         : arrowData,
    json_borders        : groupBorders,
    json_eligibility    : eligibility,
    json_hl_rules       : hlRules,
    json_explanations   : colExplanations,
    generated_timestamp : generatedTime
} = window.dashboardData;

let currentPlayerMetricMode = "%";
let currentTierChartMode    = "TIER";
let globalSearchData        = [];
let globalChartMode         = "RATE";
let c1Sub                   = "BASE";
let c2Sub                   = "BOTH";
let c3Mode                  = "MED";

document.getElementById('dashboardTitle').innerText = prefix;

const dynamicStyles = document.createElement('style');

dynamicStyles.innerHTML = `
    .highlight-best     {background-color: ${c2} !important}
    .highlight-worst    {background-color: ${c0} !important}

    td[data-songs].highlight-best:hover{color: ${c2} !important}
    td[data-songs].highlight-worst:hover{color: ${c0} !important}

    td[data-songs].highlight-best:hover::after{background-color: ${c2} !important}
    td[data-songs].highlight-worst:hover::after{background-color: ${c0} !important}
`;

document.head.appendChild(dynamicStyles);

const thickBorderColumns = new Set([
    "Player",
    "Tier",
    "Mean Over-8",
    "Lives Saved",
    "Rig Rate",
    "Solo Rig Rate",
    "Over-8 Δ",
    "Rig Δ",
    "Metric",
    "Value",
    "Team Leader",
    "Tier",
    "Lives Saved",
    "Median Vintage Hit",
    "Chant GR"
]);

window.unifiedChartLimits = {xMin: 0, xMax: 8, yMin: 1980, yMax: 2026, dtickY: 5};

if (typeof scatterData !== 'undefined' && scatterData) {
    const allXValues = [...scatterData.map(d => d.over8)];
    const allYValues = [...scatterData.map(d => d.vintage)];

    if (typeof arrowData !== 'undefined' && arrowData) {
        allXValues.push(...arrowData.map(d => d.x_start), ...arrowData.map(d => d.x_end));
        allYValues.push(...arrowData.map(d => d.y_start), ...arrowData.map(d => d.y_end));
    }

    window.unifiedChartLimits = {
        xMin    : Math.min(...allXValues) - 0.25,
        xMax    : Math.max(...allXValues) + 0.25,
        yMin    : Math.min(...allYValues) - 1,
        yMax    : Math.max(...allYValues) + 1,
        dtickY  : Math.max(2, Math.ceil((Math.max(...allYValues) - Math.min(...allYValues)) / 5))
    };
}

function updateTimeSubtitle() {
    const subtitle = document.getElementById('lastUpdatedSubtitleText');
    if (!subtitle) return;

    const differenceInMiliseconds   = Date.now() - generatedTime;
    const differenceInSeconds       = Math.floor(differenceInMiliseconds    / 1000);
    const differenceInMinutes       = Math.floor(differenceInSeconds        / 60);
    const differenceInHours         = Math.floor(differenceInMinutes        / 60);
    const differenceInDays          = Math.floor(differenceInHours          / 24);
    const differenceInWeeks         = Math.floor(differenceInDays           / 7);
    const differenceInMonths        = Math.floor(differenceInWeeks          / 4);
    const differenceInYears         = Math.floor(differenceInMonths         / 12);

    let displayString = "Last updated: ";

    if      (differenceInSeconds    < 60)   displayString += `${differenceInSeconds} seconds ago`;
    else if (differenceInMinutes    < 60)   displayString += `${differenceInMinutes} minute${differenceInMinutes    === 1 ? '' : 's'} ago`;
    else if (differenceInHours      < 24)   displayString += `${differenceInHours} hour${differenceInHours          === 1 ? '' : 's'} ago`;
    else if (differenceInDays       < 7)    displayString += `${differenceInDays} day${differenceInDays             === 1 ? '' : 's'} ago`;
    else if (differenceInWeeks      < 24)   displayString += `${differenceInWeeks} week${differenceInWeeks          === 1 ? '' : 's'} ago`;
    else if (differenceInMonths     < 24)   displayString += `${differenceInMonths} month${differenceInMonths       === 1 ? '' : 's'} ago`;
    else                                    displayString += `${differenceInYears} year${differenceInDays           === 1 ? '' : 's'} ago`;

    subtitle.innerText = displayString;
}

updateTimeSubtitle();
setInterval(updateTimeSubtitle, 1000);

function debounce(func, wait) {
    let timeout;
    return function(...args) {
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(this, args), wait);
    };
}

function trimNames(input) {
    if (!input) return '';

    let flatString  = Array.isArray(input) ? input.join('||') : String(input);
    let normalized  = flatString.replace(/\s*(?:,|\b(?<!\d)\/(?!\d)\b|・|&|×|\bfeat\.)\s*/gi, '||');
    let arr         = normalized.split('||').map(x => x.trim()).filter(Boolean);

    if (arr.length <= 3)    return arr.join(', ');
    else                    return `${arr.slice(0, 2).join(', ')}, and more`;
}

function parseVintageToFloat(vintStr) {
    const parts = vintStr.trim().split(/\s+/);    
    if (parts.length === 1 && !isNaN(parts[0])) return parseFloat(parts[0]);

    const season    = parts[0].toLowerCase();
    const year      = parseInt(parts[1]);
    let weight      = 0.0;

    if      (season === "winter")   weight = 0.00;
    else if (season === "spring")   weight = 0.25;
    else if (season === "summer")   weight = 0.50;
    else if (season === "fall")     weight = 0.75;

    return year + weight;
}

function parseFloatToVintage(val) {
    if (val === null || val === undefined || isNaN(val) || val === -Infinity) return "N/A";

    const year      = Math.floor(val);
    const remainder = Math.round((val - year) * 100) / 100;

    let season;

    if      (remainder < 0.25)  season = "Winter";
    else if (remainder < 0.50)  season = "Spring";
    else if (remainder < 0.75)  season = "Summer";
    else                        season = "Fall";
    
    return `${season}&nbsp;${year}`;
}

function hexToRgba(hex, opacity = 0.95) {
    let c = hex.replace('#', '');
    if (c.length === 3) c = c.split('').map(x => x + x).join('');

    const r = parseInt(c.substring(0, 2), 16);
    const g = parseInt(c.substring(2, 4), 16);
    const b = parseInt(c.substring(4, 6), 16);

    return `rgba(${r}, ${g}, ${b}, ${opacity})`;
}

function generateLinearColors(startHex, endHex, steps) {
    const parse = (hex) => [parseInt(hex.slice(1,3),16), parseInt(hex.slice(3,5),16), parseInt(hex.slice(5,7),16)];
    const cA    = parse(startHex), cB = parse(endHex);

    if (steps <= 1) return [startHex];
    const arr = [];

    for (let i = 0; i < steps; i++) {
        const t = i / (steps - 1);
        const r = Math.round(cA[0] + t * (cB[0] - cA[0]));
        const g = Math.round(cA[1] + t * (cB[1] - cA[1]));
        const b = Math.round(cA[2] + t * (cB[2] - cA[2]));

        arr.push(`rgb(${r},${g},${b})`);
    }

    return arr;
}

const c8Colors = generateLinearColors(c2, c0, 8);