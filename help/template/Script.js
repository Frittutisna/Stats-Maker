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

    td[data-songs].highlight-best   :hover{color: ${c2} !important}
    td[data-songs].highlight-worst  :hover{color: ${c0} !important}

    td[data-songs].highlight-best   :hover::after{background-color: ${c2} !important}
    td[data-songs].highlight-worst  :hover::after{background-color: ${c0} !important}
`;

document.head.appendChild(dynamicStyles);

const tabContainer  = document.getElementById('tabContainer');
const tourTabBtn    = document.getElementById('tourTabBtn');
const helpAnchor    = document.getElementById('globalHelpWrapper');

if (use_teams)  tourTabBtn.innerText = "Tour/Team";
else            tourTabBtn.innerText = "Tour";

if (use_teams)  helpAnchor.insertAdjacentHTML('beforebegin', `<button class="tab-btn" onclick="switchDashboardTab(event, 'tier-tab')">Tier</button>`);
                helpAnchor.insertAdjacentHTML('beforebegin', `<button class="tab-btn" onclick="switchDashboardTab(event, 'song-tab')">Song</button>`);
if (watched)    helpAnchor.insertAdjacentHTML('beforebegin', `<button class="tab-btn" onclick="switchDashboardTab(event, 'guess-tab')">Guess/List</button>`);
else            helpAnchor.insertAdjacentHTML('beforebegin', `<button class="tab-btn" onclick="switchDashboardTab(event, 'guess-tab')">Guess</button>`);
                helpAnchor.insertAdjacentHTML('beforebegin', `<button class="tab-btn" onclick="switchDashboardTab(event, 'search-tab')">Search</button>`);

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
    const subtitle = document.getElementById('lastUpdatedSubtitle');
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

function switchDashboardTab(evt, tabId) {
    document.querySelectorAll('.tab-content')   .forEach(el => el.classList.remove('active-content'));
    document.querySelectorAll('.tab-btn')       .forEach(el => el.classList.remove('active-tab'));

    document.getElementById(tabId).classList.add('active-content');
    if (evt && evt.currentTarget) evt.currentTarget.classList.add('active-tab');
    window.dispatchEvent(new Event('resize'));

    const gearWrapper = document.getElementById('globalGearWrapper');
    const helpWrapper = document.getElementById('globalHelpWrapper');

    if (gearWrapper) {
        if (['player-tab', 'tier-tab', 'search-tab'].includes(tabId) || (tabId === 'guess-tab' && watched)) gearWrapper.classList.remove    ('invisible');
        else                                                                                                gearWrapper.classList.add       ('invisible');
    }

    if (helpWrapper) {
        if (['player-tab', 'tier-tab', 'guess-tab', 'search-tab', 'song-tab'].includes(tabId))  helpWrapper.classList.remove('invisible');
        else                                                                                    helpWrapper.classList.add('invisible');
    }

    document.querySelectorAll('#globalGearWrapper > div, #globalHelpWrapper > div').forEach(el => {if (el.id.includes('Dropdown')) el.classList.add('hidden');});
}

window.toggleGlobalGear = function(event) {
    event.stopPropagation();
    const activeTab = document.querySelector('.tab-content.active-content').id;
    
    if (activeTab === 'player-tab') {
        document.getElementById("playerColumnSettingsDropdown") .classList.toggle   ("hidden");
        document.getElementById("columnSettingsDropdown")       .classList.add      ("hidden");
        document.getElementById("tierSettingsDropdown")         .classList.add      ("hidden");
        document.getElementById("guessSettingsDropdown")        .classList.add      ("hidden");
    }

    else if (activeTab === 'tier-tab') {
        document.getElementById("tierSettingsDropdown")         .classList.toggle   ("hidden");
        document.getElementById("playerColumnSettingsDropdown") .classList.add      ("hidden");
        document.getElementById("columnSettingsDropdown")       .classList.add      ("hidden");
        document.getElementById("guessSettingsDropdown")        .classList.add      ("hidden");
    }

    else if (activeTab === 'search-tab') {
        document.getElementById("columnSettingsDropdown")       .classList.toggle   ("hidden");
        document.getElementById("playerColumnSettingsDropdown") .classList.add      ("hidden");
        document.getElementById("tierSettingsDropdown")         .classList.add      ("hidden");
        document.getElementById("guessSettingsDropdown")        .classList.add      ("hidden");
    }

    else if (activeTab === 'guess-tab' && watched) {
        document.getElementById("guessSettingsDropdown")        .classList.toggle   ("hidden");
        document.getElementById("playerColumnSettingsDropdown") .classList.add      ("hidden");
        document.getElementById("columnSettingsDropdown")       .classList.add      ("hidden");
        document.getElementById("tierSettingsDropdown")         .classList.add      ("hidden");
    }
};

window.toggleGlobalHelp = function(event) {
    event.stopPropagation();
    const activeTab = document.querySelector('.tab-content.active-content').id;

    if (activeTab === 'player-tab') {
        document.getElementById("playerGuideDropdown")  .classList.toggle   ("hidden");
        document.getElementById("tierGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("guessGuideDropdown")   .classList.add      ("hidden");
        document.getElementById("songGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("searchGuideDropdown")  .classList.add      ("hidden");
    }

    else if (activeTab === 'tier-tab') {
        document.getElementById("tierGuideDropdown")    .classList.toggle   ("hidden");
        document.getElementById("playerGuideDropdown")  .classList.add      ("hidden");
        document.getElementById("guessGuideDropdown")   .classList.add      ("hidden");
        document.getElementById("songGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("searchGuideDropdown")  .classList.add      ("hidden");
    }

    else if (activeTab === 'guess-tab') {
        document.getElementById("guessGuideDropdown")   .classList.toggle   ("hidden");
        document.getElementById("playerGuideDropdown")  .classList.add      ("hidden");
        document.getElementById("tierGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("songGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("searchGuideDropdown")  .classList.add      ("hidden");
    }
    
    else if (activeTab === 'song-tab') {
        document.getElementById("songGuideDropdown")    .classList.toggle   ("hidden");
        document.getElementById("playerGuideDropdown")  .classList.add      ("hidden");
        document.getElementById("tierGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("guessGuideDropdown")   .classList.add      ("hidden");
        document.getElementById("searchGuideDropdown")  .classList.add      ("hidden");
    }

    else if (activeTab === 'search-tab') {
        document.getElementById("searchGuideDropdown")  .classList.toggle   ("hidden");
        document.getElementById("playerGuideDropdown")  .classList.add      ("hidden");
        document.getElementById("tierGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("songGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("guessGuideDropdown")   .classList.add      ("hidden");
    }
};

document.addEventListener("click", () => {document.querySelectorAll('#globalGearWrapper > div, #globalHelpWrapper > div').forEach(el => {if (el.id.includes('Dropdown')) el.classList.add('hidden');});});
const stopProp = (e) => e.stopPropagation();

[
    'playerColumnSettingsDropdown',
    'tierSettingsDropdown',
    'guessSettingsDropdown',
    'columnSettingsDropdown',
    'playerGuideDropdown',
    'tierGuideDropdown',
    'songGuideDropdown',
    'guessGuideDropdown',
    'searchGuideDropdown'
].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener("click", stopProp);
});

function getCrossProduct(o, a, b, xKey, yKey) {return (a[xKey] - o[xKey]) * (b[yKey] - o[yKey]) - (a[yKey] - o[yKey]) * (b[xKey] - o[xKey]);}

function get75PercentileHull(pts, xKey, yKey) {
    if (pts.length < 3) return null;

    const xVals = pts.map(p => p[xKey]).sort((a, b) => a - b);
    const yVals = pts.map(p => p[yKey]).sort((a, b) => a - b);

    const xMed = xVals[Math.floor(xVals.length / 2)];
    const yMed = yVals[Math.floor(yVals.length / 2)];

    const xRange = (Math.max(...xVals) - Math.min(...xVals)) || 1;
    const yRange = (Math.max(...yVals) - Math.min(...yVals)) || 1;

    const withDist = pts.map(p => {
        const dx = (p[xKey] - xMed) / xRange;
        const dy = (p[yKey] - yMed) / yRange;

        return {p, d: Math.sqrt(dx * dx + dy * dy)};
    });

    const sortedDist    = withDist.map(item => item.d).sort((a, b) => a - b);
    const threshD       = sortedDist[Math.floor(sortedDist.length * 0.75)];
    const packedPts     = withDist.filter(item => item.d < threshD).map(item => item.p);

    if (packedPts.length < 3) return null;
    packedPts.sort((a, b) => a[xKey] == b[xKey] ? a[yKey] - b[yKey] : a[xKey] - b[xKey]);

    const lower = [];

    for (let p of packedPts) {
        while (lower.length >= 2 && getCrossProduct(lower[lower.length-2], lower[lower.length-1], p, xKey, yKey) <= 0) lower.pop();
        lower.push(p);
    }

    const upper = [];

    for (let i = packedPts.length - 1; i >= 0; i--) {
        let p = packedPts[i];
        while (upper.length >= 2 && getCrossProduct(upper[upper.length-2], upper[upper.length-1], p, xKey, yKey) <= 0) upper.pop();
        upper.push(p);
    }

    upper.pop();
    lower.pop();

    const hull = lower.concat(upper);
    return {
        x: hull.map(p => p[xKey]).concat(hull[0][xKey]),
        y: hull.map(p => p[yKey]).concat(hull[0][yKey])
    };
}

const formatAndSortSongsList = (list, prefixBullets = true) => {
    return list
        .sort((a, b) => {
            const cleanA = (a.startsWith('✓') || a.startsWith('✗')) ? a.slice(2) : a;
            const cleanB = (b.startsWith('✓') || b.startsWith('✗')) ? b.slice(2) : b;

            return cleanA.toLowerCase().localeCompare(cleanB.toLowerCase());
        })

        .map(s => (s.startsWith('✓') || s.startsWith('✗') || !prefixBullets) ? s : `• ${s}`);
};

const sampleLargeSongList = (displaySongs) => {
    const ticks     = displaySongs.filter(s => s.startsWith('✓')).sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
    const crosses   = displaySongs.filter(s => s.startsWith('✗')).sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));

    return [...ticks, ...crosses];
};

let globalPlayerSortState   = {columnName: "GR", ascending: false};
let globalFilteredPlayers   = [];
let globalMetricHighlights  = {};

const playerHeadersMasterConfig = [
    {id: "player",              name: "Player",                 ascMetric: false,   teamReq: false, watchedReq: false,  def: true,  type: "text"},
    {id: "team",                name: "Team",                   ascMetric: false,   teamReq: true,  watchedReq: false,  def: false, type: "categorical",    subOptions: []},
    {id: "tier",                name: "Tier",                   ascMetric: true,    teamReq: true,  watchedReq: false,  def: false, type: "categorical",    subOptions: ["1", "2", "3", "4"]},
    {id: "elo",                 name: "Elo",                    ascMetric: false,   teamReq: true,  watchedReq: false,  def: true,  type: "range",          min: -10,   max: 200,   step: 1},
    {id: "guessrate",           name: "GR",                     ascMetric: false,   teamReq: false, watchedReq: false,  def: true,  type: "range",          min: 0,     max: 100,   step: 1},
    {id: "grdelta",             name: "GR Δ",                   ascMetric: false,   teamReq: false, watchedReq: false,  def: false, type: "range",          min: -100,  max: 100,   step: 1},
    {id: "uf",                  name: "UF",                     ascMetric: false,   teamReq: true,  watchedReq: false,  def: true,  type: "range",          min: 0,     max: 100,   step: 1},
    {id: "ufdelta",             name: "UF Δ",                   ascMetric: false,   teamReq: true,  watchedReq: false,  def: false, type: "range",          min: -100,  max: 100,   step: 1},
    {id: "score",               name: "Score",                  ascMetric: false,   teamReq: true,  watchedReq: false,  def: false, type: "range",          min: 0,     max: 100,   step: 1},
    {id: "18s",                 name: "1/8s",                   ascMetric: false,   teamReq: false, watchedReq: false,  def: true,  type: "range",          min: 0,     max: 100,   step: 1},
    {id: "28s",                 name: "2/8s",                   ascMetric: false,   teamReq: false, watchedReq: false,  def: true,  type: "range",          min: 0,     max: 100,   step: 1},
    {id: "78s",                 name: "7/8s",                   ascMetric: true,    teamReq: false, watchedReq: false,  def: true,  type: "range",          min: 0,     max: 100,   step: 1},
    {id: "meanover8",           name: "Mean Over-8",            ascMetric: true,    teamReq: false, watchedReq: false,  def: false, type: "range",          min: 0,     max: 8,     step: 0.01},
    {id: "livestaken",          name: "Lives Taken",            ascMetric: false,   teamReq: true,  watchedReq: false,  def: false, type: "range",          min: 0,     max: 100,   step: 1},
    {id: "livessaved",          name: "Lives Saved",            ascMetric: false,   teamReq: true,  watchedReq: false,  def: false, type: "range",          min: 0,     max: 100,   step: 1},
    {id: "opguessrate",         name: "OP GR",                  ascMetric: false,   teamReq: false, watchedReq: false,  def: true,  type: "range",          min: 0,     max: 100,   step: 1},
    {id: "opdelta",             name: "OP Δ",                   ascMetric: false,   teamReq: false, watchedReq: false,  def: false, type: "range",          min: -100,  max: 100,   step: 1},
    {id: "edguessrate",         name: "ED GR",                  ascMetric: false,   teamReq: false, watchedReq: false,  def: true,  type: "range",          min: 0,     max: 100,   step: 1},
    {id: "eddelta",             name: "ED Δ",                   ascMetric: false,   teamReq: false, watchedReq: false,  def: false, type: "range",          min: -100,  max: 100,   step: 1},
    {id: "inguessrate",         name: "IN GR",                  ascMetric: false,   teamReq: false, watchedReq: false,  def: true,  type: "range",          min: 0,     max: 100,   step: 1},
    {id: "indelta",             name: "IN Δ",                   ascMetric: false,   teamReq: false, watchedReq: false,  def: false, type: "range",          min: -100,  max: 100,   step: 1},
    {id: "rigs",                name: "Rigs",                   ascMetric: false,   teamReq: false, watchedReq: true,   def: true,  type: "range",          min: 0,     max: 100,   step: 1},
    {id: "rigrate",             name: "Rig Rate",               ascMetric: false,   teamReq: false, watchedReq: true,   def: false, type: "range",          min: 0,     max: 100,   step: 1},
    {id: "solorigs",            name: "Solo Rigs",              ascMetric: false,   teamReq: false, watchedReq: true,   def: false, type: "range",          min: 0,     max: 100,   step: 1},
    {id: "solorigrate",         name: "Solo Rig Rate",          ascMetric: false,   teamReq: false, watchedReq: true,   def: false, type: "range",          min: 0,     max: 100,   step: 1},
    {id: "rigover8",            name: "Rig Over-8",             ascMetric: true,    teamReq: false, watchedReq: true,   def: false, type: "range",          min: 0,     max: 8,     step: 0.01},
    {id: "over8delta",          name: "Over-8 Δ",               ascMetric: false,   teamReq: false, watchedReq: true,   def: false, type: "range",          min: -8,    max: 8,     step: 1},
    {id: "rigguessrate",        name: "Rig GR",                 ascMetric: false,   teamReq: false, watchedReq: true,   def: true,  type: "range",          min: 0,     max: 100,   step: 1},
    {id: "offguessrate",        name: "Off GR",                 ascMetric: false,   teamReq: false, watchedReq: true,   def: true,  type: "range",          min: 0,     max: 100,   step: 1},
    {id: "rigdelta",            name: "Rig Δ",                  ascMetric: false,   teamReq: false, watchedReq: true,   def: false, type: "range",          min: -1000, max: 100,   step: 0.01},
    {id: "meandifficultyhit",   name: "Mean Difficulty Hit",    ascMetric: true,    teamReq: false, watchedReq: false,  def: false, type: "range",          min: 0,     max: 100,   step: 0.01},
    {id: "medianvintagehit",    name: "Median Vintage Hit",     ascMetric: false,   teamReq: false, watchedReq: false,  def: false, type: "range",          min: 1900,  max: 2026,  step: 1},
    {id: "mediantime",          name: "Median Time",            ascMetric: true,    teamReq: false, watchedReq: false,  def: false, type: "range",          min: 0,     max: 20,    step: 0.01},
    {id: "chantguessrate",      name: "Chant GR",               ascMetric: false,   teamReq: false, watchedReq: false,  def: false, type: "range",          min: 0,     max: 100,   step: 1}
];

let activePlayerHeadersConfig = playerHeadersMasterConfig.filter(col => {
    if (col.teamReq     && !use_teams)  return false;
    if (col.watchedReq  && !watched)    return false;

    if (["GR Δ", "UF Δ", "OP Δ", "ED Δ", "IN Δ"].includes(col.name)) {
        if (!players || players.length === 0) return false;

        const areAllRowsMissingValue = players.every(p => {
            const rawItem   = p[col.name];
            const parsedVal = (rawItem !== null && typeof rawItem === 'object') ? rawItem.count : rawItem;

            return parsedVal === undefined || parsedVal === null || parsedVal === "N/A" || isNaN(parseFloat(parsedVal));
        });

        if (areAllRowsMissingValue) return false;
    }

    return true;
});

activePlayerHeadersConfig.forEach(col => {col.visible = col.def;});

const availableColumnNames = new Set(activePlayerHeadersConfig.map(c => c.name));

if      (availableColumnNames.has("Score")) thickBorderColumns.add("Score");
else if (availableColumnNames.has("GR Δ"))  thickBorderColumns.add("GR Δ");
else                                        thickBorderColumns.add("GR");

if      (availableColumnNames.has("IN Δ"))  thickBorderColumns.add("IN Δ");
else if (availableColumnNames.has("IN GR")) thickBorderColumns.add("IN GR");
else if (availableColumnNames.has("ED Δ"))  thickBorderColumns.add("ED Δ");
else if (availableColumnNames.has("ED GR")) thickBorderColumns.add("ED GR");

function initPlayerColumnSettings() {
    const container = document.getElementById("playerColumnCheckboxContainer");
    const masterChk = document.getElementById("playerAllColumnsMasterCheckbox");

    if (!container || !masterChk) return;
    container.innerHTML = "";

    if (players && players.length > 0) {
        const uniqueTeams = new Set();

        players.forEach(p => {
            if (p["Team"]) {
                const tVal = (typeof p["Team"] === 'object') ? p["Team"].count : p["Team"];
                if (tVal) uniqueTeams.add(String(tVal).trim());
            }
        });

        const teamConfig = activePlayerHeadersConfig.find(c => c.id === "team");
        if (teamConfig) teamConfig.subOptions = Array.from(uniqueTeams).sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));

        activePlayerHeadersConfig.forEach(col => {
            if (col.type === "range") {
                const numericValues = players.map(p => {
                    const raw = p[col.name];
                    const val = (raw !== null && typeof raw === 'object') ? raw.count : raw;

                    if (col.id === "medianvintagehit" && typeof val === 'string') return parseVintageToFloat(val);
                    return parseFloat(val);
                }).filter(v => !isNaN(v) && v !== -Infinity && v !== Infinity);

                if (numericValues.length > 0) {
                    col.min = Math.floor(Math.min(...numericValues));
                    col.max = Math.ceil(Math.max(...numericValues));
                }
            }
        });
    }

    function updateMasterCheckboxState() {
        const allChecked        = activePlayerHeadersConfig.every(c => c.visible);
        const noneChecked       = activePlayerHeadersConfig.every(c => !c.visible);
        masterChk.checked       = allChecked;
        masterChk.className     = "rounded accent-black";
        masterChk.indeterminate = !allChecked && !noneChecked;
    }

    masterChk.addEventListener("change", () => {
        activePlayerHeadersConfig.forEach(c => { 
            c.visible = masterChk.checked;

            if (c.type === "categorical" && c.subOptions) {
                if (masterChk.checked)  c.selectedOptions = new Set(c.subOptions.map(o => o.toLowerCase()));
                else                    c.selectedOptions.clear();
            }
        });

        container.querySelectorAll("input[type='checkbox']").forEach(chk => chk.checked = masterChk.checked);
        triggerPlayerTableRefresh();
    });

    const hasDeltaData  = activePlayerHeadersConfig.some(c => ["GR Δ", "UF Δ", "OP Δ", "ED Δ", "IN Δ"].includes(c.name));
    const metricSection = document.getElementById("playerMetricModeSection");

    if (metricSection && hasDeltaData) metricSection.classList.remove("hidden");

    activePlayerHeadersConfig.forEach(col => {
        if (currentPlayerMetricMode === "Δ" && ["GR", "UF", "OP GR", "ED GR", "IN GR"]  .includes(col.name)) col.visible = false;
        if (currentPlayerMetricMode === "%" && ["GR Δ", "UF Δ", "OP Δ", "ED Δ", "IN Δ"] .includes(col.name)) col.visible = false;
    });

    activePlayerHeadersConfig.forEach(col => {
        const colWrapper        = document.createElement("div");
        colWrapper.className    = "flex flex-col space-y-1";

        const label     = document.createElement("label");
        label.className = "flex items-center gap-2 cursor-pointer w-full text-left font-bold text-black";

        const chk       = document.createElement("input");
        chk.type        = "checkbox";
        chk.className   = "player-col-toggle-checkbox rounded accent-black";
        chk.checked     = col.visible;

        chk.addEventListener("change", () => {
            col.visible = chk.checked;

            if (col.type === "categorical" && col.subOptions) {
                if (chk.checked) {
                    col.subOptions.forEach(opt => col.selectedOptions.add(opt.toLowerCase()));
                    colWrapper.querySelectorAll("input[type='checkbox']").forEach(subChk => subChk.checked = true);
                }

                else {
                    col.selectedOptions.clear();
                    colWrapper.querySelectorAll("input[type='checkbox']").forEach(subChk => subChk.checked = false);
                }
            }

            else if (!chk.checked && col.type === "range") {
                col.currentMin = col.min;
                col.currentMax = col.max;

                const inputs = colWrapper.querySelectorAll("input[type='number']");

                if (inputs.length === 2) {
                    inputs[0].value = col.min;
                    inputs[1].value = col.max;
                }
            }

            updateMasterCheckboxState();
            triggerPlayerTableRefresh();
        });

        label.appendChild(chk);
        label.appendChild(document.createTextNode(col.name));

        if (colExplanations[col.name]) {
            const explanationIndicator      = document.createElement("span");
            explanationIndicator.className  = "-ml-1.5 text-black cursor-help select-none font-normal text-base text-bold has-explanation";
            explanationIndicator.setAttribute("data-metric", col.name);
            explanationIndicator.innerHTML  = "🛈";
            label.appendChild(explanationIndicator);
        }

        colWrapper.appendChild(label);

        if (col.type === "categorical" && col.subOptions) {
            const subContainer      = document.createElement("div");
            subContainer.className  = "pl-6 flex flex-col text-xs";

            if (col.visible)    col.selectedOptions = new Set(col.subOptions.map(o => o.toLowerCase()));
            else                col.selectedOptions = new Set();

            col.subOptions.forEach(opt => {
                const subLabel      = document.createElement("label");
                subLabel.className  = "flex items-center gap-1 cursor-pointer text-gray-700 hover:text-black py-0.5 pr-6";

                const subChk        = document.createElement("input");
                subChk.type         = "checkbox";
                subChk.className    = "rounded accent-black scale-90";
                subChk.checked      = col.visible;

                subChk.addEventListener("change", () => {
                    if (subChk.checked) col.selectedOptions.add(opt.toLowerCase());
                    else                col.selectedOptions.delete(opt.toLowerCase());

                    const totalChildren = col.subOptions.length;
                    const selectedCount = col.selectedOptions.size;

                    if (selectedCount === totalChildren) {
                        chk.checked         = true;
                        chk.indeterminate   = false;
                        col.visible         = true;
                    }

                    else if (selectedCount === 0) {
                        chk.checked         = false;
                        chk.indeterminate   = false;
                        col.visible         = false;
                    }

                    else {
                        chk.checked         = false;
                        chk.indeterminate   = true;
                        col.visible         = true;
                    }

                    updateMasterCheckboxState();
                    triggerPlayerTableRefresh();
                });

                subLabel        .appendChild(subChk);
                subLabel        .appendChild(document.createTextNode(opt));
                subContainer    .appendChild(subLabel);
            });

            colWrapper.appendChild(subContainer);
        }

        if (col.type === "range") {
            col.currentMin = col.min;
            col.currentMax = col.max;

            const inputContainer        = document.createElement("div");
            inputContainer.className    = "pl-6 flex flex-col gap-1 mt-1 w-full text-black";

            const boxWrapper        = document.createElement("div");
            boxWrapper.className    = "flex flex-col gap-1";

            const minLabel      = document.createElement("label");
            minLabel.className  = "flex items-center justify-start gap-2 font-mono";
            minLabel.innerHTML  = "Min:";
            const inputMin      = document.createElement("input");
            inputMin.type       = "number";
            inputMin.className  = "w-10 h-5 border text-center text-xs";
            inputMin.value      = col.min;

            const maxLabel      = document.createElement("label");
            maxLabel.className  = "flex items-center justify-start gap-2 font-mono";
            maxLabel.innerHTML  = "Max:";
            const inputMax      = document.createElement("input");
            inputMax.type       = "number";
            inputMax.className  = "w-10 h-5 border text-center text-xs";
            inputMax.value      = col.max;

            minLabel        .appendChild(inputMin);
            maxLabel        .appendChild(inputMax);
            boxWrapper      .appendChild(minLabel);
            boxWrapper      .appendChild(maxLabel);
            inputContainer  .appendChild(boxWrapper);

            const handleTextbookInput = () => {
                let valMin = parseFloat(inputMin.value);
                let valMax = parseFloat(inputMax.value);

                if (isNaN(valMin)) valMin = col.min;
                if (isNaN(valMax)) valMax = col.max;

                col.currentMin = valMin;
                col.currentMax = valMax;
            };

            const triggerRefreshDebounced = debounce(() => {
                if (!col.visible && (col.currentMin > col.min || col.currentMax < col.max)) {
                    col.visible = true;
                    chk.checked = true;

                    updateMasterCheckboxState();
                }

                triggerPlayerTableRefresh();
            }, 150);

            inputMin.addEventListener("input", () => { handleTextbookInput(); triggerRefreshDebounced(); });
            inputMax.addEventListener("input", () => { handleTextbookInput(); triggerRefreshDebounced(); });

            masterChk.addEventListener("change", () => {
                if (!masterChk.checked) {
                    inputMin.value = col.min;
                    inputMax.value = col.max;
                    col.currentMin = col.min;
                    col.currentMax = col.max;
                }

                else {
                    col.currentMin = isNaN(parseFloat(inputMin.value)) ? col.min : parseFloat(inputMin.value);
                    col.currentMax = isNaN(parseFloat(inputMax.value)) ? col.max : parseFloat(inputMax.value);
                }

                triggerPlayerTableRefresh();
            });

            colWrapper.appendChild(inputContainer);
        }

        container.appendChild(colWrapper);
    });

    updateMasterCheckboxState   ();
    setupTooltipListeners       ();
}

function triggerPlayerTableRefresh() {
    const playerSearchInput = document.getElementById('playerSearchInput');

    if (playerSearchInput && playerSearchInput.value.trim()) playerSearchInput.dispatchEvent(new Event('input'));

    else {
        globalFilteredPlayers = processPlayerRuntimeSettings(players || []);
        sortAndRenderPlayers();
    }
}

function processPlayerRuntimeSettings(targetPool) {
    return targetPool.filter(p => {
        for (let col of activePlayerHeadersConfig) {
            const rawVal    = p[col.name];
            let targetVal   = (rawVal !== null && typeof rawVal === 'object') ? rawVal.count : rawVal;

            if (col.type === "categorical" && col.selectedOptions) {
                if (!col.visible) continue;
                const matchString = String(targetVal || "").trim().toLowerCase();
                if (matchString && !col.selectedOptions.has(matchString)) return false;
            }

            if (col.type === "range" && col.currentMin !== undefined && col.currentMax !== undefined) {
                let checkNum = parseFloat(targetVal);
                if (col.id === "medianvintagehit" && typeof targetVal === 'string') checkNum = parseVintageToFloat(targetVal);
                if (isNaN(checkNum))                                                continue; 
                if (checkNum < col.currentMin || checkNum > col.currentMax)         return false;
            }
        }

        return true;
    });
}

function cacheGlobalHighlights() {
    if (!players || !hlRules) return;
    globalMetricHighlights = {};

    Object.keys(hlRules).forEach(metricName => {
        const rule = hlRules[metricName];

        globalMetricHighlights[metricName] = {
            bestPlayerName  : (rule.best_idx    !== undefined && players[rule.best_idx])    ? players[rule.best_idx]    ["Player"] : null,
            worstPlayerName : (rule.worst_idx   !== undefined && players[rule.worst_idx])   ? players[rule.worst_idx]   ["Player"] : null
        };
    });
}

function evaluatePlayerConstraint(playerRow, key, operator, value) {
    const aliasMap = {
        "gr"            : "guessrate",
        "usefulness"    : "uf",
        "solos"         : "18s",
        "doubles"       : "28s",
        "sevens"        : "78s",
        "opgr"          : "opguessrate",
        "edgr"          : "edguessrate",
        "ingr"          : "inguessrate",
        "riggr"         : "rigguessrate",
        "offgr"         : "offguessrate",
        "difficulty"    : "meandifficultyhit",
        "vintage"       : "medianvintagehit",
        "chantgr"       : "chantguessrate"
    };

    const lookupKey     = aliasMap[key] || key;
    const matchedHeader = activePlayerHeadersConfig.find(h => h.id === lookupKey);

    if (!matchedHeader) return false;

    let targetRaw = playerRow[matchedHeader.name];
    let targetVal = (targetRaw !== null && typeof targetRaw === 'object') ? targetRaw.count : targetRaw;

    if (key === "player" || key === "team") {
        const stringLower = String(targetVal || "").toLowerCase();
        return stringLower.includes(value.toLowerCase());
    }

    let parsedTarget = parseFloat(targetVal);
    let parsedCrit   = parseFloat(value);
    
    if (key === "vintage" && isNaN(parsedCrit)) {
        parsedTarget = parseVintageToFloat(String(targetVal));
        parsedCrit   = parseVintageToFloat(value);
    }

    if (isNaN(parsedTarget) || isNaN(parsedCrit))   return false;
    if (operator === ":"    || operator === "=")    return parsedTarget === parsedCrit;
    if (operator === "<")                           return parsedTarget <   parsedCrit;
    if (operator === ">")                           return parsedTarget >   parsedCrit;
    if (operator === "<=")                          return parsedTarget <=  parsedCrit;
    if (operator === ">=")                          return parsedTarget >=  parsedCrit;
    if (operator === "!="   || operator === "!:")   return parsedTarget !== parsedCrit;

    return false;
}

function processPlayerFiltering(rawQuery) {
    if (!rawQuery) return [...players];

    const tokens        = [];
    const tokenRegex    = /\(|\)|or\b|and\b|[a-zA-Z0-9_/-]+(?:<=|>=|!=|!:|[:<>==])"[^"]*"|[^\s"()]+|"[^"]*"/gi;

    let match;
    while ((match = tokenRegex.exec(rawQuery)) !== null) tokens.push(match[0]);

    const outputQueue   = [];
    const operatorStack = [];
    const precedence    = {'or': 1, 'and': 2};
    let expectOperator  = false;

    tokens.forEach(token => {
        const lowerToken = token.toLowerCase();

        if (expectOperator && lowerToken !== 'and' && lowerToken !== 'or' && lowerToken !== ')') {
            while (operatorStack.length && precedence[operatorStack[operatorStack.length - 1]] >= precedence['and']) outputQueue.push(operatorStack.pop());
            operatorStack.push('and');
        }

        if (lowerToken === 'and' || lowerToken === 'or') {
            while (operatorStack.length && precedence[operatorStack[operatorStack.length - 1]] >= precedence[lowerToken]) outputQueue.push(operatorStack.pop());
            operatorStack.push(lowerToken);
            expectOperator = false;
        }

        else if (token === '(') {
            operatorStack.push(token);
            expectOperator = false;
        }

        else if (token === ')') {
            while (operatorStack.length && operatorStack[operatorStack.length - 1] !== '(') outputQueue.push(operatorStack.pop());
            operatorStack.pop();
            expectOperator = true;
        }

        else {
            outputQueue.push(token);
            expectOperator = true;
        }
    });

    while (operatorStack.length) outputQueue.push(operatorStack.pop());

    return players.filter(pRow => {
        if (outputQueue.length === 0) return true;
        const evalStack = [];

        for (let token of outputQueue) {
            const lowerToken = typeof token === 'string' ? token.toLowerCase() : '';

            if (lowerToken === 'and') {
                const b = evalStack.pop();
                const a = evalStack.pop();
                evalStack.push(a && b);
            }

            else if (lowerToken === 'or') {
                const b = evalStack.pop();
                const a = evalStack.pop();
                evalStack.push(a || b);
            }

            else {
                const exprRegex = /^([a-zA-Z0-9_/-]+)(<=|>=|!=|!:|[:<>==])(.+)$/;
                const matchExpr = token.match(exprRegex);

                if (matchExpr) {
                    const key       = matchExpr[1].toLowerCase();
                    const op        = matchExpr[2];
                    let cleanVal    = matchExpr[3].trim();

                    if (cleanVal.startsWith('"') && cleanVal.endsWith('"')) cleanVal = cleanVal.slice(1, -1).trim();
                    else                                                    cleanVal = cleanVal.replace(/^"|"$/g, '').trim();

                    evalStack.push(evaluatePlayerConstraint(pRow, key, op, cleanVal));
                }

                else {
                    const cleanWord     = token.replace(/^"|"$/g, '')   .toLowerCase();
                    const nameString    = String(pRow["Player"] || "")  .toLowerCase();
                    evalStack.push(nameString.includes(cleanWord));
                }
            }
        }

        return evalStack[0];
    });
}

function sortAndRenderPlayers() {
    const activeCols    = activePlayerHeadersConfig.filter(c => c.visible);
    const table         = document.getElementById('playerStandingsTable');
    const counterNode   = document.getElementById('playerSearchCounter');

    if (!table) return;

    const currentSortField          = globalPlayerSortState.columnName;
    const configMatch               = activePlayerHeadersConfig.find(h => h.name === currentSortField);
    const sortingDirectionModifier  = globalPlayerSortState.ascending ? 1 : -1;

    globalFilteredPlayers.sort((a, b) => {
        let valA = a[currentSortField];
        let valB = b[currentSortField];

        let extractA = (valA !== null && typeof valA === 'object') ? valA.count : valA;
        let extractB = (valB !== null && typeof valB === 'object') ? valB.count : valB;

        if (extractA === undefined || extractA === null) return 1;
        if (extractB === undefined || extractB === null) return -1;

        if (currentSortField === "Player" || currentSortField === "Team") return sortingDirectionModifier * String(extractA).localeCompare(String(extractB));

        let numA = typeof extractA === 'string' ? parseFloat(extractA.replace(/[^0-9.-]/g, '')) : extractA;
        let numB = typeof extractB === 'string' ? parseFloat(extractB.replace(/[^0-9.-]/g, '')) : extractB;

        if (isNaN(numA)) return 1;
        if (isNaN(numB)) return -1;

        return (numA < numB ? -1 : numA > numB ? 1 : 0) * sortingDirectionModifier;
    });

    if (counterNode) counterNode.innerText = `${globalFilteredPlayers.length}/${players.length}`;

    if (activeCols.length === 0) {
        table.innerHTML = `<thead><tr><th>Error</th></tr></thead><tbody><tr><td class="p-2 text-center text-black">Select at least 1 column</td></tr></tbody>`;
        return;
    }

    let thead = "<thead><tr>" + activeCols.map(h => {
        let classes = [];

        if (thickBorderColumns.has(h.name)) classes.push("border-col-group");
        if (colExplanations[h.name])        classes.push("has-explanation");

        classes.push("cursor-pointer select-none");

        const isCurrentSort         = (globalPlayerSortState.columnName === h.name);
        const directionIndicator    = isCurrentSort ? (globalPlayerSortState.ascending ? "▴" : "▾")     : "▸";
        const blackBgStyle          = isCurrentSort ? ' style="background-color: black; color: white;"' : '';
        const lineBrokenName        = h.name.replace(/ /g, '<br>');

        return `<th class="${classes.join(' ')}"${blackBgStyle} data-player-metric="${h.name}" data-metric="${h.name}">${lineBrokenName}${directionIndicator}</th>`;
    }).join('') + "</tr></thead>";

    let tbody = "<tbody>";

    globalFilteredPlayers.forEach((row, idx) => {
        let groupLine = groupBorders.includes(idx) ? " border-group-line" : "";
        tbody += `<tr class="${groupLine}">`;

        activeCols.forEach(h => {
            let rawCell     = row[h.name];
            let displayVal  = (rawCell !== null && typeof rawCell === 'object') ? rawCell.count : rawCell;
            let cellStyle   = thickBorderColumns.has(h.name) ? "border-col-group " : "";

            if (globalMetricHighlights[h.name]) {
                const currentPlayerName = row["Player"];
                const isBest            = (globalMetricHighlights[h.name].bestPlayerName    === currentPlayerName);
                const isWorst           = (globalMetricHighlights[h.name].worstPlayerName   === currentPlayerName);

                if (h.name !== "Team" && h.name !== "Tier" && isBest)   cellStyle += "highlight-best ";
                if (h.name !== "Team" && h.name !== "Tier" && isWorst)  cellStyle += "highlight-worst ";
            }

            let intCols = ["Tier", "1/8s", "2/8s", "7/8s", "Lives Taken", "Lives Saved", "Rigs", "Solo Rigs"];
            let formattedVal;

            if      (h.name === "Median Vintage Hit")                               formattedVal = parseFloatToVintage(displayVal);
            else if (typeof displayVal === 'number' && !intCols.includes(h.name))   formattedVal = displayVal.toFixed(2);
            else                                                                    formattedVal = displayVal;

            let finalVal        = (h.name === "Player") ? `<b>${formattedVal}</b>` : formattedVal;
            let clickHandler    = "";

            const rawPlayerVal      = row["Player"];
            const currentPlayerName = String((rawPlayerVal !== null && typeof rawPlayerVal === 'object') ? rawPlayerVal.count : rawPlayerVal).toLowerCase();

            if (parseFloat(displayVal) > 0) {
                if      (h.name === "GR")               clickHandler = ` onclick="searchPlayerMetricFromTable('correct:${currentPlayerName}')"`;
                else if (h.name === "1/8s")             clickHandler = ` onclick="searchPlayerMetricFromTable('correct:${currentPlayerName} correct:1')"`;
                else if (h.name === "2/8s")             clickHandler = ` onclick="searchPlayerMetricFromTable('correct:${currentPlayerName} correct:2')"`;
                else if (h.name === "7/8s")             clickHandler = ` onclick="searchPlayerMetricFromTable('correct!:${currentPlayerName} correct:7')"`;
                else if (h.name === "Lives Taken")      clickHandler = ` onclick="searchPlayerMetricFromTable('lifetaken:${currentPlayerName}')"`;
                else if (h.name === "Lives Saved")      clickHandler = ` onclick="searchPlayerMetricFromTable('lifesaved:${currentPlayerName}')"`;
                else if (h.name === "OP GR")            clickHandler = ` onclick="searchPlayerMetricFromTable('correct:${currentPlayerName} songtype:op')"`;
                else if (h.name === "ED GR")            clickHandler = ` onclick="searchPlayerMetricFromTable('correct:${currentPlayerName} songtype:ed')"`;
                else if (h.name === "IN GR")            clickHandler = ` onclick="searchPlayerMetricFromTable('correct:${currentPlayerName} songtype:in')"`;
                else if (h.name === "Solo Rigs")        clickHandler = ` onclick="searchPlayerMetricFromTable('list:${currentPlayerName} list:1')"`;
                else if (h.name === "Rig GR")           clickHandler = ` onclick="searchPlayerMetricFromTable('list:${currentPlayerName} correct:${currentPlayerName}')"`;
                else if (h.name === "Off GR")           clickHandler = ` onclick="searchPlayerMetricFromTable('list!:${currentPlayerName} correct:${currentPlayerName}')"`;
                else if (h.name === "Chant GR")         clickHandler = ` onclick="searchPlayerMetricFromTable('correct:${currentPlayerName} chanting:yes')"`;
            }

            if (h.name !== "GR" && ((h.name === "Player" && rawCell && rawCell.details && rawCell.details.length > 0) || 
                (rawCell !== null && typeof rawCell === 'object' && rawCell.details && rawCell.details.length > 0))) {
                let encodedDetails = encodeURIComponent(JSON.stringify(rawCell.details));
                tbody += `<td class="${cellStyle.trim()}" data-songs="${encodedDetails}" data-metric="${h.name}"${clickHandler}>${finalVal}</td>`;
            }

            else tbody += `<td class="${cellStyle.trim()}"${clickHandler}>${finalVal}</td>`;
        });

        tbody += "</tr>";
    });

    table.innerHTML = thead + tbody + "</tbody>";

    table.querySelectorAll('thead th').forEach(th => {
        th.addEventListener('click', () => {
            const metricName = th.getAttribute('data-player-metric');

            if (globalPlayerSortState.columnName === metricName) globalPlayerSortState.ascending = !globalPlayerSortState.ascending;

            else {
                globalPlayerSortState.columnName = metricName;
                const matchObj = activePlayerHeadersConfig.find(m => m.name === metricName);
                globalPlayerSortState.ascending = matchObj ? matchObj.ascMetric : false;
            }

            sortAndRenderPlayers();
        });
    });

    setupTooltipListeners();
}

function renderPlayerTable() {
    cacheGlobalHighlights       ();
    initPlayerColumnSettings    ();

    const playerSearchInput = document.getElementById('playerSearchInput');

    const triggerPlayerQueryProcess = () => {
        const rawQuery          = playerSearchInput ? playerSearchInput.value.trim() : "";
        let textFiltered        = processPlayerFiltering(rawQuery);
        globalFilteredPlayers   = processPlayerRuntimeSettings(textFiltered);

        sortAndRenderPlayers();
    };

    if (playerSearchInput) playerSearchInput.addEventListener('input', debounce(triggerPlayerQueryProcess, 250));
    globalFilteredPlayers = processPlayerRuntimeSettings([...players]);
    sortAndRenderPlayers();
}

function renderTourTable() {
    const table = document.getElementById('tourStatsTable');
    if (!tourStats || !tourStats.length) return;

    const half          = Math.ceil(tourStats.length / 2);
    const leftSlice     = tourStats.slice(0, half);
    const rightSlice    = tourStats.slice(half);

    let thead = `
        <thead>
            <tr>
                <th class="border-col-group">Metric</th>
                <th class="border-col-group">Value</th>
                <th class="border-col-group">Metric</th>
                <th>Value</th>
            </tr>
        </thead>`;

    let tbody = "<tbody>";

    const getTourClickHandler = (metric, displayVal, encodedDetails) => {
        const key           = metric.trim();
        const fractionMatch = key.match(/^Total (\d)\/8s$/);

        if (fractionMatch)                                  return ` onclick="searchTourFraction(${fractionMatch[1]})"`;
        if (key === "Most Popular Genre"    && displayVal)  return ` onclick="searchTourMetadata('genre',   '${displayVal}')"`;
        if (key === "Most Popular Tag"      && displayVal)  return ` onclick="searchTourMetadata('tag',     '${displayVal}')"`;
        if (key === "Total 4-0s")                           return ` onclick="searchPlayerMetricFromTable('sweep:yes')"`;

        const mostFractionMatch = key.match(/^Most (\d)\/8s$/);
        if (mostFractionMatch) return ` onclick="sortPlayerColumnFromTour('${mostFractionMatch[1]}/8s', false)"`;

        if (key === "Highest GR Without 1/8s")                                                              return ` onclick="searchPlayerFilterFromTour    ('solos=0', 'GR', false)"`;
        if (key === "Lowest GR Without 1/8s")                                                               return ` onclick="searchPlayerFilterFromTour    ('solos=0', 'GR', true)"`;
        if (key === "Highest GR With 1/8s")                                                                 return ` onclick="searchPlayerFilterFromTour    ('solos>0', 'GR', false)"`;
        if (key === "Lowest GR With 1/8s")                                                                  return ` onclick="searchPlayerFilterFromTour    ('solos>0', 'GR', true)"`;
        if ((key === "Best Solo Rig Converter" || key === "Worst Solo Rig Converter") && encodedDetails)    return ` onclick="searchSoloRigConverter        (this)"`;

        return "";
    };

    for (let i = 0; i < half; i++) {
        tbody += "<tr>";
        const leftRow = leftSlice[i];

        if (leftRow) {
            let rawCell         = leftRow.Value;
            let displayVal      = (rawCell !== null && typeof rawCell === 'object') ? rawCell.count : rawCell;
            let hasExp          = !!colExplanations[leftRow.Metric];
            let metricClass     = hasExp ? "border-col-group has-explanation" : "border-col-group";
            let metricAttr      = `class='${metricClass}' data-metric="${leftRow.Metric}"`;
            let encodedDetails  = (rawCell !== null && typeof rawCell === 'object' && rawCell.details) ? encodeURIComponent(JSON.stringify(rawCell.details)) : "";
            let clickHandler    = getTourClickHandler(leftRow.Metric, displayVal, encodedDetails);

            if (encodedDetails && rawCell.details.length > 0)   tbody += `<td ${metricAttr}><b>${leftRow.Metric}</b></td><td class="border-col-group" data-songs="${encodedDetails}"${clickHandler}>${displayVal}</td>`;
            else                                                tbody += `<td ${metricAttr}><b>${leftRow.Metric}</b></td><td class="border-col-group"${clickHandler}>${displayVal}</td>`;
        }

        else tbody += `<td class="border-col-group"></td><td class="border-col-group"></td>`;

        const rightRow = rightSlice[i];

        if (rightRow) {
            let rawCell         = rightRow.Value;
            let displayVal      = (rawCell !== null && typeof rawCell === 'object') ? rawCell.count : rawCell;
            let hasExp          = !!colExplanations[rightRow.Metric];
            let metricClass     = hasExp ? "border-col-group has-explanation" : "border-col-group";
            let metricAttr      = `class='${metricClass}' data-metric="${rightRow.Metric}"`;
            let encodedDetails  = (rawCell !== null && typeof rawCell === 'object' && rawCell.details) ? encodeURIComponent(JSON.stringify(rawCell.details)) : "";
            let clickHandler    = getTourClickHandler(rightRow.Metric, displayVal, encodedDetails);

            if (encodedDetails && rawCell.details.length > 0)   tbody += `<td ${metricAttr}><b>${rightRow.Metric}</b></td><td data-songs="${encodedDetails}"${clickHandler}>${displayVal}</td>`;
            else                                                tbody += `<td ${metricAttr}><b>${rightRow.Metric}</b></td><td${clickHandler}>${displayVal}</td>`;
        }

        else tbody += `<td class="border-col-group"></td><td></td>`;

        tbody += "</tr>";
    }

    table.innerHTML = thead + tbody + "</tbody>";
}

function renderTeamTable() {
    const table     = document.getElementById('teamStatsTable');
    const spacer    = document.getElementById('teamTableSpacer');

    if (!table) return;

    if (!use_teams || !watched || !teamStats || !teamStats.length) {
        table.innerHTML                     = "";
        if (spacer) spacer.style.display    = "none";
        return;
    }

    let headers = Object.keys(teamStats[0]);

    let thead = "<thead><tr>" + headers.map(h => {
        let classes = [];

        if (thickBorderColumns.has(h))  classes.push("border-col-group");
        if (colExplanations[h])         classes.push("has-explanation");

        let classStr = classes.length > 0 ? ` class="${classes.join(' ')}"` : '';
        return `<th${classStr} data-metric="${h}">${h.replace(/ /g, '<br>')}</th>`;
    }).join('') + "</tr></thead>";

    let tbody = "<tbody>";

    teamStats.forEach((row, idx) => {
        tbody += "<tr>";

        headers.forEach(h => {
            let rawCell     = row[h];
            let displayVal  = (rawCell !== null && typeof rawCell === 'object') ? rawCell.count : rawCell;
            let cellStyle   = thickBorderColumns.has(h) ? "border-col-group " : "";

            if (teamHlRules[h]) {
                let isBest  = (teamHlRules[h].best_idx  === idx);
                let isWorst = (teamHlRules[h].worst_idx === idx);

                if      (isBest)    cellStyle += "highlight-best ";
                else if (isWorst)   cellStyle += "highlight-worst ";
            }

            let finalVal = displayVal;

            if      (h === "Win Rate" && typeof displayVal === 'number')    finalVal = isNaN(displayVal) ? "N/A" : `${displayVal.toFixed(2)}`;
            else if (h === "Team Leader")                                   finalVal = `<b>${displayVal}</b>`;

            if (rawCell !== null && typeof rawCell === 'object' && rawCell.details && rawCell.details.length > 0) {
                let encodedDetails  = encodeURIComponent(JSON.stringify(rawCell.details));
                let clickHandler    = "";

                if (h === "Total 1/8s") clickHandler = ` onclick="searchTeamSolos('${row["Team Leader"]}')"`;
                tbody += `<td class="${cellStyle.trim()}" data-songs="${encodedDetails}"${clickHandler}>${finalVal}</td>`;
            }

            else tbody += `<td class="${cellStyle.trim()}">${finalVal}</td>`;
        });

        tbody += "</tr>";
    });

    table.innerHTML = thead + tbody + "</tbody>";
}

window.updatePlayerMetricModeFromRadio = function(selectedMode) {
    currentPlayerMetricMode = selectedMode;

    activePlayerHeadersConfig.forEach(col => {
        if      (currentPlayerMetricMode === "Δ" && ["GR", "UF", "OP GR", "ED GR", "IN GR"]     .includes(col.name)) col.visible = false;
        else if (currentPlayerMetricMode === "%" && ["GR Δ", "UF Δ", "OP Δ", "ED Δ", "IN Δ"]    .includes(col.name)) col.visible = false;
        else if (currentPlayerMetricMode === "Δ" && ["GR Δ", "UF Δ", "OP Δ", "ED Δ", "IN Δ"]    .includes(col.name)) col.visible = true;
        else if (currentPlayerMetricMode === "%" && ["GR", "UF", "OP GR", "ED GR", "IN GR"]     .includes(col.name)) col.visible = col.def;
    });

    initPlayerColumnSettings    ();
    triggerPlayerTableRefresh   ();
};

function initPlayerRadioSettingsListeners() {
    document.querySelectorAll('input[name="playerMetricModeRadio"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            window.updatePlayerMetricModeFromRadio(e.target.value);
        });
    });
}

initPlayerRadioSettingsListeners();

function syncTierDropdownDOMState() {
    const isCount = globalChartMode === "COUNT";

    document.getElementById("opt_c1_base")      .innerText = isCount ? "Corrects"       : "Guess Rate";
    document.getElementById("opt_c1_over8")     .innerText = isCount ? "Over-8 Hit"     : "Over-8 Distribution";
    document.getElementById("opt_c1_rig")       .innerText = isCount ? "Rigs"           : "Rig Rate";
    document.getElementById("opt_c1_chant")     .innerText = isCount ? "Chanting Hit"   : "Chanting Guess Rate";
    document.getElementById("label_group_c1")   .innerText = "General";
    document.getElementById("label_group_c2")   .innerText = "Contribution";

    const hitLabel = document.getElementById("label_opt_c1_hit");
    const offLabel = document.getElementById("label_opt_c1_off");

    if (isCount) {
        if (hitLabel) hitLabel.classList.add("hidden");
        if (offLabel) offLabel.classList.add("hidden");

        if (c1Sub === "HIT" || c1Sub === "OFF") {
            c1Sub = "BASE";
            document.querySelector('input[name="tierSubMetricsRadio"][value="BASE"]').checked = true;
        }
    }

    else {
        if (hitLabel) hitLabel.classList.remove("hidden");
        if (offLabel) offLabel.classList.remove("hidden");
    }
}

function initTierDropdownListeners() {
    document.querySelectorAll('input[name="tierSortRadio"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            currentTierChartMode = e.target.value;
            renderTierCharts();
        });
    });

    document.querySelectorAll('input[name="tierDisplayRadio"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            globalChartMode = e.target.value;
            syncTierDropdownDOMState();
            renderTierCharts();
        });
    });

    document.querySelectorAll('input[name="tierChartGroupRadio"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            const selectedGroup = e.target.value;

            if (selectedGroup !== 'C1') document.querySelectorAll('input[name="tierSubMetricsRadio"]').forEach(r => r.checked = false);
            else {
                const checkedSub = document.querySelector('input[name="tierSubMetricsRadio"]:checked');

                if (!checkedSub) {
                    const defaultSub = document.querySelector('input[name="tierSubMetricsRadio"][value="BASE"]');

                    if (defaultSub) {
                        defaultSub.checked  = true;
                        c1Sub               = "BASE";
                    }
                }
            }

            if (selectedGroup !== 'C3') document.querySelectorAll('input[name="tierTimeRadio"]').forEach(r => r.checked = false);
            else {
                const checkedTime = document.querySelector('input[name="tierTimeRadio"]:checked');

                if (!checkedTime) {
                    const defaultTime = document.querySelector('input[name="tierTimeRadio"][value="MED"]');

                    if (defaultTime) {
                        defaultTime.checked = true;
                        c3Mode              = "MED";
                    }
                }
            }

            renderTierCharts();
        });
    });

    document.querySelectorAll('input[name="tierSubMetricsRadio"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            c1Sub = e.target.value;

            document.querySelector      ('input[name="tierChartGroupRadio"][value="C1"]').checked = true;
            document.querySelectorAll   ('input[name="tierTimeRadio"]').forEach(r => r.checked = false);

            renderTierCharts();
        });
    });

    document.querySelectorAll('input[name="tierTimeRadio"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            c3Mode = e.target.value;

            document.querySelector      ('input[name="tierChartGroupRadio"][value="C3"]').checked = true;
            document.querySelectorAll   ('input[name="tierSubMetricsRadio"]').forEach(r => r.checked = false);

            renderTierCharts();
        });
    });
}

initTierDropdownListeners();

function initGuessDropdownListeners() {
    document.querySelectorAll('input[name="guessDisplayRadio"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            const val           = e.target.value;
            const guessChart    = document.getElementById("guessChartContainer");
            const listChart     = document.getElementById("listChartContainer");

            if (val === "ALL") {
                currentGuessListViewMode = "ALL";

                if (guessChart) guessChart  .classList.remove   ("hidden");
                if (listChart)  listChart   .classList.add      ("hidden");

                window.dispatchEvent(new Event('resize'));
                updateGuessChartAxesFocus();
            } 

            else if (val === "RIG") {
                currentGuessListViewMode    = "RIG";
                currentListChartMode        = "ALL";

                if (guessChart) guessChart  .classList.add      ("hidden");
                if (listChart)  listChart   .classList.remove   ("hidden");

                renderListChart();
            } 

            else if (val === "HIT") {
                currentGuessListViewMode    = "HIT";
                currentListChartMode        = "HIT";

                if (guessChart) guessChart  .classList.add      ("hidden");
                if (listChart)  listChart   .classList.remove   ("hidden");
                renderListChart();
            }

            updateGuessHelpDropdown();
        });
    });

    document.querySelectorAll('input[name="guessFocusRadio"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            isGraphFocused = (e.target.value === "ON");

            if (currentGuessListViewMode === "ALL") updateGuessChartAxesFocus   ();
            else                                    renderListChart             ();
        });
    });
}

initGuessDropdownListeners();

function generateLinearColors(startHex, endHex, steps) {
    const parse = (hex) => [parseInt(hex.slice(1,3),16), parseInt(hex.slice(3,5),16), parseInt(hex.slice(5,7),16)];
    const cA    = parse(startHex), cB = parse(endHex);
    if (steps <= 1) return [startHex];
    const arr   = [];

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

function renderTierCharts() {
    if ((!document.getElementById('tierChart_MainMetrics') && !document.getElementById('tierChart_MainMetricsMain')) || !tierStats) return;
    if (!globalSearchData || globalSearchData.length === 0)                                                                         return;

    let checkedGroup = document.querySelector('input[name="tierChartGroupRadio"]:checked')?.value;

    if (!checkedGroup) {
        if (document.querySelector('input[name="tierSubMetricsRadio"]:checked')) {
            checkedGroup                                                                    = "C1";
            document.querySelector('input[name="tierChartGroupRadio"][value="C1"]').checked = true;
        }

        else if (document.querySelector('input[name="tierTimeRadio"]:checked')) {
            checkedGroup                                                                    = "C3";
            document.querySelector('input[name="tierChartGroupRadio"][value="C3"]').checked = true;
        }

        else {
            checkedGroup                                                                            = "C1";
            document.querySelector('input[name="tierChartGroupRadio"][value="C1"]')     .checked    = true;
            document.querySelector('input[name="tierSubMetricsRadio"][value="BASE"]')   .checked    = true;
            c1Sub = "BASE";
        }
    }

    const containerC1 = document.getElementById("container_tierChart_MainMetrics");
    const containerC2 = document.getElementById("container_tierChart_LivesMetrics");
    const containerC3 = document.getElementById("container_tierChart_TimeMetrics");

    if (containerC1) containerC1.classList.add("hidden");
    if (containerC2) containerC2.classList.add("hidden");
    if (containerC3) containerC3.classList.add("hidden");

    if (checkedGroup === "C1" && containerC1) containerC1.classList.remove("hidden");
    if (checkedGroup === "C2" && containerC2) containerC2.classList.remove("hidden");
    if (checkedGroup === "C3" && containerC3) containerC3.classList.remove("hidden");

    let gapCounter      = 0;
    const hasChanting   = globalSearchData.some(s => s.chanting && s.chanting.toLowerCase() === "yes");

    const getPlayerStringName = (pField) => {
        if (!pField) return "";
        return typeof pField === 'object' ? String(pField.count || "") : String(pField);
    };

    const truncateString = (str, maxLen = 40) => {
        if (!str) return "";
        if (str.length <= maxLen) return str;
        let cutStr = str.slice(0, maxLen);
        const lastSpace = cutStr.lastIndexOf(" ");
        if (lastSpace > 0) cutStr = cutStr.slice(0, lastSpace);
        return cutStr + " ...";
    };

    const compilePlayerStatsFromSearch = (playerName) => {
        const pClean = playerName.replace(/[★▲▼]/g, "").trim().toLowerCase();

        let totalSeen           = 0;
        let totalCorrect        = 0;
        let x8Counts            = Array(8).fill(0);
        let totalRigs           = 0;
        let rigHits             = 0;
        let offListHits         = 0;
        let totalChantSeen      = 0;
        let totalChantCorrect   = 0;
        let livesTaken          = 0;
        let livesSaved          = 0;

        let allSeenSongs        = [];
        let correctsList        = [];
        let rigsList            = [];
        let rigHitsList         = [];
        let offListHitsList     = [];
        let chantSeenList       = [];
        let chantCorrectsList   = [];
        let livesTakenList      = [];
        let livesSavedList      = [];
        let otherCorrectsList   = [];
        let x8Lists             = Array.from({length: 8}, () => []);

        globalSearchData.forEach(song => {
            const roomPlayers   = (song.room_players        || []).map(p => p.toLowerCase());
            const guessers      = (song.guessers_flat       || []).map(p => p.toLowerCase());
            const listers       = (song.listers_flat        || []).map(p => p.toLowerCase());
            const taken         = (song.lives_taken_flat    || []).map(p => p.toLowerCase());
            const saved         = (song.lives_saved_flat    || []).map(p => p.toLowerCase());

            const isPresent = roomPlayers.includes(pClean);
            if (!isPresent) return;

            totalSeen++;

            const isCorrect = guessers.includes(pClean);
            const isLister  = listers.includes(pClean);
            const isChant   = song.chanting && song.chanting.toLowerCase() === "yes";
            let typeMarker  = "";
            const rawType   = (song.type || "").toLowerCase();

            if      (rawType.includes("opening"))   typeMarker = "OP" + (song.type.match(/\d+/) || ["1"])[0];
            else if (rawType.includes("ending"))    typeMarker = "ED" + (song.type.match(/\d+/) || ["1"])[0];
            else                                    typeMarker = "IN";

            const animeTitle        = currentSearchLang === "JP" ? (song.romaji || song.english || "") : (song.english || song.romaji || "");
            const songName          = song.song || "";
            const artistName        = Array.isArray(song.artist_arr) ? song.artist_arr.join(', ') : String(song.artist_arr || '');
            const styledSongMeta    = `${truncateString(animeTitle)} (${typeMarker}): ${truncateString(songName)} by ${truncateString(artistName)}`;
            const prefixLabel       = isCorrect ? `✓ ${styledSongMeta}` : `✗ ${styledSongMeta}`;
            const bulletLabel       = `• ${styledSongMeta}`;

            allSeenSongs.push(prefixLabel);

            if (isChant) {
                totalChantSeen++;
                chantSeenList.push(prefixLabel);
            }

            if (isCorrect) {
                totalCorrect++;
                correctsList.push(bulletLabel);

                if (isChant) {
                    totalChantCorrect++;
                    chantCorrectsList.push(bulletLabel);
                }

                let numCorrect = guessers.length;

                if (numCorrect >= 1 && numCorrect <= 8) {
                    x8Counts[numCorrect - 1]++;
                    x8Lists[numCorrect - 1].push(bulletLabel);
                }

                if (!isLister) {
                    offListHits++;
                    offListHitsList.push(bulletLabel);
                }
            }

            if (isLister) {
                totalRigs++;
                rigsList.push(bulletLabel);

                if (isCorrect) {
                    rigHits++;
                    rigHitsList.push(bulletLabel);
                }
            }

            if (taken.includes(pClean)) {
                livesTaken++;
                livesTakenList.push(bulletLabel);
            }

            if (saved.includes(pClean)) {
                livesSaved++;
                livesSavedList.push(bulletLabel);
            }

            if (isCorrect && !taken.includes(pClean) && !saved.includes(pClean)) otherCorrectsList.push(bulletLabel);
        });

        return {
            totalSeen,
            totalCorrect,
            x8Counts,
            totalRigs,
            rigHits,
            offListHits,
            totalChantSeen,
            totalChantCorrect,
            livesTaken,
            livesSaved,
            allSeenSongs,
            correctsList,
            rigsList,
            rigHitsList,
            offListHitsList,
            chantSeenList,
            chantCorrectsList,
            livesTakenList,
            livesSavedList,
            otherCorrectsList,
            x8Lists
        };
    };

    const formatSampleTextList = (list) => {
        return [...list]
            .sort((a, b) => {
                const cleanA = (a.startsWith('✓') || a.startsWith('✗') || a.startsWith('• ')) ? a.replace(/^[✓✗•]\s*/, '') : a;
                const cleanB = (b.startsWith('✓') || b.startsWith('✗') || b.startsWith('• ')) ? b.replace(/^[✓✗•]\s*/, '') : b;

                return cleanA.toLowerCase().localeCompare(cleanB.toLowerCase());
            })

            .join('<br>');
    };

    const formatFractionalSample = (fractionStr, songsList) => {
        const ticks         = songsList.filter(s => s.startsWith('✓')).sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
        const crosses       = songsList.filter(s => s.startsWith('✗')).sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));

        const outputSample = [...ticks, ...crosses];
        return `<b>${fractionStr}</b><br>` + outputSample.join('<br>');
    };

    let absoluteMaxCorrectsC1 = 0;
    let absoluteMaxCorrectsC2 = 0;
    
    ["1", "2", "3", "4"].forEach((tr) => {
        if (!tierStats[tr]) return;
        tierStats[tr].forEach(p => {
            let pNameStr    = getPlayerStringName(p.Player);
            const stats     = compilePlayerStatsFromSearch(pNameStr);

            if (!pNameStr)                                  return;
            if (stats.totalCorrect > absoluteMaxCorrectsC1) absoluteMaxCorrectsC1 = stats.totalCorrect;
            if (stats.totalCorrect > absoluteMaxCorrectsC2) absoluteMaxCorrectsC2 = stats.totalCorrect;
        });
    });

    if (absoluteMaxCorrectsC1 === 0) absoluteMaxCorrectsC1 = 10;
    if (absoluteMaxCorrectsC2 === 0) absoluteMaxCorrectsC2 = 10;

    const getC1ValueAndHover = (p, mode, sub) => {
        let val         = 0;
        let hover       = "";
        let traceData   = null;
        let pNameStr    = getPlayerStringName(p.Player);
        const stats     = compilePlayerStatsFromSearch(pNameStr);

        if (sub === "BASE") {
            if (mode === "RATE") {
                val     = stats.totalSeen > 0 ? (stats.totalCorrect / stats.totalSeen) * 100 : 0;
                hover   = formatFractionalSample(`${stats.totalCorrect}/${stats.totalSeen}`, stats.allSeenSongs);
            }

            else {
                val                 = [stats.rigHits, stats.offListHits];
                traceData           = val;
                let rigsSection     = `<b>Rigs Hit Context:</b><br>${formatSampleTextList(stats.rigHitsList)}`;
                let offRigsSection  = `<b>Off Rigs Hit Context:</b><br>${formatSampleTextList(stats.offListHitsList)}`;
                hover               = `${rigsSection}<br><br>${offRigsSection}`;
            }

        }

        else if (sub === "OVER-8") {
            if (mode === "RATE") {
                let totalC  = stats.x8Counts.reduce ((a, b) => a + b, 0) || 1;
                val         = stats.x8Counts.map    (c      => (c / totalC) * 100);
            }
            
            else val = stats.x8Counts;

            traceData   = val;
            hover       = "";

        }

        else if (sub === "RIG") {
            if (!watched) return {val: 0, hover: "N/A (Watched Only)"};

            val     = mode === "RATE" ? (stats.totalSeen > 0 ? (stats.totalRigs / stats.totalSeen) * 100 : 0) : stats.totalRigs;
            hover   = formatSampleTextList(stats.rigsList);
        }

        else if (sub === "HIT") {
            if (!watched) return {val: 0, hover: "N/A (Watched Only)"};

            val     = mode === "RATE" ? (stats.totalRigs > 0 ? (stats.rigHits / stats.totalRigs) * 100 : 0) : stats.rigHits;
            hover   = formatFractionalSample(`${stats.rigHits}/${stats.totalRigs}`, stats.rigsList.map(s => stats.rigHitsList.includes(s) ? `✓ ${s.slice(2)}` : `✗ ${s.slice(2)}`));
        }

        else if (sub === "OFF") {
            if (!watched) return {val: 0, hover: "N/A (Watched Only)"};

            let totalOffSongs       = stats.totalSeen - stats.totalRigs;
            val                     = mode === "RATE" ? (totalOffSongs > 0 ? (stats.offListHits / totalOffSongs) * 100 : 0) : stats.offListHits;
            let offListSeenSongs    = stats.allSeenSongs.filter(s => !stats.rigsList.map(r => r.slice(2)).includes(s.slice(2)));
            hover                   = formatFractionalSample(`${stats.offListHits}/${totalOffSongs}`, offListSeenSongs);
        }

        else if (sub === "CHANT") {
            if (!hasChanting) return {val: 0, hover: "No Chant Songs Exist"};
            val = mode === "RATE" ? (stats.totalChantSeen > 0 ? (stats.totalChantCorrect / stats.totalChantSeen) * 100 : 0) : stats.totalChantCorrect;
            if (mode === "RATE") {
                hover = formatFractionalSample(`${stats.totalChantCorrect}/${stats.totalChantSeen}`, stats.chantSeenList);
            } else {
                hover = formatSampleTextList(stats.chantCorrectsList);
            }
        }
        return {val, hover, traceData, statsContext: stats};
    };

    const getC2ValueAndHover = (p, mode) => {
        let val             = 0;
        let hover           = "";
        let traceData       = null;
        let pNameStr        = getPlayerStringName(p.Player);
        const stats         = compilePlayerStatsFromSearch(pNameStr);
        let tk              = stats.livesTaken;
        let sv              = stats.livesSaved;
        let correctCount    = stats.totalCorrect;
        let fallbackCorrect = correctCount === 0 ? 1 : correctCount;

        if (mode === "COUNT") {
            val     = [tk, sv]; 
            hover   = "";
        }

        else {
            val     = [(tk / fallbackCorrect) * 100, (Math.max(0, correctCount - (tk + sv)) / fallbackCorrect) * 100, (sv / fallbackCorrect) * 100]; 
            hover   = "";
        }

        traceData = val;
        return {val, hover, traceData, statsContext: stats};
    };

    const buildChartData = (tabMode, sortingFn, valExtractionFn, isSubDistribution) => {
        let pool = [];

        if (tabMode === "TIER") {
            ["1", "2", "3", "4"].forEach((tr) => {
                if (!tierStats[tr] || tierStats[tr].length === 0) return;
                let arr = [...tierStats[tr]].sort(sortingFn);
                pool.push(...arr); pool.push({isSpacer: true});
            });

            if (pool.length > 0 && pool[pool.length - 1].isSpacer) pool.pop();
        }

        else {
            ["1", "2", "3", "4"].forEach((tr) => {if (tierStats[tr]) pool.push(...tierStats[tr]);});
            pool.sort(sortingFn);
        }

        let yLabels         = [];
        let customHovers    = [];
        let rawItems        = [];
        let multiData       = isSubDistribution ? Array.from({length: isSubDistribution === "X8" ? 8 : (isSubDistribution === "COUNT_BASE" ? 2 : (isSubDistribution === "COUNT_BOTH" ? 2 : 3))}, () => []) : [];
        let singleXVals     = [];

        pool.forEach(p => {
            if (p.isSpacer) {
                yLabels         .push(" ".repeat(gapCounter++));
                customHovers    .push("");
                rawItems        .push(null);

                if (isSubDistribution)  multiData.forEach(arr => arr.push(null));
                else                    singleXVals.push(null);

                return;
            }

            yLabels.push(getPlayerStringName(p.Player));
            let extracted = valExtractionFn(p);
            customHovers.push(extracted.hover);
            rawItems.push(p);

            if (isSubDistribution) for (let i = 0; i < multiData.length; i++)   multiData[i].push(extracted.traceData ? extracted.traceData[i] : 0);
            else                                                                singleXVals.push(extracted.val);
        });

        return {yLabels, customHovers, multiData, singleXVals, rawItems};
    };

    const c1Sort = (a, b) => {
        let chart1Div       = document.getElementById('tierChart_MainMetrics');
        let hiddenTraces    = (chart1Div && chart1Div.data) ? chart1Div.data.filter(t => t.visible === 'legendonly').map(t => t.name) : [];
        let extractedA      = getC1ValueAndHover(a, globalChartMode, c1Sub);
        let extractedB      = getC1ValueAndHover(b, globalChartMode, c1Sub);
        let va              = extractedA.val;
        let vb              = extractedB.val;

        if (globalChartMode === "COUNT" && c1Sub === "BASE") {
            let names   = ["Rigs Hit", "Off Rigs Hit"];
            let sa      = 0;
            let sb      = 0;

            if (!hiddenTraces.includes(names[0])) {sa += va[0]; sb += vb[0];}
            if (!hiddenTraces.includes(names[1])) {sa += va[1]; sb += vb[1];}

            return sb - sa;
        }

        else if (c1Sub === "OVER-8") {
            let sa = 0;
            let sb = 0;

            for (let i = 0; i < 8; i++) {
                let name = `${i+1}/8`;

                if (!hiddenTraces.includes(name)) {
                    sa += extractedA.traceData ? extractedA.traceData[i] : 0;
                    sb += extractedB.traceData ? extractedB.traceData[i] : 0;
                }
            }

            return sb - sa;
        }

        if (Array.isArray(va)) va = va.reduce((x, y)=>x + y, 0);
        if (Array.isArray(vb)) vb = vb.reduce((x, y)=>x + y, 0);

        return vb - va;
    };
    
    let c1LayoutSubMode = (globalChartMode === "COUNT" && c1Sub === "BASE") ? "COUNT_BASE" : (c1Sub === "OVER-8" ? "X8" : null);
    let c1Data          = buildChartData(currentTierChartMode, c1Sort, (p) => getC1ValueAndHover(p, globalChartMode, c1Sub), c1LayoutSubMode);
    let c1Traces        = [];

    if (globalChartMode === "COUNT" && c1Sub === "BASE") {
        const c1BaseColors  = [c2, c0];
        const baseNames     = ["Rigs Hit", "Off Rigs Hit"];

        let trace0Hovers = [];
        let trace1Hovers = [];

        c1Data.rawItems.forEach(p => {
            if (!p) {trace0Hovers.push(""); trace1Hovers.push(""); return;}
            let s = compilePlayerStatsFromSearch(getPlayerStringName(p.Player));
            trace0Hovers.push(formatSampleTextList(s.rigHitsList));
            trace1Hovers.push(formatSampleTextList(s.offListHitsList));
        });

        for (let i = 0; i < 2; i++) {
            c1Traces.push({
                x                   : c1Data.multiData[i]   .slice().reverse(),
                y                   : c1Data.yLabels        .slice().reverse(),
                type                : 'bar',
                orientation         : 'h',
                barmode             : 'stack',
                name                : baseNames[i],
                marker              : {color: c1BaseColors[i]},
                hovertext           : (i === 0 ? trace0Hovers : trace1Hovers).slice().reverse(),
                hoverinfo           : 'text',
                text                : c1Data.multiData[i]   .slice().reverse().map(v => v ? v.toFixed(0) : ""),
                textposition        : 'inside',
                insidetextanchor    : 'middle',
                textangle           : 0,
                textfont            : {family: 'Segoe UI', size: 15, color: 'white', weight: 'bold'}
            });
        }
    }

    else if (c1Sub === "OVER-8") {
        let x8TraceHovers = Array.from({length: 8}, () => []);

        c1Data.rawItems.forEach(p => {
            if (!p) {for (let i = 0; i < 8; i++) x8TraceHovers[i].push(""); return;}
            let s = compilePlayerStatsFromSearch(getPlayerStringName(p.Player));

            for (let i = 0; i < 8; i++) {
                let sampleList      = formatSampleTextList(s.x8Lists[i]);
                let fractionHeader  = `<b>${i + 1}/8</b>`;
                let combinedTooltip = sampleList ? `${fractionHeader}<br>${sampleList}` : fractionHeader;

                x8TraceHovers[i].push(combinedTooltip);
            }
        });

        for (let i = 0; i < 8; i++) {
            c1Traces.push({
                x                   : c1Data.multiData[i]   .slice().reverse(),
                y                   : c1Data.yLabels        .slice().reverse(),
                type                : 'bar',
                orientation         : 'h',
                barmode             : 'stack',
                name                : `${i+1}/8`,
                marker              : {color: c8Colors[i]},
                hovertext           : x8TraceHovers[i]      .slice().reverse(),
                hoverinfo           : 'text',
                text                : c1Data.multiData[i]   .slice().reverse().map(v => v ? v.toFixed(globalChartMode === "RATE" ? 1 : 0) : ""),
                textposition        : 'inside',
                insidetextanchor    : 'middle',
                textangle           : 0,
                textfont            : {family: 'Segoe UI', size: 15, color: 'white', weight: 'bold'}
            });
        }
    }

    else {
        const isRateMode    = (globalChartMode === "RATE");
        const maxCount      = Math.max(...c1Data.singleXVals.filter(v => v !== null)) || 1;

        c1Traces.push({
            x                   : c1Data.singleXVals    .slice().reverse(),
            y                   : c1Data.yLabels        .slice().reverse(),
            type                : 'bar',
            orientation         : 'h',
            barmode             : 'group',
            hovertext           : c1Data.customHovers   .slice().reverse(),
            hoverinfo           : 'text',
            text                : c1Data.singleXVals    .slice().reverse().map(v => v === null ? "" : v.toFixed(globalChartMode === "RATE" ? 2 : 0) + " "),
            textposition        : 'inside',
            insidetextanchor    : 'end',
            textangle           : 0,
            textfont            : {family: 'Segoe UI', size: 15, color: 'white', weight: 'bold'},
            marker              : {
                color       : c1Data.singleXVals.slice().reverse(),
                colorscale  : (() => {
                    if (globalChartMode === "RATE") {
                        if (c1Sub === "BASE")   return [[0, c0], [0.20, c0], [0.50, c1], [0.80, c2], [1, c2]];
                        if (c1Sub === "HIT")    return [[0, c0], [0.70, c0], [0.80, c1], [0.90, c2], [1, c2]];
                    }

                    return [[0, c0], [1, c2]];
                })(),
                cmin        : 0,
                cmax        : (() => {
                    if (globalChartMode === "RATE") {
                        if (["CHANT", "RIG", "OFF"].includes(c1Sub)) return 50;
                        return 100;
                    }

                    return Math.max(...c1Data.singleXVals.filter(v => v !== null)) || 1;
                })()
            }
        });
    }

    let displayTitleC1 = "";

    if (globalChartMode === "COUNT") {
        if      (c1Sub === "BASE")      displayTitleC1 = "Corrects";
        else if (c1Sub === "OVER-8")    displayTitleC1 = "Over-8 Hit";
        else if (c1Sub === "RIG")       displayTitleC1 = "Rigs";
        else if (c1Sub === "CHANT")     displayTitleC1 = "Chanting Hit";
    }

    else {
        if      (c1Sub === "BASE")      displayTitleC1 = "Guess Rate";
        else if (c1Sub === "OVER-8")    displayTitleC1 = "Over-8 Distribution";
        else if (c1Sub === "RIG")       displayTitleC1 = "Rig Rate";
        else if (c1Sub === "HIT")       displayTitleC1 = "Rig Guess Rate";
        else if (c1Sub === "OFF")       displayTitleC1 = "Off Rig Guess Rate";
        else if (c1Sub === "CHANT")     displayTitleC1 = "Chanting Guess Rate";
    }

    let currentChart1Div = document.getElementById('tierChart_MainMetrics');

    if (currentChart1Div && currentChart1Div.data && currentChart1Div.data.length === c1Traces.length) {
        for (let i = 0; i < c1Traces.length; i++) if(currentChart1Div.data[i].visible === 'legendonly') c1Traces[i].visible = 'legendonly';
    }

    let rangeC1;

    if (globalChartMode === "RATE") rangeC1 = [0, 105];
    else                            rangeC1 = (c1Sub === "BASE" || c1Sub === "OVER-8") ? [0, absoluteMaxCorrectsC1 + 1] : null;

    const titleC1       = `<span style="font-size: 30px;"><b>${displayTitleC1}</b></span>`;
    const hasLegends    = (c1Sub === "OVER-8" || (globalChartMode === "COUNT" && c1Sub === "BASE")) ? true : false;

    const layoutC1 = {
        font        : {family: 'Segoe UI'}, title: {text: titleC1, yref: 'container', y: 15, yanchor: 'top'},
        xaxis       : {type: 'linear', tickfont: {size: 15, color: 'black', weight: 'bold'}, fixedrange: true, showgrid: true, range: rangeC1},
        yaxis       : {type: 'category', tickfont: {size: 15, color: 'black', weight: 'bold'}, fixedrange: true, showgrid: false, ticksuffix: " "},
        bargap      : 0.0,
        barmode     : 'stack',
        margin      : {l: 150, r: 0, t: hasLegends ? 100 : 50, b: 25},
        hoverlabel  : {align: 'left', font: {family: 'Segoe UI', size: 15}}, 
        showlegend  : hasLegends,
        legend      : { 
            orientation     : "h", 
            yanchor         : "bottom",
            y               : 1.01, 
            xanchor         : "center",
            x               : 0.40, 
            traceorder      : "normal",
            entrywidthmode  : "pixels",
            entrywidth      : (globalChartMode === "COUNT" && c1Sub === "BASE") ? 40 : 20
        }
    };

    layoutC1.height = 27.5 * c1Data.yLabels.length;
    c1Traces.forEach(t => t.hoverinfo = 'none');
    Plotly.newPlot('tierChart_MainMetrics', c1Traces, layoutC1, {responsive: true, displayModeBar: false});
    let newChart1Div = document.getElementById('tierChart_MainMetrics');

    if (newChart1Div) {
        newChart1Div.removeAllListeners('plotly_legendclick');

        newChart1Div.on('plotly_legendclick', function(data) {
            let traceIndex                          = data.curveNumber;
            let currentVisibility                   = newChart1Div.data[traceIndex].visible || true;
            newChart1Div.data[traceIndex].visible   = (currentVisibility === true) ? 'legendonly' : true;

            renderTierCharts();
            return false;
        });
        
        newChart1Div.on('plotly_click', function(data) {
            if (data.event      && data.event.button    !== 0) return;
            if (!data.points    || data.points.length   === 0) return;

            const pt            = data.points[0];
            const pNameClean    = String(pt.y).trim().toLowerCase();

            if (!pNameClean) return;
            let query = "";

            if (c1Sub === "BASE") {
                if (globalChartMode === "COUNT")    query = pt.curveNumber === 0 ? `list:${pNameClean} correct:${pNameClean}` : `list!:${pNameClean} correct:${pNameClean}`;
                else                                query = `correct:${pNameClean}`;
            }

            else if (c1Sub === "OVER-8") {
                let matchX8 = pt.curveNumber + 1;
                query = `correct:${pNameClean} correct:${matchX8}`;
            }

            else if (c1Sub === "RIG")   query = `list:${pNameClean}`;
            else if (c1Sub === "HIT")   query = `list:${pNameClean} correct:${pNameClean}`;
            else if (c1Sub === "OFF")   query = `list!:${pNameClean} correct:${pNameClean}`;
            else if (c1Sub === "CHANT") query = `correct:${pNameClean} chanting:yes`;

            if (query) window.searchPlayerMetricFromTable(query);
        });

        newChart1Div.addEventListener('contextmenu', e => e.preventDefault());

        newChart1Div.on('plotly_hover', function(data) {
            if (!data.points || data.points.length === 0) return;

            const pt            = data.points[0];
            const tooltipNode   = document.getElementById('customJsTooltip');

            if (tooltipNode && pt.hovertext) {
                let traceColor = 'black';

                if (pt.fullData && pt.fullData.marker) {
                    const mColor = pt.fullData.marker.color;

                    if (Array.isArray(mColor)) {
                        const rawVal = mColor[pt.pointIndex];

                        if (rawVal !== undefined && rawVal !== null) {
                            const maxVal    = pt.fullData.marker.cmax || 100;
                            const norm      = Math.max(0, Math.min(1, rawVal / maxVal));

                            const parseHex = (hex) => {
                                let c = hex.replace('#', '');
                                if (c.length === 3) c = c.split('').map(x => x + x).join('');
                                return [parseInt(c.substring(0, 2), 16), parseInt(c.substring(2, 4), 16), parseInt(c.substring(4, 6), 16)];
                            };

                            const rgb0 = parseHex(c0);
                            const rgb1 = parseHex(c1);
                            const rgb2 = parseHex(c2);

                            let r, g, b;

                            if (globalChartMode === "RATE" && c1Sub === "BASE") {
                                if (norm <= 0.20) [r, g, b] = rgb0;

                                else if (norm <= 0.50) {
                                    let t   = (norm - 0.20) / 0.30;
                                    r       = rgb0[0] + t * (rgb1[0] - rgb0[0]);
                                    g       = rgb0[1] + t * (rgb1[1] - rgb0[1]);
                                    b       = rgb0[2] + t * (rgb1[2] - rgb0[2]);
                                }

                                else if (norm <= 0.80) {
                                    let t   = (norm - 0.50) / 0.30;
                                    r       = rgb1[0] + t * (rgb2[0] - rgb1[0]);
                                    g       = rgb1[1] + t * (rgb2[1] - rgb1[1]);
                                    b       = rgb1[2] + t * (rgb2[2] - rgb1[2]);
                                } 

                                else [r, g, b] = rgb2;
                            }

                            else if (globalChartMode === "RATE" && c1Sub === "HIT") {
                                if (norm <= 0.70) [r, g, b] = rgb0;

                                else if (norm <= 0.80) {
                                    let t = (norm - 0.70) / 0.10;
                                    r = rgb0[0] + t * (rgb1[0] - rgb0[0]);
                                    g = rgb0[1] + t * (rgb1[1] - rgb0[1]);
                                    b = rgb0[2] + t * (rgb1[2] - rgb0[2]);
                                }

                                else if (norm <= 0.90) {
                                    let t = (norm - 0.80) / 0.10;
                                    r = rgb1[0] + t * (rgb2[0] - rgb1[0]);
                                    g = rgb1[1] + t * (rgb2[1] - rgb1[1]);
                                    b = rgb1[2] + t * (rgb2[2] - rgb1[2]);
                                }

                                else [r, g, b] = rgb2;
                            }

                            else {
                                r = rgb0[0] + norm * (rgb2[0] - rgb0[0]);
                                g = rgb0[1] + norm * (rgb2[1] - rgb0[1]);
                                b = rgb0[2] + norm * (rgb2[2] - rgb0[2]);
                            }

                            traceColor = `rgb(${Math.round(r)}, ${Math.round(g)}, ${Math.round(b)})`;
                        }
                    }

                    else traceColor = mColor || 'black';
                }

                const isWhite = traceColor === 'white' || traceColor === '#ffffff' || traceColor === '#fff' || traceColor === 'rgb(255,255,255)' || traceColor === 'rgb(255, 255, 255)';

                tooltipNode.style.display           = 'block';
                tooltipNode.style.maxHeight         = '300px';
                tooltipNode.style.overflowY         = 'auto';
                tooltipNode.style.backgroundColor   = traceColor;
                tooltipNode.style.color             = isWhite ? 'black' : 'white';
                tooltipNode.style.border            = isWhite ? '1px solid black' : 'none';
                tooltipNode.innerHTML               = pt.hovertext;
                tooltipNode.style.left              = (data.event.pageX + 15) + 'px';
                tooltipNode.style.top               = (data.event.pageY + 15) + 'px';
            }
        });

        newChart1Div.on('plotly_unhover', function() {
            const tooltipNode = document.getElementById('customJsTooltip');

            if (tooltipNode && !tooltipNode.classList.contains('is-hovered')) {
                tooltipNode.style.display           = 'none';
                tooltipNode.style.backgroundColor   = 'black';
                tooltipNode.style.color             = 'white';
                tooltipNode.style.maxHeight         = '';
                tooltipNode.style.overflowY         = '';
                tooltipNode.style.border            = 'none';
            }
        });
    }

    const c2Sort = (a, b) => {
        let chart2Div       = document.getElementById('tierChart_LivesMetrics');
        let hiddenTraces    = (chart2Div && chart2Div.data) ? chart2Div.data.filter(t => t.visible === 'legendonly').map(t => t.name) : [];
        let extractedA      = getC2ValueAndHover(a, globalChartMode);
        let extractedB      = getC2ValueAndHover(b, globalChartMode);
        let va              = extractedA.val;
        let vb              = extractedB.val;
        let names           = globalChartMode === "COUNT" ? ["Lives Taken", "Lives Saved"] : ["Lives Taken", "Other Correct", "Lives Saved"];
        let sa              = 0;
        let sb              = 0;

        for (let i = 0; i < names.length; i++) {
            if (!hiddenTraces.includes(names[i])) {
                sa += extractedA.traceData ? extractedA.traceData[i] : 0;
                sb += extractedB.traceData ? extractedB.traceData[i] : 0;
            }
        }

        return sb - sa;
    };
    
    let c2LayoutSubMode = globalChartMode === "COUNT" ? "COUNT_BOTH" : "RATE_BOTH";
    let c2Data          = buildChartData(currentTierChartMode, c2Sort, (p) => getC2ValueAndHover(p, globalChartMode), c2LayoutSubMode);
    let c2Traces        = [];

    if (globalChartMode === "COUNT") {
        const c2Colors  = [c2, c0]; 
        const names     = ["Lives Taken", "Lives Saved"]; 
        let tkHovers    = [];
        let svHovers    = [];

        c2Data.rawItems.forEach(p => {
            if(!p) {tkHovers.push(""); svHovers.push(""); return;}
            let s = compilePlayerStatsFromSearch(getPlayerStringName(p.Player));
            tkHovers.push(formatSampleTextList(s.livesTakenList));
            svHovers.push(formatSampleTextList(s.livesSavedList));
        });

        for (let i = 0; i < 2; i++) {
            c2Traces.push({
                x                   : c2Data.multiData[i]               .slice().reverse(),
                y                   : c2Data.yLabels                    .slice().reverse(),
                type                : 'bar',
                orientation         : 'h',
                barmode             : 'stack',
                name                : names[i],
                marker              : {color: c2Colors[i]},
                hovertext           : (i === 0 ? tkHovers : svHovers)   .slice().reverse(),
                hoverinfo           : 'text',
                text                : c2Data.multiData[i]               .slice().reverse().map(v => v ? v.toFixed(0) : ""),
                textposition        : 'inside',
                insidetextanchor    : 'middle',
                textangle           : 0,
                textfont            : {family: 'Segoe UI', size: 15, color: 'white', weight: 'bold'}
            });
        }
    }

    else {
        const c3Colors  = [c2, c1, c0]; 
        const names     = ["Lives Taken", "Other Correct", "Lives Saved"]; 
        let tkHovers    = [];
        let othHovers   = [];
        let svHovers    = [];

        c2Data.rawItems.forEach(p => {
            if(!p) {tkHovers.push(""); othHovers.push(""); svHovers.push(""); return;}
            let s = compilePlayerStatsFromSearch(getPlayerStringName(p.Player));
            tkHovers    .push(formatSampleTextList(s.livesTakenList));
            othHovers   .push(formatSampleTextList(s.otherCorrectsList));
            svHovers    .push(formatSampleTextList(s.livesSavedList));
        });

        for (let i = 0; i < 3; i++) {
            c2Traces.push({
                x                   : c2Data.multiData[i]                                       .slice().reverse(),
                y                   : c2Data.yLabels                                            .slice().reverse(),
                type                : 'bar',
                orientation         : 'h',
                barmode             : 'stack',
                name                : names[i],
                marker              : {color: c3Colors[i]},
                hovertext           : (i === 0 ? tkHovers : (i === 1 ? othHovers : svHovers))   .slice().reverse(),
                hoverinfo           : 'text',
                text                : c2Data.multiData[i]                                       .slice().reverse().map(v => v ? v.toFixed(1) : ""),
                textposition        : 'inside',
                insidetextanchor    : 'middle',
                textangle           : 0,
                textfont            : {family: 'Segoe UI', size: 15, color: 'white', weight: 'bold'}
            });
        }
    }

    let displayTitleC2      = globalChartMode === "COUNT" ? "Contribution" : "Contribution Rate";
    let currentChart2Div    = document.getElementById('tierChart_LivesMetrics');

    if (currentChart2Div && currentChart2Div.data && currentChart2Div.data.length === c2Traces.length) {
        for (let i = 0; i < c2Traces.length; i++) if (currentChart2Div.data[i].visible === 'legendonly') c2Traces[i].visible = 'legendonly';
    }

    let rangeC2     = globalChartMode === "RATE" ? [0, 105] : [0, absoluteMaxCorrectsC2 + 1];
    const titleC2   = `<span style="font-size: 30px;"><b>${displayTitleC2}</b></span>`;

    const layoutC2 = {
        font        : {family: 'Segoe UI'}, title: {text: titleC2, yref: 'container', y: 15, yanchor: 'top'},
        xaxis       : {type: 'linear', tickfont: {size: 15, color: 'black', weight: 'bold'}, fixedrange: true, showgrid: true, range: rangeC2},
        yaxis       : {type: 'category', tickfont: {size: 15, color: 'black', weight: 'bold'}, fixedrange: true, showgrid: false, ticksuffix: " "},
        bargap      : 0.0,
        barmode     : 'stack',
        margin      : {l: 150, r: 0, t: 100, b: 25},
        hoverlabel  : {align: 'left', font: {family: 'Segoe UI', size: 15}}, 
        showlegend  : true,
        legend      : { 
            orientation     : "h", 
            yanchor         : "bottom",
            y               : 1.01, 
            xanchor         : "center",
            x               : 0.425, 
            traceorder      : "normal",
            entrywidthmode  : "pixels",
            entrywidth      : 80
        }
    };

    layoutC2.height = 27.5 * c2Data.yLabels.length;
    c2Traces.forEach(t => t.hoverinfo = 'none');
    Plotly.newPlot('tierChart_LivesMetrics', c2Traces, layoutC2, {responsive: true, displayModeBar: false});
    let newChart2Div = document.getElementById('tierChart_LivesMetrics');

    if (newChart2Div) {
        newChart2Div.removeAllListeners('plotly_legendclick');

        newChart2Div.on('plotly_legendclick', function(data) {
            let traceIndex                          = data.curveNumber;
            let currentVisibility                   = newChart2Div.data[traceIndex].visible || true;
            newChart2Div.data[traceIndex].visible   = (currentVisibility === true) ? 'legendonly' : true;

            renderTierCharts();
            return false;
        });

        newChart2Div.on('plotly_click', function(data) {
            if (data.event      && data.event.button    !== 0) return;
            if (!data.points    || data.points.length   === 0) return;

            const pt            = data.points[0];
            const pNameClean    = String(pt.y).trim().toLowerCase();

            if (!pNameClean) return;
            let query = "";

            if (globalChartMode === "COUNT") query = pt.curveNumber === 0 ? `lifetaken:${pNameClean}` : `lifesaved:${pNameClean}`;

            else {
                if      (pt.curveNumber === 0) query = `lifetaken:${pNameClean}`;
                else if (pt.curveNumber === 1) query = `correct:${pNameClean} lifetaken!:${pNameClean} lifesaved!:${pNameClean}`;
                else                           query = `lifesaved:${pNameClean}`;
            }

            if (query) window.searchPlayerMetricFromTable(query);
        });

        newChart2Div.addEventListener('contextmenu', e => e.preventDefault());

        newChart2Div.on('plotly_hover', function(data) {
            if (!data.points || data.points.length === 0) return;

            const pt            = data.points[0];
            const tooltipNode   = document.getElementById('customJsTooltip');
            
            if (tooltipNode && pt.hovertext) {
                let traceColor = 'black';

                if (pt.fullData && pt.fullData.marker) {
                    traceColor = pt.fullData.marker.color;
                    if (Array.isArray(traceColor)) traceColor = traceColor[pt.pointIndex] || 'black';
                }

                const isWhite = traceColor === 'white' || traceColor === '#ffffff' || traceColor === '#fff' || traceColor === 'rgb(255,255,255)' || traceColor === 'rgb(255, 255, 255)';

                tooltipNode.style.display           = 'block';
                tooltipNode.style.maxHeight         = '300px';
                tooltipNode.style.overflowY         = 'auto';
                tooltipNode.style.backgroundColor   = traceColor;
                tooltipNode.style.color             = isWhite ? 'black' : 'white';
                tooltipNode.style.border            = isWhite ? '1px solid black' : 'none';
                tooltipNode.innerHTML               = pt.hovertext;
                tooltipNode.style.left              = (data.event.pageX + 15) + 'px';
                tooltipNode.style.top               = (data.event.pageY + 15) + 'px';
            }
        });

        newChart2Div.on('plotly_unhover', function() {
            const tooltipNode = document.getElementById('customJsTooltip');

            if (tooltipNode && !tooltipNode.classList.contains('is-hovered')) {
                tooltipNode.style.display           = 'none';
                tooltipNode.style.backgroundColor   = 'black';
                tooltipNode.style.color             = 'white';
                tooltipNode.style.maxHeight         = '';
                tooltipNode.style.overflowY         = '';
                tooltipNode.style.border            = 'none';
            }
        });
    }

    const getC3ValueAndHover = (p, mode) => {
        let pNameStr    = getPlayerStringName(p.Player);
        let pClean      = pNameStr.replace(/[★▲▼]/g, "").trim().toLowerCase();
        let pData       = window.dashboardData.json_players.find(x => getPlayerStringName(x.Player).replace(/[★▲▼]/g, "").trim().toLowerCase() === pClean);

        if (!pData || !pData["Median Time"]) return {val: null, hover: "No times logged"};

        let det         = pData["Median Time"].details;
        let val         = pData["Median Time"].count; 
        let metricsMap  = {};

        if (det) {
            metricsMap["MIN"]   = {label: "Minimum",            val: parseFloat(det[0].split(": ")[1]).toFixed(2)};
            metricsMap["MEAN"]  = {label: "Mean",               val: parseFloat(det[1].split(": ")[1]).toFixed(2) };
            metricsMap["MED"]   = {label: "Median",             val: parseFloat(pData["Median Time"].count).toFixed(2)};
            metricsMap["MAX"]   = {label: "Maximum",            val: parseFloat(det[2].split(": ")[1]).toFixed(2)};
            metricsMap["STDEV"] = {label: "Standard Deviation", val: parseFloat(det[3].split(": ")[1]).toFixed(2)};

            val = parseFloat(metricsMap[mode].val);
        }

        let nonChartedLines = [];

        Object.keys(metricsMap).forEach(k => {
            if (k !== mode) {
                let cleanValStr = metricsMap[k].val.replace(/s$/, ""); 
                nonChartedLines.push(`${metricsMap[k].label}: ${cleanValStr}`);
            }
        });

        return {val, hover: nonChartedLines.join("<br>")};
    };

    const c3Sort = (a, b) => {
        let va = getC3ValueAndHover(a, c3Mode).val || 20;
        let vb = getC3ValueAndHover(b, c3Mode).val || 20;

        return va - vb; 
    };

    let c3Data          = buildChartData(currentTierChartMode, c3Sort, (p) => getC3ValueAndHover(p, c3Mode), null);
    const validTimes    = c3Data.singleXVals.filter(v => v !== null);
    const minTime       = Math.min(...validTimes) || 0;
    const maxTime       = Math.max(...validTimes) || 20;

    let c3Traces = [{
        x                   : c3Data.singleXVals    .slice().reverse(),
        y                   : c3Data.yLabels        .slice().reverse(),
        type                : 'bar',
        orientation         : 'h',
        hovertext           : c3Data.customHovers   .slice().reverse(),
        hoverinfo           : 'none',
        text                : c3Data.singleXVals    .slice().reverse().map(v => v === null ? "" : v.toFixed(2) + " "),
        textposition        : 'inside', 
        insidetextanchor    : 'end',
        textangle           : 0,
        textfont            : {family: 'Segoe UI', size: 15, color: 'white', weight: 'bold'},
        marker              : {
            color       : c3Data.singleXVals.slice().reverse(),
            colorscale  : [[0, c2], [1, c0]],
            cmin        : minTime,
            cmax        : maxTime
        }
    }];

    let displayTitleC3 = "";

    if      (c3Mode === "MIN")      displayTitleC3 = "Minimum Time";
    else if (c3Mode === "MEAN")     displayTitleC3 = "Mean Time";
    else if (c3Mode === "MED")      displayTitleC3 = "Median Time";
    else if (c3Mode === "MAX")      displayTitleC3 = "Maximum Time";
    else if (c3Mode === "STDEV")    displayTitleC3 = "Time Standard Deviation";

    const titleC3 = `<span style="font-size: 30px;"><b>${displayTitleC3}</b></span>`;

    const layoutC3 = {
        font        : {family: 'Segoe UI'}, title: {text: titleC3, yref: 'container', y: 15, yanchor: 'top'},
        xaxis       : {tickfont: {size: 15, color: 'black', weight: 'bold'}, fixedrange: true, showgrid: true,  range       : [0, 21], tickmode: 'array', tickvals: [0, 4, 8, 12, 16, 20]},
        yaxis       : {tickfont: {size: 15, color: 'black', weight: 'bold'}, fixedrange: true, showgrid: false, ticksuffix  : " "},
        bargap      : 0.0,
        margin      : {l: 150, r: 0, t: 50, b: 25},
        hoverlabel  : {align: 'left', font: {family: 'Segoe UI', size: 15}},
        showlegend  : false
    };

    layoutC3.height = 27.5 * c3Data.yLabels.length;
    Plotly.newPlot('tierChart_TimeMetrics', c3Traces, layoutC3, {responsive: true, displayModeBar: false});
    let newChart3Div = document.getElementById('tierChart_TimeMetrics');

    if (newChart3Div) {
        newChart3Div.addEventListener('contextmenu', e => e.preventDefault());

        newChart3Div.on('plotly_hover', function(data) {
            if (!data.points || data.points.length === 0) return;

            const pt            = data.points[0];
            const tooltipNode   = document.getElementById('customJsTooltip');

            if (tooltipNode && pt.hovertext) {
                let traceColor = 'white';

                if (pt.fullData && pt.fullData.marker) {
                    const mColor = pt.fullData.marker.color;

                    if (Array.isArray(mColor)) {
                        const rawVal = mColor[pt.pointIndex];

                        if (rawVal !== undefined && rawVal !== null) {
                            const minVal    = pt.fullData.marker.cmin || 0;
                            const maxVal    = pt.fullData.marker.cmax || 20;
                            const norm      = Math.max(0, Math.min(1, (rawVal - minVal) / (maxVal - minVal || 1)));

                            const parseHex = (hex) => {
                                let c = hex.replace('#', '');
                                if (c.length === 3) c = c.split('').map(x => x + x).join('');
                                return [parseInt(c.substring(0, 2), 16), parseInt(c.substring(2, 4), 16), parseInt(c.substring(4, 6), 16)];
                            };

                            const rgb2  = parseHex(c2);
                            const rgb0  = parseHex(c0);
                            const r     = rgb2[0] + norm * (rgb0[0] - rgb2[0]);
                            const g     = rgb2[1] + norm * (rgb0[1] - rgb2[1]);
                            const b     = rgb2[2] + norm * (rgb0[2] - rgb2[2]);
                            traceColor  = `rgb(${Math.round(r)}, ${Math.round(g)}, ${Math.round(b)})`;
                        }
                    }

                    else traceColor = mColor || 'white';
                }

                const isWhite = traceColor === 'white' || traceColor === '#ffffff' || traceColor === '#fff' || traceColor === 'rgb(255,255,255)' || traceColor === 'rgb(255, 255, 255)';

                tooltipNode.style.display           = 'block';
                tooltipNode.style.maxHeight         = '300px';
                tooltipNode.style.overflowY         = 'auto';
                tooltipNode.style.backgroundColor   = traceColor;
                tooltipNode.style.color             = isWhite ? 'black' : 'white';
                tooltipNode.style.border            = isWhite ? '1px solid black' : 'none';
                tooltipNode.innerHTML               = pt.hovertext;
                tooltipNode.style.left              = (data.event.pageX + 15) + 'px';
                tooltipNode.style.top               = (data.event.pageY + 15) + 'px';
            }
        });

        newChart3Div.on('plotly_unhover', function() {
            const tooltipNode = document.getElementById('customJsTooltip');

            if (tooltipNode && !tooltipNode.classList.contains('is-hovered')) {
                tooltipNode.style.display           = 'none';
                tooltipNode.style.backgroundColor   = 'black';
                tooltipNode.style.color             = 'white';
                tooltipNode.style.maxHeight         = '';
                tooltipNode.style.overflowY         = '';
                tooltipNode.style.border            = 'none';
            }
        });
    }
}

function setupTooltipListeners() {
    const tooltipNode = document.getElementById('customJsTooltip');
    let hideTimeout = null;

    function positionTooltip(e) {
        if (tooltipNode.classList.contains('is-hovered')) return;

        tooltipNode.style.display = 'block';
        const tooltipWidth = tooltipNode.offsetWidth; 
        const tooltipHeight = tooltipNode.offsetHeight;

        let xPos = e.pageX + 15;
        let yPos = e.pageY + 15;

        if (e.clientX + 15 + tooltipWidth   > window.innerWidth)    xPos = e.pageX - tooltipWidth   - 15;
        if (e.clientY + 15 + tooltipHeight  > window.innerHeight)   yPos = e.pageY - tooltipHeight  - 15;

        if (xPos < window.scrollX) xPos = window.scrollX + 5;
        if (yPos < window.scrollY) yPos = window.scrollY + 5;

        tooltipNode.style.left = xPos + 'px';
        tooltipNode.style.top = yPos + 'px';
    }

    function clearHideTimeout() {
        if (hideTimeout) {
            clearTimeout(hideTimeout);
            hideTimeout = null;
        }
    }

    function requestHideTooltip() {
        clearHideTimeout();

        hideTimeout = setTimeout(() => {
            if (!tooltipNode.classList.contains('is-hovered')) {
                tooltipNode.style.display           = 'none';
                tooltipNode.style.backgroundColor   = 'black';
                tooltipNode.style.color             = 'white';
                tooltipNode.style.maxHeight         = '';
                tooltipNode.style.overflowY         = '';
            }
        }, 100);
    }

    if (tooltipNode && !tooltipNode._bound) {
        tooltipNode._bound = true;

        tooltipNode.addEventListener('mouseenter', () => {
            clearHideTimeout();
            tooltipNode.classList.add('is-hovered');
        });

        tooltipNode.addEventListener('mouseleave', () => {
            tooltipNode.classList.remove('is-hovered');
            requestHideTooltip();
        });

        window.addEventListener('wheel', (e) => {
            if (tooltipNode.style.display === 'block') {
                const rect = tooltipNode.getBoundingClientRect();

                const isOverTooltip = (
                    e.clientX >= rect.left && e.clientX <= rect.right &&
                    e.clientY >= rect.top && e.clientY <= rect.bottom
                );
                
                if (tooltipNode.scrollHeight > tooltipNode.clientHeight) {
                    e.preventDefault();
                    tooltipNode.scrollTop += e.deltaY;
                }
            }
        }, {passive: false});
    }

    document.querySelectorAll('[data-metric]').forEach(th => {
        const metricKey = th.getAttribute('data-metric');
        if (!colExplanations[metricKey]) return;

        th.removeEventListener('mouseenter',    th._handlerEnter);
        th.removeEventListener('mousemove',     positionTooltip);
        th.removeEventListener('mouseleave',    th._handlerLeave);

        th._handlerEnter = (e)  => { clearHideTimeout(); tooltipNode.innerHTML = colExplanations[metricKey]; positionTooltip(e); };
        th._handlerLeave = ()   => { requestHideTooltip(); };

        th.addEventListener('mouseenter', th._handlerEnter);
        th.addEventListener('mousemove',  positionTooltip);
        th.addEventListener('mouseleave', th._handlerLeave);
    });

    document.querySelectorAll('td[data-songs]').forEach(td => {
        td.addEventListener('mouseenter', (e) => {
            try {
                clearHideTimeout();
                const songs = JSON.parse(decodeURIComponent(td.getAttribute('data-songs')));
                if (!songs || songs.length === 0) return;

                if      (td.classList.contains('highlight-best'))   {tooltipNode.style.backgroundColor = c2;        tooltipNode.style.color = 'white';}
                else if (td.classList.contains('highlight-worst'))  {tooltipNode.style.backgroundColor = c0;        tooltipNode.style.color = 'white';}
                else                                                {tooltipNode.style.backgroundColor = 'black';   tooltipNode.style.color = 'white';}

                const metricName = td.getAttribute('data-metric');

                if (metricName === "Off GR") {
                    let cleanSongs      = [...songs];
                    let fractionHeader  = "";

                    if (cleanSongs.length > 0 && (/^\d+\/\d+$/.test(cleanSongs[0]) || /^\d+-\d+-\d+$/.test(cleanSongs[0]))) {
                        fractionHeader = `<b>${cleanSongs[0]}</b>`;
                        cleanSongs.shift();
                    }

                    else {
                        let total           = cleanSongs.length;
                        let correctCount    = cleanSongs.filter(s => s.startsWith('✓')).length;
                        fractionHeader      = `<b>${correctCount}/${total}</b>`;
                    }

                    let correctSongs    = cleanSongs.filter(s => s.startsWith('✓')).map(s => s.slice(2));
                    correctSongs        = window.translateHoverText(correctSongs);
                    let displaySongs    = formatAndSortSongsList(correctSongs, true);

                    tooltipNode.style.maxHeight = '300px';
                    tooltipNode.style.overflowY = 'auto';
                    tooltipNode.innerHTML       = `${fractionHeader}<br>${displaySongs.join('<br>')}`;

                    positionTooltip(e);
                    return;
                }

                let displaySongs        = [...songs];
                displaySongs            = window.translateHoverText(displaySongs);
                const isPlayerSubHover  = td.parentNode.firstElementChild === td;

                if (songs.length === 1 && !songs[0].startsWith('✓') && !songs[0].startsWith('✗') && songs[0].includes('/')) {
                    tooltipNode.innerHTML = songs[0];
                    positionTooltip(e);
                    return;
                }

                if (songs.some(s => s.startsWith("Minimum:"))) {
                    const metricName = td.getAttribute('data-player-metric') || (activePlayerHeadersConfig[td.cellIndex] ? activePlayerHeadersConfig[td.cellIndex].name : "");
                    
                    if (metricName === "Median Vintage Hit") {
                        displaySongs = displaySongs.map(line => {
                            if (line.startsWith("Standard Deviation:")) return line.replace(/:\s*([0-9.]+)/g, (match, p1) => `: ${parseFloat(p1).toFixed(2)} years`);
                            return line.replace(/:\s*([0-9.]+)/g, (match, p1) => {return `: ${parseFloatToVintage(parseFloat(p1))}`;});
                        });
                    }

                    tooltipNode.innerHTML = displaySongs.join('<br>');
                    positionTooltip(e);
                    return;
                }

                const fractionRegex = /^\d+\/\d+$/;
                const isWLTBracket  = /^\d+-\d+-\d+$/.test(songs[0]);
                const containsRegex = fractionRegex.test(songs[0]) || isWLTBracket;
                let fractionHeader  = "";

                if (containsRegex) {
                    fractionHeader = `<b>${songs[0]}</b>`;
                    displaySongs.shift();
                }

                if (isWLTBracket) {
                    tooltipNode.style.maxHeight = '300px';
                    tooltipNode.style.overflowY = 'auto';
                    tooltipNode.innerHTML       = `${fractionHeader}<br>${displaySongs.join('<br>')}`;

                    positionTooltip(e);
                    return;
                }

                if (containsRegex)  displaySongs = sampleLargeSongList(displaySongs).map(s => (s.startsWith('✓') || s.startsWith('✗') || !isPlayerSubHover) ? s : `• ${s}`);
                else                displaySongs = formatAndSortSongsList(displaySongs, !isPlayerSubHover);

                tooltipNode.style.maxHeight = '300px';
                tooltipNode.style.overflowY = 'auto';
                tooltipNode.innerHTML       = containsRegex ? `${fractionHeader}<br>${displaySongs.join('<br>')}` : displaySongs.join('<br>');

                positionTooltip(e);
            }

            catch (err) {}
        });

        td.addEventListener('mousemove',    positionTooltip);
        td.addEventListener('mouseleave',   requestHideTooltip);
    });
}

renderPlayerTable       ();
renderTourTable         ();
renderTeamTable         ();
renderTierCharts        ();
setupTooltipListeners   ();

const xLabels = (numX === 8)
? ['5', '10', '15', '20', '25', '30', '35']
: ['5', '10', '15', '20', '25', '30', '35', '40'];

const yLabels = (numY === 8)
? [1995, 2000, 2005, 2010, 2015, 2020, 2025]
: [1990, 1995, 2000, 2005, 2010, 2015, 2020, 2025];

const matrixBins = {};

songData.forEach(s => {
    let xIdx    = Math.min(Math.floor(s.difficulty / 5), numX - 1);
    let yIdx    = (numY === 8) ? ((s.vintage < 1995) ? 0 : Math.min(Math.floor((s.vintage - 1995) / 5) + 1, 7)) : Math.min(Math.max(Math.floor((s.vintage - 1985) / 5), 0), 8);
    let key     = `${xIdx}-${yIdx}`;

    if(!matrixBins[key]) matrixBins[key] = {count: 0, over8Sum: 0};

    matrixBins[key].count++;
    matrixBins[key].over8Sum += s.correct_count;
});

let zValues     = [];
let textLabels  = [];
let annotations = [];

for (let i = 0; i < numY; i++) {
    let rowZ    = [];
    let rowText = [];

    for (let j = 0; j < numX; j++) {
        let key         = `${j}-${i}`;
        let vintageStr  = "";
        let diffStr     = "";

        if (numY === 8) {
            if      (i === 0) vintageStr = "Vintage: <1995";
            else if (i === 7) vintageStr = "Vintage: >2025";
            else {
                let startYr = 1995 + (i - 1) * 5;
                vintageStr  = `Vintage: ${startYr}-${startYr + 5}`;
            }
        }

        else {
            if      (i === 0) vintageStr = "Vintage: <1990";
            else if (i === 8) vintageStr = "Vintage: >2025";
            else {
                let startYr = 1990 + (i - 1) * 5;
                vintageStr  = `Vintage: ${startYr}-${startYr + 5}`;
            }
        }

        if (numX === 8) {
            if      (j === 0)   diffStr = "Difficulty: <5";
            else if (j === 7)   diffStr = "Difficulty: >35";
            else {
                let startDf = j         * 5;
                diffStr     = `Difficulty: ${startDf}-${startDf + 5}`;
            }
        }

        else {
            if      (j === 0)   diffStr = "Difficulty: <5";
            else if (j === 8)   diffStr = "Difficulty: >40";
            else {
                let startDf = j         * 5;
                diffStr     = `Difficulty: ${startDf}-${startDf + 5}`;
            }
        }

        if (key in matrixBins) {
            let val = matrixBins[key].over8Sum / matrixBins[key].count;
            rowZ.push(val);

            let bin_songs = matrixSongs[key] ? [...matrixSongs[key]] : [];

            bin_songs = bin_songs
                .sort((a, b) => {
                    const cleanA = (a.startsWith('✓') || a.startsWith('✗')) ? a.slice(2) : a;
                    const cleanB = (b.startsWith('✓') || b.startsWith('✗')) ? b.slice(2) : b;

                    return cleanA.toLowerCase().localeCompare(cleanB.toLowerCase());
                })

                .map(s => (s.startsWith('✓') || s.startsWith('✗')) ? s : `• ${s}`);

            rowText.push(`<b>${diffStr}<br>${vintageStr}<br>Over-8: ${val.toFixed(2)}</b>`);

            annotations.push({
                x               : j,
                y               : i,
                text            : `<b>${matrixBins[key].count}</b>`,
                font            : {family: 'Segoe UI', size: (numX > 8 ? 50 : 60), color: 'white'},
                showarrow       : false,
                captureevents   : false
            });
        }

        else {
            rowZ    .push(null);
            rowText .push(`<b>${vintageStr}<br>${diffStr}<br>Mean Over-8: N/A</b>`);
        }
    }

    zValues     .push(rowZ);
    textLabels  .push(rowText);
}

function getBgColor(val, minZ, maxZ, color0, color1, color2) {
    if (val === null || val === undefined || isNaN(val)) return 'rgba(0, 0, 0, 0)';

    let norm = Math.max(0, Math.min(1, (val - minZ) / (maxZ - minZ)));
    let r, g, b;

    const parseHex = (hex) => {
        let c = hex.replace('#', '');
        if (c.length === 3) c = c.split('').map(x => x + x).join('');
        return [parseInt(c.substring(0, 2), 16), parseInt(c.substring(2, 4), 16), parseInt(c.substring(4, 6), 16)];
    };

    const rgb0 = parseHex(color0);
    const rgb1 = parseHex(color1);
    const rgb2 = parseHex(color2);

    if (norm <= 0.375) {
        let t = norm / 0.375;
        r = rgb0[0] + t * (rgb1[0] - rgb0[0]);
        g = rgb0[1] + t * (rgb1[1] - rgb0[1]);
        b = rgb0[2] + t * (rgb1[2] - rgb0[2]);
    }

    else if (norm <= 0.625) {
        let t = (norm - 0.375) / (0.25);
        r = rgb1[0] + t * (rgb2[0] - rgb1[0]);
        g = rgb1[1] + t * (rgb2[1] - rgb1[1]);
        b = rgb1[2] + t * (rgb2[2] - rgb1[2]);
    }

    else [r, g, b] = rgb2;

    return `rgb(${Math.round(r)}, ${Math.round(g)}, ${Math.round(b)})`;
}

const bgColors = zValues.map(row => row.map(val => getBgColor(val, 0, 8, c0, c1, c2)));

Plotly.newPlot('plotlySongChart', [{
    z               : zValues,
    x               : Array.from({length: numX}, (_, i) => i),
    y               : Array.from({length: numY}, (_, i) => i),
    text            : textLabels,
    hoverinfo       : 'none',
    type            : 'heatmap',
    colorscale      : [[0, c0], [0.375, c1], [0.625, c2], [1, c2]],
    zmin            : 0,
    zmax            : 8,
    showscale       : true,
    colorbar        : {
        title       : {text: '<b>Over-8</b>', font: {family: 'Segoe UI', size: 25, color: 'black', weight: 'bold'}, side: 'right'},
        thickness   : 25,
        len         : 1.0,
        y           : 0.5,
        yanchor     : 'middle',
        x           : 1,
        xpad        : -10,
        tickmode    : 'array',
        tickvals    : [0, 3, 5, 8],
        ticktext    : ['0', '3', '5', '8'],
        tickfont    : {family: 'Segoe UI', size: 20, color: 'black', weight: 'bold'}
    }
}], {
    font        : {family: 'Segoe UI', size: 50},
    xaxis       : {
        title           : {text: '<b>Difficulty</b>', font: {family: 'Segoe UI', size: 25, color: 'black', weight: 'bold'}, pad: 5},
        tickmode        : 'array',
        tickvals        : Array.from({length: numX - 1}, (_, i) => i + 0.5),
        ticktext        : xLabels,
        tickfont        : {family: 'Segoe UI', size: 20, color: 'black', weight: 'bold'},
        showgrid        : true,
        zeroline        : false,
        showticklabels  : true,
        ticks           : 'outside',
        ticklen         : 5,
        tickcolor       : 'rgba(0, 0, 0, 0)',
        fixedrange      : true
    },
    yaxis: {
        title           : {text: '<b>Vintage</b>', font: {family: 'Segoe UI', size: 25, color: 'black', weight: 'bold'}, pad: 5},
        tickmode        : 'array',
        tickvals        : Array.from({length: numY - 1}, (_, i) => i + 0.5),
        ticktext        : yLabels,
        tickfont        : {family: 'Segoe UI', size: 20, color: 'black', weight: 'bold'},
        tickangle       : -90,
        showgrid        : true,
        zeroline        : false,
        showticklabels  : true,
        ticks           : 'outside',
        ticklen         : 5,
        tickcolor       : 'rgba(0, 0, 0, 0)',
        fixedrange      : true
    },
    annotations : annotations,
    margin      : {l: 75, r: 0, t: 0, b: 75}
}, {responsive: true, displayModeBar: false});

const songChartDiv = document.getElementById('plotlySongChart');

if (songChartDiv) {
    songChartDiv.addEventListener('contextmenu', e => e.preventDefault());

    songChartDiv.on('plotly_hover', function(data) {
        if (!data.points || data.points.length === 0) return;

        const pt    = data.points[0];
        const j     = pt.x;
        const i     = pt.y;
        const key   = `${j}-${i}`;

        if (matrixSongs && matrixSongs[key]) {
            const tooltipNode = document.getElementById('customJsTooltip');
            if (!tooltipNode) return;

            let bin_songs   = [...matrixSongs[key]];
            bin_songs       = window.translateHoverText(bin_songs);

            bin_songs = bin_songs
                .sort((a, b) => {
                    const cleanA = (a.startsWith('✓') || a.startsWith('✗')) ? a.slice(2) : a;
                    const cleanB = (b.startsWith('✓') || b.startsWith('✗')) ? b.slice(2) : b;

                    return cleanA.toLowerCase().localeCompare(cleanB.toLowerCase());
                })

                .map(s => (s.startsWith('✓') || s.startsWith('✗')) ? s : `• ${s}`);

            const baseInfo          = textLabels[i][j]  || "";
            const currentCellColor  = bgColors[i][j]    || 'black';
            
            const isWhite = currentCellColor === 'white' || currentCellColor === '#ffffff' || currentCellColor === 'rgb(255,255,255)' || currentCellColor === 'rgb(255, 255, 255)';

            tooltipNode.style.display           = 'block';
            tooltipNode.style.maxHeight         = '300px';
            tooltipNode.style.overflowY         = 'auto';
            tooltipNode.style.backgroundColor   = currentCellColor;
            tooltipNode.style.color             = isWhite ? 'black' : 'white';
            tooltipNode.style.border            = isWhite ? '1px solid black' : 'none';
            tooltipNode.innerHTML               = `${baseInfo}<br>${bin_songs.join('<br>')}`;
            const event                         = data.event;
            tooltipNode.style.left              = (event.pageX + 15) + 'px';
            tooltipNode.style.top               = (event.pageY + 15) + 'px';
        }
    });

    songChartDiv.on('plotly_unhover', function() {
        const tooltipNode = document.getElementById('customJsTooltip');

        if (tooltipNode && !tooltipNode.classList.contains('is-hovered')) {
            tooltipNode.style.display           = 'none';
            tooltipNode.style.backgroundColor   = 'black';
            tooltipNode.style.color             = 'white';
            tooltipNode.style.maxHeight         = '';
            tooltipNode.style.overflowY         = '';
            tooltipNode.style.border            = 'none';
        }
    });

    songChartDiv.on('plotly_click', function(data) {
        if (data.event      && data.event.button    !== 0) return;
        if (!data.points    || data.points.length   === 0) return;

        const pt        = data.points[0];
        const j         = pt.x;
        const i         = pt.y;
        const key       = `${j}-${i}`;
        let queryParts  = [];

        if (!(key in matrixBins)) return;

        if (numY === 8) {
            if      (i === 0) queryParts.push("vintage<1995");
            else if (i === 7) queryParts.push("vintage>=2025");
            else {
                let startYr = 1995 + (i - 1) * 5;
                let endYr   = startYr + 5;
                queryParts.push(`vintage>=${startYr}`, `vintage<${endYr}`);
            }
        }

        else {
            if      (i === 0) queryParts.push("vintage<1990");
            else if (i === 8) queryParts.push("vintage>=2025");
            else {
                let startYr = 1990 + (i - 1) * 5;
                let endYr   = startYr + 5;
                queryParts.push(`vintage>=${startYr}`, `vintage<${endYr}`);
            }
        }

        if (numX === 8) {
            if      (j === 0) queryParts.push("difficulty<5");
            else if (j === 7) queryParts.push("difficulty>35");
            else {
                let startDf = j * 5;
                let endDf   = startDf + 5;
                queryParts.push(`difficulty>${startDf}`, `difficulty<${endDf}`);
            }
        }

        else {
            if      (j === 0) queryParts.push("difficulty<5");
            else if (j === 8) queryParts.push("difficulty>40");
            else {
                let startDf = j * 5;
                let endDf   = startDf + 5;
                queryParts.push(`difficulty>${startDf}`, `difficulty<${endDf}`);
            }
        }

        if (queryParts.length > 0) {window.searchPlayerMetricFromTable(queryParts.join(" "));}
    });
}

function hexToRgba(hex, opacity = 0.95) {
    let c = hex.replace('#', '');
    if (c.length === 3) c = c.split('').map(x => x + x).join('');

    const r = parseInt(c.substring(0, 2), 16);
    const g = parseInt(c.substring(2, 4), 16);
    const b = parseInt(c.substring(4, 6), 16);

    return `rgba(${r}, ${g}, ${b}, ${opacity})`;
}

function buildScatterAnnotations(data, xKey, yKey, sizeKeyMultiplier) {
    const defaultAnn = [
        {x: 0, y: 1, xref: 'paper', yref: 'paper', text: '<b>New<br>Hard</b>', showarrow: false, font: {size: 20}, opacity: 0.75, xanchor: 'left',  yanchor: 'top'},
        {x: 1, y: 1, xref: 'paper', yref: 'paper', text: '<b>New<br>Easy</b>', showarrow: false, font: {size: 20}, opacity: 0.75, xanchor: 'right', yanchor: 'top'},
        {x: 1, y: 0, xref: 'paper', yref: 'paper', text: '<b>Old<br>Easy</b>', showarrow: false, font: {size: 20}, opacity: 0.75, xanchor: 'right', yanchor: 'bottom'},
        {x: 0, y: 0, xref: 'paper', yref: 'paper', text: '<b>Old<br>Hard</b>', showarrow: false, font: {size: 20}, opacity: 0.75, xanchor: 'left',  yanchor: 'bottom'}
    ];

    const xCenter = data.reduce((sum, d) => sum + d[xKey], 0) / data.length;
    const yCenter = data.reduce((sum, d) => sum + d[yKey], 0) / data.length;

    data.forEach(d => {
        const bubbleSize = Math.max(10, d[sizeKeyMultiplier] * 1.5);
        const baseOffset = (bubbleSize / 2) + 10;

        const dx = d[xKey] - xCenter;
        const dy = d[yKey] - yCenter;
        
        const angle = (dx === 0 && dy === 0) ? Math.PI / 4 : Math.atan2(dy, dx);

        const axVal = Math.cos(angle) * baseOffset;
        const ayVal = Math.sin(angle) * baseOffset * -1;

        const xAlign = Math.cos(angle) > 0.1 ? 'left'   : (Math.cos(angle) < -0.1 ? 'right' : 'center');
        const yAlign = Math.sin(angle) > 0.1 ? 'bottom' : (Math.sin(angle) < -0.1 ? 'top'   : 'middle');

        defaultAnn.push({
            x           : d[xKey],
            y           : d[yKey],
            text        : `<b>${d.acronym}</b>`,
            font        : {family: 'Segoe UI', size: 20, color: 'black'},
            showarrow   : true,
            arrowhead   : 0,
            arrowwidth  : 1,
            arrowcolor  : 'rgba(0, 0, 0, 0.5)',
            ax          : axVal,
            ay          : ayVal,
            xanchor     : xAlign,
            yanchor     : yAlign
        });
    });

    return defaultAnn;
}

if (scatterData) {
    const guessHull   = get75PercentileHull(scatterData, 'over8', 'vintage');
    let guessTraces   = [];

    if (guessHull) guessTraces.push({
        x           : guessHull.x,
        y           : guessHull.y,
        type        : 'scatter',
        mode        : 'lines',
        line        : {color: 'black', width: 0.5, dash: 'solid'},
        hoverinfo   : 'skip',
        showlegend  : false
    });

    guessTraces.push({
        x               : scatterData.map(d => d.over8),
        y               : scatterData.map(d => d.vintage),
        text            : scatterData.map(d => d.acronym),
        customdata      : scatterData.map(d => [d.name, d.over8.toFixed(2), d.seasonal_vintage, d.gr.toFixed(2), d.performance.toFixed(2)]),
        hovertemplate   : '<b>%{customdata[0]}</b><br>Mean Over-8: %{customdata[1]}<br>Median Vintage: %{customdata[2]}<br>GR: %{customdata[3]}<br>Score: %{customdata[4]}<extra></extra>',
        hoverlabel      : {align: 'left', font: {family: 'Segoe UI', size: 15}},
        mode            : 'markers',
        showlegend      : false,
        marker          : {
            size        : scatterData.map(d => Math.max(10, d.gr * 1.5)),
            opacity     : 0.95,
            color       : scatterData.map(d => d.performance),
            colorscale  : [[0, hexToRgba(c0)], [0.5, hexToRgba(c1)], [1, hexToRgba(c2)]],
            showscale   : true,
            colorbar    : {
                title       : {text: '<b>Score</b>', font: {family: 'Segoe UI', size: 25, color: 'black', weight: 'bold'}, side: 'right'},
                thickness   : 25,
                len         : 1.0,
                y           : 0.5,
                yanchor     : 'middle',
                x           : 1,
                tickmode    : 'array',
                tickvals    : [0, 50, 100],
                ticktext    : ['0', '50', '100'],
                tickfont    : {family: 'Segoe UI', size: 20, color: 'black', weight: 'bold'}
            },
            line        : {color: 'black', width: 0},
            cmin        : 0,
            cmax        : 100
        }
    });

    Plotly.newPlot('plotlyGuessChart', guessTraces, {
        font        : {family: 'Segoe UI'},
        xaxis       : {
            title       : {text: '<b>Over-8</b>', font: {family: 'Segoe UI', size: 25, color: 'black', weight: 'bold'}, pad: 5},
            tickfont    : {family: 'Segoe UI', size: 20, color: 'black', weight: 'bold'},
            showgrid    : true,
            tickformat  : '.1f',
            dtick       : 0.5,
            ticks       : 'outside',
            ticklen     : 5,
            tickcolor   : 'rgba(0, 0, 0, 0)',
            fixedrange  : false,
            range       : [window.unifiedChartLimits.xMin, window.unifiedChartLimits.xMax]
        },
        yaxis       : {
            title       : {text: '<b>Vintage</b>', font: {family: 'Segoe UI', size: 25, color: 'black', weight: 'bold'}, pad: 5},
            tickfont    : {family: 'Segoe UI', size: 20, color: 'black', weight: 'bold'},
            tickangle   : -90,
            showgrid    : true,
            tickformat  : 'd',
            dtick       : window.unifiedChartLimits.dtickY,
            ticks       : 'outside',
            ticklen     : 5,
            tickcolor   : 'rgba(0, 0, 0, 0)',
            fixedrange  : false,
            range       : [window.unifiedChartLimits.yMin, window.unifiedChartLimits.yMax]
        },
        margin      : {l: 75, r: 0, t: 25, b: 75},
        annotations : buildScatterAnnotations(scatterData, 'over8', 'vintage', 'gr')
    }, {responsive: true, displayModeBar: false});
}

const guessChartDiv = document.getElementById('plotlyGuessChart');

if (guessChartDiv) {
    guessChartDiv.on('plotly_click', function(data) {
        if (!data.points || data.points.length === 0) return;
        const pointData = data.points[0];

        if (pointData.customdata && pointData.customdata[0]) {
            const playerName = String(pointData.customdata[0]).trim().toLowerCase();
            window.searchPlayerMetricFromTable(`correct:${playerName}`);
        }
    });
}

let currentListChartMode        = "ALL"; 
let currentGuessListViewMode    = "ALL";
let isGraphFocused              = false;

if (watched) {
    const glBtn = document.getElementById("guessListToggleBtn");
    const fcBtn = document.getElementById("focusToggleBtn");

    if (glBtn) glBtn.classList.remove("hidden");
    if (fcBtn) fcBtn.classList.remove("hidden");
}

function getLocalizedChartBounds(dataArray, xKey, yKey) {
    if (!dataArray || dataArray.length === 0) return window.unifiedChartLimits;

    const xValues = dataArray.map(d => d[xKey]);
    const yValues = dataArray.map(d => d[yKey]);

    const xMin = Math.min(...xValues) - 0.25;
    const xMax = Math.max(...xValues) + 0.25;
    const yMin = Math.min(...yValues) - 1;
    const yMax = Math.max(...yValues) + 1;

    const dtickY = Math.max(2, Math.ceil((Math.max(...yValues) - Math.min(...yValues)) / 5));
    return {xMin, xMax, yMin, yMax, dtickY};
}

if (document.getElementById('plotlyListChart') && arrowData) {
    window.listDataPool = {
        "ALL": arrowData.map(d => ({
            acronym     : d.acronym,
            name        : d.name,
            x           : d.x_start,
            y           : d.y_start,
            size        : d.rig_rate,
            color       : d.grid_grs || d.rig_gr, 
            hoverText   : `<b>${d.name}</b><br>Rig Over-8: ${d.x_start.toFixed(2)}<br>Rig Vintage: ${d.seasonal_vintage_start}<br>Rig Rate: ${(d.grid_rate !== undefined ? d.grid_rate : d.rig_rate).toFixed(2)}<br>Rig GR: ${d.rig_gr.toFixed(2)}<extra></extra>`
        })),

        "HIT": arrowData.map(d => ({
            acronym     : d.acronym,
            name        : d.name,
            x           : d.x_end, 
            y           : d.y_end,
            size        : d.rig_rate, 
            color       : d.grid_grs || d.rig_gr, 
            hoverText   : `<b>${d.name}</b><br>Hit Rig Over-8: ${d.x_end.toFixed(2)}<br>Hit Rig Vintage: ${d.seasonal_vintage || d.seasonal_vintage_end}<br>Rig Rate: ${(d.grid_rate !== undefined ? d.grid_rate : d.rig_rate).toFixed(2)}<br>Rig GR: ${d.rig_gr.toFixed(2)}<extra></extra>`
        }))
    };

    renderListChart();
}

function updateSongHelpDropdown() {
    const dropdown = document.getElementById("songGuideDropdown");
    if (!dropdown) return;

    dropdown.innerHTML = `
        <p class="font-bold pb-1 mb-1 text-sm text-black">Example</p>
        <hr class="border-black mb-2">
        <p class="text-xs font-normal">
            <b>Difficulty:</b> 25-30<br>
            <b>Vintage:</b> 2015-2020<br>
            <b>Over-8:</b> <b><span style="color: #3232c8;">6.26</span></b><br>
            This means that, on average, <b><span style="color: #3232c8;">6.26/8</span></b> people guessed songs from <b>2015-2020</b> with a difficulty of <b>25-30</b> correctly
        </p>
    `;
}

function updateGuessHelpDropdown() {
    const dropdown = document.getElementById("guessGuideDropdown");
    if (!dropdown) return;

    let guideText = `
        <b>Display</b><br>
        Changes the dataset shown on the bubble chart<br>
        • <b>Corrects:</b> All songs guessed correctly<br>
        • <b>Rigs:</b> All songs from this player's list<br>
        • <b>Rigs Hit:</b> All songs guessed correctly from this player's list<br><br>
        <b>Focus</b><br>
        • <b>On:</b> Scales the axes to fit only the current chart<br>
        • <b>Off:</b> Uses the shared axis bounds for easy comparison between charts
    `;
    
    let exampleText = "";

    if (currentGuessListViewMode === "ALL") exampleText = `
        A <b>small, <span style="color: #3232c8;">blue</span></b> circle in the <b>bottom-left</b> means that, on average, this player:<br>
        • Has low Guess Rate (<b>small</b>), yet<br>
        • Is over-performing their Elo (<b><span style="color: #3232c8;">blue</span></b>),<br>
        • Usually hits harder (<b>left</b>) songs, and<br>
        • Prefers the older (<b>bottom</b>) ones
    `;

    else if (currentGuessListViewMode === "RIG") exampleText = `
        A <b>big, <span style="color: #c83232;">red</span></b> circle in the <b>top-right</b> means that, on average, this player's list:<br>
        • Usually has newer (<b>top</b>) songs,<br>
        • Appears a lot (<b>big</b>),<br>
        • Is difficult for the player (<b><span style="color: #c83232;">red</span></b>), yet<br>
        • Easy for others (<b>right</b>)
    `;

    else if (currentGuessListViewMode === "HIT") exampleText = `
        A <b>big, <span style="color: #c83232;">red</span></b> circle in the <b>top-right</b> means that, on average, this player:<br>
        • Focuses heavily on newer (<b>top</b>) songs from their list,<br>
        • Said list appears a lot (<b>big</b>),<br>
        • Is difficult (<b><span style="color: #c83232;">red</span></b>) for them to get right, yet<br>
        • Easy for others (<b>right</b>)
    `;

    dropdown.innerHTML = `
        <p class="font-bold pb-1 mb-1">Guide</p>
        <hr class="border-black mb-2">
        <p class="mb-2 text-xs">${guideText}</p>
        <p class="font-bold pb-1 mb-1 mt-3">Example</p>
        <hr class="border-black mb-2">
        <p class="text-xs font-normal">${exampleText}</p>
    `;
}

function updateSearchHelpDropdown() {
    const dropdown = document.getElementById("searchGuideDropdown");
    if (!dropdown) return;

    dropdown.innerHTML = `
        <p class="font-bold border-b pb-1 mb-1">Guide</p>
        <p class="mb-2 text-xs">
            Search using <code class="bg-gray-200 px-1 rounded font-mono text-xs">value</code> or <code class="bg-gray-200 px-1 rounded font-mono text-xs">columnname:value</code><br>
            You can replace <code class="bg-gray-200 px-1 rounded font-mono text-xs">:</code> with arithmetic operators (<code class="bg-gray-200 px-1 rounded font-mono text-xs">=, !:, !=, &lt;, &gt;, &lt;=, &gt;=)</code><br>
            Combine query terms using explicit <code class="bg-gray-200 px-1 rounded font-mono text-xs">and/or</code> keywords<br>
            Group precedence with <code class="bg-gray-200 px-1 rounded font-mono text-xs">(brackets)</code><br>
            Wrap multi-word values in <code class="bg-gray-200 px-1 rounded font-mono text-xs">"double-quotes"</code><br><br>
            <code class="bg-gray-200 px-1 rounded font-mono text-xs">(artist="aoi koga" or artist:"tomori kusunoki") and difficulty>20</code><br>
            returns songs by Aoi Koga or Tomori Kusunoki with difficulties above 20
        </p>
        <div class="grid grid-cols-2 gap-2 pt-1 border-t text-xs">
            <div>
                <span class="block font-mono mb-1 font-bold text-gray-700">anime, songtype, chanting, animetype, song, artist, composer, arranger, correct, list</span>
                <code class="bg-gray-200 px-1 rounded font-mono">anime:aikatsu artist:nanase</code> returns Nanase songs from Aikatsu<br>
                <code class="bg-gray-200 px-1 rounded font-mono">songtype:ed animetype:movie</code> returns movie ending songs<br>
                <code class="bg-gray-200 px-1 rounded font-mono">correct!=furlain</code> returns songs FurLain missed<br>
            </div>
            <div>
                <span class="block font-mono mb-1 font-bold text-gray-700">vintage, difficulty, correct, list</span>
                <code class="bg-gray-200 px-1 rounded font-mono">vintage&lt;"Summer 2023"</code> returns songs from anime before Summer 2023<br>
                <code class="bg-gray-200 px-1 rounded font-mono">difficulty&lt;30 list&gt;4</code> returns songs with difficulty less than 30 listed by more than 4 people<br>
                <code class="bg-gray-200 px-1 rounded font-mono">correct:0</code> returns songs no one got right<br>
            </div>
        </div>
    `;
}

function updateGuessChartAxesFocus() {
    const targetChart = document.getElementById('plotlyGuessChart');
    if (!targetChart || !scatterData) return;
    const bounds = isGraphFocused ? getLocalizedChartBounds(scatterData, 'over8', 'vintage') : window.unifiedChartLimits;

    Plotly.relayout(targetChart, {
        'xaxis.range': [bounds.xMin, bounds.xMax],
        'yaxis.range': [bounds.yMin, bounds.yMax],
        'yaxis.dtick': bounds.dtickY
    });
}

function renderListChart() {
    if (!window.listDataPool || !window.listDataPool[currentListChartMode]) return;

    const activeScatterSource   = window.listDataPool[currentListChartMode];
    const listHull              = get75PercentileHull(activeScatterSource, 'x', 'y');
    let listTraces              = [];

    if (listHull) listTraces.push({
        x           : listHull.x,
        y           : listHull.y,
        type        : 'scatter',
        mode        : 'lines',
        line        : {color: 'black', width: 0.5, dash: 'solid'},
        hoverinfo   : 'skip',
        showlegend  : false
    });

    listTraces.push({
        x               : activeScatterSource.map(d => d.x),
        y               : activeScatterSource.map(d => d.y),
        text            : activeScatterSource.map(d => d.acronym),
        hovertemplate   : activeScatterSource.map(d => d.hoverText),
        hoverlabel      : {align: 'left', font: {family: 'Segoe UI', size: 15}},
        mode            : 'markers',
        showlegend      : false,
        marker          : {
            size        : activeScatterSource.map(d => Math.max(10, d.size * 1.5)),
            opacity     : 0.95,
            color       : activeScatterSource.map(d => d.color),
            colorscale  : [[0, hexToRgba(c0)], [0.7, hexToRgba(c0)], [0.8, hexToRgba(c1)], [0.9, hexToRgba(c2)], [1, hexToRgba(c2)]],
            showscale   : true,
            colorbar    : {
                title       : {text: '<b>Rig GR</b>', font: {family: 'Segoe UI', size: 25, color: 'black', weight: 'bold'}, side: 'right'},
                thickness   : 25,
                len         : 1.0,
                y           : 0.5,
                yanchor     : 'middle',
                x           : 1,
                tickmode    : 'array',
                tickvals    : [0, 70, 80, 90, 100],
                ticktext    : ['0', '70', '80', '90', '100'],
                tickfont    : {family: 'Segoe UI', size: 20, color: 'black', weight: 'bold'}
            },
            line        : {color: 'black', width: 0},
            cmin        : 0,
            cmax        : 100
        }
    });

    const currentBounds = isGraphFocused ? getLocalizedChartBounds(activeScatterSource, 'x', 'y') : window.unifiedChartLimits;

    Plotly.newPlot('plotlyListChart', listTraces, {
        font        : {family: 'Segoe UI'},
        xaxis       : {
            title       : {text: '<b>Over-8</b>', font: {family: 'Segoe UI', size: 25, color: 'black', weight: 'bold'}, pad: 5},
            tickfont    : {family: 'Segoe UI', size: 20, color: 'black', weight: 'bold'},
            showgrid    : true,
            tickformat  : '.1f',
            dtick       : 0.5,
            ticks       : 'outside',
            ticklen     : 5,
            tickcolor   : 'rgba(0, 0, 0, 0)',
            fixedrange  : false,
            range       : [currentBounds.xMin, currentBounds.xMax]
        },
        yaxis       : {
            title       : {text: '<b>Vintage</b>', font: {family: 'Segoe UI', size: 25, color: 'black', weight: 'bold'}, pad: 5},
            tickfont    : {family: 'Segoe UI', size: 20, color: 'black', weight: 'bold'},
            tickangle   : -90,
            showgrid    : true,
            tickformat  : 'd',
            dtick       : currentBounds.dtickY,
            ticks       : 'outside',
            ticklen     : 5,
            tickcolor   : 'rgba(0, 0, 0, 0)',
            fixedrange  : false,
            range       : [currentBounds.yMin, currentBounds.yMax]
        },
        margin      : {l: 75, r: 0, t: 25, b: 75},
        annotations : buildScatterAnnotations(activeScatterSource, 'x', 'y', 'size')
    }, {responsive: true, displayModeBar: false});

    const listChartDiv = document.getElementById('plotlyListChart');

    if (listChartDiv) {
        listChartDiv.on('plotly_click', function(data) {
            if (!data.points || data.points.length === 0) return;

            const ptIndex           = data.points[0].pointIndex;
            const activeDataSource  = window.listDataPool[currentListChartMode];

            if (activeDataSource && activeDataSource[ptIndex]) {
                const playerName = String(activeDataSource[ptIndex].name).trim().toLowerCase();

                if (currentListChartMode === "HIT") window.searchPlayerMetricFromTable(`list:${playerName} correct:${playerName}`);
                else                                window.searchPlayerMetricFromTable(`list:${playerName}`);
            }
        });
    }
}

if (document.getElementById('plotlyListChart') && arrowData) {
    window.listDataPool = {
        "ALL": arrowData.map(d => ({
            acronym     : d.acronym,
            name        : d.name,
            x           : d.x_start,
            y           : d.y_start,
            size        : d.rig_rate,
            color       : d.grid_grs || d.rig_gr, 
            hoverText   : `<b>${d.name}</b><br>Rig Over-8: ${d.x_start.toFixed(2)}<br>Rig Vintage: ${d.seasonal_vintage_start}<br>Rig Rate: ${(d.grid_rate !== undefined ? d.grid_rate : d.rig_rate).toFixed(2)}<br>Rig GR: ${d.rig_gr.toFixed(2)}<extra></extra>`
        })),

        "HIT": arrowData.map(d => ({
            acronym     : d.acronym,
            name        : d.name,
            x           : d.x_end, 
            y           : d.y_end,
            size        : d.rig_rate, 
            color       : d.grid_grs || d.rig_gr, 
            hoverText   : `<b>${d.name}</b><br>Hit Rig Over-8: ${d.x_end.toFixed(2)}<br>Hit Rig Vintage: ${d.seasonal_vintage || d.seasonal_vintage_end}<br>Rig Rate: ${(d.grid_rate !== undefined ? d.grid_rate : d.rig_rate).toFixed(2)}<br>Rig GR: ${d.rig_gr.toFixed(2)}<extra></extra>`
        }))
    };

    renderListChart();
}

let globalSortState     = {columnName: "Anime", ascending: true};
let currentSearchLang   = "JP";

const searchHeadersConfig = [
    {id: "anime", name: "Anime", visible: true, type: "text"},

    {
        id          : "type",
        name        : "Song Type",
        visible     : true,
        type        : "categorical",
        subOptions  : ["Opening", "Ending", "Insert"]
    },
    {
        id          : "chanting",
        name        : "Chanting",
        visible     : false,
        type        : "categorical",
        subOptions  : ["Yes", "No"]
    },
    {
        id          : "anime_type",
        name        : "Anime Type", visible: false,
        type        : "categorical",
        subOptions  : ["TV", "Movie", "OVA", "ONA", "Special"]
    },
    {
        id          : "vintage",
        name        : "Vintage",
        visible     : false,
        type        : "range",
        min         : 1900,
        max         : 2026,
        step        : 1
    },
    {
        id          : "difficulty",  
        name        : "Difficulty",
        visible     : false,
        type        : "range",
        min         : 0,
        max         : 100,
        step        : 1
    },
    {id: "song",        name: "Song",       visible: true,  type: "text"},
    {id: "artist",      name: "Artist",     visible: true,  type: "text"},
    {id: "composer",    name: "Composer",   visible: false, type: "text"},
    {id: "arranger",    name: "Arranger",   visible: false, type: "text"},
    {
        id          : "guessers",
        name        : "Correct",
        visible     : true,
        type        : "range",
        min         : 0,
        max         : 8,
        step        : 1
    },
    {
        id          : "listers",     
        name        : "List",       
        visible     : false, 
        type        : "range",        
        min         : 0,
        max         : 8,
        step        : 1
    }
];

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

    if      (remainder < 0.25)  season = "Winter";
    else if (remainder < 0.50)  season = "Spring";
    else if (remainder < 0.75)  season = "Summer";
    else                        season = "Fall";
    
    return `${season}&nbsp;${year}`;
}

window.translateHoverText = function(textArray) {
    if (!globalSearchData || globalSearchData.length === 0) return textArray;

    return textArray.map(line => {
        if (typeof line !== 'string') return line;
        let translatedLine = line;

        for (let i = 0; i < globalSearchData.length; i++) {
            const s     = globalSearchData[i];
            const jp    = s.romaji  || "";
            const en    = s.english || "";

            if (!jp || !en || jp === en) continue;

            if (currentSearchLang === "EN") {
                if (translatedLine.includes(jp + " (OP") || translatedLine.includes(jp + " (ED") || translatedLine.includes(jp + " (IN")) {
                    translatedLine = translatedLine.replace(jp + " (", en + " (");
                    break;
                }
            }

            else {
                if (translatedLine.includes(en + " (OP") || translatedLine.includes(en + " (ED") || translatedLine.includes(en + " (IN")) {
                    translatedLine = translatedLine.replace(en + " (", jp + " (");
                    break;
                }
            }
        }

        return translatedLine;
    });
};

window.toggleSearchLanguage = function() {
    const btn           = document.getElementById("langToggleBtn");
    currentSearchLang   = currentSearchLang === "JP" ? "EN" : "JP";
    btn.innerText       = currentSearchLang;

    sortSearchData      ();
    triggerTableRefresh ();
    renderTierCharts    ();
};

function initColumnSettingsCheckboxes() {
    const container = document.getElementById("columnCheckboxContainer");
    const masterChk = document.getElementById("allColumnsMasterCheckbox");

    if (!container || !masterChk) return;
    container.innerHTML = "";

    if (globalSearchData && globalSearchData.length > 0) {
        const parsedVints = globalSearchData.map(s => s._vintageParsed) .filter(v => !isNaN(v) && v !== -Infinity);
        const parsedDiffs = globalSearchData.map(s => s._diffParsed)    .filter(d => !isNaN(d) && d !== -Infinity);

        const vintConfig  = searchHeadersConfig.find(c => c.id === "vintage");
        const diffConfig  = searchHeadersConfig.find(c => c.id === "difficulty");

        if (vintConfig && parsedVints.length > 0) {
            vintConfig.min = Math.floor (Math.min(...parsedVints));
            vintConfig.max = Math.ceil  (Math.max(...parsedVints));
        }

        if (diffConfig && parsedDiffs.length > 0) {
            diffConfig.min = 0;
            const maxDiff  = Math.max(...parsedDiffs);
            diffConfig.max = Math.ceil(maxDiff / 5) * 5; 
        }
    }

    function updateMasterCheckboxState() {
        const allChecked        = searchHeadersConfig.every(c => c.visible);
        const noneChecked       = searchHeadersConfig.every(c => !c.visible);
        masterChk.checked       = allChecked;
        masterChk.indeterminate = !allChecked && !noneChecked;
    }

    masterChk.addEventListener("change", () => {
        searchHeadersConfig.forEach(c => {
            c.visible = masterChk.checked;

            if (c.type === "categorical" && c.subOptions) {
                if (masterChk.checked)  c.selectedOptions = new Set(c.subOptions.map(o => o.toLowerCase()));
                else                    c.selectedOptions.clear();
            }
        });

        container.querySelectorAll("input[type='checkbox']").forEach(chk => {chk.checked = masterChk.checked;});
        triggerTableRefresh();
    });

    searchHeadersConfig.forEach(col => {
        const colWrapper        = document.createElement("div");
        colWrapper.className    = "flex flex-col space-y-1";

        const label     = document.createElement("label");
        label.className = "flex items-center gap-2 cursor-pointer w-full text-left font-bold";

        const chk       = document.createElement("input");
        chk.type        = "checkbox";
        chk.className   = "col-toggle-checkbox rounded accent-black";
        chk.checked     = col.visible;

        chk.addEventListener("change", () => {
            col.visible = chk.checked;

            if (col.type === "categorical" && col.subOptions) {
                if (chk.checked) {
                    col.subOptions.forEach(opt => col.selectedOptions.add(opt.toLowerCase()));
                    colWrapper.querySelectorAll("input[type='checkbox']").forEach(subChk => {subChk.checked = true;});
                }

                else {
                    col.selectedOptions.clear();
                    colWrapper.querySelectorAll("input[type='checkbox']").forEach(subChk => {subChk.checked = false;});
                }
            }

            else if (!chk.checked && col.type === "range") {
                col.currentMin = col.min;
                col.currentMax = col.max;

                const inputs = colWrapper.querySelectorAll("input[type='number']");

                if (inputs.length === 2) {
                    inputs[0].value = col.min;
                    inputs[1].value = col.max;
                }
            }

            updateMasterCheckboxState   ();
            triggerTableRefresh         ();
        });

        label       .appendChild(chk);
        label       .appendChild(document.createTextNode(col.name));
        colWrapper  .appendChild(label);

        if (col.type === "categorical" && col.subOptions) {
            const subContainer = document.createElement("div");
            subContainer.className = "pl-6 flex flex-col text-xs";

            if (col.visible)    col.selectedOptions = new Set(col.subOptions.map(o => o.toLowerCase()));
            else                col.selectedOptions = new Set();

            col.subOptions.forEach(opt => {
                const subLabel      = document.createElement("label");
                subLabel.className  = "flex items-center gap-1 cursor-pointer text-gray-700 hover:text-black py-0.5";

                const subChk        = document.createElement("input");
                subChk.type         = "checkbox";
                subChk.className    = "rounded accent-black scale-90";
                subChk.checked      = col.visible;

                subChk.addEventListener("change", () => {
                    if (subChk.checked) col.selectedOptions.add     (opt.toLowerCase());
                    else                col.selectedOptions.delete  (opt.toLowerCase());

                    const totalChildren = col.subOptions.length;
                    const selectedCount = col.selectedOptions.size;

                    if (selectedCount === totalChildren) {
                        chk.checked         = true;
                        chk.indeterminate   = false;
                        col.visible         = true;
                    }

                    else if (selectedCount === 0) {
                        chk.checked         = false;
                        chk.indeterminate   = false;
                        col.visible         = false;
                    }

                    else {
                        chk.checked         = false;
                        chk.indeterminate   = true;
                        col.visible         = true;
                    }

                    updateMasterCheckboxState   ();
                    triggerTableRefresh     ();
                });

                subLabel        .appendChild(subChk);
                subLabel        .appendChild(document.createTextNode(opt));
                subContainer    .appendChild(subLabel);
            });

            colWrapper.appendChild(subContainer);
        }

        if (col.type === "range") {
            col.currentMin = col.min;
            col.currentMax = col.max;

            const inputContainer        = document.createElement("div");
            inputContainer.className    = "pl-6 flex flex-col gap-1 mt-1 w-full text-black";

            const boxWrapper            = document.createElement("div");
            boxWrapper.className        = "flex flex-col gap-1";

            const minLabel      = document.createElement("label");
            minLabel.className  = "flex items-center justify-start gap-2 font-mono"; 
            minLabel.innerHTML  = "Min:";
            const inputMin      = document.createElement("input");
            inputMin.type       = "number";
            inputMin.min        = col.min;
            inputMin.max        = col.max;
            inputMin.value      = col.min;
            inputMin.className  = "w-10 h-5 border text-center text-xs";

            const maxLabel      = document.createElement("label");
            maxLabel.className  = "flex items-center justify-start gap-2 font-mono";
            maxLabel.innerHTML  = "Max:";
            const inputMax      = document.createElement("input");
            inputMax.type       = "number";
            inputMax.min        = col.min;
            inputMax.max        = col.max;
            inputMax.value      = col.max;
            inputMax.className  = "w-10 h-5 border text-center text-xs";

            minLabel        .appendChild(inputMin);
            maxLabel        .appendChild(inputMax);
            boxWrapper      .appendChild(minLabel);
            boxWrapper      .appendChild(maxLabel);
            inputContainer  .appendChild(boxWrapper);

            const handleTextbookInput = () => {
                let valMin = parseInt(inputMin.value);
                let valMax = parseInt(inputMax.value);

                if (isNaN(valMin)) valMin = col.min;
                if (isNaN(valMax)) valMax = col.max;

                if (valMin < col.min) valMin = col.min;
                if (valMax > col.max) valMax = col.max;

                if (valMin > valMax) {
                    if (document.activeElement === inputMin) {
                        valMin          = valMax;
                        inputMin.value  = valMin;
                    }

                    else {
                        valMax          = valMin;
                        inputMax.value  = valMax;
                    }
                }

                col.currentMin = valMin;
                col.currentMax = valMax;
            };

            const triggerRefreshDebounced = debounce(() => {
                if (!col.visible && (col.currentMin > col.min || col.currentMax < col.max)) {
                    col.visible = true;
                    chk.checked = true;

                    updateMasterCheckboxState();
                }

                triggerTableRefresh();
            }, 150);

            inputMin.addEventListener("input", () => { handleTextbookInput(); triggerRefreshDebounced(); });
            inputMax.addEventListener("input", () => { handleTextbookInput(); triggerRefreshDebounced(); });
            
            masterChk.addEventListener("change", () => {
                if (!masterChk.checked) {
                    inputMin.value = col.min;
                    inputMax.value = col.max;
                    col.currentMin = col.min;
                    col.currentMax = col.max;
                }

                else {
                    let valMin = parseInt(inputMin.value);
                    let valMax = parseInt(inputMax.value);

                    col.currentMin = isNaN(valMin) ? col.min : valMin;
                    col.currentMax = isNaN(valMax) ? col.max : valMax;
                }

                triggerTableRefresh();
            });

            colWrapper.appendChild(inputContainer);
        }

        if (col.type === "categorical" && col.subOptions) {
            const totalChildren = col.subOptions.length;
            const selectedCount = col.selectedOptions.size;

            if (selectedCount > 0 && selectedCount < totalChildren) {
                chk.checked         = false;
                chk.indeterminate   = true;
            }
        }

        container.appendChild(colWrapper);
    });

    updateMasterCheckboxState();
}

function triggerTableRefresh() {
    const searchInput = document.getElementById('songSearchInput');

    if (searchInput && searchInput.value.trim())    searchInput.dispatchEvent(new Event('input-direct'));
    else                                            renderSearchTable(globalSearchData);
}

function sortSearchData() {
    const {columnName, ascending}   = globalSortState;
    const isJp                      = currentSearchLang === "JP";

    globalSearchData.sort((a, b) => {
        let valA, valB;

        switch (columnName) {
            case "Anime"        : valA = isJp ? a._romajiLower : a._englishLower;   valB = isJp ? b._romajiLower : b._englishLower; break;
            case "Song Type"    : valA = a._typeLower;                              valB = b._typeLower;                            break;
            case "Anime Type"   : valA = a._animeTypeLower;                         valB = b._animeTypeLower;                       break;
            case "Song"         : valA = a._songLower;                              valB = b._songLower;                            break;
            case "Artist"       : valA = a._artistRawLower;                         valB = b._artistRawLower;                       break;
            case "Composer"     : valA = a._composerLower;                          valB = b._composerLower;                        break;
            case "Arranger"     : valA = a._arrangerLower;                          valB = b._arrangerLower;                        break;
            case "Chanting"     : valA = a._chantingLower;                          valB = b._chantingLower;                        break;
            case "Vintage"      : valA = a._vintageParsed;                          valB = b._vintageParsed;                        break;
            case "Genre"        : valA = a._genresCount;                            valB = b._genresCount;                          break;
            case "Tag"          : valA = a._tagsCount;                              valB = b._tagsCount;                            break;
            case "Difficulty"   : valA = a._diffParsed;                             valB = b._diffParsed;                           break;
            case "Correct"      : valA = a._guessersCount;                          valB = b._guessersCount;                        break;
            case "List"         : valA = a._listersCount;                           valB = b._listersCount;                         break;
            default: return 0;
        }

        if (valA < valB) return ascending ? -1 : 1;
        if (valA > valB) return ascending ? 1 : -1;
        return 0;
    });
}

window.handleSearchSort = function(columnHeaderName) {
    if (globalSortState.columnName === columnHeaderName) globalSortState.ascending = !globalSortState.ascending;

    else {
        globalSortState.columnName  = columnHeaderName;
        globalSortState.ascending   = true;
    }
    sortSearchData();
    triggerTableRefresh();
};

function matchNumericConstraint(numTarget, operator, numCrit) {
    if (isNaN(numTarget) || isNaN(numCrit))     return false;
    if (operator === ":" || operator === "=")   return Math.floor(numTarget) === Math.floor(numCrit);

    switch (operator) {
        case "<"    : return numTarget <    numCrit;
        case ">"    : return numTarget >    numCrit;
        case "<="   : return numTarget <=   numCrit;
        case ">="   : return numTarget >=   numCrit;
        default     : return false;
    }
}

function evaluateQuery(song, key, operator, value) {
    if (key === "anime" || key === "japanese" || key === "english") {
        const titleTarget = currentSearchLang === "JP" ? song._romajiLower : song._englishLower;
        return titleTarget.includes(value);
    }

    switch (key) {
        case "song"         : return song._songLower        .includes(value);
        case "artist"       : return song._artistRawLower   .includes(value);
        case "composer"     : return song._composerLower    .includes(value);
        case "arranger"     : return song._arrangerLower    .includes(value);
        case "animetype"    : return song._animeTypeLower   .includes(value);
        case "chanting"     : return song._chantingLower    .includes(value);

        case "genre"        : return song._genresLower  .some(g => g.includes(value));
        case "tag"          : return song._tagsLower    .some(t => t.includes(value));

        case "songtype": {
            let targetValue = value;

            if      (value === "op") targetValue = "opening";
            else if (value === "ed") targetValue = "ending";
            else if (value === "in") targetValue = "insert";

            return song._typeLower.includes(targetValue);
        }

        case "seen": {
            const isInRoom = song._roomPlayersLower.some(p => p.includes(value));
            return (operator === "!:" || operator === "!=") ? !isInRoom : isInRoom;
        }

        case "lifetaken": {
            const wasInRoom = song._roomPlayersLower    .some(p => p.includes(value));
            const gotLife   = song._livesTakenLower     .some(p => p.includes(value));
            if (operator === "!:" || operator === "!=") return wasInRoom && !gotLife;
            return gotLife;
        }

        case "lifesaved": {
            const wasInRoom = song._roomPlayersLower    .some(p => p.includes(value));
            const savedLife = song._livesSavedLower     .some(p => p.includes(value));
            if (operator === "!:" || operator === "!=") return wasInRoom && !savedLife;
            return savedLife;
        }

        case "correctteam": {
            const teamLine = (song.guessers_hover || []).find(line => line.toLowerCase().includes(value));
            if (!teamLine) return false;
            const isNone = teamLine.toLowerCase().includes(": none");
            if (operator === "!:" || operator === "!=") return isNone;
            return !isNone;
        }

        case "sweep": {
            const tids      = song.correct_teams_flat || [];
            const isSweep   = tids.length === 4 && tids.every(id => id === tids[0]);
            return (value === "yes" || value === "true") ? isSweep : !isSweep;
        }
    }

    if (key === "guessers" || key === "listers") {
        const targetArray = (key === "guessers") ? (song.guessers_flat || []) : (song.listers_flat || []);

        if (isNaN(value)) {
            const hasMatch = targetArray.some(name => name._cachedLower ? name._cachedLower.includes(value) : name.toLowerCase().includes(value));

            if (operator === "!:" || operator === "!=") {
                const roomArray = song.room_players || [];
                const wasInRoom = roomArray.some(name => name.toLowerCase().includes(value));

                return wasInRoom && !hasMatch;
            }

            return hasMatch;

        } else return matchNumericConstraint(targetArray.length, operator, parseFloat(value));
    }

    if (key === "difficulty") return matchNumericConstraint(song._diffParsed, operator, parseFloat(value));

    if (key === "vintage") {
        if (operator === ":" || operator === "=") return song._vintageLower.includes(value);
        return matchNumericConstraint(song._vintageParsed, operator, parseVintageToFloat(value));
    }

    return false;
}

function renderSearchTable(filteredSongs) {
    let runtimeFilteredList = filteredSongs.filter(song => {
        for (let col of searchHeadersConfig) {
            if (col.type === "categorical" && col.selectedOptions) {
                if (!col.visible) continue; 
                let matchValue = "";

                if (col.id === "type") {
                    const tStr = song._typeLower;

                    if      (tStr.includes("opening"))  matchValue = "opening";
                    else if (tStr.includes("ending"))   matchValue = "ending";
                    else if (tStr.includes("insert"))   matchValue = "insert";
                }

                else if (col.id === "chanting")     matchValue = song._chantingLower;
                else if (col.id === "anime_type")   matchValue = song._animeTypeLower;

                if (matchValue && !col.selectedOptions.has(matchValue)) return false;
            }

            if (col.type === "range" && col.currentMin !== undefined && col.currentMax !== undefined) {
                let targetNum = 0;

                if      (col.id === "vintage")      targetNum = song._vintageParsed;
                else if (col.id === "difficulty")   targetNum = song._diffParsed === -Infinity ? 0 : song._diffParsed;
                else if (col.id === "guessers")     targetNum = song._guessersCount;
                else if (col.id === "listers")      targetNum = song._listersCount;

                if (isNaN(targetNum) || targetNum < col.currentMin || targetNum > col.currentMax) return false;
            }
        }
        return true;
    });

    const table         = document.getElementById('searchSongsTable');
    const counterNode   = document.getElementById('searchCounter');

    if (!table)         return;
    if (counterNode)    counterNode.innerText = `${runtimeFilteredList.length}/${globalSearchData.length}`;

    const activeCols = searchHeadersConfig.filter(c => c.visible);

    if (activeCols.length === 0) {
        table.innerHTML = `<thead><tr><th>Error</th></tr></thead><tbody><tr><td class="p-2 text-center text-black">Select at least 1 column</td></tr></tbody>`;
        return;
    }

    if (runtimeFilteredList.length === 0) {
        table.innerHTML = `<thead><tr><th>Error</th></tr></thead><tbody><tr><td class="p-2 text-center text-black">No songs matched your specific constraints</td></tr></tbody>`;
        return;
    }

    let theadStr = "<tr>" + activeCols.map(c => {
        const isCurrentSort = (globalSortState.columnName === c.name);
        const indicator     = isCurrentSort ? (globalSortState.ascending ? "▴" : "▾")   : "▸";
        const activeStyles  = isCurrentSort ? 'background-color: black; color: white; ' : '';

        return `<th class="cursor-pointer select-none" style="${activeStyles}white-space: nowrap;" data-header-name="${c.name}">${c.name}${indicator}</th>`;
    }).join('') + "</tr>";
    
    table.innerHTML = `<thead>${theadStr}</thead><tbody></tbody>`;
    const tbody     = table.tBodies[0];

    table.querySelectorAll('thead th').forEach(th => {th.addEventListener('click', () => handleSearchSort(th.getAttribute('data-header-name')));});

    const fragment  = document.createDocumentFragment();
    const isJp      = currentSearchLang === "JP";

    const compVisible = activeCols.some(c => c.id === "composer");
    const arrVisible  = activeCols.some(c => c.id === "arranger");

    runtimeFilteredList.forEach(song => {
        const tr            = document.createElement('tr');
        let skipComposer    = false;
        let skipArranger    = false;

        activeCols.forEach(col => {
            if (col.id === "composer" && skipComposer) return;
            if (col.id === "arranger" && skipArranger) return;

            const td = document.createElement('td');

            switch (col.id) {
                case "anime": {
                    td.className = "text-left search-c2-text";
                    const a         = document.createElement('a');
                    a.href          = song.ann_url;
                    a.target        = "_blank";
                    a.className     = "hover:underline";
                    a.textContent   = isJp ? song.romaji : song.english;

                    td.appendChild(a);
                    break;
                }

                case "type":
                    td.className        = "text-center font-normal text-black";
                    td.style.whiteSpace = "nowrap";
                    td.textContent      = song.type;
                    break;

                case "chanting":
                    td.className    = "text-center font-normal text-black";
                    td.textContent  = song.chanting;
                    break;

                case "anime_type":
                    td.className    = "text-center font-normal text-black";
                    td.textContent  = song.anime_type;
                    break;

                case "vintage":
                    td.className        = "text-center font-normal text-black";
                    td.style.whiteSpace = "nowrap";
                    td.textContent      = song.vintage;
                    break;

                case "genre": {
                    const hasGenres = song.genres_raw && song.genres_raw.length > 0;
                    td.className    = hasGenres ? "cursor-help hover:bg-gray-100 text-center text-black font-normal" : "text-center text-black font-normal";
                    if (hasGenres) {
                        let sortedGenres = [...song.genres_raw].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
                        td.setAttribute("data-songs", encodeURIComponent(JSON.stringify(sortedGenres)));
                    }
                    td.textContent = song._genresCount;
                    break;
                }

                case "tag": {
                    const hasTags = song.tags_raw && song.tags_raw.length > 0;
                    td.className  = hasTags ? "cursor-help hover:bg-gray-100 text-center text-black font-normal" : "text-center text-black font-normal";
                    if (hasTags) {
                        let sortedTags = [...song.tags_raw].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
                        td.setAttribute("data-songs", encodeURIComponent(JSON.stringify(sortedTags)));
                    }
                    td.textContent = song._tagsCount;
                    break;
                }

                case "difficulty":
                    td.className    = "text-center font-normal font-mono text-black";
                    td.textContent  = song.difficulty;
                    break;

                case "song": {
                    td.className    = "text-left search-c2-text";
                    const a         = document.createElement('a');
                    a.href          = song.video_url;
                    a.target        = "_blank";
                    a.className     = "hover:underline";
                    a.textContent   = song.song;

                    td.appendChild(a);
                    break;
                }

                case "artist": {
                    const matchComp = compVisible   && (song.artist_raw === song.composer);
                    const matchArr  = arrVisible    && (song.composer   === song.arranger);

                    let flatString      = Array.isArray(song.artist_arr) ? song.artist_arr.join('||') : String(song.artist_arr || '');
                    let splitArtists    = flatString.replace(/\s*(?:,|\b(?<!\d)\/(?!\d)\b|・|&|×|\bfeat\.)\s*/gi, '||').split('||').map(x => x.trim()).filter(Boolean);
                    const isOverflown   = splitArtists.length > 3;

                    if (matchComp && matchArr) {
                        td.colSpan      = 3;
                        td.className    = "text-left text-black font-normal";
                        td.textContent  = trimNames(song.artist_arr || []);
                        skipComposer    = true;
                        skipArranger    = true;
                    }

                    else if (matchComp) {
                        td.colSpan      = 2;
                        td.className    = "text-left text-black font-normal";
                        td.textContent  = trimNames(song.artist_arr || []);
                        skipComposer    = true;
                    }

                    else {
                        td.className        = isOverflown ? "cursor-help hover:bg-gray-100 text-left text-black font-normal" : "text-left text-black font-normal";
                        if (isOverflown) td.setAttribute("data-songs", encodeURIComponent(JSON.stringify(splitArtists)));
                        td.textContent      = trimNames(song.artist_arr || []);
                    }

                    break;
                }

                case "composer": {
                    const matchArr = arrVisible && (song.composer === song.arranger);
                    
                    let flatComp        = Array.isArray(song.composer) ? song.composer.join('||') : String(song.composer || '');
                    let splitComps      = flatComp.replace(/\s*(?:,|\b(?<!\d)\/(?!\d)\b|・|&|×|\bfeat\.)\s*/gi, '||').split('||').map(x => x.trim()).filter(Boolean);
                    const isOverflown   = splitComps.length > 3;

                    if (matchArr) {
                        td.colSpan      = 2;
                        td.className    = "text-left font-normal text-black";
                        td.textContent  = trimNames(song.composer);
                        skipArranger    = true;
                    }

                    else {
                        td.className    = isOverflown ? "cursor-help hover:bg-gray-100 text-left text-black font-normal" : "text-left font-normal text-black";
                        if (isOverflown) td.setAttribute("data-songs", encodeURIComponent(JSON.stringify(splitComps)));
                        td.textContent  = trimNames(song.composer);
                    }

                    break;
                }

                case "arranger": {
                    let flatArr         = Array.isArray(song.arranger) ? song.arranger.join('||') : String(song.arranger || '');
                    let splitArrs       = flatArr.replace(/\s*(?:,|\b(?<!\d)\/(?!\d)\b|・|&|×|\bfeat\.)\s*/gi, '||').split('||').map(x => x.trim()).filter(Boolean);
                    const isOverflown   = splitArrs.length > 3;

                    td.className    = isOverflown ? "cursor-help hover:bg-gray-100 text-left text-black font-normal" : "text-left font-normal text-black";
                    if (isOverflown) td.setAttribute("data-songs", encodeURIComponent(JSON.stringify(splitArrs)));
                    td.textContent  = trimNames(song.arranger);
                    break;
                }

                case "guessers": {
                    const hasGuesses    = song.guessers_hover && song.guessers_hover.length > 0;
                    td.className        = hasGuesses ? "cursor-help hover:bg-gray-100 text-center text-black font-normal" : "text-center text-black font-normal";
                    if (hasGuesses) td.setAttribute("data-songs", encodeURIComponent(JSON.stringify(song.guessers_hover)));
                    td.textContent      = song._guessersCount;
                    break;
                }

                case "listers": {
                    const hasLists  = song.listers_hover && song.listers_hover.length > 0;
                    td.className    = hasLists ? "cursor-help hover:bg-gray-100 text-center text-black font-normal" : "text-center text-black font-normal";
                    if (hasLists) td.setAttribute("data-songs", encodeURIComponent(JSON.stringify(song.listers_hover)));
                    td.textContent  = song._listersCount;
                    break;
                }
            }

            tr.appendChild(td);
        });

        fragment.appendChild(tr);
    });

    tbody.appendChild(fragment);
    setupTooltipListeners();
}

fetch('Search.json')
    .then(res => res.json())

    .then(searchJson => {
        globalSearchData = searchJson.map(song => {
            if (!song.guessers_flat && song.guessers_hover) song.guessers_flat  = song.guessers_hover.map(g => g.split(' (')[0]);
            if (!song.listers_flat  && song.listers_hover)  song.listers_flat   = song.listers_hover;

            song._romajiLower       = (song.romaji              || "").toLowerCase();
            song._englishLower      = (song.english             || "").toLowerCase();
            song._songLower         = (song.song                || "").toLowerCase();
            song._artistRawLower    = (song.artist_raw          || "").toLowerCase();
            song._composerLower     = (song.composer            || "").toLowerCase();
            song._arrangerLower     = (song.arranger            || "").toLowerCase();
            song._typeLower         = (song.type                || "").toLowerCase();
            song._vintageLower      = (song.vintage             || "").toLowerCase();
            song._animeTypeLower    = (song.anime_type          || "").toLowerCase();
            song._genresLower       = (song.genres_raw          || []).map(g => g.toLowerCase());
            song._tagsLower         = (song.tags_raw            || []).map(t => t.toLowerCase());
            song._genresCount       = (song.genres_raw          || []).length;
            song._tagsCount         = (song.tags_raw            || []).length;
            song._chantingLower     = (song.chanting            || "").toLowerCase();
            song._roomPlayersLower  = (song.room_players        || []).map(p => p.toLowerCase());
            song._livesTakenLower   = (song.lives_taken_flat    || []);
            song._livesSavedLower   = (song.lives_saved_flat    || []);
            song._diffParsed        = song.difficulty === "Unrated"   ? -Infinity                     : parseFloat(song.difficulty);
            song._guessersCount     = song.guessers_flat              ? song.guessers_flat    .length : 0;
            song._listersCount      = song.listers_flat               ? song.listers_flat     .length : 0;
            song._vintageParsed     = parseVintageToFloat(song.vintage);
            song._correctTeamsLower = (song.correct_teams_flat  || []).map(tid => {
                const leader = window.dashboardData.json_teams.find(t => t._tid === tid || t.tid === tid);
                return leader ? leader["Team Leader"].toLowerCase() : "";
            });

            return song;
        });

        initColumnSettingsCheckboxes    ();
        sortSearchData                  ();
        renderSearchTable               (globalSearchData);
        renderTierCharts                ();
        updateSongHelpDropdown          ();
        updateGuessHelpDropdown         ();
        updateSearchHelpDropdown        ();

        const searchInput = document.getElementById('songSearchInput');
        if (searchInput) {

            const processQuery = (e) => {
                const rawQuery = searchInput.value.trim();

                if (!rawQuery) {
                    renderSearchTable(globalSearchData);
                    return;
                }

                const tokens        = [];
                const tokenRegex    = /\(|\)|or\b|and\b|[a-zA-Z0-9_/-]+(?:<=|>=|!=|!:|[:<>==])"[^"]*"|[^\s"()]+|"[^"]*"/gi;

                let match;
                while ((match = tokenRegex.exec(rawQuery)) !== null) tokens.push(match[0]);

                function parseToRPN(tokens) {
                    const outputQueue   = [];
                    const operatorStack = [];
                    const precedence    = {'or': 1, 'and': 2};
                    let expectOperator  = false;

                    tokens.forEach(token => {
                        const lowerToken = token.toLowerCase();

                        if (expectOperator && lowerToken !== 'and' && lowerToken !== 'or' && lowerToken !== ')') {
                            while (operatorStack.length && precedence[operatorStack[operatorStack.length - 1]] >= precedence['and']) outputQueue.push(operatorStack.pop());
                            operatorStack.push('and');
                        }

                        if (lowerToken === 'and' || lowerToken === 'or') {
                            while (operatorStack.length && precedence[operatorStack[operatorStack.length - 1]] >= precedence[lowerToken]) outputQueue.push(operatorStack.pop());
                            operatorStack.push(lowerToken);
                            expectOperator = false;
                        }

                        else if (token === '(') {
                            operatorStack.push(token);
                            expectOperator = false;
                        }

                        else if (token === ')') {
                            while (operatorStack.length && operatorStack[operatorStack.length - 1] !== '(') outputQueue.push(operatorStack.pop());
                            operatorStack.pop();
                            expectOperator = true;
                        }

                        else {
                            outputQueue.push(token);
                            expectOperator = true;
                        }
                    });

                    while (operatorStack.length) outputQueue.push(operatorStack.pop());
                    return outputQueue;
                }

                function evaluateSingleToken(song, token) {
                    const queryRegex    = /^([a-zA-Z_]+)(<=|>=|!=|!:|[:<>==])(.+)$/;
                    const parsedMatch   = token.match(queryRegex);

                    if (parsedMatch) {
                        let queryKey = parsedMatch[1].toLowerCase();
                        if (queryKey === "correct") queryKey = "guessers";
                        if (queryKey === "list")    queryKey = "listers";

                        let cleanVal = parsedMatch[3].trim();
                        if (cleanVal.startsWith('"') && cleanVal.endsWith('"')) cleanVal = cleanVal.slice(1, -1).trim();
                        else                                                    cleanVal = cleanVal.replace(/^"|"$/g, '').trim();

                        return evaluateQuery(song, queryKey, parsedMatch[2], cleanVal.toLowerCase());
                    }

                    const wordClean = token.replace(/^"|"$/g, '').toLowerCase();

                    return (
                        song._romajiLower       .includes               (wordClean)     ||
                        song._englishLower      .includes               (wordClean)     ||
                        song._songLower         .includes               (wordClean)     ||
                        song._artistRawLower    .includes               (wordClean)     ||
                        song._composerLower     .includes               (wordClean)     ||
                        song._arrangerLower     .includes               (wordClean)     ||
                        song._typeLower         .includes               (wordClean)     ||
                        song._vintageLower      .includes               (wordClean)     ||
                        song._animeTypeLower    .includes               (wordClean)     ||
                        song._chantingLower     .includes               (wordClean)     ||
                        song._genresLower       .some(g => g.includes   (wordClean))    ||
                        song._tagsLower         .some(t => t.includes   (wordClean))    ||
                        song.difficulty         .toLowerCase().includes (wordClean)
                    );
                }

                const rpnTokens = parseToRPN(tokens);

                const filtered = globalSearchData.filter(song => {
                    if (rpnTokens.length === 0) return true;
                    const stack = [];

                    for (let token of rpnTokens) {
                        const lowerToken = typeof token === 'string' ? token.toLowerCase() : '';

                        if (lowerToken === 'and') {
                            const b = stack.pop();
                            const a = stack.pop();
                            stack.push(a && b);
                        }

                        else if (lowerToken === 'or') {
                            const b = stack.pop();
                            const a = stack.pop();
                            stack.push(a || b);
                        }

                        else stack.push(evaluateSingleToken(song, token));
                    }

                    return stack[0];
                });

                renderSearchTable(filtered);
            };

            searchInput.addEventListener('input',           debounce(processQuery, 250));
            searchInput.addEventListener('input-direct',    processQuery);
        }
    })

    .catch(err => console.error("Error setting up lookup engine layout context mapping:", err));

window.searchPlayerMetricFromTable = function(filterStr) {
    const searchTabBtn  = Array.from(document.querySelectorAll('.tab-btn')).find(btn => btn.getAttribute('onclick') && btn.getAttribute('onclick').includes('search-tab'));
    const searchInput   = document.getElementById('songSearchInput');

    if (searchTabBtn && searchInput) {
        searchTabBtn.click();
        searchInput.value = "";
        searchInput.value = filterStr;
        searchInput.dispatchEvent(new Event('input-direct'));
    }
};

window.searchTourFraction = function(fractionNum) {
    const searchTabBtn  = Array.from(document.querySelectorAll('.tab-btn')).find(btn => btn.getAttribute('onclick') && btn.getAttribute('onclick').includes('search-tab'));
    const searchInput   = document.getElementById('songSearchInput');

    if (searchTabBtn && searchInput) {
        searchTabBtn.click();
        searchInput.value = "";
        searchInput.value = `correct:${fractionNum}`;
        searchInput.dispatchEvent(new Event('input-direct'));
    }
};

window.searchTourMetadata = function(type, val) {
    const searchTabBtn  = Array.from(document.querySelectorAll('.tab-btn')).find(btn => btn.getAttribute('onclick') && btn.getAttribute('onclick').includes('search-tab'));
    const searchInput   = document.getElementById('songSearchInput');

    if (searchTabBtn && searchInput) {
        const cleanVal = val.replace(/\s*\(\d+\)\s*$/, '').trim();
        searchTabBtn.click();
        searchInput.value = "";
        searchInput.value = `${type}:"${cleanVal.toLowerCase()}"`;
        searchInput.dispatchEvent(new Event('input-direct'));
    }
};

window.sortPlayerColumnFromTour = function(colName, asc) {
    const playerTabBtn      = Array.from(document.querySelectorAll('.tab-btn')).find(btn => btn.getAttribute('onclick') && btn.getAttribute('onclick').includes('player-tab'));
    const playerSearchInput = document.getElementById('playerSearchInput');

    if (playerTabBtn) {
        playerTabBtn.click();
        playerSearchInput.value = "";
        globalPlayerSortState = {columnName: colName, ascending: asc};
        sortAndRenderPlayers();
    }
};

window.searchPlayerFilterFromTour = function(filterStr, colName, asc) {
    const playerTabBtn      = Array.from(document.querySelectorAll('.tab-btn')).find(btn => btn.getAttribute('onclick') && btn.getAttribute('onclick').includes('player-tab'));
    const playerSearchInput = document.getElementById('playerSearchInput');
    
    if (playerTabBtn && playerSearchInput) {
        playerTabBtn.click();
        playerSearchInput.value = "";
        playerSearchInput.value = filterStr;
        globalFilteredPlayers   = processPlayerFiltering(filterStr);
        globalPlayerSortState   = {columnName: colName, ascending: asc};
        sortAndRenderPlayers();
    }
};

window.searchSoloRigConverter = function(el) {
    const searchTabBtn  = Array.from(document.querySelectorAll('.tab-btn')).find(btn => btn.getAttribute('onclick') && btn.getAttribute('onclick').includes('search-tab'));
    const searchInput   = document.getElementById('songSearchInput');
    
    if (searchTabBtn && searchInput) {
        const displayVal    = el.textContent || el.innerText;
        const listOwner     = displayVal.split(' (')[0].trim().toLowerCase();

        searchTabBtn.click();
        searchInput.value = "";
        searchInput.value = `list:${listOwner} list:1`;
        searchInput.dispatchEvent(new Event('input-direct'));
    }
};

window.searchTeamSolos = function(teamLeader) {
    const searchTabBtn = Array.from(document.querySelectorAll('.tab-btn')).find(btn => btn.getAttribute('onclick') && btn.getAttribute('onclick').includes('search-tab'));
    const searchInput = document.getElementById('songSearchInput');
    
    if (searchTabBtn && searchInput) {
        searchTabBtn.click();
        searchInput.value = "";
        searchInput.value = `correctteam:${teamLeader.toLowerCase()} correct:1`;
        searchInput.dispatchEvent(new Event('input-direct'));
    }
};