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

if (use_teams)  tourTabBtn.innerText = "Tour/Team";
else            tourTabBtn.innerText = "Tour";

if (use_teams)  tabContainer.insertAdjacentHTML('beforeend', `<button class="tab-btn" onclick="switchDashboardTab(event, 'tier-tab')">Tier</button>`);
                tabContainer.insertAdjacentHTML('beforeend', `<button class="tab-btn" onclick="switchDashboardTab(event, 'song-tab')">Song</button>`);
if (watched)    tabContainer.insertAdjacentHTML('beforeend', `<button class="tab-btn" onclick="switchDashboardTab(event, 'guess-tab')">Guess/List</button>`);
else            tabContainer.insertAdjacentHTML('beforeend', `<button class="tab-btn" onclick="switchDashboardTab(event, 'guess-tab')">Guess</button>`);
                tabContainer.insertAdjacentHTML('beforeend', `<button class="tab-btn" onclick="switchDashboardTab(event, 'search-tab')">Search</button>`);

const thickBorderColumns = new Set([
    "Player",
    "Tier",
    "GR",
    "Score",
    "Mean Over-8",
    "Lives Saved",
    "IN GR",
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
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active-content'));
    document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active-tab'));
    document.getElementById(tabId).classList.add('active-content');
    evt.currentTarget.classList.add('active-tab');
    window.dispatchEvent(new Event('resize'));
}

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
    const ticks     = displaySongs.filter(s => s.startsWith('✓'));
    const crosses   = displaySongs.filter(s => s.startsWith('✗'));
    const valid     = ticks.length + crosses.length;

    let tickTarget = 5;

    if (valid > 0) {
        tickTarget  = Math.round((ticks.length / valid) * 10);
        if (ticks.length > 0 && crosses.length > 0) tickTarget = Math.max(1, Math.min(9, tickTarget));
    }

    let crossTarget = 10 - tickTarget;

    if (ticks.length < tickTarget) {
        tickTarget  = ticks.length;
        crossTarget = Math.min(crosses.length, 10 - tickTarget);
    }

    else if (crosses.length < crossTarget) {
        crossTarget = crosses.length;
        tickTarget  = Math.min(ticks.length, 10 - crossTarget);
    }

    const sampledTicks      = ticks     .sort(() => Math.random() - 0.5).slice(0, tickTarget);
    const sampledCrosses    = crosses   .sort(() => Math.random() - 0.5).slice(0, crossTarget);

    return [...formatAndSortSongsList(sampledTicks), ...formatAndSortSongsList(sampledCrosses)];
};

let globalPlayerSortState   = {columnName: "GR", ascending: false};
let globalFilteredPlayers   = [];
let globalMetricHighlights  = {};

const playerHeadersMasterConfig = [
    {id: "player",              name: "Player",                 ascMetric: false,   teamReq: false, watchedReq: false,  def: true},
    {id: "team",                name: "Team",                   ascMetric: false,   teamReq: true,  watchedReq: false,  def: false},
    {id: "tier",                name: "Tier",                   ascMetric: true,    teamReq: true,  watchedReq: false,  def: false},
    {id: "elo",                 name: "Elo",                    ascMetric: false,   teamReq: true,  watchedReq: false,  def: true},
    {id: "guessrate",           name: "GR",                     ascMetric: false,   teamReq: false, watchedReq: false,  def: true},
    {id: "uf",                  name: "UF",                     ascMetric: false,   teamReq: true,  watchedReq: false,  def: true},
    {id: "score",               name: "Score",                  ascMetric: false,   teamReq: true,  watchedReq: false,  def: false},
    {id: "18s",                 name: "1/8s",                   ascMetric: false,   teamReq: false, watchedReq: false,  def: true},
    {id: "28s",                 name: "2/8s",                   ascMetric: false,   teamReq: false, watchedReq: false,  def: true},
    {id: "78s",                 name: "7/8s",                   ascMetric: true,    teamReq: false, watchedReq: false,  def: true},
    {id: "meanover8",           name: "Mean Over-8",            ascMetric: true,    teamReq: false, watchedReq: false,  def: false},
    {id: "livestaken",          name: "Lives Taken",            ascMetric: false,   teamReq: true,  watchedReq: false,  def: false},
    {id: "livessaved",          name: "Lives Saved",            ascMetric: false,   teamReq: true,  watchedReq: false,  def: false},
    {id: "opguessrate",         name: "OP GR",                  ascMetric: false,   teamReq: false, watchedReq: false,  def: true},
    {id: "edguessrate",         name: "ED GR",                  ascMetric: false,   teamReq: false, watchedReq: false,  def: true},
    {id: "inguessrate",         name: "IN GR",                  ascMetric: false,   teamReq: false, watchedReq: false,  def: true},
    {id: "rigs",                name: "Rigs",                   ascMetric: false,   teamReq: false, watchedReq: true,   def: true},
    {id: "rigrate",             name: "Rig Rate",               ascMetric: false,   teamReq: false, watchedReq: true,   def: false},
    {id: "solorigs",            name: "Solo Rigs",              ascMetric: false,   teamReq: false, watchedReq: true,   def: false},
    {id: "solorigrate",         name: "Solo Rig Rate",          ascMetric: false,   teamReq: false, watchedReq: true,   def: false},
    {id: "rigover8",            name: "Rig Over-8",             ascMetric: true,    teamReq: false, watchedReq: true,   def: false},
    {id: "over8delta",          name: "Over-8 Δ",               ascMetric: false,   teamReq: false, watchedReq: true,   def: false},
    {id: "rigguessrate",        name: "Rig GR",                 ascMetric: false,   teamReq: false, watchedReq: true,   def: true},
    {id: "offguessrate",        name: "Off GR",                 ascMetric: false,   teamReq: false, watchedReq: true,   def: true},
    {id: "rigdelta",            name: "Rig Δ",                  ascMetric: false,   teamReq: false, watchedReq: true,   def: false},
    {id: "meandifficultyhit",   name: "Mean Difficulty Hit",    ascMetric: true,    teamReq: false, watchedReq: false,  def: false},
    {id: "medianvintagehit",    name: "Median Vintage Hit",     ascMetric: false,   teamReq: false, watchedReq: false,  def: false},
    {id: "mediantime",          name: "Median Time",            ascMetric: true,    teamReq: false, watchedReq: false,  def: false},
    {id: "chantguessrate",      name: "Chant GR",               ascMetric: false,   teamReq: false, watchedReq: false,  def: false}
];

let activePlayerHeadersConfig = playerHeadersMasterConfig.filter(col => {
    if (col.teamReq     && !use_teams)  return false;
    if (col.watchedReq  && !watched)    return false;

    return true;
});

activePlayerHeadersConfig.forEach(col => {col.visible = col.def;});

window.togglePlayerColumnSettingsMenu = function(event) {
    event.stopPropagation();
    document.getElementById("playerColumnSettingsDropdown").classList.toggle("hidden");
};

document.addEventListener("click", () => {
    const pMenu = document.getElementById("playerColumnSettingsDropdown");
    if (pMenu) pMenu.classList.add("hidden");
});

if (document.getElementById("playerColumnSettingsDropdown")) document.getElementById("playerColumnSettingsDropdown").addEventListener("click", (e) => e.stopPropagation());

function initPlayerColumnSettings() {
    const container = document.getElementById("playerColumnCheckboxContainer");
    const masterChk = document.getElementById("playerAllColumnsMasterCheckbox");

    if (!container || !masterChk) return;
    container.innerHTML = "";

    const updateMasterState = () => {
        const allChecked        = activePlayerHeadersConfig.every(c => c.visible);
        const noneChecked       = activePlayerHeadersConfig.every(c => !c.visible);
        masterChk.checked       = allChecked;
        masterChk.className     = "rounded accent-black";
        masterChk.indeterminate = !allChecked && !noneChecked;
    };

    masterChk.addEventListener("change", () => {
        activePlayerHeadersConfig.forEach(c => { c.visible = masterChk.checked; });
        document.querySelectorAll(".player-col-toggle-checkbox").forEach(chk => chk.checked = masterChk.checked);
        sortAndRenderPlayers();
    });

    activePlayerHeadersConfig.forEach(col => {
        const label     = document.createElement("label");
        label.className = "flex items-center gap-2 cursor-pointer w-full text-left font-normal text-black";
        const chk       = document.createElement("input");
        chk.type        = "checkbox";
        chk.className   = "player-col-toggle-checkbox rounded accent-black";
        chk.checked     = col.visible;

        chk.addEventListener("change", () => {
            col.visible = chk.checked;

            updateMasterState       ();
            sortAndRenderPlayers    ();
        });

        label.appendChild       (chk);
        label.appendChild       (document.createTextNode(col.name));
        container.appendChild   (label);
    });

    updateMasterState();
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
                if      (h.name === "GR")               clickHandler = ` onclick="searchPlayerMetricFromTable('seen:${currentPlayerName}')"`;
                else if (h.name === "1/8s")             clickHandler = ` onclick="searchPlayerMetricFromTable('correct:${currentPlayerName} correct:1')"`;
                else if (h.name === "2/8s")             clickHandler = ` onclick="searchPlayerMetricFromTable('correct:${currentPlayerName} correct:2')"`;
                else if (h.name === "7/8s")             clickHandler = ` onclick="searchPlayerMetricFromTable('correct!:${currentPlayerName} correct:7')"`;
                else if (h.name === "Lives Taken")      clickHandler = ` onclick="searchPlayerMetricFromTable('lifetaken:${currentPlayerName}')"`;
                else if (h.name === "Lives Saved")      clickHandler = ` onclick="searchPlayerMetricFromTable('lifesaved:${currentPlayerName}')"`;
                else if (h.name === "OP GR")            clickHandler = ` onclick="searchPlayerMetricFromTable('seen:${currentPlayerName} songtype:op')"`;
                else if (h.name === "ED GR")            clickHandler = ` onclick="searchPlayerMetricFromTable('seen:${currentPlayerName} songtype:ed')"`;
                else if (h.name === "IN GR")            clickHandler = ` onclick="searchPlayerMetricFromTable('seen:${currentPlayerName} songtype:in')"`;
                else if (h.name === "Solo Rigs")        clickHandler = ` onclick="searchPlayerMetricFromTable('list:${currentPlayerName} list:1')"`;
                else if (h.name === "Rig GR")           clickHandler = ` onclick="searchPlayerMetricFromTable('list:${currentPlayerName}')"`;
                else if (h.name === "Off GR")           clickHandler = ` onclick="searchPlayerMetricFromTable('list!:${currentPlayerName}')"`;
                else if (h.name === "Chant GR")         clickHandler = ` onclick="searchPlayerMetricFromTable('seen:${currentPlayerName} chanting:yes')"`;
            }

            if ((h.name === "Player" && rawCell && rawCell.details && rawCell.details.length > 0) || 
                (rawCell !== null && typeof rawCell === 'object' && rawCell.details && rawCell.details.length > 0)) {
                let encodedDetails = encodeURIComponent(JSON.stringify(rawCell.details));
                tbody += `<td class="${cellStyle.trim()}" data-songs="${encodedDetails}"${clickHandler}>${finalVal}</td>`;
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
        const rawQuery = playerSearchInput ? playerSearchInput.value.trim() : "";
        globalFilteredPlayers = processPlayerFiltering(rawQuery);
        sortAndRenderPlayers();
    };

    if (playerSearchInput) playerSearchInput.addEventListener('input', debounce(triggerPlayerQueryProcess, 250));
    globalFilteredPlayers = [...players];
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

            let finalVal = (h === "Team Leader") ? `<b>${displayVal}</b>` : displayVal;

            if (rawCell !== null && typeof rawCell === 'object' && rawCell.details && rawCell.details.length > 0) {
                let encodedDetails  = encodeURIComponent(JSON.stringify(rawCell.details));
                let clickHandler    = (h === "Total 1/8s") ? ` onclick="searchTeamSolos('${row["Team Leader"]}')"` : "";
                tbody += `<td class="${cellStyle.trim()}" data-songs="${encodedDetails}"${clickHandler}>${finalVal}</td>`;
            }

            else tbody += `<td class="${cellStyle.trim()}">${finalVal}</td>`;
        });

        tbody += "</tr>";
    });

    table.innerHTML = thead + tbody + "</tbody>";
}

window.toggleTierChartMode = function() {
    const btn               = document.getElementById("tierModeToggleBtn");
    currentTierChartMode    = currentTierChartMode === "TIER" ? "ALL" : "TIER";
    btn.innerText           = currentTierChartMode;
    renderTierCharts();
};

window.toggleGlobalChartMode = function() {
    globalChartMode = globalChartMode === "RATE" ? "COUNT" : "RATE";

    if (globalChartMode === "COUNT" && (c1Sub === "HIT" || c1Sub === "OFF")) {
        c1Sub           = "BASE";
        const subBtn    = document.getElementById("c1_sub_btn");

        if (subBtn) subBtn.innerText = "BASE";
    }

    document.getElementById("global_mode_btn").innerText = globalChartMode;
    renderTierCharts();
};

window.toggleC1Sub = function() {
    const subs                                      = globalChartMode === "COUNT" ? ["BASE", "OVER-8", "RIG", "CHANT"] : ["BASE", "OVER-8", "RIG", "HIT", "OFF", "CHANT"];
    let idx                                         = subs.indexOf(c1Sub);
    if (idx === -1) idx                             = 0; 
    c1Sub                                           = subs[(idx + 1) % subs.length];
    document.getElementById("c1_sub_btn").innerText = c1Sub;
    renderTierCharts();
};

window.toggleC2Sub = function() {};

window.toggleC3Mode = function() {
    const modes                                         = ["MIN", "MEAN", "MED", "MAX", "STDEV"];
    let idx                                             = modes.indexOf(c3Mode);
    c3Mode                                              = modes[(idx + 1) % modes.length];
    document.getElementById("c3_mode_btn").innerText    = c3Mode;
    renderTierCharts();
};

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

    const formatSampleTextList = (list, limit = 10) => {
        let totalCount          = list.length;
        let randomizedSample    = [...list].sort(() => Math.random() - 0.5).slice(0, limit);

        randomizedSample.sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
        if (totalCount > limit) randomizedSample.push(`and ${totalCount - limit} more`);
        return randomizedSample.join('<br>');
    };

    const formatFractionalSample = (fractionStr, songsList, limit = 10) => {
        const ticks         = songsList.filter(s => s.startsWith('✓'));
        const crosses       = songsList.filter(s => s.startsWith('✗'));
        const totalValid    = ticks.length + crosses.length;

        let tickTarget = 5;

        if (totalValid > 0) {
            tickTarget = Math.round((ticks.length / totalValid) * limit);
            if (ticks.length > 0 && crosses.length > 0) tickTarget = Math.max(1, Math.min(limit - 1, tickTarget));
        }

        let crossTarget = limit - tickTarget;

        if (ticks.length < tickTarget) {
            tickTarget = ticks.length;
            crossTarget = Math.min(crosses.length, limit - tickTarget);
        }

        else if (crosses.length < crossTarget) {
            crossTarget = crosses.length;
            tickTarget = Math.min(ticks.length, limit - crossTarget);
        }

        const sampledTicks      = ticks     .sort(() => Math.random() - 0.5).slice(0, tickTarget);
        const sampledCrosses    = crosses   .sort(() => Math.random() - 0.5).slice(0, crossTarget);

        sampledTicks    .sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
        sampledCrosses  .sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));

        const outputSample = [...sampledTicks, ...sampledCrosses];

        if (totalValid > limit) outputSample.push(`and ${totalValid - limit} more`);
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
                marker              : {color: c1BaseColors[i], line: {color: 'black', width: 1}},
                hovertext           : (i === 0 ? trace0Hovers : trace1Hovers).slice().reverse(),
                hoverinfo           : 'text',
                text                : c1Data.multiData[i]   .slice().reverse().map(v => v ? v.toFixed(0) : ""),
                textposition        : 'inside',
                insidetextanchor    : 'middle',
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
                marker              : {color: c8Colors[i], line: {color: 'black', width: 1}},
                hovertext           : x8TraceHovers[i]      .slice().reverse(),
                hoverinfo           : 'text',
                text                : c1Data.multiData[i]   .slice().reverse().map(v => v ? v.toFixed(globalChartMode === "RATE" ? 1 : 0) : ""),
                textposition        : 'inside', insidetextanchor: 'middle',
                textfont            : {family: 'Segoe UI', size: 15, color: 'white', weight: 'bold'}
            });
        }
    }

    else {
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
            textfont            : {family: 'Segoe UI', size: 15, color: 'black', weight: 'bold'},
            marker              : {color: 'white', line: {color: 'black', width: 2}}
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

    layoutC1.height = 35 * c1Data.yLabels.length;
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
            if (!data.points || data.points.length === 0) return;

            const pt            = data.points[0];
            const pNameClean    = String(pt.y).trim().toLowerCase();

            if (!pNameClean) return;
            let query = "";

            if (c1Sub === "BASE") {
                if (globalChartMode === "COUNT")    query = pt.curveNumber === 0 ? `list:${pNameClean} correct:${pNameClean}` : `list!:${pNameClean} correct:${pNameClean}`;
                else                                query = `seen:${pNameClean}`;
            }

            else if (c1Sub === "OVER-8") {
                let matchX8 = pt.curveNumber + 1;
                query = `correct:${pNameClean} correct:${matchX8}`;
            }

            else if (c1Sub === "RIG") query = `list:${pNameClean}`;
            else if (c1Sub === "HIT") query = `list:${pNameClean}`;
            else if (c1Sub === "OFF") query = `list!:${pNameClean}`;

            else if (c1Sub === "CHANT") {
                if (globalChartMode === "COUNT") query = `correct:${pNameClean} chanting:yes`;
                else query = `seen:${pNameClean} chanting:yes`;
            }

            if (query) window.searchPlayerMetricFromTable(query);
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
                marker              : {color: c2Colors[i], line: {color: 'black', width: 1}},
                hovertext           : (i === 0 ? tkHovers : svHovers)   .slice().reverse(),
                hoverinfo           : 'text',
                text                : c2Data.multiData[i]               .slice().reverse().map(v => v ? v.toFixed(0) : ""),
                textposition        : 'inside',
                insidetextanchor    : 'middle',
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
                marker              : {color: c3Colors[i], line: {color: 'black', width: 1}},
                hovertext           : (i === 0 ? tkHovers : (i === 1 ? othHovers : svHovers))   .slice().reverse(),
                hoverinfo           : 'text',
                text                : c2Data.multiData[i]                                       .slice().reverse().map(v => v ? v.toFixed(1) : ""),
                textposition        : 'inside',
                insidetextanchor    : 'middle',
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

    layoutC2.height = 35 * c2Data.yLabels.length;
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
            if (!data.points || data.points.length === 0) return;

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

    let c3Data = buildChartData(currentTierChartMode, c3Sort, (p) => getC3ValueAndHover(p, c3Mode), null);

    let c3Traces = [{
        x                   : c3Data.singleXVals    .slice().reverse(),
        y                   : c3Data.yLabels        .slice().reverse(),
        type                : 'bar',
        orientation         : 'h',
        hovertext           : c3Data.customHovers   .slice().reverse(),
        hoverinfo           : 'text',
        text                : c3Data.singleXVals    .slice().reverse().map(v => v === null ? "" : v.toFixed(2) + " "),
        textposition        : 'inside', 
        insidetextanchor    : 'end',
        textfont            : {family: 'Segoe UI', size: 15, color: 'black', weight: 'bold'},
        marker              : {color: 'white', line: {color: 'black', width: 2}}
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

    layoutC3.height = 35 * c3Data.yLabels.length;
    Plotly.newPlot('tierChart_TimeMetrics', c3Traces, layoutC3, {responsive: true, displayModeBar: false});
}

function setupTooltipListeners() {
    const tooltipNode = document.getElementById('customJsTooltip');

    function positionTooltip(e) {
        tooltipNode.style.display = 'block';
        const tooltipWidth = tooltipNode.offsetWidth; const tooltipHeight = tooltipNode.offsetHeight;

        let xPos = e.pageX + 15;
        let yPos = e.pageY + 15;

        if (e.clientX + 15 + tooltipWidth   > window.innerWidth)    xPos = e.pageX - tooltipWidth   - 15;
        if (e.clientY + 15 + tooltipHeight  > window.innerHeight)   yPos = e.pageY - tooltipHeight  - 15;

        if (xPos < window.scrollX) xPos = window.scrollX + 5;
        if (yPos < window.scrollY) yPos = window.scrollY + 5;

        tooltipNode.style.left = xPos + 'px';
        tooltipNode.style.top = yPos + 'px';
    }

    document.querySelectorAll('table th[data-metric], table td[data-metric]').forEach(th => {
        const metricKey = th.getAttribute('data-metric');
        if (!colExplanations[metricKey]) return;

        th.removeEventListener('mouseenter',    th._handlerEnter);
        th.removeEventListener('mousemove',     positionTooltip);
        th.removeEventListener('mouseleave',    th._handlerLeave);

        th._handlerEnter = (e)  => {tooltipNode.innerHTML       = colExplanations[metricKey]; positionTooltip(e);};
        th._handlerLeave = ()   => {tooltipNode.style.display   = 'none'; };

        th.addEventListener('mouseenter', th._handlerEnter);
        th.addEventListener('mousemove',  positionTooltip);
        th.addEventListener('mouseleave', th._handlerLeave);
    });

    document.querySelectorAll('td[data-songs]').forEach(td => {
        td.addEventListener('mouseenter', (e) => {
            try {
                const songs = JSON.parse(decodeURIComponent(td.getAttribute('data-songs')));
                if (!songs || songs.length === 0) return;

                if      (td.classList.contains('highlight-best'))   {tooltipNode.style.backgroundColor = c2;        tooltipNode.style.color = 'white';}
                else if (td.classList.contains('highlight-worst'))  {tooltipNode.style.backgroundColor = c0;        tooltipNode.style.color = 'white';}
                else                                                {tooltipNode.style.backgroundColor = 'black';   tooltipNode.style.color = 'white';}

                let displaySongs        = [...songs];
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
                const containsRegex = fractionRegex.test(songs[0]);
                let fractionHeader  = "";

                if (containsRegex) {
                    fractionHeader = `<b>${songs[0]}</b>`;
                    displaySongs.shift();
                }

                if (containsRegex && songs.length > 10) displaySongs = sampleLargeSongList(displaySongs).map(s => (s.startsWith('✓') || s.startsWith('✗') || !isPlayerSubHover) ? s : `• ${s}`);

                else {
                    if (displaySongs.length > 10) displaySongs = displaySongs.sort(() => Math.random() - 0.5).slice(0, 10);
                    displaySongs = formatAndSortSongsList(displaySongs, !isPlayerSubHover);
                }

                const totalSongsCount = containsRegex ? (songs.length - 1) : songs.length;
                if (totalSongsCount > 10) displaySongs.push(`and ${totalSongsCount - 10} more`);
                tooltipNode.innerHTML = containsRegex ? `${fractionHeader}<br>${displaySongs.join('<br>')}` : displaySongs.join('<br>');
                positionTooltip(e);
            }

            catch (err) {}
        });

        td.addEventListener('mousemove',    positionTooltip);
        td.addEventListener('mouseleave',   () => {
            tooltipNode.style.display           = 'none';
            tooltipNode.style.backgroundColor   = 'black';
            tooltipNode.style.color             = 'white';
        });
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

            let bin_songs       = matrixSongs[key] ? [...matrixSongs[key]] : [];
            let song_hover_str  = "";

            if (bin_songs.length > 10) {
                const remainingCount = bin_songs.length - 10;
                bin_songs = bin_songs.sort(() => Math.random() - 0.5).slice(0, 10);
                bin_songs = formatAndSortSongsList(bin_songs, false);
                song_hover_str = "<br>• " + bin_songs.join("<br>• ") + "<br>and " + remainingCount + " more";
            }

            else if (bin_songs.length > 0) {
                bin_songs = formatAndSortSongsList(bin_songs, false);
                song_hover_str = "<br>• " + bin_songs.join("<br>• ");
            }

            rowText.push(`<b>${diffStr}<br>${vintageStr}<br>Over-8: ${val.toFixed(2)}</b>${song_hover_str}`);

            annotations.push({
                x               : j,
                y               : i,
                text            : `<b>${matrixBins[key].count}</b>`,
                font            : {family: 'Segoe UI', size: (numX > 8 ? 65 : 70), color: 'white'},
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
    hovertemplate   : '<span style="text-align: left; display: block;">%{text}</span><extra></extra>',
    hoverlabel      : {align: 'left', bgcolor: bgColors, font: {family: 'Segoe UI', size: 15}},
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
    margin      : {l: 75, r: 0, t: 25, b: 75}
}, {responsive: true, displayModeBar: false});

const songChartDiv = document.getElementById('plotlySongChart');

if (songChartDiv) {
    songChartDiv.on('plotly_click', function(data) {
        if (!data.points || data.points.length === 0) return;

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
        const bubbleSize = Math.max(10, d[sizeKeyMultiplier] * 2);
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
            size        : scatterData.map(d => Math.max(10, d.gr * 2)),
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
            line        : {color: 'black', width: 1},
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

let currentListChartMode = "ALL"; 

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

function renderListChart() {
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
            size        : activeScatterSource.map(d => Math.max(10, d.size * 2)),
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
            line        : {color: 'black', width: 1},
            cmin        : 0,
            cmax        : 100
        }
    });

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

if (watched) {
    const glBtn = document.getElementById("guessListToggleBtn");
    if (glBtn) glBtn.classList.remove("hidden");
}

let currentGuessListViewMode = "GUESS";

window.toggleGuessListViewMode = function() {
    const btn           = document.getElementById("guessListToggleBtn");
    const guessCtx      = document.getElementById("guessViewSubContext");
    const listCtx       = document.getElementById("listViewSubContext");
    const guessChart    = document.getElementById("guessChartContainer");
    const listChart     = document.getElementById("listChartContainer");
    const listSubToggle = document.getElementById("listModeToggleContainer");

    if (currentGuessListViewMode === "GUESS") {
        currentGuessListViewMode    = "LIST";
        btn.innerText               = "LIST";

        if (guessCtx)       guessCtx        .classList.add      ("hidden");
        if (guessChart)     guessChart      .classList.add      ("hidden");
        if (listCtx)        listCtx         .classList.remove   ("hidden");
        if (listChart)      listChart       .classList.remove   ("hidden");
        if (listSubToggle)  listSubToggle   .classList.remove   ("hidden");

        renderListChart();
    }

    else {
        currentGuessListViewMode    = "GUESS";
        btn.innerText               = "GUESS";

        if (guessCtx)       guessCtx        .classList.remove   ("hidden");
        if (guessChart)     guessChart      .classList.remove   ("hidden");
        if (listCtx)        listCtx         .classList.add      ("hidden");
        if (listChart)      listChart       .classList.add      ("hidden");
        if (listSubToggle)  listSubToggle   .classList.add      ("hidden");

        window.dispatchEvent(new Event('resize'));
    }
};

window.toggleListChartMode = function() {
    const btn               = document.getElementById("listModeToggleBtn");
    currentListChartMode    = currentListChartMode === "ALL" ? "HIT" : "ALL";
    btn.innerText           = currentListChartMode;
    const xAxisDesc         = document.getElementById("listXAxisDescription");
    const yAxisDesc         = document.getElementById("listYAxisDescription");

    if (currentListChartMode === "HIT") {
        if (xAxisDesc) xAxisDesc.innerHTML = "<b>X-Axis (Over-8):</b> Mean of correct guessers across songs that this player guessed correctly from their own list";
        if (yAxisDesc) yAxisDesc.innerHTML = "<b>Y-Axis (Vintage):</b> Median vintage across songs that this player guessed correctly from their own list";
    }

    else {
        if (xAxisDesc) xAxisDesc.innerHTML = "<b>X-Axis (Over-8):</b> Mean of correct guessers across songs from this player's list";
        if (yAxisDesc) yAxisDesc.innerHTML = "<b>Y-Axis (Vintage):</b> Median vintage across songs from this player's list";
    }

    renderListChart();
};

let globalSortState     = {columnName: "Anime", ascending: true};
let currentSearchLang   = "JP"; 

const searchHeadersConfig = [
    {id: "anime",       name: "Anime",      visible: true},
    {id: "type",        name: "Song Type",  visible: true},
    {id: "chanting",    name: "Chanting",   visible: false},
    {id: "anime_type",  name: "Anime Type", visible: false},
    {id: "vintage",     name: "Vintage",    visible: false},
    {id: "genre",       name: "Genre",      visible: false},
    {id: "tag",         name: "Tag",        visible: false},
    {id: "difficulty",  name: "Difficulty", visible: false},
    {id: "song",        name: "Song",       visible: true},
    {id: "artist",      name: "Artist",     visible: true},
    {id: "composer",    name: "Composer",   visible: false},
    {id: "arranger",    name: "Arranger",   visible: false},
    {id: "guessers",    name: "Correct",    visible: true},
    {id: "listers",     name: "List",       visible: false}
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

window.toggleSearchLanguage = function() {
    const btn           = document.getElementById("langToggleBtn");
    currentSearchLang   = currentSearchLang === "JP" ? "EN" : "JP";
    btn.innerText       = currentSearchLang;

    sortSearchData      ();
    triggerTableRefresh ();
};

window.togglePlayerGuideMenu = function(event) {
    event.stopPropagation();
    document.getElementById("playerGuideDropdown").classList.toggle("hidden");
};

window.toggleSongGuideMenu = function(event) {
    event.stopPropagation();
    document.getElementById("songGuideDropdown").classList.toggle("hidden");
};

document.addEventListener("click", () => {
    const pMenu     = document.getElementById("playerColumnSettingsDropdown");
    const sMenu     = document.getElementById("columnSettingsDropdown");
    const pGuide    = document.getElementById("playerGuideDropdown");
    const sGuide    = document.getElementById("songGuideDropdown");

    if (pMenu)  pMenu   .classList.add("hidden");
    if (sMenu)  sMenu   .classList.add("hidden");
    if (pGuide) pGuide  .classList.add("hidden");
    if (sGuide) sGuide  .classList.add("hidden");
});

if (document.getElementById("playerGuideDropdown")) document.getElementById("playerGuideDropdown")  .addEventListener("click", (e) => e.stopPropagation());
if (document.getElementById("songGuideDropdown"))   document.getElementById("songGuideDropdown")    .addEventListener("click", (e) => e.stopPropagation());

window.toggleColumnSettingsMenu = function(event) {
    event.stopPropagation();
    const menu = document.getElementById("columnSettingsDropdown");
    menu.classList.toggle("hidden");
};

document.addEventListener("click", () => {
    const menu = document.getElementById("columnSettingsDropdown");
    if (menu) menu.classList.add("hidden");
});

if (document.getElementById("columnSettingsDropdown")) document.getElementById("columnSettingsDropdown").addEventListener("click", (e) => {e.stopPropagation();});

function initColumnSettingsCheckboxes() {
    const container = document.getElementById("columnCheckboxContainer");
    const masterChk = document.getElementById("allColumnsMasterCheckbox");

    if (!container || !masterChk) return;
    container.innerHTML = "";

    function updateMasterCheckboxState() {
        const allChecked        = searchHeadersConfig.every(c => c.visible);
        const noneChecked       = searchHeadersConfig.every(c => !c.visible);
        masterChk.checked       = allChecked;
        masterChk.indeterminate = !allChecked && !noneChecked;
    }

    masterChk.addEventListener("change", () => {
        searchHeadersConfig.forEach(c => { c.visible = masterChk.checked; });
        document.querySelectorAll(".col-toggle-checkbox").forEach(chk => {chk.checked = masterChk.checked;});
        triggerTableRefresh();
    });

    searchHeadersConfig.forEach(col => {
        const label     = document.createElement("label");
        label.className = "flex items-center gap-2 cursor-pointer w-full text-left";
        
        const chk       = document.createElement("input");
        chk.type        = "checkbox";
        chk.className   = "col-toggle-checkbox rounded accent-black";
        chk.checked     = col.visible;
        
        chk.addEventListener("change", () => {
            col.visible = chk.checked;
            updateMasterCheckboxState();
            triggerTableRefresh();
        });

        label       .appendChild(chk);
        label       .appendChild(document.createTextNode(col.name));
        container   .appendChild(label);
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
    const table         = document.getElementById('searchSongsTable');
    const counterNode   = document.getElementById('searchCounter');

    if (!table)         return;
    if (counterNode)    counterNode.innerText = `${filteredSongs.length}/${globalSearchData.length}`;

    const activeCols = searchHeadersConfig.filter(c => c.visible);

    if (activeCols.length === 0) {
        table.innerHTML = `<thead><tr><th>Error</th></tr></thead><tbody><tr><td class="p-2 text-center text-black">Select at least 1 column</td></tr></tbody>`;
        return;
    }

    if (filteredSongs.length === 0) {
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

    filteredSongs.forEach(song => {
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