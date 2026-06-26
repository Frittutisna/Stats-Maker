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
    "Guess Rate",
    "Score",
    "Mean Over-8",
    "Lives Saved",
    "IN Guess Rate",
    "Rig Rate",
    "Solo Rig Rate",
    "Over-8 Delta",
    "Rig Delta",
    "Metric",
    "Value",
    "Team Leader",
    "Tier",
    "Lives Saved",
    "Chanting Guess Rate"
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
    else if (differenceInMinutes    < 60)   displayString += `${differenceInMinutes} minute${differenceInMinutes === 1 ? '' : 's'} ago`;
    else if (differenceInHours      < 24)   displayString += `${differenceInHours} hour${differenceInHours === 1 ? '' : 's'} ago`;
    else if (differenceInDays       < 7)    displayString += `${differenceInDays} day${differenceInDays === 1 ? '' : 's'} ago`;
    else if (differenceInWeeks      < 24)   displayString += `${differenceInWeeks} week${differenceInWeeks === 1 ? '' : 's'} ago`;
    else if (differenceInMonths     < 24)   displayString += `${differenceInMonths} month${differenceInMonths === 1 ? '' : 's'} ago`;
    else                                    displayString += `${differenceInYears} year${differenceInDays === 1 ? '' : 's'} ago`;

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

function renderPlayerTable() {
    const table = document.getElementById('playerStandingsTable');
    if(!players || !players.length) return;
    let headers = Object.keys(players[0]);

    let thead = "<thead><tr>" + headers.map(h => {
        let classes = [];

        if (thickBorderColumns.has(h))  classes.push("border-col-group");
        if (colExplanations[h])         classes.push("has-explanation");

        let classStr = classes.length > 0 ? ` class="${classes.join(' ')}"` : '';
        return `<th${classStr} data-metric="${h}">${h.replace(/ /g, '<br>')}</th>`;
    }).join('') + "</tr></thead>";

    let tbody = "<tbody>";

    players.forEach((row, idx) => {
        let groupLine = groupBorders.includes(idx) ? " border-group-line" : "";
        tbody += `<tr class="${groupLine}">`;

        headers.forEach(h => {
            let rawCell     = row[h];
            let displayVal  = (rawCell !== null && typeof rawCell === 'object') ? rawCell.count : rawCell;
            let cellStyle   = thickBorderColumns.has(h) ? "border-col-group " : "";

            if (hlRules[h]) {
                let isBest  = (hlRules[h].best_idx  === idx);
                let isWorst = (hlRules[h].worst_idx === idx);

                if      (isBest)    cellStyle += "highlight-best ";
                else if (isWorst)   cellStyle += "highlight-worst ";
            }

            let intCols = ["1/8s", "2/8s", "7/8s", "Lives Taken", "Lives Saved", "Rigs", "Solo Rigs"];

            let formattedVal    = (typeof displayVal === 'number' && !intCols.includes(h))  ? displayVal.toFixed(2)     : displayVal;
            let finalVal        = (h === "Player")                                          ? `<b>${formattedVal}</b>`  : formattedVal;

            if ((h === "Player" && rawCell && rawCell.details && rawCell.details.length > 0) || (rawCell !== null && typeof rawCell === 'object' && rawCell.details && rawCell.details.length > 0)) {
                let encodedDetails = encodeURIComponent(JSON.stringify(rawCell.details));
                tbody += `<td class="${cellStyle.trim()}" data-songs="${encodedDetails}">${finalVal}</td>`;
            }

            else tbody += `<td class="${cellStyle.trim()}">${finalVal}</td>`;
        });

        tbody += "</tr>";
    });

    table.innerHTML = thead + tbody + "</tbody>";
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

    for (let i = 0; i < half; i++) {
        tbody += "<tr>";
        const leftRow = leftSlice[i];

        if (leftRow) {
            let rawCell     = leftRow.Value;
            let displayVal  = (rawCell !== null && typeof rawCell === 'object') ? rawCell.count : rawCell;
            let hasExp      = !!colExplanations[leftRow.Metric];
            let metricClass = hasExp ? "border-col-group has-explanation" : "border-col-group";
            let metricAttr  = `class='${metricClass}' data-metric="${leftRow.Metric}"`;

            if (rawCell !== null && typeof rawCell === 'object' && rawCell.details && rawCell.details.length > 0) {
                let encodedDetails = encodeURIComponent(JSON.stringify(rawCell.details));
                tbody += `<td ${metricAttr}><b>${leftRow.Metric}</b></td><td class="border-col-group" data-songs="${encodedDetails}">${displayVal}</td>`;
            }

            else tbody += `<td ${metricAttr}><b>${leftRow.Metric}</b></td><td class="border-col-group">${displayVal}</td>`;

        }

        else tbody += `<td class="border-col-group"></td><td class="border-col-group"></td>`;

        const rightRow = rightSlice[i];

        if (rightRow) {
            let rawCell     = rightRow.Value;
            let displayVal  = (rawCell !== null && typeof rawCell === 'object') ? rawCell.count : rawCell;
            let hasExp      = !!colExplanations[rightRow.Metric];
            let metricClass = hasExp ? "border-col-group has-explanation" : "border-col-group";
            let metricAttr  = `class='${metricClass}' data-metric="${rightRow.Metric}"`;

            if (rawCell !== null && typeof rawCell === 'object' && rawCell.details && rawCell.details.length > 0) {
                let encodedDetails = encodeURIComponent(JSON.stringify(rawCell.details));
                tbody += `<td ${metricAttr}><b>${rightRow.Metric}</b></td><td data-songs="${encodedDetails}">${displayVal}</td>`;
            }

            else tbody += `<td ${metricAttr}><b>${rightRow.Metric}</b></td><td>${displayVal}</td>`;
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
                let encodedDetails = encodeURIComponent(JSON.stringify(rawCell.details));
                tbody += `<td class="${cellStyle.trim()}" data-songs="${encodedDetails}">${finalVal}</td>`;
            }

            else tbody += `<td class="${cellStyle.trim()}">${finalVal}</td>`;
        });

        tbody += "</tr>";
    });

    table.innerHTML = thead + tbody + "</tbody>";
}

let currentTierChartMode = "TIER";

window.toggleTierChartMode = function() {
    const btn               = document.getElementById("tierModeToggleBtn");
    currentTierChartMode    = currentTierChartMode === "TIER" ? "ALL" : "TIER";
    btn.innerText           = currentTierChartMode;

    renderTierCharts();
};

function renderTierCharts() {
    if (!document.getElementById('tierChart_GuessRate') || !tierStats) return;

    const metrics = [
        {key: "Guess Rate",             title: "Guess Rate",            isAsc: false,   isRate: true,   isInt: false,   hoverDisabled: false, isTime: false},
        {key: "Lives Taken",            title: "Lives Taken",           isAsc: false,   isRate: false,  isInt: true,    hoverDisabled: false, isTime: false},
        {key: "Lives Saved",            title: "Lives Saved",           isAsc: false,   isRate: false,  isInt: true,    hoverDisabled: false, isTime: false},
        {key: "Contribution Rate",      title: "Contribution Rate",     isAsc: false,   isRate: true,   isInt: false,   hoverDisabled: false, isTime: false},
        {key: "Median Time",            title: "Median Time",           isAsc: true,    isRate: false,  isInt: false,   hoverDisabled: false, isTime: true},
        {key: "Chanting Guess Rate",    title: "Chanting Guess Rate",   isAsc: false,   isRate: true,   isInt: false,   hoverDisabled: false, isTime: false}
    ];

    const divIds = [
        "tierChart_GuessRate",
        "tierChart_LivesTaken",
        "tierChart_LivesSaved",
        "tierChart_ContributionRate",
        "tierChart_MedianTime",
        "tierChart_ChantingGuessRate"
    ];

    let gapCounter = 0;

    metrics.forEach((metric, mIdx) => {
        let xVals           = [];
        let yVals           = [];
        let customHovers    = [];

        const sortComparator = (a, b) => {
            let va = (a[metric.key] !== null && typeof a[metric.key] === 'object') ? a[metric.key].count : a[metric.key];
            let vb = (b[metric.key] !== null && typeof b[metric.key] === 'object') ? b[metric.key].count : b[metric.key];

            if (va === null || va === undefined) return 1;
            if (vb === null || vb === undefined) return -1;

            const fractionalMetrics = new Set(["Guess Rate", "Contribution Rate", "Chanting Guess Rate"]);

            if (va === vb && fractionalMetrics.has(metric.key)) {
                let numA = a[metric.key] && a[metric.key].details ? parseInt(a[metric.key].details[0].split('/')[0]) || 0 : 0;
                let numB = b[metric.key] && b[metric.key].details ? parseInt(b[metric.key].details[0].split('/')[0]) || 0 : 0;

                if (numA !== numB) return numB - numA;
            }

            return metric.isAsc ? va - vb : vb - va;
        };

        let playerPool = [];

        if (currentTierChartMode === "TIER") {
            ["1", "2", "3", "4"].forEach((tr) => {
                if (!tierStats[tr] || tierStats[tr].length === 0) return;
                let playersInTier = [...tierStats[tr]];
                playersInTier.sort(sortComparator);

                playerPool.push(...playersInTier);
                playerPool.push({isSpacer: true});
            });

            if (playerPool.length > 0 && playerPool[playerPool.length - 1].isSpacer) playerPool.pop();
        }

        else {
            ["1", "2", "3", "4"].forEach((tr) => {if (tierStats[tr]) playerPool.push(...tierStats[tr]);});
            playerPool.sort(sortComparator);
        }

        playerPool.forEach(p => {
            if (p.isSpacer) {
                xVals           .push(null);
                yVals           .push(" ".repeat(gapCounter++));
                customHovers    .push("");
                return;
            }

            let rawVal      = p[metric.key];
            let val         = (rawVal !== null && typeof rawVal === 'object') ? rawVal.count : rawVal;
            let finalVal    = 0;

            if (val !== null && val !== undefined && val !== Infinity) finalVal = metric.isInt ? Math.round(val) : Number(val.toFixed(2));

            xVals.push(finalVal);
            yVals.push(p.Player);

            if (!metric.hoverDisabled) {
                if (rawVal !== null && typeof rawVal === 'object' && rawVal.details && rawVal.details.length > 0) {
                    let displaySongs    = [...rawVal.details];
                    let fractionHeader  = "";
                    const fractionRegex = /^\d+\/\d+$/;

                    if (fractionRegex.test(displaySongs[0])) {
                        fractionHeader = `<b>${displaySongs[0]}</b>`;
                        displaySongs.shift(); 
                    }

                    if (metric.isTime) customHovers.push(displaySongs.join('<br>'));

                    else if (displaySongs.length > 10) {
                        displaySongs = sampleLargeSongList(displaySongs);
                        displaySongs.push(`and ${rawVal.details.length - 1 - 10} more`);
                        customHovers.push(fractionHeader ? `${fractionHeader}<br>${displaySongs.join('<br>')}` : displaySongs.join('<br>'));
                    }

                    else {
                        displaySongs = formatAndSortSongsList(displaySongs);
                        customHovers.push(fractionHeader ? `${fractionHeader}<br>${displaySongs.join('<br>')}` : displaySongs.join('<br>'));
                    }
                }

                else {
                    let detailKey   = metric.key + " Details";
                    let songs       = p[detailKey] || [];

                    if (songs.length > 0) {
                        let displaySongs = [...songs];

                        if (songs.length > 10) {
                            displaySongs = displaySongs.sort(() => Math.random() - 0.5).slice(0, 10);
                            displaySongs = formatAndSortSongsList(displaySongs, false);
                            customHovers.push("• " + displaySongs.join("<br>• ") + "<br>and " + (songs.length - 10) + " more");
                        }

                        else {
                            displaySongs = formatAndSortSongsList(displaySongs, false);
                            customHovers.push("• " + displaySongs.join("<br>• "));
                        }
                    }

                    else customHovers.push("No songs logged");
                }
            } 
            else customHovers.push("");
        });

        xVals           .reverse();
        yVals           .reverse();
        customHovers    .reverse();

        const trace = {
            x                   : xVals,
            y                   : yVals,
            type                : 'bar',
            orientation         : 'h',
            text                : xVals.map(v => v === null ? "" : (metric.isInt ? v.toFixed(0) : v.toFixed(2)) + " "),
            textposition        : 'inside',
            insidetextanchor    : 'end',
            textfont            : {family: 'Segoe UI', size: 15, color: 'black', weight: 'bold'},
            marker              : {color: 'white', line: {color: 'black', width: 2}}
        };

        if (metric.hoverDisabled) trace.hoverinfo = 'skip';

        else {
            trace.hovertext = customHovers;
            trace.hoverinfo = 'text';
        }

        const explanation = colExplanations[metric.key];

        const titleText = explanation 
            ? `<span style="font-size: 30px;"><b>${metric.title}</b></span><br><span style="font-size: 15px; font-weight: normal; color: 'black';">${explanation}</span>`
            : `<span style="font-size: 30px;"><b>${metric.title}</b></span>`;

        const layout = {
            font        : {family: 'Segoe UI'},
            title       : {text: titleText, font: {family: 'Segoe UI', size: 15, color: 'black'}, yref: 'container', y: 15, yanchor: 'top'},
            xaxis       : {tickfont: {family: 'Segoe UI', size: 15, color: 'black', weight: 'bold'}, fixedrange: true, showgrid: true},
            yaxis       : {tickfont: {family: 'Segoe UI', size: 15, color: 'black', weight: 'bold'}, fixedrange: true, showgrid: false, ticksuffix: "  " },
            bargap      : 0.0,
            margin      : {l: 150, r: 0, t: 100, b: 25},
            hoverlabel  : {align: 'left', font: {family: 'Segoe UI', size: 15}}
        };

        if      (metric.isRate) {layout.xaxis.tickmode = 'array'; layout.xaxis.tickvals = [0, 20, 40, 60, 80, 100]; layout.xaxis.range = [0, 105];}
        else if (metric.isTime) {layout.xaxis.tickmode = 'array'; layout.xaxis.tickvals = [0, 4, 8, 12, 16, 20];    layout.xaxis.range = [0, 21];}

        const totalPlayers  = playerPool.filter(p => !p.isSpacer).length;
        const totalSpacers  = playerPool.filter(p => p.isSpacer).length;
        layout.height       = 35 * (totalPlayers + totalSpacers);

        Plotly.newPlot(divIds[mIdx], [trace], layout, {responsive: true, displayModeBar: false});

        if (colExplanations[metric.key]) {setTimeout(() => {
            const titleEl = document.querySelector(`#${divIds[mIdx]} .g-title`);

            if (titleEl) {
                titleEl.style.cursor        = 'help';
                titleEl.style.pointerEvents = 'all';
                const cleanEl               = titleEl.cloneNode(true);

                titleEl.parentNode.replaceChild(cleanEl, titleEl);

                cleanEl.addEventListener('mouseenter', (e) => {
                    const tooltipNode = document.getElementById('customJsTooltip');
                    tooltipNode.innerHTML = colExplanations[metric.key]; tooltipNode.style.display = 'block';
                });

                cleanEl.addEventListener('mousemove', (e) => {
                    const tooltipNode = document.getElementById('customJsTooltip');

                    let xPos = e.pageX + 15;
                    let yPos = e.pageY + 15;

                    if (xPos + 450 > window.innerWidth + window.scrollX) xPos = e.pageX - 465;
                    tooltipNode.style.left = xPos + 'px'; tooltipNode.style.top = yPos + 'px';
                });

                cleanEl.addEventListener('mouseleave', () => { document.getElementById('customJsTooltip').style.display = 'none'; });
            }
        }, 300);}
    });
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

        th.addEventListener('mouseenter',   (e) => {tooltipNode.innerHTML = colExplanations[metricKey]; positionTooltip(e);});
        th.addEventListener('mousemove',    positionTooltip);
        th.addEventListener('mouseleave',   () => {tooltipNode.style.display = 'none';});
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
        hovertemplate   : '<b>%{customdata[0]}</b><br>Mean Over-8: %{customdata[1]}<br>Median Vintage: %{customdata[2]}<br>Guess Rate: %{customdata[3]}<br>Score: %{customdata[4]}<extra></extra>',
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
            hoverText   : `<b>${d.name}</b><br>Rig Over-8: ${d.x_start.toFixed(2)}<br>Rig Vintage: ${d.seasonal_vintage_start}<br>Rig Rate: ${(d.grid_rate !== undefined ? d.grid_rate : d.rig_rate).toFixed(2)}<br>Rig Guess Rate: ${d.rig_gr.toFixed(2)}<extra></extra>`
        })),

        "HIT": arrowData.map(d => ({
            acronym     : d.acronym,
            name        : d.name,
            x           : d.x_end, 
            y           : d.y_end,
            size        : d.rig_rate, 
            color       : d.grid_grs || d.rig_gr, 
            hoverText   : `<b>${d.name}</b><br>Hit Rig Over-8: ${d.x_end.toFixed(2)}<br>Hit Rig Vintage: ${d.seasonal_vintage || d.seasonal_vintage_end}<br>Rig Rate: ${(d.grid_rate !== undefined ? d.grid_rate : d.rig_rate).toFixed(2)}<br>Rig Guess Rate: ${d.rig_gr.toFixed(2)}<extra></extra>`
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
                title       : {text: '<b>Rig Guess Rate</b>', font: {family: 'Segoe UI', size: 25, color: 'black', weight: 'bold'}, side: 'right'},
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
        if (xAxisDesc) xAxisDesc.innerHTML = "<b>X-Axis:</b> Mean of correct guessers across songs that this player guessed correctly from their own list";
        if (yAxisDesc) yAxisDesc.innerHTML = "<b>Y-Axis:</b> Median vintage across songs that this player guessed correctly from their own list";
    }

    else {
        if (xAxisDesc) xAxisDesc.innerHTML = "<b>X-Axis:</b> Mean of correct guessers across songs from this player's list";
        if (yAxisDesc) yAxisDesc.innerHTML = "<b>Y-Axis:</b> Median vintage across songs from this player's list";
    }

    renderListChart();
};

let globalSearchData    = [];
let globalSortState     = {columnName: "Anime", ascending: true};
let currentSearchLang   = "JP"; 

const searchHeadersConfig = [
    {id: "anime",       name: "Anime",      visible: true},
    {id: "type",        name: "Song Type",  visible: true},
    {id: "chanting",    name: "Chanting",   visible: false},
    {id: "anime_type",  name: "Anime Type", visible: false},
    {id: "vintage",     name: "Vintage",    visible: false},
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

    let arr = Array.isArray(input) ? input : input.split(',').map(x => x.trim());
    arr     = arr.filter(Boolean);

    if (arr.length <= 3)    return arr.join(', ');
    else                    return `${arr.slice(0, 2).join(', ')}, and more`;
}

function parseVintageToFloat(vintStr) {
    const parts = vintStr.trim().split(/\s+/);    
    if (parts.length === 1 && !isNaN(parts[0])) return parseFloat(parts[0]);

    const season        = parts[0].toLowerCase();
    const year          = parseInt(parts[1]);
    let seasonWeight    = 0.0;

    if      (season === "winter")   seasonWeight = 0.1;
    else if (season === "spring")   seasonWeight = 0.2;
    else if (season === "summer")   seasonWeight = 0.3;
    else if (season === "fall")     seasonWeight = 0.4;

    return year + seasonWeight;
}

window.toggleSearchLanguage = function() {
    const btn           = document.getElementById("langToggleBtn");
    currentSearchLang   = currentSearchLang === "JP" ? "EN" : "JP";
    btn.innerText       = currentSearchLang;

    sortSearchData      ();
    triggerTableRefresh ();
};

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
        chk.className   = "col-toggle-checkbox rounded text-black focus:ring-black";
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
        case "songtype"     : return song._typeLower        .includes(value);
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
        const indicator = (globalSortState.columnName === c.name) ? (globalSortState.ascending ? " ▲" : " ▼") : " ▶";
        return `<th class="cursor-pointer select-none" style="white-space: nowrap;" data-header-name="${c.name}">${c.name}${indicator}</th>`;
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
                        const isOverflown   = song.artist_arr && song.artist_arr.length > 3;
                        td.className        = isOverflown ? "cursor-help hover:bg-gray-100 text-left text-black font-normal" : "text-left text-black font-normal";
                        if (isOverflown) td.setAttribute("data-songs", encodeURIComponent(JSON.stringify(song.artist_arr)));
                        td.textContent      = trimNames(song.artist_arr || []);
                    }

                    break;
                }

                case "composer": {
                    const matchArr = arrVisible && (song.composer === song.arranger);

                    if (matchArr) {
                        td.colSpan      = 2;
                        td.className    = "text-left font-normal text-black";
                        td.textContent  = trimNames(song.composer);
                        skipArranger    = true;
                    }

                    else {
                        td.className    = "text-left font-normal text-black";
                        td.textContent  = trimNames(song.composer);
                    }

                    break;
                }

                case "arranger":
                    td.className    = "text-left font-normal text-black";
                    td.textContent  = trimNames(song.arranger);
                    break;

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

            song._romajiLower     = (song.romaji        || "").toLowerCase();
            song._englishLower    = (song.english       || "").toLowerCase();
            song._songLower       = (song.song          || "").toLowerCase();
            song._artistRawLower  = (song.artist_raw    || "").toLowerCase();
            song._composerLower   = (song.composer      || "").toLowerCase();
            song._arrangerLower   = (song.arranger      || "").toLowerCase();
            song._typeLower       = (song.type          || "").toLowerCase();
            song._vintageLower    = (song.vintage       || "").toLowerCase();
            song._animeTypeLower  = (song.anime_type    || "").toLowerCase();
            song._chantingLower   = (song.chanting      || "").toLowerCase();
            song._vintageParsed   = parseVintageToFloat(song.vintage);
            song._diffParsed      = song.difficulty === "Unrated"   ? -Infinity                     : parseFloat(song.difficulty);
            song._guessersCount   = song.guessers_flat              ? song.guessers_flat    .length : 0;
            song._listersCount    = song.listers_flat               ? song.listers_flat     .length : 0;

            return song;
        });

        initColumnSettingsCheckboxes    ();
        sortSearchData                  ();
        renderSearchTable               (globalSearchData);

        const searchInput = document.getElementById('songSearchInput');
        if (searchInput) {

            const processQuery = (e) => {
                const rawQuery = searchInput.value.trim();

                if (!rawQuery) {
                    renderSearchTable(globalSearchData);
                    return;
                }

                const tokens        = [];
                const tokenRegex    = /\(|\)|or\b|and\b|[^\s"()]+|"[^"]*"/gi;

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

                        return evaluateQuery(song, queryKey, parsedMatch[2], parsedMatch[3].replace(/^"|"$/g, '').toLowerCase().trim());
                    }

                    const wordClean = token.replace(/^"|"$/g, '').toLowerCase();

                    return (
                        song._romajiLower       .includes(wordClean) ||
                        song._englishLower      .includes(wordClean) ||
                        song._songLower         .includes(wordClean) ||
                        song._artistRawLower    .includes(wordClean) ||
                        song._composerLower     .includes(wordClean) ||
                        song._arrangerLower     .includes(wordClean) ||
                        song._typeLower         .includes(wordClean) ||
                        song._vintageLower      .includes(wordClean) ||
                        song._animeTypeLower    .includes(wordClean) ||
                        song._chantingLower     .includes(wordClean) ||
                        song.difficulty         .toLowerCase().includes(wordClean)
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