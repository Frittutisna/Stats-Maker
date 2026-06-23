// Destructure application contexts from fetched dynamic runtime global cache
const {
    prefix, use_teams, watched, num_x: numX, num_y: numY, c0, c1, c2,
    json_players: players, json_tour_stats: tourStats, json_teams: teamStats,
    json_team_hl_rules: teamHlRules, json_tier_merged: tierStats, json_songs: songData,
    json_matrix_songs: matrixSongs, json_scatter: scatterData, json_arrows: arrowData,
    json_borders: groupBorders, json_eligibility: eligibility, json_hl_rules: hlRules,
    json_explanations: colExplanations, generated_timestamp: generatedTime
} = window.dashboardData;

document.getElementById('dashboardTitle').innerText = prefix;

// Setup custom programmatic accent style definitions based on internal properties
const dynamicStyles = document.createElement('style');
dynamicStyles.innerHTML = `
    .highlight-best { background-color: ${c2} !important; }
    .highlight-worst { background-color: ${c0} !important; }
    td[data-songs].highlight-best:hover { color: ${c2} !important; }
    td[data-songs].highlight-worst:hover { color: ${c0} !important; }
    td[data-songs].highlight-best:hover::after { background-color: ${c2} !important; }
    td[data-songs].highlight-worst:hover::after { background-color: ${c0} !important; }
`;
document.head.appendChild(dynamicStyles);

// Reconstruct structural components based on tour contextual flags
const tabContainer = document.getElementById('tabContainer');
if (use_teams && watched) {
    tabContainer.insertAdjacentHTML('beforeend', `<button class="tab-btn" onclick="switchDashboardTab(event, 'team-tab')">Team</button>`);
    document.getElementById('team-tab-container').outerHTML = `<div id='team-tab' class='tab-content'><div class='table-center-wrapper'><table class='main-table' id='teamStatsTable'></table></div></div>`;
}
if (use_teams) {
    tabContainer.insertAdjacentHTML('beforeend', `<button class="tab-btn" onclick="switchDashboardTab(event, 'tier-tab')">Tier</button>`);
    document.getElementById('tier-tab-container').outerHTML = `<div id='tier-tab' class='tab-content'><div class='max-w-[1200px] mx-auto space-y-8 bg-white p-6 rounded shadow-md border border-gray-300'><div id='tierChart_GuessRate'></div><div id='tierChart_LivesTaken'></div><div id='tierChart_LivesSaved'></div><div id='tierChart_ContributionRate'></div><div id='tierChart_MedianTime'></div><div id='tierChart_ChantingGuessRate'></div></div></div>`;
}
tabContainer.insertAdjacentHTML('beforeend', `<button class="tab-btn" onclick="switchDashboardTab(event, 'song-tab')">Song</button>`);
tabContainer.insertAdjacentHTML('beforeend', `<button class="tab-btn" onclick="switchDashboardTab(event, 'guess-tab')">Guess</button>`);
if (watched) {
    tabContainer.insertAdjacentHTML('beforeend', `<button class="tab-btn" onclick="switchDashboardTab(event, 'list-tab')">List</button>`);
    document.getElementById('list-tab-container').outerHTML = `<div id='list-tab' class='tab-content'><div class='max-w-[1200px] mx-auto border border-gray-300 p-4 bg-white rounded shadow-md'><div class='mb-4 text-lg text-black space-y-1'><p><b>X-Axis:</b> Mean of correct guessers across songs from this player\'s list</p><p><b>Y-Axis:</b> Median vintage across songs from this player\'s list</p><p><b>Size (Rig Rate)</b></p></div><div id='plotlyListChart' style='width:100%; height:750px;'></div></div></div>`;
}

const colBorders = new Set(["Player", "Guess Rate", "Score", "Mean Over-8", "Lives Saved", "IN Guess Rate", "Rig Rate", "Solo Rig Rate", "Over-8 Delta", "Rig Delta", "Metric", "Value", "Team Leader", "Tier", "Lives Saved", "Chanting Guess Rate"]);

function updateTimeAgoSubtitle() {
    const subNode = document.getElementById('lastUpdatedSubtitle');
    if (!subNode) return;

    const diffMs = Date.now() - generatedTime;
    const diffSec = Math.floor(diffMs / 1000);
    const diffMin = Math.floor(diffSec / 60);
    const diffHr = Math.floor(diffMin / 60);
    const diffDays = Math.floor(diffHr / 24);

    let displayString = "Last updated: ";
    if (diffSec < 60) displayString += `${diffSec} seconds ago`;
    else if (diffMin < 60) displayString += `${diffMin} minute${diffMin === 1 ? '' : 's'} ago`;
    else if (diffHr < 24) displayString += `${diffHr} hour${diffHr === 1 ? '' : 's'} ago`;
    else displayString += `${diffDays} day${diffDays === 1 ? '' : 's'} ago`;

    subNode.innerText = displayString;
}
updateTimeAgoSubtitle();
setInterval(updateTimeAgoSubtitle, 1000);

function switchDashboardTab(evt, tabId) {
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active-content'));
    document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active-tab'));
    document.getElementById(tabId).classList.add('active-content');
    evt.currentTarget.classList.add('active-tab');
    window.dispatchEvent(new Event('resize'));
}

function crossProduct(o, a, b, xKey, yKey) {
    return (a[xKey] - o[xKey]) * (b[yKey] - o[yKey]) - (a[yKey] - o[yKey]) * (b[xKey] - o[xKey]);
}

function get75PercentileHull(pts, xKey, yKey) {
    if (pts.length < 3) return null;
    const xVals = pts.map(p => p[xKey]).sort((a,b) => a-b);
    const yVals = pts.map(p => p[yKey]).sort((a,b) => a-b);
    const medX = xVals[Math.floor(xVals.length / 2)];
    const medY = yVals[Math.floor(yVals.length / 2)];
    const xRange = (Math.max(...xVals) - Math.min(...xVals)) || 1;
    const yRange = (Math.max(...yVals) - Math.min(...yVals)) || 1;

    const withDist = pts.map(p => {
        const dx = (p[xKey] - medX) / xRange;
        const dy = (p[yKey] - medY) / yRange;
        return { p, d: Math.sqrt(dx*dx + dy*dy) };
    });

    const sortedDist = withDist.map(item => item.d).sort((a,b) => a-b);
    const threshD = sortedDist[Math.floor(sortedDist.length * 0.75)];
    const packedPts = withDist.filter(item => item.d < threshD).map(item => item.p);

    if (packedPts.length < 3) return null;
    packedPts.sort((a, b) => a[xKey] == b[xKey] ? a[yKey] - b[yKey] : a[xKey] - b[xKey]);
    
    const lower = [];
    for (let p of packedPts) {
        while (lower.length >= 2 && crossProduct(lower[lower.length-2], lower[lower.length-1], p, xKey, yKey) <= 0) lower.pop();
        lower.push(p);
    }
    const upper = [];
    for (let i = packedPts.length - 1; i >= 0; i--) {
        let p = packedPts[i];
        while (upper.length >= 2 && crossProduct(upper[upper.length-2], upper[upper.length-1], p, xKey, yKey) <= 0) upper.pop();
        upper.push(p);
    }
    upper.pop(); lower.pop();
    const hull = lower.concat(upper);
    return {
        x: hull.map(p => p[xKey]).concat(hull[0][xKey]),
        y: hull.map(p => p[yKey]).concat(hull[0][yKey])
    };
}

function renderPlayerTable() {
    const table = document.getElementById('playerStandingsTable');
    if(!players || !players.length) return;

    let headers = Object.keys(players[0]);
    let thead = "<thead><tr>" + headers.map(h => {
        let classes = [];
        if (colBorders.has(h)) classes.push("border-col-group");
        if (colExplanations[h]) classes.push("has-explanation");
        let classStr = classes.length > 0 ? ` class="${classes.join(' ')}"` : '';
        return `<th${classStr} data-metric="${h}">${h.replace(/ /g, '<br>')}</th>`;
    }).join('') + "</tr></thead>";

    let tbody = "<tbody>";
    players.forEach((row, idx) => {
        let groupLine = groupBorders.includes(idx) ? " border-group-line" : "";
        tbody += `<tr class="${groupLine}">`;
        
        headers.forEach(h => {
            let rawCell = row[h];
            let displayVal = (rawCell !== null && typeof rawCell === 'object') ? rawCell.count : rawCell;
            let cellStyle = colBorders.has(h) ? "border-col-group " : "";
            
            if (hlRules[h]) {
                let isBest = (hlRules[h].best_idx === idx);
                let isWorst = (hlRules[h].worst_idx === idx);
                if (isBest) cellStyle += "highlight-best ";
                else if (isWorst) cellStyle += "highlight-worst ";
            }

            let formattedVal = (typeof displayVal === 'number' && h !== "1/8s" && h !== "2/8s" && h !== "7/8s" && h !== "Lives Taken" && h !== "Lives Saved" && h !== "Rigs" && h !== "Solo Rigs") ? displayVal.toFixed(2) : displayVal;
            let finalVal = (h === "Player") ? `<b>${formattedVal}</b>` : formattedVal;
            
            if (h === "Player" && rawCell && rawCell.details && rawCell.details.length > 0) {
                let encodedDetails = encodeURIComponent(JSON.stringify(rawCell.details));
                tbody += `<td class="${cellStyle.trim()}" data-songs="${encodedDetails}">${finalVal}</td>`;
            } else if (rawCell !== null && typeof rawCell === 'object' && rawCell.details && rawCell.details.length > 0) {
                let encodedDetails = encodeURIComponent(JSON.stringify(rawCell.details));
                tbody += `<td class="${cellStyle.trim()}" data-songs="${encodedDetails}">${finalVal}</td>`;
            } else {
                tbody += `<td class="${cellStyle.trim()}">${finalVal}</td>`;
            }
        });
        tbody += "</tr>";
    });
    table.innerHTML = thead + tbody + "</tbody>";
}

function renderTourTable() {
    const table = document.getElementById('tourStatsTable');
    if(!tourStats) return;
    
    let tourHeaders = ['Metric', 'Value'];
    let thead = "<thead><tr>" + tourHeaders.map(h => {
        let classes = [];
        if (colBorders.has(h)) classes.push("border-col-group");
        if (colExplanations[h]) classes.push("has-explanation");
        let classStr = classes.length > 0 ? ` class="${classes.join(' ')}"` : '';
        return `<th${classStr} data-metric="${h}">${h.replace(/ /g, '<br>')}</th>`;
    }).join('') + "</tr></thead>";

    let tbody = "";
    tourStats.forEach(row => {
        let rawCell = row.Value;
        let displayVal = (rawCell !== null && typeof rawCell === 'object') ? rawCell.count : rawCell;
        let metricClass = colExplanations[row.Metric] ? "border-col-group has-explanation" : "border-col-group";
        if (rawCell !== null && typeof rawCell === 'object' && rawCell.details && rawCell.details.length > 0) {
            let encodedDetails = encodeURIComponent(JSON.stringify(rawCell.details));
            tbody += `<tr><td class='${metricClass}'><b>${row.Metric}</b></td><td data-songs="${encodedDetails}">${displayVal}</td></tr>`;
        } else {
            tbody += `<tr><td class='${metricClass}'><b>${row.Metric}</b></td><td>${displayVal}</td></tr>`;
        }
    });
    table.innerHTML = thead + tbody + "</tbody>";
}

function renderTeamTable() {
    const table = document.getElementById('teamStatsTable');
    if(!table || !teamStats || !teamStats.length) return;
    
    let headers = Object.keys(teamStats[0]);
    let thead = "<thead><tr>" + headers.map(h => {
        let classes = [];
        if (colBorders.has(h)) classes.push("border-col-group");
        if (colExplanations[h]) classes.push("has-explanation");
        let classStr = classes.length > 0 ? ` class="${classes.join(' ')}"` : '';
        return `<th${classStr} data-metric="${h}">${h.replace(/ /g, '<br>')}</th>`;
    }).join('') + "</tr></thead>";
    
    let tbody = "<tbody>";
    teamStats.forEach((row, idx) => {
        tbody += "<tr>";
        headers.forEach(h => {
            let rawCell = row[h];
            let displayVal = (rawCell !== null && typeof rawCell === 'object') ? rawCell.count : rawCell;
            let cellStyle = colBorders.has(h) ? "border-col-group " : "";
            
            if (teamHlRules[h]) {
                let isBest = (teamHlRules[h].best_idx === idx);
                let isWorst = (teamHlRules[h].worst_idx === idx);
                if (isBest) cellStyle += "highlight-best ";
                else if (isWorst) cellStyle += "highlight-worst ";
            }
            
            let finalVal = (h === "Team Leader") ? `<b>${displayVal}</b>` : displayVal;
            
            if (rawCell !== null && typeof rawCell === 'object' && rawCell.details && rawCell.details.length > 0) {
                let encodedDetails = encodeURIComponent(JSON.stringify(rawCell.details));
                tbody += `<td class="${cellStyle.trim()}" data-songs="${encodedDetails}">${finalVal}</td>`;
            } else {
                tbody += `<td class="${cellStyle.trim()}">${finalVal}</td>`;
            }
        });
        tbody += "</tr>";
    });
    table.innerHTML = thead + tbody + "</tbody>";
}

function renderTierCharts() {
    if (!document.getElementById('tierChart_GuessRate') || !tierStats) return;
    const metrics = [
        { key: "Guess Rate", title: "Guess Rate", isAsc: false, isRate: true, hoverDisabled: false },
        { key: "Lives Taken", title: "Lives Taken", isAsc: false, isRate: false, isInt: true },
        { key: "Lives Saved", title: "Lives Saved", isAsc: false, isRate: false, isInt: true },
        { key: "Contribution Rate", title: "Contribution Rate", isAsc: false, isRate: true, hoverDisabled: false },
        { key: "Median Time", title: "Median Time", isAsc: true, isRate: false, isTime: true, hoverDisabled: true },
        { key: "Chanting Guess Rate", title: "Chanting Guess Rate", isAsc: false, isRate: true, hoverDisabled: false }
    ];

    const divIds = [
        "tierChart_GuessRate", "tierChart_LivesTaken", "tierChart_LivesSaved",
        "tierChart_ContributionRate", "tierChart_MedianTime", "tierChart_ChantingGuessRate"
    ];

    let gapCounter = 0;
    metrics.forEach((metric, mIdx) => {
        let xVals = [];
        let yVals = [];
        let customHovers = [];

        ["1", "2", "3", "4"].forEach((tr, tIdx) => {
            if (!tierStats[tr] || tierStats[tr].length === 0) return;

            let playersInTier = [...tierStats[tr]];
            playersInTier.sort((a, b) => {
                let va = (a[metric.key] !== null && typeof a[metric.key] === 'object') ? a[metric.key].count : a[metric.key];
                let vb = (b[metric.key] !== null && typeof b[metric.key] === 'object') ? b[metric.key].count : b[metric.key];
                if (va === null || va === undefined) return 1;
                if (vb === null || vb === undefined) return -1;
                if (va === vb && (metric.key === "Guess Rate" || metric.key === "Chanting Guess Rate")) {
                    let numA = a[metric.key] && a[metric.key].details ? parseInt(a[metric.key].details[0].split('/')[0]) || 0 : 0;
                    let numB = b[metric.key] && b[metric.key].details ? parseInt(b[metric.key].details[0].split('/')[0]) || 0 : 0;
                    if (numA !== numB) return numB - numA;
                }
                return metric.isAsc ? va - vb : vb - va;
            });

            if (xVals.length > 0) {
                xVals.push(null);
                yVals.push(" ".repeat(gapCounter++));
                customHovers.push("");
            }

            playersInTier.forEach(p => {
                let rawVal = p[metric.key];
                let val = (rawVal !== null && typeof rawVal === 'object') ? rawVal.count : rawVal;
                let finalVal = 0;
                if (val !== null && val !== undefined && val !== Infinity) {
                    finalVal = metric.isInt ? Math.round(val) : Number(val.toFixed(2));
                }

                xVals.push(finalVal);
                yVals.push(p.Player);

                if (!metric.hoverDisabled) {
                    if (rawVal !== null && typeof rawVal === 'object' && rawVal.details && rawVal.details.length > 0) {
                        customHovers.push(rawVal.details[0]);
                    } else {
                        let detailKey = metric.key + " Details";
                        let songs = p[detailKey] || [];
                        if (songs.length > 0) {
                            let displaySongs = [...songs];
                            if (songs.length > 10) {
                                displaySongs = displaySongs.sort(() => Math.random() - 0.5).slice(0, 10);
                                displaySongs.sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
                                customHovers.push("• " + displaySongs.join("<br>• ") + "<br>and " + (songs.length - 10) + " more");
                            } else {
                                displaySongs.sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
                                customHovers.push("• " + displaySongs.join("<br>• "));
                            }
                        } else {
                            customHovers.push("• No songs logged");
                        }
                    }
                } else {
                    customHovers.push("");
                }
            });
        });

        xVals.reverse(); yVals.reverse(); customHovers.reverse();

        const trace = {
            x: xVals, y: yVals, type: 'bar', orientation: 'h',
            text: xVals.map(v => v === null ? "" : (metric.isInt ? v.toFixed(0) : v.toFixed(2))),
            textposition: 'inside', insidetextanchor: 'end',
            textfont: { family: 'Segoe UI', size: 14, color: 'black', weight: 'bold' },
            marker: { color: 'white', line: { color: 'black', width: 2 } }
        };

        if (metric.hoverDisabled) trace.hoverinfo = 'skip';
        else { trace.hovertext = customHovers; trace.hoverinfo = 'text'; }

        const explanation = colExplanations[metric.key];
        const titleText = explanation 
            ? `<span style="font-size: 30px;"><b>${metric.title}</b></span><br><span style="font-size: 15px; font-weight: normal; color: 'black';">${explanation}</span>`
            : `<span style="font-size: 30px;"><b>${metric.title}</b></span>`;

        const layout = {
            font: { family: 'Segoe UI' },
            title: { text: titleText, font: { family: 'Segoe UI', size: 14, color: 'black' }, y: 0.95, yanchor: 'top' },
            xaxis: { tickfont: { family: 'Segoe UI', size: 16, color: 'black', weight: 'bold' }, showgrid: true, zeroline: true, fixedrange: true },
            yaxis: { tickfont: { family: 'Segoe UI', size: 16, color: 'black', weight: 'bold' }, type: 'category', fixedrange: true, ticksuffix: "  " },
            bargap: 0.0, margin: { l: 200, r: 50, t: 100, b: 100 }, height: 145 + (yVals.length * 30),
            hoverlabel: { align: 'left', font: { family: 'Segoe UI', size: 15 }}
        };

        if (metric.isRate) { layout.xaxis.tickmode = 'array'; layout.xaxis.tickvals = [0, 20, 40, 60, 80, 100]; layout.xaxis.range = [0, 105]; }
        else if (metric.isTime) { layout.xaxis.tickmode = 'array'; layout.xaxis.tickvals = [0, 4, 8, 12, 16, 20]; layout.xaxis.range = [0, 21]; }

        Plotly.newPlot(divIds[mIdx], [trace], layout, { responsive: true, displayModeBar: false });

        if (colExplanations[metric.key]) {
            setTimeout(() => {
                const titleEl = document.querySelector(`#${divIds[mIdx]} .g-title`);
                if (titleEl) {
                    titleEl.style.cursor = 'help'; titleEl.style.pointerEvents = 'all';
                    titleEl.addEventListener('mouseenter', (e) => {
                        const tooltipNode = document.getElementById('customJsTooltip');
                        tooltipNode.innerHTML = colExplanations[metric.key]; tooltipNode.style.display = 'block';
                    });
                    titleEl.addEventListener('mousemove', (e) => {
                        const tooltipNode = document.getElementById('customJsTooltip');
                        let xPos = e.pageX + 15; let yPos = e.pageY + 15;
                        if (xPos + 450 > window.innerWidth + window.scrollX) xPos = e.pageX - 465;
                        tooltipNode.style.left = xPos + 'px'; tooltipNode.style.top = yPos + 'px';
                    });
                    titleEl.addEventListener('mouseleave', () => { document.getElementById('customJsTooltip').style.display = 'none'; });
                }
            }, 300);
        }
    });
}

function setupTooltipListeners() {
    const tooltipNode = document.getElementById('customJsTooltip');
    function positionTooltip(e) {
        tooltipNode.style.display = 'block';
        const tooltipWidth = tooltipNode.offsetWidth; const tooltipHeight = tooltipNode.offsetHeight;
        let xPos = e.pageX + 15; let yPos = e.pageY + 15;
        if (e.clientX + 15 + tooltipWidth > window.innerWidth) xPos = e.pageX - tooltipWidth - 15;
        if (e.clientY + 15 + tooltipHeight > window.innerHeight) yPos = e.pageY - tooltipHeight - 15;
        if (xPos < window.scrollX) xPos = window.scrollX + 5;
        if (yPos < window.scrollY) yPos = window.scrollY + 5;
        tooltipNode.style.left = xPos + 'px'; tooltipNode.style.top = yPos + 'px';
    }

    document.querySelectorAll('table th[data-metric]').forEach(th => {
        const metricKey = th.getAttribute('data-metric');
        if (!colExplanations[metricKey]) return;
        th.addEventListener('mouseenter', (e) => { tooltipNode.innerHTML = colExplanations[metricKey]; positionTooltip(e); });
        th.addEventListener('mousemove', positionTooltip);
        th.addEventListener('mouseleave', () => { tooltipNode.style.display = 'none'; });
    });

    document.querySelectorAll('#tourStatsTable tr td:first-child').forEach(td => {
        const metricKey = td.innerText.trim();
        if (!colExplanations[metricKey]) return;
        td.addEventListener('mouseenter', (e) => { tooltipNode.innerHTML = colExplanations[metricKey]; positionTooltip(e); });
        td.addEventListener('mousemove', positionTooltip);
        td.addEventListener('mouseleave', () => { tooltipNode.style.display = 'none'; });
    });

    document.querySelectorAll('td[data-songs]').forEach(td => {
        td.addEventListener('mouseenter', (e) => {
            try {
                const songs = JSON.parse(decodeURIComponent(td.getAttribute('data-songs')));
                if(!songs || songs.length === 0) return;
                let displaySongs = [...songs];
                const isPlayerSubHover = td.parentNode.firstElementChild === td;
                
                if (songs.length === 1 && !songs[0].startsWith('✓') && !songs[0].startsWith('✗') && songs[0].includes('/')) {
                    tooltipNode.innerHTML = songs[0]; positionTooltip(e); return;
                }

                if (songs.length > 10) {
                    displaySongs = displaySongs.sort(() => Math.random() - 0.5).slice(0, 10);
                    displaySongs.sort((a, b) => {
                        const cleanA = (a.startsWith('✓') || a.startsWith('✗')) ? a.slice(2) : a;
                        const cleanB = (b.startsWith('✓') || b.startsWith('✗')) ? b.slice(2) : b;
                        return cleanA.toLowerCase().localeCompare(cleanB.toLowerCase());
                    });
                    displaySongs = displaySongs.map(s => (s.startsWith('✓') || s.startsWith('✗')) ? s : isPlayerSubHover ? s : `• ${s}`);
                    displaySongs.push(`and ${songs.length - 10} more`);
                } else {
                    displaySongs.sort((a, b) => {
                        const cleanA = (a.startsWith('✓') || a.startsWith('✗')) ? a.slice(2) : a;
                        const cleanB = (b.startsWith('✓') || b.startsWith('✗')) ? b.slice(2) : b;
                        return cleanA.toLowerCase().localeCompare(cleanB.toLowerCase());
                    });
                    displaySongs = displaySongs.map(s => (s.startsWith('✓') || s.startsWith('✗')) ? s : isPlayerSubHover ? s : `• ${s}`);
                }
                tooltipNode.innerHTML = displaySongs.join('<br>'); positionTooltip(e);
            } catch(err) {}
        });
        td.addEventListener('mousemove', positionTooltip);
        td.addEventListener('mouseleave', () => { tooltipNode.style.display = 'none'; });
    });
}

// Render Lifecycle Invocation
renderPlayerTable();
renderTourTable();
renderTeamTable();
renderTierCharts();
setupTooltipListeners();

// Initialize Heatmaps & Diagrams Matrix Configurations
const xLabels = (numX === 8) ? ['5', '10', '15', '20', '25', '30', '35'] : ['5', '10', '15', '20', '25', '30', '35', '40'];
const yLabels = (numY === 8) ? [1990, 1995, 2000, 2005, 2010, 2015, 2020, 2025] : [1985, 1990, 1995, 2000, 2005, 2010, 2015, 2020, 2025];

const matrixBins = {};
songData.forEach(s => {
    let xIdx = Math.min(Math.floor(s.difficulty / 5), numX - 1);
    let yIdx = (numY === 8) ? ((s.vintage < 1990) ? 0 : Math.min(Math.floor((s.vintage - 1990) / 5) + 1, 7)) : Math.min(Math.max(Math.floor((s.vintage - 1985) / 5), 0), 8);
    let key = `${xIdx}-${yIdx}`;
    if(!matrixBins[key]) matrixBins[key] = { count: 0, over8Sum: 0 };
    matrixBins[key].count++; matrixBins[key].over8Sum += s.correct_count;
});

let zValues = [], textLabels = [], annotations = [];
for (let i = 0; i < numY; i++) {
    let let_rowZ = [], let_rowText = [];
    for (let j = 0; j < numX; j++) {
        let key = `${j}-${i}`;
        if (key in matrixBins) {
            let val = matrixBins[key].over8Sum / matrixBins[key].count;
            let_rowZ.push(val);
            let bin_songs = matrixSongs[key] ? [...matrixSongs[key]] : [];
            let song_hover_str = "";
            
            if (bin_songs.length > 10) {
                const remainingCount = bin_songs.length - 10;
                bin_songs = bin_songs.sort(() => Math.random() - 0.5).slice(0, 10);
                bin_songs.sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
                song_hover_str = "<br>• " + bin_songs.join("<br>• ") + "<br>and " + remainingCount + " more";
            } else if (bin_songs.length > 0) {
                bin_songs.sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
                song_hover_str = "<br>• " + bin_songs.join("<br>• ");
            }
            let_rowText.push(`Mean Over-8: ${val.toFixed(2)}${song_hover_str}`);
            annotations.push({
                x: j, y: i, text: `<b>${matrixBins[key].count}</b>`,
                font: { family: 'Segoe UI', size: (numX > 8 ? 48 : 55), color: 'white' },
                showarrow: false, captureevents: false
            });
        } else { let_rowZ.push(null); let_rowText.push(''); }
    }
    zValues.push(let_rowZ); textLabels.push(let_rowText);
}

Plotly.newPlot('plotlySongChart', [{
    z: zValues, x: Array.from({length: numX}, (_, i) => i), y: Array.from({length: numY}, (_, i) => i),
    text: textLabels, hovertemplate: '<span style="text-align: left; display: block;">%{text}</span><extra></extra>',
    hoverlabel: { align: 'left' }, type: 'heatmap', colorscale: [[0, c0], [0.375, c1], [0.625, c2], [1, c2]],
    zmin: 0, zmax: 8, showscale: true,
    colorbar: {
        title: { text: '<b>Over-8</b>', font: { family: 'Segoe UI', size: 25, color: 'black', weight: 'bold' }, side: 'right' },
        thickness: 25, len: 1.0, y: 0.5, yanchor: 'middle', x: 1, xpad: -20,
        tickmode: 'array', tickvals: [0, 3, 5, 8], ticktext: ['0', '3', '5', '8'],
        tickfont: { family: 'Segoe UI', size: 20, color: 'black', weight: 'bold' }
    }
}], {
    font: { family: 'Segoe UI' },
    xaxis: {
        title: { text: '<b>Difficulty</b>', font: { family: 'Segoe UI', size: 25, color: 'black', weight: 'bold' }, pad: 5 },
        tickmode: 'array', tickvals: Array.from({length: numX - 1}, (_, i) => i + 0.5), ticktext: xLabels,
        tickfont: { family: 'Segoe UI', size: 20, color: 'black', weight: 'bold' },
        showgrid: true, zeroline: false, showticklabels: true, ticks: '', fixedrange: true
    },
    yaxis: {
        title: { text: '<b>Vintage</b>', font: { family: 'Segoe UI', size: 25, color: 'black', weight: 'bold' }, pad: 5 },
        tickmode: 'array', tickvals: Array.from({length: numY - 1}, (_, i) => i + 0.5), ticktext: yLabels,
        tickfont: { family: 'Segoe UI', size: 20, color: 'black', weight: 'bold' },
        tickangle: -90, showgrid: true, zeroline: false, showticklabels: true, ticks: '', fixedrange: true
    },
    annotations: annotations, margin: { l: 60, r: 0, t: 30, b: 55 }
}, {responsive: true, displayModeBar: false});

if (document.getElementById('plotlyListChart') && arrowData) {
    const listHull = get75PercentileHull(arrowData, 'x_start', 'y_start');
    let listTraces = [];
    if (listHull) {
        listTraces.push({ x: listHull.x, y: listHull.y, type: 'scatter', mode: 'lines', line: { color: 'black', width: 0.5, dash: 'solid' }, hoverinfo: 'skip', showlegend: false });
    }
    listTraces.push({
        x: arrowData.map(d => d.x_start), y: arrowData.map(d => d.y_start), text: arrowData.map(d => d.acronym),
        customdata: arrowData.map(d => [d.name, d.x_start.toFixed(2), d.seasonal_vintage_start, d.rig_rate.toFixed(2), d.rig_gr.toFixed(2)]),
        hovertemplate: '<b>%{customdata[0]}</b><br>Rig Over-8: %{customdata[1]}<br>Rig Vintage: %{customdata[2]}<br>Rig Rate: %{customdata[3]}<br>Rig Guess Rate: %{customdata[4]}<extra></extra>',
        mode: 'markers+text', textposition: 'top center', textfont: { family: 'Segoe UI', size: 20, weight: 'bold', color: 'black' }, showlegend: false,
        marker: {
            size: arrowData.map(d => Math.max(25, Math.pow(d.rig_rate, 2) * 0.025)), opacity: 1, color: arrowData.map(d => d.grid_grs || d.rig_gr),
            colorscale: [[0, c0], [0.7, c0], [0.8, c1], [0.9, c2], [1, c2]], showscale: true,
            colorbar: {
                title: { text: '<b>Rig Guess Rate</b>', font: { family: 'Segoe UI', size: 25, color: 'black', weight: 'bold' }, side: 'right' },
                thickness: 25, len: 1.0, y: 0.5, yanchor: 'middle', x: 1, xpad: -20,
                tickmode: 'array', tickvals: [0, 70, 80, 90, 100], ticktext: ['0', '70', '80', '90', '100'],
                tickfont: { family: 'Segoe UI', size: 20, color: 'black', weight: 'bold' }
            },
            line: { color: 'black', width: 1 }, cmin: 0, cmax: 100
        }
    });

    Plotly.newPlot('plotlyListChart', listTraces, {
        font: { family: 'Segoe UI' },
        xaxis: { title: { text: '<b>Over-8</b>', font: { family: 'Segoe UI', size: 25, color: 'black', weight: 'bold' }, pad: 5 }, tickfont: { family: 'Segoe UI', size: 20, color: 'black', weight: 'bold' }, showgrid: true, tickformat: '.1f', dtick: 0.5, fixedrange: false },
        yaxis: { title: { text: '<b>Vintage</b>', font: { family: 'Segoe UI', size: 25, color: 'black', weight: 'bold' }, pad: 5 }, tickfont: { family: 'Segoe UI', size: 20, color: 'black', weight: 'bold' }, tickangle: -90, showgrid: true, tickformat: '.0f', dtick: Math.max(2, Math.ceil((Math.max(...arrowData.map(d => d.y_start)) - Math.min(...arrowData.map(d => d.y_start))) / 5)), fixedrange: false },
        margin: { l: 60, r: 0, t: 30, b: 55 },
        annotations: [
            { x: 0, y: 1, xref: 'paper', yref: 'paper', text: '<b>New<br>Hard</b>', showarrow: false, font: { size: 20 }, opacity: 0.75, xanchor: 'left', yanchor: 'top' },
            { x: 1, y: 1, xref: 'paper', yref: 'paper', text: '<b>New<br>Easy</b>', showarrow: false, font: { size: 20 }, opacity: 0.75, xanchor: 'right', yanchor: 'top' },
            { x: 1, y: 0, xref: 'paper', yref: 'paper', text: '<b>Old<br>Easy</b>', showarrow: false, font: { size: 20 }, opacity: 0.75, xanchor: 'right', yanchor: 'bottom' },
            { x: 0, y: 0, xref: 'paper', yref: 'paper', text: '<b>Old<br>Hard</b>', showarrow: false, font: { size: 20 }, opacity: 0.75, xanchor: 'left', yanchor: 'bottom' }
        ]
    }, {responsive: true, displayModeBar: false});
}

if(scatterData) {
    const guessHull = get75PercentileHull(scatterData, 'over8', 'vintage');
    let guessTraces = [];
    if (guessHull) {
        guessTraces.push({ x: guessHull.x, y: guessHull.y, type: 'scatter', mode: 'lines', line: { color: 'black', width: 0.5, dash: 'solid' }, hoverinfo: 'skip', showlegend: false });
    }
    guessTraces.push({
        x: scatterData.map(d => d.over8), y: scatterData.map(d => d.vintage), text: scatterData.map(d => d.acronym),
        customdata: scatterData.map(d => [d.name, d.over8.toFixed(2), d.seasonal_vintage, d.gr.toFixed(2), d.performance.toFixed(2)]),
        hovertemplate: '<b>%{customdata[0]}</b><br>Mean Over-8: %{customdata[1]}<br>Median Vintage: %{customdata[2]}<br>Guess Rate: %{customdata[3]}<br>Score: %{customdata[4]}<extra></extra>',
        mode: 'markers+text', textposition: 'top center', textfont: { family: 'Segoe UI', size: 20, weight: 'bold', color: 'black' }, showlegend: false,
        marker: {
            size: scatterData.map(d => Math.max(25, Math.pow(d.gr, 2) * 0.025)), opacity: 1, color: scatterData.map(d => d.performance),
            colorscale: [[0, c0], [0.5, c1], [1, c2]], showscale: true,
            colorbar: {
                title: { text: '<b>Score</b>', font: { family: 'Segoe UI', size: 25, color: 'black', weight: 'bold' }, side: 'right' },
                thickness: 25, len: 1.0, y: 0.5, yanchor: 'middle', x: 1, xpad: -20,
                tickmode: 'array', tickvals: [0, 50, 100], ticktext: ['0', '50', '100'],
                tickfont: { family: 'Segoe UI', size: 20, color: 'black', weight: 'bold' }
            },
            line: { color: 'black', width: 1 }, cmin: 0, cmax: 100
        }
    });

    Plotly.newPlot('plotlyGuessChart', guessTraces, {
        font: { family: 'Segoe UI' },
        xaxis: { title: { text: '<b>Over-8</b>', font: { family: 'Segoe UI', size: 25, color: 'black', weight: 'bold' }, pad: 5 }, tickfont: { family: 'Segoe UI', size: 20, color: 'black', weight: 'bold' }, showgrid: true, tickformat: '.1f', dtick: 0.5, fixedrange: false },
        yaxis: { title: { text: '<b>Vintage</b>', font: { family: 'Segoe UI', size: 25, color: 'black', weight: 'bold' }, pad: 5 }, tickfont: { family: 'Segoe UI', size: 20, color: 'black', weight: 'bold' }, tickangle: -90, showgrid: true, tickformat: '.0f', dtick: Math.max(2, Math.ceil((Math.max(...scatterData.map(d => d.vintage)) - Math.min(...scatterData.map(d => d.vintage))) / 5)), fixedrange: false },
        margin: { l: 60, r: 0, t: 30, b: 55 },
        annotations: [
            { x: 0, y: 1, xref: 'paper', yref: 'paper', text: '<b>New<br>Hard</b>', showarrow: false, font: { size: 20 }, opacity: 0.75, xanchor: 'left', yanchor: 'top' },
            { x: 1, y: 1, xref: 'paper', yref: 'paper', text: '<b>New<br>Easy</b>', showarrow: false, font: { size: 20 }, opacity: 0.75, xanchor: 'right', yanchor: 'top' },
            { x: 1, y: 0, xref: 'paper', yref: 'paper', text: '<b>Old<br>Easy</b>', showarrow: false, font: { size: 20 }, opacity: 0.75, xanchor: 'right', yanchor: 'bottom' },
            { x: 0, y: 0, xref: 'paper', yref: 'paper', text: '<b>Old<br>Hard</b>', showarrow: false, font: { size: 20 }, opacity: 0.75, xanchor: 'left', yanchor: 'bottom' }
        ]
    }, {responsive: true, displayModeBar: false});
}