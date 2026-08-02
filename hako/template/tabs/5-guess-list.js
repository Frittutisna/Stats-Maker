let currentListChartMode        = "ALL"; 
let currentGuessListViewMode    = "ALL";
let isGraphFocused              = false;

if (watched) {
    const glBtn = document.getElementById("guessListToggleBtn");
    const fcBtn = document.getElementById("focusToggleBtn");

    if (glBtn) glBtn.classList.remove("hidden");
    if (fcBtn) fcBtn.classList.remove("hidden");
}

function getCrossProduct(o, a, b, xKey, yKey) {
    return (a[xKey] - o[xKey]) * (b[yKey] - o[yKey]) - (a[yKey] - o[yKey]) * (b[xKey] - o[xKey]);
}

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
    packedPts.sort((a, b) => a[xKey] === b[xKey] ? a[yKey] - b[yKey] : a[xKey] - b[xKey]);

    const lower = [];
    const upper = [];

    for (let p of packedPts) {
        while (lower.length >= 2 && getCrossProduct(lower[lower.length - 2], lower[lower.length - 1], p, xKey, yKey) <= 0) lower.pop();
        lower.push(p);
    }

    for (let i = packedPts.length - 1; i >= 0; i--) {
        let p = packedPts[i];
        while (upper.length >= 2 && getCrossProduct(upper[upper.length - 2], upper[upper.length - 1], p, xKey, yKey) <= 0) upper.pop();
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

function updateGuessHelpDropdown() {
    const dropdown = document.getElementById("guessGuideDropdown");
    if (!dropdown) return;

    let gearText = watched ? `<br><br>Click the <b class="bg-black text-white px-1 rounded">⚙</b> button to configure your view settings<br>` : '';

    let guideText = `
        <b>Display</b><br>
        Changes the dataset shown on the bubble chart<br>
        • <b>Corrects:</b> All songs guessed correctly<br>
        • <b>Rigs:</b> All songs from this player's list<br>
        • <b>Rigs Hit:</b> All songs guessed correctly from this player's list<br><br>
        <b>Focus</b><br>
        • <b>On:</b> Scales the axes to fit only the current chart<br>
        • <b>Off:</b> Uses the shared axis bounds for easy comparison between charts${gearText}
    `;
    
    let exampleText = "";

    if (currentGuessListViewMode === "ALL") {
        exampleText = `
            A <b>small, <span style="color: #3232c8;">blue</span></b> circle in the <b>bottom-left</b> means that, on average, this player:<br>
            • Has low Guess Rate (<b>small</b>), yet<br>
            • Is over-performing their Elo (<b><span style="color: #3232c8;">blue</span></b>),<br>
            • Usually hits harder (<b>left</b>) songs, and<br>
            • Prefers the older (<b>bottom</b>) ones
        `;
    }

    else if (currentGuessListViewMode === "RIG") {
        exampleText = `
            A <b>big, <span style="color: #c83232;">red</span></b> circle in the <b>top-right</b> means that, on average, this player's list:<br>
            • Usually has newer (<b>top</b>) songs,<br>
            • Appears a lot (<b>big</b>),<br>
            • Is difficult for the player (<b><span style="color: #c83232;">red</span></b>), yet<br>
            • Easy for others (<b>right</b>)
        `;
    }

    else if (currentGuessListViewMode === "HIT") {
        exampleText = `
            A <b>big, <span style="color: #c83232;">red</span></b> circle in the <b>top-right</b> means that, on average, this player:<br>
            • Focuses heavily on newer (<b>top</b>) songs from their list,<br>
            • Said list appears a lot (<b>big</b>),<br>
            • Is difficult (<b><span style="color: #c83232;">red</span></b>) for them to get right, yet<br>
            • Easy for others (<b>right</b>)
        `;
    }

    dropdown.innerHTML = `
        <p class="font-bold pb-1 mb-1">Guide</p>
        <hr class="border-black mb-2">
        <p class="mb-2 text-xs">${guideText}</p>
        <p class="font-bold pb-1 mb-1 mt-3">Example</p>
        <hr class="border-black mb-2">
        <p class="text-xs font-normal">${exampleText}</p>
    `;
}

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

function renderGuessScatterChart() {
    const targetNode = document.getElementById('plotlyGuessChart');
    if (!targetNode || !scatterData) return;

    const guessHull = get75PercentileHull(scatterData, 'over8', 'vintage');
    let guessTraces = [];

    if (guessHull) {
        guessTraces.push({
            x           : guessHull.x,
            y           : guessHull.y,
            type        : 'scatter',
            mode        : 'lines',
            line        : {color: 'black', width: 0.5, dash: 'solid'},
            hoverinfo   : 'skip',
            showlegend  : false
        });
    }

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

    const guessChartDiv = document.getElementById('plotlyGuessChart');

    if (guessChartDiv) {
        guessChartDiv.on('plotly_click', function(data) {
            if (!data.points || data.points.length === 0) return;
            const pointData = data.points[0];

            if (pointData.customdata && pointData.customdata[0]) {
                const playerName = String(pointData.customdata[0]).trim().toLowerCase();
                if (typeof window.searchPlayerMetricFromTable === 'function') window.searchPlayerMetricFromTable(`correct:${playerName}`);
            }
        });
    }
}

if (document.getElementById('plotlyListChart') && typeof arrowData !== 'undefined' && arrowData) {
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
            x           : d.x_hit, 
            y           : d.y_hit,
            size        : d.rig_rate, 
            color       : d.grid_grs || d.rig_gr, 
            hoverText   : `<b>${d.name}</b><br>Hit Rig Over-8: ${(d.x_hit !== undefined ? d.x_hit : d.x_end).toFixed(2)}<br>Hit Rig Vintage: ${d.seasonal_vintage || d.seasonal_vintage_end}<br>Rig Rate: ${(d.grid_rate !== undefined ? d.grid_rate : d.rig_rate).toFixed(2)}<br>Rig GR: ${d.rig_gr.toFixed(2)}<extra></extra>`
        }))
    };
}

function renderListChart() {
    if (!window.listDataPool || !window.listDataPool[currentListChartMode]) return;

    const activeScatterSource   = window.listDataPool[currentListChartMode];
    const listHull              = get75PercentileHull(activeScatterSource, 'x', 'y');
    let listTraces              = [];

    if (listHull) {
        listTraces.push({
            x           : listHull.x,
            y           : listHull.y,
            type        : 'scatter',
            mode        : 'lines',
            line        : {color: 'black', width: 0.5, dash: 'solid'},
            hoverinfo   : 'skip',
            showlegend  : false
        });
    }

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

                if (typeof window.searchPlayerMetricFromTable === 'function') {
                    if (currentListChartMode === "HIT") window.searchPlayerMetricFromTable(`list:${playerName} correct:${playerName}`);
                    else                                window.searchPlayerMetricFromTable(`list:${playerName}`);
                }
            }
        });
    }
}

initGuessDropdownListeners  ();
updateGuessHelpDropdown     ();
renderGuessScatterChart     ();

if (window.listDataPool) renderListChart();