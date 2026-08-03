const xLabels       = (numX === 8) ? ['5',  '10', '15', '20', '25', '30', '35'] : ['5',     '10', '15', '20', '25', '30', '35', '40'];
const yLabels       = (numY === 8) ? [1995, 2000, 2005, 2010, 2015, 2020, 2025] : [1990,    1995, 2000, 2005, 2010, 2015, 2020, 2025];
const matrixBins    = {};

if (typeof songData !== 'undefined' && songData) {
    songData.forEach(s => {
        let xIdx    = Math.min(Math.floor(s.difficulty / 5), numX - 1);
        let yIdx    = (numY === 8) ? ((s.vintage < 1995) ? 0 : Math.min(Math.floor((s.vintage - 1995) / 5) + 1, 7)) : Math.min(Math.max(Math.floor((s.vintage - 1985) / 5), 0), 8);
        let key     = `${xIdx}-${yIdx}`;

        if (!matrixBins[key]) matrixBins[key] = {count: 0, over8Sum: 0};

        matrixBins[key].count++;
        matrixBins[key].over8Sum += s.correct_count;
    });
}

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
                let startDf = j * 5;
                diffStr     = `Difficulty: ${startDf}-${startDf + 5}`;
            }
        }

        else {
            if      (j === 0)   diffStr = "Difficulty: <5";
            else if (j === 8)   diffStr = "Difficulty: >40";
            else {
                let startDf = j * 5;
                diffStr     = `Difficulty: ${startDf}-${startDf + 5}`;
            }
        }

        if (key in matrixBins) {
            let val = matrixBins[key].over8Sum / matrixBins[key].count;
            rowZ.push(val);

            let bin_songs = matrixSongs && matrixSongs[key] ? [...matrixSongs[key]] : [];

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
            rowZ.push(null);
            rowText.push(`<b>${vintageStr}<br>${diffStr}<br>Mean Over-8: N/A</b>`);
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

function updateSongHelpDropdown() {
    const dropdown = document.getElementById("songGuideDropdown");
    if (!dropdown) return;

    dropdown.innerHTML = `
        <p class="font-bold pb-1 mb-1 text-sm text-black">Example</p>
        <hr class="border-black mb-2">
        <p class="text-xs font-normal">
            <b>Difficulty:</b> 25-30<br>
            <b>Vintage:</b> 2015-2020<br>
            <b>Over-8:</b> 6.26<br>
            This means that, on average, <b>6.26/8</b> people guessed songs from <b>2015-2020</b> with a difficulty of <b>25-30</b> correctly
        </p>
    `;
}

function renderSongHeatmap() {
    const targetNode = document.getElementById('plotlySongChart');
    if (!targetNode) return;

    const isDark    = window.isDarkMode;
    const textColor = isDark ? '#ffffff' : '#323232';
    const paperBg   = isDark ? '#323232' : 'rgba(0, 0, 0, 0)';
    const plotBg    = isDark ? '#323232' : 'rgba(0, 0, 0, 0)';
    const gridColor = isDark ? '#3c3c3c' : '#f0f0f0';

    Plotly.newPlot('plotlySongChart', [{
        z           : zValues,
        x           : Array.from({length: numX}, (_, i) => i),
        y           : Array.from({length: numY}, (_, i) => i),
        text        : textLabels,
        hoverinfo   : 'none',
        type        : 'heatmap',
        colorscale  : [[0, c0], [0.375, c1], [0.625, c2], [1, c2]],
        zmin        : 0,
        zmax        : 8,
        showscale   : true,
        colorbar    : {
            title       : {text: '<b>Over-8</b>', font: {family: 'Segoe UI', size: 25, color: textColor, weight: 'bold'}, side: 'right'},
            thickness   : 25,
            len         : 1.0,
            y           : 0.5,
            yanchor     : 'middle',
            x           : 1,
            xpad        : -10,
            tickmode    : 'array',
            tickvals    : [0, 3, 5, 8],
            ticktext    : ['0', '3', '5', '8'],
            tickfont    : {family: 'Segoe UI', size: 20, color: textColor, weight: 'bold'}
        }
    }], {
        font            : {family: 'Segoe UI', size: 50, color: textColor},
        paper_bgcolor   : paperBg,
        plot_bgcolor    : plotBg,
        xaxis           : {
            title           : {text: '<b>Difficulty</b>', font: {family: 'Segoe UI', size: 25, color: textColor, weight: 'bold'}, pad: 5},
            tickmode        : 'array',
            tickvals        : Array.from({length: numX - 1}, (_, i) => i + 0.5),
            ticktext        : xLabels,
            tickfont        : {family: 'Segoe UI', size: 20, color: textColor, weight: 'bold'},
            gridcolor       : gridColor,
            showgrid        : true,
            zeroline        : false,
            showticklabels  : true,
            ticks           : 'outside',
            ticklen         : 5,
            tickcolor       : 'rgba(0, 0, 0, 0)',
            fixedrange      : true
        },
        yaxis: {
            title           : {text: '<b>Vintage</b>', font: {family: 'Segoe UI', size: 25, color: textColor, weight: 'bold'}, pad: 5},
            tickmode        : 'array',
            tickvals        : Array.from({length: numY - 1}, (_, i) => i + 0.5),
            ticktext        : yLabels,
            tickfont        : {family: 'Segoe UI', size: 20, color: textColor, weight: 'bold'},
            tickangle       : -90,
            gridcolor       : gridColor,
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

                let bin_songs = [...matrixSongs[key]];
                if (typeof window.translateHoverText === 'function') bin_songs = window.translateHoverText(bin_songs);

                bin_songs = bin_songs
                    .sort((a, b) => {
                        const cleanA = (a.startsWith('✓') || a.startsWith('✗')) ? a.slice(2) : a;
                        const cleanB = (b.startsWith('✓') || b.startsWith('✗')) ? b.slice(2) : b;
                        return cleanA.toLowerCase().localeCompare(cleanB.toLowerCase());
                    })
                    .map(s => (s.startsWith('✓') || s.startsWith('✗')) ? s : `• ${s}`);

                const baseInfo          = textLabels    [i][j]          || "";
                const currentCellColor  = bgColors      [i][j]          || 'black';
                const isWhite           = currentCellColor === 'white'  || currentCellColor === '#ffffff' || currentCellColor === 'rgb(255,255,255)' || currentCellColor === 'rgb(255, 255, 255)';

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

            if (queryParts.length > 0 && typeof window.searchPlayerMetricFromTable === 'function') window.searchPlayerMetricFromTable(queryParts.join(" "));
        });
    }
}

updateSongHelpDropdown  ();
renderSongHeatmap       ();