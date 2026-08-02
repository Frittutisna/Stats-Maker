function syncTierDropdownDOMState() {
    const isCount = globalChartMode === "COUNT";

    const optBase   = document.getElementById("opt_c1_base");
    const optOver8  = document.getElementById("opt_c1_over8");
    const optRig    = document.getElementById("opt_c1_rig");
    const optChant  = document.getElementById("opt_c1_chant");
    const groupC1   = document.getElementById("label_group_c1");
    const groupC2   = document.getElementById("label_group_c2");

    if (optBase)    optBase     .innerText = isCount ? "Corrects"       : "Guess Rate";
    if (optOver8)   optOver8    .innerText = isCount ? "Over-8 Hit"     : "Over-8 Distribution";
    if (optRig)     optRig      .innerText = isCount ? "Rigs"           : "Rig Rate";
    if (optChant)   optChant    .innerText = isCount ? "Chanting Hit"   : "Chanting Guess Rate";
    if (groupC1)    groupC1     .innerText = "General";
    if (groupC2)    groupC2     .innerText = "Contribution";

    const hitLabel = document.getElementById("label_opt_c1_hit");
    const offLabel = document.getElementById("label_opt_c1_off");
    const rigLabel = document.getElementById("opt_c1_rig")?.closest('label');

    if (!watched) {
        if (hitLabel) hitLabel.classList.add("hidden");
        if (offLabel) offLabel.classList.add("hidden");
        if (rigLabel) rigLabel.classList.add("hidden");

        if (c1Sub === "RIG" || c1Sub === "HIT" || c1Sub === "OFF") {
            c1Sub = "BASE";

            const defaultBaseRadio = document.querySelector('input[name="tierSubMetricsRadio"][value="BASE"]');
            if (defaultBaseRadio) defaultBaseRadio.checked = true;
        }
    }

    else if (isCount) {
        if (hitLabel) hitLabel.classList.add("hidden");
        if (offLabel) offLabel.classList.add("hidden");

        if (c1Sub === "HIT" || c1Sub === "OFF") {
            c1Sub = "BASE";

            const defaultBaseRadio = document.querySelector('input[name="tierSubMetricsRadio"][value="BASE"]');
            if (defaultBaseRadio) defaultBaseRadio.checked = true;
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

            syncTierDropdownDOMState    ();
            renderTierCharts            ();
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

            const c1GroupRadio = document.querySelector('input[name="tierChartGroupRadio"][value="C1"]');
            if (c1GroupRadio) c1GroupRadio.checked = true;

            document.querySelectorAll('input[name="tierTimeRadio"]').forEach(r => r.checked = false);
            renderTierCharts();
        });
    });

    document.querySelectorAll('input[name="tierTimeRadio"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            c3Mode = e.target.value;

            const c3GroupRadio = document.querySelector('input[name="tierChartGroupRadio"][value="C3"]');
            if (c3GroupRadio) c3GroupRadio.checked = true;

            document.querySelectorAll('input[name="tierSubMetricsRadio"]').forEach(r => r.checked = false);
            renderTierCharts();
        });
    });
}

function updateTierHelpDropdown() {
    const dropdown = document.getElementById("tierGuideDropdown");
    if (!dropdown) return;

    dropdown.innerHTML = `
        <p class="font-bold border-b pb-1 mb-1">Guide</p>
        <p class="mb-1 pt-1 text-xs">
            <b class="font-bold">Sort</b><br>
            • <b>Global:</b> Sorts and ranks all players globally<br>
            • <b>Tier:</b> Sorts and ranks players within their respective Tiers<br><br>
            <b class="font-bold">Display</b><br>
            • <b>Rate:</b> Shows percentages<br>
            • <b>Count:</b> Shows raw counts<br><br>
            <b class="font-bold">Chart</b><br>
            • <b>General:</b> Charts various metrics (Correct, Over-8, (Off) Rig, Chanting)<br>
            • <b>Contribution:</b> Charts direct contribution to life count (Lives Taken/Saved, Other Correct)<br>
            • <b>Time:</b> Charts Time aggregate metrics (Minimum, Mean, Median, Maximum, or Standard Deviation)<br><br>
            Click the <b class="bg-black text-white px-1 rounded">⚙</b> button to configure your view settings<br>
            Only one chart is visible at any time
        </p>
    `;
}

function renderTierCharts() {
    if ((!document.getElementById('tierChart_MainMetrics') && !document.getElementById('tierChart_MainMetricsMain')) || !tierStats) return;
    if (!globalSearchData || globalSearchData.length === 0)                                                                         return;

    let checkedGroup = document.querySelector('input[name="tierChartGroupRadio"]:checked')?.value;

    if (!checkedGroup) {
        if (document.querySelector('input[name="tierSubMetricsRadio"]:checked')) {
            checkedGroup = "C1";

            const c1Radio = document.querySelector('input[name="tierChartGroupRadio"][value="C1"]');
            if (c1Radio) c1Radio.checked = true;
        }

        else if (document.querySelector('input[name="tierTimeRadio"]:checked')) {
            checkedGroup = "C3";

            const c3Radio = document.querySelector('input[name="tierChartGroupRadio"][value="C3"]');
            if (c3Radio) c3Radio.checked = true;
        }

        else {
            checkedGroup = "C1";

            const c1Radio       = document.querySelector('input[name="tierChartGroupRadio"][value="C1"]');
            const baseSubRadio  = document.querySelector('input[name="tierSubMetricsRadio"][value="BASE"]');

            if (c1Radio)        c1Radio.checked         = true;
            if (baseSubRadio)   baseSubRadio.checked    = true;

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
        if (!str)                   return "";
        if (str.length <= maxLen)   return str;

        let cutStr      = str.slice(0, maxLen);
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

        let x8Lists = Array.from({length: 8}, () => []);

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

            if (isCorrect && !taken.includes(pClean) && !saved.includes(pClean)) {
                otherCorrectsList.push(bulletLabel);
            }
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
        const outputSample  = [...ticks, ...crosses];

        return `<b>${fractionStr}</b><br>` + outputSample.join('<br>');
    };

    let absoluteMaxCorrectsC1 = 0;
    let absoluteMaxCorrectsC2 = 0;

    ["1", "2", "3", "4"].forEach((tr) => {
        if (!tierStats[tr]) return;

        tierStats[tr].forEach(p => {
            let pNameStr    = getPlayerStringName(p.Player);
            const stats     = compilePlayerStatsFromSearch(pNameStr);

            if (!pNameStr) return;

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
                if (watched) {
                    val                 = [stats.rigHits, stats.offListHits];
                    traceData           = val;
                    let rigsSection     = `<b>Rigs Hit Context:</b><br>${formatSampleTextList(stats.rigHitsList)}`;
                    let offRigsSection  = `<b>Off Rigs Hit Context:</b><br>${formatSampleTextList(stats.offListHitsList)}`;
                    hover               = `${rigsSection}<br><br>${offRigsSection}`;
                }

                else {
                    val     = stats.totalCorrect;
                    hover   = "";
                }
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

            if (mode === "RATE") {
                let sortedOffHits   = formatAndSortSongsList([...stats.offListHitsList], false);
                hover               = `<b>${stats.offListHits}/${totalOffSongs}</b><br>${sortedOffHits.join('<br>')}`;
            }

            else {
                let offListSeenSongs    = stats.allSeenSongs.filter(s => !stats.rigsList.map(r => r.slice(2)).includes(s.slice(2)));
                hover                   = formatFractionalSample(`${stats.offListHits}/${totalOffSongs}`, offListSeenSongs);
            }
        }

        else if (sub === "CHANT") {
            if (!hasChanting) return {val: 0, hover: "No Chant Songs Exist"};
            val = mode === "RATE" ? (stats.totalChantSeen > 0 ? (stats.totalChantCorrect / stats.totalChantSeen) * 100 : 0) : stats.totalChantCorrect;

            if (mode === "RATE")    hover = formatFractionalSample(`${stats.totalChantCorrect}/${stats.totalChantSeen}`, stats.chantSeenList);
            else                    hover = formatSampleTextList(stats.chantCorrectsList);
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
                yLabels.push(" ".repeat(gapCounter++));
                customHovers.push("");
                rawItems.push(null);

                if (isSubDistribution)  multiData.forEach(arr => arr.push(null));
                else                    singleXVals.push(null);

                return;
            }

            yLabels.push(getPlayerStringName(p.Player));
            let extracted = valExtractionFn(p);
            customHovers.push(extracted.hover);
            rawItems.push(p);

            if (isSubDistribution)  for (let i = 0; i < multiData.length; i++) multiData[i].push(extracted.traceData ? extracted.traceData[i] : 0);
            else                    singleXVals.push(extracted.val);
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

        if (globalChartMode === "COUNT" && c1Sub === "BASE" && watched) {
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

        if (Array.isArray(va)) va = va.reduce((x, y) => x + y, 0);
        if (Array.isArray(vb)) vb = vb.reduce((x, y) => x + y, 0);

        return vb - va;
    };

    let c1LayoutSubMode = (globalChartMode === "COUNT" && c1Sub === "BASE" && watched) ? "COUNT_BASE" : (c1Sub === "OVER-8" ? "X8" : null);
    let c1Data          = buildChartData(currentTierChartMode, c1Sort, (p) => getC1ValueAndHover(p, globalChartMode, c1Sub), c1LayoutSubMode);
    let c1Traces        = [];

    if (globalChartMode === "COUNT" && c1Sub === "BASE" && watched) {
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
                x                   : c1Data.multiData[i].slice().reverse(),
                y                   : c1Data.yLabels.slice().reverse(),
                type                : 'bar',
                orientation         : 'h',
                barmode             : 'stack',
                name                : baseNames[i],
                marker              : {color: c1BaseColors[i]},
                hovertext           : (i === 0 ? trace0Hovers : trace1Hovers).slice().reverse(),
                hoverinfo           : 'text',
                text                : c1Data.multiData[i].slice().reverse().map(v => v ? v.toFixed(0) : ""),
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
                x                   : c1Data.multiData[i].slice().reverse(),
                y                   : c1Data.yLabels.slice().reverse(),
                type                : 'bar',
                orientation         : 'h',
                barmode             : 'stack',
                name                : `${i+1}/8`,
                marker              : {color: c8Colors[i]},
                hovertext           : x8TraceHovers[i].slice().reverse(),
                hoverinfo           : 'text',
                text                : c1Data.multiData[i].slice().reverse().map(v => v ? v.toFixed(globalChartMode === "RATE" ? 1 : 0) : ""),
                textposition        : 'inside',
                insidetextanchor    : 'middle',
                textangle           : 0,
                textfont            : {family: 'Segoe UI', size: 15, color: 'white', weight: 'bold'}
            });
        }
    }

    else {
        c1Traces.push({
            x                   : c1Data.singleXVals.slice().reverse(),
            y                   : c1Data.yLabels.slice().reverse(),
            type                : 'bar',
            orientation         : 'h',
            barmode             : 'group',
            hovertext           : c1Data.customHovers.slice().reverse(),
            hoverinfo           : 'text',
            text                : c1Data.singleXVals.slice().reverse().map(v => v === null ? "" : v.toFixed(globalChartMode === "RATE" ? 2 : 0) + " "),
            textposition        : 'inside',
            insidetextanchor    : 'end',
            textangle           : 0,
            textfont            : {family: 'Segoe UI', size: 15, color: 'white', weight: 'bold'},
            marker              : {
                color       : c1Data.singleXVals.slice().reverse(),
                colorscale  : (() => {
                    if (globalChartMode === "RATE" && c1Sub === "HIT") return [[0, c0], [0.70, c0], [0.80, c1], [0.90, c2], [1, c2]];
                    return [[0, c0], [1, c2]];
                })(),
                cmin        : 0,
                cmax        : Math.max(...c1Data.singleXVals.filter(v => v !== null)) || 1
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
        for (let i = 0; i < c1Traces.length; i++) if (currentChart1Div.data[i].visible === 'legendonly') c1Traces[i].visible = 'legendonly';
    }

    let rangeC1;

    if (globalChartMode === "RATE") rangeC1 = [0, 105];
    else                            rangeC1 = (c1Sub === "BASE" || c1Sub === "OVER-8") ? [0, absoluteMaxCorrectsC1 + 1] : null;

    const titleC1       = `<span style="font-size: 30px;"><b>${displayTitleC1}</b></span>`;
    const hasLegends    = (c1Sub === "OVER-8" || (globalChartMode === "COUNT" && c1Sub === "BASE"));

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

            if (query && typeof window.searchPlayerMetricFromTable === 'function') window.searchPlayerMetricFromTable(query);
        });

        newChart1Div.addEventListener('contextmenu', e => e.preventDefault());

        newChart1Div.on('plotly_hover', function(data) {
            if (!data.points || data.points.length === 0 || (globalChartMode === "RATE" && c1Sub === "BASE")) return;

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
                                r = rgb0[0] + norm * (rgb2[0] - rgb0[0]);
                                g = rgb0[1] + norm * (rgb2[1] - rgb0[1]);
                                b = rgb0[2] + norm * (rgb2[2] - rgb0[2]);
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
            if (!p) {tkHovers.push(""); svHovers.push(""); return;}
            let s = compilePlayerStatsFromSearch(getPlayerStringName(p.Player));

            tkHovers.push(formatSampleTextList(s.livesTakenList));
            svHovers.push(formatSampleTextList(s.livesSavedList));
        });

        for (let i = 0; i < 2; i++) {
            c2Traces.push({
                x                   : c2Data.multiData[i].slice().reverse(),
                y                   : c2Data.yLabels.slice().reverse(),
                type                : 'bar',
                orientation         : 'h',
                barmode             : 'stack',
                name                : names[i],
                marker              : {color: c2Colors[i]},
                hovertext           : (i === 0 ? tkHovers : svHovers).slice().reverse(),
                hoverinfo           : 'text',
                text                : c2Data.multiData[i].slice().reverse().map(v => v ? v.toFixed(0) : ""),
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
            if (!p) {tkHovers.push(""); othHovers.push(""); svHovers.push(""); return}
            let s = compilePlayerStatsFromSearch(getPlayerStringName(p.Player));

            tkHovers    .push(formatSampleTextList(s.livesTakenList));
            othHovers   .push(formatSampleTextList(s.otherCorrectsList));
            svHovers    .push(formatSampleTextList(s.livesSavedList));
        });

        for (let i = 0; i < 3; i++) {
            c2Traces.push({
                x                   : c2Data.multiData[i].slice().reverse(),
                y                   : c2Data.yLabels.slice().reverse(),
                type                : 'bar',
                orientation         : 'h',
                barmode             : 'stack',
                name                : names[i],
                marker              : {color: c3Colors[i]},
                hovertext           : (i === 0 ? tkHovers : (i === 1 ? othHovers : svHovers)).slice().reverse(),
                hoverinfo           : 'text',
                text                : c2Data.multiData[i].slice().reverse().map(v => v ? v.toFixed(1) : ""),
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

            if (query && typeof window.searchPlayerMetricFromTable === 'function') window.searchPlayerMetricFromTable(query);
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
        x                   : c3Data.singleXVals.slice().reverse(),
        y                   : c3Data.yLabels.slice().reverse(),
        type                : 'bar',
        orientation         : 'h',
        hovertext           : c3Data.customHovers.slice().reverse(),
        hoverinfo           : 'none',
        text                : c3Data.singleXVals.slice().reverse().map(v => v === null ? "" : v.toFixed(2) + " "),
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
        xaxis       : {tickfont: {size: 15, color: 'black', weight: 'bold'}, fixedrange: true, showgrid: true,  range: [0, 21], tickmode: 'array', tickvals: [0, 4, 8, 12, 16, 20]},
        yaxis       : {tickfont: {size: 15, color: 'black', weight: 'bold'}, fixedrange: true, showgrid: false, ticksuffix: " "},
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

initTierDropdownListeners   ();
syncTierDropdownDOMState    ();
updateTierHelpDropdown      ();
renderTierCharts            ();