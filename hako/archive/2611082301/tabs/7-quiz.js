let quizActivePool          = [];
let quizIndex               = 0;
let quizScore               = 0;
let quizTimerId             = null;
let quizIsPaused            = false;
let quizLeewayActive        = false;
let quizTimeRemaining       = 0;
let quizSoundLimit          = 20;
let quizNoSoundLimit        = 0;
let quizCurrentAudio        = null;
let quizAudioTimeoutId      = null;
let quizReplayActive        = false;
let quizLeewayTimeoutId     = null;
let quizRevealTimeElapsed   = 0;
let quizMatrixStates        = {seen: {}, correct: {}, list: {}};
let quizCurrentMode         = "entry";
let quizSampleMode          = "actual";

const quizConfig = [
    {id: "type",        name: "Song Type",  type: "categorical",    subOptions: ["Opening", "Ending", "Insert"]},
    {id: "chanting",    name: "Chanting",   type: "categorical",    subOptions: ["Yes", "No"]},
    {id: "anime_type",  name: "Anime Type", type: "categorical",    subOptions: ["TV", "Movie", "OVA", "ONA", "Special"]},
    {id: "vintage",     name: "Vintage",    type: "range",          min: 1900,  max: 2026},
    {id: "difficulty",  name: "Difficulty", type: "range",          min: 0,     max: 100},
    {id: "guessers",    name: "Correct",    type: "range",          min: 0,     max: 8},
    {id: "listers",     name: "List",       type: "range",          min: 0,     max: 8}
];

function getNextQuizCycleState(curr) {
    if (curr === " ") return "+";
    if (curr === "+") return "~";
    if (curr === "~") return "-";

    return " ";
}

function initQuizSettingsCheckboxes() {
    const container = document.getElementById("quizColumnCheckboxContainer");
    const masterChk = document.getElementById("quizColumnsMasterCheckbox");

    if (!container || !masterChk) return;
    container.innerHTML = "";

    const searchVints = globalSearchData.map(s => s._vintageParsed).filter(v => !isNaN(v) && v !== -Infinity);
    const searchDiffs = globalSearchData.map(s => s._diffParsed).filter(d => !isNaN(d) && d !== -Infinity);

    const vintC = quizConfig.find(c => c.id === "vintage");
    const diffC = quizConfig.find(c => c.id === "difficulty");

    if (vintC && searchVints.length > 0) {vintC.min = Math.floor(Math.min(...searchVints)); vintC.max = Math.ceil(Math.max(...searchVints));}
    if (diffC && searchDiffs.length > 0) {diffC.min = 0; diffC.max = Math.ceil(Math.max(...searchDiffs) / 5) * 5;}

    function updateMasterState() {
        const allChecked    = quizConfig.every(c => c.visible);
        const noneChecked   = quizConfig.every(c => !c.visible);

        masterChk.checked       = allChecked;
        masterChk.className     = "rounded accent-black";
        masterChk.indeterminate = !allChecked && !noneChecked;
    }

    masterChk.addEventListener("change", () => {
        const isChecked = masterChk.checked;

        quizConfig.forEach(c => {
            c.visible = isChecked;

            if (c.type === "categorical" && c.subOptions) {
                if (isChecked) c.selectedOptions = new Set(c.subOptions.map(o => o.toLowerCase()));
                else c.selectedOptions.clear();
            }

            else if (!isChecked && c.type === "range") {
                c.currentMin = c.min;
                c.currentMax = c.max;
            }
        });

        const allInputs = container.getElementsByTagName('input');

        for (let i = 0; i < allInputs.length; i++) {
            if (allInputs[i].type === 'checkbox') {
                allInputs[i].checked        = isChecked;
                allInputs[i].indeterminate  = false;
            }
        }
    });

    quizConfig.forEach(col => {
        col.visible             = true;
        const colWrapper        = document.createElement("div");
        colWrapper.className    = "flex flex-col space-y-1";

        const label     = document.createElement("label");
        label.className = "flex items-center gap-2 cursor-pointer w-full text-left font-bold text-black";

        const chk       = document.createElement("input");
        chk.type        = "checkbox";
        chk.className   = "rounded accent-black";
        chk.checked     = true;

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
                if (inputs.length === 2) { inputs[0].value = col.min; inputs[1].value = col.max; }
            }

            updateMasterState();
        });

        label       .appendChild(chk);
        label       .appendChild(document.createTextNode(col.name));
        colWrapper  .appendChild(label);

        if (col.type === "categorical" && col.subOptions) {
            const subContainer      = document.createElement("div");
            subContainer.className  = "pl-6 flex flex-col text-xs";
            col.selectedOptions     = new Set(col.subOptions.map(o => o.toLowerCase()));

            col.subOptions.forEach(opt => {
                const subLabel      = document.createElement("label");
                subLabel.className  = "flex items-center gap-1 cursor-pointer text-gray-700 hover:text-black py-0.5 pr-6";
                const subChk        = document.createElement("input");
                subChk.type         = "checkbox";
                subChk.className    = "rounded accent-black";
                subChk.checked      = true;

                subChk.addEventListener("change", () => {
                    if (subChk.checked) col.selectedOptions.add     (opt.toLowerCase());
                    else                col.selectedOptions.delete  (opt.toLowerCase());

                    const total     = col.subOptions.length;
                    const selected  = col.selectedOptions.size;

                    if      (selected === total)    {chk.checked = true;    chk.indeterminate = false;  col.visible = true;}
                    else if (selected === 0)        {chk.checked = false;   chk.indeterminate = false;  col.visible = false;}
                    else                            {chk.checked = false;   chk.indeterminate = true;   col.visible = true;}

                    updateMasterState();
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
            inputMin.type       = "number"; inputMin.value = col.min;
            inputMin.className  = "w-10 h-5 border text-center text-xs";

            const maxLabel      = document.createElement("label");
            maxLabel.className  = "flex items-center justify-start gap-2 font-mono";
            maxLabel.innerHTML  = "Max:";
            const inputMax      = document.createElement("input");
            inputMax.type       = "number"; inputMax.value = col.max;
            inputMax.className  = "w-10 h-5 border text-center text-xs";

            const syncRanges = () => {
                let vMin = parseFloat(inputMin.value);
                let vMax = parseFloat(inputMax.value);

                col.currentMin = isNaN(vMin) ? col.min : vMin;
                col.currentMax = isNaN(vMax) ? col.max : vMax;

                if (!col.visible && (col.currentMin > col.min || col.currentMax < col.max)) {
                    col.visible = true;
                    chk.checked = true;

                    updateMasterState();
                }
            };

            inputMin.addEventListener("input", syncRanges);
            inputMax.addEventListener("input", syncRanges);

            minLabel        .appendChild(inputMin);
            maxLabel        .appendChild(inputMax);
            boxWrapper      .appendChild(minLabel);
            boxWrapper      .appendChild(maxLabel);
            inputContainer  .appendChild(boxWrapper);
            colWrapper      .appendChild(inputContainer);
        }

        container.appendChild(colWrapper);
    });

    const separator     = document.createElement("div");
    separator.className = "border-b my-1.5";

    container.appendChild(separator);

    const matrices = [{key: "seen", name: "Seen"}, {key: "correct", name: "Correct"}, {key: "list", name: "List"}];
    if (quizMatrixStates._sectionEnabled === undefined) {quizMatrixStates._sectionEnabled = {seen: true, correct: true, list: true};}

    matrices.forEach(m => {
        const groupWrapper      = document.createElement("div");
        groupWrapper.className  = "flex flex-col select-none";

        const headerLabel       = document.createElement("label");
        headerLabel.className   = "flex items-center gap-1.5 font-bold text-black block text-xs cursor-pointer mt-0.5";

        const sectionToggle     = document.createElement("input");
        sectionToggle.type      = "checkbox";
        sectionToggle.className = "rounded accent-black";
        sectionToggle.checked   = quizMatrixStates._sectionEnabled[m.key];

        headerLabel     .appendChild(sectionToggle);
        headerLabel     .appendChild(document.createTextNode(m.name));
        groupWrapper    .appendChild(headerLabel);

        const subContainer      = document.createElement("div");
        subContainer.className  = "pl-4 flex flex-col gap-0.5 text-xs w-full";

        const updateSubContainerOpacity = () => {
            subContainer.style.opacity          = quizMatrixStates._sectionEnabled[m.key] ? "1"     : "0.5";
            subContainer.style.pointerEvents    = quizMatrixStates._sectionEnabled[m.key] ? "auto"  : "none";
        };

        sectionToggle.addEventListener("change", () => {
            quizMatrixStates._sectionEnabled[m.key] = sectionToggle.checked;
            updateSubContainerOpacity();
        });

        if (typeof players !== 'undefined' && players) {
            players.forEach(pRow => {
                const rawPlayer = pRow["Player"];
                const pName     = typeof rawPlayer === 'object' ? String(rawPlayer.count || "") : String(rawPlayer);

                if (quizMatrixStates[m.key][pName] === undefined) quizMatrixStates[m.key][pName] = " ";

                const subLabel              = document.createElement("label");
                subLabel.className          = "flex items-center gap-1 cursor-pointer text-gray-700 hover:text-black py-0.5 pr-6 w-full min-w-max";
                const indicator             = document.createElement("span");
                const isBlank               = quizMatrixStates[m.key][pName] === " ";
                indicator.className         = `font-mono font-bold text-center border inline-flex items-center justify-center w-3.5 h-3.5 bg-white text-black border-gray-400 align-middle select-none shrink-0 p-0 text-sm self-center scale-90 ${isBlank ? "bg-white text-black border-gray-400" : "bg-black text-white border-black"}`;
                indicator.innerText         = quizMatrixStates[m.key][pName];
                indicator.style.lineHeight  = "1";

                subLabel.appendChild(indicator);
                subLabel.appendChild(document.createTextNode(pName));

                subLabel.addEventListener("click", (e) => {
                    e.preventDefault    ();
                    e.stopPropagation   ();

                    const currentState          = quizMatrixStates[m.key][pName];
                    const currentActiveCount    = Object.keys(quizMatrixStates[m.key]).filter(k => quizMatrixStates[m.key][k] !== " ").length;

                    if (currentState === " " && currentActiveCount >= 2) {
                        alert(`You can only configure ${m.name} filters for up to 2 players`);
                        return;
                    }

                    const nextState                 = getNextQuizCycleState(currentState);
                    quizMatrixStates[m.key][pName]  = nextState;
                    indicator.innerText             = nextState;

                    if (nextState === " ")  indicator.className = "font-mono font-bold text-center border inline-flex items-center justify-center w-3.5 h-3.5 bg-white text-black border-gray-400 align-middle select-none shrink-0 p-0 text-sm self-center scale-90";
                    else                    indicator.className = "font-mono font-bold text-center border inline-flex items-center justify-center w-3.5 h-3.5 bg-black text-white border-black align-middle select-none shrink-0 p-0 text-sm font-bold self-center scale-90";
                });

                subContainer.appendChild(subLabel);
            });
        }

        updateSubContainerOpacity();

        groupWrapper    .appendChild(subContainer);
        container       .appendChild(groupWrapper);
    });

    updateMasterState();
}

function updateQuizHelpDropdown() {
    const dropdown = document.getElementById("quizGuideDropdown");
    if (!dropdown) return;

    let guideText = `
        Click the <b class="bg-black text-white px-1 rounded font-normal inline-block w-5 h-4 text-center">⚙</b> 
        button to configure your quiz song pool<br>
    `;

    let filterText = `
        Configure <b>Seen</b>, <b>Correct</b>, and <b>List</b> player filters as follows<br>
        • <b class="bg-white text-black border border-gray-400 px-1 rounded inline-block w-4 h-4 text-center">&nbsp;</b>&nbsp;Ignore this player<br>
        • <b class="bg-black text-white px-1 rounded inline-block w-4 h-4 text-center">+</b>&nbsp;Force this player<br>
        • <b class="bg-black text-white px-1 rounded inline-block w-4 h-4 text-center">~</b>&nbsp;Find at least one player that match<br>
        • <b class="bg-black text-white px-1 rounded inline-block w-4 h-4 text-center">-</b>&nbsp;Force against this player<br>
    `;

    let exampleText = `
        Setting <b>HakoHoka</b> to <b class="bg-black text-white px-1 rounded inline-block w-4 h-4 text-center">+</b> 
        and <b>florenz</b> to <b class="bg-black text-white px-1 rounded inline-block w-4 h-4 text-center">-</b> under <b>Correct</b> means:<br>
        • The quiz pool will only draw songs that <b>HakoHoka</b> guessed correctly, and<br>
        • Completely skip any songs that <b>florenz</b> got right<br>
        You may only set filters for up to two players within each category
    `;

    dropdown.innerHTML = `
        <p class="font-bold pb-1 mb-1">Guide</p>
        <hr class="border-black mb-2">
        <p class="mb-2 text-xs">${guideText}</p>
        <p class="font-bold pb-1 mb-1 mt-3">Filter</p>
        <hr class="border-black mb-2">
        <p class="mb-2 text-xs">${filterText}</p>
        <p class="font-bold pb-1 mb-1 mt-3">Example</p>
        <hr class="border-black mb-2">
        <p class="text-xs font-normal">${exampleText}</p>
    `;
}

function parseColorToRGB(hex) {
    let c = hex.replace('#', '');
    if (c.length === 3) c = c.split('').map(x => x + x).join('');
    return [parseInt(c.substring(0, 2), 16), parseInt(c.substring(2, 4), 16), parseInt(c.substring(4, 6), 16)];
}

function updateQuizTimerUI(secondsVal) {
    const displayNode = document.getElementById("quizTimerValue");
    if (!displayNode) return;

    displayNode.innerText   = Math.ceil(secondsVal);
    displayNode.style.color = ""; 
}

function normalizeQuizString(str) {
    if (!str) return "";

    return str.toString().toLowerCase()
        .replace(/[★▲▼]/g, " ")
        .replace(/[^a-z0-9\u3040-\u309f\u30a0-\u30ff\u4e00-\u9faf\s]/g, " ")
        .replace(/\s+/g, " ").trim();
}

function enforceQuizInputBounds() {
    const countNode = document.getElementById("quizSongCountInput");
    let cVal        = parseInt(countNode.value);
    const maxSongs  = (globalSearchData && globalSearchData.length > 0) ? globalSearchData.length : 35;

    if (isNaN(cVal) || cVal < 5) countNode.value = 5;
    if (cVal > maxSongs)         countNode.value = maxSongs;

    const soundNode = document.getElementById("quizSoundTimeInput");
    let sVal        = parseInt(soundNode.value);

    if (isNaN(sVal) || sVal < 1) soundNode.value = 1;
    if (sVal > 60)               soundNode.value = 60;

    const extraNode = document.getElementById("quizExtraTimeInput");
    let eVal        = parseInt(extraNode.value);

    if (isNaN(eVal) || eVal < 0) extraNode.value = 0;
    if (eVal > 60)               extraNode.value = 60;
}

["quizSongCountInput", "quizSoundTimeInput", "quizExtraTimeInput"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener("change", enforceQuizInputBounds);
});

function evaluateMatrixConditions(song) {
    const matrices = ["seen", "correct", "list"];

    for (let mKey of matrices) {
        if (quizMatrixStates._sectionEnabled && quizMatrixStates._sectionEnabled[mKey] === false) continue;

        let hasMust         = false;
        let mustMatched     = false;
        let activeOrCount   = 0;

        for (let pName in quizMatrixStates[mKey]) if (quizMatrixStates[mKey][pName] === "~") activeOrCount++;

        let hasOr       = activeOrCount > 1;
        let orMatched   = false;

        for (let pName in quizMatrixStates[mKey]) {
            const state = quizMatrixStates[mKey][pName];
            if (state === " ") continue;

            const pClean        = pName.replace(/[★▲▼]/g, "").trim().toLowerCase();
            const roomPlayers   = (song.room_players    || []).map(p => p.toLowerCase());
            const guessers      = (song.guessers_flat   || []).map(p => p.toLowerCase());
            const listers       = (song.listers_flat    || []).map(p => p.toLowerCase());
            let isTrue          = false;

            if      (mKey === "seen")       isTrue = roomPlayers.includes(pClean);
            else if (mKey === "correct")    isTrue = guessers.includes(pClean);
            else if (mKey === "list")       isTrue = listers.includes(pClean);

            if (state === "+") {
                hasMust = true;
                if (!isTrue) return false; 
                mustMatched = true;
            }

            else if (state === "-" && isTrue)           return false;
            else if (state === "~" && isTrue && hasOr)  orMatched = true;
        }

        if (hasOr && !orMatched) return false;
    }

    return true;
}

window.toggleQuizDropdown = function(event, dropdownId) {
    event.stopPropagation();

    if (dropdownId === 'quizSampleDropdown')    document.getElementById('quizModeDropdown')     .classList.add('hidden');
    if (dropdownId === 'quizModeDropdown')      document.getElementById('quizSampleDropdown')   .classList.add('hidden');

    document.getElementById(dropdownId).classList.toggle('hidden');
};

window.selectQuizDropdownOption = function(btnId, textId, value, label, dropdownId, event) {
    event.stopPropagation();

    const btn = document.getElementById(btnId);
    btn.setAttribute('data-value', value);

    document.getElementById(textId).innerText = label;
    document.getElementById(dropdownId).classList.add('hidden');
};

function startQuizEngine() {
    enforceQuizInputBounds();

    quizSoundLimit      = parseInt(document.getElementById("quizSoundTimeInput").value);
    quizNoSoundLimit    = parseInt(document.getElementById("quizExtraTimeInput").value);
    const count         = parseInt(document.getElementById("quizSongCountInput").value);
    quizCurrentMode     = document.getElementById("quizModeSelect").getAttribute("data-value");
    quizSampleMode      = document.getElementById("quizSampleSelect").getAttribute("data-value");

    const pool = globalSearchData.filter(song => {
        if (!song.video_url) return false;

        for (let col of quizConfig) {
            if (col.type === "categorical" && col.selectedOptions) {
                if (!col.visible) continue;
                let val = "";

                if (col.id === "type") {
                    const t = song._typeLower;

                    if      (t.includes("opening")) val = "opening";
                    else if (t.includes("ending"))  val = "ending";
                    else if (t.includes("insert"))  val = "insert";
                }

                else if (col.id === "chanting")     val = song._chantingLower;
                else if (col.id === "anime_type")   val = song._animeTypeLower;

                if (val && !col.selectedOptions.has(val)) return false;
            }

            if (col.type === "range" && col.currentMin !== undefined && col.currentMax !== undefined) {
                let val = 0;

                if      (col.id === "vintage")      val = song._vintageParsed;
                else if (col.id === "difficulty")   val = song._diffParsed === -Infinity ? 0 : song._diffParsed;
                else if (col.id === "guessers")     val = song._guessersCount;
                else if (col.id === "listers")      val = song._listersCount;

                if (isNaN(val) || val < col.currentMin || val > col.currentMax) return false;
            }
        }

        return evaluateMatrixConditions(song);
    });

    if (pool.length === 0) {
        alert("No tracks match your current settings");
        return;
    }

    quizActivePool      = [...pool].sort(() => Math.random() - 0.5).slice(0, count);
    quizIndex           = 0;
    quizScore           = 0;
    quizIsPaused        = false;
    quizLeewayActive    = false;
    quizReplayActive    = false;

    document.getElementById("quizSetupScreen")      .classList.add      ("hidden");
    document.getElementById("quizActiveScreen")     .classList.remove   ("hidden");
    document.getElementById("globalQuizPauseBtn")   .classList.remove   ("hidden");
    document.getElementById("globalQuizReturnBtn")  .classList.remove   ("hidden");

    document.getElementById("globalQuizPauseBtn")   .innerText = "❚❚";
    document.getElementById("quizScoreValue")       .innerText = `1/${quizActivePool.length}: 0`;

    if (quizCurrentMode === "mc") {
        document.getElementById("quizInterfaceBody")    .classList.add      ("hidden");
        document.getElementById("quizMcInterfaceBody")  .classList.remove   ("hidden");
    }

    else {
        document.getElementById("quizMcInterfaceBody")  .classList.add      ("hidden");
        document.getElementById("quizInterfaceBody")    .classList.remove   ("hidden");
    }

    executeQuizTrack();
}

function executeQuizTrack() {
    clearQuizTimers();

    quizLeewayActive        = false;
    quizReplayActive        = false;
    quizRevealTimeElapsed   = 0;

    if (quizIndex >= quizActivePool.length) {
        exitQuizEngine();
        return;
    }

    const scoreNode = document.getElementById("quizScoreValue");

    if (scoreNode) {
        scoreNode.innerText     = `${quizIndex + 1}/${quizActivePool.length}: ${quizScore}`;
        scoreNode.style.color   = "";
    }

    document.getElementById("quizScoreValue").innerText = `${quizIndex + 1}/${quizActivePool.length}: ${quizScore}`;

    const label     = document.getElementById("quizResolutionLabel");
    label.innerText = "";

    label.removeAttribute("title");

    const inputEl = document.getElementById("quizAnimeInput");
    if (inputEl) {inputEl.value = ""; inputEl.disabled = false;}

    const song = quizActivePool[quizIndex];

    if (quizCurrentMode === "mc") {
        const isJp              = currentSearchLang === "JP";
        const correctTitle      = isJp ? (song.romaji || song.english || "") : (song.english || song.romaji || "");
        const forbiddenTitles   = new Set();

        if (song.romaji)                            forbiddenTitles.add(song.romaji     .toLowerCase().trim());
        if (song.english)                           forbiddenTitles.add(song.english    .toLowerCase().trim());
        if (song.alts && Array.isArray(song.alts))  song.alts.forEach(alt => {if (alt) forbiddenTitles.add(alt.toLowerCase().trim());});

        const poolSource            = window.localAnimeNamesPool || [];
        const distractorScoringPool = [];

        for (let i = 0; i < poolSource.length; i++) {
            let currentTitle = poolSource[i];
            if (!currentTitle) continue;

            if (currentTitle.includes(" || ")) {
                const parts = currentTitle.split(" || ");
                currentTitle = isJp ? parts[0] : parts[1];
            }

            if (forbiddenTitles.has(currentTitle.toLowerCase().trim())) continue;

            const similarityScore = getQuizStringSimilarity(correctTitle, currentTitle);
            distractorScoringPool.push({title: currentTitle, score: similarityScore});
        }

        distractorScoringPool.sort((a, b) => b.score - a.score);

        const optionsArray = distractorScoringPool.slice(0, 3).map(item => item.title);
        optionsArray.push(correctTitle);
        optionsArray.sort(() => Math.random() - 0.5);

        for (let i = 1; i <= 4; i++) {
            const btn = document.getElementById(`quizMcOpt${i}`);

            if (btn) {
                const optTitle              = optionsArray[i - 1];
                btn.innerText               = optTitle;
                btn.disabled                = false;
                btn.style.backgroundColor   = "";
                btn.style.color             = "";
                btn.style.fontWeight        = "";

                btn.removeAttribute ("data-selected");
                btn.setAttribute    ("data-quiz-hover", encodeURIComponent(optTitle));
            }
        }

        if (typeof setupTooltipListeners === 'function') setupTooltipListeners();
    }

    const anchor        = document.getElementById("quizAudioAnchor");
    anchor.innerHTML    = `<audio id="quizAudioPlayer" src="${song.video_url}" preload="auto"></audio>`;
    quizCurrentAudio    = document.getElementById("quizAudioPlayer");
    quizTimeRemaining   = quizSoundLimit + quizNoSoundLimit;

    updateQuizTimerUI(quizTimeRemaining);
    let isAudioPlaying = false;

    if (quizCurrentAudio) {
        quizCurrentAudio.volume = 0.1;
        quizCurrentAudio.loop   = true;

        quizCurrentAudio.addEventListener('loadedmetadata', () => {
            const maxStart = Math.max(0, quizCurrentAudio.duration - quizSoundLimit);
            let startPos = 0;

            if      (quizSampleMode === "start")  startPos = 0;
            else if (quizSampleMode === "actual") startPos = Math.min(song.start_sample || 0, maxStart);
            else if (quizSampleMode === "random") startPos = Math.random() * maxStart;
            else if (quizSampleMode === "end")    startPos = maxStart;

            quizCurrentAudio.currentTime = startPos;
        });

        quizCurrentAudio.addEventListener('playing', () => {
            if (!isAudioPlaying) {
                isAudioPlaying = true;

                quizTimerId = setInterval(() => {
                    if (quizLeewayActive) return;

                    if (!quizReplayActive) {
                        if (isAudioPlaying) {
                            quizTimeRemaining -= 1;

                            if (quizTimeRemaining <= 0) {
                                quizTimeRemaining = 0;

                                updateQuizTimerUI   (0);
                                resolveQuizItem     (false);

                                return;
                            }

                            updateQuizTimerUI(quizTimeRemaining);
                            if (quizTimeRemaining <= quizNoSoundLimit && quizCurrentAudio && !quizCurrentAudio.paused) quizCurrentAudio.pause();
                        }
                    }

                    else {
                        quizRevealTimeElapsed += 1;
                        let displayCountdown = Math.max(0, (quizSoundLimit + quizNoSoundLimit) - quizRevealTimeElapsed);
                        updateQuizTimerUI(displayCountdown);

                        if (displayCountdown <= 0) {
                            if      (quizIsPaused)                          quizLeewayActive = true;
                            else if (quizIndex >= quizActivePool.length)    exitQuizEngine();
                            else                                            executeQuizTrack();
                        }
                    }
                }, 1000);
            }

            if (!quizAudioTimeoutId && !quizReplayActive) {
                quizAudioTimeoutId = setTimeout(() => {if (quizCurrentAudio && !quizReplayActive) quizCurrentAudio.pause()}, quizSoundLimit * 1000);
            }
        });

        quizCurrentAudio.play().catch(() => {});
    }

    if (quizCurrentMode === "entry" && inputEl) inputEl.focus();

    if (inputEl) {
        inputEl.oninput     = null; 
        inputEl.onkeydown   = handleQuizInputKeyDown;
    }
}

function handleQuizInputKeyDown(e) {
    if (quizCurrentMode === "mc") return;

    if (e.key === "Enter") {
        e.preventDefault();

        if (!quizReplayActive && !quizLeewayActive) {
            const song          = quizActivePool[quizIndex];
            const animeAns      = document.getElementById("quizAnimeInput").value;
            const normAnimeAns  = normalizeQuizString(animeAns);
            const normRomaji    = normalizeQuizString(song.romaji);
            const normEnglish   = normalizeQuizString(song.english);
            const isCorrect     = normAnimeAns.length > 0 && (
                normRomaji  === normAnimeAns || 
                normEnglish === normAnimeAns || 
                (song.alts && song.alts.some(alt => normalizeQuizString(alt) === normAnimeAns))
            );

            resolveQuizItem(isCorrect);
        }
    }
}

function handleQuizSkipClick() {
    if (quizReplayActive || quizLeewayActive) {
        if (quizIsPaused) return; 
        clearQuizTimers();

        if (quizCurrentAudio) {quizCurrentAudio.pause(); quizCurrentAudio = null;}
        executeQuizTrack();
    }

    else {
        const song          = quizActivePool[quizIndex];
        const animeAns      = document.getElementById("quizAnimeInput").value;
        const normAnimeAns  = normalizeQuizString(animeAns);
        const normRomaji    = normalizeQuizString(song.romaji);
        const normEnglish   = normalizeQuizString(song.english);
        const isCorrect     = normAnimeAns.length > 0 && (
            normRomaji  === normAnimeAns || 
            normEnglish === normAnimeAns || 
            (song.alts && song.alts.some(alt => normalizeQuizString(alt) === normAnimeAns))
        );

        resolveQuizItem(isCorrect);
    }
}

function handleMcOptionClick(btn) {
    if (quizReplayActive || quizLeewayActive) return;

    const selectedTitle = btn.innerText;
    const song          = quizActivePool[quizIndex];
    const isJp          = currentSearchLang === "JP";
    const correctTitle  = isJp ? (song.romaji || song.english || "") : (song.english || song.romaji || "");

    btn.setAttribute("data-selected", "true");
    resolveQuizItem(selectedTitle === correctTitle);
}

function handleMcSkipClick() {
    if (quizReplayActive || quizLeewayActive) {
        if (quizIsPaused) return; 
        clearQuizTimers();

        if (quizCurrentAudio) {quizCurrentAudio.pause(); quizCurrentAudio = null;}
        executeQuizTrack();
    }

    else resolveQuizItem(false);
}

function resolveQuizItem(isCorrect) {
    closeQuizAutocomplete();

    if (quizAudioTimeoutId)     {clearTimeout(quizAudioTimeoutId);  quizAudioTimeoutId  = null}
    if (quizLeewayTimeoutId)    {clearTimeout(quizLeewayTimeoutId); quizLeewayTimeoutId = null}

    const song = quizActivePool[quizIndex];
    if (isCorrect) quizScore++;

    const displayIndex  = Math.min(quizIndex + 1, quizActivePool.length);
    const scoreNode     = document.getElementById("quizScoreValue");

    if (scoreNode) scoreNode.innerText = `${displayIndex}/${quizActivePool.length}: ${quizScore}`

    const inputEl = document.getElementById("quizAnimeInput");
    if (inputEl) inputEl.disabled = true;

    if (quizCurrentMode === "mc") {
        const isJp          = currentSearchLang === "JP";
        const correctTitle  = isJp ? (song.romaji || song.english || "") : (song.english || song.romaji || "");

        for (let i = 1; i <= 4; i++) {
            const btn = document.getElementById(`quizMcOpt${i}`);

            if (btn) {
                btn.disabled = true;

                if (btn.innerText === correctTitle) {
                    btn.style.backgroundColor   = c2;
                    btn.style.color             = "white";
                    btn.style.fontWeight        = "bold";
                }

                else if (btn.getAttribute("data-selected") === "true") {
                    btn.style.backgroundColor   = c0;
                    btn.style.color             = "white";
                    btn.style.fontWeight        = "bold";
                }
            }
        }
    }

    let typeMarker  = "";
    const rawType   = (song.type || "").toLowerCase();

    if      (rawType.includes("opening"))   typeMarker = "OP" + (song.type.match(/\d+/) || ["1"])[0];
    else if (rawType.includes("ending"))    typeMarker = "ED" + (song.type.match(/\d+/) || ["1"])[0];
    else                                    typeMarker = "IN";

    const displayTitle  = currentSearchLang === "JP" ? (song.romaji || song.english || "") : (song.english || song.romaji || "");
    const artistName    = Array.isArray(song.artist_arr) ? song.artist_arr.join(', ') : String(song.artist_arr || song.artist_raw || '');    
    const label         = document.getElementById("quizResolutionLabel");
    
    if (label) {
        label.style.pointerEvents   = "auto";
        const fullLine              = `${displayTitle} (${typeMarker}): ${song.song || ""} by ${artistName}`;
        let shortLine               = displayTitle;

        if (shortLine.length > 40) {
            let sub         = shortLine.substring(0, 40);
            let lastSpace   = sub.lastIndexOf(" ");
            shortLine       = lastSpace > 0 ? shortLine.substring(0, lastSpace) + " ..." : shortLine.substring(0, 37) + "...";
        }

        label.innerText = shortLine;

        label.removeAttribute   ("title"); 
        label.setAttribute      ("data-quiz-hover", encodeURIComponent(fullLine));

        label.style.cursor          = "help";
        label.style.pointerEvents   = "auto";
    }

    quizReplayActive        = true;
    quizRevealTimeElapsed   = 0;

    updateQuizTimerUI(quizSoundLimit + quizNoSoundLimit);

    if (quizCurrentAudio) {
        quizCurrentAudio.loop = true;
        quizCurrentAudio.play().catch(() => {});
    }

    if (typeof setupTooltipListeners === 'function') setupTooltipListeners(); 
    quizIndex++;
}

function toggleQuizPause() {
    if (quizIndex >= quizActivePool.length && quizActivePool.length > 0) return;
    quizIsPaused = !quizIsPaused;

    const pauseBtn = document.getElementById("globalQuizPauseBtn");
    if (pauseBtn) pauseBtn.innerText = quizIsPaused ? "▶" : "❚❚";

    if (quizIsPaused && quizLeewayTimeoutId) {
        clearTimeout(quizLeewayTimeoutId);
        quizLeewayTimeoutId = null;
    }

    else if (quizLeewayActive) {
        quizLeewayActive    = false;
        quizLeewayTimeoutId = setTimeout(() => {if (!quizIsPaused) executeQuizTrack();}, 3000);
    }
}

function clearQuizTimers() {
    if (quizTimerId) {
        clearInterval   (quizTimerId);
        clearTimeout    (quizTimerId);

        quizTimerId = null;
    }

    if (quizAudioTimeoutId) {
        clearTimeout(quizAudioTimeoutId);
        quizAudioTimeoutId = null;
    }

    if (quizLeewayTimeoutId) {
        clearTimeout(quizLeewayTimeoutId);
        quizLeewayTimeoutId = null;
    }
}

function finishQuizEngine() {
    clearQuizTimers();
    if (quizCurrentAudio) {quizCurrentAudio.pause(); quizCurrentAudio = null}

    const resLabel = document.getElementById("quizResolutionLabel");
    const timerVal = document.getElementById("quizTimerValue");
    const pauseBtn = document.getElementById("globalQuizPauseBtn");

    if (resLabel) {
        resLabel.innerText = `Score: ${quizScore}/${quizActivePool.length}`;
        resLabel.removeAttribute("title");
    }

    if (timerVal) timerVal.innerText = "0";
    if (pauseBtn) pauseBtn.classList.add("hidden");
}

function exitQuizEngine() {
    closeQuizAutocomplete   ();
    clearQuizTimers         ();

    if (quizCurrentAudio) {quizCurrentAudio.pause(); quizCurrentAudio = null}

    const audioAnchor   = document.getElementById("quizAudioAnchor");
    const activeScr     = document.getElementById("quizActiveScreen");
    const setupScr      = document.getElementById("quizSetupScreen");
    const pauseBtn      = document.getElementById("globalQuizPauseBtn");
    const returnBtn     = document.getElementById("globalQuizReturnBtn");

    if (audioAnchor) audioAnchor.innerHTML = "";

    if (activeScr)   activeScr  .classList.add      ("hidden");
    if (setupScr)    setupScr   .classList.remove   ("hidden");
    if (pauseBtn)    pauseBtn   .classList.add      ("hidden");
    if (returnBtn)   returnBtn  .classList.add      ("hidden");
}

const quizInput         = document.getElementById("quizAnimeInput");
const quizAutoDropdown  = document.getElementById("quizAutocompleteDropdown");
let currentFocusIndex   = -1;

if (quizInput && quizAutoDropdown) {
    document.addEventListener("click", (e) => {if (e.target !== quizInput && e.target !== quizAutoDropdown) closeQuizAutocomplete();});

    quizInput.addEventListener("input", () => {
        const query = quizInput.value.trim();

        if (query.length < 3 || quizReplayActive || quizLeewayActive) {
            closeQuizAutocomplete();
            return;
        }

        fetchAnimeSuggestions(query);
    });

    quizInput.addEventListener("keydown", (e) => {
        const items = quizAutoDropdown.getElementsByTagName("button");
        if (!items.length) return;

        if (e.key === "ArrowDown") {
            e.preventDefault();
            currentFocusIndex++;
            setActiveAutocompleteItem(items);
        }

        else if (e.key === "ArrowUp") {
            e.preventDefault();
            currentFocusIndex--;
            setActiveAutocompleteItem(items);
        }

        else if (e.key === "Escape") {
            e.preventDefault        ();
            closeQuizAutocomplete   ();
        }

        else if (e.key === "Enter" && quizAutoDropdown.classList.contains("hidden") === false) {
            if (currentFocusIndex > -1 && items[currentFocusIndex]) {
                e.preventDefault    ();
                e.stopPropagation   ();

                items[currentFocusIndex].click();
            }
        }
    });
}

function getQuizStringSimilarity(str1, str2) {
    const s1 = str1.toLowerCase().replace(/\s+/g, '');
    const s2 = str2.toLowerCase().replace(/\s+/g, '');

    if (s1 === s2)                      return 1.0;
    if (s1.length < 2 || s2.length < 2) return 0.0;

    const getBigrams = (str) => {
        const bigrams = new Set();
        for (let i = 0; i < str.length - 1; i++) bigrams.add(str.substring(i, i + 2));
        return bigrams;
    };

    const b1 = getBigrams(s1);
    const b2 = getBigrams(s2);

    let intersection = 0;
    b1.forEach(gram => {if (b2.has(gram)) intersection++;});

    return (2.0 * intersection) / (b1.size + b2.size);
}

function fetchAnimeSuggestions(query) {
    const normalizedQuery   = query.toLowerCase().trim();
    const pool              = window.localAnimeNamesPool || [];
    const matchedTitles     = [];
    const isJp              = currentSearchLang === "JP"; 

    for (let i = 0; i < pool.length; i++) {
        let title = pool[i];
        if (!title) continue;

        if (title.includes(" || ")) {
            const parts = title.split(" || ");
            title = isJp ? parts[0] : parts[1];
        }

        if (title.toLowerCase().includes(normalizedQuery) && !matchedTitles.includes(title)) matchedTitles.push(title); 
    }

    matchedTitles.sort((a, b) => {
        const similarityA = getQuizStringSimilarity(query, a);
        const similarityB = getQuizStringSimilarity(query, b);

        if (a.length    !== b.length)       return a.length     - b.length;
        if (similarityB !== similarityA)    return similarityB  - similarityA;

        return a.localeCompare(b);
    });

    const topFiveMatches = matchedTitles.slice(0, 5);
    renderAutocompleteOptions(topFiveMatches);
}

function renderAutocompleteOptions(titlesArray) {
    quizAutoDropdown.innerHTML  = "";
    currentFocusIndex           = -1;

    if (!titlesArray.length) {
        closeQuizAutocomplete();
        return;
    }

    titlesArray.forEach(title => {
        const btn       = document.createElement("button");
        btn.type        = "button";
        btn.className   = "w-full h-8 text-left px-3 py-0 hover:bg-black hover:text-white transition-colors cursor-pointer border-b last:border-0 border-gray-100 whitespace-nowrap overflow-hidden text-ellipsis bg-white text-black font-medium flex items-center";
        btn.textContent = title;

        btn.addEventListener("click", (e) => {
            e.preventDefault();
            quizInput.value = title;
            closeQuizAutocomplete();
            quizInput.focus();
        });

        quizAutoDropdown.appendChild(btn);
    });

    quizAutoDropdown.classList.remove("hidden");
}

function setActiveAutocompleteItem(items) {
    for (let i = 0; i < items.length; i++) {
        items[i].classList.remove   ("bg-black", "text-white");
        items[i].classList.add      ("bg-white", "text-black");
    }

    if (currentFocusIndex >= items.length)  currentFocusIndex = 0;
    if (currentFocusIndex < 0)              currentFocusIndex = items.length - 1;

    if (items[currentFocusIndex]) {
        items[currentFocusIndex].classList.remove   ("bg-white", "text-black");
        items[currentFocusIndex].classList.add      ("bg-black", "text-white");

        items[currentFocusIndex].scrollIntoView({block: "nearest"});
    }
}

function closeQuizAutocomplete() {
    if (quizAutoDropdown) {
        quizAutoDropdown.classList.add("hidden");
        quizAutoDropdown.innerHTML = "";
    }

    currentFocusIndex = -1;
}

const checkIntervalForData = setInterval(() => {
    if (globalSearchData && globalSearchData.length > 0 && typeof players !== 'undefined' && players && players.length > 0) {
        clearInterval               (checkIntervalForData);
        initQuizSettingsCheckboxes  ();
        updateQuizHelpDropdown      ();
    }
}, 100);