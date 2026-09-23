let globalSortState     = {columnName: "Anime", ascending: true};
let currentSearchLang   = "JP";
let searchMatrixStates  = {seen: {}, correct: {}, list: {}};

const searchHeadersConfig = [
    {
        id          : "anime",
        name        : "Anime",
        visible     : true,
        type        : "text"
    },
    {
        id          : "alt",
        name        : "Alt",
        visible     : false,
        type        : "range",
        min         : 0,
        max         : 20,
        step        : 1
    },
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
    {
        id          : "song",
        name        : "Song",
        visible     : true,
        type        : "text",
        locked      : true
    },
    {
        id          : "artist",
        name        : "Artist",
        visible     : true,
        type        : "text"},
    {
        id          : "composer",
        name        : "Composer",
        visible     : false,
        type        : "text"
    },
    {
        id          : "arranger",
        name        : "Arranger",
        visible     : false,
        type        : "text"
    },
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

window.toggleSearchLanguage = function() {
    const btn               = document.getElementById("langToggleBtn");
    currentSearchLang       = currentSearchLang === "JP" ? "EN" : "JP";
    if (btn) btn.innerText  = currentSearchLang;

    sortSearchData      ();
    triggerTableRefresh ();

    if (typeof renderTierCharts === 'function') renderTierCharts();
};

function initColumnSettingsCheckboxes() {
    const container = document.getElementById("columnCheckboxContainer");
    const masterChk = document.getElementById("allColumnsMasterCheckbox");

    if (!container || !masterChk) return;
    container.innerHTML = "";

    if (globalSearchData && globalSearchData.length > 0) {
        const parsedVints = globalSearchData.map(s => s._vintageParsed).filter(v => !isNaN(v) && v !== -Infinity);
        const parsedDiffs = globalSearchData.map(s => s._diffParsed).filter(d => !isNaN(d) && d !== -Infinity);

        const vintConfig  = searchHeadersConfig.find(c => c.id === "vintage");
        const diffConfig  = searchHeadersConfig.find(c => c.id === "difficulty");

        if (vintConfig && parsedVints.length > 0) {
            vintConfig.min = Math.floor(Math.min(...parsedVints));
            vintConfig.max = Math.ceil(Math.max(...parsedVints));
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
        const isChecked = masterChk.checked;

        searchHeadersConfig.forEach(c => {
            if (c.locked) {
                c.visible = true;
                return;
            }

            c.visible = isChecked;

            if (c.type === "categorical" && c.subOptions) {
                if (isChecked)  c.selectedOptions = new Set(c.subOptions.map(o => o.toLowerCase()));
                else            c.selectedOptions.clear();
            }
        });

        if (!isChecked) {
            searchHeadersConfig.forEach(c => {
                if (c.type === "range") {
                    c.currentMin = c.min;
                    c.currentMax = c.max;
                }
            });
        }

        const individualCheckboxes = container.getElementsByClassName('col-toggle-checkbox');
        for (let i = 0; i < individualCheckboxes.length; i++) if (!individualCheckboxes[i].disabled) individualCheckboxes[i].checked = isChecked;

        const allInputs = container.getElementsByTagName('input');

        for (let i = 0; i < allInputs.length; i++) {
            if (allInputs[i].type === 'checkbox' && !allInputs[i].disabled) {
                allInputs[i].checked        = isChecked;
                allInputs[i].indeterminate  = false;
            }
        }

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

        if (col.locked) {
            col.visible     = true;
            chk.checked     = true;
            chk.disabled    = true;
        }

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
            const subContainer      = document.createElement("div");
            subContainer.className  = "pl-6 flex flex-col text-xs";

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

                    updateMasterCheckboxState   ();
                    triggerTableRefresh         ();
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

            if (col.id !== "alt") {
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

    const separator     = document.createElement("div");
    separator.className = "border-b my-1.5";

    container.appendChild(separator);

    const matrices = [{key: "seen", name: "Seen"}, {key: "correct", name: "Correct"}, {key: "list", name: "List"}];
    if (searchMatrixStates._sectionEnabled === undefined) searchMatrixStates._sectionEnabled = {seen: true, correct: true, list: true};

    matrices.forEach(m => {
        const groupWrapper      = document.createElement("div");
        groupWrapper.className  = "flex flex-col select-none";

        const headerLabel       = document.createElement("label");
        headerLabel.className   = "flex items-center gap-1.5 font-bold text-black block text-xs cursor-pointer mt-0.5";

        const sectionToggle     = document.createElement("input");
        sectionToggle.type      = "checkbox";
        sectionToggle.className = "rounded accent-black";
        sectionToggle.checked   = searchMatrixStates._sectionEnabled[m.key];

        headerLabel     .appendChild(sectionToggle);
        headerLabel     .appendChild(document.createTextNode(m.name));
        groupWrapper    .appendChild(headerLabel);

        const subContainer      = document.createElement("div");
        subContainer.className  = "pl-4 flex flex-col gap-0.5 text-xs w-full";

        const updateSubContainerOpacity = () => {
            subContainer.style.opacity          = searchMatrixStates._sectionEnabled[m.key] ? "1"     : "0.5";
            subContainer.style.pointerEvents    = searchMatrixStates._sectionEnabled[m.key] ? "auto"  : "none";
        };

        sectionToggle.addEventListener("change", () => {
            searchMatrixStates._sectionEnabled[m.key] = sectionToggle.checked;

            updateSubContainerOpacity   ();
            triggerTableRefresh         (); 
        });

        if (typeof players !== 'undefined' && players) {
            players.forEach(pRow => {
                const rawPlayer = pRow["Player"];
                const pName     = typeof rawPlayer === 'object' ? String(rawPlayer.count || "") : String(rawPlayer);

                if (searchMatrixStates[m.key][pName] === undefined) searchMatrixStates[m.key][pName] = " ";

                const subLabel              = document.createElement("label");
                subLabel.className          = "flex items-center gap-1 cursor-pointer text-gray-700 hover:text-black py-0.5 pr-6 w-full min-w-max";
                const indicator             = document.createElement("span");
                const isBlank               = searchMatrixStates[m.key][pName] === " ";
                indicator.className         = `font-mono font-bold text-center border inline-flex items-center justify-center w-3.5 h-3.5 bg-white text-black border-gray-400 align-middle select-none shrink-0 p-0 text-sm self-center scale-90 ${isBlank ? "bg-white text-black border-gray-400" : "bg-black text-white border-black"}`;
                indicator.innerText         = searchMatrixStates[m.key][pName];
                indicator.style.lineHeight  = "1";

                subLabel.appendChild(indicator);
                subLabel.appendChild(document.createTextNode(pName));

                subLabel.addEventListener("click", (e) => {
                    e.preventDefault    ();
                    e.stopPropagation   ();

                    const currentState          = searchMatrixStates[m.key][pName];
                    const currentActiveCount    = Object.keys(searchMatrixStates[m.key]).filter(k => searchMatrixStates[m.key][k] !== " ").length;

                    if (currentState === " " && currentActiveCount >= 2) {
                        alert(`You can only configure ${m.name} filters for up to 2 players`);
                        return;
                    }

                    let nextState = " ";

                    if      (currentState === " ") nextState = "+";
                    else if (currentState === "+") nextState = "~";
                    else if (currentState === "~") nextState = "-";

                    searchMatrixStates[m.key][pName]    = nextState;
                    indicator.innerText                 = nextState;

                    if (nextState === " ")  indicator.className = "font-mono font-bold text-center border inline-flex items-center justify-center w-3.5 h-3.5 bg-white text-black border-gray-400 align-middle select-none shrink-0 p-0 text-sm self-center scale-90";
                    else                    indicator.className = "font-mono font-bold text-center border inline-flex items-center justify-center w-3.5 h-3.5 bg-black text-white border-black align-middle select-none shrink-0 p-0 text-sm font-bold self-center scale-90";

                    triggerTableRefresh(); 
                });

                subContainer.appendChild(subLabel);
            });
        }

        updateSubContainerOpacity();

        groupWrapper    .appendChild(subContainer);
        container       .appendChild(groupWrapper);
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
            case "Alt"          : valA = a._altsCount;                              valB = b._altsCount;                            break;
            default             : return 0;
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

    sortSearchData      ();
    triggerTableRefresh ();
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

        case "alt": {
            if (isNaN(value))   return song._altsLower.some(altName => altName.includes(value));
            else                return matchNumericConstraint(song._altsCount, operator, parseFloat(value));
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

function evaluateSearchMatrixConditions(song) {
    const matrices = ["seen", "correct", "list"];

    for (let mKey of matrices) {
        if (searchMatrixStates._sectionEnabled && searchMatrixStates._sectionEnabled[mKey] === false) continue;

        let hasMust         = false;
        let mustMatched     = false;
        let activeOrCount   = 0;

        for (let pName in searchMatrixStates[mKey]) if (searchMatrixStates[mKey][pName] === "~") activeOrCount++;

        let hasOr       = activeOrCount > 1;
        let orMatched   = false;

        for (let pName in searchMatrixStates[mKey]) {
            const state = searchMatrixStates[mKey][pName];
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

                else if (col.id === "chanting")         matchValue = song._chantingLower;
                else if (col.id === "anime_type")       matchValue = song._animeTypeLower;

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

        return evaluateSearchMatrixConditions(song);
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

    table.querySelectorAll('thead th').forEach(th => {th.addEventListener('click', () => handleSearchSort(th.getAttribute('data-header-name')))});

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
                    td.textContent  = song.difficulty === "Unrated" ? "" : song.difficulty;
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
                    const matchArr      = arrVisible && (song.composer === song.arranger);
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
                        td.className = isOverflown ? "cursor-help hover:bg-gray-100 text-left text-black font-normal" : "text-left font-normal text-black";
                        if (isOverflown) td.setAttribute("data-songs", encodeURIComponent(JSON.stringify(splitComps)));
                        td.textContent = trimNames(song.composer);
                    }

                    break;
                }

                case "arranger": {
                    let flatArr         = Array.isArray(song.arranger) ? song.arranger.join('||') : String(song.arranger || '');
                    let splitArrs       = flatArr.replace(/\s*(?:,|\b(?<!\d)\/(?!\d)\b|・|&|×|\bfeat\.)\s*/gi, '||').split('||').map(x => x.trim()).filter(Boolean);
                    const isOverflown   = splitArrs.length > 3;

                    td.className = isOverflown ? "cursor-help hover:bg-gray-100 text-left text-black font-normal" : "text-left font-normal text-black";
                    if (isOverflown) td.setAttribute("data-songs", encodeURIComponent(JSON.stringify(splitArrs)));
                    td.textContent = trimNames(song.arranger);
                    break;
                }

                case "guessers": {
                    const hasGuesses    = song.guessers_hover && song.guessers_hover.length > 0;
                    td.className        = hasGuesses ? "cursor-help hover:bg-gray-100 text-center text-black font-normal" : "text-center text-black font-normal";

                    if (hasGuesses) td.setAttribute("data-songs", encodeURIComponent(JSON.stringify(song.guessers_hover)));
                    td.textContent = song._guessersCount === 0 ? "" : song._guessersCount;
                    break;
                }

                case "listers": {
                    const hasLists  = song.listers_hover && song.listers_hover.length > 0;
                    td.className    = hasLists ? "cursor-help hover:bg-gray-100 text-center text-black font-normal" : "text-center text-black font-normal";

                    if (hasLists) td.setAttribute("data-songs", encodeURIComponent(JSON.stringify(song.listers_hover)));
                    td.textContent = song._listersCount === 0 ? "" : song._listersCount;
                    break;
                }

                case "alt": {
                    const hasAlts   = song.alts && song.alts.length > 0;
                    td.className    = hasAlts ? "cursor-help hover:bg-gray-100 text-center text-black font-normal font-mono" : "text-center text-black font-normal font-mono";

                    if (hasAlts) {
                        let sortedAlts = [...song.alts].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
                        td.setAttribute("data-songs", encodeURIComponent(JSON.stringify(sortedAlts)));
                    }

                    td.textContent = song._altsCount === 0 ? "" : song._altsCount;
                    break;
                }
            }

            tr.appendChild(td);
        });

        fragment.appendChild(tr);
    });

    tbody.appendChild(fragment);
    if (typeof setupTooltipListeners === 'function') setupTooltipListeners();
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
            returns songs by Aoi Koga or Tomori Kusunoki with difficulties above 20<br><br>
            Click the <b class="bg-black text-white px-1 w-4 h-4 rounded">⚙</b> button to configure your view settings<br>
        </p>
        <p class="font-bold border-b pb-1 mb-1 mt-3">Filter</p>
        <p class="mb-2 text-xs">
            Configure <b>Seen</b>, <b>Correct</b>, and <b>List</b> player matrix filters as follows:<br>
            • <b class="bg-white text-black border border-gray-400 px-1 rounded inline-block w-4 h-4 text-center">&nbsp;</b>&nbsp;Ignore this player<br>
            • <b class="bg-black text-white px-1 rounded inline-block w-4 h-4 text-center">+</b>&nbsp;Force this player<br>
            • <b class="bg-black text-white px-1 rounded inline-block w-4 h-4 text-center">~</b>&nbsp;Find at least one player that match<br>
            • <b class="bg-black text-white px-1 rounded inline-block w-4 h-4 text-center">-</b>&nbsp;Force against this player<br>
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

Promise.all([
    fetch('jsons/search.json')  .then(res => res.json()),
    fetch('jsons/name.json')    .then(res => res.json())
]).then(([searchJson, nameJson]) => {
    window.localAnimeNamesPool = Array.isArray(nameJson) ? nameJson : [];

    globalSearchData = searchJson.map(song => {
        if (!song.guessers_flat && song.guessers_hover) song.guessers_flat  = song.guessers_hover.map(g => g.split(' (')[0]);
        if (!song.listers_flat  && song.listers_hover)  song.listers_flat   = song.listers_hover;

        song._altsLower         = (song.alts                || []).map(a => String(a).toLowerCase());
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

        song._diffParsed    = song.difficulty === "Unrated"   ? -Infinity                   : parseFloat(song.difficulty);
        song._altsCount     = Array.isArray(song.alts)        ? song.alts.length            : 0;
        song._guessersCount = song.guessers_flat              ? song.guessers_flat.length   : 0;
        song._listersCount  = song.listers_flat               ? song.listers_flat.length    : 0;

        song._vintageParsed = parseVintageToFloat(song.vintage);

        song._correctTeamsLower = (song.correct_teams_flat  || []).map(tid => {
            const leader = window.dashboardData.json_teams ? window.dashboardData.json_teams.find(t => t._tid === tid || t.tid === tid) : null;
            return leader ? leader["Team Leader"].toLowerCase() : "";
        });

        return song;
    });

    const countInput = document.getElementById("quizSongCountInput");
    if (countInput && globalSearchData.length > 0) countInput.setAttribute("max", globalSearchData.length);

    initColumnSettingsCheckboxes    ();
    sortSearchData                  ();
    renderSearchTable               (globalSearchData);

    if (typeof renderTierCharts === 'function') renderTierCharts();
    updateSearchHelpDropdown();

    const searchInput = document.getElementById('songSearchInput');

    if (searchInput) {
        const processQuery = (e) => {
            const rawQuery = searchInput.value.trim();
            if (!rawQuery) {renderSearchTable(globalSearchData); return}

            const tokens        = [];
            const tokenRegex    = /\(|\)|or\b|and\b|[a-zA-Z0-9_/-]+(=|>=|!=|!:|[:<>==])"[^"]*"|[^\s"()]+|"[^"]*"/gi;

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
 
                    song._altsLower     .some(a => a.includes(wordClean)) ||
                    song._genresLower   .some(g => g.includes(wordClean)) ||
                    song._tagsLower     .some(t => t.includes(wordClean)) ||

                    song.difficulty.toLowerCase().includes(wordClean)
                );
            }

            const rpnTokens = parseToRPN(tokens);

            const filtered = globalSearchData.filter(song => {
                if (rpnTokens.length === 0) return true;
                const stack = [];

                for (let token of rpnTokens) {
                    const lowerToken = typeof token === 'string' ? token.toLowerCase() : '';

                    if (lowerToken === 'and') {
                        const b = stack.pop(); const a = stack.pop();
                        stack.push(a && b);
                    }

                    else if (lowerToken === 'or') {
                        const b = stack.pop(); const a = stack.pop();
                        stack.push(a || b);
                    }

                    else stack.push(evaluateSingleToken(song, token));
                }

                return stack[0];
            });

            renderSearchTable(filtered);
        };

        searchInput.addEventListener('input',        debounce(processQuery, 250));
        searchInput.addEventListener('input-direct', processQuery);
    }
}).catch(err => console.error("Error setting up lookup engine layout context mapping:", err));