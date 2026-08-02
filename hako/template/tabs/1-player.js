let globalPlayerSortState   = {columnName: "GR", ascending: false};
let globalFilteredPlayers   = [];
let globalMetricHighlights  = {};

const playerHeadersMasterConfig = [
    {id: "player",              name: "Player",                 ascMetric: false,   teamReq: false, watchedReq: false,  def: true,  type: "text",           locked: true},
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
                    col.min = Math.floor    (Math.min(...numericValues));
                    col.max = Math.ceil     (Math.max(...numericValues));
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
        const isChecked = masterChk.checked;

        activePlayerHeadersConfig.forEach(c => { 
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
            activePlayerHeadersConfig.forEach(c => {
                if (c.type === "range") {
                    c.currentMin = c.min;
                    c.currentMax = c.max;
                }
            });
        }

        const individualCheckboxes = container.getElementsByClassName('player-col-toggle-checkbox');
        for (let i = 0; i < individualCheckboxes.length; i++) individualCheckboxes[i].checked = isChecked;

        const allInputs = container.getElementsByTagName('input');

        for (let i = 0; i < allInputs.length; i++) {
            if (allInputs[i].type === 'checkbox') {
                allInputs[i].checked        = isChecked;
                allInputs[i].indeterminate  = false;
            }
        }

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
            explanationIndicator.innerHTML = "⦾";
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

                    if (cleanVal.startsWith('"') && cleanVal.endsWith('"')) cleanVal = cleanVal.slice(1, -1)            .trim();
                    else                                                    cleanVal = cleanVal.replace(/^"|"$/g, '')   .trim();

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
    document.querySelectorAll('input[name="playerMetricModeRadio"]').forEach(radio => {radio.addEventListener('change', (e) => {window.updatePlayerMetricModeFromRadio(e.target.value)})});
}

function updatePlayerHelpDropdown() {
    const dropdown = document.getElementById("playerGuideDropdown");
    if (!dropdown) return;

    const hasDeltaMetrics = activePlayerHeadersConfig.some(c => ["GR Δ", "UF Δ", "OP Δ", "ED Δ", "IN Δ"].includes(c.name));
    let configSectionText = "";

    if (hasDeltaMetrics) {
        configSectionText = `
            <br><br>
            Click the <b class="bg-black text-white px-1 rounded">⚙</b> button to configure your view settings<br>
            • <b>Current:</b> Shows empirical metrics for (OP/ED/IN) GR and UF<br>
            • <b>Delta:</b> Compares them against their historical baselines
        `;
    }

    dropdown.innerHTML = `
        <p class="font-bold border-b pb-1 mb-1">Guide</p>
        <p class="mb-1 pt-1 text-xs">
            Search using <code class="bg-gray-200 px-1 rounded font-mono text-xs">value</code> or <code class="bg-gray-200 px-1 rounded font-mono text-xs">columnname:value</code><br><br>
            You can replace <code class="bg-gray-200 px-1 rounded font-mono text-xs">:</code> with arithmetic operators (<code class="bg-gray-200 px-1 rounded font-mono text-xs">=, !:, !=, &lt;, &gt;, &lt;=, &gt;=)</code>,<br>
            <code class="bg-gray-200 px-1 rounded font-mono text-xs">uf</code> with 
            <code class="bg-gray-200 px-1 rounded font-mono text-xs">usefulness</code>, 
            <code class="bg-gray-200 px-1 rounded font-mono text-xs">guess rate</code> with 
            <code class="bg-gray-200 px-1 rounded font-mono text-xs">gr</code>, 
            <code class="bg-gray-200 px-1 rounded font-mono text-xs">1/8s</code> with 
            <code class="bg-gray-200 px-1 rounded font-mono text-xs">solos</code>, 
            <code class="bg-gray-200 px-1 rounded font-mono text-xs">2/8s</code> with 
            <code class="bg-gray-200 px-1 rounded font-mono text-xs">doubles</code>, and 
            <code class="bg-gray-200 px-1 rounded font-mono text-xs">7/8s</code> with 
            <code class="bg-gray-200 px-1 rounded font-mono text-xs">sevens</code><br><br>
            Combine query terms using explicit <code class="bg-gray-200 px-1 rounded font-mono text-xs">and/or</code> keywords<br>
            Group precedence with <code class="bg-gray-200 px-1 rounded font-mono text-xs">(brackets)</code><br>
            Wrap multi-word values in <code class="bg-gray-200 px-1 rounded font-mono text-xs">"double-quotes"</code><br><br>
            <code class="bg-gray-200 px-1 rounded font-mono text-xs">elo&lt;90 or (opgr&gt;80 edguessrate&lt;30)</code><br>
            returns players that either have an Elo below 90 or those who have OP and ED Guess Rates above 80% and below 30%${configSectionText}
        </p>
    `;
}

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

initPlayerRadioSettingsListeners    ();
renderPlayerTable                   ();
updatePlayerHelpDropdown            ();