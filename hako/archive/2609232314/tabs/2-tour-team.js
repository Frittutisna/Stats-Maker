function renderTourTable() {
    const table = document.getElementById('tourStatsTable');
    if (!tourStats || !tourStats.length) return;

    const half          = Math.ceil(tourStats.length / 2);
    const leftSlice     = tourStats.slice(0, half);
    const rightSlice    = tourStats.slice(half);

    let thead = `<thead>
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

    if (!use_teams || !teamStats || !teamStats.length) {
        table.innerHTML                     = "";
        if (spacer) spacer.style.display    = "none";
        return;
    }

    let headers = Object.keys(teamStats[0]);

    if (!watched) {
        const skipped = new Set(["Rig Synergy", "Off Synergy", "Shared Rigs"]);
        headers = headers.filter(h => !skipped.has(h));
    }

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

            if      (h === "Win Record")    finalVal = displayVal;
            else if (h === "Team Leader")   finalVal = `<b>${displayVal}</b>`;

            if (rawCell !== null && typeof rawCell === 'object' && rawCell.details && rawCell.details.length > 0) {
                let encodedDetails  = encodeURIComponent(JSON.stringify(rawCell.details));
                let clickHandler    = "";

                if (h === "Total 1/8s") clickHandler = ` onclick="searchTeamSolos('${row["Team Leader"]}')"`;
                tbody += `<td class="${cellStyle.trim()}" data-songs="${encodedDetails}" data-metric="${h}"${clickHandler}>${finalVal}</td>`;
            }

            else tbody += `<td class="${cellStyle.trim()}">${finalVal}</td>`;
        });

        tbody += "</tr>";
    });

    table.innerHTML = thead + tbody + "</tbody>";
}

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
    const searchTabBtn  = Array.from(document.querySelectorAll('.tab-btn')).find(btn => btn.getAttribute('onclick') && btn.getAttribute('onclick').includes('search-tab'));
    const searchInput   = document.getElementById('songSearchInput');

    if (searchTabBtn && searchInput) {
        searchTabBtn.click();

        searchInput.value = "";
        searchInput.value = `correctteam:${teamLeader.toLowerCase()} correct:1`;

        searchInput.dispatchEvent(new Event('input-direct'));
    }
};

renderTourTable();
renderTeamTable();