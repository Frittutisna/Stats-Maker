const tabContainer  = document.getElementById('tabContainer');
const tourTabBtn    = document.getElementById('tourTabBtn');
const helpAnchor    = document.getElementById('globalHelpWrapper');

if (use_teams)  tourTabBtn.innerText = "Tour/Team";
else            tourTabBtn.innerText = "Tour";

if (use_teams)  helpAnchor.insertAdjacentHTML('beforebegin', `<button class="tab-btn" onclick="switchDashboardTab(event, 'tier-tab')">Tier</button>`);
                helpAnchor.insertAdjacentHTML('beforebegin', `<button class="tab-btn" onclick="switchDashboardTab(event, 'song-tab')">Song</button>`);
if (watched)    helpAnchor.insertAdjacentHTML('beforebegin', `<button class="tab-btn" onclick="switchDashboardTab(event, 'guess-tab')">Guess/List</button>`);
else            helpAnchor.insertAdjacentHTML('beforebegin', `<button class="tab-btn" onclick="switchDashboardTab(event, 'guess-tab')">Guess</button>`);
                helpAnchor.insertAdjacentHTML('beforebegin', `<button class="tab-btn" onclick="switchDashboardTab(event, 'search-tab')">Search</button>`);
                helpAnchor.insertAdjacentHTML('beforebegin', `<button class="tab-btn" onclick="switchDashboardTab(event, 'quiz-tab')">Quiz</button>`);

function switchDashboardTab(evt, tabId) {
    if (tabId !== 'quiz-tab' && typeof exitQuizEngine === 'function') exitQuizEngine();

    document.querySelectorAll('.tab-content')   .forEach(el => el.classList.remove('active-content'));
    document.querySelectorAll('.tab-btn')       .forEach(el => el.classList.remove('active-tab'));

    const targetTab = document.getElementById(tabId);

    if (targetTab)                  targetTab.classList.add('active-content');
    if (evt && evt.currentTarget)   evt.currentTarget.classList.add('active-tab');

    window.dispatchEvent(new Event('resize'));

    const gearWrapper = document.getElementById('globalGearWrapper');
    const helpWrapper = document.getElementById('globalHelpWrapper');

    if (gearWrapper) {
        if (['player-tab', 'tier-tab', 'search-tab', 'quiz-tab'].includes(tabId) || (tabId === 'guess-tab' && watched)) gearWrapper.classList.remove    ('invisible');
        else                                                                                                            gearWrapper.classList.add       ('invisible');
    }

    if (helpWrapper) {
        if (['player-tab', 'tier-tab', 'guess-tab', 'search-tab', 'song-tab', 'quiz-tab'].includes(tabId))  helpWrapper.classList.remove    ('invisible');
        else                                                                                                helpWrapper.classList.add       ('invisible');
    }

    if (tabId !== 'quiz-tab') {
        const pauseBtn  = document.getElementById("globalQuizPauseBtn");
        const returnBtn = document.getElementById("globalQuizReturnBtn");

        if (pauseBtn)   pauseBtn    .classList.add("hidden");
        if (returnBtn)  returnBtn   .classList.add("hidden");
    }

    document.querySelectorAll('#globalGearWrapper > div, #globalHelpWrapper > div').forEach(el => {
        if (el.id.includes('Dropdown')) el.classList.add('hidden');
    });
}

window.toggleGlobalGear = function(event) {
    event.stopPropagation();
    const activeTab = document.querySelector('.tab-content.active-content').id;

    if (activeTab === 'player-tab') {
        document.getElementById("playerColumnSettingsDropdown") .classList.toggle   ("hidden");
        document.getElementById("columnSettingsDropdown")       .classList.add      ("hidden");
        document.getElementById("tierSettingsDropdown")         .classList.add      ("hidden");
        document.getElementById("guessSettingsDropdown")        .classList.add      ("hidden");
        document.getElementById("quizSettingsDropdown")         .classList.add      ("hidden");
    }

    else if (activeTab === 'tier-tab') {
        document.getElementById("tierSettingsDropdown")         .classList.toggle   ("hidden");
        document.getElementById("playerColumnSettingsDropdown") .classList.add      ("hidden");
        document.getElementById("columnSettingsDropdown")       .classList.add      ("hidden");
        document.getElementById("guessSettingsDropdown")        .classList.add      ("hidden");
        document.getElementById("quizSettingsDropdown")         .classList.add      ("hidden");
    }

    else if (activeTab === 'search-tab') {
        document.getElementById("columnSettingsDropdown")       .classList.toggle   ("hidden");
        document.getElementById("playerColumnSettingsDropdown") .classList.add      ("hidden");
        document.getElementById("tierSettingsDropdown")         .classList.add      ("hidden");
        document.getElementById("guessSettingsDropdown")        .classList.add      ("hidden");
        document.getElementById("quizSettingsDropdown")         .classList.add      ("hidden");
    }

    else if (activeTab === 'guess-tab' && watched) {
        document.getElementById("guessSettingsDropdown")        .classList.toggle   ("hidden");
        document.getElementById("playerColumnSettingsDropdown") .classList.add      ("hidden");
        document.getElementById("columnSettingsDropdown")       .classList.add      ("hidden");
        document.getElementById("tierSettingsDropdown")         .classList.add      ("hidden");
        document.getElementById("quizSettingsDropdown")         .classList.add      ("hidden");
    }

    else if (activeTab === 'quiz-tab') {
        document.getElementById("quizSettingsDropdown")         .classList.toggle("hidden");
        document.getElementById("playerColumnSettingsDropdown") .classList.add("hidden");
        document.getElementById("columnSettingsDropdown")       .classList.add("hidden");
        document.getElementById("tierSettingsDropdown")         .classList.add("hidden");
        document.getElementById("guessSettingsDropdown")        .classList.add("hidden");
    }
};

window.toggleGlobalHelp = function(event) {
    event.stopPropagation();
    const activeTab = document.querySelector('.tab-content.active-content').id;

    if (activeTab === 'player-tab') {
        document.getElementById("playerGuideDropdown")  .classList.toggle   ("hidden");
        document.getElementById("tierGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("guessGuideDropdown")   .classList.add      ("hidden");
        document.getElementById("songGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("searchGuideDropdown")  .classList.add      ("hidden");
        document.getElementById("quizGuideDropdown")    .classList.add      ("hidden");
    }

    else if (activeTab === 'tier-tab') {
        document.getElementById("tierGuideDropdown")    .classList.toggle   ("hidden");
        document.getElementById("playerGuideDropdown")  .classList.add      ("hidden");
        document.getElementById("guessGuideDropdown")   .classList.add      ("hidden");
        document.getElementById("songGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("searchGuideDropdown")  .classList.add      ("hidden");
        document.getElementById("quizGuideDropdown")    .classList.add      ("hidden");
    }

    else if (activeTab === 'guess-tab') {
        document.getElementById("guessGuideDropdown")   .classList.toggle   ("hidden");
        document.getElementById("playerGuideDropdown")  .classList.add      ("hidden");
        document.getElementById("tierGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("songGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("searchGuideDropdown")  .classList.add      ("hidden");
        document.getElementById("quizGuideDropdown")    .classList.add      ("hidden");
    }

    else if (activeTab === 'song-tab') {
        document.getElementById("songGuideDropdown")    .classList.toggle   ("hidden");
        document.getElementById("playerGuideDropdown")  .classList.add      ("hidden");
        document.getElementById("tierGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("guessGuideDropdown")   .classList.add      ("hidden");
        document.getElementById("searchGuideDropdown")  .classList.add      ("hidden");
        document.getElementById("quizGuideDropdown")    .classList.add      ("hidden");
    }

    else if (activeTab === 'search-tab') {
        document.getElementById("searchGuideDropdown")  .classList.toggle   ("hidden");
        document.getElementById("playerGuideDropdown")  .classList.add      ("hidden");
        document.getElementById("tierGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("songGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("guessGuideDropdown")   .classList.add      ("hidden");
        document.getElementById("quizGuideDropdown")    .classList.add      ("hidden");
    }

    else if (activeTab === 'quiz-tab') {
        document.getElementById("quizGuideDropdown")    .classList.toggle   ("hidden");
        document.getElementById("playerGuideDropdown")  .classList.add      ("hidden");
        document.getElementById("tierGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("guessGuideDropdown")   .classList.add      ("hidden");
        document.getElementById("songGuideDropdown")    .classList.add      ("hidden");
        document.getElementById("searchGuideDropdown")  .classList.add      ("hidden");
    }
};

window.toggleArchiveDropdown = function(event) {
    event.stopPropagation();
    document.getElementById("archiveDropdown").classList.toggle("hidden");

    document.querySelectorAll('#globalGearWrapper > div, #globalHelpWrapper > div').forEach(el => {
        if (el.id.includes('Dropdown')) el.classList.add('hidden');
    });
};

async function populateArchiveDropdown() {
    const dropdown = document.getElementById("archiveDropdown");
    if (!dropdown) return;

    try {
        const response = await fetch("https://api.github.com/repos/frittutisna/Stats-Maker/contents/hako/archive");
        if (!response.ok) throw new Error("Failed to scan archive directory");

        const files             = await response.json();
        const timestampPattern  = /^\d{10}$/;
        const pastToursArchive  = files
            .filter(item => item.type === "dir" && timestampPattern.test(item.name))
            .map(item => {
                const name              = item.name;
                const formattedLabel    = `20${name.substring(0, 2)}/${name.substring(2, 4)}/${name.substring(4, 6)} ${name.substring(6, 8)}:${name.substring(8, 10)} JST`;

                return {id: name, label: formattedLabel};
            })
            .sort((a, b) => b.id.localeCompare(a.id));

        if (pastToursArchive.length === 0) {
            dropdown.innerHTML = `<p class="px-4 py-2 text-gray-500 italic text-center">No archives found</p>`;
            return;
        }

        dropdown.innerHTML = pastToursArchive.map(tour => `
            <a href="https://frittutisna.github.io/Stats-Maker/hako/archive/${tour.id}/index.html?update=1" 
               class="px-2 py-1 text-black border-b last:border-0 block text-center font-bold no-underline transition-colors">
               ${tour.label}
            </a>
        `).join('');
    }

    catch (err) {
        console.error("Failed to scan for Archive dropdown:", err);
        dropdown.innerHTML = `<p class="px-4 py-2 text-red-500 text-center font-bold">Error loading archives</p>`;
    }
}

populateArchiveDropdown();

document.addEventListener("click", () => {
    document.querySelectorAll('#globalGearWrapper > div, #globalHelpWrapper > div').forEach(el => {
        if (el.id.includes('Dropdown')) el.classList.add('hidden');
    });

    const archiveMenu       = document.getElementById('archiveDropdown');
    const quizSampleDrop    = document.getElementById('quizSampleDropdown');
    const quizModeDrop      = document.getElementById('quizModeDropdown');

    if (archiveMenu)    archiveMenu     .classList.add('hidden');
    if (quizSampleDrop) quizSampleDrop  .classList.add('hidden');
    if (quizModeDrop)   quizModeDrop    .classList.add('hidden');
});

const stopProp = (e) => e.stopPropagation();

[
    'playerColumnSettingsDropdown',
    'tierSettingsDropdown',
    'guessSettingsDropdown',
    'columnSettingsDropdown',
    'quizSettingsDropdown',
    'playerGuideDropdown',
    'tierGuideDropdown',
    'songGuideDropdown',
    'guessGuideDropdown',
    'searchGuideDropdown',
    'quizGuideDropdown',
    'archiveDropdown'
].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener("click", stopProp);
});

const formatAndSortSongsList = (list, prefixBullets = true) => {return list
    .sort((a, b) => {
        const cleanA = (a.startsWith('✓') || a.startsWith('✗')) ? a.slice(2) : a;
        const cleanB = (b.startsWith('✓') || b.startsWith('✗')) ? b.slice(2) : b;
        return cleanA.toLowerCase().localeCompare(cleanB.toLowerCase());
    })
    .map(s => (s.startsWith('✓') || s.startsWith('✗') || !prefixBullets) ? s : `• ${s}`);
};

const sampleLargeSongList = (displaySongs) => {
    const ticks     = displaySongs.filter(s => s.startsWith('✓')).sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
    const crosses   = displaySongs.filter(s => s.startsWith('✗')).sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));

    return [...ticks, ...crosses];
};

window.translateHoverText = function(textArray) {
    if (!globalSearchData || globalSearchData.length === 0) return textArray;

    return textArray.map(line => {
        if (typeof line !== 'string') return line;
        let translatedLine = line;

        for (let i = 0; i < globalSearchData.length; i++) {
            const s     = globalSearchData[i];
            const jp    = s.romaji  || "";
            const en    = s.english || "";

            if (!jp || !en || jp === en) continue;

            if (currentSearchLang === "EN") {
                if (translatedLine.includes(jp + " (OP") || translatedLine.includes(jp + " (ED") || translatedLine.includes(jp + " (IN")) {
                    translatedLine = translatedLine.replace(jp + " (", en + " (");
                    break;
                }
            }

            else {
                if (translatedLine.includes(en + " (OP") || translatedLine.includes(en + " (ED") || translatedLine.includes(en + " (IN")) {
                    translatedLine = translatedLine.replace(en + " (", jp + " (");
                    break;
                }
            }
        }

        return translatedLine;
    });
};

function setupTooltipListeners() {
    const tooltipNode   = document.getElementById('customJsTooltip');
    let hideTimeout     = null;

    function positionTooltip(e) {
        if (tooltipNode.classList.contains('is-hovered')) return;
        tooltipNode.style.display = 'block';

        const tooltipWidth  = tooltipNode.offsetWidth; 
        const tooltipHeight = tooltipNode.offsetHeight;

        let xPos = e.pageX + 15;
        let yPos = e.pageY + 15;

        if (e.clientX + 15 + tooltipWidth   > window.innerWidth)    xPos = e.pageX - tooltipWidth   - 15;
        if (e.clientY + 15 + tooltipHeight  > window.innerHeight)   yPos = e.pageY - tooltipHeight  - 15;

        if (xPos < window.scrollX) xPos = window.scrollX + 5;
        if (yPos < window.scrollY) yPos = window.scrollY + 5;

        tooltipNode.style.left  = xPos + 'px';
        tooltipNode.style.top   = yPos + 'px';
    }

    function clearHideTimeout() {
        if (hideTimeout) {
            clearTimeout(hideTimeout);
            hideTimeout = null;
        }
    }

    function requestHideTooltip() {
        clearHideTimeout();

        hideTimeout = setTimeout(() => {
            if (!tooltipNode.classList.contains('is-hovered')) {
                tooltipNode.style.display           = 'none';
                tooltipNode.style.backgroundColor   = 'black';
                tooltipNode.style.color             = 'white';
                tooltipNode.style.maxHeight         = '';
                tooltipNode.style.overflowY         = '';
            }
        }, 100);
    }

    if (tooltipNode && !tooltipNode._bound) {
        tooltipNode._bound = true;

        tooltipNode.addEventListener('mouseenter', () => {
            clearHideTimeout();
            tooltipNode.classList.add('is-hovered');
        });

        tooltipNode.addEventListener('mouseleave', () => {
            tooltipNode.classList.remove('is-hovered');
            requestHideTooltip();
        });

        window.addEventListener('wheel', (e) => {
            if (tooltipNode.style.display === 'block') {
                const rect          = tooltipNode.getBoundingClientRect();
                const isOverTooltip = (
                    e.clientX >= rect.left && e.clientX <= rect.right &&
                    e.clientY >= rect.top && e.clientY <= rect.bottom
                );

                if (tooltipNode.scrollHeight > tooltipNode.clientHeight) {
                    e.preventDefault();
                    tooltipNode.scrollTop += e.deltaY;
                }
            }
        }, {passive: false});
    }

    document.querySelectorAll('[data-metric]').forEach(th => {
        const metricKey = th.getAttribute('data-metric');
        if (!colExplanations[metricKey]) return;

        th.removeEventListener('mouseenter',    th._handlerEnter);
        th.removeEventListener('mousemove',     positionTooltip);
        th.removeEventListener('mouseleave',    th._handlerLeave);

        th._handlerEnter = (e)  => { clearHideTimeout(); tooltipNode.innerHTML = colExplanations[metricKey]; positionTooltip(e); };
        th._handlerLeave = ()   => { requestHideTooltip(); };

        th.addEventListener('mouseenter', th._handlerEnter);
        th.addEventListener('mousemove',  positionTooltip);
        th.addEventListener('mouseleave', th._handlerLeave);
    });

    document.querySelectorAll('td[data-songs]').forEach(td => {
        td.addEventListener('mouseenter', (e) => {
            try {
                clearHideTimeout();

                const songs = JSON.parse(decodeURIComponent(td.getAttribute('data-songs')));
                if (!songs || songs.length === 0) return;

                if      (td.classList.contains('highlight-best'))   {tooltipNode.style.backgroundColor = c2;        tooltipNode.style.color = 'white';}
                else if (td.classList.contains('highlight-worst'))  {tooltipNode.style.backgroundColor = c0;        tooltipNode.style.color = 'white';}
                else                                                {tooltipNode.style.backgroundColor = 'black';   tooltipNode.style.color = 'white';}

                const metricName = td.getAttribute('data-metric');

                if (metricName === "Off GR") {
                    let cleanSongs      = [...songs];
                    let fractionHeader  = "";

                    if (cleanSongs.length > 0 && (/^\d+\/\d+$/.test(cleanSongs[0]) || /^\d+-\d+-\d+$/.test(cleanSongs[0]))) {
                        fractionHeader = `<b>${cleanSongs[0]}</b>`;
                        cleanSongs.shift();
                    }

                    else {
                        let total           = cleanSongs.length;
                        let correctCount    = cleanSongs.filter(s => s.startsWith('✓')).length;
                        fractionHeader      = `<b>${correctCount}/${total}</b>`;
                    }

                    let correctSongs    = cleanSongs.filter(s => s.startsWith('✓')).map(s => s.slice(2));
                    correctSongs        = window.translateHoverText(correctSongs);
                    let displaySongs    = formatAndSortSongsList(correctSongs, true);

                    tooltipNode.style.maxHeight = '300px';
                    tooltipNode.style.overflowY = 'auto';
                    tooltipNode.innerHTML       = `${fractionHeader}<br>${displaySongs.join('<br>')}`;

                    positionTooltip(e);
                    return;
                }

                let displaySongs        = [...songs];
                displaySongs            = window.translateHoverText(displaySongs);
                const isPlayerSubHover  = td.parentNode.firstElementChild === td;

                if (songs.length === 1 && !songs[0].startsWith('✓') && !songs[0].startsWith('✗') && songs[0].includes('/')) {
                    tooltipNode.innerHTML = songs[0];
                    positionTooltip(e);
                    return;
                }

                if (songs.some(s => s.startsWith("Minimum:"))) {
                    const metricName = td.getAttribute('data-player-metric') || (typeof activePlayerHeadersConfig !== 'undefined' && activePlayerHeadersConfig[td.cellIndex] ? activePlayerHeadersConfig[td.cellIndex].name : "");

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

                const isTeamWinRecordColumn = td.getAttribute('data-metric') === "Win Record" || /^\d+-\d+-\d+$/.test(songs[0]);
                const fractionRegex         = /^\d+\/\d+$/;
                const containsRegex         = fractionRegex.test(songs[0]);
                let fractionHeader          = "";

                if (isTeamWinRecordColumn) {
                    const hasHeader = /^\d+-\d+-\d+$/.test(songs[0]);
                    if (hasHeader) displaySongs.shift();

                    tooltipNode.style.maxHeight = '300px';
                    tooltipNode.style.overflowY = 'auto';
                    tooltipNode.innerHTML       = displaySongs.join('<br>');

                    positionTooltip(e);
                    return;
                }

                if (containsRegex) {
                    fractionHeader = `<b>${songs[0]}</b>`;
                    displaySongs.shift();
                }

                if (containsRegex)  displaySongs = sampleLargeSongList(displaySongs).map(s => (s.startsWith('✓') || s.startsWith('✗') || !isPlayerSubHover) ? s : `• ${s}`);
                else                displaySongs = formatAndSortSongsList(displaySongs, !isPlayerSubHover);

                tooltipNode.style.maxHeight = '300px';
                tooltipNode.style.overflowY = 'auto';
                tooltipNode.innerHTML       = containsRegex ? `${fractionHeader}<br>${displaySongs.join('<br>')}` : displaySongs.join('<br>');

                positionTooltip(e);
            }

            catch (err) {}
        });

        td.addEventListener('mousemove',    positionTooltip);
        td.addEventListener('mouseleave',   requestHideTooltip);
    });

    document.querySelectorAll('[data-quiz-hover]').forEach(el => {
        el.removeEventListener('mouseenter',    el._quizEnter);
        el.removeEventListener('mousemove',     el._quizMove);
        el.removeEventListener('mouseleave',    el._quizLeave);

        el._quizEnter = (e) => {
            clearHideTimeout();
            if (el.id !== "quizResolutionLabel" && el.scrollWidth <= el.clientWidth) return;

            const rawText = decodeURIComponent(el.getAttribute('data-quiz-hover'));
            if (!rawText) return;

            tooltipNode.style.backgroundColor   = 'black';
            tooltipNode.style.color             = 'white';
            tooltipNode.style.border            = 'none';
            tooltipNode.innerHTML               = `<b>${rawText}</b>`;

            positionTooltip(e);
        };

        el._quizMove = (e) => {
            if (el.id !== "quizResolutionLabel" && el.scrollWidth <= el.clientWidth) return;
            positionTooltip(e);
        };

        el._quizLeave = () => { 
            requestHideTooltip();

            tooltipNode.style.display   = 'none';
            tooltipNode.innerHTML       = '';
        };

        el.addEventListener('mouseenter', el._quizEnter);
        el.addEventListener('mousemove',  el._quizMove);
        el.addEventListener('mouseleave', el._quizLeave);
    });
}