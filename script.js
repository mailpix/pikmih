// State
let rawExcelData = []; // Full parsed sheet
let headers = [];
let allDataInColumn = []; // Unfiltered data from column
let participants = []; // Current active participants pool (array of objects {name, row})
let winners = []; // Array of objects {name, row, time}

// Audio FX
const winSound = new Audio('https://assets.mixkit.co/active_storage/sfx/2015/2015-preview.mp3');

// Sintesis suara ticking dinamis "ting" bersih
const audioCtx = new (window.AudioContext || window.webkitAudioContext)();

function playTick() {
    if (audioCtx.state === 'suspended') audioCtx.resume();
    const osc = audioCtx.createOscillator();
    const gainNode = audioCtx.createGain();

    osc.type = 'sine';
    osc.frequency.setValueAtTime(800, audioCtx.currentTime);
    osc.frequency.exponentialRampToValueAtTime(40, audioCtx.currentTime + 0.05);

    gainNode.gain.setValueAtTime(1.0, audioCtx.currentTime);
    gainNode.gain.exponentialRampToValueAtTime(0.001, audioCtx.currentTime + 0.05);

    osc.connect(gainNode);
    gainNode.connect(audioCtx.destination);
    osc.start();
    osc.stop(audioCtx.currentTime + 0.05);
}

function playSpinWheelTicks(durationSec) {
    let startTime = Date.now();
    let durationMs = durationSec * 1000;

    function nextTick() {
        let elapsed = Date.now() - startTime;
        if (elapsed > durationMs) return;

        playTick();

        let progress = elapsed / durationMs;
        // Sinkronisasi kurva suara: mencapai top speed di 30% awal, lalu melambat panjang di 70% sisa waktu
        let normalized;
        if (progress < 0.3) {
            normalized = 1 - (progress / 0.3); // dari 1 ke 0 (makin cepat)
        } else {
            normalized = (progress - 0.3) / 0.7; // dari 0 ke 1 (makin lambat perlahan)
        }
        let currentDelay = 50 + Math.pow(normalized, 3) * 500;

        setTimeout(nextTick, currentDelay);
    }

    nextTick();
}

// DOM Elements
const sidebar = document.getElementById('sidebar');
const overlay = document.getElementById('sidebar-overlay');
const fileUpload = document.getElementById('file-upload');
const columnSelectGroup = document.getElementById('column-select-group');
const columnSelect = document.getElementById('column-select');
const applyDataBtn = document.getElementById('apply-data-btn');
const pesertaTableBody = document.querySelector('#peserta-table tbody');
const totalPesertaSpan = document.getElementById('total-peserta');
const resetDataBtn = document.getElementById('reset-data-btn');
const winnerCountInput = document.getElementById('winner-count');
const slotsContainer = document.getElementById('slots-container');
const spinBtn = document.getElementById('spin-btn');
const pemenangTableBody = document.querySelector('#pemenang-table tbody');
const clearWinnersBtn = document.getElementById('clear-winners-btn');
const exportWinnersBtn = document.getElementById('export-winners-btn');
const rangeStart = document.getElementById('range-start');
const rangeEnd = document.getElementById('range-end');

// Init
document.addEventListener('DOMContentLoaded', () => {
    loadState();
    setupEventListeners();
    renderSlots(parseInt(winnerCountInput.value));
});

function setupEventListeners() {
    document.getElementById('hamburger-btn').addEventListener('click', () => toggleSidebar(true));
    document.getElementById('close-sidebar-btn').addEventListener('click', () => toggleSidebar(false));
    overlay.addEventListener('click', () => toggleSidebar(false));

    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            e.target.classList.add('active');
            document.getElementById(e.target.dataset.target).classList.add('active');
        });
    });

    fileUpload.addEventListener('change', handleFileUpload);
    columnSelect.addEventListener('change', processSelectedColumn);
    applyDataBtn.addEventListener('click', () => { applyFilter(); alert("Filter berhasil diterapkan!"); });

    resetDataBtn.addEventListener('click', () => {
        if (confirm("Yakin ingin mereset seluruh data peserta?")) {
            participants = [];
            allDataInColumn = [];
            rawExcelData = [];
            headers = [];
            fileUpload.value = '';
            columnSelectGroup.classList.add('hidden');
            saveState();
            renderPeserta();
            checkReady();
            renderSlots(parseInt(winnerCountInput.value));
        }
    });

    winnerCountInput.addEventListener('change', (e) => {
        let val = parseInt(e.target.value);
        if (isNaN(val) || val < 1) val = 1;
        if (val > 6) val = 6;
        e.target.value = val;
        if (!isSpinning) renderSlots(val);
    });

    pesertaTableBody.addEventListener('click', (e) => {
        if (e.target.classList.contains('btn-delete')) {
            const index = parseInt(e.target.dataset.index);
            participants.splice(index, 1);
            saveState();
            renderPeserta();
            checkReady();
        }
    });

    spinBtn.addEventListener('click', startSpin);

    clearWinnersBtn.addEventListener('click', () => {
        if (confirm("Hapus semua hasil pemenang?")) {
            winners = [];
            saveState();
            renderPemenang();
        }
    });

    exportWinnersBtn.addEventListener('click', exportWinners);
}

function toggleSidebar(show) {
    if (show) {
        sidebar.classList.add('active');
        overlay.classList.add('active');
    } else {
        sidebar.classList.remove('active');
        overlay.classList.remove('active');
    }
}

// Data Handling
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        rawExcelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (rawExcelData.length > 0) {
            headers = rawExcelData[0] || [];
            columnSelect.innerHTML = '';
            headers.forEach((h, i) => {
                const opt = document.createElement('option');
                opt.value = i;
                opt.textContent = h || `Kolom ${i + 1}`;
                columnSelect.appendChild(opt);
            });
            columnSelectGroup.classList.remove('hidden');

            processSelectedColumn();
        }
    };
    reader.readAsArrayBuffer(file);
}

function processSelectedColumn() {
    if (rawExcelData.length === 0) return;
    const colIndex = parseInt(columnSelect.value);

    allDataInColumn = [];
    for (let i = 1; i < rawExcelData.length; i++) {
        if (rawExcelData[i] && rawExcelData[i][colIndex] !== undefined && rawExcelData[i][colIndex] !== null) {
            let name = rawExcelData[i][colIndex].toString().trim();
            if (name) {
                allDataInColumn.push(name);
            }
        }
    }

    rangeStart.value = 1;
    rangeEnd.value = allDataInColumn.length;

    applyFilter();
}

function applyFilter() {
    const start = parseInt(rangeStart.value) - 1;
    const end = parseInt(rangeEnd.value) - 1;

    participants = [];
    let count = 0;
    for (let i = 0; i < allDataInColumn.length; i++) {
        if (i >= start && i <= end) {
            let name = allDataInColumn[i];
            if (name) {
                participants.push({ name: name, row: i + 2 });
                count++;
            }
        }
    }

    saveState();
    renderPeserta();
    checkReady();
    if (!isSpinning) renderSlots(parseInt(winnerCountInput.value));
}

function renderPeserta() {
    totalPesertaSpan.textContent = participants.length;
    pesertaTableBody.innerHTML = '';

    const displayCount = Math.min(participants.length, 100);

    for (let i = 0; i < displayCount; i++) {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${i + 1}</td>
            <td>${participants[i].name} <small style="color:#888">(Baris ${participants[i].row})</small></td>
            <td><button class="btn-danger btn-sm btn-delete" data-index="${i}">X</button></td>
        `;
        pesertaTableBody.appendChild(tr);
    }

    if (participants.length > 100) {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td colspan="3" style="text-align:center; color:var(--text-muted)">... dan ${participants.length - 100} lainnya ...</td>`;
        pesertaTableBody.appendChild(tr);
    }
}

function renderPemenang() {
    pemenangTableBody.innerHTML = '';
    winners.forEach((w, i) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${i + 1}</td>
            <td>${w.name}</td>
            <td>${w.time || '-'}</td>
        `;
        pemenangTableBody.appendChild(tr);
    });
}

function checkReady() {
    if (participants.length > 0) {
        spinBtn.disabled = false;
    } else {
        spinBtn.disabled = true;
    }
}

function exportWinners() {
    if (winners.length === 0) return alert("Belum ada pemenang!");
    const ws_data = [["No", "Nama Pemenang", "Baris Asal Excel", "Waktu Menang"]];
    winners.forEach((w, i) => {
        ws_data.push([i + 1, w.name, w.row, w.time || '-']);
    });
    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pemenang");
    XLSX.writeFile(wb, "Data_Pemenang_Undian.xlsx");
}

// Storage
function saveState() {
    localStorage.setItem('slot_participants', JSON.stringify(participants));
    localStorage.setItem('slot_winners', JSON.stringify(winners));
}

function loadState() {
    const savedP = localStorage.getItem('slot_participants');
    const savedW = localStorage.getItem('slot_winners');
    if (savedP) participants = JSON.parse(savedP);
    if (savedW) winners = JSON.parse(savedW);
    renderPeserta();
    renderPemenang();
    checkReady();
}

// Slot Logic
let isSpinning = false;
let currentWinnerCount = 1;

function renderSlots(count) {
    currentWinnerCount = count;
    slotsContainer.innerHTML = '';
    slotsContainer.className = 'slots-container count-' + count;

    for (let i = 0; i < count; i++) {
        const box = document.createElement('div');
        box.className = 'slot-box';

        const inner = document.createElement('div');
        inner.className = 'slot-inner';
        inner.id = `slot-inner-${i}`;

        if (participants.length > 0) {
            // Idle Animation Placeholder
            const idleItems = 10;
            let idleList = [];
            for (let j = 0; j < idleItems; j++) {
                idleList.push(participants[Math.floor(Math.random() * participants.length)]);
            }
            const fullList = [...idleList, ...idleList]; // Double for seamless loop

            fullList.forEach(obj => {
                const item = document.createElement('div');
                item.className = 'slot-item';
                item.innerHTML = `<span class="slot-item-text">${obj.name}</span>`;
                inner.appendChild(item);
            });
            inner.classList.add('idle-anim');
        } else {
            // Empty state
            const item = document.createElement('div');
            item.className = 'slot-item';
            item.innerHTML = `<span class="slot-item-text">? ? ?</span>`;
            inner.appendChild(item);
        }

        box.appendChild(inner);
        slotsContainer.appendChild(box);
    }
}

async function startSpin() {
    if (isSpinning || participants.length === 0) return;

    let wCount = parseInt(winnerCountInput.value);
    if (wCount > participants.length) {
        alert("Jumlah pemenang melebihi sisa peserta! Akan mengundi sesuai sisa peserta.");
        wCount = participants.length;
        winnerCountInput.value = wCount;
        renderSlots(wCount);
    }

    isSpinning = true;
    spinBtn.disabled = true;
    toggleSidebar(false);

    document.querySelectorAll('.slot-box').forEach(b => b.classList.remove('winner'));

    let roundWinners = [];
    let roundWinnerIndices = [];

    for (let i = 0; i < wCount; i++) {
        let rIndex;
        do {
            rIndex = Math.floor(Math.random() * participants.length);
        } while (roundWinnerIndices.includes(rIndex));

        roundWinnerIndices.push(rIndex);
        roundWinners.push({ obj: participants[rIndex], index: rIndex });
    }

    const spinItemsCount = 80;

    for (let i = 0; i < wCount; i++) {
        const inner = document.getElementById(`slot-inner-${i}`);
        inner.classList.remove('idle-anim');
        inner.innerHTML = '';
        inner.style.transition = 'none';
        inner.style.transform = 'translateY(0)';

        for (let j = 0; j < spinItemsCount - 1; j++) {
            const item = document.createElement('div');
            item.className = 'slot-item';
            let obj = participants[Math.floor(Math.random() * participants.length)];
            item.innerHTML = `<span class="slot-item-text">${obj.name}</span>`;
            inner.appendChild(item);
        }

        const winningItem = document.createElement('div');
        winningItem.className = 'slot-item';
        winningItem.innerHTML = `<span class="slot-item-text">${roundWinners[i].obj.name}</span>`;
        inner.appendChild(winningItem);
    }

    void slotsContainer.offsetWidth;

    const promises = [];
    for (let i = 0; i < wCount; i++) {
        promises.push(new Promise(resolve => {
            // Jeda 2 detik antar slot mulai berputar
            const startDelay = i * 2000;

            setTimeout(() => {
                const inner = document.getElementById(`slot-inner-${i}`);
                inner.parentElement.classList.add('spinning');

                const targetPercentage = -((spinItemsCount - 1) / spinItemsCount) * 100;

                // Durasi masing-masing putaran dibuat seragam 5 detik
                const duration = 7;

                // Kurva khusus: melesat cepat di awal, dan pengereman sangat panjang di akhir
                inner.style.transition = `transform ${duration}s cubic-bezier(0.3, 0, 0.1, 1)`;
                inner.style.transform = `translateY(${targetPercentage}%)`;

                // Play dynamic ticking sound SAAT slot ini mulai berputar
                playSpinWheelTicks(duration);

                setTimeout(() => {
                    inner.parentElement.classList.remove('spinning');
                    inner.parentElement.classList.add('winner');

                    let wSound = winSound.cloneNode();
                    wSound.volume = 0.2; // Ini volumenya diturunkan ke 10%
                    wSound.play().catch(e => console.log('Audio error:', e));

                    fireConfetti();
                    resolve();
                }, duration * 1000);

            }, startDelay);
        }));
    }

    await Promise.all(promises);

    roundWinnerIndices.sort((a, b) => b - a);

    const timeNow = new Date().toLocaleString('id-ID');

    roundWinners.forEach(rw => {
        winners.push({
            name: rw.obj.name,
            row: rw.obj.row,
            time: timeNow
        });
    });

    roundWinnerIndices.forEach(idx => {
        participants.splice(idx, 1);
    });

    saveState();
    renderPeserta();
    renderPemenang();

    isSpinning = false;
    checkReady();
}

function fireConfetti() {
    const duration = 2 * 1000;
    const end = Date.now() + duration;

    (function frame() {
        confetti({
            particleCount: 5,
            angle: 60,
            spread: 55,
            origin: { x: 0 },
            colors: ['#FFD700', '#FFFFFF', '#00A2E9']
        });
        confetti({
            particleCount: 5,
            angle: 120,
            spread: 55,
            origin: { x: 1 },
            colors: ['#FFD700', '#FFFFFF', '#00A2E9']
        });

        if (Date.now() < end) {
            requestAnimationFrame(frame);
        }
    }());
}
