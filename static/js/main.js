function addActivity() {
    const container = document.getElementById('activitiesContainer');
    const row = document.createElement('div');
    row.className = 'activity-row';
    row.innerHTML = `
        <span class="row-no"></span>
        <textarea name="activity_text[]" rows="2" placeholder="Describe the work done…"></textarea>
        <input type="text" name="activity_delay[]" placeholder="Nil / Describe delay">
        <button type="button" class="btn-remove-row" onclick="removeRow(this)">✕</button>
    `;
    container.appendChild(row);
    updateRowNumbers();
    row.querySelector('textarea').focus();
}

function addDelivery() {
    const container = document.getElementById('deliveriesContainer');
    const row = document.createElement('div');
    row.className = 'list-row';
    row.innerHTML = `
        <span class="row-no"></span>
        <input type="text" name="delivery[]" placeholder="Describe delivery…">
        <button type="button" class="btn-remove-row" onclick="removeRow(this)">✕</button>
    `;
    container.appendChild(row);
    updateRowNumbers();
    row.querySelector('input').focus();
}

function addCollection() {
    const container = document.getElementById('collectionsContainer');
    const row = document.createElement('div');
    row.className = 'list-row';
    row.innerHTML = `
        <span class="row-no"></span>
        <input type="text" name="collection[]" placeholder="Describe collection…">
        <button type="button" class="btn-remove-row" onclick="removeRow(this)">✕</button>
    `;
    container.appendChild(row);
    updateRowNumbers();
    row.querySelector('input').focus();
}

function addCustomLabour() {
    const container = document.getElementById('customLabourContainer');
    const row = document.createElement('div');
    row.className = 'custom-resource-row';
    row.innerHTML = `
        <input type="text" name="custom_labour_name[]" placeholder="Category name…">
        <input type="number" name="custom_labour_count[]" min="0" value="0" class="count-input">
        <button type="button" class="btn-remove-row" onclick="removeRow(this)">✕</button>
    `;
    container.appendChild(row);
    row.querySelector('input').focus();
}

function addCustomPlant() {
    const container = document.getElementById('customPlantContainer');
    const row = document.createElement('div');
    row.className = 'custom-resource-row';
    row.innerHTML = `
        <input type="text" name="custom_plant_name[]" placeholder="Equipment name…">
        <input type="number" name="custom_plant_count[]" min="0" value="0" class="count-input">
        <button type="button" class="btn-remove-row" onclick="removeRow(this)">✕</button>
    `;
    container.appendChild(row);
    row.querySelector('input').focus();
}

function removeRow(btn) {
    const row = btn.parentElement;
    row.remove();
    updateRowNumbers();
}

function updateRowNumbers() {
    document.querySelectorAll('#activitiesContainer .activity-row').forEach((row, i) => {
        const span = row.querySelector('.row-no');
        if (span) span.textContent = i + 1;
    });
    document.querySelectorAll('#deliveriesContainer .list-row').forEach((row, i) => {
        const span = row.querySelector('.row-no');
        if (span) span.textContent = i + 1;
    });
    document.querySelectorAll('#collectionsContainer .list-row').forEach((row, i) => {
        const span = row.querySelector('.row-no');
        if (span) span.textContent = i + 1;
    });
}

function updateLabourTotal() {
    let total = 0;
    document.querySelectorAll('.count-input').forEach(inp => {
        total += parseInt(inp.value) || 0;
    });
    const el = document.getElementById('labourTotalVal');
    if (el) el.textContent = total;
}

const DAYS = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];

function updateDayOfWeek() {
    const dateInput = document.getElementById('entryDate');
    const dayDisplay = document.getElementById('dayDisplay');
    if (!dateInput || !dayDisplay) return;
    const val = dateInput.value;
    if (val) {
        const [y, m, d] = val.split('-').map(Number);
        const dt = new Date(y, m - 1, d);
        dayDisplay.value = DAYS[dt.getDay()];
    } else {
        dayDisplay.value = '';
    }
}

document.addEventListener('DOMContentLoaded', () => {
    updateRowNumbers();
    updateLabourTotal();
    updateDayOfWeek();

    const dateInput = document.getElementById('entryDate');
    if (dateInput) dateInput.addEventListener('change', updateDayOfWeek);

    document.querySelectorAll('.count-input').forEach(inp => {
        inp.addEventListener('input', updateLabourTotal);
    });

    const form = document.getElementById('diaryForm');
    if (form) {
        form.addEventListener('input', (e) => {
            if (e.target.classList.contains('count-input')) updateLabourTotal();
        });
    }
});
