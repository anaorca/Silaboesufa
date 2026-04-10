const courseSelect = document.getElementById('courseSelect');
const moduleNameInput = document.getElementById('moduleName');
const moduleCodeInput = document.getElementById('moduleCode');
const currentModuleTitle = document.getElementById('currentModuleTitle');
const selectionInstructionArea = document.getElementById('selectionInstructionArea');
const totalCreditsEl = document.getElementById('totalCredits');
const creditsInput = document.getElementById('credits');
const totalHoursInput = document.getElementById('totalHours');
const syncHoursInput = document.getElementById('syncHours');
const asyncHoursInput = document.getElementById('asyncHours');
const hoursValidation = document.getElementById('hoursValidation');
const modeVirtual = document.getElementById('modeVirtual');
const modePresencial = document.getElementById('modePresencial');
const objectiveTextarea = document.getElementById('courseObjective');
const rapsListContainer = document.getElementById('rapsList');
const unitsContainer = document.getElementById('unitsContainer');
const btnAddEval = document.getElementById('btnAddEval');
const btnHistory = document.getElementById('btnHistory');
const btnSettings = document.getElementById('btnSettings');
const btnExportPDF = document.getElementById('btnExportPDF');

// Persistence State
let syllabusDb = [];
let currentCourseIndex = -1;
let editingRAPIndex = -1;
let history = JSON.parse(localStorage.getItem('syllabus_history') || '[]');
async function init() {
    try {
        const response = await fetch('tga_syllabus_db.json');
        syllabusDb = await response.json();
        
        // Clean up data if necessary
        syllabusDb = syllabusDb.map((course, idx) => ({
            ...course,
            id: idx, // Ensure sequential IDs for the selector
            sigla: course.sigla || 'N/A'
        }));

        populateDropdown();
        setupEventListeners();
        loadSettings(); // Load saved AI config
        
        // --- LÓGICA DE ADMINISTRADOR ---
        // Solo muestra el botón de configuración si la URL tiene ?admin=true
        const urlParams = new URLSearchParams(window.location.search);
        if (urlParams.get('admin') === 'true') {
            if (btnSettings) btnSettings.style.display = 'inline-flex';
            console.log("✨ Modo Administrador Activado");
        }
    } catch (error) {
        console.error("Error loading DB:", error);
    }
}

function populateDropdown() {
    // Clear existing options except the first two
    while (courseSelect.options.length > 2) {
        courseSelect.remove(2);
    }
    
    syllabusDb.forEach(course => {
        const option = document.createElement('option');
        option.value = course.id;
        option.textContent = `${course.modulo} (${course.sigla})`;
        courseSelect.appendChild(option);
    });
}

function setupEventListeners() {
    // Course Selection via Dropdown
    courseSelect.addEventListener('change', (e) => {
        const val = e.target.value;
        if (val === "NEW") {
            handleNewCourse();
        } else if (val !== "") {
            const course = syllabusDb.find(c => c.id == val);
            if (course) selectCourse(course);
        }
    });

    // New Course Button (Header)
    btnNewCourse.addEventListener('click', () => {
        courseSelect.value = "NEW";
        handleNewCourse();
    });

    // Modality Changes
    [modeVirtual, modePresencial].forEach(radio => {
        radio.addEventListener('change', () => {
            if (currentCourseIndex !== -1) {
                const course = syllabusDb.find(c => c.id == currentCourseIndex);
                updateHoursBasedOnModality(course);
            } else {
                updateHoursBasedOnModality(null);
            }
        });
    });

    // Credits / Hours Logic
    creditsInput.addEventListener('input', () => {
        updateTotalHours();
        const course = currentCourseIndex !== -1 ? syllabusDb.find(c => c.id == currentCourseIndex) : null;
        if (course) course.creditos = parseInt(creditsInput.value) || 0;
        updateHoursBasedOnModality(course);
        validateHours();
    });

    moduleNameInput.addEventListener('input', () => {
        const course = syllabusDb.find(c => c.id == currentCourseIndex);
        if (course) course.modulo = moduleNameInput.value;
    });

    moduleCodeInput.addEventListener('input', () => {
        const course = syllabusDb.find(c => c.id == currentCourseIndex);
        if (course) course.sigla = moduleCodeInput.value;
    });

    syncHoursInput.addEventListener('input', () => {
        const course = syllabusDb.find(c => c.id == currentCourseIndex);
        if (course) course.horas_sincronicas = parseInt(syncHoursInput.value) || 0;
        validateHours();
    });
    
    asyncHoursInput.addEventListener('input', () => {
        const course = syllabusDb.find(c => c.id == currentCourseIndex);
        if (course) course.horas_asincronicas = parseInt(asyncHoursInput.value) || 0;
        validateHours();
    });

    // Add RAP
    document.getElementById('btnAddRAP').addEventListener('click', () => {
        openRAPModal(-1);
    });

    // Add Unit
    if (btnAddUnit) btnAddUnit.addEventListener('click', addUnit);
    const btnAddUnitBottom = document.getElementById('btnAddUnitBottom');
    if (btnAddUnitBottom) btnAddUnitBottom.addEventListener('click', addUnit);
    if (btnAddEval) btnAddEval.addEventListener('click', addEvaluation);

    // Modal Actions
    const btnCancelRAP = document.getElementById('btnCancelRAP');
    if (btnCancelRAP) btnCancelRAP.addEventListener('click', closeRAPModal);
    
    const btnConfirmRAP = document.getElementById('btnConfirmRAP');
    if (btnConfirmRAP) btnConfirmRAP.addEventListener('click', saveRAP);

    const btnCorrectRAPIA = document.getElementById('btnCorrectRAPIA');
    if (btnCorrectRAPIA) btnCorrectRAPIA.addEventListener('click', optimizeRAPWithIA);

    // Objective Sync
    objectiveTextarea.addEventListener('input', () => {
        const course = syllabusDb.find(c => c.id == currentCourseIndex);
        if (course) course.objetivo_general = objectiveTextarea.value;
    });

    // Save Button
    document.getElementById('btnSave').addEventListener('click', () => {
        saveSyllabusToHistory();
    });

    // IA Assist Buttons
    document.getElementById('btnAIAssistUnits').addEventListener('click', autoAlignSyllabusWithIA);
    document.getElementById('btnAIAssistEval').addEventListener('click', suggestEvaluations);

    // History & Settings Buttons
    if (btnHistory) btnHistory.addEventListener('click', openHistoryModal);
    if (btnSettings) btnSettings.addEventListener('click', openSettingsModal);
    if (btnExportPDF) btnExportPDF.addEventListener('click', exportToPDF);
    const btnExportWord = document.getElementById('btnExportWord');
    if (btnExportWord) btnExportWord.addEventListener('click', exportToWord);
}

function handleNewCourse() {
    clearForm();
    
    // Create new temporary entry in DB to allow interaction
    const newId = syllabusDb.length;
    const newCourse = {
        id: newId,
        modulo: "Nueva Asignatura",
        sigla: "",
        creditos: 1,
        unidades: [],
        evaluaciones: [],
        resultados_aprendizaje: []
    };
    syllabusDb.push(newCourse);
    currentCourseIndex = newId;
    
    applyVirtualRatio();
    if (selectionInstructionArea) selectionInstructionArea.style.display = 'none';
    document.getElementById('currentModuleTitle').textContent = "Nueva Asignatura";
    
    // Refresh UI for new course
    renderUnits([]);
    renderEvaluations([]);
    renderRAPs([]);
}

function selectCourse(course) {
    if (selectionInstructionArea) selectionInstructionArea.style.display = 'none';
    currentCourseIndex = course.id;
    const areaInfo = course.area ? ` [${course.area}]` : '';
    document.getElementById('currentModuleTitle').textContent = `Editando: ${course.modulo}${areaInfo}`;

    // Auto-detect modality
    const name = (course.modulo || '').toLowerCase();
    const sigla = (course.sigla || '').toLowerCase();
    if (name.includes('delta') || name.includes('práctica') || name.includes('practica') || sigla.includes('delta')) {
        modePresencial.checked = true;
    } else {
        modeVirtual.checked = true;
    }

    // Fill Form
    moduleNameInput.value = course.modulo || '';
    moduleCodeInput.value = course.sigla || '';
    creditsInput.value = course.creditos || 1;
    
    updateTotalHours();
    updateHoursBasedOnModality(course);
    validateHours();

    objectiveTextarea.value = course.objetivo_general || '';
    
    // UI Hint if information is missing
    const subtitle = document.querySelector('.subtitle');
    let dbStatus = "";
    if (course.unidades && course.unidades.length > 0) {
        dbStatus = `<br><span style="color: #4db6ac; font-size: 0.9rem;">✅ Estructura presencial oficial cargada (${course.unidades.length} unidades).</span>`;
    }

    if (!course.objetivo_general && (!course.resultados_aprendizaje || course.resultados_aprendizaje.length === 0)) {
        subtitle.innerHTML = `<span style="color: var(--primary-light);">⚠ Sin información previa:</span> Favor completar los objetivos y RAPs.${dbStatus}`;
    } else {
        subtitle.innerHTML = `Seleccione una asignatura existente o cree una nueva.${dbStatus}`;
    }

    // RAPs
    renderRAPs(course.resultados_aprendizaje || []);
    
    // Units & Topics (We only keep the official content structure)
    const virtualizedUnits = (course.unidades || []).map(unit => {
        let tema = "";
        let subtemas = "";
        
        if (Array.isArray(unit.temas)) {
            tema = unit.temas[0] || "";
            subtemas = unit.temas.slice(1).join('\n');
        } else if (typeof unit.temas === 'string') {
            const split = unit.temas.split('\n');
            tema = split[0] || "";
            subtemas = split.slice(1).join('\n');
        }

        return {
            ...unit,
            tema: unit.tema || tema,
            subtemas: unit.subtemas || subtemas,
            actividades: unit.actividades || "", // Clear for virtual redesign if empty
            recursos: unit.recursos || "",    // Clear for virtual redesign if empty
            rap_id: unit.rap_id               // Keep if already aligned
        };
    });
    
    // Keep existing evaluations if coming from history/load, else start fresh for virtual redesign
    if (!course.evaluaciones || course.evaluaciones.length === 0) {
        course.evaluaciones = [];
    }
    
    renderUnits(virtualizedUnits);
    renderEvaluations(course.evaluaciones);
    
    // Update the DB object in memory as well
    course.unidades = virtualizedUnits;
}

function renderUnits(units) {
    if (!unitsContainer) return;
    unitsContainer.innerHTML = '';
    
    const course = currentCourseIndex !== -1 ? syllabusDb.find(c => c.id == currentCourseIndex) : null;
    const raps = course?.resultados_aprendizaje || [];

    if (units.length === 0) {
        unitsContainer.innerHTML = '<div style="padding: 20px; text-align: center; color: var(--text-dim);">No hay unidades definidas. Pulse + Nueva Unidad.</div>';
        return;
    }

    units.forEach((unit, uIdx) => {
        // Automatic Alignment Logic (Background)
        if (unit.rap_id === undefined && raps.length > 0) {
            unit.rap_id = findBestRAPMatch((unit.nombre || "") + " " + (unit.tema || ""), raps);
        }

        const unitRow = document.createElement('div');
        unitRow.className = 'unit-row';
        
        let subtemasText = "";
        if (Array.isArray(unit.subtemas)) {
            subtemasText = unit.subtemas.join('\n');
        } else {
            subtemasText = unit.subtemas || "";
        }

        let linkedRAPText = unit.rap_id !== undefined ? `Vínculo IA: RAP ${unit.rap_id + 1}` : "Analizando Vínculo...";

        unitRow.innerHTML = `
            <div class="unit-col-label">
                <textarea class="unit-textarea-label" placeholder="Nombre de la Unidad..." oninput="updateUnitName(${uIdx}, this.value)">${unit.nombre || ''}</textarea>
                <div class="alignment-status-ia">
                    <span class="icon-sparkle">✨</span> ${linkedRAPText}
                </div>
            </div>
            <div class="unit-col-content">
                <textarea class="topics-textarea" placeholder="Tema principal..." oninput="updateUnitTema(${uIdx}, this.value)">${unit.tema || ''}</textarea>
            </div>
            <div class="unit-col-content">
                <textarea class="topics-textarea" placeholder="Subtemas..." oninput="updateUnitSubtemas(${uIdx}, this.value)">${subtemasText}</textarea>
            </div>
            <div class="unit-col-content">
                <textarea class="topics-textarea" placeholder="Actividades (AVA)..." oninput="updateUnitActivities(${uIdx}, this.value)">${unit.actividades || ''}</textarea>
            </div>
            <div class="unit-col-content">
                <textarea class="topics-textarea" placeholder="Recursos (AVA)..." oninput="updateUnitResources(${uIdx}, this.value)">${unit.recursos || ''}</textarea>
            </div>
            <div class="unit-col-actions">
                <button class="btn-delete-row" onclick="deleteUnit(${uIdx})" title="Eliminar fila">✕</button>
            </div>
        `;
        unitsContainer.appendChild(unitRow);
    });
}

function updateUnitAlignment(uIdx, rapId) {
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    if (course && course.unidades[uIdx]) {
        course.unidades[uIdx].rap_id = rapId === "" ? undefined : parseInt(rapId);
        renderUnits(course.unidades);
    }
}

function updateUnitActivities(uIdx, value) {
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    if (course && course.unidades[uIdx]) {
        course.unidades[uIdx].actividades = value;
    }
}

function updateUnitResources(uIdx, value) {
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    if (course && course.unidades[uIdx]) {
        course.unidades[uIdx].recursos = value;
    }
}

function updateUnitTema(uIdx, value) {
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    if (course && course.unidades[uIdx]) {
        course.unidades[uIdx].tema = value;
    }
}

function updateUnitSubtemas(uIdx, value) {
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    if (course && course.unidades[uIdx]) {
        course.unidades[uIdx].subtemas = value;
    }
}

function updateUnitName(uIdx, value) {
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    if (course && course.unidades[uIdx]) {
        course.unidades[uIdx].nombre = value;
    }
}

function addUnit() {
    if (currentCourseIndex === -1) {
        alert("Seleccione una asignatura primero");
        return;
    }
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    if (!course.unidades) course.unidades = [];
    course.unidades.push({ nombre: "Nueva Unidad", tema: "", subtemas: "", actividades: "", recursos: "" });
    renderUnits(course.unidades);
}

function deleteUnit(uIdx) {
    if (!confirm("¿Eliminar toda la unidad y sus temas?")) return;
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    if (course && course.unidades) {
        course.unidades.splice(uIdx, 1);
        renderUnits(course.unidades);
    }
}

function addEvaluation() {
    if (currentCourseIndex === -1) {
        alert("Seleccione una asignatura primero");
        return;
    }
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    if (!course.evaluaciones) course.evaluaciones = [];
    course.evaluaciones.push({
        actividad: "Nueva Actividad de Evaluación",
        instrumento: "Rúbrica",
        criterios: "",
        instrucciones: ""
    });
    renderEvaluations(course.evaluaciones);
}

function renderEvaluations(evals) {
    evaluationsContainer.innerHTML = '';
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    const raps = course?.resultados_aprendizaje || [];

    evals.forEach((ev, eIdx) => {
        // Automatic Alignment Logic (Internal)
        if (ev.rap_id === undefined && raps.length > 0) {
            ev.rap_id = findBestRAPMatch(ev.actividad + " " + ev.descripcion, raps);
        }

        const evalCard = document.createElement('div');
        evalCard.className = 'evaluation-card';
        
        let detailLabel = "Criterios de Evaluación / Rúbrica";
        if (ev.instrumento === "Lista de Cotejo") detailLabel = "Puntos de Verificación (Checklist)";
        if (ev.instrumento === "Examen / Cuestionario") detailLabel = "Temas a evaluar / Preguntas clave";
        if (ev.instrumento === "Otros") detailLabel = "Defina sus elementos de evaluación";

        let linkedRAPText = ev.rap_id !== undefined ? `Vínculo IA: RAP ${ev.rap_id + 1}` : "Analizando Vínculo...";

        evalCard.innerHTML = `
            <div class="eval-card-header">
                <textarea class="input-eval-title" oninput="updateEvalData(${eIdx}, 'actividad', this.value)" placeholder="Nombre de la Actividad de Evaluación..." rows="2">${ev.actividad}</textarea>
                <div class="eval-card-actions">
                    <div class="alignment-status-ia">
                        <span class="icon-sparkle">✨</span> ${linkedRAPText}
                    </div>
                    <select class="select-instrument" onchange="updateEvalData(${eIdx}, 'instrumento', this.value); renderEvaluations(syllabusDb.find(c => c.id == currentCourseIndex).evaluaciones)">
                        <option value="Rúbrica" ${ev.instrumento === 'Rúbrica' ? 'selected' : ''}>Rúbrica</option>
                        <option value="Lista de Cotejo" ${ev.instrumento === 'Lista de Cotejo' ? 'selected' : ''}>Lista de Cotejo</option>
                        <option value="Examen / Cuestionario" ${ev.instrumento === 'Examen / Cuestionario' ? 'selected' : ''}>Examen / Cuestionario</option>
                        <option value="Guía de Observación" ${ev.instrumento === 'Guía de Observación' ? 'selected' : ''}>Guía de Observación</option>
                        <option value="Mapa Mental / Infografía" ${ev.instrumento === 'Mapa Mental / Infografía' ? 'selected' : ''}>Mapa Mental / Infografía</option>
                        <option value="Portafolio" ${ev.instrumento === 'Portafolio' ? 'selected' : ''}>Portafolio</option>
                        <option value="Otros" ${ev.instrumento === 'Otros' ? 'selected' : ''}>Otros</option>
                    </select>
                    <button class="btn-delete-row" onclick="deleteEvaluation(${eIdx})" title="Eliminar Actividad">✕</button>
                </div>
            </div>
            <div class="eval-card-body">
                <div class="eval-group full-width">
                    <label>Descripción del Reto (Evaluación Auténtica):</label>
                    <textarea oninput="updateEvalData(${eIdx}, 'descripcion', this.value)" placeholder="Describa el reto o problema real que el estudiante debe resolver...">${ev.descripcion || ''}</textarea>
                </div>
                <div class="eval-group">
                    <label>Instrucciones Paso a Paso:</label>
                    <textarea oninput="updateEvalData(${eIdx}, 'instrucciones', this.value)" placeholder="Guía clara para el estudiante...">${ev.instrucciones || ''}</textarea>
                </div>
                <div class="eval-group">
                    <label>${detailLabel}:</label>
                    <textarea oninput="updateEvalData(${eIdx}, 'criterios', this.value)" placeholder="Verbo en infinitivo + objeto + condición...">${ev.criterios || ''}</textarea>
                </div>
            </div>
        `;
        evaluationsContainer.appendChild(evalCard);
    });
}

function updateEvalData(eIdx, field, value) {
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    if (course && course.evaluaciones[eIdx]) {
        course.evaluaciones[eIdx][field] = value;
    }
}

function deleteEvaluation(eIdx) {
    if (!confirm("¿Eliminar toda esta actividad de evaluación?")) return;
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    if (course && course.evaluaciones) {
        course.evaluaciones.splice(eIdx, 1);
        renderEvaluations(course.evaluaciones);
    }
}

function renderRAPs(raps) {
    rapsListContainer.innerHTML = '';
    raps.forEach((rap, index) => {
        const li = document.createElement('li');
        li.className = 'rap-item';
        li.innerHTML = `
            <div class="rap-text">RAP ${index + 1}: ${rap}</div>
            <div class="rap-actions">
                <button class="btn btn-secondary btn-small" onclick="openRAPModal(${index})">✎</button>
                <button class="btn btn-danger btn-small" onclick="deleteRAP(${index})">✕</button>
            </div>
        `;
        rapsListContainer.appendChild(li);
    });
}

function clearForm() {
    moduleNameInput.value = '';
    moduleCodeInput.value = '';
    creditsInput.value = 1;
    syncHoursInput.value = 0;
    asyncHoursInput.value = 0;
    updateTotalHours();
    validateHours();
    objectiveTextarea.value = '';
    rapsListContainer.innerHTML = '';
    unitsContainer.innerHTML = '';
}

function updateTotalHours() {
    const credits = parseInt(creditsInput.value) || 0;
    totalHoursInput.value = credits * 48;
}

function updateHoursBasedOnModality(course) {
    if (modeVirtual.checked) {
        applyVirtualRatio();
    } else {
        // Presencial: Use original values if available, else default to all sync
        if (course) {
            syncHoursInput.value = course.horas_sincronicas || course.horas_ad || 0;
            asyncHoursInput.value = course.horas_asincronicas || course.horas_ti || 0;
        } else {
            syncHoursInput.value = parseInt(totalHoursInput.value) || 0;
            asyncHoursInput.value = 0;
        }
    }
}

function applyVirtualRatio() {
    const total = parseInt(totalHoursInput.value) || 0;
    // Ratio 1 sync : 2 async -> Total = 3 units
    const sync = Math.floor(total / 3);
    const async = total - sync;
    
    syncHoursInput.value = sync;
    asyncHoursInput.value = async;
}

function validateHours() {
    const totalRequired = parseInt(totalHoursInput.value) || 0;
    const sync = parseInt(syncHoursInput.value) || 0;
    const async = parseInt(asyncHoursInput.value) || 0;
    const sum = sync + async;

    if (sum === totalRequired && totalRequired > 0) {
        if (modeVirtual.checked) {
            // Check ratio 1:2
            const isRatioCorrect = (async === sync * 2) || (async === (sync * 2) + 1) || (async === (sync * 2) - 1);
            if (isRatioCorrect) {
                hoursValidation.textContent = "✓ Distribución virtual correcta (Relación 1:2).";
                hoursValidation.className = "validation-msg success";
            } else {
                hoursValidation.textContent = "⚠ La distribución no sigue exactamente la relación 1:2.";
                hoursValidation.className = "validation-msg error";
            }
        } else {
            hoursValidation.textContent = "✓ Distribución presencial validada.";
            hoursValidation.className = "validation-msg success";
        }
    } else if (sum !== totalRequired) {
        hoursValidation.textContent = `⚠ La suma (${sum}h) no coincide con el total (${totalRequired}h).`;
        hoursValidation.className = "validation-msg error";
    } else {
        hoursValidation.style.display = 'none';
    }
}

// RAP Modal Logic
function openRAPModal(index) {
    editingRAPIndex = index;
    const modal = document.getElementById('rapModal');
    const editor = document.getElementById('rapEditor');
    
    if (index === -1) {
        editor.value = '';
    } else {
        const raps = getCurrentRAPs();
        editor.value = raps[index];
    }
    
    modal.style.display = 'flex';
}

function closeRAPModal() {
    document.getElementById('rapModal').style.display = 'none';
}

function getCurrentRAPs() {
    return Array.from(rapsListContainer.querySelectorAll('.rap-text')).map(el => el.textContent);
}

function saveRAP() {
    const text = document.getElementById('rapEditor').value.trim();
    if (!text) return;

    let raps = getCurrentRAPs();
    if (editingRAPIndex === -1) {
        raps.push(text);
    } else {
        raps[editingRAPIndex] = text;
    }

    renderRAPs(raps);
    
    // Sync with DB
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    if (course) course.resultados_aprendizaje = raps;
    
    closeRAPModal();
}

function deleteRAP(index) {
    if (confirm('¿Está seguro de eliminar este resultado de aprendizaje?')) {
        let raps = getCurrentRAPs();
        raps.splice(index, 1);
        renderRAPs(raps);
        
        // Sync with DB
        const course = syllabusDb.find(c => c.id == currentCourseIndex);
        if (course) course.resultados_aprendizaje = raps;
    }
}

// Start app
// AI Pedagogical Assistant Logic
async function showIALoading(text) {
    const overlay = document.getElementById('iaOverlay');
    const statusText = document.getElementById('iaProcessingText');
    statusText.textContent = `✨ IA: ${text}`;
    overlay.style.display = 'flex';
    await new Promise(r => setTimeout(r, 1500));
}

function hideIALoading() {
    document.getElementById('iaOverlay').style.display = 'none';
}

async function autoAlignSyllabusWithIA() {
    if (currentCourseIndex === -1) return alert("Seleccione una asignatura primero");
    
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    const raps = course.resultados_aprendizaje || [];

    if (raps.length === 0) return alert("Primero agregue los Resultados de Aprendizaje (RAP).");

    const prompt = `Actúa como Diseñador Instruccional Senior experto en Ambientes Virtuales de Aprendizaje (AVA) para la ESUFA. 
    Modalidad: ${course.modality || 'Virtual'}.
    Asignatura: ${course.modulo}
    Objetivo: ${course.objetivo_general}
    Resultados (RAP): ${raps.join(' | ')}
    Unidades: ${JSON.stringify(course.unidades)}
    
    INSTRUCCIONES DE DISEÑO INSTRUCCIONAL (METODOLOGÍAS ACTIVAS):
    No te limites a cambiar nombres. Genera ACTIVIDADES REALES DE APRENDIZAJE ACTIVO:
    1. Si es VIRTUAL: Sugiere actividades como: 
       - "Simulación de escenarios técnicos mediante software específico".
       - "Resolución de Casos de Estudio en equipos virtuales (Teams/Moodle)".
       - "Construcción de Wikis colaborativas sobre doctrina técnica".
       - "Laboratorios virtuales o simuladores de vuelo/mantenimiento".
       - "Debates sincrónicos sobre ética o táctica en la fuerza".
    2. RECURSOS: Deben ser digitales (OVAs, Simuladores, Guías interactivas, Videotutoriales técnicos).
    3. EVALUACIÓN: Debe ser evaluación auténtica (Proyectos, Portafolios, Rúbricas de desempeño en simulador).
    
    PROHIBIDO: Cualquier mención a "clases" tradicionales, lecturas pasivas o "clases magistrales". Todo debe centrarse en el HACER del estudiante.

    Devuelve un JSON con: { unidades: [{ nombre, tema, subtemas, actividades, recursos, rap_id }], evaluaciones: [{ actividad, descripcion, instrucciones, criterios, instrumento, rap_id }] }.`;

    try {
        await showIALoading("Analizando Alineamiento Constructivo Global...");
        const response = await callGeminiAI(prompt);
        // Extract JSON from response (handling potential markdown)
        const jsonMatch = response.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
            const data = JSON.parse(jsonMatch[0]);
            
            // Fix: Flatten any AI-returned objects or arrays to strings
            const flatten = (val) => {
                if (val === null || val === undefined) return '';
                if (Array.isArray(val)) {
                    return val.map(item => {
                        if (typeof item === 'object' && item !== null) {
                            return item.descripcion || item.actividad || item.nombre || item.texto || item.tema || JSON.stringify(item);
                        }
                        return item;
                    }).join('\n');
                }
                if (typeof val === 'object' && val !== null) {
                    return val.descripcion || val.actividad || val.nombre || val.texto || val.tema || JSON.stringify(val);
                }
                return String(val);
            };

            course.unidades = (data.unidades || []).map(u => ({
                ...u,
                actividades: flatten(u.actividades),
                recursos: flatten(u.recursos),
                tema: flatten(u.tema),
                subtemas: flatten(u.subtemas || u.temas || u.contenido)
            }));
            
            course.evaluaciones = (data.evaluaciones || []).map(ev => ({
                ...ev,
                descripcion: flatten(ev.descripcion),
                instrucciones: flatten(ev.instrucciones),
                criterios: flatten(ev.criterios)
            }));

            renderUnits(course.unidades);
            renderEvaluations(course.evaluaciones);
        }
    } catch (error) {
        console.error("AI Error:", error);
        alert("Error al conectar con la IA: " + error.message);
    } finally {
        hideIALoading();
    }
}

function decideInstrumentAndActivity(unitName, rapText) {
    const rap = rapText.toLowerCase();
    
    // Virtual Simulation / Interactive
    if (rap.includes("simular") || rap.includes("operar") || rap.includes("ejecutar")) {
        return {
            actividad: `Reto en Simulador Digital: ${unitName}`,
            instrumento: "Guía de Observación (Rúbrica de ejecución en línea)",
            descripcion: `Práctica realizada a través de herramientas de simulación web o software específico disponible en el AVA.`
        };
    }
    
    // Digital Synthesis
    if (rap.includes("sintetizar") || rap.includes("organizar") || rap.includes("explicar")) {
        return {
            actividad: `Propuesta de diseño (Genial.ly/Canva): ${unitName}`,
            instrumento: "Rúbrica de Producto Digital",
            descripcion: `Entrega de un recurso multimedia donde se estructuren los conceptos clave de la unidad en entorno virtual.`
        };
    }
    
    // Automated Knowledge
    if (rap.includes("identificar") || rap.includes("definir") || rap.includes("conocer")) {
        return {
            actividad: `Cuestionario de Validación Técnica: ${unitName}`,
            instrumento: "Examen / Cuestionario (Autocalificable)",
            descripcion: `Evaluación de conceptos fundamentales a través del centro de calificaciones del AVA Blackboard.`
        };
    }
    
    // Default to Virtual Analysis
    return {
        actividad: `Estudio de Caso / Análisis de Problemas: ${unitName}`,
        instrumento: "Rúbrica de Análisis Crítico",
        descripcion: `Resolución de un escenario técnico de la Fuerza Aérea basado en documentación digital y cargue de informe en el AVA.`
    };
}

function findBestRAPMatch(text, raps) {
    text = text.toLowerCase();
    let bestMatch = 0, maxScore = -1;
    raps.forEach((rap, idx) => {
        const score = rap.toLowerCase().split(/\s+/).filter(w => w.length > 4 && text.includes(w)).length;
        if (score > maxScore) { maxScore = score; bestMatch = idx; }
    });
    return bestMatch;
}

async function optimizeRAPWithIA() {
    const editor = document.getElementById('rapEditor');
    if (!editor.value) return;

    try {
        await showIALoading("Optimizando RAP con Taxonomía de Bloom...");
        const response = await callGeminiAI(`Corrije y mejora este Resultado de Aprendizaje (RAP) usando un verbo de la Taxonomía de Bloom (Nivel 4 o superior) y asegúrate de que sea medible: "${editor.value}". Devuelve solo el texto corregido.`);
        editor.value = response.trim();
    } catch (error) {
        alert("Error IA: " + error.message);
    } finally {
        hideIALoading();
    }
}

async function suggestActivities() {
    if (currentCourseIndex === -1) return;
    await showIALoading("Sugiriendo Actividades y Recursos (AVA/Presencial)...");
    
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    course.unidades?.forEach(unit => {
        if (!unit.actividades) {
            unit.actividades = "Taller colaborativo en AVA sobre " + unit.nombre + " con resolución de casos reales.";
        }
        if (!unit.recursos) {
            unit.recursos = "Plataforma Blackboard, Biblioteca ESUFA, Material multimedia interactivo.";
        }
    });
    
    renderUnits(course.unidades);
    hideIALoading();
}

async function suggestEvaluations() {
    if (currentCourseIndex === -1) return;
    await showIALoading("Diseñando Evaluación Auténtica e Instrumentos...");
    
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    if (!course.evaluaciones || course.evaluaciones.length === 0) {
        course.evaluaciones = [{
            actividad: "Reto de Integración Profesional",
            descripcion: "Simulación de un escenario real donde el estudiante debe aplicar los conocimientos de la asignatura para resolver un problema operativo de la Fuerza Aérea.",
            instrucciones: "1. Analizar el caso. 2. Propner solución técnica. 3. Sustentar resultados.",
            criterios: "Claridad técnica, Coherencia con la doctrina militar, Capacidad de síntesis.",
            instrumento: "Rúbrica"
        }];
    } else {
        course.evaluaciones.forEach(ev => {
            ev.actividad = "Proyecto de Aplicación: " + ev.actividad;
            ev.instrumento = "Rúbrica";
        });
    }
    
    renderEvaluations(course.evaluaciones);
    hideIALoading();
}

// --- NEW UTILITIES ---

// 1. AI Calling Logic
async function callGeminiAI(userPrompt) {
    const endpoint = localStorage.getItem('ai_endpoint') || 'https://silaboesufa.monicacastillaluna.workers.dev';
    const apiKey = localStorage.getItem('ai_api_key') || '';
    
    if (!endpoint && !apiKey) {
        throw new Error("No se ha configurado la IA. Por favor, ingrese la URL o API Key en Ajustes.");
    }

    // Standard Gemini Payload
    const body = { 
        contents: [{ parts: [{ text: userPrompt }] }],
        generationConfig: { temperature: 0.7, topP: 0.95, topK: 40, maxOutputTokens: 2048 }
    };

    const model = document.getElementById('aiModel')?.value || localStorage.getItem('ai_model') || 'gemini-flash-latest';
    const url = endpoint || `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    
    const maxRetries = 3;
    let lastError = null;

    for (let i = 0; i < maxRetries; i++) {
        try {
            console.log(`Intento AI ${i + 1}/${maxRetries} via:`, endpoint ? "Proxy (" + endpoint + ")" : "Direct API");
            
            const response = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(body)
            });

            if (response.status === 429 || response.status === 503) {
                console.warn(`Servidor ocupado (${response.status}). Reintentando en 2.5s (${i+1}/${maxRetries})...`);
                await new Promise(r => setTimeout(r, 2500));
                continue;
            }

            if (!response.ok) {
                let errDetail = "Sin detalle";
                try {
                    const errJson = await response.json();
                    if (errJson.error) {
                        errDetail = errJson.error.message || JSON.stringify(errJson.error);
                    } else {
                        errDetail = JSON.stringify(errJson);
                    }
                } catch (e) {
                    errDetail = await response.text().catch(() => "Error ilegible");
                }
                throw new Error(`Error ${response.status}: ${errDetail}`);
            }

            const data = await response.json();
            
            // Standard Gemini response path
            if (data.candidates && data.candidates[0]?.content?.parts) {
                return data.candidates[0].content.parts[0].text;
            }
            // Proxy response paths
            if (data.text) return data.text;
            if (data.response) return data.response;
            if (typeof data === 'string') return data;
            
            if (data.error) {
                throw new Error(`Error de la IA: ${data.error.message || "Error desconocido"}`);
            }
            
            throw new Error("Formato de respuesta desconocido.");

        } catch (error) {
            lastError = error;
            if (i === maxRetries - 1) throw error;
            console.warn(`Fallo en intento ${i + 1}. Reintentando...`, error.message);
            await new Promise(r => setTimeout(r, 1500));
        }
    }
    throw lastError;
}

async function testAIConnection() {
    const btn = event.target;
    const originalText = btn.textContent;
    btn.textContent = "⌛ Probando...";
    btn.disabled = true;

    try {
        const response = await callGeminiAI("Hola, ¿estás activo? Responde solo 'SÍ'.");
        alert("¡Conexión exitosa! La IA dice: " + response);
    } catch (error) {
        console.error("Test Error:", error);
        alert("Fallo en la prueba: " + error.message);
    } finally {
        btn.textContent = originalText;
        btn.disabled = false;
    }
}

// 2. Settings Management
function openSettingsModal() { document.getElementById('settingsModal').style.display = 'flex'; }
function closeSettingsModal() { document.getElementById('settingsModal').style.display = 'none'; }

function saveSettings() {
    let endpoint = document.getElementById('aiEndpoint').value.trim();
    const apiKey = document.getElementById('aiApiKey').value.trim();
    const model = document.getElementById('aiModel').value;
    
    // Auto-fix protocol if missing
    if (endpoint && !endpoint.startsWith('http')) {
        endpoint = 'https://' + endpoint;
        document.getElementById('aiEndpoint').value = endpoint;
    }

    localStorage.setItem('ai_endpoint', endpoint);
    localStorage.setItem('ai_api_key', apiKey);
    localStorage.setItem('ai_model', model);
    closeSettingsModal();
    alert("Configuración guardada. Recuerde que su Worker debe estar 'Deplyed' y ser público en Cloudflare.");
}

function loadSettings() {
    document.getElementById('aiEndpoint').value = localStorage.getItem('ai_endpoint') || 'https://silaboesufa.monicacastillaluna.workers.dev';
    document.getElementById('aiApiKey').value = localStorage.getItem('ai_api_key') || '';
    if (document.getElementById('aiModel')) {
        document.getElementById('aiModel').value = localStorage.getItem('ai_model') || 'gemini-flash-latest';
    }
}

// 3. History Management
function saveSyllabusToHistory() {
    if (currentCourseIndex === -1) return alert("Nada que guardar.");
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    
    const history = JSON.parse(localStorage.getItem('syllabus_history') || '[]');
    
    // Buscar si ya existe una versión de este mismo curso
    const existingIdx = history.findIndex(item => 
        item.data.modulo === course.modulo && item.data.sigla === course.sigla
    );

    if (existingIdx !== -1) {
        const confirmUpdate = confirm(`Ya existe una versión de "${course.modulo}". ¿Desea ACTUALIZAR la versión existente? (Cancelar para guardar como una versión nueva)`);
        
        if (confirmUpdate) {
            history[existingIdx].timestamp = new Date().toISOString();
            history[existingIdx].data = JSON.parse(JSON.stringify(course));
            localStorage.setItem('syllabus_history', JSON.stringify(history));
            alert(`✅ Versión actualizada con éxito.`);
            return;
        }
    }

    const savedItem = {
        timestamp: new Date().toISOString(),
        id: Date.now(),
        data: JSON.parse(JSON.stringify(course))
    };
    
    history.unshift(savedItem);
    localStorage.setItem('syllabus_history', JSON.stringify(history.slice(0, 50)));
    alert(`✅ "${course.modulo}" guardado como nueva versión.`);
}



function openHistoryModal() {
    const list = document.getElementById('historyList');
    const history = JSON.parse(localStorage.getItem('syllabus_history') || '[]');
    
    if (history.length === 0) {
        list.innerHTML = `
            <div style="text-align: center; color: var(--text-muted); padding: 40px;">
                <span style="font-size: 3rem; display: block; margin-bottom: 1rem;">empty</span>
                <p>No hay sílabos guardados aún.</p>
            </div>`;
    } else {
        list.innerHTML = history.map(item => `
            <div class="history-item">
                <div class="history-item-info">
                    <h4>${item.data.modulo}</h4>
                    <p>📅 ${new Date(item.timestamp).toLocaleString()}</p>
                </div>
                <div class="history-actions">
                    <button class="btn btn-secondary btn-small" onclick="loadSyllabusFromHistory(${item.id})">
                        <span class="icon">📂</span> Cargar
                    </button>
                    <button class="btn btn-icon btn-danger" onclick="deleteHistoryItem(${item.id})" title="Eliminar">
                        ✕
                    </button>
                </div>
            </div>
        `).join('');
    }
    document.getElementById('historyModal').style.display = 'flex';
}

function closeHistoryModal() { document.getElementById('historyModal').style.display = 'none'; }

function loadSyllabusFromHistory(id) {
    const history = JSON.parse(localStorage.getItem('syllabus_history') || '[]');
    const item = history.find(i => i.id === id);
    if (item) {
        // Find existing or push new to DB
        const existingIdx = syllabusDb.findIndex(c => c.sigla === item.data.sigla && c.modulo === item.data.modulo);
        if (existingIdx !== -1) {
            syllabusDb[existingIdx] = item.data;
            selectCourse(syllabusDb[existingIdx]);
        } else {
            item.data.id = syllabusDb.length;
            syllabusDb.push(item.data);
            selectCourse(item.data);
        }
        closeHistoryModal();
    }
}

function deleteHistoryItem(id) {
    if (!confirm("¿Eliminar este registro del historial?")) return;
    let history = JSON.parse(localStorage.getItem('syllabus_history') || '[]');
    history = history.filter(i => i.id !== id);
    localStorage.setItem('syllabus_history', JSON.stringify(history));
    openHistoryModal();
}

// 4. PDF Export
async function exportToPDF() {
    const element = document.querySelector('.editor-area');
    const courseName = document.getElementById('moduleName').value || 'Syllabus';

    // 1. Prepare UI - Force everything to be visible
    document.body.classList.add('pdf-export-mode');
    document.body.style.overflow = 'visible';
    
    const textareas = element.querySelectorAll('textarea');
    textareas.forEach(tx => {
        tx.setAttribute('data-old-height', tx.style.height);
        tx.style.height = 'auto';
        tx.style.height = (tx.scrollHeight + 10) + 'px';
        tx.innerHTML = tx.value; // Important for some renderers
    });

    const opt = {
        margin: [0.5, 0.5],
        filename: `Syllabus_${courseName.replace(/ /g, '_')}.pdf`,
        image: { type: 'jpeg', quality: 1.0 },
        html2canvas: { 
            scale: 2, 
            useCORS: true, 
            logging: false,
            letterRendering: true,
            windowWidth: 1200, // Force a wide window for rendering
            scrollY: 0
        },
        jsPDF: { unit: 'in', format: 'letter', orientation: 'portrait' },
        pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
    };

    try {
        alert("Generando PDF... Capturando TODO el contenido.");
        // Use from(element) directly on the real, modified UI
        await html2pdf().set(opt).from(element).save();
    } catch (err) {
        console.error("PDF Export Error:", err);
        alert("Error al generar PDF.");
    } finally {
        // Restore heights and UI mode
        textareas.forEach(tx => tx.style.height = tx.getAttribute('data-old-height') || '');
        document.body.style.overflow = '';
        document.body.classList.remove('pdf-export-mode');
    }
}

async function checkAvailableModels() {
    const apiKey = document.getElementById('aiApiKey').value.trim() || localStorage.getItem('ai_api_key');
    if (!apiKey) return alert("Ingrese una API Key en el campo de abajo para diagnosticar.");
    
    const btn = event.target;
    btn.textContent = "⌛ Diagnosticando...";
    
    try {
        console.log("Diagnosticando con clave:", apiKey.substring(0, 6) + "...");
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
        
        if (!response.ok) {
            const err = await response.json();
            throw new Error(err.error?.message || "Error de red");
        }

        const data = await response.json();
        
        if (!data.models || data.models.length === 0) {
            alert("⚠️ La clave funciona pero Google no te ha asignado ningún modelo aún.");
            return;
        }

        const models = data.models
            .map(m => `- ${m.name.replace('models/', '')} (${m.displayName})`)
            .join('\n');
            
        alert("✅ CLAVE VÁLIDA. Modelos disponibles para ti:\n\n" + models + "\n\nUse uno de estos nombres en su Worker.");
    } catch (error) {
        console.error("DIAGNOSTIC ERROR:", error);
        alert("❌ ERROR DE DIAGNÓSTICO: " + error.message + "\n\nSugerencia: Intente generar una clave nueva en Google AI Studio.");
    } finally {
        btn.textContent = "📋 Lista Modelos";
    }
}

// 5. Word Export (Structured Clean Generation)
function exportToWord() {
    const courseName = document.getElementById('moduleName').value || 'Syllabus';
    const courseCode = document.getElementById('moduleCode').value || '';
    const credits = document.getElementById('credits').value || '1';
    const objective = document.getElementById('courseObjective').value || '';
    
    // Header
    let docHtml = `
        <h1 style="text-align: center; color: #1a237e;">ESUFA - SÍLABO ESTRATÉGICO</h1>
        <h2 style="text-align: center;">${courseName.toUpperCase()}</h2>
        <p style="text-align: center;"><strong>Sigla:</strong> ${courseCode} | <strong>Créditos:</strong> ${credits}</p>
        <hr>
        <h3 style="color: #1a237e; background: #f0f0f0; padding: 5px;">I. OBJETIVO GENERAL</h3>
        <p>${objective.replace(/\n/g, '<br>')}</p>
    `;

    // 1. RAPs
    docHtml += `<h3 style="color: #1a237e; background: #f0f0f0; padding: 5px;">II. RESULTADOS DE APRENDIZAJE (RAP)</h3><ul>`;
    const raps = document.querySelectorAll('.rap-text');
    if (raps.length > 0) {
        raps.forEach(rap => {
            docHtml += `<li style="margin-bottom: 5px;">${rap.innerText}</li>`;
        });
    } else {
        docHtml += `<li>No definidos aún.</li>`;
    }
    docHtml += `</ul>`;

    // 2. Units (The core table)
    docHtml += `<h3 style="color: #1a237e; background: #f0f0f0; padding: 5px;">III. ESTRUCTURA DE ENSEÑANZA Y APRENDIZAJE</h3>`;
    const units = document.querySelectorAll('.unit-row');
    if (units.length > 0) {
        units.forEach((row, i) => {
            const name = row.querySelector('.unit-textarea-label')?.value || 'Sin Nombre';
            const textareas = row.querySelectorAll('.topics-textarea');
            const tema = textareas[0]?.value || '';
            const subtemas = textareas[1]?.value || '';
            const activities = textareas[2]?.value || '';
            const resources = textareas[3]?.value || '';
            docHtml += `
                <div style="border: 1px solid #ccc; margin-bottom: 15px; padding: 10px;">
                    <h4 style="margin: 0 0 10px 0; color: #1565c0;">UNIDAD ${i+1}: ${name}</h4>
                    <table width="100%" border="0" cellspacing="0" cellpadding="5">
                        <tr><td width="20%"><strong>Tema Principal:</strong></td><td>${tema.replace(/\n/g, '<br>')}</td></tr>
                        <tr><td width="20%"><strong>Subtemas:</strong></td><td>${subtemas.replace(/\n/g, '<br>')}</td></tr>
                        <tr><td width="20%"><strong>Actividades AVA:</strong></td><td>${activities.replace(/\n/g, '<br>')}</td></tr>
                        <tr><td width="20%"><strong>Recursos:</strong></td><td>${resources.replace(/\n/g, '<br>')}</td></tr>
                    </table>
                </div>
            `;
        });
    } else {
        docHtml += `<p>No se han definido unidades para este sílabo.</p>`;
    }

    // 3. Evaluations (Using internal DB for accuracy)
    docHtml += `<h3 style="color: #1a237e; background: #f0f0f0; padding: 5px;">IV. PLAN DE EVALUACIÓN AUTÉNTICA</h3>`;
    const course = syllabusDb.find(c => c.id == currentCourseIndex);
    
    if (course.evaluaciones && course.evaluaciones.length > 0) {
        course.evaluaciones.forEach((ev, idx) => {
            const rapLink = ev.rap_id !== undefined ? ` (Vinculado a: RAP ${ev.rap_id + 1})` : "";
            
            docHtml += `
                <div style="border: 1px solid #999; margin: 15px 0; padding: 15px; background: #fafafa;">
                    <table width="100%" border="0" cellspacing="0" cellpadding="3">
                        <tr><td width="20%"><strong>Actividad ${idx+1}:</strong></td><td style="font-size: 1.1rem; color: #1a237e;"><strong>${ev.actividad?.toUpperCase()}</strong>${rapLink}</td></tr>
                        <tr><td><strong>Instrumento:</strong></td><td>${ev.instrumento}</td></tr>
                    </table>
                    <div style="margin-top: 10px;">
                        <p style="margin-bottom: 5px;"><strong>Descripción del Reto:</strong></p>
                        <p style="margin-left: 15px; color: #444;">${(ev.descripcion || 'No definida').replace(/\n/g, '<br>')}</p>
                        
                        <p style="margin-bottom: 5px; margin-top: 10px;"><strong>Instrucciones:</strong></p>
                        <p style="margin-left: 15px; color: #444;">${(ev.instrucciones || 'No definidas').replace(/\n/g, '<br>')}</p>
                        
                        <p style="margin-bottom: 5px; margin-top: 10px;"><strong>Criterios / Rúbrica:</strong></p>
                        <p style="margin-left: 15px; color: #444;">${(ev.criterios || 'No definidos').replace(/\n/g, '<br>')}</p>
                    </div>
                </div>
            `;
        });
    } else {
        docHtml += `<p>No hay evaluaciones definidas.</p>`;
    }

    // Meta Footer
    docHtml += `<hr><p style="font-size: 0.8rem; text-align: center; color: #777;">Generado por Syllabus Virtualizer ESUFA | ${new Date().toLocaleDateString()}</p>`;

    // Download logic
    const blob = new Blob(['\ufeff', `
        <!DOCTYPE html><html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head><meta charset='utf-8'><title>Syllabus</title>
        <style>
            body { font-family: 'Times New Roman', serif; font-size: 11pt; }
            h1, h2, h3 { font-family: 'Arial', sans-serif; }
            table { border-collapse: collapse; }
        </style></head><body>${docHtml}</body></html>
    `], { type: 'application/msword' });
    
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `Syllabus_${courseName.replace(/ /g, '_')}.doc`;
    link.click();
    
    alert("✓ Word generado con éxito (Limpio).");
}

init();

