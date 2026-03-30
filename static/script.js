/**
 * script.js — Control de Acceso Alcorta Shopping
 * ================================================
 * Maneja toda la interactividad del frontend: escaneo de DNI,
 * visualización del resultado, actualización del dashboard en tiempo
 * real y gestión del panel de administración.
 *
 * Una sola entrada (DOMContentLoaded) detecta en qué página está el usuario
 * según los elementos presentes en el DOM y activa solo la lógica correspondiente.
 *
 * Autor: Roberto Laforcada
 */

document.addEventListener('DOMContentLoaded', () => {

    // =========================================================================
    // FUNCIONES GLOBALES (compartidas entre las páginas de admisión y admin)
    // =========================================================================

    /**
     * Carga los registros de acceso del día actual y los muestra en una tabla.
     * Los registros se muestran en orden inverso (el más reciente primero).
     * Colorea cada fila según el resultado (verde/rojo) y el tipo de permiso.
     *
     * @param {HTMLElement} targetTableBody - Elemento <tbody> donde insertar las filas.
     */
    async function loadDailyRecords(targetTableBody) {
        if (!targetTableBody) return;

        try {
            const response = await fetch('/get_daily_records');
            const data     = await response.json();
            targetTableBody.innerHTML = '';

            if (data.success && data.records.length > 0) {
                // Mostrar el más reciente primero
                data.records.reverse().forEach(record => {
                    const row = targetTableBody.insertRow();

                    // Aplicar clase de color según resultado
                    if (record.Resultado === 'VERDE') row.classList.add('fila-verde');
                    else if (record.Resultado === 'ROJO') row.classList.add('fila-roja');

                    // Los registros FAO llevan borde amarillo adicional
                    if (record.Tipo_Permiso === 'FAO') row.classList.add('fila-fao');

                    // Extraer solo la hora (HH:MM:SS) del campo que puede venir como datetime
                    const hora = record.Hora_Ingreso
                        ? (record.Hora_Ingreso.split(' ')[1] || record.Hora_Ingreso)
                        : '';

                    row.insertCell().textContent = hora;
                    row.insertCell().textContent = record.DNI || '';
                    row.insertCell().textContent = record['Nombre y Apellido'] || '';
                    row.insertCell().textContent = record.Num_Permiso || '';
                    row.insertCell().textContent = record.Local || '';
                    row.insertCell().textContent = record.Tarea || '';
                });

            } else {
                // Sin registros: mostrar mensaje centrado
                const row  = targetTableBody.insertRow();
                const cell = row.insertCell(0);
                cell.colSpan     = 6;
                cell.textContent = data.message || 'No hay registros para mostrar.';
                cell.style.textAlign = 'center';
            }

        } catch (error) {
            console.error('Error al cargar registros diarios:', error);
        }
    }

    /**
     * Actualiza los contadores del dashboard (Total, Permitidos, Rechazados)
     * consultando al servidor el resumen estadístico del día.
     * Solo actúa si el elemento #statTotal existe en la página actual.
     */
    async function updateStats() {
        const statTotalEl = document.getElementById('statTotal');
        if (!statTotalEl) return;

        try {
            const response = await fetch('/get_daily_stats');
            const stats    = await response.json();
            statTotalEl.textContent = stats.total;
            document.getElementById('statPermitidos').textContent  = stats.permitidos;
            document.getElementById('statRechazados').textContent  = stats.rechazados;
        } catch (error) {
            console.error('Error al actualizar estadísticas:', error);
        }
    }


    // =========================================================================
    // PÁGINA DE ADMISIÓN (index.html)
    // Activa solo si existe el campo de entrada de DNI (#dniInput)
    // =========================================================================

    const dniInput = document.getElementById('dniInput');
    if (dniInput) {
        // Referencias a los elementos del panel de resultado
        const resultadoDiv      = document.getElementById('resultado');
        const mensajeResultado  = document.getElementById('mensajeResultado');
        const nombrePersona     = document.getElementById('nombrePersona');
        const infoLocal         = document.getElementById('infoLocal');
        const infoTarea         = document.getElementById('infoTarea');
        const infoVence         = document.getElementById('infoVence');
        const ingresosListBody  = document.getElementById('ingresosListBody');

        // Mantener el foco siempre en el input para capturar el lector de código de barras
        dniInput.focus();
        dniInput.addEventListener('keydown', (event) => {
            if (event.key === 'Enter') {
                event.preventDefault();
                verificarDNI();
            }
        });

        /**
         * Procesa el valor escaneado o escrito en el campo DNI.
         *
         * Los lectores de código de barras del DNI argentino emiten una cadena
         * completa con todos los datos del documento. Esta función extrae el DNI
         * buscando una secuencia de 7 u 8 dígitos consecutivos en esa cadena.
         *
         * Si no encuentra ese patrón (ej. ingreso manual), toma todos los dígitos
         * de la cadena y los une como DNI.
         */
        async function verificarDNI() {
            const scannerData = dniInput.value.trim();
            if (!scannerData) return;

            let dniParaVerificar = '';

            // Buscar la primera secuencia de 7 u 8 dígitos en la cadena escaneada.
            // Esto cubre tanto DNI de 7 como de 8 dígitos, y funciona con
            // los distintos formatos de código de barras del DNI argentino.
            const match = scannerData.match(/\b\d{7,8}\b/);

            if (match) {
                dniParaVerificar = match[0];
            } else {
                // Fallback: extraer todos los dígitos y usarlos si superan los 7
                const numeros = scannerData.replace(/\D/g, '');
                if (numeros.length >= 7) {
                    dniParaVerificar = numeros;
                }
            }

            if (!dniParaVerificar) {
                mostrarResultado('red', 'No se pudo encontrar un DNI válido en el código.', '', '', '', '');
                dniInput.value = '';
                dniInput.focus();
                return;
            }

            try {
                const response = await fetch('/verificar_dni', {
                    method:  'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body:    JSON.stringify({ dni: dniParaVerificar }),
                });
                const data = await response.json();

                // Generar texto de vencimiento según el resultado del acceso
                const vencimientoTexto = (data.vence && data.vence !== 'N/A')
                    ? (data.acceso === 'PERMITIDO' ? `Vence: ${data.vence}` : `Vencido el: ${data.vence}`)
                    : '';

                mostrarResultado(
                    data.acceso === 'PERMITIDO' ? 'green' : 'red',
                    data.mensaje,
                    data.nombre,
                    data.local,
                    data.tarea,
                    vencimientoTexto
                );

            } catch (error) {
                console.error('Error al verificar DNI:', error);
                mostrarResultado('red', 'Error de comunicación con el servidor.', '', '', '', '');

            } finally {
                // Siempre limpiar el input y actualizar la vista
                dniInput.value = '';
                dniInput.focus();
                updateStats();
                loadDailyRecords(ingresosListBody);
            }
        }

        /**
         * Actualiza visualmente el panel de resultado con el color, mensaje
         * y datos de la persona según la respuesta del servidor.
         *
         * @param {string} colorClass     - 'green' para permitido, 'red' para denegado.
         * @param {string} mensaje        - Texto principal del resultado.
         * @param {string} nombre         - Nombre completo de la persona.
         * @param {string} localInfo      - Local o marca asociada.
         * @param {string} tareaInfo      - Descripción de la tarea (FAO).
         * @param {string} venceInfo      - Texto de vencimiento del permiso.
         */
        function mostrarResultado(colorClass, mensaje, nombre, localInfo, tareaInfo, venceInfo) {
            resultadoDiv.className    = `result-display luz-${colorClass}`;
            mensajeResultado.textContent = mensaje;
            nombrePersona.textContent = (nombre && nombre !== 'No Encontrado')
                ? `Nombre: ${nombre}` : '';

            // Limpiar campos secundarios siempre
            infoLocal.textContent = '';
            infoTarea.textContent = '';
            infoVence.textContent = '';

            // Mostrar detalles adicionales solo en accesos permitidos
            if (colorClass === 'green') {
                infoLocal.textContent = (localInfo && localInfo !== 'N/A') ? `Local: ${localInfo}` : '';
                infoTarea.textContent = (tareaInfo && tareaInfo !== 'N/A') ? `Tarea: ${tareaInfo}` : '';
                infoVence.textContent = venceInfo;
            }
        }

        // Carga inicial al abrir la página
        loadDailyRecords(ingresosListBody);
        updateStats();
    }


    // =========================================================================
    // PÁGINA DE LOGIN (login.html)
    // Activa solo si existe el botón #loginBtn
    // =========================================================================

    const loginBtn = document.getElementById('loginBtn');
    if (loginBtn) {
        const usernameInput = document.getElementById('username');
        const passwordInput = document.getElementById('password');
        const loginMessage  = document.getElementById('loginMessage');

        loginBtn.addEventListener('click', async () => {
            const username = usernameInput.value.trim();
            const password = passwordInput.value.trim();

            try {
                const response = await fetch('/perform_login', {
                    method:  'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body:    JSON.stringify({ username, password }),
                });
                const data = await response.json();

                if (data.success) {
                    window.location.href = '/admin';
                } else {
                    loginMessage.textContent = data.message;
                }
            } catch (error) {
                loginMessage.textContent = 'Error de comunicación con el servidor.';
            }
        });

        // Permitir enviar el formulario con Enter desde los campos de texto
        [usernameInput, passwordInput].forEach(input => {
            input.addEventListener('keydown', (event) => {
                if (event.key === 'Enter') loginBtn.click();
            });
        });
    }


    // =========================================================================
    // PÁGINA DE ADMINISTRACIÓN (admin.html)
    // Activa solo si existe la tabla de registros del admin (#adminIngresosListBody)
    // =========================================================================

    const adminIngresosListBody = document.getElementById('adminIngresosListBody');
    if (adminIngresosListBody) {

        // Carga inicial de registros del día en la tabla del admin
        loadDailyRecords(adminIngresosListBody);

        // --- Sección: Subir Archivos Excel ---

        const fapInput   = document.getElementById('fapFileInput');
        const faoInput   = document.getElementById('faoFileInput');
        const fapFileName = document.getElementById('fapFileName');
        const faoFileName = document.getElementById('faoFileName');

        // Mostrar el nombre del archivo seleccionado en la UI
        if (fapInput) {
            fapInput.addEventListener('change', () => {
                fapFileName.textContent = fapInput.files.length > 0
                    ? fapInput.files[0].name : 'Ningún archivo seleccionado';
            });
        }
        if (faoInput) {
            faoInput.addEventListener('change', () => {
                faoFileName.textContent = faoInput.files.length > 0
                    ? faoInput.files[0].name : 'Ningún archivo seleccionado';
            });
        }

        // Enviar los archivos seleccionados al servidor
        const uploadExcelBtn = document.getElementById('uploadExcelBtn');
        uploadExcelBtn.addEventListener('click', async () => {
            const uploadStatus = document.getElementById('uploadStatus');
            const formData     = new FormData();

            if (fapInput.files.length > 0) formData.append('fap_file', fapInput.files[0]);
            if (faoInput.files.length > 0) formData.append('fao_file', faoInput.files[0]);

            if (!formData.entries().next().value) {
                uploadStatus.textContent = 'Por favor, seleccione al menos un archivo.';
                return;
            }

            uploadStatus.textContent = 'Cargando...';
            try {
                const response = await fetch('/upload_excel', { method: 'POST', body: formData });
                const data     = await response.json();
                uploadStatus.textContent = data.message;
            } catch (error) {
                uploadStatus.textContent = 'Error de comunicación con el servidor.';
            }
        });

        // --- Sección: Agregar Excepción de Acceso ---

        const addExcepcionBtn = document.getElementById('addExcepcionBtn');
        addExcepcionBtn.addEventListener('click', async () => {
            const excepcionData = {
                nombre:         document.getElementById('nombreExcepcion').value.trim(),
                apellido:       document.getElementById('apellidoExcepcion').value.trim(),
                dni:            document.getElementById('dniExcepcion').value.trim(),
                local:          document.getElementById('localExcepcion').value.trim(),
                quien_autoriza: document.getElementById('autorizaExcepcion').value.trim(),
            };
            const excepcionStatus = document.getElementById('excepcionStatus');
            excepcionStatus.textContent = 'Guardando...';

            try {
                const response = await fetch('/add_excepcion', {
                    method:  'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body:    JSON.stringify(excepcionData),
                });
                const data = await response.json();
                excepcionStatus.textContent = data.message;

                // Limpiar el formulario si la operación fue exitosa
                if (data.success) {
                    document.querySelector('.excepcion-form').reset();
                }
            } catch (error) {
                excepcionStatus.textContent = 'Error al procesar la respuesta del servidor.';
            }
        });

        // --- Sección: Enviar Reporte Diario por Email ---

        const enviarReporteBtn = document.getElementById('enviarReporteBtn');
        enviarReporteBtn.addEventListener('click', async () => {
            const reporteStatus = document.getElementById('reporteStatus');
            reporteStatus.textContent = 'Enviando...';

            try {
                const response = await fetch('/enviar_reporte_diario', { method: 'POST' });
                const data     = await response.json();
                reporteStatus.textContent = data.message;
            } catch (error) {
                reporteStatus.textContent = 'Error de comunicación con el servidor.';
            }
        });
    }

});