// functions.js - Taskpane UI para listar y aplicar sensitivity labels
// Requiere: Mailbox requirement set mínimo que soporta sensitivityLabelsCatalog (ej. 1.13+)

let labelCatalog = [];

Office.onReady(() => {
  // Inicializar la UI cuando Office esté listo
  initUi();
});

function initUi() {
  const loadingEl = document.getElementById('loading');
  const selectEl = document.getElementById('labelsSelect');
  const applyBtn = document.getElementById('applyBtn');
  const noLabelsEl = document.getElementById('no-labels');
  const statusEl = document.getElementById('status');
  const labelForSelect = document.getElementById('labelForSelect');

  loadingEl.style.display = 'block';
  noLabelsEl.style.display = 'none';
  selectEl.style.display = 'none';
  applyBtn.style.display = 'none';
  labelForSelect.style.display = 'none';
  statusEl.textContent = '';

  // 1) Comprobar si el catálogo está habilitado en el buzón
  Office.context.sensitivityLabelsCatalog.getIsEnabledAsync(function(isEnabledResult) {
    loadingEl.style.display = 'none';
    if (isEnabledResult.status !== Office.AsyncResultStatus.Succeeded || !isEnabledResult.value) {
      statusEl.textContent = "Catálogo de etiquetas no disponible en este buzón.";
      return;
    }

    // 2) Obtener catálogo de etiquetas
    Office.context.sensitivityLabelsCatalog.getAsync(function(catalogResult) {
      if (catalogResult.status !== Office.AsyncResultStatus.Succeeded) {
        statusEl.textContent = "Error obteniendo catálogo de etiquetas: " + (catalogResult.error && catalogResult.error.message);
        return;
      }

      labelCatalog = catalogResult.value || [];
      if (!labelCatalog || labelCatalog.length === 0) {
        noLabelsEl.style.display = 'block';
        statusEl.textContent = "No hay etiquetas de sensibilidad publicadas para este buzón.";
        return;
      }

      // 3) Llenar select con etiquetas
      populateSelect(labelCatalog);
    });
  });

  applyBtn.addEventListener('click', function() {
    const selIndex = selectEl.selectedIndex;
    if (selIndex < 0) {
      statusEl.textContent = "Selecciona primero una etiqueta.";
      return;
    }
    // El value del option es el índice en labelCatalog
    const selectedIndex = parseInt(selectEl.value, 10);
    const selectedLabel = labelCatalog[selectedIndex];
    if (!selectedLabel) {
      statusEl.textContent = "Etiqueta seleccionada inválida.";
      return;
    }
    applyLabelToItem(selectedLabel);
  });
}

function populateSelect(catalog) {
  const selectEl = document.getElementById('labelsSelect');
  const applyBtn = document.getElementById('applyBtn');
  const loadingEl = document.getElementById('loading');
  const labelForSelect = document.getElementById('labelForSelect');

  // Limpiar
  selectEl.innerHTML = '';

  // Orden simple por nombre (opcional)
  catalog.sort((a,b) => ( (a.name||'').localeCompare(b.name||'') ));

  for (let i = 0; i < catalog.length; i++) {
    const lbl = catalog[i];
    const opt = document.createElement('option');
    opt.value = i; // guardamos índice
    const desc = lbl.description ? " — " + lbl.description : "";
    opt.text = (lbl.name || "(sin nombre)") + desc;
    selectEl.appendChild(opt);
  }

  loadingEl.style.display = 'none';
  labelForSelect.style.display = 'block';
  selectEl.style.display = 'block';
  applyBtn.style.display = 'inline-block';
}

function applyLabelToItem(labelDetails) {
  const statusEl = document.getElementById('status');
  statusEl.textContent = "Aplicando etiqueta: " + (labelDetails.name || "(sin nombre)") + "...";

  // labelDetails viene directamente del catálogo; la API acepta el objeto.
  Office.context.mailbox.item.sensitivityLabel.setAsync(labelDetails, function(setResult) {
    if (setResult.status === Office.AsyncResultStatus.Succeeded) {
      statusEl.textContent = "Etiqueta aplicada: " + (labelDetails.name || "(sin nombre)");
    } else {
      statusEl.textContent = "Error al aplicar etiqueta: " + (setResult.error && setResult.error.message);
    }
  });
}
