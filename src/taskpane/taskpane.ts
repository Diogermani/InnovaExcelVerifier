import "./taskpane.css";
import { getConfig } from "../config/configService";

// Register initialization callback
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initTaskpane();
  }
});

async function initTaskpane() {
  const configGrid = document.getElementById("config-grid");
  if (!configGrid) return;

  try {
    const config = await getConfig();
    
    // Map configurations to readable Portuguese items
    const items = [
      { label: "Domínio da Empresa", value: config.companyDomain },
      { label: "Colunas de Alerta", value: `${config.startColumn} em diante` },
      { label: "Idiomas Suportados", value: config.language },
      { label: "Validar Abas Ocultas", value: config.validateHiddenSheets ? "Sim" : "Não" },
      { label: "Permitir Envio com Alerta", value: config.allowUserOverride ? "Sim" : "Não" },
      { label: "Extensões Monitoradas", value: config.supportedExtensions.join(", ") }
    ];

    configGrid.innerHTML = items
      .map(
        (item) => `
        <div class="config-item">
          <span class="config-label">${item.label}</span>
          <span class="config-value">${item.value}</span>
        </div>
      `
      )
      .join("");
  } catch (error) {
    configGrid.innerHTML = `
      <div class="config-item" style="border-color: rgba(239, 68, 68, 0.2); background: rgba(239, 68, 68, 0.05);">
        <span class="config-label" style="color: #ef4444;">Falha ao carregar configurações.</span>
      </div>
    `;
  }
}
