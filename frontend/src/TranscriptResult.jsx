import { useState } from "react";
import { useLocation, useNavigate } from "react-router-dom";
import "./TranscriptResult.css";

// ── Label maps (mirrors backend) ─────────────────────────────────────────────
const SECTION_LABELS = {
  entree:             "Entrée",
  salon:              "Salon",
  cuisine:            "Cuisine",
  salle_de_bain:      "Salle de Bain",
  chambres:           "Chambres",
  balcon_terrasse:    "Balcon / Terrasse",
  equipements:        "Équipements",
  controle_sensoriel: "Contrôle Sensoriel",
};

const ITEM_LABELS = {
  porte_entree: "Porte d'entrée", sol: "Sol", murs: "Murs",
  lumiere: "Lumière", interphone_digicode: "Interphone / Digicode",
  odeur: "Odeur", sol_murs: "Sol & Murs",
  poussiere_meubles: "Poussière / Meubles", canape: "Canapé",
  coussins: "Coussins", table_basse: "Table basse", rideaux: "Rideaux",
  volets: "Volets", fenetres: "Fenêtres", eclairage: "Éclairage",
  television: "Télévision", decoration: "Décoration",
  plan_travail: "Plan de travail", evier: "Évier", vaisselle: "Vaisselle",
  ustensiles: "Ustensiles", plaques_cuisson: "Plaques de cuisson",
  hotte: "Hotte", micro_ondes_four: "Micro-ondes / Four",
  refrigerateur: "Réfrigérateur",
  machine_cafe_bouilloire: "Machine à café / Bouilloire",
  placards: "Placards", poubelle: "Poubelle",
  produits_menagers: "Produits ménagers", lavabo: "Lavabo",
  miroir: "Miroir", douche_baignoire: "Douche / Baignoire",
  wc: "WC", papier_toilette: "Papier toilette", lumieres: "Lumières",
  poussiere: "Poussière", literie: "Literie", lit: "Lit",
  matelas: "Matelas", placard_dressing: "Placard / Dressing",
  rideaux_volets: "Rideaux / Volets",
  mobilier_exterieur: "Mobilier extérieur", proprete: "Propreté",
  garde_corps: "Garde-corps", interrupteurs: "Interrupteurs",
  prises: "Prises", wifi: "WiFi", eau_chaude: "Eau chaude",
  chauffage: "Chauffage", climatisation: "Climatisation",
  machine_laver: "Machine à laver", portes_serrures: "Portes / Serrures",
  cles: "Clés", odeur_generale: "Odeur générale",
  temperature: "Température", bruit: "Bruit",
  lumiere_ambiante: "Lumière ambiante",
};

// ── Component ─────────────────────────────────────────────────────────────────
export default function TranscriptResult() {
  const { state } = useLocation();
  const navigate  = useNavigate();
  const [downloading, setDownloading] = useState(false);
  const [dlError, setDlError]         = useState("");

  if (!state?.transcript) {
    return (
      <div className="tr-page">
        <div className="tr-card">
          <p className="tr-empty">Aucun résultat disponible.</p>
          <button className="tr-btn tr-btn--primary" onClick={() => navigate("/")}>
            ← Retour
          </button>
        </div>
      </div>
    );
  }

  const { transcript, inspection, saved, id } = state;
  const sections = inspection?.sections ?? {};

  // Compute summary counts
  let total = 0, conformeCount = 0, nonConformeCount = 0;
  Object.values(sections).forEach((sec) => {
    Object.values(sec).forEach((item) => {
      if (!item || typeof item !== "object") return;
      if (item.conforme === true)  { conformeCount++;    total++; }
      if (item.conforme === false) { nonConformeCount++; total++; }
    });
  });

  // ── Excel download ──────────────────────────────────────────────────────────
  async function handleDownloadExcel() {
    setDownloading(true);
    setDlError("");
    try {
      const res = await fetch("http://localhost:5000/export", {
        method:  "POST",
        headers: { "Content-Type": "application/json" },
        body:    JSON.stringify({ inspection, id }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `Erreur serveur : ${res.status}`);
      }

      // Trigger browser download
      const blob     = await res.blob();
      const url      = URL.createObjectURL(blob);
      const a        = document.createElement("a");
      const code     = inspection?.code_appartement || id || "inspection";
      a.href         = url;
      a.download     = `inspection_${code}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      setDlError("Téléchargement échoué : " + err.message);
    } finally {
      setDownloading(false);
    }
  }

  // ── Render ──────────────────────────────────────────────────────────────────
  return (
    <div className="tr-page">
      <div className="tr-card">

        {/* Header */}
        <div className="tr-header">
          <h1 className="tr-title">Rapport d'Inspection</h1>
          {saved && id && (
            <span className="tr-badge tr-badge--saved">✔ Sauvegardé · ID {id}</span>
          )}
        </div>

        {/* Meta info */}
        {inspection && (
          <div className="tr-meta">
            {[
              ["Date",             inspection.date],
              ["Propriétaire",     inspection.proprietaire],
              ["Contrôleur",       inspection.controleur],
              ["Code appartement", inspection.code_appartement],
              ["Adresse",          inspection.adresse],
            ].filter(([, v]) => v).map(([label, value]) => (
              <div key={label} className="tr-meta__row">
                <span className="tr-meta__label">{label}</span>
                <span className="tr-meta__value">{value}</span>
              </div>
            ))}
          </div>
        )}

        {/* Summary pills */}
        <div className="tr-summary">
          <div className="tr-pill tr-pill--total">
            <span className="tr-pill__num">{total}</span>
            <span className="tr-pill__lbl">Évalués</span>
          </div>
          <div className="tr-pill tr-pill--ok">
            <span className="tr-pill__num">{conformeCount}</span>
            <span className="tr-pill__lbl">Conformes</span>
          </div>
          <div className="tr-pill tr-pill--nok">
            <span className="tr-pill__num">{nonConformeCount}</span>
            <span className="tr-pill__lbl">Non conformes</span>
          </div>
        </div>

        {/* ── Excel download button ── */}
        <div className="tr-actions">
          <button
            className={`tr-btn tr-btn--excel ${downloading ? "tr-btn--loading" : ""}`}
            onClick={handleDownloadExcel}
            disabled={downloading}
          >
            {downloading ? (
              <>
                <span className="tr-spinner" />
                Génération…
              </>
            ) : (
              <>
                <ExcelIcon />
                Télécharger Excel
              </>
            )}
          </button>

          <button className="tr-btn tr-btn--ghost" onClick={() => navigate("/")}>
            ← Nouvelle inspection
          </button>
        </div>

        {dlError && <div className="tr-error">{dlError}</div>}

        {/* Sections */}
        <div className="tr-sections">
          {Object.entries(SECTION_LABELS).map(([sKey, sLabel]) => {
            const items = sections[sKey];
            if (!items || !Object.keys(items).length) return null;
            return (
              <div key={sKey} className="tr-section">
                <h2 className="tr-section__title">{sLabel}</h2>
                <table className="tr-table">
                  <thead>
                    <tr>
                      <th>Élément</th>
                      <th>Conforme ?</th>
                      <th>Commentaire</th>
                    </tr>
                  </thead>
                  <tbody>
                    {Object.entries(items).map(([iKey, iData]) => {
                      const conforme    = iData?.conforme;
                      const commentaire = iData?.commentaire || "";
                      return (
                        <tr key={iKey}>
                          <td>{ITEM_LABELS[iKey] || iKey}</td>
                          <td>
                            {conforme === true  && <span className="tr-badge tr-badge--ok">✔ Conforme</span>}
                            {conforme === false && <span className="tr-badge tr-badge--nok">✘ Non conforme</span>}
                            {conforme === null  && <span className="tr-badge tr-badge--na">—</span>}
                          </td>
                          <td className="tr-comment">{commentaire}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            );
          })}
        </div>

        {/* Transcript accordion */}
        <details className="tr-transcript">
          <summary>Transcription brute</summary>
          <p>{transcript}</p>
        </details>

      </div>
    </div>
  );
}

function ExcelIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">
      <path d="M14 0H5.5L4 1.5V4H1.5L0 5.5v9L1.5 16H10l1.5-1.5V13h2.5L15.5 11.5V2L14 0zm0 11.5H11.5V5.5L10 4H5.5V2h8v9.5zM9.5 14H2V6h7.5v8z"/>
      <path d="M3.5 11.5l1-2 1 2h1l-1.5-2.5L6.5 7h-1l-1 2-1-2h-1l1.5 2.5L3.5 12h1z" fill="white"/>
    </svg>
  );
}