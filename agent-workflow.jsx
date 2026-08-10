import React, { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  Upload, CheckCircle2, AlertTriangle, XCircle, ArrowRight, Loader2,
  Database, FileSpreadsheet, Sparkles, ChevronRight, RotateCcw,
  ShieldCheck, GitMerge, FileText, Truck, Car
} from "lucide-react";

/* ---------- Design tokens ----------
  Background  #0B0D12   panel #12151C   panel-2 #171B24
  border      #232838
  text        #E7EAF2   muted #8891A6
  accent (pipeline indigo) #7C83FD
  signal ok   #4ADE9A   warn #F5B942   error #F1655C
  mono: "JetBrains Mono", body: "Inter"
------------------------------------- */

const ASSURANCE_MAP = {
  "STAR": "STAR", "STAR ASSURANCE": "STAR",
  "GAT": "GAT", "GAT ASSURANCES": "GAT",
  "COMAR": "COMAR", "MAGHREBIA": "MAGHREBIA",
  "LLOYD": "LLOYD TUNISIEN", "LLOYD TUNISIEN": "LLOYD TUNISIEN",
  "AMI": "AMI ASSURANCES", "ASTREE": "ASTREE",
  "BH ASSURANCE": "BH ASSURANCE", "ZITOUNA": "ZITOUNA TAKAFUL",
};
const BRAND_ALIAS = {
  "VW": "VOLKSWAGEN", "VOLKSWAGEN": "VOLKSWAGEN",
  "MERCEDES": "MERCEDES-BENZ", "MERCEDES BENZ": "MERCEDES-BENZ",
  "PEUGEOT": "PEUGEOT", "RENAULT": "RENAULT", "CITROEN": "CITROEN",
  "HYUNDAI": "HYUNDAI", "KIA": "KIA", "TOYOTA": "TOYOTA",
  "FIAT": "FIAT", "SEAT": "SEAT", "SKODA": "SKODA",
};
const INVALID_ENTRIES = ["SEAT", "PASSAGER AVOIR"];

const norm = (s) => (s ?? "").toString().trim().toUpperCase().replace(/\s+/g, " ");

function detectType(headers) {
  const h = headers.map(norm);
  const insuranceHints = ["ASSUREUR", "ASSURANCE", "MARQUE", "REPARATION", "SINISTRE"];
  const supplierHints = ["FOURNISSEUR", "ACHAT", "COMMANDE", "MONTANT TND", "DELAI"];
  const iScore = insuranceHints.filter((k) => h.some((x) => x.includes(k))).length;
  const sScore = supplierHints.filter((k) => h.some((x) => x.includes(k))).length;
  if (iScore === 0 && sScore === 0) return null;
  return iScore >= sScore ? "assurance" : "fournisseur";
}

function guessCol(headers, keywords) {
  const idx = headers.findIndex((h) => keywords.some((k) => norm(h).includes(k)));
  return idx >= 0 ? headers[idx] : null;
}

function rowHash(row) {
  return Object.values(row).join("|").toUpperCase().replace(/\s+/g, "");
}

async function askClaude(prompt) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1000,
      messages: [{ role: "user", content: prompt }],
    }),
  });
  const data = await res.json();
  return (data.content || []).map((b) => b.text || "").join("\n");
}

const STEPS = [
  { key: "upload", label: "Import", icon: Upload },
  { key: "mapping", label: "Détection", icon: FileSpreadsheet },
  { key: "quality", label: "Contrôle qualité", icon: ShieldCheck },
  { key: "merge", label: "Fusion", icon: GitMerge },
  { key: "report", label: "Rapport", icon: FileText },
];

export default function WorkflowAgent() {
  const [stepIdx, setStepIdx] = useState(0);
  const [type, setType] = useState(null);
  const [headers, setHeaders] = useState([]);
  const [rawRows, setRawRows] = useState([]);
  const [colMap, setColMap] = useState({});
  const [normalizedRows, setNormalizedRows] = useState([]);
  const [issues, setIssues] = useState({ duplicates: [], missing: [], invalid: [] });
  const [history, setHistory] = useState([]);
  const [historyLoaded, setHistoryLoaded] = useState(false);
  const [mergePreview, setMergePreview] = useState({ added: 0, skipped: 0 });
  const [report, setReport] = useState("");
  const [loadingReport, setLoadingReport] = useState(false);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const fileInput = useRef(null);

  const storageKey = type === "assurance" ? "wf-assurance-history" : "wf-fournisseur-history";

  useEffect(() => {
    if (!type) return;
    (async () => {
      try {
        const res = await window.storage.get(storageKey);
        setHistory(res ? JSON.parse(res.value) : []);
      } catch {
        setHistory([]);
      }
      setHistoryLoaded(true);
    })();
  }, [type]);

  const reset = () => {
    setStepIdx(0); setType(null); setHeaders([]); setRawRows([]);
    setColMap({}); setNormalizedRows([]); setIssues({ duplicates: [], missing: [], invalid: [] });
    setMergePreview({ added: 0, skipped: 0 }); setReport(""); setFileName(""); setError("");
    setHistoryLoaded(false);
  };

  const handleFile = useCallback((file) => {
    setError("");
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (!json.length) { setError("Le fichier semble vide."); return; }
        const hdrs = Object.keys(json[0]);
        const detected = detectType(hdrs);
        if (!detected) {
          setError("Type de données non reconnu. Vérifie les en-têtes du fichier (Assureur/Marque ou Fournisseur/Achat).");
          return;
        }
        setHeaders(hdrs);
        setRawRows(json);
        setType(detected);

        const map = detected === "assurance"
          ? {
              assureur: guessCol(hdrs, ["ASSUREUR", "ASSURANCE"]),
              marque: guessCol(hdrs, ["MARQUE"]),
              modele: guessCol(hdrs, ["MODELE", "MODÈLE"]),
              montant: guessCol(hdrs, ["MONTANT", "TND"]),
              date: guessCol(hdrs, ["DATE", "MOIS"]),
            }
          : {
              fournisseur: guessCol(hdrs, ["FOURNISSEUR"]),
              montant: guessCol(hdrs, ["MONTANT", "TND"]),
              date: guessCol(hdrs, ["DATE", "MOIS"]),
              delai: guessCol(hdrs, ["DELAI", "DÉLAI"]),
            };
        setColMap(map);
        setStepIdx(1);
      } catch (err) {
        setError("Impossible de lire ce fichier. Vérifie qu'il s'agit bien d'un .xlsx valide.");
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const runNormalization = () => {
    const rows = rawRows.map((r) => {
      const out = { ...r };
      if (type === "assurance" && colMap.assureur) {
        const v = norm(r[colMap.assureur]);
        out.__assureurNorm = ASSURANCE_MAP[v] || r[colMap.assureur];
      }
      if (type === "assurance" && colMap.marque) {
        const v = norm(r[colMap.marque]);
        const prefix = Object.keys(BRAND_ALIAS).find((k) => v.startsWith(k));
        out.__marqueNorm = prefix ? BRAND_ALIAS[prefix] : r[colMap.marque];
      }
      out.__hash = rowHash(r);
      return out;
    });
    setNormalizedRows(rows);

    const seen = new Set();
    const duplicates = [];
    const missing = [];
    const invalid = [];
    const keyCol = type === "assurance" ? colMap.assureur : colMap.fournisseur;

    rows.forEach((r, i) => {
      if (seen.has(r.__hash)) duplicates.push(i); else seen.add(r.__hash);
      if (keyCol && !norm(r[keyCol])) missing.push(i);
      if (colMap.montant && (r[colMap.montant] === "" || r[colMap.montant] == null)) missing.push(i);
      const checkVal = norm(type === "assurance" ? r[colMap.assureur] : r[colMap.fournisseur]);
      if (INVALID_ENTRIES.includes(checkVal)) invalid.push(i);
    });
    setIssues({ duplicates: [...new Set(duplicates)], missing: [...new Set(missing)], invalid: [...new Set(invalid)] });
    setStepIdx(2);
  };

  const goMerge = () => {
    const existingHashes = new Set(history.map((h) => h.__hash));
    const cleanRows = normalizedRows.filter((r, i) =>
      !issues.invalid.includes(i) && !issues.duplicates.includes(i)
    );
    const toAdd = cleanRows.filter((r) => !existingHashes.has(r.__hash));
    setMergePreview({ added: toAdd.length, skipped: cleanRows.length - toAdd.length });
    setStepIdx(3);
  };

  const confirmMerge = async () => {
    const existingHashes = new Set(history.map((h) => h.__hash));
    const cleanRows = normalizedRows.filter((r, i) =>
      !issues.invalid.includes(i) && !issues.duplicates.includes(i)
    );
    const toAdd = cleanRows.filter((r) => !existingHashes.has(r.__hash));
    const merged = [...history, ...toAdd];
    try {
      await window.storage.set(storageKey, JSON.stringify(merged));
      setHistory(merged);
      setStepIdx(4);
      generateReport(merged, toAdd);
    } catch {
      setError("Échec de l'enregistrement dans le stockage.");
    }
  };

  const generateReport = async (merged, added) => {
    setLoadingReport(true);
    setReport("");
    try {
      let summary = "";
      if (type === "assurance") {
        const byAssureur = {};
        merged.forEach((r) => {
          const k = r.__assureurNorm || r[colMap.assureur] || "Inconnu";
          byAssureur[k] = (byAssureur[k] || 0) + 1;
        });
        summary = `Total dossiers cumulés: ${merged.length}. Nouveaux dossiers importés: ${added.length}. Répartition par assureur: ${JSON.stringify(byAssureur)}.`;
      } else {
        const byFournisseur = {};
        merged.forEach((r) => {
          const k = r[colMap.fournisseur] || "Inconnu";
          byFournisseur[k] = (byFournisseur[k] || 0) + 1;
        });
        summary = `Total commandes cumulées: ${merged.length}. Nouvelles commandes importées: ${added.length}. Répartition par fournisseur: ${JSON.stringify(byFournisseur)}.`;
      }
      const prompt = `Tu es un analyste opérations. Voici un résumé de données ${type === "assurance" ? "de réparation véhicule/assurance" : "d'achats fournisseurs"} en Tunisie (TND) après import mensuel:\n${summary}\nRédige un résumé exécutif en français, 4 à 6 phrases, factuel et direct, destiné à un directeur. Mentionne les points notables (concentration, évolution, alertes éventuelles). Pas de préambule, pas de titre, texte brut uniquement.`;
      const text = await askClaude(prompt);
      setReport(text.trim());
    } catch {
      setReport("Le rapport n'a pas pu être généré (connexion à l'API indisponible). Les données ont bien été fusionnées dans l'historique.");
    } finally {
      setLoadingReport(false);
    }
  };

  const totalIssues = issues.duplicates.length + issues.missing.length + issues.invalid.length;

  return (
    <div style={{
      minHeight: "100vh", background: "#0B0D12", color: "#E7EAF2",
      fontFamily: "Inter, -apple-system, sans-serif", padding: "32px 20px",
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap');
        * { box-sizing: border-box; }
        .mono { font-family: 'JetBrains Mono', monospace; }
        .btn { cursor: pointer; border: none; border-radius: 8px; font-weight: 600; transition: all .15s ease; font-family: inherit; }
        .btn-primary { background: #7C83FD; color: #0B0D12; padding: 11px 20px; font-size: 14px; }
        .btn-primary:hover { background: #9198FF; }
        .btn-primary:disabled { background: #2B2F3D; color: #565C6E; cursor: not-allowed; }
        .btn-ghost { background: transparent; color: #8891A6; padding: 10px 16px; font-size: 13px; border: 1px solid #232838; }
        .btn-ghost:hover { border-color: #7C83FD; color: #E7EAF2; }
        .panel { background: #12151C; border: 1px solid #232838; border-radius: 12px; }
        .fade-in { animation: fadeIn .35s ease; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
        .row-hover:hover { background: #171B24; }
        @media (prefers-reduced-motion: reduce) { .fade-in { animation: none; } }
      `}</style>

      <div style={{ maxWidth: 880, margin: "0 auto" }}>
        {/* Header */}
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 28 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 34, height: 34, borderRadius: 8, background: "#7C83FD", display: "flex", alignItems: "center", justifyContent: "center" }}>
              <Sparkles size={18} color="#0B0D12" />
            </div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 16 }}>Agent Workflow</div>
              <div className="mono" style={{ fontSize: 11, color: "#565C6E" }}>import · normalisation · contrôle · fusion · rapport</div>
            </div>
          </div>
          {type && (
            <button className="btn btn-ghost" onClick={reset} style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <RotateCcw size={13} /> Nouveau
            </button>
          )}
        </div>

        {/* Pipeline stepper */}
        <div className="panel" style={{ padding: "18px 20px", marginBottom: 24, display: "flex", alignItems: "center" }}>
          {STEPS.map((s, i) => {
            const Icon = s.icon;
            const active = i === stepIdx;
            const done = i < stepIdx;
            return (
              <React.Fragment key={s.key}>
                <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 6, minWidth: 64 }}>
                  <div style={{
                    width: 34, height: 34, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center",
                    background: done ? "#4ADE9A" : active ? "#7C83FD" : "#171B24",
                    border: active || done ? "none" : "1px solid #232838",
                  }}>
                    {done ? <CheckCircle2 size={16} color="#0B0D12" /> : <Icon size={15} color={active ? "#0B0D12" : "#565C6E"} />}
                  </div>
                  <span style={{ fontSize: 11, color: active ? "#E7EAF2" : "#565C6E", fontWeight: active ? 600 : 400 }}>{s.label}</span>
                </div>
                {i < STEPS.length - 1 && (
                  <div style={{ flex: 1, height: 1, background: i < stepIdx ? "#4ADE9A" : "#232838", margin: "0 4px 20px" }} />
                )}
              </React.Fragment>
            );
          })}
        </div>

        {error && (
          <div className="panel fade-in" style={{ padding: 14, marginBottom: 16, borderColor: "#F1655C", display: "flex", gap: 10, alignItems: "flex-start" }}>
            <XCircle size={16} color="#F1655C" style={{ flexShrink: 0, marginTop: 2 }} />
            <span style={{ fontSize: 13, color: "#E7EAF2" }}>{error}</span>
          </div>
        )}

        {/* STEP 0: Upload */}
        {stepIdx === 0 && (
          <div className="panel fade-in" style={{ padding: 40 }}>
            <div
              onDrop={(e) => { e.preventDefault(); if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); }}
              onDragOver={(e) => e.preventDefault()}
              onClick={() => fileInput.current?.click()}
              style={{
                border: "1.5px dashed #232838", borderRadius: 10, padding: "48px 20px",
                textAlign: "center", cursor: "pointer",
              }}
            >
              <Upload size={28} color="#7C83FD" style={{ marginBottom: 12 }} />
              <div style={{ fontWeight: 600, marginBottom: 4 }}>Dépose ton fichier Excel mensuel</div>
              <div style={{ fontSize: 13, color: "#8891A6" }}>Fournisseurs ou assurance véhicule — détection automatique du type</div>
              <input ref={fileInput} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
                onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])} />
            </div>
            <div style={{ display: "flex", gap: 20, marginTop: 24, fontSize: 12, color: "#565C6E" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 6 }}><Truck size={13} /> Achats fournisseurs</div>
              <div style={{ display: "flex", alignItems: "center", gap: 6 }}><Car size={13} /> Réparations assurance</div>
            </div>
          </div>
        )}

        {/* STEP 1: Detection / mapping */}
        {stepIdx === 1 && (
          <div className="panel fade-in" style={{ padding: 28 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
              {type === "assurance" ? <Car size={18} color="#7C83FD" /> : <Truck size={18} color="#7C83FD" />}
              <span style={{ fontWeight: 600 }}>
                Type détecté : {type === "assurance" ? "Assurance véhicule" : "Fournisseurs"}
              </span>
            </div>
            <div className="mono" style={{ fontSize: 12, color: "#565C6E", marginBottom: 20 }}>
              {fileName} · {rawRows.length} lignes · {headers.length} colonnes
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 24 }}>
              {Object.entries(colMap).map(([k, v]) => (
                <div key={k} style={{ background: "#171B24", borderRadius: 8, padding: "10px 12px" }}>
                  <div style={{ fontSize: 11, color: "#565C6E", textTransform: "uppercase", marginBottom: 3 }}>{k}</div>
                  <div className="mono" style={{ fontSize: 13, color: v ? "#E7EAF2" : "#F1655C" }}>{v || "non trouvée"}</div>
                </div>
              ))}
            </div>
            <button className="btn btn-primary" onClick={runNormalization} style={{ display: "flex", alignItems: "center", gap: 6 }}>
              Normaliser & analyser <ArrowRight size={14} />
            </button>
          </div>
        )}

        {/* STEP 2: Quality control */}
        {stepIdx === 2 && (
          <div className="panel fade-in" style={{ padding: 28 }}>
            <div style={{ fontWeight: 600, marginBottom: 18 }}>Contrôle qualité</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginBottom: 22 }}>
              {[
                { label: "Doublons", n: issues.duplicates.length, icon: AlertTriangle, color: "#F5B942" },
                { label: "Champs manquants", n: issues.missing.length, icon: AlertTriangle, color: "#F5B942" },
                { label: "Entrées invalides", n: issues.invalid.length, icon: XCircle, color: "#F1655C" },
              ].map((c) => (
                <div key={c.label} style={{ background: "#171B24", borderRadius: 10, padding: 16 }}>
                  <c.icon size={15} color={c.n > 0 ? c.color : "#4ADE9A"} />
                  <div className="mono" style={{ fontSize: 26, fontWeight: 700, marginTop: 8 }}>{c.n}</div>
                  <div style={{ fontSize: 12, color: "#8891A6" }}>{c.label}</div>
                </div>
              ))}
            </div>
            {totalIssues === 0 ? (
              <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: "#4ADE9A", marginBottom: 20 }}>
                <CheckCircle2 size={15} /> Aucune anomalie détectée — données prêtes à fusionner.
              </div>
            ) : (
              <div style={{ fontSize: 13, color: "#8891A6", marginBottom: 20, lineHeight: 1.6 }}>
                Les doublons et entrées invalides (ex. {INVALID_ENTRIES.join(", ")}) seront exclus automatiquement de la fusion.
                Les lignes à champs manquants sont conservées mais signalées.
              </div>
            )}
            <button className="btn btn-primary" onClick={goMerge} style={{ display: "flex", alignItems: "center", gap: 6 }}>
              Continuer vers la fusion <ArrowRight size={14} />
            </button>
          </div>
        )}

        {/* STEP 3: Merge preview */}
        {stepIdx === 3 && (
          <div className="panel fade-in" style={{ padding: 28 }}>
            <div style={{ fontWeight: 600, marginBottom: 6 }}>Fusion avec l'historique</div>
            <div style={{ fontSize: 13, color: "#8891A6", marginBottom: 20 }}>
              Historique actuel : <span className="mono">{history.length}</span> lignes {historyLoaded ? "" : "(chargement...)"}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 24 }}>
              <div style={{ background: "#171B24", borderRadius: 10, padding: 16 }}>
                <div className="mono" style={{ fontSize: 26, fontWeight: 700, color: "#4ADE9A" }}>+{mergePreview.added}</div>
                <div style={{ fontSize: 12, color: "#8891A6" }}>nouvelles lignes à ajouter</div>
              </div>
              <div style={{ background: "#171B24", borderRadius: 10, padding: 16 }}>
                <div className="mono" style={{ fontSize: 26, fontWeight: 700, color: "#565C6E" }}>{mergePreview.skipped}</div>
                <div style={{ fontSize: 12, color: "#8891A6" }}>déjà présentes (ignorées)</div>
              </div>
            </div>
            <button className="btn btn-primary" onClick={confirmMerge} style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <GitMerge size={14} /> Confirmer la fusion
            </button>
          </div>
        )}

        {/* STEP 4: Report */}
        {stepIdx === 4 && (
          <div className="panel fade-in" style={{ padding: 28 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
              <CheckCircle2 size={17} color="#4ADE9A" />
              <span style={{ fontWeight: 600 }}>Fusion terminée</span>
            </div>
            <div className="mono" style={{ fontSize: 12, color: "#565C6E", marginBottom: 20 }}>
              Historique {type === "assurance" ? "assurance" : "fournisseurs"} : {history.length} lignes cumulées
            </div>

            <div style={{ fontSize: 13, fontWeight: 600, color: "#8891A6", marginBottom: 10, display: "flex", alignItems: "center", gap: 6 }}>
              <FileText size={14} /> Résumé exécutif
            </div>
            {loadingReport ? (
              <div style={{ display: "flex", alignItems: "center", gap: 8, color: "#8891A6", fontSize: 13, padding: "20px 0" }}>
                <Loader2 size={15} className="mono" style={{ animation: "spin 1s linear infinite" }} />
                Génération du rapport...
              </div>
            ) : (
              <div style={{ background: "#171B24", borderRadius: 10, padding: 18, fontSize: 14, lineHeight: 1.7, color: "#E7EAF2" }}>
                {report}
              </div>
            )}
            <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
          </div>
        )}
      </div>
    </div>
  );
}
