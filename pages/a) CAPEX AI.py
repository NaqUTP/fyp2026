# ======================================================================================
# CAPEX AI RT2026  -  FULL APP (single file, Streamlit)
# Models: Random Forest + Gradient Boosting + MLP (Deep Learning)
#
# This build FIXES and IMPROVES the previous version:
#   FIX A: Monte Carlo now runs on REAL base feature values (dataset means or user
#          overrides), not empty {} that produced all-NaN simulations.
#   FIX B: Monte Carlo tab is usable directly from any trained dataset, not only from
#          a Project, and shows per-feature base values you can edit before simulating.
#   FIX C: Prediction inputs are always visible once a model is trained, with clearer
#          guidance and a base-values helper.
#   ADD : Facility-type aware note on cost drivers; cleaner PETRONAS teal UI.
#
# Run:  streamlit run "capex_ai.py"
# ======================================================================================

import io
import json
import re
import requests
import numpy as np
import pandas as pd
import streamlit as st

try:
    from sklearn.impute import KNNImputer, SimpleImputer
    from sklearn.model_selection import train_test_split
    from sklearn.preprocessing import StandardScaler
    from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
    from sklearn.dummy import DummyRegressor
    from sklearn.pipeline import Pipeline
    from sklearn.metrics import mean_squared_error, r2_score, mean_absolute_error
except Exception as e:
    st.error(f"Missing dependency: scikit-learn.\n\nAdd scikit-learn to requirements.txt.\n\nDetails: {e}")
    st.stop()

try:
    import torch
    import torch.nn as nn
    from torch.utils.data import DataLoader, TensorDataset
    TORCH_AVAILABLE = True
except ImportError:
    TORCH_AVAILABLE = False

import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="CAPEX AI RT2026", page_icon="💠", layout="wide", initial_sidebar_state="expanded")

PETRONAS = {"teal": "#00A19B", "teal_dark": "#008C87", "purple": "#6C4DD3",
            "white": "#FFFFFF", "black": "#0E1116", "border": "rgba(0,0,0,0.10)"}

SHAREPOINT_LINKS = {
    "Shallow Water": "https://petronas.sharepoint.com/sites/your-site/shallow-water",
    "Deep Water": "https://petronas.sharepoint.com/sites/your-site/deep-water",
    "Onshore": "https://petronas.sharepoint.com/sites/your-site/onshore",
    "Uncon": "https://petronas.sharepoint.com/sites/your-site/uncon",
    "CCS": "https://petronas.sharepoint.com/sites/your-site/ccs",
}

st.markdown(f"""<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
:root{{
  --bg:#0E1116; --panel:#171B22; --panel2:#1E232C; --elev:#232A34;
  --text:#E6E9EF; --muted:#9AA3B2; --border:rgba(255,255,255,.10);
  --teal:#00A19B; --purple:#6C4DD3;
}}
html,body,[class*="css"]{{font-family:'Inter',sans-serif;}}
.stApp,[data-testid="stAppViewContainer"]{{background:var(--bg) !important;color:var(--text) !important;padding-top:.5rem;}}
[data-testid="stHeader"]{{background:transparent !important;}}
#MainMenu,footer{{visibility:hidden;}}
h1,h2,h3,h4,h5,h6,p,span,label,li,div,.stMarkdown{{color:var(--text) !important;}}
.stCaption,[data-testid="stCaptionContainer"],small{{color:var(--muted) !important;}}

/* sidebar */
[data-testid="stSidebar"]{{background:linear-gradient(180deg,#12161C 0%,#0E1116 100%) !important;border-right:1px solid var(--border);}}
[data-testid="stSidebar"] *{{color:var(--text) !important;}}

/* hero */
.petronas-hero{{border-radius:20px;padding:26px 30px;margin:6px 0 18px 0;color:#fff !important;background:linear-gradient(135deg,{PETRONAS["teal"]},{PETRONAS["purple"]},#0A0D12);background-size:200% 200%;animation:heroGradient 8s ease-in-out infinite;box-shadow:0 10px 30px rgba(0,0,0,.5);border:1px solid var(--border);}}
@keyframes heroGradient{{0%{{background-position:0% 50%}}50%{{background-position:100% 50%}}100%{{background-position:0% 50%}}}}
.petronas-hero h1{{margin:0 0 5px;font-weight:800;letter-spacing:.3px;color:#fff !important;}}
.petronas-hero p{{margin:0;opacity:.92;font-weight:500;color:#fff !important;}}

/* buttons */
.stButton>button,.stDownloadButton>button,.petronas-button{{border-radius:10px;padding:.6rem 1.1rem;font-weight:600;color:#fff !important;border:none;background:linear-gradient(to right,{PETRONAS["teal"]},{PETRONAS["purple"]});background-size:200% auto;transition:background-position .85s ease,transform .2s ease,box-shadow .25s ease;text-decoration:none;display:inline-block;}}
.stButton>button:hover,.stDownloadButton>button:hover,.petronas-button:hover{{background-position:right center;transform:translateY(-1px);box-shadow:0 6px 18px rgba(0,161,155,.35);}}

/* inputs: dark fields with light text */
[data-testid="stTextInput"] input,[data-testid="stNumberInput"] input,textarea,.stTextArea textarea{{background:var(--panel2) !important;color:var(--text) !important;border:1px solid var(--border) !important;border-radius:9px !important;}}
[data-baseweb="select"]>div{{background:var(--panel2) !important;border:1px solid var(--border) !important;color:var(--text) !important;}}
[data-baseweb="select"] *{{color:var(--text) !important;}}
[data-baseweb="popover"],[role="listbox"]{{background:var(--panel2) !important;color:var(--text) !important;}}
[data-testid="stTextInput"] input::placeholder,textarea::placeholder{{color:#5B6472 !important;}}
[data-testid="stFileUploaderDropzone"]{{background:var(--panel2) !important;border:1px dashed var(--border) !important;color:var(--text) !important;}}
[data-testid="stFileUploaderDropzone"] *{{color:var(--text) !important;}}
/* the "Browse files" button inside the uploader (a secondary button) */
[data-testid="stFileUploaderDropzone"] button,[data-testid="baseButton-secondary"]{{background:linear-gradient(to right,{PETRONAS["teal"]},{PETRONAS["purple"]}) !important;color:#fff !important;border:none !important;border-radius:8px !important;font-weight:600 !important;}}
[data-testid="stFileUploaderDropzone"] button *,[data-testid="baseButton-secondary"] *{{color:#fff !important;}}
[data-testid="stFileUploaderDropzone"] button svg{{fill:#fff !important;stroke:#fff !important;}}
[data-testid="stFileUploaderDropzone"] small{{color:var(--muted) !important;}}
[data-testid="stWidgetLabel"] label,[data-testid="stWidgetLabel"] p{{color:var(--text) !important;}}

/* metrics */
[data-testid="stMetric"]{{background:var(--panel) !important;border:1px solid var(--border);border-radius:12px;padding:12px 14px;}}
[data-testid="stMetricValue"]{{color:var(--text) !important;}}
[data-testid="stMetricLabel"]{{color:var(--muted) !important;}}

/* tabs */
.stTabs [role="tablist"]{{display:flex;gap:8px;border-bottom:none;padding-bottom:6px;flex-wrap:wrap;}}
.stTabs [role="tab"]{{background:var(--panel) !important;color:var(--text) !important;border-radius:8px;padding:10px 18px;border:1px solid var(--border);font-weight:600;transition:all .3s ease;}}
.stTabs [role="tab"] p{{color:var(--text) !important;}}
.stTabs [role="tab"]:hover{{background:linear-gradient(to right,{PETRONAS["teal"]},{PETRONAS["purple"]}) !important;}}
.stTabs [role="tab"][aria-selected="true"]{{background:linear-gradient(to right,{PETRONAS["teal"]},{PETRONAS["purple"]}) !important;border-color:transparent;box-shadow:0 4px 16px rgba(0,0,0,.4);}}
.stTabs [role="tab"][aria-selected="true"] p{{color:#fff !important;}}

/* dataframes + tables */
[data-testid="stDataFrame"],[data-testid="stTable"]{{background:var(--panel) !important;border:1px solid var(--border);border-radius:10px;}}
[data-testid="stExpander"]{{background:var(--panel) !important;border:1px solid var(--border) !important;border-radius:10px;}}
[data-testid="stExpander"] summary,[data-testid="stExpander"] p{{color:var(--text) !important;}}

/* alerts stay readable on dark */
[data-testid="stAlert"]{{border-radius:10px;}}

/* chat */
[data-testid="stChatMessage"]{{background:var(--panel) !important;border:1px solid var(--border);border-radius:12px;}}
.stChatInput textarea,[data-testid="stChatInput"] textarea{{background:var(--panel2) !important;color:var(--text) !important;}}

/* radio / checkbox labels */
[data-testid="stRadio"] label,[data-testid="stCheckbox"] label{{color:var(--text) !important;}}
hr{{border-color:var(--border) !important;}}
</style>""", unsafe_allow_html=True)

st.markdown("""<div class="petronas-hero"><h1>CAPEX AI RT2026</h1><p>Data-driven CAPEX prediction · Random Forest, Gradient Boosting &amp; MLP</p></div>""", unsafe_allow_html=True)

# ---- session state ----------------------------------------------------------
for key, default in [("datasets", {}), ("predictions", {}), ("processed_excel_files", set()),
                     ("_last_metrics", None), ("projects", {}), ("uploader_nonce", 0), ("widget_nonce", 0)]:
    if key not in st.session_state:
        st.session_state[key] = default


# ---- helpers ----------------------------------------------------------------
def toast(msg, icon="✅"):
    try:
        st.toast(f"{icon} {msg}")
    except Exception:
        st.success(msg)

def is_junk_col(colname):
    h = str(colname).strip().upper()
    return (not h) or h.startswith("UNNAMED") or h in {"INDEX", "IDX"}

def currency_from_header(header):
    h = (header or "").strip().upper()
    if "€" in h: return "€"
    if "£" in h: return "£"
    if "$" in h: return "$"
    if re.search(r"\bUSD\b", h): return "USD"
    if re.search(r"\b(MYR|RM)\b", h): return "RM"
    return ""

def get_currency_symbol(df, target_col=None):
    if df is None or df.empty: return ""
    if target_col and target_col in df.columns:
        return currency_from_header(str(target_col))
    for c in reversed(df.columns):
        if not is_junk_col(c):
            return currency_from_header(str(c))
    return ""

def cost_breakdown(base_pred, sst_pct, owners_pct, cont_pct, esc_pct):
    base_pred = float(base_pred)
    owners_cost = round(base_pred * (owners_pct / 100), 2)
    sst_cost    = round(base_pred * (sst_pct / 100), 2)
    contingency = round((base_pred + owners_cost) * (cont_pct / 100), 2)
    escalation  = round((base_pred + owners_cost) * (esc_pct / 100), 2)
    grand_total = round(base_pred + owners_cost + sst_cost + contingency + escalation, 2)
    return owners_cost, sst_cost, contingency, escalation, grand_total

def project_components_df(proj):
    rows = []
    for c in proj.get("components", []):
        rows.append({"Component": c["component_type"], "Dataset": c["dataset"],
                     "Base CAPEX": float(c["prediction"]),
                     "Owner's Cost": float(c["breakdown"]["owners_cost"]),
                     "Contingency": float(c["breakdown"]["contingency_cost"]),
                     "Escalation": float(c["breakdown"]["escalation_cost"]),
                     "SST": float(c["breakdown"]["sst_cost"]),
                     "Grand Total": float(c["breakdown"]["grand_total"])})
    return pd.DataFrame(rows)

def project_totals(proj):
    dfc = project_components_df(proj)
    if dfc.empty:
        return {k: 0.0 for k in ("capex_sum", "owners", "cont", "esc", "sst", "grand_total")}
    return {"capex_sum": float(dfc["Base CAPEX"].sum()), "owners": float(dfc["Owner's Cost"].sum()),
            "cont": float(dfc["Contingency"].sum()), "esc": float(dfc["Escalation"].sum()),
            "sst": float(dfc["SST"].sum()), "grand_total": float(dfc["Grand Total"].sum())}

def dataset_feature_means(ds_name):
    """Real base values for Monte Carlo: the mean of each feature column.
    This is the fix for the all-NaN Monte Carlo bug."""
    df = st.session_state.datasets.get(ds_name)
    fcols = st.session_state.get(f"feature_cols__{ds_name}", [])
    means = {}
    if df is not None:
        for c in fcols:
            if c in df.columns:
                v = pd.to_numeric(df[c], errors="coerce").mean()
                means[c] = float(v) if pd.notna(v) else 0.0
            else:
                means[c] = 0.0
    return means


# ---- data preprocessing -----------------------------------------------------

# =============================================================================
# DATA PREPARATION ENGINE
# Maps messy raw columns to the canonical cost-driver schema per facility type,
# converts units, drops identifier/metadata columns, and flags out-of-range rows.
# =============================================================================
PREP_SCHEMAS = {
    "Pipeline": {
        "features": {
            "Diameter_mm":  ["diameter", "dia", "od", "nominal size", "ppl_size", "size"],
            "Capacity_bpd": ["capacity", "throughput", "flow", "boe", "bpd"],
            "Length_km":    ["length", "len", "distance", "km"],
        },
        "target_aliases": ["cost", "capex", "price", "mmusd"],
        "units": {"Diameter_mm": [("in", 25.4), ("inch", 25.4)], "Length_km": [("m", 0.001), ("mile", 1.60934)]},
        "ranges": {"Diameter_mm": (10, 2000), "Capacity_bpd": (0, 5_000_000), "Length_km": (0, 5000)},
    },
    "WHP": {
        "features": {
            "Water_Depth_m":     ["water depth", "depth", "wtr_dpth", "wd"],
            "Num_Wells":         ["num wells", "wells", "well count", "n_wells", "slots"],
            "Topsides_Weight_t": ["topside", "topsides weight", "deck weight"],
            "Jacket_Weight_t":   ["jacket", "jacket weight", "substructure"],
            "Is_Unmanned":       ["unmanned", "manned", "is_unmanned"],
            "Remoteness_km":     ["remoteness", "distance to shore", "shore distance"],
        },
        "target_aliases": ["capex", "cost", "mmusd"],
        "units": {"Water_Depth_m": [("ft", 0.3048), ("feet", 0.3048)],
                  "Topsides_Weight_t": [("kg", 0.001), ("lb", 0.000453592)],
                  "Jacket_Weight_t": [("kg", 0.001), ("lb", 0.000453592)]},
        "ranges": {"Water_Depth_m": (0, 500), "Num_Wells": (0, 60), "Topsides_Weight_t": (0, 60000),
                   "Jacket_Weight_t": (0, 60000), "Is_Unmanned": (0, 1), "Remoteness_km": (0, 2000)},
    },
    "CPP": {
        "features": {
            "Water_Depth_m":       ["water depth", "depth", "wtr_dpth", "wd"],
            "Structure_Type":      ["structure", "structure type", "type", "fcl_kind", "kind"],
            "Slot_Capacity_Gross": ["slot", "slot capacity", "slots", "num wells", "wells"],
            "Design_Life_years":   ["design life", "lifetime", "design_life"],
            "Has_Drilling":        ["drilling", "has_drilling", "drill"],
            "Has_Storage":         ["storage", "has_storage"],
            "Has_Quarter":         ["quarter", "quarters", "accommodation"],
        },
        "target_aliases": ["capex", "cost", "mmusd"],
        "units": {"Water_Depth_m": [("ft", 0.3048), ("feet", 0.3048)]},
        "ranges": {"Water_Depth_m": (0, 2000), "Slot_Capacity_Gross": (0, 80), "Design_Life_years": (0, 60),
                   "Has_Drilling": (0, 1), "Has_Storage": (0, 1), "Has_Quarter": (0, 1)},
    },
}
PREP_DROP_HINTS = ["year", "date", "startup", "sanction", "vintage", "id", "identifier", "uid", "guid",
                   "npdid", "api", "name", "field", "block", "area", "lease", "complex", "segment",
                   "status", "phase", "basis", "source", "operator", "owner name", "note", "comment",
                   "remark", "url", "link", "code"]

def _prep_norm(s):
    return re.sub(r"[^a-z0-9 ]", " ", str(s).lower()).strip()

def _prep_is_meta(header):
    nh = _prep_norm(header)
    return any(re.search(rf"\b{re.escape(h)}\b", nh) or nh == h for h in PREP_DROP_HINTS)

def _prep_to_num(series):
    def one(v):
        if pd.isna(v): return np.nan
        parts = [p for p in re.split(r"[;,/]", str(v)) if re.search(r"\d", p)]
        nums = []
        for p in parts:
            p = re.sub(r"[^0-9.\-]", "", p)
            try: nums.append(float(p))
            except Exception: pass
        if not nums:
            c = re.sub(r"[^0-9.\-]", "", str(v))
            try: return float(c)
            except Exception: return np.nan
        return sum(nums) / len(nums)
    return series.map(one)

def _prep_match(headers, aliases):
    normed = {h: _prep_norm(h) for h in headers}
    for h, nh in normed.items():
        if nh in aliases: return h
    for h, nh in normed.items():
        for a in aliases:
            if a in nh or nh in a: return h
    return None

def prep_detect(df):
    headers = list(df.columns)
    scores = {}
    for ft, spec in PREP_SCHEMAS.items():
        hits = sum(1 for al in spec["features"].values() if _prep_match(headers, al) is not None)
        scores[ft] = round(hits / len(spec["features"]), 2)
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else None, scores

def prep_run(df, facility_type):
    spec = PREP_SCHEMAS[facility_type]
    headers = list(df.columns)
    report = {"mapped": {}, "dropped_metadata": [], "extra_kept": [], "unmapped": [], "flagged": {}, "conversions": []}
    out = pd.DataFrame()
    for canon, aliases in spec["features"].items():
        raw = _prep_match(headers, aliases)
        if raw is None:
            report["unmapped"].append(canon); continue
        report["mapped"][raw] = canon
        col = df[raw].astype(str) if canon == "Structure_Type" else _prep_to_num(df[raw])
        if canon in spec.get("units", {}):
            nh = _prep_norm(raw)
            for hint, factor in spec["units"][canon]:
                if re.search(rf"\b{re.escape(hint)}\b", nh):
                    col = col * factor
                    report["conversions"].append(f"{raw}: x{factor} -> {canon}"); break
        out[canon] = col
    # target
    tgt = _prep_match(headers, spec["target_aliases"])
    if tgt is not None:
        report["mapped"][tgt] = "CAPEX_MMUSD"
        t = _prep_to_num(df[tgt])
        if re.search(r"\busd\b", _prep_norm(tgt)) and "mm" not in _prep_norm(tgt):
            if t.median(skipna=True) and t.median(skipna=True) > 10000:
                t = t / 1e6; report["conversions"].append(f"{tgt}: /1e6 -> CAPEX_MMUSD")
        out["CAPEX_MMUSD"] = t
    # unknown columns: drop metadata/ids, keep genuine numeric extras
    known = set(report["mapped"].keys())
    for h in headers:
        if h in known: continue
        if _prep_is_meta(h):
            report["dropped_metadata"].append(h); continue
        num = _prep_to_num(df[h])
        if num.notna().mean() > 0.6:
            safe = re.sub(r"[^0-9a-zA-Z]+", "_", str(h)).strip("_")
            out[f"extra__{safe}"] = num; report["extra_kept"].append(h)
        else:
            report["unmapped"].append(h)
    # range flags
    for canon, (lo, hi) in spec.get("ranges", {}).items():
        if canon in out.columns:
            bad = ((out[canon] < lo) | (out[canon] > hi)) & out[canon].notna()
            if bad.any(): report["flagged"][canon] = int(bad.sum())
    # move target last, drop rows with no target
    if "CAPEX_MMUSD" in out.columns:
        out = out[out["CAPEX_MMUSD"].notna()].copy()
        cols = [c for c in out.columns if c != "CAPEX_MMUSD"] + ["CAPEX_MMUSD"]
        out = out[cols]
    report["final_rows"] = len(out)
    return out, report


class DataPreprocessor:
    @staticmethod
    def clean_dataframe(df):
        df = df.copy()
        bad = [c for c in df.columns if is_junk_col(c)]
        if bad:
            df = df.drop(columns=bad)
        return df

    @staticmethod
    def extract_features_target(df):
        if df is None or df.empty:
            raise ValueError("Empty dataset")
        target_col = df.columns[-1]
        feature_cols = [c for c in df.columns if c != target_col]
        if not feature_cols:
            raise ValueError("No feature columns found")
        X = df[feature_cols].copy()
        y = pd.to_numeric(df[target_col], errors="coerce")
        if y.isna().sum() / len(y) > 0.8:
            raise ValueError(f"Target column '{target_col}' has too many missing values")
        return X, y, target_col

    @staticmethod
    def validate_feature_columns(X):
        X = X.copy()
        for col in X.columns:
            if X[col].dtype == object:
                X[col] = pd.to_numeric(X[col], errors="coerce")
        return X


# ---- MLP --------------------------------------------------------------------
class CapexMLP(nn.Module if TORCH_AVAILABLE else object):
    def __init__(self, n_features):
        if not TORCH_AVAILABLE:
            raise ImportError("PyTorch not installed.")
        super().__init__()
        self.net = nn.Sequential(
            nn.Linear(n_features, 128), nn.BatchNorm1d(128), nn.ReLU(), nn.Dropout(0.3),
            nn.Linear(128, 64), nn.BatchNorm1d(64), nn.ReLU(), nn.Dropout(0.2),
            nn.Linear(64, 32), nn.ReLU(), nn.Linear(32, 1))
    def forward(self, x):
        return self.net(x).squeeze(1)

class MLPWrapper:
    def __init__(self, n_features, epochs=200, lr=0.001, batch_size=32, patience=20, random_state=42):
        self.n_features = n_features; self.epochs = epochs; self.lr = lr
        self.batch_size = batch_size; self.patience = patience; self.random_state = random_state
        self.model = None; self.scaler_X = StandardScaler(); self.scaler_y = StandardScaler()
        self.train_losses = []; self.val_losses = []; self.imputer = SimpleImputer(strategy="median")

    def fit(self, X, y):
        torch.manual_seed(self.random_state); np.random.seed(self.random_state)
        Xi = self.imputer.fit_transform(X)
        Xs = self.scaler_X.fit_transform(Xi).astype(np.float32)
        ys = self.scaler_y.fit_transform(y.reshape(-1, 1)).ravel().astype(np.float32)
        nv = max(1, int(len(Xs) * 0.15))
        Xtr, Xv, ytr, yv = Xs[:-nv], Xs[-nv:], ys[:-nv], ys[-nv:]
        loader = DataLoader(TensorDataset(torch.from_numpy(Xtr), torch.from_numpy(ytr)),
                            batch_size=self.batch_size, shuffle=True)
        self.model = CapexMLP(self.n_features)
        opt = torch.optim.Adam(self.model.parameters(), lr=self.lr)
        sch = torch.optim.lr_scheduler.ReduceLROnPlateau(opt, patience=10, factor=0.5)
        crit = nn.MSELoss(); best_val = float("inf"); best_state = None; no_imp = 0
        self.train_losses = []; self.val_losses = []
        for _ in range(self.epochs):
            self.model.train(); ep = 0.0
            for Xb, yb in loader:
                opt.zero_grad(); loss = crit(self.model(Xb), yb); loss.backward(); opt.step()
                ep += loss.item() * len(Xb)
            ep /= len(Xtr); self.model.eval()
            with torch.no_grad():
                vl = crit(self.model(torch.from_numpy(Xv)), torch.from_numpy(yv)).item()
            self.train_losses.append(ep); self.val_losses.append(vl); sch.step(vl)
            if vl < best_val:
                best_val = vl; best_state = {k: v.clone() for k, v in self.model.state_dict().items()}; no_imp = 0
            else:
                no_imp += 1
                if no_imp >= self.patience:
                    break
        if best_state:
            self.model.load_state_dict(best_state)
        self.model.eval(); return self

    def predict(self, X):
        Xi = self.imputer.transform(X)
        Xs = self.scaler_X.transform(Xi).astype(np.float32)
        with torch.no_grad():
            ys = self.model(torch.from_numpy(Xs)).numpy()
        return self.scaler_y.inverse_transform(ys.reshape(-1, 1)).ravel()


# ---- model pipeline ---------------------------------------------------------
class ModelPipeline:
    MODEL_CANDIDATES = {
        "RandomForest": lambda rs=42: RandomForestRegressor(n_estimators=200, random_state=rs, n_jobs=-1),
        "GradientBoosting": lambda rs=42: GradientBoostingRegressor(n_estimators=200, learning_rate=0.05,
                                                                    max_depth=4, subsample=0.8, random_state=rs),
    }

    @classmethod
    def create_pipeline(cls, model_name, random_state=42):
        if model_name not in cls.MODEL_CANDIDATES:
            model_name = "RandomForest"
        ctor = cls.MODEL_CANDIDATES[model_name]
        try:
            model = ctor(random_state)
        except Exception:
            model = ctor()
        return Pipeline([("imputer", SimpleImputer(strategy="median")), ("model", model)])

    @classmethod
    @st.cache_resource(show_spinner=False)
    def train_all_cached(_cls, X, y, test_size=0.20, random_state=42,
                         mlp_epochs=200, mlp_lr=0.001, mlp_batch=32, mlp_patience=20):
        Xa = X.values.astype(np.float32); ya = y.values.astype(np.float32)
        Xtr, Xte, ytr, yte = train_test_split(Xa, ya, test_size=test_size, random_state=random_state)
        results = {}
        for name in ("RandomForest", "GradientBoosting"):
            pipe = _cls.create_pipeline(name, random_state); pipe.fit(Xtr, ytr); yp = pipe.predict(Xte)
            results[name] = {"pipeline": pipe, "r2": round(float(r2_score(yte, yp)), 4),
                             "rmse": round(float(np.sqrt(mean_squared_error(yte, yp))), 4),
                             "mae": round(float(mean_absolute_error(yte, yp)), 4),
                             "y_test": yte, "y_pred": yp, "type": "sklearn"}
        if TORCH_AVAILABLE and len(Xtr) >= 20:
            mlp = MLPWrapper(Xtr.shape[1], mlp_epochs, mlp_lr, mlp_batch, mlp_patience, random_state)
            mlp.fit(Xtr, ytr); ypm = mlp.predict(Xte)
            results["MLP"] = {"pipeline": mlp, "r2": round(float(r2_score(yte, ypm)), 4),
                              "rmse": round(float(np.sqrt(mean_squared_error(yte, ypm))), 4),
                              "mae": round(float(mean_absolute_error(yte, ypm)), 4),
                              "y_test": yte, "y_pred": ypm,
                              "train_losses": mlp.train_losses, "val_losses": mlp.val_losses, "type": "mlp"}
        else:
            reason = "PyTorch not installed." if not TORCH_AVAILABLE else "Too few rows for MLP (need 20+)."
            results["MLP"] = {"pipeline": None, "r2": None, "rmse": None, "mae": None,
                              "y_test": yte, "y_pred": np.zeros_like(yte),
                              "train_losses": [], "val_losses": [], "type": "mlp", "error": reason}
        # baseline: predict the mean (reference point for R2)
        dummy = DummyRegressor(strategy="mean").fit(Xtr, ytr)
        base_r2 = round(float(r2_score(yte, dummy.predict(Xte))), 4)
        valid = {k: v for k, v in results.items() if v["r2"] is not None}
        best = max(valid, key=lambda k: valid[k]["r2"]); bm = valid[best]
        return {"rf": results["RandomForest"], "gb": results["GradientBoosting"], "mlp": results["MLP"],
                "best": best, "pipeline": bm["pipeline"], "feature_cols": list(X.columns),
                "model": best, "r2": bm["r2"], "rmse": bm["rmse"], "mae": bm["mae"],
                "baseline_r2": base_r2}

    @staticmethod
    def prepare_prediction_input(feature_cols, payload):
        row = {}
        for col in feature_cols:
            val = payload.get(col, np.nan)
            if val is None or (isinstance(val, str) and val.strip() == ""):
                row[col] = np.nan
            elif isinstance(val, (int, float, np.number)):
                row[col] = float(val)
            else:
                try:
                    row[col] = float(val)
                except Exception:
                    row[col] = np.nan
        return pd.DataFrame([row], columns=feature_cols)


def monte_carlo_simulation(pipeline, feature_cols, base_values, n_simulations=1000, feature_uncertainty=0.10):
    """Perturb each base feature value with Gaussian noise and predict each scenario.
    base_values MUST contain real numbers (see dataset_feature_means)."""
    np.random.seed(42)
    base_array = np.array([float(base_values.get(c, 0.0)) for c in feature_cols], dtype=float)
    base_array = np.nan_to_num(base_array, nan=0.0)
    preds = []
    for _ in range(int(n_simulations)):
        noise = np.random.normal(0, feature_uncertainty, len(base_array))
        sim = base_array * (1 + noise)
        sim_df = pd.DataFrame([sim], columns=feature_cols)
        try:
            preds.append(float(pipeline.predict(sim_df)[0]))
        except Exception:
            preds.append(float("nan"))
    return pd.DataFrame({"prediction": preds}).dropna()


# ---- navigation shortcuts ---------------------------------------------------
for col, label in zip(st.columns(5), ["SHALLOW WATER", "DEEP WATER", "ONSHORE", "UNCON", "CCS"]):
    with col:
        url = SHAREPOINT_LINKS.get(label.title(), "#")
        st.markdown(f'<a href="{url}" target="_blank" rel="noopener" class="petronas-button" '
                    f'style="width:100%;text-align:center;display:inline-block;">{label}</a>',
                    unsafe_allow_html=True)

tab_prep, tab_data, tab_pb, tab_mc, tab_compare, tab_ai = st.tabs(
    ["🧹 Data Preparation", "📊 Data & Models", "🏗️ Project Builder", "🎲 Monte Carlo", "🔀 Compare Projects", "🤖 AI Advisor"])



# =============================================================================
# TAB 0 - DATA PREPARATION
# =============================================================================
with tab_prep:
    st.markdown('<h3 style="margin-top:0;color:#E6E9EF;">🧹 Data Preparation</h3>', unsafe_allow_html=True)
    st.caption("Turn a messy raw CSV into a clean, training-ready dataset. The tool detects the "
               "facility type, maps columns to the correct cost drivers, converts units, and drops "
               "identifier columns such as names, ids and years.")

    raw_file = st.file_uploader("Upload a raw CSV to clean", type="csv",
                                key=f"prep_uploader_{st.session_state.uploader_nonce}")
    if raw_file is not None:
        try:
            raw_df = pd.read_csv(raw_file)
        except Exception as e:
            st.error(f"Could not read CSV: {e}"); raw_df = None

        if raw_df is not None:
            st.markdown("##### Raw preview")
            st.dataframe(raw_df.head(8), use_container_width=True)

            suggested, scores = prep_detect(raw_df)
            score_str = "  ·  ".join(f"{k} {v}" for k, v in sorted(scores.items(), key=lambda kv: -kv[1]))
            st.caption(f"Detection scores: {score_str}")

            types = list(PREP_SCHEMAS.keys())
            default_idx = types.index(suggested) if suggested in types else 0
            colf1, colf2 = st.columns([1, 1])
            with colf1:
                facility_type = st.selectbox("Facility type (confirm or override)", types, index=default_idx,
                                             key="prep_facility_type")
            with colf2:
                store_name = st.text_input("Save cleaned dataset as", value=f"{facility_type}_clean.csv",
                                           key="prep_store_name")

            drivers = list(PREP_SCHEMAS[facility_type]["features"].keys())
            st.caption(f"Cost drivers for **{facility_type}**: {', '.join(drivers)}  →  target: CAPEX_MMUSD")

            if st.button("Run cleaning", type="primary", key="prep_run_btn"):
                try:
                    clean_df, rep = prep_run(raw_df, facility_type)
                    st.session_state["_prep_clean_df"] = clean_df
                    st.session_state["_prep_report"] = rep
                    st.session_state["_prep_store_name"] = store_name
                    toast("Cleaning complete.")
                except Exception as e:
                    st.error(f"Cleaning failed: {e}")

    # show result if present
    if st.session_state.get("_prep_clean_df") is not None:
        clean_df = st.session_state["_prep_clean_df"]
        rep = st.session_state["_prep_report"]
        st.divider()
        st.markdown("##### Cleaning summary")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Final rows", rep.get("final_rows", len(clean_df)))
        m2.metric("Cost drivers mapped", len([v for v in rep["mapped"].values() if v != "CAPEX_MMUSD"]))
        m3.metric("Dropped (metadata)", len(rep["dropped_metadata"]))
        m4.metric("Kept for review", len(rep["extra_kept"]))

        detail = []
        for raw, canon in rep["mapped"].items():
            detail.append({"Raw column": raw, "Resolved as": canon, "Action": "cost driver" if canon != "CAPEX_MMUSD" else "target"})
        for c in rep["dropped_metadata"]:
            detail.append({"Raw column": c, "Resolved as": "—", "Action": "dropped: identifier/metadata"})
        for c in rep["extra_kept"]:
            detail.append({"Raw column": c, "Resolved as": f"extra__{re.sub(r'[^0-9a-zA-Z]+','_',c).strip('_')}", "Action": "kept: review before training"})
        for c in rep["unmapped"]:
            detail.append({"Raw column": c, "Resolved as": "—", "Action": "unmapped (left out)"})
        st.dataframe(pd.DataFrame(detail), use_container_width=True)

        if rep["conversions"]:
            st.caption("Unit conversions applied: " + "; ".join(rep["conversions"]))
        if rep["flagged"]:
            st.warning("Rows outside expected range (flagged, not removed): " +
                       ", ".join(f"{k}: {v}" for k, v in rep["flagged"].items()))
        if rep["extra_kept"]:
            st.info("Columns kept as 'extra__' are unknown numeric fields that might be real cost drivers. "
                    "Review them; drop any that are actually another form of the cost (target leakage).")

        st.markdown("##### Clean preview")
        st.dataframe(clean_df.head(10), use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇️ Download clean CSV", data=clean_df.to_csv(index=False),
                               file_name=st.session_state.get("_prep_store_name", "clean.csv"),
                               mime="text/csv", key="prep_download_btn")
        with c2:
            if st.button("➡️ Send to Data & Models", key="prep_send_btn"):
                name = st.session_state.get("_prep_store_name", "clean.csv")
                st.session_state.datasets[name] = clean_df
                st.session_state.predictions.setdefault(name, [])
                toast(f"'{name}' added to Data & Models."); st.rerun()


# =============================================================================
# TAB 1 - DATA & MODELS
# =============================================================================
with tab_data:
    st.markdown('<h3 style="margin-top:0;color:#E6E9EF;">📁 Data</h3>', unsafe_allow_html=True)
    uploaded_files = st.file_uploader("Upload CSV files (the last column is treated as the CAPEX target)",
                                      type="csv", accept_multiple_files=True,
                                      key=f"csv_uploader_{st.session_state.uploader_nonce}")
    if uploaded_files:
        for up in uploaded_files:
            if up.name not in st.session_state.datasets:
                try:
                    df = DataPreprocessor.clean_dataframe(pd.read_csv(up))
                    st.session_state.datasets[up.name] = df
                    st.session_state.predictions.setdefault(up.name, [])
                except Exception as e:
                    st.error(f"Failed to read {up.name}: {e}")
        toast("Dataset(s) added.")

    cA, cB = st.columns(2)
    with cA:
        if st.button("🧹 Clear predictions", key="clear_preds_btn"):
            st.session_state.predictions = {k: [] for k in st.session_state.predictions}
            toast("Predictions cleared.", "🧹"); st.rerun()
    with cB:
        if st.button("🗂️ Clear all data", key="clear_datasets_btn"):
            st.session_state.datasets = {}; st.session_state.predictions = {}
            st.session_state.processed_excel_files = set(); st.session_state._last_metrics = None
            st.session_state.uploader_nonce += 1; st.session_state.widget_nonce += 1
            toast("All data cleared.", "🗂️"); st.rerun()

    st.divider()

    if not st.session_state.datasets:
        st.info("Upload a dataset to begin. The last column is used as the CAPEX target.")
    else:
        ds_name_data = st.selectbox("Active dataset", list(st.session_state.datasets.keys()), key="active_dataset_data")
        df_active = st.session_state.datasets[ds_name_data]
        target_col_active = df_active.columns[-1]
        currency_active = get_currency_symbol(df_active, target_col_active)
        colA, colB, colC, colD2 = st.columns(4)
        colA.metric("Rows", f"{df_active.shape[0]:,}")
        colB.metric("Columns", f"{df_active.shape[1]:,}")
        colC.metric("Currency", currency_active or "—")
        colD2.caption(f"Target column: **{target_col_active}**")
        with st.expander("Preview (first 10 rows)", expanded=False):
            st.dataframe(df_active.head(10), use_container_width=True)

        # ---- training ----
        st.divider()
        st.markdown('<h3 style="margin-top:0;color:#E6E9EF;">⚙️ Model Training</h3>', unsafe_allow_html=True)
        ds_name_model = st.selectbox("Dataset for training", list(st.session_state.datasets.keys()), key="ds_model")
        df_model = st.session_state.datasets[ds_name_model]
        data_ok = False
        try:
            X, y, target_col = DataPreprocessor.extract_features_target(df_model)
            X = DataPreprocessor.validate_feature_columns(X)
            st.success(f"Ready — **{X.shape[1]} features**, target: **{target_col}**")
            c1, c2, c3 = st.columns(3)
            c1.metric("Features", X.shape[1]); c2.metric("Samples", X.shape[0])
            valid_n = int(y.notna().sum()); c3.metric("Valid targets", f"{valid_n} ({valid_n/len(y)*100:.0f}%)")
            data_ok = True
        except Exception as e:
            st.error(f"Data preparation failed: {e}")

        if data_ok:
            split_col, btn_col = st.columns([3, 1])
            with split_col:
                test_size = st.slider("Test set size", 0.10, 0.40, 0.20, 0.05, key="train_test_size")
                train_pct = round((1 - test_size) * 100); test_pct = round(test_size * 100)
                st.caption(f"Train {train_pct}%  ·  Test {test_pct}%")
            with btn_col:
                st.write(""); st.write("")
                run_train = st.button("🚀 Train RF, GB & MLP", key="run_training_btn", type="primary")

            with st.expander("🧠 MLP Hyperparameters", expanded=False):
                if not TORCH_AVAILABLE:
                    st.warning("PyTorch not installed — MLP will be skipped. Add `torch` to requirements.txt.")
                mc1, mc2, mc3, mc4 = st.columns(4)
                with mc1: mlp_epochs = st.number_input("Max Epochs", 50, 500, 200, 50, key="mlp_epochs")
                with mc2: mlp_lr = st.select_slider("Learning Rate", [0.0001, 0.0005, 0.001, 0.005, 0.01], value=0.001, key="mlp_lr")
                with mc3: mlp_batch = st.selectbox("Batch Size", [16, 32, 64, 128], index=1, key="mlp_batch")
                with mc4: mlp_patience = st.number_input("Early Stop Patience", 5, 50, 20, 5, key="mlp_patience")

            if run_train:
                try:
                    with st.spinner("Training Random Forest, Gradient Boosting, and MLP…"):
                        metrics = ModelPipeline.train_all_cached(X, y, float(test_size), 42,
                                                                 int(mlp_epochs), float(mlp_lr),
                                                                 int(mlp_batch), int(mlp_patience))
                    st.session_state._last_metrics = metrics
                    st.session_state[f"trained_model__{ds_name_model}"] = metrics
                    st.session_state[f"current_pipeline__{ds_name_model}"] = metrics["pipeline"]
                    st.session_state[f"feature_cols__{ds_name_model}"] = metrics["feature_cols"]
                    try:
                        knn = KNNImputer(n_neighbors=5); knn.fit(X)
                        st.session_state[f"knn_imputer_{ds_name_model}"] = knn
                    except Exception:
                        pass
                    toast("Training complete! 🎉")

                    rf, gb, mlp = metrics["rf"], metrics["gb"], metrics["mlp"]
                    mlp_r2 = mlp["r2"] if mlp["r2"] is not None else float("nan")
                    mlp_rmse = mlp["rmse"] if mlp["rmse"] is not None else float("nan")
                    mlp_mae = mlp["mae"] if mlp["mae"] is not None else float("nan")
                    compare_df = pd.DataFrame({
                        "Metric": ["R² Score ↑", "RMSE ↓", "MAE ↓"],
                        "Random Forest": [rf["r2"], rf["rmse"], rf["mae"]],
                        "Gradient Boosting": [gb["r2"], gb["rmse"], gb["mae"]],
                        "MLP (Deep Learning)": [mlp_r2, mlp_rmse, mlp_mae]})
                    st.markdown("##### Model Comparison — RF vs GB vs MLP")
                    st.dataframe(compare_df, use_container_width=True, hide_index=True)
                    st.caption(f"Baseline (predict the mean) R² = {metrics['baseline_r2']}. "
                               f"A useful model should clearly beat this.")

                    winner = metrics["best"]
                    winner_label = {"RandomForest": "Random Forest", "GradientBoosting": "Gradient Boosting",
                                    "MLP": "MLP (Deep Learning)"}.get(winner, winner)
                    st.success(f"**{winner_label}** selected as active model (highest R²)")
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Model", winner_label); m2.metric("R²", f"{metrics['r2']:.4f}")
                    m3.metric("RMSE", f"{metrics['rmse']:,.2f}"); m4.metric("MAE", f"{metrics['mae']:,.2f}")

                    st.markdown("##### Actual vs Predicted — All Models")
                    fig_scatter = go.Figure(); all_vals = []
                    for key, label, colour in [("rf", "Random Forest", "#00A19B"),
                                               ("gb", "Gradient Boosting", "#6C4DD3"),
                                               ("mlp", "MLP", "#F4801A")]:
                        m = metrics[key]
                        if m["r2"] is None: continue
                        fig_scatter.add_trace(go.Scatter(x=m["y_test"], y=m["y_pred"], mode="markers",
                                                         marker=dict(color=colour, opacity=0.55, size=6), name=label))
                        all_vals.extend(list(m["y_test"])); all_vals.extend(list(m["y_pred"]))
                    if all_vals:
                        lo, hi = float(min(all_vals)), float(max(all_vals))
                        fig_scatter.add_trace(go.Scatter(x=[lo, hi], y=[lo, hi], mode="lines",
                                                         line=dict(color="#AAB2C0", dash="dash", width=1.5), name="Perfect fit"))
                    fig_scatter.update_layout(xaxis_title="Actual CAPEX", yaxis_title="Predicted CAPEX",
                                              height=400, margin=dict(l=0, r=0, t=10, b=0),
                                              paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                                              legend=dict(orientation="h", y=-0.18))
                    st.plotly_chart(fig_scatter, use_container_width=True)

                    st.markdown("##### Feature Importance — RF vs GB")
                    fi_left, fi_right = st.columns(2)
                    for container, (label, bkey) in zip([fi_left, fi_right],
                                                        [("Random Forest", "rf"), ("Gradient Boosting", "gb")]):
                        pipe = metrics[bkey]["pipeline"]
                        fi_df = pd.DataFrame({"Feature": metrics["feature_cols"],
                                              "Importance": pipe.named_steps["model"].feature_importances_}
                                             ).sort_values("Importance", ascending=True)
                        fig_fi = go.Figure(go.Bar(x=fi_df["Importance"], y=fi_df["Feature"], orientation="h",
                                                  marker_color="#00A19B" if bkey == "rf" else "#6C4DD3"))
                        fig_fi.update_layout(title=label, xaxis_title="Importance",
                                             height=max(260, 32 * len(fi_df)), margin=dict(l=0, r=0, t=35, b=0),
                                             paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                        with container:
                            st.plotly_chart(fig_fi, use_container_width=True)
                except Exception as e:
                    st.error(f"Training failed: {e}")

        # ---- PREDICT (inputs always visible once a model exists) ----
        st.divider()
        st.markdown('<h3 style="margin-top:0;color:#E6E9EF;">🎯 Predict CAPEX</h3>', unsafe_allow_html=True)
        ds_name_pred = st.selectbox("Dataset for prediction", list(st.session_state.datasets.keys()), key="ds_pred")
        df_pred = st.session_state.datasets[ds_name_pred]

        if f"current_pipeline__{ds_name_pred}" not in st.session_state:
            st.warning("Train a model for this dataset first (Model Training section above).")
        else:
            pipeline = st.session_state[f"current_pipeline__{ds_name_pred}"]
            feature_cols = st.session_state[f"feature_cols__{ds_name_pred}"]
            target_col = df_pred.columns[-1]
            currency_pred = get_currency_symbol(df_pred, target_col)
            meta = st.session_state.get(f"trained_model__{ds_name_pred}", {})
            active_model = meta.get("best", "—")
            active_label = {"RandomForest": "Random Forest", "GradientBoosting": "Gradient Boosting",
                            "MLP": "MLP (Deep Learning)"}.get(active_model, active_model)
            st.info(f"Active model: **{active_label}**  ·  R² {meta.get('r2', 0):.4f}")

            cf1, cf2 = st.columns(2)
            with cf1:
                sst_pct = st.number_input("SST (%)", 0.0, 100.0, 0.0, 0.5, key="pred_sst")
                owners_pct = st.number_input("Owner's Cost (%)", 0.0, 100.0, 0.0, 0.5, key="pred_owner")
            with cf2:
                cont_pct = st.number_input("Contingency (%)", 0.0, 100.0, 0.0, 0.5, key="pred_cont")
                esc_pct = st.number_input("Escalation (%)", 0.0, 100.0, 0.0, 0.5, key="pred_esc")

            project_name = st.text_input("Project name", placeholder="e.g. Offshore Pipeline Replacement 2026", key="pred_project_name")

            st.markdown("##### Feature values")
            fill_means = st.checkbox("Prefill with dataset average values", value=True, key="prefill_means")
            means = dataset_feature_means(ds_name_pred)
            st.caption(f"Enter values for **{len(feature_cols)}** features. Blank = imputed with the median.")
            input_values = {}
            for i in range(0, len(feature_cols), 3):
                cols = st.columns(3)
                for j, feat in enumerate(feature_cols[i:i + 3]):
                    with cols[j]:
                        default = f"{means.get(feat, 0.0):.2f}" if fill_means else ""
                        val = st.text_input(feat, value=default, key=f"input_{feat}_{ds_name_pred}")
                        if str(val).strip() in ("", "nan"):
                            input_values[feat] = np.nan
                        else:
                            try:
                                input_values[feat] = float(val)
                            except Exception:
                                input_values[feat] = np.nan

            if st.button("Run Prediction", key="run_pred_btn", type="primary"):
                try:
                    pred_input = ModelPipeline.prepare_prediction_input(feature_cols, input_values)
                    base_pred = float(pipeline.predict(pred_input)[0])
                    owners_cost, sst_cost, contingency, escalation, grand_total = cost_breakdown(
                        base_pred, sst_pct, owners_pct, cont_pct, esc_pct)
                    result = {"Project Name": project_name or "Untitled", "Model Used": active_label,
                              "Base CAPEX": round(base_pred, 2), "Owner's Cost": owners_cost,
                              "SST Cost": sst_cost, "Contingency": contingency, "Escalation": escalation,
                              "Grand Total": grand_total}
                    for col in feature_cols:
                        result[col] = pred_input[col].iloc[0]
                    st.session_state.predictions.setdefault(ds_name_pred, []).append(result)
                    toast("Prediction added!")
                    r1, r2c, r3, r4, r5 = st.columns(5)
                    r1.metric("Base CAPEX", f"{currency_pred} {base_pred:,.2f}")
                    r2c.metric("Owner's Cost", f"{currency_pred} {owners_cost:,.2f}")
                    r3.metric("SST", f"{currency_pred} {sst_cost:,.2f}")
                    r4.metric("Contingency", f"{currency_pred} {contingency:,.2f}")
                    r5.metric("Grand Total", f"{currency_pred} {grand_total:,.2f}")
                except Exception as e:
                    st.error(f"Prediction failed: {e}")

        # ---- RESULTS + DOWNLOAD ----
        st.divider()
        st.markdown('<h3 style="margin-top:0;color:#E6E9EF;">📄 Results</h3>', unsafe_allow_html=True)
        ds_name_res = st.selectbox("Dataset for results", list(st.session_state.datasets.keys()), key="ds_results")
        preds = st.session_state.predictions.get(ds_name_res, [])
        if preds:
            df_preds = pd.DataFrame(preds)
            display_cols = [c for c in ["Project Name", "Model Used", "Base CAPEX", "Owner's Cost",
                                        "SST Cost", "Contingency", "Escalation", "Grand Total"] if c in df_preds.columns]
            st.dataframe(df_preds[display_cols], use_container_width=True, height=300)
            c1, c2 = st.columns(2)
            with c1:
                bio = io.BytesIO(); df_preds.to_excel(bio, index=False, engine="openpyxl"); bio.seek(0)
                st.download_button("⬇️ Download Excel", data=bio, file_name=f"{ds_name_res}_predictions.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key="download_excel_btn")
            with c2:
                st.download_button("⬇️ Download CSV", data=df_preds.to_csv(index=False),
                                   file_name=f"{ds_name_res}_predictions.csv", mime="text/csv", key="download_csv_btn")
        else:
            st.info("No predictions yet.")


# =============================================================================
# TAB 2 - PROJECT BUILDER
# =============================================================================
with tab_pb:
    st.markdown('<h4 style="margin-top:0;color:#E6E9EF;">🏗️ Project Builder</h4>', unsafe_allow_html=True)
    st.caption("Assemble multi-component CAPEX projects from trained models.")
    if not st.session_state.datasets:
        st.info("No datasets loaded. Load data in the Data & Models tab first.")
    else:
        colA, colB = st.columns([2, 1])
        with colA:
            new_proj = st.text_input("New project name", placeholder="e.g. CAPEX 2026", key="pb_new_project_name")
        with colB:
            st.write("")
            if new_proj and new_proj not in st.session_state.projects:
                if st.button("Create project", key="pb_create_project_btn"):
                    st.session_state.projects[new_proj] = {"components": [], "currency": "",
                                                           "cost_factors": {"sst_pct": 0.0, "owners_pct": 0.0,
                                                                            "cont_pct": 0.0, "esc_pct": 0.0}}
                    toast(f"Project '{new_proj}' created."); st.rerun()

        if not st.session_state.projects:
            st.info("Create a project above, then add components.")
        else:
            proj_sel = st.selectbox("Select project", list(st.session_state.projects.keys()), key="pb_project_select")
            proj = st.session_state.projects[proj_sel]

            st.markdown("##### Project cost factors")
            cf1, cf2 = st.columns(2)
            with cf1:
                proj["cost_factors"]["sst_pct"] = st.number_input("SST (%)", 0.0, 100.0, proj["cost_factors"].get("sst_pct", 0.0), 0.5, key=f"pb_sst_{proj_sel}")
                proj["cost_factors"]["owners_pct"] = st.number_input("Owner's Cost (%)", 0.0, 100.0, proj["cost_factors"].get("owners_pct", 0.0), 0.5, key=f"pb_owners_{proj_sel}")
            with cf2:
                proj["cost_factors"]["cont_pct"] = st.number_input("Contingency (%)", 0.0, 100.0, proj["cost_factors"].get("cont_pct", 0.0), 0.5, key=f"pb_cont_{proj_sel}")
                proj["cost_factors"]["esc_pct"] = st.number_input("Escalation (%)", 0.0, 100.0, proj["cost_factors"].get("esc_pct", 0.0), 0.5, key=f"pb_esc_{proj_sel}")

            st.markdown("##### Add component")
            dataset_for_comp = st.selectbox("Dataset for component", sorted(st.session_state.datasets.keys()), key="pb_dataset_for_component")
            df_comp = st.session_state.datasets[dataset_for_comp]
            curr_comp = get_currency_symbol(df_comp, df_comp.columns[-1])
            if f"current_pipeline__{dataset_for_comp}" not in st.session_state:
                st.warning(f"Train a model for '{dataset_for_comp}' in the Data & Models tab first.")
            else:
                pipeline_comp = st.session_state[f"current_pipeline__{dataset_for_comp}"]
                feat_comp = st.session_state[f"feature_cols__{dataset_for_comp}"]
                meta_comp = st.session_state.get(f"trained_model__{dataset_for_comp}", {})
                label_comp = {"RandomForest": "Random Forest", "GradientBoosting": "Gradient Boosting",
                              "MLP": "MLP (Deep Learning)"}.get(meta_comp.get("best", "—"), meta_comp.get("best", "—"))
                st.info(f"Model: **{label_comp}**  ·  R² {meta_comp.get('r2', 0):.4f}")
                component_type = st.text_input("Component type", placeholder="e.g. Pipeline, Platform, FPSO", key=f"pb_component_type_{proj_sel}")
                means_comp = dataset_feature_means(dataset_for_comp)
                comp_inputs = {}
                for i in range(0, len(feat_comp), 2):
                    cols = st.columns(2)
                    for j, feat in enumerate(feat_comp[i:i + 2]):
                        with cols[j]:
                            val = st.text_input(feat, value=f"{means_comp.get(feat, 0.0):.2f}", key=f"pb_{feat}_{proj_sel}_{dataset_for_comp}")
                            if str(val).strip() in ("", "nan"):
                                comp_inputs[feat] = np.nan
                            else:
                                try:
                                    comp_inputs[feat] = float(val)
                                except Exception:
                                    comp_inputs[feat] = np.nan
                if st.button("➕ Add component", key=f"pb_add_comp_{proj_sel}"):
                    if not component_type:
                        st.error("Enter a component type.")
                    else:
                        try:
                            pi = ModelPipeline.prepare_prediction_input(feat_comp, comp_inputs)
                            bp = float(pipeline_comp.predict(pi)[0])
                            cf = proj["cost_factors"]
                            oc, sc, cc, ec, gt = cost_breakdown(bp, cf["sst_pct"], cf["owners_pct"], cf["cont_pct"], cf["esc_pct"])
                            proj["components"].append({"component_type": component_type, "dataset": dataset_for_comp,
                                                       "model_used": meta_comp.get("best", "—"), "prediction": bp,
                                                       "breakdown": {"owners_cost": oc, "sst_cost": sc, "contingency_cost": cc,
                                                                     "escalation_cost": ec, "grand_total": gt}})
                            proj["currency"] = curr_comp
                            toast(f"Component '{component_type}' added."); st.rerun()
                        except Exception as e:
                            st.error(f"Failed to add component: {e}")

            comps = proj.get("components", [])
            if comps:
                st.markdown("##### Components")
                df_comps = pd.DataFrame([{"Component": c["component_type"], "Dataset": c["dataset"],
                                          "Model": c.get("model_used", "—"),
                                          "Base CAPEX": f"{curr_comp} {c['prediction']:,.2f}",
                                          "Grand Total": f"{curr_comp} {c['breakdown']['grand_total']:,.2f}"} for c in comps])
                st.dataframe(df_comps, use_container_width=True)
                totals = project_totals(proj)
                t1, t2, t3 = st.columns(3)
                t1.metric("Total Base CAPEX", f"{curr_comp} {totals['capex_sum']:,.2f}")
                t2.metric("Total SST", f"{curr_comp} {totals['sst']:,.2f}")
                t3.metric("Grand Total", f"{curr_comp} {totals['grand_total']:,.2f}")
                for idx, comp in enumerate(comps):
                    c1, c2, c3 = st.columns([3, 2, 1])
                    c1.write(f"**{comp['component_type']}** — {comp.get('model_used', '—')}")
                    c2.write(f"GT: {curr_comp} {comp['breakdown']['grand_total']:,.2f}")
                    with c3:
                        if st.button("🗑️", key=f"del_comp_{proj_sel}_{idx}"):
                            comps.pop(idx); st.rerun()
                proj_json = json.dumps(proj, indent=2, default=float)
                st.download_button("⬇️ Download project (JSON)", data=proj_json, file_name=f"{proj_sel}.json",
                                   mime="application/json", key=f"dl_json_{proj_sel}")
            else:
                st.info("No components yet.")


# =============================================================================
# TAB 3 - MONTE CARLO  (FIXED: runs on real base values, usable per-dataset)
# =============================================================================
with tab_mc:
    st.markdown('<h3 style="margin-top:0;color:#E6E9EF;">🎲 Monte Carlo Analysis</h3>', unsafe_allow_html=True)
    st.caption("Simulate cost uncertainty by perturbing the feature values around a base case.")

    trained = [ds for ds in st.session_state.datasets if f"current_pipeline__{ds}" in st.session_state]
    if not trained:
        st.info("Train a model in the Data & Models tab first, then simulate here.")
    else:
        mode = st.radio("Simulate from", ["A trained dataset", "A project"], horizontal=True, key="mc_mode")

        if mode == "A trained dataset":
            ds_mc = st.selectbox("Trained dataset", trained, key="mc_dataset")
            pipe = st.session_state[f"current_pipeline__{ds_mc}"]
            fcols = st.session_state[f"feature_cols__{ds_mc}"]
            means = dataset_feature_means(ds_mc)

            st.markdown("##### Base case (edit any value; defaults are dataset averages)")
            base_values = {}
            for i in range(0, len(fcols), 3):
                cols = st.columns(3)
                for j, feat in enumerate(fcols[i:i + 3]):
                    with cols[j]:
                        base_values[feat] = st.number_input(feat, value=float(round(means.get(feat, 0.0), 2)),
                                                             key=f"mc_base_{feat}_{ds_mc}")

            mc1, mc2, mc3 = st.columns(3)
            with mc1: n_sims = st.number_input("Simulations", 100, 20000, 1000, 100, key="mc_n_sims")
            with mc2: feat_unc = st.slider("Feature uncertainty (%)", 1.0, 50.0, 10.0, 1.0, key="mc_feat_unc")
            with mc3: budget = st.number_input("Budget threshold", 0.0, value=float(round(pipe.predict(pd.DataFrame([[base_values[c] for c in fcols]], columns=fcols))[0] * 1.2, 2)), step=10.0, key="mc_budget")

            if st.button("Run Monte Carlo", type="primary", key="mc_run"):
                try:
                    with st.spinner("Running simulations…"):
                        sims = monte_carlo_simulation(pipe, fcols, base_values, int(n_sims), feat_unc / 100)
                    vals = sims["prediction"].values
                    if len(vals) == 0:
                        st.warning("No valid simulations generated.")
                    else:
                        p50, p80, p90 = np.percentile(vals, 50), np.percentile(vals, 80), np.percentile(vals, 90)
                        exceed = (vals > budget).mean() * 100 if budget > 0 else 0.0
                        rc1, rc2, rc3, rc4 = st.columns(4)
                        rc1.metric("P50", f"{p50:,.1f}"); rc2.metric("P80", f"{p80:,.1f}")
                        rc3.metric("P90", f"{p90:,.1f}"); rc4.metric(f"P(> {budget:,.0f})", f"{exceed:.1f}%")
                        fig = px.histogram(x=vals, nbins=50, title="Cost distribution",
                                           labels={"x": "Predicted CAPEX", "y": "Frequency"},
                                           color_discrete_sequence=["#00A19B"])
                        if budget > 0:
                            fig.add_vline(x=budget, line_dash="dash", line_color="red",
                                          annotation_text=f"Budget: {budget:,.0f}")
                        fig.add_vline(x=p50, line_dash="dot", line_color="#6C4DD3", annotation_text="P50")
                        st.plotly_chart(fig, use_container_width=True)
                except Exception as e:
                    st.error(f"Monte Carlo failed: {e}")

        else:  # project mode
            proj_trained = {p: pr for p, pr in st.session_state.projects.items() if pr.get("components")}
            if not proj_trained:
                st.info("Build a project with components first (Project Builder tab).")
            else:
                proj_sel_mc = st.selectbox("Project", list(proj_trained.keys()), key="mc_project_select")
                comps_mc = st.session_state.projects[proj_sel_mc]["components"]
                mc1, mc2, mc3 = st.columns(3)
                with mc1: n_sims = st.number_input("Simulations", 100, 20000, 1000, 100, key="mc_pn_sims")
                with mc2: feat_unc = st.slider("Feature uncertainty (%)", 1.0, 50.0, 10.0, 1.0, key="mc_pfeat_unc")
                with mc3: budget = st.number_input("Budget threshold", 0.0, value=1000.0, step=10.0, key="mc_pbudget")
                if st.button("Run Monte Carlo", type="primary", key="mc_prun"):
                    try:
                        with st.spinner("Running simulations…"):
                            all_sims = []
                            for comp in comps_mc:
                                ds = comp["dataset"]
                                if f"current_pipeline__{ds}" not in st.session_state:
                                    st.warning(f"No trained model for {ds}"); continue
                                pipe = st.session_state[f"current_pipeline__{ds}"]
                                fcols = st.session_state[f"feature_cols__{ds}"]
                                # FIX: use the dataset means as the base case, not empty {}
                                base = dataset_feature_means(ds)
                                sims = monte_carlo_simulation(pipe, fcols, base, int(n_sims), feat_unc / 100)
                                all_sims.append(sims["prediction"].values)
                        if all_sims:
                            L = min(len(a) for a in all_sims)
                            total = np.sum([a[:L] for a in all_sims], axis=0)
                            p50, p80, p90 = np.percentile(total, 50), np.percentile(total, 80), np.percentile(total, 90)
                            exceed = (total > budget).mean() * 100 if budget > 0 else 0.0
                            rc1, rc2, rc3, rc4 = st.columns(4)
                            rc1.metric("P50", f"{p50:,.0f}"); rc2.metric("P80", f"{p80:,.0f}")
                            rc3.metric("P90", f"{p90:,.0f}"); rc4.metric(f"P(> {budget:,.0f})", f"{exceed:.1f}%")
                            fig = px.histogram(x=total, nbins=50, title="Total project cost distribution",
                                               labels={"x": "Total CAPEX", "y": "Frequency"},
                                               color_discrete_sequence=["#00A19B"])
                            fig.add_vline(x=budget, line_dash="dash", line_color="red", annotation_text=f"Budget: {budget:,.0f}")
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.warning("No valid simulations generated.")
                    except Exception as e:
                        st.error(f"Monte Carlo failed: {e}")


# =============================================================================
# TAB 4 - COMPARE PROJECTS
# =============================================================================
with tab_compare:
    st.markdown('<h3 style="margin-top:0;color:#E6E9EF;">🔀 Compare Projects</h3>', unsafe_allow_html=True)
    if len(st.session_state.projects) < 2:
        st.info("Create at least 2 projects in the Project Builder to compare.")
    else:
        proj_names = list(st.session_state.projects.keys())
        sel_projs = st.multiselect("Select projects to compare", proj_names, default=proj_names[:2])
        if len(sel_projs) < 2:
            st.warning("Select at least 2 projects.")
        else:
            cmp_data = []
            for pn in sel_projs:
                t = project_totals(st.session_state.projects[pn])
                cmp_data.append({"Project": pn, "Components": len(st.session_state.projects[pn].get("components", [])),
                                 "Base CAPEX": t["capex_sum"], "SST": t["sst"], "Owner's Cost": t["owners"],
                                 "Contingency": t["cont"], "Escalation": t["esc"], "Grand Total": t["grand_total"]})
            df_cmp = pd.DataFrame(cmp_data)
            st.dataframe(df_cmp, use_container_width=True)
            viz_type = st.selectbox("Chart type", ["Bar Chart", "Stacked Bar"], key="viz_type")
            if viz_type == "Bar Chart":
                fig_cmp = px.bar(df_cmp, x="Project", y="Grand Total", title="Grand Total by Project",
                                 text="Grand Total", color_discrete_sequence=["#00A19B"])
                fig_cmp.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
            else:
                melt = df_cmp.melt(id_vars=["Project"], value_vars=["Base CAPEX", "SST", "Owner's Cost", "Contingency", "Escalation"],
                                   var_name="Cost Type", value_name="Amount")
                fig_cmp = px.bar(melt, x="Project", y="Amount", color="Cost Type", title="Cost Breakdown by Project", barmode="stack")
            st.plotly_chart(fig_cmp, use_container_width=True)


# =============================================================================
# TAB 5 - AI ADVISOR
# =============================================================================
with tab_ai:
    st.markdown('<h3 style="margin-top:0;color:#E6E9EF;">🤖 AI CAPEX Advisor</h3>', unsafe_allow_html=True)
    st.caption("Ask about CAPEX, cost drivers, what-if scenarios, or project risks.")

    if "ai_messages" not in st.session_state: st.session_state.ai_messages = []
    if "ai_backend" not in st.session_state: st.session_state.ai_backend = "Anthropic API (Claude)"

    with st.expander("⚙️ AI backend settings", expanded=False):
        backend = st.radio("Choose AI backend", ["Anthropic API (Claude)", "Ollama (local, open-source)"],
                           horizontal=True, key="ai_backend_select")
        st.session_state.ai_backend = backend
        if backend == "Anthropic API (Claude)":
            st.caption("Reads `anthropic_api_key` from Streamlit secrets (.streamlit/secrets.toml).")
        else:
            st.caption("Requires Ollama running locally (ollama serve; ollama pull llama3).")
            ollama_model = st.selectbox("Ollama model", ["llama3", "mistral", "deepseek-r1:7b", "qwen3:8b"], key="ollama_model_select")
            ollama_url = st.text_input("Ollama base URL", value="http://localhost:11434", key="ollama_url")

    def build_context_summary():
        lines = []
        if st.session_state.datasets:
            lines.append("LOADED DATASETS:")
            for ds_name, df in st.session_state.datasets.items():
                tgt = df.columns[-1]; feats = [c for c in df.columns if c != tgt]
                try:
                    rng = f"{pd.to_numeric(df[tgt], errors='coerce').min():.1f}–{pd.to_numeric(df[tgt], errors='coerce').max():.1f}"
                except Exception:
                    rng = "?"
                lines.append(f"  - {ds_name}: {len(df)} rows, features={feats}, target={tgt}, CAPEX range={rng}")
        if st.session_state.get("_last_metrics"):
            m = st.session_state["_last_metrics"]
            lines.append(f"LAST TRAINED MODEL: {m.get('model','?')}, R²={m.get('r2','?')}, baseline R²={m.get('baseline_r2','?')}")
        return "\n".join(lines) if lines else "No datasets loaded yet."

    SYSTEM_PROMPT = """You are a senior cost engineer and CAPEX expert with deep experience in offshore and onshore oil and gas projects (wellhead platforms, CPPs, FPSOs, pipelines, topsides). You cover parametric CAPEX estimation, cost driver analysis (water depth, weight, capacity, well count), what-if and sensitivity analysis, and Monte Carlo cost risk. Use the user's loaded data context below. Be specific with numbers, use MM USD unless told otherwise, and be honest about uncertainty.

CURRENT APP CONTEXT:
{context}"""

    st.markdown("##### Quick questions")
    chip_cols = st.columns(3)
    chips = [
        ("📊 Main cost driver?", "Looking at my loaded dataset, what is the main cost driver for CAPEX?"),
        ("💡 Why is CAPEX high?", "What are the top 5 reasons for CAPEX overruns in offshore oil and gas?"),
        ("🌊 Water depth impact?", "How does water depth affect CAPEX for offshore platforms?"),
        ("⚠️ What if costs rise 20%?", "If material and labour costs rise 20% due to inflation, how would that affect my estimates?"),
        ("🔍 Compare my projects", "Compare the projects I built. Which is most cost efficient and why?"),
        ("📈 CAPEX vs production?", "What is the typical relationship between production capacity and FPSO CAPEX?"),
    ]
    for idx, (label, prompt_text) in enumerate(chips):
        with chip_cols[idx % 3]:
            if st.button(label, key=f"chip_{idx}", use_container_width=True):
                st.session_state.ai_messages.append({"role": "user", "content": prompt_text})

    st.divider()
    for msg in st.session_state.ai_messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    def call_anthropic(messages, system):
        import json as _json, urllib.request as _req, urllib.error as _err
        api_key = st.secrets.get("anthropic_api_key", "")
        if not api_key:
            return "❌ Anthropic API key not found. Add `anthropic_api_key` to .streamlit/secrets.toml."
        payload = _json.dumps({"model": "claude-sonnet-4-20250514", "max_tokens": 1000,
                               "system": system, "messages": messages}).encode()
        req = _req.Request("https://api.anthropic.com/v1/messages", data=payload,
                           headers={"Content-Type": "application/json", "x-api-key": api_key,
                                    "anthropic-version": "2023-06-01"}, method="POST")
        try:
            with _req.urlopen(req, timeout=30) as resp:
                data = _json.loads(resp.read())
        except _err.HTTPError as e:
            return f"❌ Anthropic API error {e.code}: {e.read().decode('utf-8', errors='replace')}"
        except Exception as e:
            return f"❌ Request failed: {e}"
        return " ".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")

    def call_ollama(messages, system, model, base_url):
        import json as _json, urllib.request as _req, urllib.error as _err
        all_msgs = [{"role": "system", "content": system}] + messages
        payload = _json.dumps({"model": model, "messages": all_msgs, "stream": False}).encode()
        req = _req.Request(f"{base_url.rstrip('/')}/api/chat", data=payload,
                           headers={"Content-Type": "application/json"}, method="POST")
        try:
            with _req.urlopen(req, timeout=60) as resp:
                data = _json.loads(resp.read())
        except Exception as e:
            return f"❌ Cannot reach Ollama at {base_url}. Is it running? Error: {e}"
        return data.get("message", {}).get("content", "No response from Ollama.")

    user_input = st.chat_input("Ask about CAPEX, cost drivers, what-if scenarios...")
    if user_input:
        st.session_state.ai_messages.append({"role": "user", "content": user_input})
        with st.chat_message("user"):
            st.markdown(user_input)
        system_with_ctx = SYSTEM_PROMPT.format(context=build_context_summary())
        with st.chat_message("assistant"):
            with st.spinner("Thinking..."):
                if st.session_state.ai_backend == "Anthropic API (Claude)":
                    reply = call_anthropic(st.session_state.ai_messages, system_with_ctx)
                else:
                    om = st.session_state.get("ollama_model_select", "llama3")
                    ou = st.session_state.get("ollama_url", "http://localhost:11434")
                    reply = call_ollama(st.session_state.ai_messages, system_with_ctx, model=om, base_url=ou)
                st.markdown(reply)
                st.session_state.ai_messages.append({"role": "assistant", "content": reply})

    if st.session_state.ai_messages:
        if st.button("🗑️ Clear chat", key="clear_ai_chat"):
            st.session_state.ai_messages = []; st.rerun()
