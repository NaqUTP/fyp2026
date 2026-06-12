# ======================================================================================
# CAPEX AI RT2026 — FULL APP
# Models : Random Forest + Gradient Boosting + MLP (Deep Learning)
# FIX 1  : AI Advisor — added x-api-key + anthropic-version headers
# FIX 2  : Removed all st.stop() calls inside tab blocks so AI Advisor always renders
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
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

st.set_page_config(page_title="CAPEX AI RT2026", page_icon="💠", layout="wide", initial_sidebar_state="expanded")

PETRONAS = {"teal":"#00A19B","teal_dark":"#008C87","purple":"#6C4DD3","white":"#FFFFFF","black":"#0E1116","border":"rgba(0,0,0,0.10)"}
SHAREPOINT_LINKS = {"Shallow Water":"https://petronas.sharepoint.com/sites/your-site/shallow-water","Deep Water":"https://petronas.sharepoint.com/sites/your-site/deep-water","Onshore":"https://petronas.sharepoint.com/sites/your-site/onshore","Uncon":"https://petronas.sharepoint.com/sites/your-site/uncon","CCS":"https://petronas.sharepoint.com/sites/your-site/ccs"}

st.markdown(f"""<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html,body{{font-family:'Inter',sans-serif;}}
[data-testid="stAppViewContainer"]{{background:{PETRONAS["white"]};color:{PETRONAS["black"]};padding-top:.5rem;}}
#MainMenu,footer{{visibility:hidden;}}
[data-testid="stSidebar"]{{background:linear-gradient(180deg,{PETRONAS["teal"]} 0%,{PETRONAS["teal_dark"]} 100%) !important;color:#fff !important;border-top-right-radius:16px;border-bottom-right-radius:16px;box-shadow:0 6px 20px rgba(0,0,0,.15);}}
[data-testid="stSidebar"] *{{color:#fff !important;}}
[data-testid="collapsedControl"]{{position:fixed !important;top:50% !important;left:10px !important;transform:translateY(-50%) !important;z-index:9999 !important;}}
.petronas-hero{{border-radius:20px;padding:28px 32px;margin:6px 0 18px 0;color:#fff;background:linear-gradient(135deg,{PETRONAS["teal"]},{PETRONAS["purple"]},{PETRONAS["black"]});background-size:200% 200%;animation:heroGradient 8s ease-in-out infinite,fadeIn .8s ease-in-out,heroPulse 5s ease-in-out infinite;box-shadow:0 10px 24px rgba(0,0,0,.12);}}
@keyframes heroGradient{{0%{{background-position:0% 50%}}50%{{background-position:100% 50%}}100%{{background-position:0% 50%}}}}
@keyframes fadeIn{{from{{opacity:0;transform:translateY(10px)}}to{{opacity:1;transform:translateY(0)}}}}
@keyframes heroPulse{{0%{{box-shadow:0 0 16px rgba(0,161,155,.45)}}25%{{box-shadow:0 0 26px rgba(108,77,211,.55)}}50%{{box-shadow:0 0 36px rgba(0,161,155,.55)}}75%{{box-shadow:0 0 26px rgba(108,77,211,.55)}}100%{{box-shadow:0 0 16px rgba(0,161,155,.45)}}}}
.petronas-hero h1{{margin:0 0 5px;font-weight:800;letter-spacing:.3px;}}
.petronas-hero p{{margin:0;opacity:.9;font-weight:500;}}
.stButton>button,.stDownloadButton>button,.petronas-button{{border-radius:10px;padding:.6rem 1.1rem;font-weight:600;color:#fff !important;border:none;background:linear-gradient(to right,{PETRONAS["teal"]},{PETRONAS["purple"]});background-size:200% auto;transition:background-position .85s ease,transform .2s ease,box-shadow .25s ease;text-decoration:none;display:inline-block;}}
.stButton>button:hover,.stDownloadButton>button:hover,.petronas-button:hover{{background-position:right center;transform:translateY(-1px);box-shadow:0 6px 16px rgba(0,0,0,.18);}}
.stTabs [role="tablist"]{{display:flex;gap:8px;border-bottom:none;padding-bottom:6px;}}
.stTabs [role="tab"]{{background:#fff;color:{PETRONAS["black"]};border-radius:8px;padding:10px 18px;border:1px solid {PETRONAS["border"]};font-weight:600;transition:all .3s ease;position:relative;}}
.stTabs [role="tab"]:hover{{background:linear-gradient(to right,{PETRONAS["teal"]},{PETRONAS["purple"]});color:#fff;}}
.stTabs [role="tab"][aria-selected="true"]{{background:linear-gradient(to right,{PETRONAS["teal"]},{PETRONAS["purple"]});color:#fff;border-color:transparent;box-shadow:0 4px 16px rgba(0,0,0,.15);}}
.stTabs [role="tab"][aria-selected="true"]::after{{content:'';position:absolute;left:10%;bottom:-3px;width:80%;height:3px;background:linear-gradient(90deg,{PETRONAS["teal"]},{PETRONAS["purple"]},{PETRONAS["teal"]});background-size:200% 100%;border-radius:2px;animation:glowSlide 2.5s linear infinite;}}
@keyframes glowSlide{{0%{{background-position:0% 50%}}50%{{background-position:100% 50%}}100%{{background-position:0% 50%}}}}
</style>""", unsafe_allow_html=True)

st.markdown("""<div class="petronas-hero"><h1>CAPEX AI RT2026</h1><p>Data-driven CAPEX prediction · Random Forest &amp; Gradient Boosting</p></div>""", unsafe_allow_html=True)

# ── auth ──────────────────────────────────────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
APPROVED_EMAILS  = [str(e).strip().lower() for e in st.secrets.get("emails", [])]
correct_password = st.secrets.get("password", None)
if not st.session_state.authenticated:
    with st.form("login_form"):
        st.markdown("#### 🔐 Access Required")
        email    = (st.text_input("Email Address", key="login_email") or "").strip().lower()
        password = st.text_input("Access Password", type="password", key="login_pwd")
        if st.form_submit_button("Login"):
            if (email in APPROVED_EMAILS) and (password == correct_password):
                st.session_state.authenticated = True; st.success("✅ Access granted."); st.rerun()
            else:
                st.error("❌ Invalid credentials.")
    st.stop()

for key, default in [("datasets",{}),("predictions",{}),("processed_excel_files",set()),("_last_metrics",None),("projects",{}),("uploader_nonce",0),("widget_nonce",0)]:
    if key not in st.session_state:
        st.session_state[key] = default

# ── helpers ───────────────────────────────────────────────────────────────────
def toast(msg, icon="✅"):
    try:    st.toast(f"{icon} {msg}")
    except: st.success(msg)

def is_junk_col(colname):
    h = str(colname).strip().upper()
    return (not h) or h.startswith("UNNAMED") or h in {"INDEX","IDX"}

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
    if target_col and target_col in df.columns: return currency_from_header(str(target_col))
    for c in reversed(df.columns):
        if not is_junk_col(c): return currency_from_header(str(c))
    return ""

def cost_breakdown(base_pred, sst_pct, owners_pct, cont_pct, esc_pct):
    base_pred   = float(base_pred)
    owners_cost = round(base_pred * (owners_pct/100), 2)
    sst_cost    = round(base_pred * (sst_pct/100), 2)
    contingency = round((base_pred + owners_cost) * (cont_pct/100), 2)
    escalation  = round((base_pred + owners_cost) * (esc_pct/100), 2)
    grand_total = round(base_pred + owners_cost + sst_cost + contingency + escalation, 2)
    return owners_cost, sst_cost, contingency, escalation, grand_total

def project_components_df(proj):
    rows = []
    for c in proj.get("components", []):
        rows.append({"Component":c["component_type"],"Dataset":c["dataset"],"Base CAPEX":float(c["prediction"]),
                     "Owner's Cost":float(c["breakdown"]["owners_cost"]),"Contingency":float(c["breakdown"]["contingency_cost"]),
                     "Escalation":float(c["breakdown"]["escalation_cost"]),"SST":float(c["breakdown"]["sst_cost"]),"Grand Total":float(c["breakdown"]["grand_total"])})
    return pd.DataFrame(rows)

def project_totals(proj):
    dfc = project_components_df(proj)
    if dfc.empty: return {k:0.0 for k in ("capex_sum","owners","cont","esc","sst","grand_total")}
    return {"capex_sum":float(dfc["Base CAPEX"].sum()),"owners":float(dfc["Owner's Cost"].sum()),
            "cont":float(dfc["Contingency"].sum()),"esc":float(dfc["Escalation"].sum()),
            "sst":float(dfc["SST"].sum()),"grand_total":float(dfc["Grand Total"].sum())}

GITHUB_USER = "Hafizuddin-Abd-Rahman-Dev-Upstream"; REPO_NAME = "Cost-Predictor"; BRANCH = "main"; DATA_FOLDER = "pages/data_CAPEX"

@st.cache_data(ttl=600, show_spinner=False)
def fetch_json(url):
    r = requests.get(url, timeout=15); r.raise_for_status(); return r.json()

@st.cache_data(ttl=600, show_spinner=False)
def list_csvs_from_manifest(folder_path):
    url = f"https://raw.githubusercontent.com/{GITHUB_USER}/{REPO_NAME}/{BRANCH}/{folder_path}/files.json"
    try:
        data = fetch_json(url); return [str(x) for x in data] if isinstance(data, list) else []
    except Exception as e:
        st.error(f"Failed to load CSV manifest: {e}"); return []

NOK_TO_USD = 0.092
NOD_INVESTMENTS_URL = "https://factpages.sodir.no/public?/Factpages/external/tableview/field_investment_yearly&rs:Command=Render&rc:Toolbar=false&rc:Parameters=f&IpAddress=not_used&CultureCode=en&rs:Format=CSV&Top100=false"
NOD_FIELDS_URL = "https://factpages.sodir.no/public?/Factpages/external/tableview/field&rs:Command=Render&rc:Toolbar=false&rc:Parameters=f&IpAddress=not_used&CultureCode=en&rs:Format=CSV&Top100=false"

@st.cache_data(ttl=3600, show_spinner=False)
def load_nod_data():
    inv = pd.read_csv(NOD_INVESTMENTS_URL, sep=";", encoding="utf-8-sig"); inv.columns = [c.strip() for c in inv.columns]
    col_map = {}
    for c in inv.columns:
        cl = c.lower()
        if "field" in cl and "name" in cl: col_map[c] = "Field"
        elif "year" in cl: col_map[c] = "Year"
        elif "invest" in cl: col_map[c] = "Investment_MNOK"
    inv = inv.rename(columns=col_map)
    keep = [c for c in ["Field","Year","Investment_MNOK"] if c in inv.columns]
    inv = inv[keep].dropna(subset=["Investment_MNOK"])
    inv["Investment_MNOK"] = pd.to_numeric(inv["Investment_MNOK"], errors="coerce")
    inv = inv[inv["Investment_MNOK"] > 0].dropna(subset=["Investment_MNOK"])
    ft = inv.groupby("Field").agg(Total_Investment_MNOK=("Investment_MNOK","sum"),Peak_Year_Investment=("Investment_MNOK","max"),First_Year=("Year","min"),Last_Year=("Year","max"),Num_Active_Years=("Year","count")).reset_index()
    try:
        fld = pd.read_csv(NOD_FIELDS_URL, sep=";", encoding="utf-8-sig"); fld.columns = [c.strip() for c in fld.columns]
        cm2 = {}
        for c in fld.columns:
            cl = c.lower()
            if "field" in cl and "name" in cl: cm2[c] = "Field"
            elif "water" in cl and "depth" in cl: cm2[c] = "Water_Depth_m"
            elif "kind" in cl or "type" in cl: cm2[c] = "Field_Type"
            elif "status" in cl: cm2[c] = "Status"
        fld = fld.rename(columns=cm2)
        kf = [c for c in ["Field","Water_Depth_m","Field_Type","Status"] if c in fld.columns]
        merged = ft.merge(fld[kf].dropna(subset=["Field"]), on="Field", how="left")
    except Exception: merged = ft
    merged["Development_Years"] = merged["Last_Year"] - merged["First_Year"] + 1
    merged["CAPEX_MMUSD"] = (merged["Total_Investment_MNOK"] * NOK_TO_USD / 1000).round(2)
    if "Water_Depth_m" in merged.columns: merged["Water_Depth_m"] = pd.to_numeric(merged["Water_Depth_m"], errors="coerce")
    if "Field_Type" in merged.columns: merged["Field_Type_Code"] = merged["Field_Type"].astype("category").cat.codes
    fc = [c for c in ["Total_Investment_MNOK","Peak_Year_Investment","First_Year","Development_Years","Num_Active_Years","Water_Depth_m","Field_Type_Code"] if c in merged.columns]
    result = merged[fc + ["CAPEX_MMUSD"]].dropna(subset=["CAPEX_MMUSD"])
    return result[result["CAPEX_MMUSD"] > 0].reset_index(drop=True)

# =============================================================================
# DATA PREPROCESSOR
# =============================================================================
class DataPreprocessor:
    @staticmethod
    def clean_dataframe(df):
        df = df.copy(); bad = [c for c in df.columns if is_junk_col(c)]
        if bad: df = df.drop(columns=bad)
        return df

    @staticmethod
    def extract_features_target(df):
        if df is None or df.empty: raise ValueError("Empty dataset")
        target_col = df.columns[-1]; feature_cols = [c for c in df.columns if c != target_col]
        if not feature_cols: raise ValueError("No feature columns found")
        X = df[feature_cols].copy(); y = pd.to_numeric(df[target_col], errors="coerce")
        if y.isna().sum() / len(y) > 0.8: raise ValueError(f"Target column '{target_col}' has too many missing values")
        return X, y, target_col

    @staticmethod
    def validate_feature_columns(X):
        X = X.copy()
        for col in X.columns:
            if X[col].dtype == object: X[col] = pd.to_numeric(X[col], errors="coerce")
        return X

# =============================================================================
# MLP
# =============================================================================
class CapexMLP(nn.Module if TORCH_AVAILABLE else object):
    def __init__(self, n_features):
        if not TORCH_AVAILABLE: raise ImportError("PyTorch not installed.")
        super().__init__()
        self.net = nn.Sequential(nn.Linear(n_features,128),nn.BatchNorm1d(128),nn.ReLU(),nn.Dropout(0.3),nn.Linear(128,64),nn.BatchNorm1d(64),nn.ReLU(),nn.Dropout(0.2),nn.Linear(64,32),nn.ReLU(),nn.Linear(32,1))
    def forward(self, x): return self.net(x).squeeze(1)

class MLPWrapper:
    def __init__(self, n_features, epochs=200, lr=0.001, batch_size=32, patience=20, random_state=42):
        self.n_features=n_features; self.epochs=epochs; self.lr=lr; self.batch_size=batch_size
        self.patience=patience; self.random_state=random_state; self.model=None
        self.scaler_X=StandardScaler(); self.scaler_y=StandardScaler()
        self.train_losses=[]; self.val_losses=[]; self.imputer=SimpleImputer(strategy="median")

    def fit(self, X, y):
        torch.manual_seed(self.random_state); np.random.seed(self.random_state)
        Xi = self.imputer.fit_transform(X); Xs = self.scaler_X.fit_transform(Xi).astype(np.float32)
        ys = self.scaler_y.fit_transform(y.reshape(-1,1)).ravel().astype(np.float32)
        nv = max(1, int(len(Xs)*0.15)); Xtr,Xv,ytr,yv = Xs[:-nv],Xs[-nv:],ys[:-nv],ys[-nv:]
        loader = DataLoader(TensorDataset(torch.from_numpy(Xtr),torch.from_numpy(ytr)), batch_size=self.batch_size, shuffle=True)
        self.model = CapexMLP(self.n_features)
        opt = torch.optim.Adam(self.model.parameters(), lr=self.lr)
        sch = torch.optim.lr_scheduler.ReduceLROnPlateau(opt, patience=10, factor=0.5)
        crit = nn.MSELoss(); best_val=float("inf"); best_state=None; no_imp=0
        self.train_losses=[]; self.val_losses=[]
        for _ in range(self.epochs):
            self.model.train(); ep=0.0
            for Xb,yb in loader:
                opt.zero_grad(); loss=crit(self.model(Xb),yb); loss.backward(); opt.step(); ep+=loss.item()*len(Xb)
            ep/=len(Xtr); self.model.eval()
            with torch.no_grad(): vl=crit(self.model(torch.from_numpy(Xv)),torch.from_numpy(yv)).item()
            self.train_losses.append(ep); self.val_losses.append(vl); sch.step(vl)
            if vl < best_val: best_val=vl; best_state={k:v.clone() for k,v in self.model.state_dict().items()}; no_imp=0
            else:
                no_imp+=1
                if no_imp>=self.patience: break
        if best_state: self.model.load_state_dict(best_state)
        self.model.eval(); return self

    def predict(self, X):
        Xi = self.imputer.transform(X); Xs = self.scaler_X.transform(Xi).astype(np.float32)
        with torch.no_grad(): ys = self.model(torch.from_numpy(Xs)).numpy()
        return self.scaler_y.inverse_transform(ys.reshape(-1,1)).ravel()

# =============================================================================
# MODEL PIPELINE
# =============================================================================
class ModelPipeline:
    MODEL_CANDIDATES = {
        "RandomForest":    lambda rs=42: RandomForestRegressor(n_estimators=200,max_depth=None,min_samples_split=2,min_samples_leaf=1,random_state=rs,n_jobs=-1),
        "GradientBoosting":lambda rs=42: GradientBoostingRegressor(n_estimators=200,learning_rate=0.05,max_depth=4,subsample=0.8,random_state=rs),
    }

    @classmethod
    def create_pipeline(cls, model_name, random_state=42):
        if model_name not in cls.MODEL_CANDIDATES: model_name = "RandomForest"
        ctor = cls.MODEL_CANDIDATES[model_name]
        try: model = ctor(random_state)
        except: model = ctor()
        return Pipeline([("imputer",SimpleImputer(strategy="median")),("model",model)])

    @classmethod
    @st.cache_resource(show_spinner=False)
    def train_all_cached(_cls, X, y, test_size=0.20, random_state=42, mlp_epochs=200, mlp_lr=0.001, mlp_batch=32, mlp_patience=20):
        Xa = X.values.astype(np.float32); ya = y.values.astype(np.float32)
        Xtr,Xte,ytr,yte = train_test_split(Xa, ya, test_size=test_size, random_state=random_state)
        results = {}
        for name in ("RandomForest","GradientBoosting"):
            pipe = _cls.create_pipeline(name, random_state); pipe.fit(Xtr,ytr); yp = pipe.predict(Xte)
            results[name] = {"pipeline":pipe,"r2":round(float(r2_score(yte,yp)),4),"rmse":round(float(np.sqrt(mean_squared_error(yte,yp))),4),"mae":round(float(mean_absolute_error(yte,yp)),4),"y_test":yte,"y_pred":yp,"type":"sklearn"}
        if TORCH_AVAILABLE:
            mlp = MLPWrapper(Xtr.shape[1],mlp_epochs,mlp_lr,mlp_batch,mlp_patience,random_state); mlp.fit(Xtr,ytr); ypm = mlp.predict(Xte)
            results["MLP"] = {"pipeline":mlp,"r2":round(float(r2_score(yte,ypm)),4),"rmse":round(float(np.sqrt(mean_squared_error(yte,ypm))),4),"mae":round(float(mean_absolute_error(yte,ypm)),4),"y_test":yte,"y_pred":ypm,"train_losses":mlp.train_losses,"val_losses":mlp.val_losses,"type":"mlp"}
        else:
            results["MLP"] = {"pipeline":None,"r2":None,"rmse":None,"mae":None,"y_test":yte,"y_pred":np.zeros_like(yte),"train_losses":[],"val_losses":[],"type":"mlp","error":"PyTorch not installed."}
        valid = {k:v for k,v in results.items() if v["r2"] is not None}
        best = max(valid, key=lambda k: valid[k]["r2"]); bm = valid[best]
        return {"rf":results["RandomForest"],"gb":results["GradientBoosting"],"mlp":results["MLP"],"best":best,"pipeline":bm["pipeline"],"feature_cols":list(X.columns),"model":best,"r2":bm["r2"],"rmse":bm["rmse"],"mae":bm["mae"]}

    @classmethod
    @st.cache_resource(show_spinner=False)
    def train_both_cached(_cls, X, y, test_size=0.20, random_state=42):
        return _cls.train_all_cached(X, y, test_size, random_state)

    @staticmethod
    def prepare_prediction_input(feature_cols, payload):
        row = {}
        for col in feature_cols:
            val = payload.get(col, np.nan)
            if val is None or (isinstance(val, str) and val.strip() == ""): row[col] = np.nan
            elif isinstance(val, (int,float,np.number)): row[col] = float(val)
            else:
                try: row[col] = float(val)
                except: row[col] = np.nan
        return pd.DataFrame([row], columns=feature_cols)

def monte_carlo_simulation(pipeline, feature_cols, base_values, n_simulations=1000, feature_uncertainty=0.05):
    np.random.seed(42); base_array = np.array([float(base_values.get(c, np.nan)) for c in feature_cols]); preds = []
    for _ in range(n_simulations):
        noise = np.random.normal(0, feature_uncertainty, len(base_array)); sim_feats = base_array * (1 + noise)
        sim_df = pd.DataFrame([sim_feats], columns=feature_cols)
        try: preds.append(float(pipeline.predict(sim_df)[0]))
        except: preds.append(0.0)
    return pd.DataFrame({"prediction": preds})

# ── nav bar ───────────────────────────────────────────────────────────────────
for col, label in zip(st.columns(5), ["SHALLOW WATER","DEEP WATER","ONSHORE","UNCON","CCS"]):
    with col:
        url = SHAREPOINT_LINKS.get(label.title(), "#")
        st.markdown(f'<a href="{url}" target="_blank" rel="noopener" class="petronas-button" style="width:100%;text-align:center;display:inline-block;">{label}</a>', unsafe_allow_html=True)

tab_data, tab_pb, tab_mc, tab_compare, tab_ai = st.tabs(["📊 Data","🏗️ Project Builder","🎲 Monte Carlo","🔀 Compare Projects","🤖 AI Advisor"])

# =============================================================================
# TAB 1 — DATA
# =============================================================================
with tab_data:
    st.markdown('<h3 style="margin-top:0;color:#000;">📁 Data</h3>', unsafe_allow_html=True)
    c1, c2 = st.columns([1.2,1])
    with c1:
        data_source = st.radio("Choose data source", ["Upload CSV","Load from Server","Norwegian Offshore Directorate (Live)"], horizontal=True, key="data_source")
    with c2:
        st.caption("Enterprise Storage (SharePoint)")
        st.markdown('<a href="https://petronas.sharepoint.com/sites/ecm_ups_coe/confidential/DFE%20Cost%20Engineering/Forms/AllItems.aspx" target="_blank" rel="noopener" class="petronas-button">Open Enterprise Storage</a>', unsafe_allow_html=True)

    uploaded_files = []
    if data_source == "Upload CSV":
        uploaded_files = st.file_uploader("Upload CSV files (max 200 MB)", type="csv", accept_multiple_files=True, key=f"csv_uploader_{st.session_state.uploader_nonce}")
    elif data_source == "Load from Server":
        github_csvs = list_csvs_from_manifest(DATA_FOLDER)
        if github_csvs:
            sel = st.selectbox("Choose CSV from GitHub", github_csvs, key="github_csv_select")
            if st.button("Load selected CSV", key="load_github_csv_btn"):
                raw_url = f"https://raw.githubusercontent.com/{GITHUB_USER}/{REPO_NAME}/{BRANCH}/{DATA_FOLDER}/{sel}"
                try:
                    df = DataPreprocessor.clean_dataframe(pd.read_csv(raw_url))
                    st.session_state.datasets[sel] = df; st.session_state.predictions.setdefault(sel, [])
                    toast(f"Loaded: {sel}"); st.rerun()
                except Exception as e: st.error(f"Error loading CSV: {e}")
        else: st.info("No CSV files found in GitHub folder.")
    elif data_source == "Norwegian Offshore Directorate (Live)":
        st.markdown("Pulls **real** North Sea field investment data from [factpages.sodir.no](https://factpages.sodir.no).")
        nod_col1, nod_col2 = st.columns([2,1])
        with nod_col1: st.caption(f"Source: NOD · Licence: NLOD · 1 NOK ≈ {NOK_TO_USD} USD")
        with nod_col2: nod_rate = st.number_input("NOK → USD rate", 0.01, 1.0, NOK_TO_USD, 0.001, format="%.3f", key="nod_rate")
        if st.button("Load NOD Field Data", key="load_nod_btn", type="primary"):
            try:
                with st.spinner("Fetching from Norwegian Offshore Directorate..."):
                    df_nod = load_nod_data()
                    if abs(nod_rate - NOK_TO_USD) > 0.0001: df_nod["CAPEX_MMUSD"] = (df_nod["Total_Investment_MNOK"] * nod_rate / 1000).round(2)
                ds_key = "NOD_FieldInvestments"; st.session_state.datasets[ds_key] = df_nod; st.session_state.predictions.setdefault(ds_key, [])
                mn, mx = df_nod["CAPEX_MMUSD"].min(), df_nod["CAPEX_MMUSD"].max()
                st.success(f"✅ **{len(df_nod)} real fields** · {df_nod.shape[1]} features · CAPEX range: ${mn:,.0f}M – ${mx:,.0f}M")
                with st.expander("Preview NOD data", expanded=True): st.dataframe(df_nod.head(10), use_container_width=True)
                st.rerun()
            except Exception as e: st.error(f"Failed to load NOD data: {e}")

    if uploaded_files:
        for up in uploaded_files:
            if up.name not in st.session_state.datasets:
                try:
                    df = DataPreprocessor.clean_dataframe(pd.read_csv(up))
                    st.session_state.datasets[up.name] = df; st.session_state.predictions.setdefault(up.name, [])
                except Exception as e: st.error(f"Failed to read {up.name}: {e}")
        toast("Dataset(s) added.")

    st.divider()
    cA,cB,cC,cD = st.columns(4)
    with cA:
        if st.button("🧹 Clear predictions", key="clear_preds_btn"):
            st.session_state.predictions = {k:[] for k in st.session_state.predictions}; toast("Predictions cleared.","🧹"); st.rerun()
    with cB:
        if st.button("🧺 Clear history", key="clear_processed_btn"):
            st.session_state.processed_excel_files = set(); toast("History cleared.","🧺"); st.rerun()
    with cC:
        if st.button("🔁 Refresh", key="refresh_manifest_btn"):
            list_csvs_from_manifest.clear(); fetch_json.clear(); toast("Refreshed.","🔁"); st.rerun()
    with cD:
        if st.button("🗂️ Clear all data", key="clear_datasets_btn"):
            st.session_state.datasets={}; st.session_state.predictions={}; st.session_state.processed_excel_files=set()
            st.session_state._last_metrics=None; st.session_state.uploader_nonce+=1; st.session_state.widget_nonce+=1
            toast("All data cleared.","🗂️"); st.rerun()
    st.divider()

    if not st.session_state.datasets:
        st.info("Upload or load a dataset to proceed.")
    else:
        ds_name_data = st.selectbox("Active dataset", list(st.session_state.datasets.keys()), key="active_dataset_data")
        df_active = st.session_state.datasets[ds_name_data]; target_col_active = df_active.columns[-1]; currency_active = get_currency_symbol(df_active, target_col_active)
        colA,colB,colC,colD2 = st.columns(4)
        colA.metric("Rows", f"{df_active.shape[0]:,}"); colB.metric("Columns", f"{df_active.shape[1]:,}"); colC.metric("Currency", currency_active or "—"); colD2.caption(f"Target column: **{target_col_active}**")
        with st.expander("Preview (first 10 rows)", expanded=False): st.dataframe(df_active.head(10), use_container_width=True)

        st.divider()
        st.markdown('<h3 style="margin-top:0;color:#000;">⚙️ Model Training</h3>', unsafe_allow_html=True)
        ds_name_model = st.selectbox("Dataset for training", list(st.session_state.datasets.keys()), key="ds_model")
        df_model = st.session_state.datasets[ds_name_model]; data_ok = False
        try:
            X, y, target_col = DataPreprocessor.extract_features_target(df_model)
            st.success(f"✅ Ready — **{X.shape[1]} features**, target: **{target_col}**")
            col1,col2,col3 = st.columns(3); col1.metric("Features", X.shape[1]); col2.metric("Samples", X.shape[0])
            valid_n = int(y.notna().sum()); col3.metric("Valid targets", f"{valid_n} ({valid_n/len(y)*100:.1f}%)"); data_ok = True
        except Exception as e: st.error(f"Data preparation failed: {e}")

        if data_ok:
            st.markdown("##### Train / Test Split")
            split_col, btn_col = st.columns([3,1])
            with split_col:
                test_size = st.slider("Test set size", 0.10, 0.40, 0.20, 0.05, key="train_test_size")
                train_pct = round((1-test_size)*100); test_pct = round(test_size*100); n_total = len(X); n_train = int(n_total*(1-test_size)); n_test = n_total-n_train
                st.markdown(f"""<div style="display:flex;height:14px;border-radius:7px;overflow:hidden;margin-top:4px;"><div style="width:{train_pct}%;background:#00A19B;"></div><div style="width:{test_pct}%;background:#6C4DD3;"></div></div><div style="display:flex;justify-content:space-between;font-size:12px;margin-top:4px;"><span style="color:#00A19B;font-weight:600;">🟦 Train {train_pct}% &nbsp;({n_train} rows)</span><span style="color:#6C4DD3;font-weight:600;">🟪 Test {test_pct}% &nbsp;({n_test} rows)</span></div>""", unsafe_allow_html=True)
            with btn_col:
                st.write(""); st.write("")
                run_train = st.button("🚀 Train RF, GB & MLP", key="run_training_btn", type="primary")

            with st.expander("🧠 MLP Hyperparameters", expanded=False):
                if not TORCH_AVAILABLE: st.warning("⚠️ PyTorch not installed — MLP will be skipped. Add `torch` to requirements.txt.")
                mc1,mc2,mc3,mc4 = st.columns(4)
                with mc1: mlp_epochs = st.number_input("Max Epochs", 50, 500, 200, 50, key="mlp_epochs")
                with mc2: mlp_lr = st.select_slider("Learning Rate", [0.0001,0.0005,0.001,0.005,0.01], value=0.001, key="mlp_lr")
                with mc3: mlp_batch = st.selectbox("Batch Size", [16,32,64,128], index=1, key="mlp_batch")
                with mc4: mlp_patience = st.number_input("Early Stop Patience", 5, 50, 20, 5, key="mlp_patience")

            if run_train:
                try:
                    with st.spinner("Training Random Forest, Gradient Boosting, and MLP…"):
                        metrics = ModelPipeline.train_all_cached(X, y, float(test_size), 42, int(mlp_epochs), float(mlp_lr), int(mlp_batch), int(mlp_patience))
                    st.session_state._last_metrics = metrics
                    st.session_state[f"trained_model__{ds_name_model}"] = metrics
                    st.session_state[f"current_pipeline__{ds_name_model}"] = metrics["pipeline"]
                    st.session_state[f"feature_cols__{ds_name_model}"] = metrics["feature_cols"]
                    knn_imputer = KNNImputer(n_neighbors=5); knn_imputer.fit(X); st.session_state[f"knn_imputer_{ds_name_model}"] = knn_imputer
                    toast("Training complete! 🎉")
                    if not TORCH_AVAILABLE: st.warning("⚠️ MLP skipped — add `torch` to requirements.txt.")

                    st.markdown("##### Model Comparison — RF vs GB vs MLP")
                    rf, gb, mlp = metrics["rf"], metrics["gb"], metrics["mlp"]
                    mlp_r2 = mlp["r2"] if mlp["r2"] is not None else float("nan")
                    mlp_rmse = mlp["rmse"] if mlp["rmse"] is not None else float("nan")
                    mlp_mae = mlp["mae"] if mlp["mae"] is not None else float("nan")
                    compare_df = pd.DataFrame({"Metric":["R² Score ↑","RMSE ↓","MAE ↓"],"Random Forest":[rf["r2"],rf["rmse"],rf["mae"]],"Gradient Boosting":[gb["r2"],gb["rmse"],gb["mae"]],"MLP (Deep Learning)":[mlp_r2,mlp_rmse,mlp_mae]})
                    def highlight_best_3(row):
                        vals = {"Random Forest":row["Random Forest"],"Gradient Boosting":row["Gradient Boosting"],"MLP (Deep Learning)":row["MLP (Deep Learning)"]}
                        vv = {k:v for k,v in vals.items() if not (isinstance(v,float) and np.isnan(v))}
                        if not vv: return [""]*len(row)
                        bc = max(vv, key=vv.get) if row["Metric"].endswith("↑") else min(vv, key=vv.get)
                        return ["background-color:#d4f5f3;font-weight:700" if c==bc else "" for c in row.index]
                    styled = compare_df.style.apply(highlight_best_3, axis=1).format({"Random Forest":"{:.4f}","Gradient Boosting":"{:.4f}","MLP (Deep Learning)":lambda v: f"{v:.4f}" if not np.isnan(v) else "N/A"})
                    st.dataframe(styled, use_container_width=True, hide_index=True)

                    winner = metrics["best"]; winner_label = {"RandomForest":"Random Forest","GradientBoosting":"Gradient Boosting","MLP":"MLP (Deep Learning)"}.get(winner, winner)
                    st.success(f"✅ **{winner_label}** selected as active model (highest R²)")
                    m1,m2,m3,m4 = st.columns(4); m1.metric("Model",winner_label); m2.metric("R²",f"{metrics['r2']:.4f}"); m3.metric("RMSE",f"{metrics['rmse']:,.2f}"); m4.metric("MAE",f"{metrics['mae']:,.2f}")

                    st.markdown("##### Actual vs Predicted — All Models")
                    fig_scatter = go.Figure(); all_vals = []
                    for key, label, colour in [("rf","Random Forest","#00A19B"),("gb","Gradient Boosting","#6C4DD3"),("mlp","MLP","#F4801A")]:
                        m = metrics[key]
                        if m["r2"] is None: continue
                        fig_scatter.add_trace(go.Scatter(x=m["y_test"],y=m["y_pred"],mode="markers",marker=dict(color=colour,opacity=0.55,size=6),name=label))
                        all_vals.extend(m["y_test"].tolist()); all_vals.extend(m["y_pred"].tolist())
                    lo,hi = float(min(all_vals)),float(max(all_vals))
                    fig_scatter.add_trace(go.Scatter(x=[lo,hi],y=[lo,hi],mode="lines",line=dict(color="#888",dash="dash",width=1.5),name="Perfect fit"))
                    fig_scatter.update_layout(xaxis_title="Actual CAPEX (MM USD)",yaxis_title="Predicted CAPEX (MM USD)",height=400,margin=dict(l=0,r=0,t=10,b=0),paper_bgcolor="white",plot_bgcolor="white",legend=dict(orientation="h",y=-0.18))
                    st.plotly_chart(fig_scatter, use_container_width=True)

                    if TORCH_AVAILABLE and metrics["mlp"]["train_losses"]:
                        st.markdown("##### MLP Training & Validation Loss Curve")
                        train_l,val_l = metrics["mlp"]["train_losses"],metrics["mlp"]["val_losses"]; epochs_ran = list(range(1,len(train_l)+1))
                        fig_loss = go.Figure()
                        fig_loss.add_trace(go.Scatter(x=epochs_ran,y=train_l,mode="lines",name="Train Loss",line=dict(color="#00A19B",width=2)))
                        fig_loss.add_trace(go.Scatter(x=epochs_ran,y=val_l,mode="lines",name="Val Loss",line=dict(color="#F4801A",width=2,dash="dot")))
                        fig_loss.update_layout(xaxis_title="Epoch",yaxis_title="MSE Loss (scaled)",height=320,margin=dict(l=0,r=0,t=10,b=0),paper_bgcolor="white",plot_bgcolor="white",legend=dict(orientation="h",y=-0.22))
                        st.plotly_chart(fig_loss, use_container_width=True)

                    st.markdown("##### Feature Importance — RF vs GB")
                    fi_left,fi_right = st.columns(2)
                    for container,(label,bkey) in zip([fi_left,fi_right],[("Random Forest","rf"),("Gradient Boosting","gb")]):
                        pipe = metrics[bkey]["pipeline"]
                        fi_df = pd.DataFrame({"Feature":metrics["feature_cols"],"Importance":pipe.named_steps["model"].feature_importances_}).sort_values("Importance",ascending=True)
                        fig_fi = go.Figure(go.Bar(x=fi_df["Importance"],y=fi_df["Feature"],orientation="h",marker_color="#00A19B" if bkey=="rf" else "#6C4DD3"))
                        fig_fi.update_layout(title=label,xaxis_title="Importance",height=max(260,32*len(fi_df)),margin=dict(l=0,r=0,t=35,b=0),paper_bgcolor="white",plot_bgcolor="white")
                        with container: st.plotly_chart(fig_fi, use_container_width=True)
                except Exception as e: st.error(f"Training failed: {e}")

        st.divider()
        st.markdown('<h3 style="margin-top:0;color:#000;">📈 Visualization</h3>', unsafe_allow_html=True)
        ds_name_viz = st.selectbox("Dataset for visualization", list(st.session_state.datasets.keys()), key="ds_viz")
        df_viz = st.session_state.datasets[ds_name_viz]; target_col_viz = df_viz.columns[-1]; currency_viz = get_currency_symbol(df_viz,target_col_viz) or "USD"; numeric_df_viz = df_viz.select_dtypes(include=[np.number])
        viz_tab1,viz_tab2,viz_tab3,viz_tab4 = st.tabs(["🔥 Correlation Matrix","🌐 3D Cost Surface","📊 Distribution","🔗 Scatter Matrix"])
        with viz_tab1:
            try:
                if len(numeric_df_viz.columns) > 1:
                    corr = numeric_df_viz.corr(); fig_cor = px.imshow(corr,text_auto=".2f",aspect="auto",color_continuous_scale="RdBu_r",zmin=-1,zmax=1)
                    fig_cor.update_layout(margin=dict(l=0,r=0,t=10,b=0),paper_bgcolor="white"); st.plotly_chart(fig_cor, use_container_width=True)
                else: st.info("Need at least 2 numeric columns.")
            except Exception as e: st.error(f"Correlation error: {e}")
        with viz_tab2:
            st.caption("Drag to rotate · Scroll to zoom")
            cols_lower = {c.lower():c for c in df_viz.columns}; depth_col=cols_lower.get("water_depth_m"); topsides_col=cols_lower.get("topsides_weight_t"); length_col=cols_lower.get("length_km"); diam_col=cols_lower.get("diameter_inch")
            if depth_col and topsides_col: x_col,y_col = depth_col,topsides_col
            elif length_col and diam_col: x_col,y_col = length_col,diam_col
            else:
                all_feat = [c for c in numeric_df_viz.columns if c!=target_col_viz]; x_col=all_feat[0] if len(all_feat)>=1 else None; y_col=all_feat[1] if len(all_feat)>=2 else None
            if x_col and y_col:
                all_num_cols = [c for c in numeric_df_viz.columns if c!=target_col_viz]
                c3d_1,c3d_2 = st.columns(2)
                with c3d_1: x_col = st.selectbox("X axis", all_num_cols, index=all_num_cols.index(x_col), key="viz3d_x")
                with c3d_2: y_col = st.selectbox("Y axis", all_num_cols, index=all_num_cols.index(y_col) if y_col in all_num_cols else 0, key="viz3d_y")
                plot_3d = df_viz[[x_col,y_col,target_col_viz]].dropna()
                if len(plot_3d) >= 3:
                    fig_3d = go.Figure(data=[go.Scatter3d(x=plot_3d[x_col],y=plot_3d[y_col],z=plot_3d[target_col_viz],mode="markers",marker=dict(size=5,color=plot_3d[target_col_viz],colorscale="Teal",colorbar=dict(title=dict(text=f"CAPEX ({currency_viz}M)",side="right"),thickness=14,len=0.6),opacity=0.85,line=dict(width=0.4,color="rgba(0,0,0,0.25)")),hovertemplate=f"<b>{x_col}:</b> %{{x:,.1f}}<br><b>{y_col}:</b> %{{y:,.1f}}<br><b>CAPEX:</b> {currency_viz} %{{z:,.2f}}M<extra></extra>")])
                    fig_3d.update_layout(scene=dict(xaxis=dict(title=x_col),yaxis=dict(title=y_col),zaxis=dict(title=f"CAPEX ({currency_viz}M)"),camera=dict(eye=dict(x=1.6,y=1.6,z=0.8))),margin=dict(l=0,r=0,t=10,b=0),height=540,paper_bgcolor="white")
                    st.plotly_chart(fig_3d, use_container_width=True)
                    s1,s2,s3,s4 = st.columns(4); s1.metric("Min CAPEX",f"{currency_viz} {plot_3d[target_col_viz].min():,.1f}M"); s2.metric("Max CAPEX",f"{currency_viz} {plot_3d[target_col_viz].max():,.1f}M"); s3.metric("Mean CAPEX",f"{currency_viz} {plot_3d[target_col_viz].mean():,.1f}M"); s4.metric("Data points",f"{len(plot_3d):,}")
                else: st.warning("Not enough data after removing nulls.")
            else: st.info("Need at least 2 numeric feature columns.")
        with viz_tab3:
            dist_col = st.selectbox("Select column", numeric_df_viz.columns.tolist(), index=len(numeric_df_viz.columns)-1, key="viz_dist_col")
            d1,d2 = st.columns(2)
            with d1:
                fig_hist = px.histogram(df_viz,x=dist_col,nbins=30,title=f"Distribution — {dist_col}",color_discrete_sequence=["#00A19B"],template="plotly_white")
                fig_hist.update_layout(height=320,margin=dict(l=0,r=0,t=35,b=0),showlegend=False); st.plotly_chart(fig_hist, use_container_width=True)
            with d2:
                fig_box = px.box(df_viz,y=dist_col,title=f"Box Plot — {dist_col}",color_discrete_sequence=["#6C4DD3"],template="plotly_white")
                fig_box.update_layout(height=320,margin=dict(l=0,r=0,t=35,b=0),showlegend=False); st.plotly_chart(fig_box, use_container_width=True)
            st.markdown("##### Descriptive Statistics"); st.dataframe(df_viz[dist_col].describe().to_frame().T.style.format("{:,.2f}"), use_container_width=True)
        with viz_tab4:
            num_cols_list = numeric_df_viz.columns.tolist()
            if len(num_cols_list) >= 3:
                default_sel = num_cols_list[:min(6,len(num_cols_list))]
                selected_scatter = st.multiselect("Choose columns (max 6)", num_cols_list, default=default_sel, key="viz_scatter_cols")
                if len(selected_scatter) >= 2:
                    fig_sm = px.scatter_matrix(df_viz[selected_scatter].dropna(),dimensions=selected_scatter,color=df_viz[target_col_viz],color_continuous_scale="Teal",title="Scatter Matrix",template="plotly_white",opacity=0.6)
                    fig_sm.update_traces(marker=dict(size=3)); fig_sm.update_layout(height=600,margin=dict(l=0,r=0,t=35,b=0)); st.plotly_chart(fig_sm, use_container_width=True)
                else: st.info("Select at least 2 columns.")
            else: st.info("Need at least 3 numeric columns.")

        st.divider()
        st.markdown('<h3 style="margin-top:0;color:#000;">🎯 Predict</h3>', unsafe_allow_html=True)
        ds_name_pred = st.selectbox("Dataset for prediction", list(st.session_state.datasets.keys()), key="ds_pred")
        df_pred = st.session_state.datasets[ds_name_pred]
        if f"current_pipeline__{ds_name_pred}" not in st.session_state:
            st.warning("⚠️ Please train a model first using the Model Training section above.")
        else:
            pipeline = st.session_state[f"current_pipeline__{ds_name_pred}"]; feature_cols = st.session_state[f"feature_cols__{ds_name_pred}"]
            target_col = df_pred.columns[-1]; currency_pred = get_currency_symbol(df_pred, target_col)
            trained_meta = st.session_state.get(f"trained_model__{ds_name_pred}", {}); active_model = trained_meta.get("best","—")
            active_label = {"RandomForest":"Random Forest","GradientBoosting":"Gradient Boosting","MLP":"MLP (Deep Learning)"}.get(active_model, active_model)
            st.info(f"🤖 Active model: **{active_label}**")
            cf1,cf2 = st.columns(2)
            with cf1: sst_pct=st.number_input("SST (%)",0.0,100.0,0.0,0.5,key="pred_sst"); owners_pct=st.number_input("Owner's Cost (%)",0.0,100.0,0.0,0.5,key="pred_owner")
            with cf2: cont_pct=st.number_input("Contingency (%)",0.0,100.0,0.0,0.5,key="pred_cont"); esc_pct=st.number_input("Escalation (%)",0.0,100.0,0.0,0.5,key="pred_esc")
            project_name = st.text_input("Project Name", placeholder="e.g., Offshore Pipeline Replacement 2026", key="pred_project_name")
            st.markdown("##### Feature Values"); st.caption(f"Enter values for **{len(feature_cols)}** features. Leave blank = NaN (will be imputed).")
            input_values = {}
            for i in range(0, len(feature_cols), 3):
                cols = st.columns(3)
                for j, feat in enumerate(feature_cols[i:i+3]):
                    with cols[j]:
                        val = st.text_input(feat, value="", key=f"input_{feat}_{ds_name_pred}")
                        if val.strip() in ("","nan"): input_values[feat] = np.nan
                        else:
                            try: input_values[feat] = float(val)
                            except: input_values[feat] = np.nan
            use_knn = st.checkbox("🔮 Use KNN imputation for missing features", value=False, key=f"use_knn_{ds_name_pred}")
            if st.button("Run Prediction", key="run_pred_btn", type="primary"):
                try:
                    pred_input=ModelPipeline.prepare_prediction_input(feature_cols,input_values); original_inputs=input_values.copy(); imputation_method="pipeline median"
                    if use_knn:
                        knn_key = f"knn_imputer_{ds_name_pred}"
                        if knn_key in st.session_state:
                            arr=st.session_state[knn_key].transform(pred_input); pred_input=pd.DataFrame(arr,columns=feature_cols); imputation_method="KNN"
                        else: st.warning("KNN imputer unavailable — using median fallback.")
                    base_pred = float(pipeline.predict(pred_input)[0])
                    st.markdown("##### Feature Values Used")
                    comp_rows = []
                    for col in feature_cols:
                        u=original_inputs.get(col); v=pred_input[col].iloc[0]
                        comp_rows.append({"Feature":col,"Your Input":f"{u:,.2f}" if (u is not None and not pd.isna(u)) else "—","Value Used":f"{v:,.2f}" if isinstance(v,(int,float)) else str(v),"Source":"User provided" if (u is not None and not pd.isna(u)) else f"Imputed ({imputation_method})"})
                    st.dataframe(pd.DataFrame(comp_rows), use_container_width=True, height=min(400,35*len(feature_cols)))
                    owners_cost,sst_cost,contingency,escalation,grand_total = cost_breakdown(base_pred,sst_pct,owners_pct,cont_pct,esc_pct)
                    result = {"Project Name":project_name,"Model Used":active_label,"Base CAPEX":round(base_pred,2),"Owner's Cost":owners_cost,"SST Cost":sst_cost,"Contingency":contingency,"Escalation":escalation,"Grand Total":grand_total,"Target":round(base_pred,2)}
                    for col in feature_cols: result[col] = pred_input[col].iloc[0]
                    st.session_state.predictions.setdefault(ds_name_pred,[]).append(result); toast("Prediction added!")
                    r1,r2,r3,r4,r5 = st.columns(5)
                    r1.metric("Base CAPEX",f"{currency_pred} {base_pred:,.2f}"); r2.metric("Owner's Cost",f"{currency_pred} {owners_cost:,.2f}"); r3.metric("SST",f"{currency_pred} {sst_cost:,.2f}"); r4.metric("Contingency",f"{currency_pred} {contingency:,.2f}"); r5.metric("Grand Total",f"{currency_pred} {grand_total:,.2f}")
                except Exception as e: st.error(f"Prediction failed: {e}")

            st.markdown("---"); st.markdown("##### Batch Prediction (Excel)")
            excel_file = st.file_uploader("Upload Excel for batch prediction", type=["xlsx"], key=f"batch_excel_{st.session_state.widget_nonce}")
            if excel_file:
                file_id = f"{excel_file.name}_{excel_file.size}_{ds_name_pred}"
                if file_id not in st.session_state.processed_excel_files:
                    try:
                        batch_df=pd.read_excel(excel_file); missing_cols=[c for c in feature_cols if c not in batch_df.columns]
                        if missing_cols: st.error(f"Missing columns in Excel: {missing_cols}")
                        else:
                            X_batch=DataPreprocessor.validate_feature_columns(batch_df[feature_cols]); preds_batch=pipeline.predict(X_batch)
                            for i,(_, row) in enumerate(batch_df.iterrows()):
                                bp=float(preds_batch[i]); oc,sc,cc,ec,gt=cost_breakdown(bp,sst_pct,owners_pct,cont_pct,esc_pct)
                                r={"Project Name":str(row.get("Project Name",f"Project {i+1}")),"Model Used":active_model,"Base CAPEX":round(bp,2),"Owner's Cost":oc,"SST Cost":sc,"Contingency":cc,"Escalation":ec,"Grand Total":gt,"Target":round(bp,2)}
                                for feat in feature_cols: r[feat]=row.get(feat,np.nan)
                                st.session_state.predictions.setdefault(ds_name_pred,[]).append(r)
                            st.session_state.processed_excel_files.add(file_id); st.success(f"✅ Processed {len(batch_df)} rows")
                    except Exception as e: st.error(f"Batch processing failed: {e}")
                else: st.info("This file has already been processed.")

        st.divider(); st.markdown('<h3 style="margin-top:0;color:#000;">📄 Results</h3>', unsafe_allow_html=True)
        ds_name_res = st.selectbox("Dataset for results", list(st.session_state.datasets.keys()), key="ds_results")
        preds = st.session_state.predictions.get(ds_name_res, [])
        if preds:
            df_preds=pd.DataFrame(preds); display_cols=[c for c in ["Project Name","Model Used","Base CAPEX","Owner's Cost","SST Cost","Contingency","Escalation","Grand Total"] if c in df_preds.columns]
            df_display=df_preds[display_cols].copy()
            for col in display_cols[2:]: df_display[col]=df_display[col].apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "")
            st.dataframe(df_display, use_container_width=True, height=300)
            col1,col2 = st.columns(2)
            with col1:
                bio=io.BytesIO(); df_preds.to_excel(bio,index=False,engine="openpyxl"); bio.seek(0)
                st.download_button("⬇️ Download Excel", data=bio, file_name=f"{ds_name_res}_predictions.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_excel_btn")
            with col2:
                st.download_button("⬇️ Download CSV", data=df_preds.to_csv(index=False), file_name=f"{ds_name_res}_predictions.csv", mime="text/csv", key="download_csv_btn")
            if st.button("🗑️ Clear predictions", key="clear_predictions_btn"):
                st.session_state.predictions[ds_name_res]=[]; st.rerun()
        else: st.info("No predictions yet.")


# =============================================================================
# TAB 2 — PROJECT BUILDER  (no st.stop — uses if/else)
# =============================================================================
with tab_pb:
    st.markdown('<h4 style="margin-top:0;color:#000;">🏗️ Project Builder</h4>', unsafe_allow_html=True)
    st.caption("Assemble multi-component CAPEX projects from trained models.")
    if not st.session_state.datasets:
        st.info("No datasets loaded. Please load data in the Data tab first.")
    else:
        colA,colB = st.columns([2,1])
        with colA: new_proj = st.text_input("New Project Name", placeholder="e.g., CAPEX 2026", key="pb_new_project_name")
        with colB:
            if new_proj and new_proj not in st.session_state.projects:
                if st.button("Create Project", key="pb_create_project_btn"):
                    st.session_state.projects[new_proj]={"components":[],"currency":"","cost_factors":{"sst_pct":0.0,"owners_pct":0.0,"cont_pct":0.0,"esc_pct":0.0}}
                    toast(f"Project '{new_proj}' created."); st.rerun()
        if not st.session_state.projects:
            st.info("Create a project above, then add components.")
        else:
            proj_sel = st.selectbox("Select project", list(st.session_state.projects.keys()), key="pb_project_select")
            proj = st.session_state.projects[proj_sel]
            st.markdown("##### Project Cost Factors")
            cf1,cf2 = st.columns(2)
            with cf1:
                proj["cost_factors"]["sst_pct"]=st.number_input("SST (%)",0.0,100.0,proj["cost_factors"].get("sst_pct",0.0),0.5,key=f"pb_sst_{proj_sel}")
                proj["cost_factors"]["owners_pct"]=st.number_input("Owner's Cost (%)",0.0,100.0,proj["cost_factors"].get("owners_pct",0.0),0.5,key=f"pb_owners_{proj_sel}")
            with cf2:
                proj["cost_factors"]["cont_pct"]=st.number_input("Contingency (%)",0.0,100.0,proj["cost_factors"].get("cont_pct",0.0),0.5,key=f"pb_cont_{proj_sel}")
                proj["cost_factors"]["esc_pct"]=st.number_input("Escalation (%)",0.0,100.0,proj["cost_factors"].get("esc_pct",0.0),0.5,key=f"pb_esc_{proj_sel}")
            st.markdown("##### Add Component")
            dataset_for_comp = st.selectbox("Dataset for component", sorted(st.session_state.datasets.keys()), key="pb_dataset_for_component")
            df_comp = st.session_state.datasets[dataset_for_comp]; curr_comp = get_currency_symbol(df_comp, df_comp.columns[-1])
            if f"current_pipeline__{dataset_for_comp}" not in st.session_state:
                st.warning(f"Please train a model for '{dataset_for_comp}' in the Data tab first.")
            else:
                pipeline_comp=st.session_state[f"current_pipeline__{dataset_for_comp}"]; feat_comp=st.session_state[f"feature_cols__{dataset_for_comp}"]
                meta_comp=st.session_state.get(f"trained_model__{dataset_for_comp}",{}); best_comp=meta_comp.get("best","—")
                label_comp={"RandomForest":"Random Forest","GradientBoosting":"Gradient Boosting","MLP":"MLP (Deep Learning)"}.get(best_comp,best_comp)
                st.info(f"🤖 Model: **{label_comp}** | R²: **{meta_comp.get('r2',0):.4f}**")
                component_type = st.text_input("Component type", placeholder="e.g., Pipeline, Platform, FPSO", key=f"pb_component_type_{proj_sel}")
                comp_inputs = {}
                for i in range(0, len(feat_comp), 2):
                    cols = st.columns(2)
                    for j,feat in enumerate(feat_comp[i:i+2]):
                        with cols[j]:
                            val = st.text_input(feat, placeholder="Enter value", key=f"pb_{feat}_{proj_sel}_{dataset_for_comp}")
                            if val.strip() in ("","nan"): comp_inputs[feat]=np.nan
                            else:
                                try: comp_inputs[feat]=float(val)
                                except: comp_inputs[feat]=np.nan
                if st.button("➕ Add Component", key=f"pb_add_comp_{proj_sel}"):
                    if not component_type: st.error("Please enter a component type.")
                    else:
                        try:
                            pi=ModelPipeline.prepare_prediction_input(feat_comp,comp_inputs); bp=float(pipeline_comp.predict(pi)[0])
                            cf=proj["cost_factors"]; oc,sc,cc,ec,gt=cost_breakdown(bp,cf["sst_pct"],cf["owners_pct"],cf["cont_pct"],cf["esc_pct"])
                            proj["components"].append({"component_type":component_type,"dataset":dataset_for_comp,"model_used":meta_comp.get("best","—"),"prediction":bp,"breakdown":{"owners_cost":oc,"sst_cost":sc,"contingency_cost":cc,"escalation_cost":ec,"grand_total":gt}})
                            proj["currency"]=curr_comp; toast(f"Component '{component_type}' added."); st.rerun()
                        except Exception as e: st.error(f"Failed to add component: {e}")
            st.markdown("---"); comps = proj.get("components",[])
            if comps:
                st.markdown("##### Components")
                df_comps = pd.DataFrame([{"Component":c["component_type"],"Dataset":c["dataset"],"Model":c.get("model_used","—"),"Base CAPEX":f"{curr_comp} {c['prediction']:,.2f}","Grand Total":f"{curr_comp} {c['breakdown']['grand_total']:,.2f}"} for c in comps])
                st.dataframe(df_comps, use_container_width=True)
                totals=project_totals(proj); t1,t2,t3=st.columns(3)
                t1.metric("Total Base CAPEX",f"{curr_comp} {totals['capex_sum']:,.2f}"); t2.metric("Total SST",f"{curr_comp} {totals['sst']:,.2f}"); t3.metric("Total Grand Total",f"{curr_comp} {totals['grand_total']:,.2f}")
                st.markdown("##### Manage Components")
                for idx,comp in enumerate(comps):
                    c1,c2,c3=st.columns([3,2,1]); c1.write(f"**{comp['component_type']}** — {comp.get('model_used','—')}"); c1.caption(f"{comp['dataset']} | Base: {curr_comp} {comp['prediction']:,.2f}"); c2.write(f"Grand Total: {curr_comp} {comp['breakdown']['grand_total']:,.2f}")
                    with c3:
                        if st.button("🗑️", key=f"del_comp_{proj_sel}_{idx}"): comps.pop(idx); st.rerun()
                st.markdown("---")
                proj_json = json.dumps(proj, indent=2, default=float)
                st.download_button("⬇️ Download Project (JSON)", data=proj_json, file_name=f"{proj_sel}.json", mime="application/json", key=f"dl_json_{proj_sel}")
            else: st.info("No components yet.")
            st.markdown("##### Import Project")
            up_json = st.file_uploader("Upload project JSON", type=["json"], key=f"import_{proj_sel}")
            if up_json:
                try: st.session_state.projects[proj_sel]=json.load(up_json); toast("Project imported."); st.rerun()
                except Exception as e: st.error(f"Import failed: {e}")


# =============================================================================
# TAB 3 — MONTE CARLO  (no st.stop — uses if/else)
# =============================================================================
with tab_mc:
    st.markdown('<h3 style="margin-top:0;color:#000;">🎲 Monte Carlo Analysis</h3>', unsafe_allow_html=True)
    if not st.session_state.projects:
        st.info("Create a project in the Project Builder tab first.")
    else:
        proj_sel_mc=st.selectbox("Select project", list(st.session_state.projects.keys()), key="mc_project_select")
        proj_mc=st.session_state.projects[proj_sel_mc]; comps_mc=proj_mc.get("components",[])
        if not comps_mc:
            st.warning("This project has no components. Add components in the Project Builder tab.")
        else:
            mc1,mc2 = st.columns(2)
            with mc1: n_sims=st.number_input("Simulations",100,10000,1000,100,key="mc_n_sims"); feat_unc=st.slider("Feature uncertainty (%)",0.0,50.0,10.0,1.0,key="mc_feat_unc")
            with mc2: budget=st.number_input("Budget threshold (MM USD)",0.0,value=1000.0,step=10.0,key="mc_budget")
            if st.button("Run Monte Carlo", type="primary", key="mc_run"):
                try:
                    with st.spinner("Running simulations…"):
                        all_sims=[]
                        for comp in comps_mc:
                            ds=comp["dataset"]
                            if f"current_pipeline__{ds}" not in st.session_state: st.warning(f"No trained model for {ds}"); continue
                            pipe=st.session_state[f"current_pipeline__{ds}"]; fcols=st.session_state[f"feature_cols__{ds}"]
                            sims=monte_carlo_simulation(pipe,fcols,{},int(n_sims),feat_unc/100); all_sims.append(sims["prediction"].values)
                    if all_sims:
                        total_sims=np.sum(all_sims,axis=0); p50=np.percentile(total_sims,50); p80=np.percentile(total_sims,80); p90=np.percentile(total_sims,90); exceed_pct=(total_sims>budget).mean()*100
                        rc1,rc2,rc3,rc4=st.columns(4); rc1.metric("P50",f"${p50:,.0f}M"); rc2.metric("P80",f"${p80:,.0f}M"); rc3.metric("P90",f"${p90:,.0f}M"); rc4.metric(f"P(>${budget:,.0f}M)",f"{exceed_pct:.1f}%")
                        fig_mc=px.histogram(x=total_sims,nbins=50,title="Total Cost Distribution",labels={"x":"Total Cost (MM USD)","y":"Frequency"},color_discrete_sequence=["#00A19B"])
                        fig_mc.add_vline(x=budget,line_dash="dash",line_color="red",annotation_text=f"Budget: ${budget:,.0f}M"); st.plotly_chart(fig_mc, use_container_width=True)
                    else: st.warning("No valid simulations generated.")
                except Exception as e: st.error(f"Monte Carlo failed: {e}")


# =============================================================================
# TAB 4 — COMPARE PROJECTS  (no st.stop — uses if/else)
# =============================================================================
with tab_compare:
    st.markdown('<h3 style="margin-top:0;color:#000;">🔀 Compare Projects</h3>', unsafe_allow_html=True)
    if len(st.session_state.projects) < 2:
        st.info("Create at least 2 projects in the Project Builder to compare.")
    else:
        proj_names=list(st.session_state.projects.keys()); sel_projs=st.multiselect("Select projects to compare",proj_names,default=proj_names[:2])
        if len(sel_projs) < 2: st.warning("Please select at least 2 projects.")
        else:
            cmp_data=[]
            for pn in sel_projs:
                p=st.session_state.projects[pn]; t=project_totals(p)
                cmp_data.append({"Project":pn,"Components":len(p.get("components",[])),"Base CAPEX":t["capex_sum"],"SST":t["sst"],"Owner's Cost":t["owners"],"Contingency":t["cont"],"Escalation":t["esc"],"Grand Total":t["grand_total"]})
            df_cmp=pd.DataFrame(cmp_data); st.markdown("##### Comparison Table")
            st.dataframe(df_cmp.style.format({c:"{:,.2f}" for c in df_cmp.columns if c not in ("Project","Components")}), use_container_width=True)
            viz_type=st.selectbox("Chart type",["Bar Chart","Stacked Bar"],key="viz_type")
            if viz_type=="Bar Chart":
                fig_cmp=px.bar(df_cmp,x="Project",y="Grand Total",title="Grand Total by Project",text="Grand Total",color_discrete_sequence=["#00A19B"]); fig_cmp.update_traces(texttemplate="%{text:,.0f}",textposition="outside")
            else:
                melt=df_cmp.melt(id_vars=["Project"],value_vars=["Base CAPEX","SST","Owner's Cost","Contingency","Escalation"],var_name="Cost Type",value_name="Amount")
                fig_cmp=px.bar(melt,x="Project",y="Amount",color="Cost Type",title="Cost Breakdown by Project",barmode="stack")
            st.plotly_chart(fig_cmp, use_container_width=True)


# =============================================================================
# TAB 5 — AI ADVISOR  (always renders — no st.stop() anywhere above in tab blocks)
# =============================================================================
with tab_ai:
    st.markdown('<h3 style="margin-top:0;color:#000;">🤖 AI CAPEX Advisor</h3>', unsafe_allow_html=True)
    st.caption("Ask anything about CAPEX, cost drivers, what-if scenarios, or project risks.")

    if "ai_messages" not in st.session_state: st.session_state.ai_messages = []
    if "ai_backend" not in st.session_state: st.session_state.ai_backend = "Anthropic API (Claude)"

    with st.expander("⚙️ AI Backend Settings", expanded=False):
        backend = st.radio("Choose AI backend", ["Anthropic API (Claude)","Ollama (local, open-source)"], horizontal=True, key="ai_backend_select")
        st.session_state.ai_backend = backend
        if backend == "Anthropic API (Claude)":
            st.caption("Reads `anthropic_api_key` from Streamlit secrets. Add to `.streamlit/secrets.toml`: `anthropic_api_key = \"sk-ant-...\"`")
        else:
            st.caption("Requires Ollama running locally. Install: https://ollama.com · Run: `ollama serve` then `ollama pull llama3`")
            ollama_model = st.selectbox("Ollama model", ["llama3","llama3:70b","mistral","deepseek-r1:7b","qwen3:8b","gemma3:9b"], key="ollama_model_select")
            ollama_url = st.text_input("Ollama base URL", value="http://localhost:11434", key="ollama_url")

    def build_context_summary():
        lines = []
        if st.session_state.datasets:
            lines.append("LOADED DATASETS:")
            for ds_name, df in st.session_state.datasets.items():
                tgt=df.columns[-1]; feats=[c for c in df.columns if c!=tgt]
                lines.append(f"  - {ds_name}: {len(df)} rows, features={feats}, target={tgt}, CAPEX range={df[tgt].min():.1f}–{df[tgt].max():.1f} MM USD")
        if st.session_state.predictions:
            for ds_name, preds in st.session_state.predictions.items():
                if preds:
                    lines.append(f"PREDICTIONS for {ds_name} ({len(preds)} records):")
                    for p in preds[-3:]: lines.append(f"  - {p.get('Project Name','?')}: Base={p.get('Base CAPEX','?')} MM USD, Grand Total={p.get('Grand Total','?')} MM USD")
        if st.session_state.projects:
            lines.append("PROJECTS:")
            for pn, proj in st.session_state.projects.items():
                t=project_totals(proj); lines.append(f"  - {pn}: {len(proj.get('components',[]))} components, Grand Total={t['grand_total']:.1f} MM USD")
        if st.session_state.get("_last_metrics"):
            m=st.session_state["_last_metrics"]; lines.append(f"LAST TRAINED MODEL: {m.get('model','?')}, R²={m.get('r2','?')}, RMSE={m.get('rmse','?')}")
        return "\n".join(lines) if lines else "No datasets or predictions loaded yet."

    SYSTEM_PROMPT = """You are a senior cost engineer and CAPEX expert with 25 years of experience in offshore and onshore oil & gas projects. You specialise in:
- Wellhead platforms (WHP), Central Processing Platforms (CPP), FPSOs, pipelines, topsides
- Parametric and deterministic CAPEX estimation
- Cost driver analysis (water depth, topsides weight, production capacity, number of wells)
- What-if scenario analysis and sensitivity analysis
- Industry cost indices (UCCI, CERA)
- Monte Carlo cost risk analysis
- PETRONAS, North Sea, West Africa, and Southeast Asia project contexts
- Comparing estimated vs actual spend and explaining variances
- Risk factors: schedule delays, material price inflation, contractor availability

You have access to the user's loaded datasets, predictions, and project data (shown below). Use this context when answering questions about their specific data.
When asked "why is the CAPEX like this?" — explain the key cost drivers from the data.
When asked "what would happen if..." — give a quantified scenario answer.
When asked about cost overruns — cite common causes (scope creep, weather, fabrication delays).
Always be specific with numbers where possible. Use MM USD as currency unless told otherwise.
Keep answers concise but technically rigorous. If you are uncertain, say so clearly.

CURRENT APP CONTEXT:
{context}"""

    st.markdown("##### Quick questions")
    chip_cols = st.columns(3)
    chips = [
        ("📊 Main cost driver?",      "Looking at my loaded dataset, what is the main cost driver for CAPEX? Which feature has the strongest influence?"),
        ("💡 Why is CAPEX high?",     "Why might the CAPEX be higher than expected for an offshore platform? What are the top 5 reasons for CAPEX overruns in oil and gas?"),
        ("🌊 Water depth impact?",    "How does water depth affect CAPEX for offshore platforms? Give me a parametric breakdown of cost vs water depth."),
        ("⚠️ What if costs rise 20%?","If material and labour costs increase by 20% due to inflation, how would that affect the CAPEX estimates in my projects?"),
        ("🔍 Compare my projects",    "Compare the projects I have built in the Project Builder. Which has the best cost efficiency and why?"),
        ("📈 CAPEX vs production?",   "What is the typical relationship between production capacity (kbopd) and CAPEX for an FPSO? Give me a parametric rule of thumb."),
    ]
    for idx,(label,prompt_text) in enumerate(chips):
        with chip_cols[idx%3]:
            if st.button(label, key=f"chip_{idx}", use_container_width=True):
                st.session_state.ai_messages.append({"role":"user","content":prompt_text})

    st.divider()
    for msg in st.session_state.ai_messages:
        with st.chat_message(msg["role"]): st.markdown(msg["content"])

    def call_anthropic(messages, system):
        import json as _json, urllib.request as _req, urllib.error as _err
        api_key = st.secrets.get("anthropic_api_key", "")
        if not api_key: return "❌ **Anthropic API key not found.**\n\nAdd to `.streamlit/secrets.toml`:\n\n```toml\nanthropic_api_key = \"sk-ant-...\"\n```\n\nThen restart the app."
        payload = _json.dumps({"model":"claude-sonnet-4-20250514","max_tokens":1000,"system":system,"messages":messages}).encode()
        req = _req.Request("https://api.anthropic.com/v1/messages", data=payload,
                           headers={"Content-Type":"application/json","x-api-key":api_key,"anthropic-version":"2023-06-01"}, method="POST")
        try:
            with _req.urlopen(req, timeout=30) as resp: data = _json.loads(resp.read())
        except _err.HTTPError as e: return f"❌ Anthropic API error {e.code}: {e.read().decode('utf-8',errors='replace')}"
        except Exception as e: return f"❌ Request failed: {e}"
        return " ".join(b.get("text","") for b in data.get("content",[]) if b.get("type")=="text")

    def call_ollama(messages, system, model, base_url):
        import json as _json, urllib.request as _req, urllib.error as _err
        all_msgs = [{"role":"system","content":system}] + messages
        payload = _json.dumps({"model":model,"messages":all_msgs,"stream":False}).encode()
        req = _req.Request(f"{base_url.rstrip('/')}/api/chat", data=payload, headers={"Content-Type":"application/json"}, method="POST")
        try:
            with _req.urlopen(req, timeout=60) as resp: data = _json.loads(resp.read())
        except _err.URLError as e: return f"❌ Cannot reach Ollama at `{base_url}`. Make sure Ollama is running (`ollama serve`). Error: {e.reason}"
        except Exception as e: return f"❌ Ollama request failed: {e}"
        return data.get("message",{}).get("content","No response from Ollama.")

    user_input = st.chat_input("Ask me anything about CAPEX, cost drivers, what-if scenarios...")
    if user_input:
        st.session_state.ai_messages.append({"role":"user","content":user_input})
        with st.chat_message("user"): st.markdown(user_input)
        system_with_ctx = SYSTEM_PROMPT.format(context=build_context_summary())
        with st.chat_message("assistant"):
            with st.spinner("Thinking..."):
                backend_sel = st.session_state.ai_backend
                if backend_sel == "Anthropic API (Claude)":
                    reply = call_anthropic(st.session_state.ai_messages, system_with_ctx)
                else:
                    om=st.session_state.get("ollama_model_select","llama3"); ou=st.session_state.get("ollama_url","http://localhost:11434")
                    reply = call_ollama(st.session_state.ai_messages, system_with_ctx, model=om, base_url=ou)
                st.markdown(reply); st.session_state.ai_messages.append({"role":"assistant","content":reply})

    if st.session_state.ai_messages:
        ctrl1,ctrl2 = st.columns([1,4])
        with ctrl1:
            if st.button("🗑️ Clear chat", key="clear_ai_chat"): st.session_state.ai_messages=[]; st.rerun()
        with ctrl2:
            chat_export = "\n\n".join(f"{m['role'].upper()}:\n{m['content']}" for m in st.session_state.ai_messages)
            st.download_button("⬇️ Export chat", data=chat_export, file_name="capex_ai_chat.txt", mime="text/plain", key="export_ai_chat")
