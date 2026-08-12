import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path
from datetime import datetime

from optimizer import (
    TARGETS, FE_LOWER, FE_UPPER, get_default_chemistry,
    solve_blend_with_compensation, calculate_cost_breakdown,
    quality_checks, quality_table, redistribute_adjustment,
    what_if_analysis, compute_achieved
)

st.set_page_config(page_title="Sinter Burden Control", page_icon="🏭", layout="wide")

# ---------- INDUSTRIAL THEME ----------
st.markdown(r'''
<style>
:root{--bg:#0b1015;--p:#121a21;--p2:#17232b;--l:#2a3a44;--t:#edf3f6;--m:#91a1ab;--s:#4f8fb8;--g:#35c47a;--a:#e6a63a;--r:#d95757;--o:#e47732}
html,body,[class*="css"]{font-family:Inter,system-ui,sans-serif}
.stApp{background:radial-gradient(circle at 75% 0%,#24475b22,transparent 28%),var(--bg);color:var(--t)}
.block-container{max-width:1800px;padding:.75rem 1rem 2rem}
[data-testid="stSidebar"]{background:#0e161c;border-right:1px solid #263740;width:220px!important;min-width:220px!important}
[data-testid="stSidebar"] > div:first-child{width:220px!important}
section[data-testid="stSidebar"]{width:220px!important}
h1,h2,h3{color:var(--t)!important}.sub,.small{color:var(--m);font-size:.7rem}.eyebrow{color:#7194a7;font-size:.58rem;letter-spacing:.16em;font-weight:900}
.panel{background:linear-gradient(145deg,#141e25,#10171d);border:1px solid var(--l);border-radius:11px;padding:.8rem}
.panel-title{font-size:.6rem;font-weight:900;letter-spacing:.12em;color:#9eb2bd;text-transform:uppercase;margin-bottom:.6rem}
.hero{background:linear-gradient(110deg,#17242d,#111a21 65%,#18242b);border:1px solid #314651;border-radius:12px;padding:.8rem}
.kpi{border:1px solid var(--l);border-radius:10px;padding:.7rem;min-height:88px;background:#121b22}
.kpi-label{font-size:.55rem;letter-spacing:.12em;font-weight:900;color:#82949f}.kpi-value{font-size:1.15rem;font-weight:900;margin-top:.2rem}.kpi-sub{font-size:.58rem;color:#778a95}
.kpi-s{border-left:3px solid var(--s)}.kpi-g{border-left:3px solid var(--g)}.kpi-a{border-left:3px solid var(--a)}.kpi-r{border-left:3px solid var(--r)}.kpi-o{border-left:3px solid var(--o)}
.badge{display:inline-block;border-radius:999px;padding:.17rem .45rem;font-size:.53rem;font-weight:900}.ok{background:#123524;color:#64d695;border:1px solid #276a47}.out{background:#351616;color:#ff7d7d;border:1px solid #6e3030}.warn{background:#34270f;color:#f1bd5e;border:1px solid #705322}.info{background:#122b39;color:#79b9dc;border:1px solid #2c5870}
.notice{padding:.55rem .65rem;border-radius:8px;border:1px solid #304550;background:#15222a;font-size:.64rem}.notice-w{border-color:#6a5025;background:#261e11}
.table-wrap{overflow-x:auto;overflow-y:visible;border:1px solid #263740;border-radius:8px}.pretty{width:100%;border-collapse:collapse;font-size:.72rem}.pretty th{background:#16232b;color:#b7c9d3;padding:.55rem;text-align:left;font-size:.68rem;letter-spacing:.02em}.pretty td{padding:.5rem;border-top:1px solid #22313a;white-space:nowrap}.pretty tr.total{background:#1a2932;font-weight:900}.pretty tr.iron td:first-child{border-left:3px solid #4f8fb8}.pretty tr.flux td:first-child{border-left:3px solid #35c47a}.pretty tr.recycle td:first-child{border-left:3px solid #e6a63a}.pretty tr.fuel td:first-child{border-left:3px solid #d95757}
.qitem{padding:.45rem 0;border-bottom:1px solid #23333c}.qtop{display:flex;justify-content:space-between;font-size:.63rem;font-weight:800}.qsub{font-size:.56rem;color:#8497a2;margin-top:.15rem}.meter{height:4px;background:#22313a;border-radius:99px;margin-top:.25rem}.meter div{height:100%;background:#35c47a;border-radius:99px}.meter.bad div{background:#d95757}
.navhead{font-size:.54rem;letter-spacing:.14em;font-weight:900;color:#617985;margin:.7rem .2rem .25rem}.side{font-size:.58rem;color:#9aacb6;border:1px solid #263740;border-radius:8px;padding:.55rem;background:#10191f}
.stButton>button{border-radius:7px!important;background:#142029!important;border:1px solid #2b414c!important;color:#e5eef2!important;font-weight:800!important;font-size:.64rem!important}.stButton>button[kind="primary"]{background:linear-gradient(135deg,#2d6f96,#245876)!important}
[data-testid="stDataEditor"]{border:1px solid #293b45;border-radius:8px;overflow:hidden}.stSlider label{font-size:.70rem!important}
.footer{border-top:1px solid #22323b;margin-top:1.2rem;padding-top:.5rem;color:#627681;font-size:.54rem;text-align:right}
</style>
''',unsafe_allow_html=True)

GROUPS=["Iron_ore","Flux","Recycle","Fuel"]
LABEL={"Iron_ore":"Iron Ore","Flux":"Flux","Recycle":"Recycle","Fuel":"Fuel"}
COLORS={"Iron_ore":"#4f8fb8","Flux":"#35c47a","Recycle":"#e6a63a","Fuel":"#d95757"}
MASTER=["Material","Group","Fe","SiO2","Al2O3","CaO","MgO","LOI","Tech_Min","Tech_Max"]

# ---------- EXCEL ----------
def clean_cols(df):
    df=df.copy(); df.columns=df.columns.astype(str).str.strip().str.replace(r"\s+","_",regex=True); return df

def load_primary(f):
    df = clean_cols(pd.read_excel(f))

    required = MASTER + ["Price_Rs_t", "Available_Tonnes"]
    miss = [c for c in required if c not in df.columns]
    if miss:
        raise ValueError(
            "Master Excel missing mandatory column(s): " + ", ".join(miss) +
            ". Required: Material, Group, Fe, SiO2, Al2O3, CaO, MgO, LOI, "
            "Tech_Min, Tech_Max, Price_Rs_t, Available_Tonnes."
        )

    type_col = next((c for c in ["Material_Type", "Type", "MaterialType"] if c in df.columns), None)
    if type_col is None:
        df["Material_Type"] = "Primary"
    else:
        df["Material_Type"] = (
            df[type_col].astype(str).str.strip().str.lower()
            .map({"primary": "Primary", "alternative": "Alternative"})
        )
        if df["Material_Type"].isna().any():
            raise ValueError("Material_Type must contain only Primary or Alternative.")

    df = df[required + ["Material_Type"]].copy()
    df["Material"] = df["Material"].astype(str).str.strip()
    df["Group"] = df["Group"].astype(str).str.strip()

    if df["Material"].duplicated().any():
        raise ValueError("Duplicate material names in Excel.")
    if not set(df["Group"]).issubset(GROUPS):
        raise ValueError("Invalid Group. Use Iron_ore, Flux, Recycle or Fuel.")

    df = df.set_index("Material")
    for c in MASTER[2:]:
        df[c] = pd.to_numeric(df[c], errors="raise")

    for c in ["Price_Rs_t", "Available_Tonnes", "Tech_Max"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
        if df[c].isna().any():
            raise ValueError(f"{c} contains blank/non-numeric values.")
        if (df[c] < 0).any():
            raise ValueError(f"{c} cannot contain negative values.")

    return df


def split_material_types(df):
    primary = df.index[df["Material_Type"].eq("Primary")].tolist()
    alternatives = df.index[df["Material_Type"].eq("Alternative")].tolist()
    return primary, alternatives

def load_alt(f):
    df = load_primary(f).copy()
    df["Material_Type"] = "Alternative"
    df["Group"] = "Iron_ore"
    return df


# ---------- STATE ----------
if "df" not in st.session_state:
    st.session_state.df=get_default_chemistry().copy()
    st.session_state.source="Built-in Master Chemistry"
    st.session_state.primary=list(st.session_state.df.index)
    st.session_state.alts=[]; st.session_state.alt_on={}
    st.session_state.avail={m:True for m in st.session_state.df.index}
    st.session_state.result=None; st.session_state.manual_base=None
    st.session_state.manual=None; st.session_state.changed=False
    st.session_state.whatif=None; st.session_state.runs=0
if "nav" not in st.session_state: st.session_state.nav="Dashboard"

# Self-heal sessions created by older dashboard versions.
if "df" in st.session_state:
    if "Material_Type" not in st.session_state.df.columns:
        st.session_state.df["Material_Type"] = "Primary"
    p, a = split_material_types(st.session_state.df)
    st.session_state.primary = p
    st.session_state.alts = a
    if "alt_on" not in st.session_state:
        st.session_state.alt_on = {m: False for m in a}
    else:
        for m in a:
            st.session_state.alt_on.setdefault(m, False)
    if "avail" not in st.session_state:
        st.session_state.avail = {m: (False if m in a else True) for m in st.session_state.df.index}


def reset_primary(df,source):
    df = df.copy()
    if "Material_Type" not in df.columns:
        df["Material_Type"] = "Primary"
    primary, alternatives = split_material_types(df)

    st.session_state.df = df
    st.session_state.source = source
    st.session_state.primary = primary
    st.session_state.alts = alternatives
    st.session_state.alt_on = {m: False for m in alternatives}
    st.session_state.avail = {m: (False if m in alternatives else True) for m in df.index}
    st.session_state.result = None
    st.session_state.manual_base = None
    st.session_state.manual = None
    st.session_state.changed = False
    st.session_state.whatif = None
    st.session_state.runs = 0

def add_alt(df):
    dup=set(df.index)&set(st.session_state.df.index)
    if dup: raise ValueError("Alternative material already exists: "+", ".join(sorted(dup)))
    st.session_state.df=pd.concat([st.session_state.df,df])
    for m in df.index:
        st.session_state.alts.append(m); st.session_state.alt_on[m]=False; st.session_state.avail[m]=False
    st.session_state.changed=True

def active_df():
    df=st.session_state.df.copy()
    for m in df.index:
        if m in st.session_state.alts:
            if not st.session_state.alt_on.get(m,False): df.loc[m,"Available_Tonnes"]=0
        elif not st.session_state.avail.get(m,True): df.loc[m,"Available_Tonnes"]=0
    return df

# ---------- DISPLAY ----------
def chip(g): return f'<span class="badge info">{LABEL.get(g,g)}</span>'

def table(df,money=set(),status=None):
    rows=[]
    for _,r in df.iterrows():
        m=str(r.get("Material","")); g=str(r.get("Group","")); cl="total" if m=="TOTAL" else g.replace("_"," ").lower().replace(" ","")
        cells=[]
        for c in df.columns:
            v=r[c]
            if pd.isna(v): v=""
            if c=="Group" and v!="": v=chip(str(v))
            elif c in money and v!="":
                try:v=f"₹{float(v):,.2f}"
                except:pass
            elif isinstance(v,(float,int)) and not isinstance(v,bool): v=f"{float(v):,.2f}"
            if status==c:
                s=str(v); cls="ok" if any(x in s.lower() for x in ["ok","available","optimal","pass","feasible"]) else "out" if any(x in s.lower() for x in ["out","critical","unavailable","review"]) else "warn"
                v=f'<span class="badge {cls}">● {s}</span>'
            cells.append(f"<td>{v}</td>")
        rows.append(f'<tr class="{cl}">{"".join(cells)}</tr>')
    return '<div class="table-wrap"><table class="pretty"><thead><tr>'+''.join(f"<th>{c}</th>" for c in df.columns)+'</tr></thead><tbody>'+''.join(rows)+'</tbody></table></div>'

def material_sequence(df=None):
    """Single display/analysis sequence for all material tables.

    Primary materials always retain their original master-Excel order.
    Alternative materials are appended in their uploaded/master order.
    No table is allowed to reorder these materials by quantity or cost.
    """
    seq = []

    for m in st.session_state.primary:
        if df is None or m in df.index:
            seq.append(m)

    for m in st.session_state.alts:
        if m not in seq and (df is None or m in df.index):
            seq.append(m)

    return seq


def breakdown(blend, df):
    b, c, total = calculate_cost_breakdown(blend, df)

    order = material_sequence(df)
    order_pos = {m: i for i, m in enumerate(order)}

    rows = []
    for m in order:
        if m not in df.index:
            continue

        kg = float(blend.get(m, 0.0))
        price = float(df.loc[m, "Price_Rs_t"])
        cost = kg * price / 1000.0

        rows.append({
            "Material": m,
            "Group": str(df.loc[m, "Group"]),
            "kg/t": kg,
            "% of Burden": (kg / total * 100.0 if total else 0.0),
            "Cost Rs/t": cost,
            "% of Cost": (cost / c * 100.0 if c else 0.0)
        })

    out = pd.DataFrame(rows)

    if len(out):
        # Absolute material order — never sort by quantity/cost.
        out["_material_order"] = out["Material"].map(order_pos)
        out = (
            out.sort_values("_material_order", kind="stable")
               .drop(columns="_material_order")
               .reset_index(drop=True)
        )

    totalrow = pd.DataFrame([{
        "Material": "TOTAL",
        "Group": "",
        "kg/t": total,
        "% of Burden": 100.0,
        "Cost Rs/t": c,
        "% of Cost": 100.0
    }])

    return pd.concat([out, totalrow], ignore_index=True), c, total

def donut(vals,center,unit):
    d=pd.DataFrame({"Group":GROUPS,"Value":[vals.get(g,0) for g in GROUPS]}); d=d[d.Value>0]
    fig=px.pie(d,names="Group",values="Value",hole=.65,color="Group",color_discrete_map=COLORS)
    fig.update_layout(height=285,margin=dict(l=2,r=2,t=2,b=2),paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)",font=dict(color="#dfe8ed",size=9),showlegend=True)
    fig.update_traces(textinfo="percent",textposition="inside")
    fig.add_annotation(text=f"<b>{center:,.1f}</b><br><span style='font-size:9px'>{unit}</span>",x=.5,y=.5,showarrow=False,font=dict(size=14,color="#fff"))
    return fig

def qpanel(ach):
    q=quality_table(ach,TARGETS); out=""
    for _,r in q.iterrows():
        ok=str(r.Status)=="OK"
        out+=f'<div class="qitem"><div class="qtop"><span>{r.KPI}</span><span class="badge {"ok" if ok else "out"}">{r.Status}</span></div><div class="qsub">Achieved <b>{float(r.Achieved):.4f}</b> | Target {r.Target}</div><div class="meter {"bad" if not ok else ""}"><div style="width:100%"></div></div></div>'
    return out

def page(title,sub): st.markdown(f'<div class="eyebrow">HOSPET ALLOY STEEL PLANT</div><h2>{title}</h2><div class="sub">{sub}</div>',unsafe_allow_html=True)

# ---------- EDITORS ----------
def chemistry_editor(key="dashboard_chemistry"):
    data = []
    for m in st.session_state.primary:
        r = st.session_state.df.loc[m]
        data.append({
            "Material": m,
            "Group": r.Group,
            "Fe (%)": r.Fe,
            "SiO₂ (%)": r.SiO2,
            "Al₂O₃ (%)": r.Al2O3,
            "CaO (%)": r.CaO,
            "MgO (%)": r.MgO,
            "LOI (%)": r.LOI,
            "Tech Min": r.Tech_Min,
            "Tech Max": r.Tech_Max,
        })

    n = len(data)
    editor_height = max(470, 38 * (n + 1) + 20)

    ed = st.data_editor(
        pd.DataFrame(data),
        hide_index=True,
        use_container_width=True,
        height=editor_height,
        key=key,
        disabled=["Material", "Group"],
        column_config={
            "Material": st.column_config.TextColumn("Material"),
            "Group": st.column_config.TextColumn("Group"),
            "Fe (%)": st.column_config.NumberColumn("Fe (%) ✎", min_value=0, step=.01, format="%.2f"),
            "SiO₂ (%)": st.column_config.NumberColumn("SiO₂ (%) ✎", min_value=0, step=.01, format="%.2f"),
            "Al₂O₃ (%)": st.column_config.NumberColumn("Al₂O₃ (%) ✎", min_value=0, step=.01, format="%.2f"),
            "CaO (%)": st.column_config.NumberColumn("CaO (%) ✎", min_value=0, step=.01, format="%.2f"),
            "MgO (%)": st.column_config.NumberColumn("MgO (%) ✎", min_value=0, step=.01, format="%.2f"),
            "LOI (%)": st.column_config.NumberColumn("LOI (%) ✎", min_value=0, step=.01, format="%.2f"),
            "Tech Min": st.column_config.NumberColumn("Tech Min", min_value=0, step=1, format="%.0f"),
            "Tech Max": st.column_config.NumberColumn("Tech Max ✎", min_value=0, step=1, format="%.0f"),
        }
    )

    for _, r in ed.iterrows():
        m = r["Material"]
        for src_col, dst_col in [
            ("Fe (%)", "Fe"), ("SiO₂ (%)", "SiO2"), ("Al₂O₃ (%)", "Al2O3"),
            ("CaO (%)", "CaO"), ("MgO (%)", "MgO"), ("LOI (%)", "LOI"),
            ("Tech Min", "Tech_Min"), ("Tech Max", "Tech_Max")
        ]:
            st.session_state.df.loc[m, dst_col] = float(r[src_col])

    st.session_state.changed = True


def primary_editor(key):
    data = []
    for m in st.session_state.primary:
        if m not in st.session_state.df.index:
            continue
        r = st.session_state.df.loc[m]
        data.append({
            "Material": m,
            "Availability": st.session_state.avail.get(m, True),
            "Price (₹/t)": r.Price_Rs_t,
            "RM Stock (t)": r.Available_Tonnes,
            "Tech Max (kg/t)": r.Tech_Max
        })

    ed = st.data_editor(
        pd.DataFrame(data),
        hide_index=True,
        use_container_width=True,
        height=max(300, 38 * (len(data) + 1) + 20),
        key=key,
        disabled=["Material"],
        column_config={
            "Availability": st.column_config.CheckboxColumn(
                "Availability ●",
                help="OFF excludes this primary material from optimization."
            ),
            "Price (₹/t)": st.column_config.NumberColumn(
                "Price ₹/t ✎", min_value=0, step=1, format="₹ %.0f"
            ),
            "RM Stock (t)": st.column_config.NumberColumn(
                "RM Stock t ✎", min_value=0, step=100, format="%.0f"
            ),
            "Tech Max (kg/t)": st.column_config.NumberColumn(
                "Tech Max kg/t ✎", min_value=0, step=1, format="%.0f"
            )
        }
    )

    for _, r in ed.iterrows():
        m = r["Material"]
        st.session_state.df.loc[m, "Price_Rs_t"] = float(r["Price (₹/t)"])
        st.session_state.df.loc[m, "Available_Tonnes"] = float(r["RM Stock (t)"])
        st.session_state.df.loc[m, "Tech_Max"] = float(r["Tech Max (kg/t)"])
        st.session_state.avail[m] = bool(r["Availability"])

    st.session_state.changed = True

def alt_editor():
    if not st.session_state.alts:
        st.info("No alternative materials are loaded. Add Material_Type = Alternative rows to the Master Excel.")
        return

    data = []
    for m in st.session_state.alts:
        r = st.session_state.df.loc[m]
        data.append({
            "Material": m,
            "Include in Mix": st.session_state.alt_on.get(m, False),
            "Fe": r.Fe, "SiO2": r.SiO2, "Al2O3": r.Al2O3,
            "CaO": r.CaO, "MgO": r.MgO, "LOI": r.LOI,
            "Price (₹/t)": r.Price_Rs_t,
            "RM Stock (t)": r.Available_Tonnes,
            "Tech Min": r.Tech_Min,
            "Tech Max (kg/t)": r.Tech_Max
        })

    ed = st.data_editor(
        pd.DataFrame(data),
        hide_index=True,
        use_container_width=True,
        height=max(250, 38 * (len(data) + 1) + 20),
        key="alt_editor",
        disabled=["Material"],
        column_config={
            "Include in Mix": st.column_config.CheckboxColumn(
                "Include in Mix ●",
                help="OFF = completely excluded. ON = eligible, but optimizer does not have to use it."
            ),
            "Fe": st.column_config.NumberColumn("Fe ✎", min_value=0, step=.01),
            "SiO2": st.column_config.NumberColumn("SiO₂ ✎", min_value=0, step=.01),
            "Al2O3": st.column_config.NumberColumn("Al₂O₃ ✎", min_value=0, step=.01),
            "CaO": st.column_config.NumberColumn("CaO ✎", min_value=0, step=.01),
            "MgO": st.column_config.NumberColumn("MgO ✎", min_value=0, step=.01),
            "LOI": st.column_config.NumberColumn("LOI ✎", min_value=0, step=.01),
            "Price (₹/t)": st.column_config.NumberColumn("Price ₹/t ✎", min_value=0, step=1, format="₹ %.0f"),
            "RM Stock (t)": st.column_config.NumberColumn("RM Stock t ✎", min_value=0, step=100),
            "Tech Min": st.column_config.NumberColumn("Tech Min ✎", min_value=0, step=1),
            "Tech Max (kg/t)": st.column_config.NumberColumn("Tech Max kg/t ✎", min_value=0, step=1)
        }
    )

    for _, r in ed.iterrows():
        m = r["Material"]
        st.session_state.alt_on[m] = bool(r["Include in Mix"])
        for c in ["Fe","SiO2","Al2O3","CaO","MgO","LOI"]:
            st.session_state.df.loc[m,c] = float(r[c])
        st.session_state.df.loc[m,"Tech_Min"] = float(r["Tech Min"])
        st.session_state.df.loc[m,"Tech_Max"] = float(r["Tech Max (kg/t)"])
        st.session_state.df.loc[m,"Price_Rs_t"] = float(r["Price (₹/t)"])
        st.session_state.df.loc[m,"Available_Tonnes"] = float(r["RM Stock (t)"])
        st.session_state.avail[m] = bool(r["Include in Mix"])

    st.session_state.changed = True

def alternative():
    page(
        "Alternative Raw Material",
        "Contingency materials from the same Master Chemistry Excel."
    )
    st.markdown(
        '<div class="notice">OFF = excluded from optimization. ON = eligible but not forced. '
        'Chemistry, price, RM Stock and Tech Max are editable here. Material names are taken from the uploaded Excel.</div>',
        unsafe_allow_html=True
    )
    st.write("")
    st.markdown(
        '<div class="panel"><div class="panel-title">'
        'ALTERNATIVE RAW MATERIALS — CONTINGENCY CONTROL</div>',
        unsafe_allow_html=True
    )
    alt_editor()
    st.markdown('</div>', unsafe_allow_html=True)


def composition(kind):
    if not result or not result["blend"]: page("Composition","Run optimizer first."); return
    bd,cost,total=breakdown(result["blend"],result["df"]); base=bd[bd.Material!="TOTAL"]
    vals={g:float(base.loc[base.Group==g,"Cost Rs/t"].sum() if kind=="cost" else base.loc[base.Group==g,"kg/t"].sum()) for g in GROUPS}
    title="Cost Structure & Cost Drivers" if kind=="cost" else "Burden Mix & Material Contribution"; unit="₹/t" if kind=="cost" else "kg/t"; center=cost if kind=="cost" else total
    page(title,"Sequential view: Iron Ore → Flux → Recycle → Fuel.")
    a,b=st.columns([1,1.6])
    with a: st.markdown(f'<div class="panel"><div class="panel-title">{kind.upper()} DISTRIBUTION</div>',unsafe_allow_html=True); st.plotly_chart(donut(vals,center,unit),use_container_width=True,config={"displayModeBar":False}); st.markdown('</div>',unsafe_allow_html=True)
    with b:
        rows=[{"Material":LABEL[g],"Group":g,("Cost Rs/t" if kind=="cost" else "kg/t"):vals[g],("% of Cost" if kind=="cost" else "% of Burden"):(vals[g]/center*100 if center else 0)} for g in GROUPS]
        rows.append({"Material":"TOTAL","Group":"","Cost Rs/t":cost,"% of Cost":100} if kind=="cost" else {"Material":"TOTAL","Group":"","kg/t":total,"% of Burden":100})
        st.markdown('<div class="panel"><div class="panel-title">GROUP CONTRIBUTION</div>'+table(pd.DataFrame(rows),{"Cost Rs/t"})+'</div>',unsafe_allow_html=True)
    st.write(""); st.markdown('<div class="panel"><div class="panel-title">MATERIAL LEVEL</div>'+table(bd.round(2),{"Cost Rs/t"})+'</div>',unsafe_allow_html=True)

def results():
    page("Optimized Sinter Recipe","Latest solver result.")
    if not result or not result["blend"]: st.info("Run optimizer first."); return
    bd,cost,total=breakdown(result["blend"],result["df"]); a=result["achieved"]; ok=all(quality_checks(a,TARGETS).values())
    cc=st.columns(4)
    for col,l,v,s,cl in [(cc[0],"TOTAL COST",f"₹{cost:,.2f}/t","Current optimum","kpi-s"),(cc[1],"BURDEN",f"{total:,.1f} kg/t","Optimized","kpi-g"),(cc[2],"Fe",f"{a['Fe']:.3f}%","Target band","kpi-a"),(cc[3],"QUALITY","PASS" if ok else "REVIEW","Gate","kpi-g" if ok else "kpi-r")]: col.markdown(f'<div class="kpi {cl}"><div class="kpi-label">{l}</div><div class="kpi-value">{v}</div><div class="kpi-sub">{s}</div></div>',unsafe_allow_html=True)
    st.write(""); a,b=st.columns([1,2]); a.markdown('<div class="panel"><div class="panel-title">QUALITY</div>'+qpanel(a=a if False else result["achieved"])+'</div>',unsafe_allow_html=True); b.markdown('<div class="panel"><div class="panel-title">RECIPE</div>'+table(bd.round(2),{"Cost Rs/t"})+'</div>',unsafe_allow_html=True)

def manual():
    page("Manual Burden Control","Adjust the optimized burden and see burden, cost and chemistry impact inside this tab only.")
    if not result or not result["blend"]: st.info("Run optimizer first."); return
    df=result["df"]; base=st.session_state.manual_base or result["blend"].copy(); st.session_state.manual_base=base.copy()
    adj=[m for m in base if df.loc[m,"Group"] in ("Iron_ore","Flux")]; fixed=[m for m in base if df.loc[m,"Group"] in ("Recycle","Fuel")]
    st.markdown('<div class="notice">Iron Ore ±15% • Flux ±10% • Recycle/Fuel fixed • total burden preserved.</div>',unsafe_allow_html=True)
    req={}; cols=st.columns(2)
    for i,m in enumerate(adj):
        b=float(base[m]); r=.15 if df.loc[m,"Group"]=="Iron_ore" else .10; mn=max(0,b*(1-r)); mx=max(mn+1,b*(1+r)); key="man_"+m
        try:
            current = float(st.session_state.get(key, b))
        except (TypeError, ValueError):
            current = b
        current = max(float(mn), min(float(current), float(mx)))
        st.session_state[key] = current
        with cols[i%2]:
            req[m] = st.slider(
                f"{m} — kg/t",
                min_value=float(mn),
                max_value=float(mx),
                value=float(current),
                step=0.5,
                key=key
            )
    adjusted=redistribute_adjustment(base,df,req)
    for m in fixed: adjusted[m]=base[m]
    ach=compute_achieved(adjusted,df,1000); ac=sum(adjusted[m]*df.loc[m,"Price_Rs_t"]/1000 for m in adjusted); bc=float(result["cost"] or 0); total=sum(adjusted); ok=all(quality_checks(ach,TARGETS).values())
    cc=st.columns(5)
    for col,l,v,s,cl in [(cc[0],"BASE COST",f"₹{bc:,.2f}","Optimized","kpi-s"),(cc[1],"ADJUSTED COST",f"₹{ac:,.2f}",f"{ac-bc:+,.2f}","kpi-o"),(cc[2],"BURDEN",f"{total:,.1f}","Preserved","kpi-g"),(cc[3],"Fe",f"{ach['Fe']:.3f}%","After adjustment","kpi-a"),(cc[4],"QUALITY","PASS" if ok else "REVIEW","After adjustment","kpi-g" if ok else "kpi-r")]: col.markdown(f'<div class="kpi {cl}"><div class="kpi-label">{l}</div><div class="kpi-value">{v}</div><div class="kpi-sub">{s}</div></div>',unsafe_allow_html=True)
    st.write(""); a,b=st.columns(2)
    gv={g:sum(adjusted.get(m,0) for m in adjusted if df.loc[m,"Group"]==g) for g in GROUPS}; gc={g:sum(adjusted.get(m,0)*df.loc[m,"Price_Rs_t"]/1000 for m in adjusted if df.loc[m,"Group"]==g) for g in GROUPS}
    with a: st.markdown('<div class="panel"><div class="panel-title">ADJUSTED BURDEN COMPOSITION</div>',unsafe_allow_html=True);st.plotly_chart(donut(gv,total,"kg/t"),use_container_width=True,config={"displayModeBar":False});st.markdown('</div>',unsafe_allow_html=True)
    with b: st.markdown('<div class="panel"><div class="panel-title">ADJUSTED COST COMPOSITION</div>',unsafe_allow_html=True);st.plotly_chart(donut(gc,ac,"₹/t"),use_container_width=True,config={"displayModeBar":False});st.markdown('</div>',unsafe_allow_html=True)
    st.write(""); a,b=st.columns([1,2])
    with a: st.markdown('<div class="panel"><div class="panel-title">ACHIEVED CHEMISTRY</div>'+qpanel(ach)+'</div>',unsafe_allow_html=True)
    with b:
        bd,_,_=breakdown(adjusted,df); st.markdown('<div class="panel"><div class="panel-title">ADJUSTED BURDEN & COST</div>'+table(bd.round(2),{"Cost Rs/t"})+'</div>',unsafe_allow_html=True)
    st.write(""); st.markdown('<div class="panel"><div class="panel-title">OPTIMIZED vs MANUAL</div>',unsafe_allow_html=True)
    cmp=pd.DataFrame([["Cost",bc,ac,ac-bc],["Burden",sum(base.values()),total,total-sum(base.values())],["Fe",result["achieved"]["Fe"],ach["Fe"],ach["Fe"]-result["achieved"]["Fe"]],["SiO2",result["achieved"]["SiO2"],ach["SiO2"],ach["SiO2"]-result["achieved"]["SiO2"]],["Basicity",result["achieved"]["Basicity"],ach["Basicity"],ach["Basicity"]-result["achieved"]["Basicity"]]],columns=["Parameter","Optimized","Manual","Change"])
    st.dataframe(cmp.round(4),hide_index=True,use_container_width=True); st.markdown('</div>',unsafe_allow_html=True)
    if st.button("↩ RESET TO OPTIMIZED",use_container_width=True):
        for m in adj: st.session_state["man_"+m]=float(base[m])
        st.rerun()

def whatif():
    page("Scenario & Material Risk","Test material unavailability.")
    if st.button("▶ RUN MATERIAL SHORTAGE SCENARIOS",type="primary"):
        with st.spinner("Evaluating…"): st.session_state.whatif=what_if_analysis(active_df(),TARGETS)
    if st.session_state.whatif is not None:
        w=st.session_state.whatif.copy(); st.markdown('<div class="panel">'+table(w.fillna("—"),status="Status")+'</div>',unsafe_allow_html=True)
    else: st.info("Run the scenario analysis.")

def bottleneck():
    page("Quality Constraint Pressure","Identify the constraints closest to limits.")
    if not result or not result["achieved"]: st.info("Run optimizer first."); return
    q=quality_table(result["achieved"],TARGETS); st.markdown('<div class="panel">'+table(q.round(4),status="Status")+'</div>',unsafe_allow_html=True)

def reports():
    page("Reports & Export","Export the latest optimized recipe.")
    if not result or not result["blend"]: st.info("Run optimizer first."); return
    bd,_,_=breakdown(result["blend"],result["df"]); st.markdown('<div class="panel">'+table(bd.round(2),{"Cost Rs/t"})+'</div>',unsafe_allow_html=True)
    st.download_button("⬇ DOWNLOAD OPTIMIZATION REPORT",bd.to_csv(index=False).encode(),"sinter_optimization_report.csv","text/csv",use_container_width=True)

def settings():
    page("Upload & Settings","Secondary upload/settings page. The primary Master Chemistry Excel is also permanently available on the Dashboard.")
    a,b=st.columns([1.4,1])
    with a:
        st.markdown('<div class="panel"><div class="panel-title">PRIMARY CHEMISTRY EXCEL</div>',unsafe_allow_html=True)
        f=st.file_uploader("Upload primary chemistry",type=["xlsx"],key="primary_upload")
        if f:
            try:
                df=load_primary(f); st.success(f"{len(df)} materials validated.")
                if st.button("ACTIVATE PRIMARY CHEMISTRY",type="primary",use_container_width=True): reset_primary(df,"Uploaded • "+f.name); st.rerun()
            except Exception as e: st.error(str(e))
        st.markdown('<div class="small">Excel contains chemistry + Tech Min/Max. Optional Price_Rs_t and Available_Tonnes can also be loaded; all commercial values remain editable in the dashboard.</div>',unsafe_allow_html=True); st.markdown('</div>',unsafe_allow_html=True)
    with b:
        st.markdown('<div class="panel"><div class="panel-title">SYSTEM</div>',unsafe_allow_html=True)
        if st.button("↺ RESTORE BUILT-IN MASTER",use_container_width=True): reset_primary(get_default_chemistry(),"Built-in Master Chemistry"); st.rerun()
        st.markdown('<div class="notice">Primary chemistry can be uploaded from the Dashboard at any time. After activation, chemistry remains editable in the Dashboard chemistry table.</div>',unsafe_allow_html=True); st.markdown('</div>',unsafe_allow_html=True)

# ---------- ROUTE ----------
pages={"Dashboard":dashboard,"RM Stock":rm_stock,"Optimization Results":results,"Manual Burden Control":manual,"Alternative Raw Material":alternative,"Burden Composition":lambda:composition("burden"),"Cost Composition":lambda:composition("cost"),"What-if Analysis":whatif,"Bottleneck Analysis":bottleneck,"Reports":reports,"Upload & Settings":settings}
pages[st.session_state.nav]()
st.markdown('<div class="footer">Sinter Burden Control • Hospet Alloy Steel Plant • Production decision-support interface</div>',unsafe_allow_html=True)
