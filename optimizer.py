
"""
Sinter Burden Optimizer v22.4
Core optimization engine extracted/adapted from the supplied Colab backend.

This module contains:
- Built-in/default chemistry
- Optional Excel loading
- Material availability/compensation rules
- PuLP/CBC optimization
- Quality calculations
- Manual burden redistribution helpers
- What-if scenario analysis

The Streamlit UI lives in app.py.
"""

from io import BytesIO
import pandas as pd
import pulp

# ============================================================================
# 1. MATERIAL RANKING & COMPENSATION RULES
# ============================================================================

IRON_ORE_RANK = {
    "Lloyds_HG": 1, "MILL_SCALE": 2,
    "SIOM_MG": 3, "KIOM_MG": 3, "DIOM_LG": 4
}
FLUX_RANK = {"LIMESTONE": 1, "DOLOMITE": 2, "QUICKLIME": 3}

MILL_SCALE_MAX_BURDEN_PCT = 0.15

IRON_ORE_MAX_PCT_BASE = {
    "Lloyds_HG": 0.25, "MILL_SCALE": 0.29,
    "SIOM_MG": 0.29, "KIOM_MG": 0.29, "DIOM_LG": 0.29
}
IRON_ORE_MAX_PCT_RELAXED = {
    "Lloyds_HG": 0.35, "MILL_SCALE": 0.35,
    "SIOM_MG": 0.40, "KIOM_MG": 0.40, "DIOM_LG": 0.40
}
IRON_ORE_MAX_PCT_EMERGENCY = {
    "Lloyds_HG": 0.45, "MILL_SCALE": 0.45,
    "SIOM_MG": 0.50, "KIOM_MG": 0.50, "DIOM_LG": 0.50
}
IRON_ORE_MIN_PCT = {
    "Lloyds_HG": 0.03, "MILL_SCALE": 0.03,
    "SIOM_MG": 0.03, "KIOM_MG": 0.03, "DIOM_LG": 0.03
}
MAX_IRON_ORE_PORTION = 0.80

FLUX_MAX_PCT_BASE = {"LIMESTONE": 0.60, "DOLOMITE": 0.45, "QUICKLIME": 0.35}
FLUX_MAX_PCT_RELAXED = {"LIMESTONE": 0.75, "DOLOMITE": 0.55, "QUICKLIME": 0.45}
FLUX_MAX_PCT_EMERGENCY = {"LIMESTONE": 0.90, "DOLOMITE": 0.65, "QUICKLIME": 0.55}
FLUX_MIN_PCT = {"LIMESTONE": 0.05, "DOLOMITE": 0.05, "QUICKLIME": 0.02}
MAX_FLUX_PORTION = 0.25

FE_TARGET = 54.0
FE_TOLERANCE = 0.3
FE_LOWER = FE_TARGET - FE_TOLERANCE
FE_UPPER = FE_TARGET + FE_TOLERANCE
FE_CENTER_WEIGHT = 2.0

DEVIATION_WEIGHTS = {
    "Fe": 5.0,
    "Basicity": 6.0,
    "CaO": 5.0,
    "MgO": 4.0,
    "Al2O3": 3.0,
    "SiO2": 2.0,
    "Al2O3_SiO2_ratio": 2.0,
}

PIN_TOLERANCE = 1e-3
FLUX_BASELINE_INCREASE_CAP = 0.05

ADJUSTMENT_RANGES = {
    "Iron_ore": 0.15,
    "Flux": 0.10,
    "Recycle": 0.00,
    "Fuel": 0.00,
}

TARGETS = {
    "Fe_min": 54.0,
    "SiO2_max": 5.8,
    "Al2O3_max": 4.5,
    "Al2O3_SiO2_max": 0.98,
    "Basicity_min": 1.9,
    "Basicity_max": 2.0,
    "MgO_min": 2.2,
    "MgO_max": 2.4,
    "CaO_min": 10.5,
    "CaO_max": 11.5,
}

# ============================================================================
# 2. DEFAULT CHEMISTRY
# ============================================================================

def get_default_chemistry():
    data = {
        "Material": [
            "MILL_SCALE", "Lloyds_HG", "DIOM_LG", "SIOM_MG", "KIOM_MG",
            "Solid_Waste", "IOL_Fines", "FLUE_DUST",
            "DOLOMITE", "LIMESTONE", "QUICKLIME", "COKE_BREEZE"
        ],
        "Group": [
            "Iron_ore", "Iron_ore", "Iron_ore", "Iron_ore", "Iron_ore",
            "Recycle", "Recycle", "Recycle",
            "Flux", "Flux", "Flux", "Fuel"
        ],
        "Fe": [68.34, 63.52, 57.17, 59.34, 58.41, 50.0, 60.00, 47.02, 0.54, 0.88, 0.01, 0],
        "SiO2": [2.00, 3.86, 12.39, 6.92, 5.75, 6.00, 5.00, 7.07, 4.72, 4.48, 2.50, 2.8],
        "Al2O3": [2.72, 2.27, 2.93, 3.72, 5.48, 4.50, 3.00, 4.50, 0.95, 1.19, 0.61, 0],
        "CaO": [0, 0.022, 0.058, 0.256, 0.157, 1.122, 8.79, 1.10, 30.02, 48.71, 89.00, 0],
        "MgO": [0, 0.034, 0.114, 0.331, 0.018, 0.06, 1.52, 0.29, 18.75, 2.59, 1.57, 0],
        "LOI": [2.50, 2.29, 4.00, 3.45, 4.62, 3.00, 3.00, 15.00, 42.00, 40.00, 5.00, 70.00],
        "Tech_Min": [0, 0, 0, 0, 0, 30, 120, 25, 30, 0, 40, 65],
        "Tech_Max": [220, 200, 200, 200, 300, 30, 120, 25, 200, 250, 65, 75],
        "Available_Tonnes": [2000, 10000, 6000, 8000, 5000, 5000, 5000, 3000, 10000, 15000, 5000, 9999],
        "Price_Rs_t": [7800, 7820, 4600, 4600, 4900, 1000, 5577, 500, 1340, 1355, 9200, 15022],
    }
    df = pd.DataFrame(data).set_index("Material")
    for mat in df[df["Group"] == "Recycle"].index:
        fixed_rate = df.loc[mat, "Tech_Min"]
        df.loc[mat, "Tech_Max"] = fixed_rate
    return df


# ============================================================================
# 3. OPTIONAL EXCEL LOADER
# ============================================================================

def load_chemistry_from_excel(uploaded_file):
    """Load the same schema used by the supplied backend.

    Excel is optional. The app can always fall back to get_default_chemistry().
    """
    df = pd.read_excel(uploaded_file, index_col="Material")
    required_cols = [
        "Group", "Fe", "SiO2", "Al2O3", "CaO", "MgO", "LOI",
        "Tech_Min", "Tech_Max", "Available_Tonnes", "Price_Rs_t"
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError("Excel missing required columns: " + ", ".join(missing))

    rename_map = {"INTERNAL_FINES": "Solid_Waste", "RETURN_SINTER": "IOL_Fines"}
    df = df.rename(index=rename_map)

    for mat in df[df["Group"] == "Recycle"].index:
        fixed_rate = df.loc[mat, "Tech_Min"]
        df.loc[mat, "Tech_Max"] = fixed_rate

    if "LIME_POWDER" in df.index:
        df = df.drop("LIME_POWDER")

    return df


# ============================================================================
# 4. HELPERS
# ============================================================================

def get_iron_ore_tier(df, iron_ores):
    unavailable = [
        m for m in iron_ores
        if df.loc[m, "Available_Tonnes"] <= 0 or df.loc[m, "Tech_Max"] == 0
    ]
    n = len(unavailable)
    if n == 0:
        return IRON_ORE_MAX_PCT_BASE.copy(), unavailable, "All iron ores available"
    if n <= 2:
        return IRON_ORE_MAX_PCT_RELAXED.copy(), unavailable, f"{n} iron ore(s) unavailable"
    return IRON_ORE_MAX_PCT_EMERGENCY.copy(), unavailable, f"{n} iron ores unavailable"


def get_flux_tier(df, fluxes):
    unavailable = [
        m for m in fluxes
        if df.loc[m, "Available_Tonnes"] <= 0 or df.loc[m, "Tech_Max"] == 0
    ]
    n = len(unavailable)
    if n == 0:
        return FLUX_MAX_PCT_BASE.copy(), unavailable, "All fluxes available"
    if n == 1:
        return FLUX_MAX_PCT_RELAXED.copy(), unavailable, "1 flux unavailable"
    return FLUX_MAX_PCT_EMERGENCY.copy(), unavailable, f"{n} fluxes unavailable"


def check_fuel_gate(df):
    fuels = [m for m in df.index if df.loc[m, "Group"] == "Fuel"]
    problems = []
    for mat in fuels:
        tech_min = df.loc[mat, "Tech_Min"]
        available = df.loc[mat, "Available_Tonnes"]
        tech_max = df.loc[mat, "Tech_Max"]
        if tech_min > 0 and (available <= 0 or tech_max <= 0):
            problems.append(
                f"{mat}: requires >= {tech_min} kg/t; "
                f"available={available}, Tech_Max={tech_max}"
            )
    return (False, problems) if problems else (True, [])


def build_bounds(df, production_tonnes):
    bounds = {}
    for mat in df.index:
        tech_min = float(df.loc[mat, "Tech_Min"])
        tech_max = float(df.loc[mat, "Tech_Max"])
        available = float(df.loc[mat, "Available_Tonnes"])
        if available <= 0 or tech_max == 0:
            bounds[mat] = (0, 0)
        else:
            inv_cap = (available / production_tonnes) * 1000
            bounds[mat] = (tech_min, min(tech_max, inv_cap))
    return bounds


def add_structural_constraints(
    prob, x, df, bounds, iron_ores, fluxes, iron_ore_max_pct,
    unavailable_iron, flux_max_pct, unavailable_flux, OUT,
    baseline_flux_portion=None
):
    non_fuel = [m for m in x if df.loc[m, "Group"] != "Fuel"]
    mass = pulp.lpSum(
        x[m] * (1 - df.loc[m, "LOI"] / 100) for m in non_fuel
    )
    prob += mass >= OUT - 2, "Mass_Balance_Lower"
    prob += mass <= OUT + 2, "Mass_Balance_Upper"

    total_iron_ore = pulp.lpSum(x[m] for m in iron_ores)
    total_flux = pulp.lpSum(x[m] for m in fluxes)
    total_burden = pulp.lpSum(
        x[m] for m in df.index if df.loc[m, "Group"] != "Fuel"
    )

    for mat in iron_ores:
        if mat in unavailable_iron:
            prob += x[mat] == 0, f"{mat}_unavailable"
        else:
            max_pct = iron_ore_max_pct.get(mat, 0.29)
            prob += x[mat] <= max_pct * total_iron_ore + 0.001, f"{mat}_max_pct"
            prob += x[mat] >= IRON_ORE_MIN_PCT.get(mat, 0.03) * total_iron_ore - 0.001, f"{mat}_min_pct"

    if "MILL_SCALE" in x:
        prob += x["MILL_SCALE"] <= MILL_SCALE_MAX_BURDEN_PCT * total_burden, "MILL_SCALE_Burden_Cap"

    prob += total_iron_ore <= MAX_IRON_ORE_PORTION * OUT, "Max_Iron_Ore_Portion"

    for mat in fluxes:
        if mat in unavailable_flux:
            prob += x[mat] == 0, f"{mat}_unavailable"
        else:
            max_pct = flux_max_pct.get(mat, 0.5)
            min_pct = FLUX_MIN_PCT.get(mat, 0.02)
            prob += x[mat] <= max_pct * total_flux + 0.001, f"{mat}_max_pct"
            prob += x[mat] >= min_pct * total_flux - 0.001, f"{mat}_min_pct"

    prob += total_flux <= MAX_FLUX_PORTION * OUT, "Max_Flux_Portion"

    if baseline_flux_portion is not None:
        prob += (
            total_flux <= (baseline_flux_portion + FLUX_BASELINE_INCREASE_CAP) * OUT,
            "Flux_Baseline_Cap"
        )

    return total_iron_ore, total_flux, total_burden


def compute_achieved(blend, df, OUT=1000):
    Fe = sum(blend[m] * df.loc[m, "Fe"] / 100 for m in blend) / OUT * 100
    SiO2 = sum(blend[m] * df.loc[m, "SiO2"] / 100 for m in blend) / OUT * 100
    Al2O3 = sum(blend[m] * df.loc[m, "Al2O3"] / 100 for m in blend) / OUT * 100
    CaO = sum(blend[m] * df.loc[m, "CaO"] / 100 for m in blend) / OUT * 100
    MgO = sum(blend[m] * df.loc[m, "MgO"] / 100 for m in blend) / OUT * 100

    achieved = {"Fe": Fe, "SiO2": SiO2, "Al2O3": Al2O3, "CaO": CaO, "MgO": MgO}

    if SiO2 > 0:
        achieved["Basicity"] = CaO / SiO2
        achieved["Al2O3/SiO2"] = Al2O3 / SiO2
        achieved["B4"] = (CaO + MgO) / (SiO2 + Al2O3)
    else:
        achieved["Basicity"] = 0
        achieved["Al2O3/SiO2"] = 0
        achieved["B4"] = 0

    return achieved


def build_soft_vars_and_constraints(prob, xr, df, OUT, targets, fe_lo, fe_hi, suffix=""):
    Fe_s = pulp.lpSum(xr[m] * df.loc[m, "Fe"] / 100 for m in xr)
    SiO2_s = pulp.lpSum(xr[m] * df.loc[m, "SiO2"] / 100 for m in xr)
    Al2O3_s = pulp.lpSum(xr[m] * df.loc[m, "Al2O3"] / 100 for m in xr)
    CaO_s = pulp.lpSum(xr[m] * df.loc[m, "CaO"] / 100 for m in xr)
    MgO_s = pulp.lpSum(xr[m] * df.loc[m, "MgO"] / 100 for m in xr)

    Fe_under = pulp.LpVariable(f"Fe_under{suffix}", lowBound=0)
    Fe_over = pulp.LpVariable(f"Fe_over{suffix}", lowBound=0)
    SiO2_over = pulp.LpVariable(f"SiO2_over{suffix}", lowBound=0)
    Al2O3_over = pulp.LpVariable(f"Al2O3_over{suffix}", lowBound=0)
    ratio_over = pulp.LpVariable(f"ratio_over{suffix}", lowBound=0)
    Bas_under = pulp.LpVariable(f"Bas_under{suffix}", lowBound=0)
    Bas_over = pulp.LpVariable(f"Bas_over{suffix}", lowBound=0)
    MgO_under = pulp.LpVariable(f"MgO_under{suffix}", lowBound=0)
    MgO_over = pulp.LpVariable(f"MgO_over{suffix}", lowBound=0)
    CaO_under = pulp.LpVariable(f"CaO_under{suffix}", lowBound=0)
    CaO_over = pulp.LpVariable(f"CaO_over{suffix}", lowBound=0)
    Fe_center_dev = pulp.LpVariable(f"Fe_center_dev{suffix}", lowBound=0)

    prob += Fe_s + Fe_under >= fe_lo, f"Fe_lo_soft{suffix}"
    prob += Fe_s - Fe_over <= fe_hi, f"Fe_hi_soft{suffix}"
    prob += SiO2_s - SiO2_over <= targets["SiO2_max"] * OUT / 100, f"SiO2_soft{suffix}"
    prob += Al2O3_s - Al2O3_over <= targets["Al2O3_max"] * OUT / 100, f"Al2O3_soft{suffix}"
    prob += (Al2O3_s - targets["Al2O3_SiO2_max"] * SiO2_s) - ratio_over <= 0, f"Ratio_soft{suffix}"
    prob += (CaO_s - targets["Basicity_min"] * SiO2_s) + Bas_under >= 0, f"Basicity_lo_soft{suffix}"
    prob += (CaO_s - targets["Basicity_max"] * SiO2_s) - Bas_over <= 0, f"Basicity_hi_soft{suffix}"
    prob += MgO_s + MgO_under >= targets["MgO_min"] * OUT / 100, f"MgO_lo_soft{suffix}"
    prob += MgO_s - MgO_over <= targets["MgO_max"] * OUT / 100, f"MgO_hi_soft{suffix}"
    prob += CaO_s + CaO_under >= targets["CaO_min"] * OUT / 100, f"CaO_lo_soft{suffix}"
    prob += CaO_s - CaO_over <= targets["CaO_max"] * OUT / 100, f"CaO_hi_soft{suffix}"

    prob += Fe_s - (FE_TARGET * OUT / 100) <= Fe_center_dev, f"Fe_center_pos{suffix}"
    prob += (FE_TARGET * OUT / 100) - Fe_s <= Fe_center_dev, f"Fe_center_neg{suffix}"

    slacks = {
        "Fe_under": Fe_under, "Fe_over": Fe_over,
        "SiO2_over": SiO2_over, "Al2O3_over": Al2O3_over,
        "ratio_over": ratio_over,
        "Bas_under": Bas_under, "Bas_over": Bas_over,
        "MgO_under": MgO_under, "MgO_over": MgO_over,
        "CaO_under": CaO_under, "CaO_over": CaO_over,
        "Fe_center_dev": Fe_center_dev,
    }
    sums = {"Fe": Fe_s, "SiO2": SiO2_s, "Al2O3": Al2O3_s, "CaO": CaO_s, "MgO": MgO_s}
    return slacks, sums


def weighted_deviation_expr(slacks, targets, OUT):
    W = DEVIATION_WEIGHTS
    return (
        W["Fe"] * ((slacks["Fe_under"] + slacks["Fe_over"]) / (FE_TOLERANCE * OUT / 100)) +
        W["SiO2"] * (slacks["SiO2_over"] / (targets["SiO2_max"] * OUT / 100)) +
        W["Al2O3"] * (slacks["Al2O3_over"] / (targets["Al2O3_max"] * OUT / 100)) +
        W["Al2O3_SiO2_ratio"] * (
            slacks["ratio_over"] / max(targets["Al2O3_SiO2_max"] * OUT / 100, 1e-6)
        ) +
        W["Basicity"] * (
            (slacks["Bas_under"] + slacks["Bas_over"]) /
            max((targets["Basicity_max"] - targets["Basicity_min"]) * OUT / 100, 1e-6)
        ) +
        W["MgO"] * (
            (slacks["MgO_under"] + slacks["MgO_over"]) /
            max((targets["MgO_max"] - targets["MgO_min"]) * OUT / 100, 1e-6)
        ) +
        W["CaO"] * (
            (slacks["CaO_under"] + slacks["CaO_over"]) /
            max((targets["CaO_max"] - targets["CaO_min"]) * OUT / 100, 1e-6)
        ) +
        FE_CENTER_WEIGHT * (
            slacks["Fe_center_dev"] / (FE_TOLERANCE * OUT / 100)
        )
    )


# ============================================================================
# 5. MAIN SOLVER
# ============================================================================

def solve_blend_with_compensation(
    df, production_tonnes=1000, targets=None, baseline_blend=None,
    enforce_b4=False, b4_min=1.8, b4_max=2.0
):
    targets = targets or TARGETS
    OUT = 1000

    iron_ores = [m for m in df.index if df.loc[m, "Group"] == "Iron_ore"]
    fluxes = [m for m in df.index if df.loc[m, "Group"] == "Flux"]

    fuel_ok, fuel_problems = check_fuel_gate(df)
    if not fuel_ok:
        diagnostics = ["PRODUCTION IMPOSSIBLE: Fuel requirement cannot be met."]
        diagnostics.extend(fuel_problems)
        return "No_Production", None, None, None, diagnostics, False

    iron_ore_max_pct, unavailable_iron, _ = get_iron_ore_tier(df, iron_ores)
    flux_max_pct, unavailable_flux, _ = get_flux_tier(df, fluxes)

    bounds = build_bounds(df, production_tonnes)
    diagnostics = []

    if unavailable_iron:
        diagnostics.append("Compensating for missing ore(s): " + ", ".join(unavailable_iron))
    if unavailable_flux:
        diagnostics.append("Compensating for missing flux(es): " + ", ".join(unavailable_flux))

    fe_lo = FE_LOWER * OUT / 100
    fe_hi = FE_UPPER * OUT / 100

    baseline_flux_portion = None
    if baseline_blend:
        baseline_flux_portion = sum(
            baseline_blend.get(m, 0) for m in fluxes
        ) / OUT

    prob = pulp.LpProblem("Sinter_Burden_Opt", pulp.LpMinimize)
    x = {
        m: pulp.LpVariable(
            f"x_{m}", lowBound=bounds[m][0], upBound=bounds[m][1]
        )
        for m in df.index
    }

    prob += pulp.lpSum(
        x[m] * df.loc[m, "Price_Rs_t"] / 1000 for m in x
    ), "Total_Cost"

    add_structural_constraints(
        prob, x, df, bounds, iron_ores, fluxes,
        iron_ore_max_pct, unavailable_iron,
        flux_max_pct, unavailable_flux, OUT,
        baseline_flux_portion
    )

    Fe_sum = pulp.lpSum(x[m] * df.loc[m, "Fe"] / 100 for m in x)
    SiO2_sum = pulp.lpSum(x[m] * df.loc[m, "SiO2"] / 100 for m in x)
    Al2O3_sum = pulp.lpSum(x[m] * df.loc[m, "Al2O3"] / 100 for m in x)
    CaO_sum = pulp.lpSum(x[m] * df.loc[m, "CaO"] / 100 for m in x)
    MgO_sum = pulp.lpSum(x[m] * df.loc[m, "MgO"] / 100 for m in x)

    prob += Fe_sum >= fe_lo
    prob += Fe_sum <= fe_hi
    prob += SiO2_sum <= targets["SiO2_max"] * OUT / 100
    prob += Al2O3_sum <= targets["Al2O3_max"] * OUT / 100
    prob += Al2O3_sum - targets["Al2O3_SiO2_max"] * SiO2_sum <= 0
    prob += CaO_sum >= targets["Basicity_min"] * SiO2_sum
    prob += CaO_sum <= targets["Basicity_max"] * SiO2_sum
    prob += MgO_sum >= targets["MgO_min"] * OUT / 100
    prob += MgO_sum <= targets["MgO_max"] * OUT / 100
    prob += CaO_sum >= targets["CaO_min"] * OUT / 100
    prob += CaO_sum <= targets["CaO_max"] * OUT / 100

    if enforce_b4:
        prob += (CaO_sum + MgO_sum) - b4_min * (SiO2_sum + Al2O3_sum) >= 0
        prob += (CaO_sum + MgO_sum) - b4_max * (SiO2_sum + Al2O3_sum) <= 0

    prob.solve(pulp.PULP_CBC_CMD(msg=0))
    status = pulp.LpStatus[prob.status]

    if status == "Optimal":
        blend = {m: round(x[m].value(), 2) for m in x}
        total_cost = pulp.value(prob.objective)
        achieved = compute_achieved(blend, df, OUT)
        return status, blend, total_cost, achieved, diagnostics, False

    # Phase 1: minimize weighted deviation
    prob1 = pulp.LpProblem("Phase1_MinDeviation", pulp.LpMinimize)
    x1 = {
        m: pulp.LpVariable(
            f"x1_{m}", lowBound=bounds[m][0], upBound=bounds[m][1]
        )
        for m in df.index
    }

    add_structural_constraints(
        prob1, x1, df, bounds, iron_ores, fluxes,
        iron_ore_max_pct, unavailable_iron,
        flux_max_pct, unavailable_flux, OUT,
        baseline_flux_portion
    )

    slacks1, _ = build_soft_vars_and_constraints(
        prob1, x1, df, OUT, targets, fe_lo, fe_hi, suffix="_p1"
    )
    prob1 += weighted_deviation_expr(slacks1, targets, OUT)
    prob1.solve(pulp.PULP_CBC_CMD(msg=0))

    if pulp.LpStatus[prob1.status] != "Optimal":
        diagnostics.append("Even relaxed problem is infeasible.")
        return "Infeasible", None, None, None, diagnostics, True

    blend1 = {m: round(x1[m].value(), 2) for m in x1}
    achieved1 = compute_achieved(blend1, df, OUT)
    phase1_slack_values = {
        k: (v.value() or 0.0) for k, v in slacks1.items()
    }

    dev_report = {
        "Fe": phase1_slack_values["Fe_under"] + phase1_slack_values["Fe_over"],
        "SiO2": phase1_slack_values["SiO2_over"],
        "Al2O3": phase1_slack_values["Al2O3_over"],
        "Al2O3/SiO2": phase1_slack_values["ratio_over"],
        "Basicity": phase1_slack_values["Bas_under"] + phase1_slack_values["Bas_over"],
        "MgO": phase1_slack_values["MgO_under"] + phase1_slack_values["MgO_over"],
        "CaO": phase1_slack_values["CaO_under"] + phase1_slack_values["CaO_over"],
    }
    diagnostics.append(
        f"Most binding constraint: {max(dev_report, key=dev_report.get)}"
    )

    # Phase 2: minimize cost while pinning phase-1 deviation
    prob2 = pulp.LpProblem("Phase2_MinCost", pulp.LpMinimize)
    x2 = {
        m: pulp.LpVariable(
            f"x2_{m}", lowBound=bounds[m][0], upBound=bounds[m][1]
        )
        for m in df.index
    }

    add_structural_constraints(
        prob2, x2, df, bounds, iron_ores, fluxes,
        iron_ore_max_pct, unavailable_iron,
        flux_max_pct, unavailable_flux, OUT,
        baseline_flux_portion
    )

    slacks2, _ = build_soft_vars_and_constraints(
        prob2, x2, df, OUT, targets, fe_lo, fe_hi, suffix="_p2"
    )

    for key, var in slacks2.items():
        p1_val = phase1_slack_values.get(key, 0.0)
        prob2 += var <= p1_val + PIN_TOLERANCE

    prob2 += pulp.lpSum(
        x2[m] * df.loc[m, "Price_Rs_t"] / 1000 for m in x2
    )
    prob2.solve(pulp.PULP_CBC_CMD(msg=0))

    if pulp.LpStatus[prob2.status] == "Optimal":
        blend2 = {m: round(x2[m].value(), 2) for m in x2}
        cost2 = sum(
            blend2[m] * df.loc[m, "Price_Rs_t"] / 1000
            for m in blend2
        )
        achieved2 = compute_achieved(blend2, df, OUT)
        return "Infeasible", blend2, cost2, achieved2, diagnostics, True

    return "Infeasible", blend1, None, achieved1, diagnostics, True


# ============================================================================
# 6. UI/DASHBOARD HELPERS
# ============================================================================

def get_baseline_blend(df, targets=None):
    targets = targets or TARGETS
    result = solve_blend_with_compensation(
        df, 1000, targets, baseline_blend=None
    )
    if result[0] == "Optimal":
        return result[1], result[2], result[3]
    return None, None, None


def calculate_cost_breakdown(blend, df):
    rows = []
    total_cost = 0.0
    total_burden = sum(blend.values())

    for mat, qty in blend.items():
        cost = qty * float(df.loc[mat, "Price_Rs_t"]) / 1000
        total_cost += cost
        rows.append({
            "Material": mat,
            "Group": df.loc[mat, "Group"],
            "kg/t": qty,
            "% of Burden": (qty / total_burden * 100) if total_burden else 0,
            "Cost Rs/t": cost,
            "% of Cost": 0.0,
        })

    out = pd.DataFrame(rows)
    if total_cost:
        out["% of Cost"] = out["Cost Rs/t"] / total_cost * 100
    return out, total_cost, total_burden


def quality_checks(achieved, targets=None):
    targets = targets or TARGETS
    checks = {
        "Fe": FE_LOWER - 0.01 <= achieved["Fe"] <= FE_UPPER + 0.01,
        "SiO2": achieved["SiO2"] <= targets["SiO2_max"] + 0.01,
        "Al2O3": achieved["Al2O3"] <= targets["Al2O3_max"] + 0.01,
        "Al2O3/SiO2": achieved["Al2O3/SiO2"] <= targets["Al2O3_SiO2_max"] + 0.005,
        "Basicity": targets["Basicity_min"] - 0.01 <= achieved["Basicity"] <= targets["Basicity_max"] + 0.01,
        "MgO": targets["MgO_min"] - 0.01 <= achieved["MgO"] <= targets["MgO_max"] + 0.01,
        "CaO": targets["CaO_min"] - 0.01 <= achieved["CaO"] <= targets["CaO_max"] + 0.01,
    }
    return checks


def quality_table(achieved, targets=None):
    targets = targets or TARGETS
    rows = [
        ("Fe (%)", achieved["Fe"], f"{FE_LOWER:.1f} – {FE_UPPER:.1f}", "Fe"),
        ("SiO2 (%)", achieved["SiO2"], f"≤ {targets['SiO2_max']:.2f}", "SiO2"),
        ("Al2O3 (%)", achieved["Al2O3"], f"≤ {targets['Al2O3_max']:.2f}", "Al2O3"),
        ("Al2O3/SiO2", achieved["Al2O3/SiO2"], f"≤ {targets['Al2O3_SiO2_max']:.2f}", "Al2O3/SiO2"),
        ("Basicity", achieved["Basicity"], f"{targets['Basicity_min']:.2f} – {targets['Basicity_max']:.2f}", "Basicity"),
        ("MgO (%)", achieved["MgO"], f"{targets['MgO_min']:.2f} – {targets['MgO_max']:.2f}", "MgO"),
        ("CaO (%)", achieved["CaO"], f"{targets['CaO_min']:.2f} – {targets['CaO_max']:.2f}", "CaO"),
        ("B4", achieved["B4"], "1.8 – 2.2 (info)", "B4"),
    ]
    checks = quality_checks(achieved, targets)
    data = []
    for label, value, target, key in rows:
        data.append({
            "KPI": label,
            "Achieved": value,
            "Target": target,
            "Status": "OK" if key == "B4" else ("OK" if checks[key] else "OUT"),
        })
    return pd.DataFrame(data)


def redistribute_adjustment(baseline_blend, df, requested_values):
    """Preserve total burden by reducing other adjustable materials proportionally."""
    adjusted = baseline_blend.copy()

    adjustable = [
        m for m in baseline_blend
        if df.loc[m, "Group"] in ("Iron_ore", "Flux")
    ]

    if not adjustable:
        return adjusted

    # Apply requested values.
    for mat in adjustable:
        if mat in requested_values:
            adjusted[mat] = float(requested_values[mat])

    baseline_total_adjustable = sum(baseline_blend[m] for m in adjustable)
    requested_total_adjustable = sum(adjusted[m] for m in adjustable)
    delta = requested_total_adjustable - baseline_total_adjustable

    # The UI normally changes one material at a time.
    # Reduce/increase all other adjustable materials proportionally.
    if abs(delta) < 1e-9:
        return adjusted

    changed = max(
        adjustable,
        key=lambda m: abs(adjusted[m] - baseline_blend[m])
    )
    others = [m for m in adjustable if m != changed]
    total_others = sum(baseline_blend[m] for m in others)

    if total_others <= 0:
        return adjusted

    for mat in others:
        share = baseline_blend[mat] / total_others
        adjusted[mat] = max(0.0, adjusted[mat] - delta * share)

    return {m: round(v, 2) for m, v in adjusted.items()}


def what_if_analysis(df, targets=None):
    targets = targets or TARGETS
    base = solve_blend_with_compensation(df, 1000, targets)
    base_status, base_blend, base_cost = base[0], base[1], base[2]

    if base_status != "Optimal":
        return pd.DataFrame([{
            "Scenario": "Base Case",
            "Status": "Base case not optimal",
            "Cost Rs/t": None,
            "Cost Increase": None,
        }])

    results = []
    candidates = [
        m for m in df.index
        if df.loc[m, "Group"] in ("Iron_ore", "Flux", "Fuel")
        and df.loc[m, "Available_Tonnes"] > 0
    ]

    for mat in candidates:
        scenario_df = df.copy()
        scenario_df.loc[mat, "Available_Tonnes"] = 0

        result = solve_blend_with_compensation(
            scenario_df, 1000, targets, baseline_blend=base_blend
        )
        status, _, cost, _, _, fallback = result

        if status == "No_Production":
            label = "NO PRODUCTION"
            inc = None
        elif status == "Optimal" and not fallback:
            label = "Feasible"
            inc = (cost - base_cost) if cost is not None else None
        elif result[1] is not None:
            label = "Quality relaxed"
            inc = (cost - base_cost) if cost is not None else None
        else:
            label = "Hard infeasible"
            inc = None

        results.append({
            "Missing Material": mat,
            "Group": df.loc[mat, "Group"],
            "Status": label,
            "Cost Rs/t": cost,
            "Cost Increase": inc,
            "Increase %": (inc / base_cost * 100) if inc is not None and base_cost else None,
        })

    return pd.DataFrame(results)
