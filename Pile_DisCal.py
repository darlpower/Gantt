import numpy as np
import math

# ==============================================================================
# 1. 專案參數設定 (Project Input)
# ==============================================================================

# --- A. 外力條件 (External Loads) ---
H0 = 1520  # kN (水平力)
V0 = 4200  # kN (垂直力)
M0 = 2310  # kN.m (彎矩)

# --- B. 樁剛度參數 (Pile Stiffness) ---
Kv = 131567  # kN/m
K1 = 16381  # kN/m
K2 = 9922  # kN/rad
K3 = 9922  # kN/rad
K4 = 14020  # kN.m/rad

# --- C. 樁幾何配置 (Pile Geometry) ---
# ★關鍵修正★：確保這裡產生的是包含 14 個數值的陣列
# 前排(x=-1.25) 7支, 後排(x=+1.25) 7支
x_coords = np.array([-1.25] * 7 + [1.25] * 7)
theta = 0  # 垂直樁 (rad)
D = 0.239  # m (樁徑)

# --- D. 地層與樁端參數 (Soil & Tip Parameters) ---
soil_layers = [
    {"id": 1, "type": "sand", "Li": 6.5, "N": 10, "desc": "崖錐層(砂質)"},
    {"id": 2, "type": "sand", "Li": 11.2, "N": 23, "desc": "砂質土"},
    {"id": 3, "type": "sand", "Li": 1.3, "N": 50, "desc": "砂質土"}
]

# 樁端評估設定
tip_condition = {
    "standard": "Japan_2017",  # 選項: 'Japan_2017' 或 'Taiwan_Building'
    "method": "Bored",  # 工法
    "soil_type": "sand",  # 樁底土質
    "N_val": 50  # 樁底 N 值
}


# ==============================================================================
# 2. 核心計算函數庫 (Functions)
# ==============================================================================

def get_geometry(d):
    """計算幾何參數"""
    return math.pi * d, (math.pi * d ** 2) / 4


def get_qd_value(standard, method, soil_type, N_val):
    """依據規範計算樁端極限支承度 qd"""
    qd = 0
    note = ""

    if standard == 'Japan_2017':
        if method == 'Bored':
            if soil_type in ['sand', 'clay']:
                qd = min(110 * N_val, 3300)
                note = "110N (Max 3300)"
            elif soil_type == 'gravel':
                qd = min(160 * N_val, 8000)
                note = "160N (Max 8000)"

    elif standard == 'Taiwan_Building':
        # 1 tf/m2 = 9.81 kN/m2
        if method == 'Bored' and soil_type == 'sand':
            eff_N = min(N_val, 50)
            qd = (7.5 * eff_N) * 9.81
            note = "7.5N tf/m2"

    if qd == 0: qd, note = 3000, "Default Fixed"
    return qd, note


def solve_structure(K_vals, x_arr, loads, theta_rad):
    """
    求解變位矩陣 (修正版)
    修正說明：強制將剛度項擴展為陣列後再加總，確保計算到 n 支樁的總和。
    """
    kv, k1, k2, k3, k4 = K_vals
    sin_t, cos_t = np.sin(theta_rad), np.cos(theta_rad)
    n_piles = len(x_arr)  # 應該是 14

    # --- 建立係數 (Aij) ---
    # 使用 np.sum 時，若內容是常數，必須乘上樁數，或使用 np.full_like 確保維度正確

    # Axx: 水平剛度總和
    term_xx = k1 * cos_t ** 2 + kv * sin_t ** 2
    Axx = term_xx * n_piles

    # Axy: 耦合項總和
    term_xy = (kv - k1) * sin_t * cos_t
    Axy = term_xy * n_piles

    # Ax_alpha: (因為涉及 x_arr，直接 sum 即可)
    Ax_a = np.sum((kv - k1) * x_arr * sin_t * cos_t - k2 * cos_t)

    # Ayy: 垂直剛度總和
    term_yy = kv * cos_t ** 2 + k1 * sin_t ** 2
    Ayy = term_yy * n_piles

    # Ay_alpha:
    Ay_a = np.sum((kv * cos_t ** 2 + k1 * sin_t ** 2) * x_arr + k2 * sin_t)

    # A_alpha_alpha: 旋轉剛度總和
    Aa_a = np.sum((kv * cos_t ** 2 + k1 * sin_t ** 2) * x_arr ** 2 + (k2 + k3) * x_arr * sin_t + k4)

    # --- 組合矩陣並求解 ---
    K_mat = np.array([
        [Axx, Axy, Ax_a],
        [Axy, Ayy, Ay_a],
        [Ax_a, Ay_a, Aa_a]
    ])
    F_vec = np.array(loads)

    try:
        sol = np.linalg.solve(K_mat, F_vec)
        return sol, K_mat
    except np.linalg.LinAlgError:
        return [0, 0, 0], K_mat


# ==============================================================================
# 3. 主程式執行 (Main Execution)
# ==============================================================================

print("=" * 65)
print(f"{'結構與地工整合計算報告 (修正版)':^55}")
print("=" * 65)

# --- Step 1: 幾何與參數準備 ---
Ug, Ag = get_geometry(D)
print(f"【幾何資訊】 樁數 n={len(x_coords)}, D={D} m, Ug={Ug:.4f} m, Ag={Ag:.4f} m2")

# --- Step 2: 結構變位求解 ---
stiffness_inputs = (Kv, K1, K2, K3, K4)
loads_inputs = [H0, V0, M0]
(delta_x, delta_y, alpha), K_matrix = solve_structure(stiffness_inputs, x_coords, loads_inputs, theta)

print("-" * 65)
print("【結構變位求解結果】 (目標: dx~7.29mm, dy~2.28mm)")
print(f"水平變位 (delta_x) : {delta_x * 1000:.4f} mm")
print(f"垂直變位 (delta_y) : {delta_y * 1000:.4f} mm")
print(f"旋轉角   (alpha)   : {alpha:.6f} rad")

# --- Step 3: 單樁受力計算 ---
print("-" * 65)
print(f"{'樁位置(x)':^10} | {'PNi (軸力)':^12} | {'PHi (剪力)':^12} | {'Mti (彎矩)':^12}")
print("-" * 65)

PN_list = []
sin_t, cos_t = np.sin(theta), np.cos(theta)
printed_indices = [0, 7]  # 只印出前後排代表

for i, x in enumerate(x_coords):
    # 局部變位
    d_xi = delta_x * cos_t - (delta_y + alpha * x) * sin_t
    d_yi = delta_x * sin_t + (delta_y + alpha * x) * cos_t

    # 內力計算
    PNi = Kv * d_yi
    PHi = K1 * d_xi - K2 * alpha
    Mti = -K3 * d_xi + K4 * alpha

    PN_list.append(PNi)

    if i in printed_indices:
        pos_name = "前排" if x < 0 else "後排"
        print(f"{pos_name}({x:5.2f}) | {PNi:12.2f} | {PHi:12.2f} | {Mti:12.2f}")

PN_max = max(PN_list)
print("-" * 65)
print(f"** 最大軸壓力 P_max = {PN_max:.2f} kN (作為檢核依據)")

# --- Step 4: 極限承載力 Ru 計算 ---
print("-" * 65)
print("【極限承載力估算 Ru = Rf + Rp】")
Rf = 0

# 4.1 周面摩擦力
for layer in soil_layers:
    raw_tau = 5 * layer["N"] if layer["type"] == "sand" else 10 * layer["N"]
    limit = 200 if layer["type"] == "sand" else 150
    tau = min(raw_tau, limit)

    f_i = Ug * layer["Li"] * tau
    Rf += f_i

# 4.2 樁端支承力
qd, qd_rule = get_qd_value(tip_condition['standard'], tip_condition['method'],
                           tip_condition['soil_type'], tip_condition['N_val'])
Rp = qd * Ag
Ru = Rf + Rp

print(f"總周面摩擦力 (Rf) : {Rf:.2f} kN")
print(f"樁端支承力   (Rp) : {Rp:.2f} kN (Rule: {qd_rule})")
print(f"單樁極限承載力(Ru) : {Ru:.2f} kN")

# --- Step 5: 安全性判定 ---
print("=" * 65)
print(f"{'最終安全性檢核 (Safety Check)':^55}")
print("=" * 65)
fs = Ru / PN_max
print(f"需求軸力 (Demand P_max) : {PN_max:.2f} kN")
print(f"供應能力 (Supply Ru)    : {Ru:.2f} kN")
print(f"安全係數 (F.S.)         : {fs:.2f}")

if fs >= 3.0:
    print(">> 判定結果: 【OK】 安全係數 > 3.0 (常時)")
elif fs >= 2.0:
    print(">> 判定結果: 【注意】 2.0 < 安全係數 < 3.0 (僅滿足地震時需求)")
else:
    print(">> 判定結果: 【NG】 安全係數不足")
print("=" * 65)