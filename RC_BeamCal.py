import numpy as np
import matplotlib.pyplot as plt
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import platform
import os
#藉由輸入設計彎矩、梁寬、梁深、配筋量等設計條件，計算梁應變，確認是否為拉力控制斷面。
# ==========================================
# 1. 輸入參數定義 (雙排筋邏輯)
# ==========================================
params = {
    "b": 120,  # 梁寬 (cm)
    "h": 150,  # 梁深 (cm)
    "num_bars": 8,  # 鋼筋總支數
    "bar_diameter": 3.22,  # 主筋直徑 (D32)
    "stirrup_dia": 1.59,  # 箍筋直徑 (D16)
    "gap": 2.5,  # 鋼筋淨間距
    "Mu": 121.16,  # 作用彎矩 (tf-m)
    "C": 4,  # 保護層厚度 (cm)
    "fc": 280,  # 混凝土強度 (kgf/cm2)
    "fy": 4200  # 鋼筋強度 (kgf/cm2)
}


# ==========================================
# 2. 力學計算邏輯 (dt 由保護層計算, c 改為 xb)
# ==========================================
def calculate_rc_mechanics(p):
    # (1) 計算鋼筋面積 As
    As = p['num_bars'] * (np.pi * (p['bar_diameter'] ** 2) / 4)

    # (2) 計算有效高度 dt (雙排筋公式)
    # dt = h - C - stirrup - bar_dia - gap/2
    dt = p['h'] - p['C'] - p['stirrup_dia'] - p['bar_diameter'] - (p['gap'] / 2)

    # (3) 壓力塊深度 a
    a = (As * p['fy']) / (0.85 * p['fc'] * p['b'])

    # (4) 中性軸深度 xb (原本為 c)
    beta1 = 0.85 if p['fc'] <= 280 else max(0.65, 0.85 - 0.05 * (p['fc'] - 280) / 70)
    xb = a / beta1

    # (5) 鋼筋應變 eps_t
    eps_cu = 0.003
    eps_t = eps_cu * (dt - xb) / xb

    # (6) 斷面判定
    eps_y = p['fy'] / 2040000
    if eps_t >= 0.005:
        section_type = "拉力控制斷面 (Tension-Controlled)"
        judgment_formula = f"εt ({eps_t:.4f}) >= 0.005"
        plot_logic = r"$\epsilon_t \geq 0.005$"  # Matplotlib 專用 TeX 語法
    elif eps_t <= eps_y:
        section_type = "壓力控制斷面 (Compression-Controlled)"
        judgment_formula = f"εt ({eps_t:.4f}) <= εy ({eps_y:.4f})"
        plot_logic = r"$\epsilon_t \leq \epsilon_y$"
    else:
        section_type = "過渡斷面 (Transition)"
        judgment_formula = f"εy ({eps_y:.4f}) < εt ({eps_t:.4f}) < 0.005"
        plot_logic = r"$\epsilon_y < \epsilon_t < 0.005$"

    return locals()


# ==========================================
# 3. 繪製應變圖 (垂直標註 a 與 dt)
# ==========================================
def generate_strain_plot(res):
    plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'Arial Unicode MS', 'SimHei']
    plt.rcParams['axes.unicode_minus'] = False

    p, h, dt, xb, a = res['p'], res['p']['h'], res['dt'], res['xb'], res['a']
    eps_t, section_type, plot_logic = res['eps_t'], res['section_type'], res['plot_logic']

    fig, ax = plt.subplots(figsize=(8, 10))
    ax.plot([0, 0], [0, h], 'k-', lw=2, alpha=0.3)
    ax.axhline(y=h, color='k', lw=2);
    ax.axhline(y=0, color='k', lw=2)

    # 應變分布線
    eps_cu_plot = -0.003
    ax.plot([eps_cu_plot, eps_t], [h, h - dt], 'r-o', lw=4, markersize=12)

    # --- 垂直標註區 ---
    # 標註 a
    ax.annotate('', xy=(-0.004, h - a), xytext=(-0.004, h),
                arrowprops=dict(arrowstyle='<->', color='darkorange', lw=2))
    ax.text(-0.0042, h - a / 2, f'a = {a:.2f} cm', color='darkorange', rotation=90, va='center', ha='right',
            fontweight='bold')

    # 標註 dt
    ax.annotate('', xy=(-0.006, h - dt), xytext=(-0.006, h),
                arrowprops=dict(arrowstyle='<->', color='blue', lw=2))
    ax.text(-0.0062, h - dt / 2, f'dt = {dt:.1f} cm', color='blue', rotation=90, va='center', ha='right',
            fontweight='bold')

    # 標示中性軸 xb
    ax.axhline(y=h - xb, color='gray', linestyle='--', lw=1.5)
    ax.text(eps_t / 2, h - xb + 2, f'中性軸 xb = {xb:.1f} cm', color='gray', fontsize=10)

    # 應變標註
    ax.text(eps_cu_plot, h + 3, f'εcu = 0.003', color='blue', fontweight='bold', ha='center')
    ax.text(eps_t, h - dt - 8, f'εt = {eps_t:.4f}', color='red', fontweight='bold', ha='center')

    # 判定顯示
    #ax.text(0, -20, f"判定：{section_type}", fontsize=14, color='darkgreen', fontweight='bold',
    #        ha='center', bbox=dict(facecolor='white', edgecolor='darkgreen', boxstyle='round,pad=0.5'))
    result_box_txt = f"【判定結果】\n{section_type}\n依據：{plot_logic}"
    ax.text(0.5, 1.05, result_box_txt, transform=ax.transAxes,
            fontsize=13, color='darkgreen', fontweight='bold',
            ha='center', va='bottom',
            bbox=dict(facecolor='#f0fff0', edgecolor='darkgreen', boxstyle='round,pad=0.8'))

    ax.set_title("梁斷面應變分析圖", fontsize=16, pad=85)
    ax.set_ylabel("深度方向 (cm)");
    ax.set_xlabel("應變值 (Strain)")
    ax.set_xlim(-0.008, max(0.008, eps_t * 1.5))
    ax.set_ylim(-15, h + 25)
    ax.grid(True, linestyle=':', alpha=0.5)

    plt.tight_layout()
    plt.savefig("strain_plot.png", dpi=300);
    plt.close()


# ==========================================
# 4. 產生 PDF 計算書 (含公式流程)
# ==========================================
class RCReport(FPDF):
    def __init__(self, font_path):
        super().__init__()
        self.font_name = "Chinese"
        if os.path.exists(font_path):
            self.add_font(self.font_name, "", font_path)
        else:
            self.font_name = "Arial"

    def header(self):
        self.set_font(self.font_name, size=18)
        self.cell(0, 10, "梁應變計算書", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
        self.ln(5)

    def chapter_body(self, res):
        p = res['p']
        self.set_font(self.font_name, size=14);
        self.cell(0, 10, "一、 設計參數與幾何", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_font(self.font_name, size=11)
        items = [f"梁寬 b = {p['b']} cm, 梁深 h = {p['h']} cm", f"保護層 C = {p['C']} cm",
                 f"鋼筋: {p['num_bars']} 支 D32 (D={p['bar_diameter']}cm), 箍筋 D16 (D={p['stirrup_dia']}cm)"]
        for i in items: self.cell(0, 7, i, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        self.ln(5);
        self.set_font(self.font_name, size=14);
        self.cell(0, 10, "二、 計算流程 (假設為雙排筋)", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_font(self.font_name, size=11)

        # 顯示公式與代入數值
        f = [
            f"1. 鋼筋總面積 (As) = {p['num_bars']} * (π * {p['bar_diameter']}^2 / 4) = {res['As']:.2f} cm2",
            f"2. 有效高度 (dt) = h - C - stirrup - bar_dia - gap/2",
            f"   dt = {p['h']} - {p['C']} - {p['stirrup_dia']} - {p['bar_diameter']} - {p['gap'] / 2} = {res['dt']:.1f} cm",
            f"3. 壓力塊深度 (a) = (As * fy) / (0.85 * fc' * b) = {res['a']:.2f} cm",
            f"4. 中性軸深度 (xb): xb = a / β1 (此處 β1 = {res['beta1']:.2f})",
            f"5. 拉力筋應變 (εt) = 0.003 * (dt - xb) / xb = {res['eps_t']:.4f}"
        ]
        for line in f: self.cell(0, 8, line, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        self.ln(5);
        self.cell(0, 10, f"判定結果：{res['section_type']} (依據：{res['judgment_formula']})",
                  new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.image("strain_plot.png", x=25, w=150)


# ==========================================
# 5. 執行
# ==========================================
if __name__ == "__main__":
    target_font = "C:/Windows/Fonts/msjh.ttc" if platform.system() == "Windows" else "/System/Library/Fonts/STHeiti Light.ttc"
    results = calculate_rc_mechanics(params)
    generate_strain_plot(results)
    pdf = RCReport(target_font);
    pdf.add_page();
    pdf.chapter_body(results)
    pdf.output("RC_DoubleLayer_Report.pdf")
    print(f"完成！dt={results['dt']:.1f}cm, xb={results['xb']:.2f}cm")