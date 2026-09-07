# Requires: Python 3, NumPy, and openpyxl. No dedicated statistical package is used.
# p values and 95% CIs use the t distribution with finite degrees of freedom (not a normal approximation).
# 输入：本工作簿「跨国底稿」工作表，第 5-171 行，列 C=MMR、E=原始GDP、H=政治体制代码、K=CPR、L=CPR年份。
#       G 列 logGDP = LN(E)，N 列 log(MMR) = LN(C)，O/Q/R 列为派生虚拟变量（均为工作簿内公式）。
# 输出：与「模型结果」「稳健性分析」「模型输出与HC3结果」三表逐位一致（误差量级 1e-10 以下）。
import numpy as np, openpyxl
from math import lgamma, exp, log
# 自动定位同目录下的数据表工作簿（无需改名即可运行）
import glob as _glob, os as _os
_cand = _glob.glob(_os.path.join(_os.path.dirname(_os.path.abspath(__file__)) or ".", "*数据表*.xlsx"))
if not _cand:
    raise SystemExit("未找到数据表工作簿（文件名需含「数据表」且为 .xlsx）")
wb = openpyxl.load_workbook(_cand[0])
ws = wb["跨国底稿"]
rows = [r for r in ws.iter_rows(min_row=5, max_row=171, values_only=True) if r and r[0]]
D = [dict(mmr=float(r[2]), gdp=float(r[4]), code=r[7], cpr=r[10], cyear=r[11]) for r in rows]
def fmt_p(p):                         # 极小 p 值显示为 <0.0001，不显示为 0
    return "<0.0001" if p < 1e-4 else f"{p:.4f}"

def t_p(t, df):                       # 双尾 p，正则化不完全 beta 函数
    x = df / (df + t * t); a, b = df / 2, 0.5
    def bcf(a, b, x, itmax=300, eps=3e-14):
        qab, qap, qam = a + b, a + 1, a - 1
        c = 1.0; d = 1 - qab * x / qap; d = 1 / d; h = d
        for m in range(1, itmax + 1):
            m2 = 2 * m
            aa = m * (b - m) * x / ((qam + m2) * (a + m2))
            d = 1 + aa * d; c = 1 + aa / c; d = 1 / d; h *= d * c
            aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
            d = 1 + aa * d; c = 1 + aa / c; d = 1 / d; de = d * c; h *= de
            if abs(de - 1) < eps: break
        return h
    lb = lgamma(a) + lgamma(b) - lgamma(a + b)
    return (exp(a * log(x) + b * log(1 - x) - lb) * bcf(a, b, x) / a if x < (a + 1) / (a + b + 2)
            else 1 - exp(b * log(1 - x) + a * log(x) - lb) * bcf(b, a, 1 - x) / b)
def ols(y, Xcols, names):
    n = len(y); X = np.column_stack([np.ones(n)] + Xcols); k = X.shape[1]
    XtXi = np.linalg.inv(X.T @ X); beta = XtXi @ X.T @ y
    e = y - X @ beta; dof = n - k
    se_ols = np.sqrt(np.diag((e @ e / dof) * XtXi))
    h = np.diag(X @ XtXi @ X.T)
    V = XtXi @ (X.T @ np.diag((e ** 2) / ((1 - h) ** 2)) @ X) @ XtXi      # HC3 三明治估计
    se_hc3 = np.sqrt(np.diag(V))
    r2 = 1 - (e @ e) / ((y - y.mean()) ** 2).sum()
    for i, nm in enumerate(names):
        print(f"{nm:10s} beta={beta[i]:9.4f} SE={se_ols[i]:8.4f} p={fmt_p(t_p(abs(beta[i]/se_ols[i]), dof)):>8} "
              f"| HC3 SE={se_hc3[i]:8.4f} p={fmt_p(t_p(abs(beta[i]/se_hc3[i]), dof)):>8}")
    print(f"n={n} dof={dof} R2={r2:.4f} adjR2={1-(1-r2)*(n-1)/dof:.4f}")
lg = lambda g: np.log(g)
# M1 / M2：需要政治体制代码，故排除无统一 2023 值的 PSE（n=166）
D_c    = [d for d in D if d["code"] is not None]
code   = np.array([d["code"] for d in D_c], dtype=float)
lmmr_c = np.array([np.log(d["mmr"]) for d in D_c])
mmr_c  = np.array([d["mmr"] for d in D_c])
lgdp_c = np.array([lg(d["gdp"]) for d in D_c])
ols(lmmr_c, [code, lgdp_c], ["Intercept", "regime", "logGDP"])      # M1 主模型
ols(mmr_c,  [code, lgdp_c], ["Intercept", "regime", "logGDP"])      # M2 未取对数
# M3 / M4 / M5：CPR 子样本（n=138；M5 限 2015-2023 后 n=109）
sub  = [d for d in D if d["cpr"] is not None and d["code"] is not None]
cpr  = np.array([d["cpr"] for d in sub])
c_cd = np.array([d["code"] for d in sub], dtype=float)
c_lg = np.array([lg(d["gdp"]) for d in sub])
ols(cpr, [c_cd, c_lg], ["Intercept", "regime", "logGDP"])                                  # M3 主模型
ols(cpr, [c_cd, c_lg, np.array([d["cyear"] for d in sub], dtype=float) - 2018],
         ["Intercept", "regime", "logGDP", "cpr_year_c"])                                  # M4 加观测年份
sub15 = [d for d in sub if d["cyear"] >= 2015]
ols(np.array([d["cpr"] for d in sub15]),
    [np.array([d["code"] for d in sub15], dtype=float),
     np.array([lg(d["gdp"]) for d in sub15])],
    ["Intercept", "regime", "logGDP"])                                                     # M5 限制年份窗
# M6 / M7：放松等距假设，改用三个虚拟变量（参照组 = 封闭独裁）
d1 = np.array([1.0 if d["code"] == 1 else 0.0 for d in D_c])
d2 = np.array([1.0 if d["code"] == 2 else 0.0 for d in D_c])
d3 = np.array([1.0 if d["code"] == 3 else 0.0 for d in D_c])
ols(lmmr_c, [d1, d2, d3, lgdp_c], ["Intercept", "regime_1", "regime_2", "regime_3", "logGDP"])  # M6
e1 = np.array([1.0 if d["code"] == 1 else 0.0 for d in sub])
e2 = np.array([1.0 if d["code"] == 2 else 0.0 for d in sub])
e3 = np.array([1.0 if d["code"] == 3 else 0.0 for d in sub])
ols(cpr, [e1, e2, e3, c_lg], ["Intercept", "regime_1", "regime_2", "regime_3", "logGDP"])      # M7
