"""Мини-исследование: месяцы рождения игроков НХЛ сезона 2024/25.
Данные: составы НХЛ (api-web.nhle.com), выгрузка github.com/mforrest13/NHL-stats.
Запуск: python3 analysis.py  (нужны scipy и matplotlib)"""
import csv, collections
from scipy.stats import chisquare
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

rows = list({r["id"]: r for r in csv.DictReader(open("data/nhl_players_2024_25.csv"))}.values())
DAYS = [90.25, 91, 92, 92]  # дней в кварталах (с учётом високосных лет)

def quarters(sub):
    m = collections.Counter(int(r["birthDate"][5:7]) for r in sub)
    return [sum(m[i] for i in range(s, s + 3)) for s in (1, 4, 7, 10)], m

groups = {"Все игроки": rows}
for code, name in [("CAN", "Канада"), ("USA", "США"), ("SWE", "Швеция"), ("RUS", "Россия")]:
    groups[name] = [r for r in rows if r["birthCountry"] == code]

res = {}
for name, sub in groups.items():
    q, m = quarters(sub); n = sum(q)
    p = chisquare(q, [n * d / 365.25 for d in DAYS]).pvalue
    res[name] = [100 * x / n for x in q]
    print(f"{name:11s} n={n:3d}  I={q[0]:3d} II={q[1]:3d} III={q[2]:3d} IV={q[3]:3d}  "
          f"доли%={[round(x,1) for x in res[name]]}  I/IV={q[0]/q[3]:.2f}  p={p:.2g}")
    print("            по месяцам:", [m[i] for i in range(1, 13)])

fig, ax = plt.subplots(figsize=(9, 5), dpi=200)
labels = ["I квартал\nянв–мар", "II квартал\nапр–июн", "III квартал\nиюл–сен", "IV квартал\nокт–дек"]
x = range(4); w = 0.36
for off, (name, col) in zip((-w/2 - 0.01, w/2 + 0.01), (("Канада", "#2a78d6"), ("Россия", "#eb6834"))):
    bars = ax.bar([i + off for i in x], res[name], w, color=col, label=f"{name} (n={len(groups[name])})")
    for b, v in zip(bars, res[name]):
        ax.text(b.get_x() + b.get_width()/2, v + 0.6, f"{v:.0f}%", ha="center", fontsize=10, color="#333")
ax.axhline(25, color="#888", lw=1, ls="--", label="если бы рождались равномерно (25%)")
ax.set_xticks(list(x)); ax.set_xticklabels(labels)
ax.set_ylabel("доля игроков, %"); ax.set_ylim(0, 42)
ax.set_title("Когда родились игроки НХЛ сезона 2024/25", loc="left", fontsize=13, weight="bold")
for s in ("top", "right"): ax.spines[s].set_visible(False)
ax.legend(frameon=False, loc="upper center")
fig.tight_layout(); fig.savefig("chart_nhl_quarters.png")
print("chart saved")
