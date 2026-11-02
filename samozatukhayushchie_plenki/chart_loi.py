"""График: кислородный индекс (LOI) распространённых плёнок и материалов.
LOI — минимальная доля кислорода (%), при которой материал ещё горит. В воздухе 21 %.
Значения — из справочников и статей (см. список источников в zashchita.md)."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

data = [  # (название, LOI %)
    ("Полиэтилен (пакеты)", 17.4),
    ("Полипропилен", 17.4),
    ("Хлопок", 18),
    ("ПЭТ (бутылки, плёнки)", 21),
    ("Хлопок + ДНК сельди", 28),
    ("Каптон (полиимид)", 37),
    ("Жёсткий ПВХ", 45),
]
names, vals = zip(*data)
cols = ["#eb6834" if v <= 21 else "#2a9d6f" for v in vals]
fig, ax = plt.subplots(figsize=(9, 5), dpi=200)
bars = ax.barh(names, vals, color=cols)
ax.set_ylim(6.6, -1.1)
for b, v in zip(bars, vals):
    ax.text(v - 0.6, b.get_y() + b.get_height()/2, f"{v:g} %".replace(".", ","), va="center", ha="right",
            fontsize=10, color="white", weight="bold")
ax.axvline(21, color="#555", ls="--", lw=1.2)
ax.text(21.4, -0.75, "в воздухе 21 % кислорода", fontsize=9, color="#555")
ax.set_xlim(0, 52)
ax.set_xlabel("кислородный индекс (LOI), %")
ax.set_title("Сколько кислорода нужно материалу, чтобы гореть", loc="left", fontsize=13, weight="bold")
for s in ("top", "right"): ax.spines[s].set_visible(False)
fig.tight_layout(); fig.savefig("chart_loi.png")
print("ok")
