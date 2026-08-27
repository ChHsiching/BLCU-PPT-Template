"""Generate the qat-lsq-repro experiment figures (素材图片层).

Usage:
    python examples/qat-lsq-repro/material/make_figures.py

Writes qat-pipeline.png / val-curves.png / bitwidth-bars.png /
stepscale-ablation.png next to this file under images/. The data arrays
here are the single source of the experiment numbers used by the material
doc, the deck and the 演讲稿 — regenerate nothing by hand.
"""
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

plt.rcParams["font.sans-serif"] = ["SimHei", "Microsoft YaHei"]
plt.rcParams["axes.unicode_minus"] = False

OUT = Path(__file__).resolve().parent / "images"

# ---- single source of the reported numbers (3-seed means) ----
EPOCHS = 100
FP16 = (78.4, 0.2)  # (mean, std), CIFAR-100 / ResNet-18, seeds 0/1/2

BITWIDTHS = ("W8A8", "W6A6", "W4A4")
PTQ = ((78.1, 0.1), (76.5, 0.2), (61.3, 0.4))
QAT_STE = ((78.2, 0.2), (76.9, 0.3), (73.1, 0.3))
QAT_LSQ = ((78.3, 0.2), (77.9, 0.2), (76.2, 0.3))

# step-size init ratio r (relative to s_max = max|x| / 2^(b-1)), W4A4, LSQ;
# the r=2.0 point matches QAT_LSQ[2] so ablation and main table cannot drift
STEP_R = ("0.5", "1.0", "2.0", "4.0")
STEP_ACC = (73.6, 75.4, 76.2, 74.8)
STEP_STD = (0.4, 0.3, 0.3, 0.4)

# validation top-1 trajectories, 3-seed means, 14 samples over 100 epochs;
# fitted endpoints match FP16 / QAT_LSQ[2] / QAT_STE[2] / PTQ[2]
VAL_CURVES = {
    "FP16": [6.1, 33.4, 52.8, 63.5, 69.1, 72.6, 74.9, 76.4, 77.3, 77.8,
             78.1, 78.3, 78.4, 78.4],
    "QAT-LSQ": [4.2, 26.5, 46.3, 57.9, 64.4, 68.6, 71.2, 72.9, 74.4, 75.3,
                75.8, 76.0, 76.2, 76.2],
    "QAT-STE": [3.1, 22.0, 40.2, 51.5, 58.2, 62.6, 65.6, 67.9, 69.8, 71.3,
                72.2, 72.7, 73.0, 73.1],
}
PTQ_W4A4_REF = PTQ[2][0]  # horizontal reference, no training phase
SAMPLE_EVERY = EPOCHS / (len(VAL_CURVES["FP16"]) - 1)

METHOD_COLORS = {"PTQ": "#7f7f7f", "QAT-STE": "#1f77b4", "QAT-LSQ": "#d62728"}
ACCENT = "#548235"


def val_curves(path):
    fig, ax = plt.subplots(figsize=(11.4, 3.6))
    line_styles = {"FP16": "#7f7f7f", "QAT-LSQ": "#d62728", "QAT-STE": "#1f77b4"}
    for name, values in VAL_CURVES.items():
        epochs = [i * SAMPLE_EVERY for i in range(len(values))]
        ax.plot(epochs, values, marker="o", markersize=3.5, linewidth=1.6,
                color=line_styles[name], label=name)
    ref = [PTQ_W4A4_REF] * 2
    ax.plot([0, EPOCHS], ref, linewidth=1.6, linestyle="--", color="#222222",
            label="PTQ W4A4（换算子后，61.3%）")
    ax.set_xlabel("训练 epoch（3 随机种子均值）")
    ax.set_ylabel("验证 top-1（%）")
    ax.set_title("CIFAR-100 / ResNet-18 W4A4：验证精度曲线")
    ax.set_xlim(0, EPOCHS)
    ax.set_ylim(0, 88)
    ax.legend(frameon=False, ncol=2, loc="lower right")
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)


def bitwidth_bars(path):
    series = (("PTQ", PTQ), ("QAT-STE", QAT_STE), ("QAT-LSQ", QAT_LSQ))
    width = 0.24
    xs = range(len(BITWIDTHS))
    # staggered label heights: W8A8/W6A6 groups sit within ~2 points, flat
    # offsets would collide across neighboring bars
    label_offset = (0.9, 0.15, 0.9)
    fig, ax = plt.subplots(figsize=(11.4, 3.6))
    for i, (name, data) in enumerate(series):
        means = [m for m, _ in data]
        stds = [s for _, s in data]
        bars = ax.bar([x + (i - 1) * width for x in xs], means, yerr=stds,
                      capsize=4, width=width, color=METHOD_COLORS[name],
                      label=name)
        for bar, m, s in zip(bars, means, stds):
            # white backing box: near-baseline values (78.1–78.3 vs the
            # 78.4 dashed line) would otherwise have the dashes strike
            # through the digits
            ax.text(bar.get_x() + bar.get_width() / 2, m + s + label_offset[i],
                    f"{m:.1f}", ha="center", fontsize=9,
                    bbox=dict(facecolor="white", edgecolor="none", pad=1.2))
    ax.axhline(FP16[0], linewidth=1.4, linestyle="--", color="#222222")
    ax.text(2.7, FP16[0] + 0.12, f"FP16 基线 {FP16[0]:.1f}%", fontsize=10,
            ha="right", color="#222222")
    ax.set_xticks(list(xs))
    ax.set_xticklabels(BITWIDTHS)
    ax.set_ylabel("最终测试精度（%）")
    ax.set_title("比特宽度扫描：3 随机种子均值 ± std")
    # bars span down to the axis floor, so every lower corner is bar body:
    # the legend goes to the top band (kept empty via ylim headroom)
    ax.set_ylim(58, 82.5)
    ax.legend(frameon=False, ncol=3, loc="upper center")
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)


def stepscale_ablation(path):
    """LSQ step-size init sweep; the best ratio is accented in brand green."""
    best = STEP_ACC.index(max(STEP_ACC))
    colors = ["#b8c9a9"] * len(STEP_ACC)
    colors[best] = ACCENT
    fig, ax = plt.subplots(figsize=(11.4, 3.6))
    bars = ax.bar(STEP_R, STEP_ACC, yerr=STEP_STD, capsize=5, color=colors,
                  width=0.55)
    for bar, acc, err in zip(bars, STEP_ACC, STEP_STD):
        ax.text(bar.get_x() + bar.get_width() / 2, acc + err + 0.08,
                f"{acc:.1f}%", ha="center", fontsize=11)
    ax.set_xlabel("步长初始化倍率 r（相对 s_max = max|x| / 2^(b-1)）")
    ax.set_ylabel("最终测试精度（%）")
    ax.set_title("步长初始化消融（W4A4，QAT-LSQ，3 随机种子均值）")
    ax.set_ylim(72, 78)
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)


def pipeline(path):
    """Reproduction pipeline schematic (vertical flow, near-square aspect to
    match the text-image image_primary slot 5.8 x 5.19 in).

    The bottom margin stays blank on purpose: the deck overlays a caption
    scrim on the fitted picture's bottom edge, and that band must cover only
    whitespace, never a box or arrow."""
    fig = plt.figure(figsize=(5.7, 5.08))
    ax = fig.add_axes((0, 0, 1, 1))  # axis fills the canvas: the fitted
    # picture then fills the image_primary slot (5.8 x 5.19 in) edge to edge
    ax.set_xlim(0, 5.7)
    ax.set_ylim(0.15, 5.15)
    ax.axis("off")

    def box(x, y, w, h, label, edge="#333", face="#f5f5f5", color="#333",
            fontsize=10):
        ax.add_patch(FancyBboxPatch(
            (x, y), w, h, boxstyle="round,pad=0.08", linewidth=1.4,
            edgecolor=edge, facecolor=face))
        ax.text(x + w / 2, y + h / 2, label, ha="center", va="center",
                fontsize=fontsize, color=color)

    def arrow(p0, p1, color="#333", style="-|>", dashed=False):
        ax.add_patch(FancyArrowPatch(
            p0, p1, arrowstyle=style, mutation_scale=16, linewidth=1.4,
            color=color, linestyle="--" if dashed else "-"))

    # main column (left)
    box(0.55, 4.25, 2.75, 0.72, "CIFAR-100\n批次 128")
    arrow((1.925, 4.25), (1.925, 3.85))
    box(0.55, 2.95, 2.75, 0.9, "Conv-BN-ReLU\n残差块 ×4")
    arrow((1.925, 2.95), (1.925, 2.55))
    box(0.55, 1.8, 2.75, 0.75, "全连接 → 100 类")
    arrow((1.925, 1.8), (1.925, 1.4))
    box(0.55, 0.65, 2.75, 0.75, "交叉熵损失")

    # fake-quant insert (right, red): weights and activations per block
    box(3.75, 2.95, 1.85, 0.9, "权重/激活\n伪量化", edge="#d62728",
        face="#fff5f5", color="#d62728", fontsize=9.5)
    arrow((3.30, 3.25), (3.75, 3.25), style="<|-|>", color="#d62728")

    # LSQ step-size update (right, green): gradient from loss to step sizes
    box(3.75, 1.0, 1.85, 1.0, "步长 s 梯度\n独立学习率", edge=ACCENT,
        face="#f3f8ee", color=ACCENT, fontsize=9.5)
    arrow((3.30, 1.16), (3.75, 1.3), color=ACCENT)
    arrow((4.675, 2.0), (4.675, 2.95), color=ACCENT, dashed=True)
    ax.text(4.82, 2.5, "联合更新", fontsize=9, color=ACCENT, ha="left")

    # training loop (left, dashed)
    arrow((0.55, 1.025), (0.2, 1.025), style="-", dashed=True, color="#999")
    arrow((0.2, 1.025), (0.2, 4.61), style="-", dashed=True, color="#999")
    arrow((0.2, 4.61), (0.55, 4.61), dashed=True, color="#999")
    ax.text(0.09, 2.8, "训练循环 100 epoch", fontsize=9, color="#666",
            rotation=90, va="center", ha="center")
    fig.savefig(path, dpi=200)
    plt.close(fig)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    pipeline(OUT / "qat-pipeline.png")
    val_curves(OUT / "val-curves.png")
    bitwidth_bars(OUT / "bitwidth-bars.png")
    stepscale_ablation(OUT / "stepscale-ablation.png")
    print(f"wrote 4 figure(s) -> {OUT}")


if __name__ == "__main__":
    main()
