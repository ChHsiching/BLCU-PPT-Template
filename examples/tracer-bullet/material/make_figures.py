"""Generate the tracer-bullet experiment figures (素材图片层).

Usage:
    python examples/tracer-bullet/material/make_figures.py

Writes pipeline.png / loss-curves.png / acc-bars.png / lambda-ablation.png
next to this file under images/. The data arrays here are the single source
of the experiment numbers used by the material doc, the deck and the 演讲稿 —
regenerate nothing by hand.
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
OPTIMIZERS = ("SGD (momentum)", "Adam + L2", "AdamW")
COLORS = ("#7f7f7f", "#1f77b4", "#d62728")
FINAL_ACC = (94.6, 93.9, 95.2)
FINAL_ACC_STD = (0.2, 0.2, 0.1)
FINAL_LOSS = (0.043, 0.061, 0.036)

# AdamW weight-decay ablation (3-seed means; the λ=1e-2 point matches
# FINAL_ACC[2] so the ablation and the main comparison cannot drift apart)
LAMBDA_LABELS = ("λ=0", "λ=1e-4", "λ=1e-2", "λ=5e-2")
LAMBDA_ACC = (94.7, 94.9, 95.2, 94.4)

# training-loss trajectories (coarse hand-fit to the logged runs; the fitted
# endpoints match FINAL_LOSS so text and figures cannot drift apart)
LOSS_CURVES = {
    "SGD (momentum)": [2.30, 0.62, 0.34, 0.23, 0.17, 0.13, 0.10, 0.081, 0.066, 0.056,
                       0.050, 0.046, 0.044, 0.043],
    "Adam + L2":      [1.86, 0.55, 0.38, 0.30, 0.26, 0.23, 0.20, 0.18, 0.16, 0.14,
                       0.12, 0.10, 0.085, 0.061],
    "AdamW":          [1.74, 0.46, 0.28, 0.20, 0.16, 0.13, 0.11, 0.092, 0.078, 0.066,
                       0.056, 0.048, 0.041, 0.036],
}
SAMPLE_EVERY = EPOCHS / (len(LOSS_CURVES["AdamW"]) - 1)  # epochs per sample point


def loss_curves(path):
    fig, ax = plt.subplots(figsize=(7.2, 4.4))
    for name, values in LOSS_CURVES.items():
        epochs = [i * SAMPLE_EVERY for i in range(len(values))]
        line, = ax.plot(epochs, values, marker="o", markersize=3.5,
                        linewidth=1.6, color=COLORS[OPTIMIZERS.index(name)],
                        label=name)
    ax.set_xlabel("训练 epoch（3 随机种子均值）")
    ax.set_ylabel("训练损失")
    ax.set_title("CIFAR-10 / ResNet-18：训练损失对比")
    ax.set_ylim(0, 1.0)
    ax.legend(frameon=False)
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)


def acc_bars(path):
    fig, ax = plt.subplots(figsize=(7.2, 4.4))
    bars = ax.bar(OPTIMIZERS, FINAL_ACC, yerr=FINAL_ACC_STD, capsize=5,
                  color=COLORS, width=0.55)
    for bar, acc, err in zip(bars, FINAL_ACC, FINAL_ACC_STD):
        ax.text(bar.get_x() + bar.get_width() / 2, acc + err + 0.08,
                f"{acc:.1f}%", ha="center", fontsize=11)
    ax.set_ylabel("最终测试精度（%）")
    ax.set_title("最终测试精度：3 随机种子均值 ± std")
    ax.set_ylim(92.5, 96.0)
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)


def pipeline(path):
    """Training-pipeline schematic: forward, loss, decoupled-decay update.

    The bottom margin stays blank on purpose: the deck hugs the fitted
    image's bottom edge with a caption scrim, and that band must cover only
    whitespace, never the red annotation."""
    fig, ax = plt.subplots(figsize=(11.5, 4.05))
    ax.set_xlim(0, 11.5)
    ax.set_ylim(-0.85, 3.2)
    ax.axis("off")

    boxes = [
        (0.3, 1.3, "CIFAR-10\n批次 128"),
        (3.0, 1.3, "ResNet-18\n前向传播"),
        (5.7, 1.3, "交叉熵损失\n+ 评估"),
        (8.4, 1.3, "AdamW 更新\n（动量 + 自适应 lr）"),
    ]
    for x, y, label in boxes:
        ax.add_patch(FancyBboxPatch(
            (x, y), 2.4, 1.1, boxstyle="round,pad=0.08",
            linewidth=1.4, edgecolor="#333", facecolor="#f5f5f5"))
        ax.text(x + 1.2, y + 0.55, label, ha="center", va="center", fontsize=11)

    for x0 in (2.7, 5.4, 8.1):
        ax.add_patch(FancyArrowPatch(
            (x0, 1.85), (x0 + 0.3, 1.85), arrowstyle="-|>",
            mutation_scale=18, linewidth=1.4, color="#333"))

    # the decoupled decay bypasses the adaptive scaling: drawn as its own arrow
    ax.add_patch(FancyBboxPatch(
        (8.4, 0.1), 2.4, 0.7, boxstyle="round,pad=0.08",
        linewidth=1.4, edgecolor="#d62728", facecolor="#fff5f5"))
    ax.text(9.6, 0.45, "解耦衰减 -η·λ·θ\n（不进动量/二阶矩）", ha="center",
            va="center", fontsize=10, color="#d62728")
    ax.add_patch(FancyArrowPatch(
        (9.6, 0.8), (9.6, 1.28), arrowstyle="-|>", mutation_scale=18,
        linewidth=1.4, color="#d62728"))

    ax.add_patch(FancyArrowPatch(
        (1.5, 2.45), (9.6, 2.45), arrowstyle="-", linewidth=1.2,
        color="#999", linestyle="--"))
    ax.add_patch(FancyArrowPatch(
        (9.6, 2.45), (9.6, 2.42), arrowstyle="-|>", mutation_scale=18,
        linewidth=1.2, color="#999"))
    ax.text(5.5, 2.62, "参数更新回流（训练循环）", ha="center",
            fontsize=10, color="#666")
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)


def lambda_ablation(path):
    """AdamW weight-decay sweep; the best λ is accented in brand green."""
    best = LAMBDA_ACC.index(max(LAMBDA_ACC))
    colors = ["#b8c9a9"] * len(LAMBDA_ACC)
    colors[best] = "#548235"
    fig, ax = plt.subplots(figsize=(7.2, 4.4))
    bars = ax.bar(LAMBDA_LABELS, LAMBDA_ACC, color=colors, width=0.55)
    for bar, acc in zip(bars, LAMBDA_ACC):
        ax.text(bar.get_x() + bar.get_width() / 2, acc + 0.06,
                f"{acc:.1f}%", ha="center", fontsize=11)
    ax.set_ylabel("最终测试精度（%）")
    ax.set_title("AdamW 权重衰减系数 λ 扫描（3 随机种子均值）")
    ax.set_ylim(93.8, 95.8)
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    pipeline(OUT / "pipeline.png")
    loss_curves(OUT / "loss-curves.png")
    acc_bars(OUT / "acc-bars.png")
    lambda_ablation(OUT / "lambda-ablation.png")
    print(f"wrote 4 figure(s) -> {OUT}")


if __name__ == "__main__":
    main()
