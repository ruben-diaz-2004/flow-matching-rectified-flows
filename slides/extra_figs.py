"""Figuras auxiliares para las diapositivas: ecuaciones (mathtext), schedules y portada.

Uso:  ..\\.venv\\Scripts\\python extra_figs.py      (desde la carpeta slides/)
"""
import math, os
import numpy as np
import torch, torch.nn as nn
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

OUT = "assets"
os.makedirs(f"{OUT}/eq", exist_ok=True)
plt.rcParams["mathtext.fontset"] = "cm"

INK, TEAL, CRIMSON, GOLD, NAVY = "#1F2933", "#00798C", "#D1495B", "#EDAE49", "#14213D"

# ---------------------------------------------------------------- ecuaciones
EQS = {
    "ode": r"\frac{d\,x_t}{dt} = v_\theta(x_t,\,t), \qquad x_0 \sim p_0",
    "pushforward": r"p_t = [\phi_t]_{\#}\; p_0",
    "continuity": r"\partial_t\, p_t(x) + \nabla\!\cdot\!\left(p_t(x)\, v_t(x)\right) = 0",
    "logp": r"\frac{d}{dt}\,\log p_t(x_t) = -\,\nabla\!\cdot v_\theta(x_t,t)",
    "lfm": r"\mathcal{L}_{FM}(\theta) = \mathbb{E}_{t,\;x\sim p_t}\,\|\,v_\theta(x,t) - u_t(x)\,\|^2",
    "condpath": r"x_t = (1-t)\,x_0 + t\,x_1",
    "condvel": r"u_t(x_t \mid x_1) = \frac{d x_t}{dt} = x_1 - x_0",
    "marginal_p": r"p_t(x) = \int p_t(x\mid x_1)\, q(x_1)\, dx_1",
    "marginal_u": r"u_t(x) = \mathbb{E}\,[\,x_1 - x_0 \mid x_t = x\,]",
    "lcfm": r"\mathcal{L}_{CFM}(\theta) = \mathbb{E}_{t,\,x_0,\,x_1}\,\|\,v_\theta(x_t,t) - (x_1 - x_0)\,\|^2",
    "theorem": r"\nabla_\theta\, \mathcal{L}_{FM}(\theta) = \nabla_\theta\, \mathcal{L}_{CFM}(\theta)",
    "proof1": r"\|v_\theta - u\|^2 = \|v_\theta\|^2 - 2\,\langle v_\theta,\, u\rangle + \|u\|^2",
    "proof2": r"\mathbb{E}_{p_t(x)}\langle v_\theta(x), u_t(x)\rangle = \mathbb{E}_{q(x_1)\,p_t(x|x_1)}\langle v_\theta(x), u_t(x|x_1)\rangle",
    "gausspath": r"x_t = \alpha_t\, x_1 + \sigma_t\, x_0, \qquad u_t = \dot\alpha_t\, x_1 + \dot\sigma_t\, x_0",
    "x0hat": r"\hat x_0 = x_t - t\, v_\theta",
    "x1hat": r"\hat x_1 = x_t + (1-t)\, v_\theta",
    "score": r"\nabla \log p_t(x_t) = -\frac{x_t - t\, v_\theta}{1-t}",
    "sde": r"dx = \left[\,v_\theta + \frac{g_t^2}{2}\,\nabla\log p_t\right]\,dt + g_t\, dW_t",
    "euler": r"x_{t+h} = x_t + h\; v_\theta(x_t,\, t)",
    "straight": r"S(Z) = \int_0^1 \mathbb{E}\,\|\,(Z_1 - Z_0) - \dot Z_t\,\|^2\; dt",
    "cost": r"\mathbb{E}\,[c(Z_1 - Z_0)] \;\leq\; \mathbb{E}\,[c(X_1 - X_0)] \quad \forall\, c\ \mathrm{convexa}",
    "bound": r"\min_{k\leq K}\, S(Z^k) \;\leq\; \frac{\mathbb{E}\,\|X_1 - X_0\|^2}{K}",
    "reflow": r"(Z_0,\, Z_1) = \left(X_0,\ \mathrm{ODE}_{v}(X_0)\right)",
    "otcfm": r"\pi^\star = \mathrm{arg\,min}_{\pi}\ \sum_{i,j}\pi_{ij}\,\|x_0^{(i)} - x_1^{(j)}\|^2",
    "lognorm": r"\pi_{ln}(t;\,m,s) = \frac{1}{s\sqrt{2\pi}}\,\frac{1}{t(1-t)}\,\exp\!\left(-\frac{(\mathrm{logit}(t)-m)^2}{2s^2}\right)",
    "meanflow": r"u(x_t, r, t) = \frac{1}{t-r}\int_r^t v(x_\tau, \tau)\, d\tau",
}


def render_eq(name, tex, color=INK, size=28):
    fig = plt.figure(figsize=(0.01, 0.01))
    fig.text(0, 0, f"${tex}$", fontsize=size, color=color)
    fig.savefig(f"{OUT}/eq/{name}.png", dpi=300, transparent=True, bbox_inches="tight", pad_inches=0.04)
    plt.close(fig)


for k, v in EQS.items():
    render_eq(k, v)
    render_eq(k + "_w", v, color="white")
print(len(EQS), "ecuaciones")

# ---------------------------------------------------------------- schedules
t = np.linspace(0, 1, 200)
fig, axs = plt.subplots(1, 2, figsize=(10, 4))
ax = axs[0]
ax.plot(t, t, c=TEAL, lw=2.5, label=r"FM-OT / RF: $\alpha_t = t$")
ax.plot(t, 1 - t, c=TEAL, lw=2.5, ls="--", label=r"$\sigma_t = 1-t$")
ax.plot(t, np.sin(np.pi * t / 2), c=CRIMSON, lw=2.5, label=r"VP (coseno): $\alpha_t=\sin\frac{\pi t}{2}$")
ax.plot(t, np.cos(np.pi * t / 2), c=CRIMSON, lw=2.5, ls="--", label=r"$\sigma_t=\cos\frac{\pi t}{2}$")
ax.set_xlabel("t   (0 = ruido, 1 = datos)"); ax.set_title("Schedules  $(\\alpha_t, \\sigma_t)$")
ax.legend(fontsize=11, frameon=False); ax.spines[["top", "right"]].set_visible(False)
ax = axs[1]
x0, x1 = np.array([-1.0, 1.4]), np.array([1.6, -0.2])
ot = (1 - t)[:, None] * x0 + t[:, None] * x1
vp = np.cos(np.pi * t / 2)[:, None] * x0 + np.sin(np.pi * t / 2)[:, None] * x1
ax.plot(*ot.T, c=TEAL, lw=3, label="camino lineal (FM-OT / RF)")
ax.plot(*vp.T, c=CRIMSON, lw=3, label="camino VP (difusión)")
ax.scatter(*x0, s=120, c="#9aa5b1", zorder=3); ax.annotate("$x_0$ (ruido)", x0, xytext=(-10, 12), textcoords="offset points", fontsize=12)
ax.scatter(*x1, s=120, c=CRIMSON, zorder=3); ax.annotate("$x_1$ (dato)", x1, xytext=(-20, -22), textcoords="offset points", fontsize=12)
ax.set_aspect("equal"); ax.set_xticks([]); ax.set_yticks([]); ax.set_title("Camino condicional de un par $(x_0, x_1)$")
ax.legend(fontsize=11, frameon=False, loc="upper right"); ax.set_xlim(-1.6, 2.3); ax.set_ylim(-0.8, 2.0)
plt.tight_layout(); plt.savefig(f"{OUT}/schedules.png", dpi=200, bbox_inches="tight"); plt.close()

# ---------------------------------------------------------------- portada
torch.manual_seed(3)


def sample_moons(n, noise=0.05):
    a = torch.rand(n) * math.pi
    up = torch.rand(n) < 0.5
    x = torch.where(up, torch.cos(a), 1 - torch.cos(a)); y = torch.where(up, torch.sin(a), 0.5 - torch.sin(a))
    return (torch.stack([x, y], 1) + noise * torch.randn(n, 2) - torch.tensor([0.5, 0.25])) * 2.0


net = nn.Sequential(nn.Linear(3, 256), nn.SiLU(), nn.Linear(256, 256), nn.SiLU(), nn.Linear(256, 256), nn.SiLU(), nn.Linear(256, 2))
opt = torch.optim.Adam(net.parameters(), 1e-3)
for it in range(5000):
    x1 = sample_moons(512); x0 = torch.randn_like(x1); tt = torch.rand(512, 1)
    loss = ((net(torch.cat([(1 - tt) * x0 + tt * x1, tt], 1)) - (x1 - x0)) ** 2).mean()
    opt.zero_grad(); loss.backward(); opt.step()

with torch.no_grad():
    x = torch.randn(700, 2) * 1.0; traj = [x]
    for i in range(100):
        x = x + 0.01 * net(torch.cat([x, torch.full((x.shape[0], 1), i / 100)], 1)); traj.append(x)
traj = torch.stack(traj).numpy()

fig = plt.figure(figsize=(8, 7.5)); ax = fig.add_axes([0, 0, 1, 1]); fig.patch.set_facecolor(NAVY); ax.set_facecolor(NAVY)
from matplotlib.collections import LineCollection
for i in range(traj.shape[1]):
    pts = traj[:, i, :]
    segs = np.stack([pts[:-1], pts[1:]], 1)
    lc = LineCollection(segs, cmap="viridis", array=np.linspace(0, 1, len(segs)), linewidths=0.7, alpha=0.55)
    lc.set_cmap(matplotlib.colors.LinearSegmentedColormap.from_list("fm", ["#5c6b80", TEAL, CRIMSON]))
    ax.add_collection(lc)
ax.scatter(*traj[-1].T, s=3, c=CRIMSON, zorder=3, alpha=0.9)
ax.set_xlim(-3.3, 3.3); ax.set_ylim(-3.1, 3.1); ax.axis("off")
plt.savefig(f"{OUT}/cover.png", dpi=200, facecolor=NAVY); plt.close()
print("ok")
