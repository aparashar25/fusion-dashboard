"""IRR and NPV that behave like Excel's IRR(values, 0.01) and NPV(rate, values)."""
from __future__ import annotations
import numpy as np


def irr_monthly(cf, guess: float = 0.01) -> float:
    with np.errstate(all="ignore"):
        return _irr(cf, guess)


def _irr(cf, guess):
    cf = np.asarray(cf, float)
    if cf.size < 2 or not (np.any(cf > 0) and np.any(cf < 0)):
        return float("nan")
    t = np.arange(cf.size)
    r = guess
    for _ in range(200):                      # Newton from Excel's guess
        if r <= -1:
            break
        v = (1 + r) ** -t
        f = np.dot(cf, v)
        fp = np.dot(-t * cf, v / (1 + r))
        if fp == 0 or not np.isfinite(fp):
            break
        step = f / fp
        r -= step
        if abs(step) < 1e-13:
            if r > -0.99 and np.isfinite(r):
                return float(r)
            break
    # fallback: bracket and bisect on (-0.999999, 10)
    def npv(x):
        return np.dot(cf, (1 + x) ** -t)
    grid = np.concatenate([np.linspace(-0.999999, -0.5, 50), np.linspace(-0.5, 0.5, 200), np.linspace(0.5, 10, 50)])
    vals = [npv(x) for x in grid]
    roots = []
    for a, b, fa, fb in zip(grid[:-1], grid[1:], vals[:-1], vals[1:]):
        if np.sign(fa) != np.sign(fb):
            lo, hi = a, b
            for _ in range(200):
                mid = (lo + hi) / 2
                if np.sign(npv(mid)) == np.sign(npv(lo)):
                    lo = mid
                else:
                    hi = mid
            roots.append((lo + hi) / 2)
    if not roots:
        return float("nan")
    return float(min(roots, key=lambda x: abs(x - guess)))   # the root nearest Excel's guess


def irr_annual(cf) -> float:
    m = irr_monthly(cf)
    return float((1 + m) ** 12 - 1) if np.isfinite(m) else float("nan")


def npv_excel(rate: float, values) -> float:
    """Excel NPV: first value discounted one period."""
    v = np.asarray(values, float)
    return float(np.sum(v / (1 + rate) ** np.arange(1, v.size + 1)))
