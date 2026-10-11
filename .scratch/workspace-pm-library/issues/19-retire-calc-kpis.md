# 19: Retire calc_kpis

**Status:** resolved

**Blocked by:** 14, and the user's Looker Studio report going live

**What to build:** Remove the `calc_kpis` dashboard rendering and its formatting code once the Looker report built on `fact_presence` replaces it.

## Comments

_2026-10-12 — calc_kpis removed at the user's request (ADR 0036), and `refreshKpis` now only rebuilds fact_presence._
