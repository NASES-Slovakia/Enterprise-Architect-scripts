# Enterprise-Architect-scripts

Tento repozitár združuje skripty používané architektonickým tímom nad Enterprise Architect (EA).

Cieľ:

- mať jedno spoločné miesto pre skripty,
- vedieť ich verzovať (Git),
- jednoducho ich aktualizovať u všetkých členov tímu,
- oddeliť **externé skripty** (spúšťané mimo EA) a **interné EA skripty** (v Scripting okne).

---

## Štruktúra

```text
Enterprise-Architect-scripts/
├─ README.md                   ← tento súbor
├─ external/                   ← skripty spúšťané mimo EA (cscript, cmd)
└─ ea-scripts/                 ← skripty určené ako šablóny do EA Scripting
    ├─ diagram/                ← skripty pre skupinu „Diagram“
    └─ project-browser/        ← (prípadne neskôr) skripty pre „Project Browser“
