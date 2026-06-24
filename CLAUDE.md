# CLAUDE.MD -- Academic Project Development with Claude Code

<!-- HOW TO USE: Replace [BRACKETED PLACEHOLDERS] with your project info.
     Customize Beamer environments and CSS classes for your theme.
     Keep this file under ~150 lines — Claude loads it every session.
     See the guide at docs/workflow-guide.html for full documentation. -->

**Project:** Малошумящий усилитель (МШУ) на GaN с коэффициентом шума менее 1 дБ
**Type:** Диссертация + статьи (русский язык; статьи — английский)
**Institution:** ТУСУР (Томский государственный университет систем управления и радиоэлектроники)
**Author:** ARTisEZZ
**Branch:** main

## Предметная область

- **Тема:** Малошумящий усилитель (МШУ) на GaN для СВЧ-диапазона, целевой NF < 1 дБ
- **Направления:** СВЧ-электроника, микроэлектроника, полупроводниковые приборы
- **Инструменты моделирования:** Comsol, Silvaco/Centaurus TCAD, ADS, CST Studio, Matlab, Mathcad, Excel
- **Документооборот:** Word (основной), AutoCAD (чертежи), Figma (схемы/иллюстрации)
- **Языки написания:** Диссертация — русский; статьи — английский

---

## Core Principles

- **Plan first** -- enter plan mode before non-trivial tasks; save plans to `quality_reports/plans/`
- **Verify after** -- compile/render and confirm output at the end of every task
- **Single source of truth** -- Beamer `.tex` is authoritative; Quarto `.qmd` derives from it
- **Quality gates** -- nothing ships below 80/100
- **[LEARN] tags** -- when corrected, save `[LEARN:category] wrong → right` to [MEMORY.md](MEMORY.md)

Cross-session context lives in [MEMORY.md](MEMORY.md); past plans, specs, and session logs are in [quality_reports/](quality_reports/).

---

## Folder Structure

```
C_C_disser/
├── CLAUDE.md                    # Этот файл — читается каждую сессию
├── MEMORY.md                    # Контекст между сессиями
├── .claude/                     # Правила, скиллы, агенты, хуки
├── Bibliography_base.bib        # Единая база библиографии (BibTeX)
├── Figures/                     # Рисунки, графики, схемы
│   ├── simulations/             # Графики из ADS, Comsol, CST
│   ├── schematics/              # Принципиальные схемы
│   └── photos/                  # Фото макетов и образцов
├── Dissertation/                # Текст диссертации (Word / LaTeX)
├── Articles/                    # Статьи (EN)
│   ├── article_01_lna_gan/
│   └── article_02_microelectronics/
├── Slides/                      # Презентации (Beamer .tex)
├── Quarto/                      # Веб-версии слайдов
├── scripts/                     # Вспомогательные скрипты
├── quality_reports/             # Планы, логи сессий, отчёты проверки
├── explorations/                # Черновики и эксперименты
├── master_supporting_docs/      # Источники, статьи для изучения
└── templates/                   # Шаблоны документов ТУСУР
```

---

## Commands

```bash
# LaTeX (3-pass, XeLaTeX only)
cd Slides && TEXINPUTS=../Preambles:$TEXINPUTS xelatex -interaction=nonstopmode file.tex
BIBINPUTS=..:$BIBINPUTS bibtex file
TEXINPUTS=../Preambles:$TEXINPUTS xelatex -interaction=nonstopmode file.tex
TEXINPUTS=../Preambles:$TEXINPUTS xelatex -interaction=nonstopmode file.tex

# Deploy Quarto to GitHub Pages
./scripts/sync_to_docs.sh LectureN

# Quality score
python scripts/quality_score.py Quarto/file.qmd

# Palette sync (LaTeX ↔ SCSS)
./scripts/check-palette-sync.sh

# Surface-count sync (README ↔ CLAUDE.md ↔ guide ↔ landing page)
./scripts/check-surface-sync.sh
```

**Palette contract:** color names in `Preambles/header.tex` must match SCSS variables in `Quarto/theme-template.scss`. See [`Preambles/README.md`](Preambles/README.md).

---

## Quality Thresholds (advisory)

| Score | Checkpoint | Meaning |
|-------|------|---------|
| 80 | Commit | Good enough to save |
| 90 | PR | Ready for deployment |
| 95 | Excellence | Aspirational |

Enforced by `/commit` (halts + asks for override) **and** — once you run `./scripts/install-hooks.sh` — by a real git pre-commit hook (`.githooks/pre-commit`) that runs the surface-sync + quality (≥80) gates on every commit. Bypass sparingly with `SKIP_QUALITY_GATE=1` or `--no-verify`.

---

## Skills Quick Reference

The full table of all skills lives in [README.md](README.md#skills-claudeskills). Most-used, by workflow:

- **Slides / teaching:** `/create-lecture` `/compile-latex` `/deploy` `/qa-quarto` `/slide-excellence` `/syllabus` `/teach-from-paper` `/scaffold-exercises`
- **Papers / review:** `/review-paper` (`--peer`) `/seven-pass-review` `/respond-to-referees` `/verify-claims` `/proofread` `/humanize` `/submission-disclosures`
- **Data / reproducibility:** `/data-analysis` `/did-event-study` `/simulation-study` `/audit-reproducibility` `/diagnose` `/replication-package` `/capture-environment` `/power-analysis` `/disclosure-check`
- **Research / writing:** `/interview-me` `/lit-review` `/research-ideation` `/preregister` `/grant-proposal` `/data-management-plan`
- **Meta / workflow:** `/commit` `/learn` `/new-skill` `/checkpoint` `/context-status` `/deep-audit` `/coauthor-brief` `/triage-inbox`

Stata (`/stata-replication`), R packages (`/r-package-check`), TikZ (`/extract-tikz`, `/new-diagram`), and more — see the README for the complete index.

---

<!-- CUSTOMIZE: Replace placeholder rows ([your-env], [.your-class]) with your own.
     Delete the rows marked "(example — delete)" once you've added yours. -->

## Beamer Custom Environments

| Environment | Effect | Use Case |
| --- | --- | --- |
| `[your-env]` | [Description] | [When to use] |
| `keybox` | Gold background box | Key points *(example — delete)* |
| `definitionbox[Title]` | Blue-bordered titled box | Formal definitions *(example — delete)* |

## Quarto CSS Classes

| Class | Effect | Use Case |
| --- | --- | --- |
| `[.your-class]` | [Description] | [When to use] |
| `.smaller` | 85% font | Dense content *(example — delete)* |
| `.positive` | Green bold | Good annotations *(example — delete)* |

---

## Текущее состояние проекта

| Документ | Статус | Описание |
| --- | --- | --- |
| Диссертация — Глава 1 | 🔲 Не начата | Обзор литературы по МШУ на GaN |
| Диссертация — Глава 2 | 🔲 Не начата | Моделирование в Comsol / TCAD |
| Статья 1 (EN) | 🔲 Не начата | LNA GaN, NF < 1 dB — основные результаты |
| Статья 2 (EN) | 🔲 Не начата | Тема по микроэлектронике |

## Ключевые технические параметры МШУ

- Целевой коэффициент шума: **NF < 1 дБ**
- Технология: **GaN HEMT**
- Инструменты симуляции: Comsol, Silvaco TCAD, ADS, CST
