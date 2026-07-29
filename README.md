# Portable Aimbot Target

Shared team training equipment — switch it on and practise. Packs into one checked bag.

Auto-aim can only be tested against a moving target, and producing one today means tying up a second robot plus someone to drive it. This removes that dependency, holds a commanded speed so two sessions a week apart are comparable, and travels to the competition in one checked bag.

**Primary users:** aimbot team members — the people testing and tuning auto-aim. Not the builder.
**Status:** pending decisions D1–D6 and open questions Q1–Q7.
**Deadlines:** workshop-usable by end of November 2026; travel-ready by end of February 2027 (Q1).
**Budget:** ~$910 all phases, or $550 for Phases 1–2 if staged.

---

## Contents

| Path | What it is |
|---|---|
| `01_Planning/Portable-Aimbot-Target-PRD.docx` | The PRD. Users, architecture, mechanical and electronics requirements, transport envelope, phases, risks. Start here. |
| `01_Planning/Portable-Aimbot-Target-Roadmap.xlsx` | Working tracker: 26-week roadmap, dashboard, weight budget, case sizing and packing manifest, BOM, risks, metrics, open questions. Edit this one as the project runs. |
| `02_Reference/` | Read-only PDF of the PRD for sharing. |
| `99_Generators/` | Scripts that produced the documents. Edit the script, re-run, regenerate. |
| `00_Archive/` | Superseded earlier scope. Safe to delete. |

The workbook is the living document; the PRD is the reasoning behind it. If they disagree, the workbook is newer.

---

## The design constraint that shapes everything

The primary user did not build this and will not read the PRD. They will pick it up at 9 p.m. with a robot half-assembled behind them. Anything that requires knowing a trick, remembering an order, or asking a question is a defect — not a documentation gap.

That drives four things: parts keyed so they cannot go together wrong, labelled to match their foam pocket and the setup card; joints verifiable by sight rather than by torque spec; a boot self-test that reports every sensor channel good or bad, so a user with no basis for suspecting the data still knows when to trust it; and Phase 1 acceptance being a teammate running a full session unaided, not the builder demonstrating it.

---

## The four architectural decisions

**One head, three bases.** Head, workshop base, rail base and travel base share a single frozen interface: bolt pattern, vertical datum, one 24 V + data connector. Freeze it in Phase 1. A usable machine reaches the team in six weeks because the first base is a lump of steel; the travel base is then designed against geometry that has stopped changing; and if the travel base proves too hard, the workshop keeps its equipment.

**Slot to locate, latch to clamp.** Slotting is right for assembly speed and wrong for a machine that vibrates. Slots position parts and generate no preload, so a design relying on slot friction works perfectly on the bench and is loose by the third session. Geometry locates, a dedicated latch clamps, neither does the other's job.

**TDOA processing lives on the rotating head.** Only digested impact events cross the slip ring, as digital packets. Contact noise is fatal to microsecond analogue timing and irrelevant to digital — and this cuts conductors across the interface from 16+ to about 4, making the slip ring smaller, cheaper and longer-lived.

**Build to the small armor format (~135 mm), not the large (~235 mm).** The plates are custom, so the format is yours. Small packs far smaller, is the harder and more realistic target, and TDOA still works comfortably — a wave crosses it in ~70 µs, which an STM32 timer resolves with enormous margin.

---

## Two constraints that bite

**Weight, not size.** Size fails visibly — a part fits the case or it doesn't. Weight accumulates invisibly and is discovered at check-in. Projected 18.6 kg against a 23 kg limit, but the stand and case are 44% of flying mass and the two least-defined items. Weigh parts as they're made; the `Weight Budget` tab is built for it.

**No lithium batteries in checked baggage.** A hard airline rule worldwide, and a standard RM pack far exceeds the carry-on allowance. Nothing battery-powered flies with this kit. Hence a universal-input mains PSU (100–240 V covers both Singapore and mainland China) with a DC input that also takes a borrowed pack on site.

Note that 158 cm is a *sum* of three dimensions, not a length — so case shape is a free design variable. Pick it after you know the largest part you can't shrink. The `Packing` tab models this.

---

## Gates

| Week | Test | Action on failure |
|---|---|---|
| 6 | Does a team member reach for it unprompted, and run it unaided? | Understand why not before adding a rail, sensors or a suitcase to it |
| 10 | Prototype plate ≤ 25 mm RMS on held-out grid positions | Don't build four. Ship with stock plates and referee counting — still good training equipment |
| 20–21 | Full dry-run packout under 23 kg and 158 cm | Decide deliberately: oversize fee, freight, or workshop-only. Not at the check-in desk |
| 24 | Teammate assembles from the case, unaided, under 10 min | The design is ambiguous. Fix keying and labelling, re-run |
| any | Two consecutive sessions needing tools to get running | Stop adding features. Everything here is downstream of it working |

---

## Answer first

**Q1** — which competition and date is the travel deadline? RMUL 2026's mainland sites ran March–May 2026, so this plan assumes the 2027 cycle around March 2027. If the real date is earlier, Phase 4 has to move ahead of Phase 3.

**Q7** — 3D printing and machining access, or everything cut by hand? Every mechanical effort estimate depends on it.

---

## Two things to keep in view

**Capacity.** 82 solo-days of scope against ~65 available at 2.5 days/week. The intended sacrifice order is Phase 4 polish, then Phase 3 UI depth, then the rail. Phase 1 is never sacrificed.

**M2 and M8 are the pair that matter.** Every other metric can be satisfied by a machine only its builder can operate. If teammates aren't using it within a month of Phase 1, something is wrong that no specification captured.

---

## Sources

- [RoboMaster GM6020 Brushless DC Motor User Guide](https://rm-static.djicdn.com/tem/17348/RoboMaster%20GM6020%20Brushless%20DC%20Motor%20User%20Guide.pdf) — 468 g, 1.2 N·m rated, 320 rpm, 18 mm hollow shaft
- [RoboMaster 2026 University Series Communication Protocol](https://bbs-web-static.robomaster.com/e43c29c066b443fab64614397cbfd3761764749912683/RoboMaster%202026%20University%20Series%20Communication%20Protocol%20V1.0.0%EF%BC%8820251203%EF%BC%89.pdf)
- [Armor Modules AM02 & AM12](https://www.robomaster.com/en-US/products/components/detail/1424) — 20 Hz max detection, 19 N min test pressure
- [RoboMaster University League](https://www.robomaster.com/en-US/robo/college-league) — 2026 mainland sites ran March–May 2026
- [A Novel Differential Time-of-Arrival Estimation Technique](https://www.ncbi.nlm.nih.gov/pmc/articles/PMC5676669/) — piezo TDOA impact localisation
- [A Study on Impact Force Detection Based on Piezoelectric Sensing](https://pmc.ncbi.nlm.nih.gov/articles/PMC9323614/)
# Moving_Aimbot_Target
