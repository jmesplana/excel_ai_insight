# Framework 2 (Ebola) — how the feedback coding works

> **Implementation update:** The discussion below describes the original flat v1 configuration.
> The app now supports dataset-independent v2 dependencies and derived lookups.
> Import `examples/ebola-framework2-hierarchical.jev-config.json` for sequential
> Type → valid Topic → leaf Code classification plus a deterministic sous-dimension lookup.
> Uncertain parents block downstream decisions for review. The v1 example remains available.
> All framework content lives in that JSON, with no Ebola-specific runtime behavior.
> See [JEV_CONFIGURATION.md](JEV_CONFIGURATION.md) for the configurable workflow and reporting contract.

Why the Jev configuration does not reproduce the Excel sheet's three-level
structure, what that costs, and what the codes actually are.

| | |
| --- | --- |
| Source sheet | `examples/sample_framework_bvd.xlsx`, tab `framework`, table `FrameWork2` (`A2:BQ34`) |
| Parsed tree | `framework2_taxonomy.json` — rebuild: `python scripts/build_framework2_taxonomy.py` |
| Topic grid | `framework2_topic_grid.json` — rebuild: `python scripts/build_framework2_topic_grid.py` |
| Jev config | `examples/ebola-framework2.jev-config.json` — rebuild: `python scripts/build_framework2_jev_config.py` |
| Overlap figures | `python scripts/build_framework2_overlap.py` (§3) |

---

## 1. The taxonomy, by the numbers

The sheet is a three-level tree: **Type → Sous-dimension → Code**.

| Level | Count | Fits one Jev question? (cap 255) |
| --- | --- | --- |
| Type | 7 | Yes |
| Sous-dimension | 61 | Yes |
| Code (leaf) | **444 unique** | **No — 189 over the cap** |

Codes per type:

| Type | Sous-dimensions | Codes |
| --- | --- | --- |
| Observations, perceptions ou croyances | 12 | 158 |
| Questions | 12 | 140 |
| Demandes | 12 | 62 |
| Suggestions | 9 | 32 |
| Signalements et préoccupations | 9 | 31 |
| Appréciations | 5 | 15 |
| Allégations et Signalement d'incident | 2 | 7 |

Totals: 7 types, 61 sous-dimensions, 445 leaf paths (444 unique labels).

---

## 2. What the Excel structure is for

The sheet drives its levels with data-validation dropdowns. The mechanism is
visible in the workbook's defined names:

```
_MainFrame2List = INDEX(FrameWork2[#Data], 0,
                        MATCH(<value you picked>, FrameWork2[#Headers], 0))
```

Pick a value → `MATCH` finds the column whose **header** is that value → that
column's body becomes the next dropdown's list. Header = parent, column body =
children.

It is split three ways because **a human cannot pick from a 444-item
dropdown.** Narrowing 7 → 61 → ≤27 makes each pick scrollable. The constraint
being solved is human eyes and a mouse.

---

## 3. The key finding: two dimensions at the sous-dimension level

The **sous-dimension** level is not an independent layer. It is the Type name
concatenated onto a shared topic vocabulary. Verified across all 61
sous-dimensions:

- **61 of 61** sous-dimension labels belong to exactly one Type.
- **(Type, Topic) → Sous-dimension is a function** — 61 grid entries, zero
  duplicate keys, all 61 sous-dimensions reachable.

So at that level the framework is **two independent dimensions**:

- **Type (7)** — the speech act: observation, question, demand, appreciation…
- **Topic (12)** — what it is about: vaccins, MSP, traitement, réponse…

### But the leaf level is genuinely type-specific

This does **not** extend downward. Measured leaf-vocabulary overlap between
types that share a topic (`scripts/build_framework2_overlap.py`, of the shorter list,
how much is echoed in the longer):

| Topic | Types | Overlap | | Topic | Types | Overlap |
| --- | --- | --- | --- | --- | --- | --- |
| MSP | 6 | 41% | | Réponse | 7 | 15% |
| Services de santé | 6 | 37% | | Comportements | 4 | 12% |
| Impact de l'épidémie | 5 | 29% | | Épidémie | 3 | 11% |
| Acteurs de la réponse | 7 | 26% | | Traitement | 5 | 9% |
| Sûreté et sécurité | 5 | 24% | | Maladie | 4 | 7% |
| Vaccins | 6 | 21% | | Autres | 3 | 0% |

**Mean overlap: 19%.** Four-fifths of each leaf list is unique to its type.

This is a substantive property of the framework, not redundancy. What a
community *observes* about a treatment and what it *asks* about one are
different things: observations carry beliefs, rumour and misinformation
(`Croyances concernant l'utilisation de remèdes naturels pour le traitement`),
while questions carry information gaps (`Questions sur l'efficacité des
traitements approuvés`). RCCE teams act on those differently — rumours route to
message correction, questions route to information provision. The asymmetry is
deliberate.

> **Correction.** An earlier draft of this document claimed the per-type lists
> "run in near-lockstep, item for item, with only the prefix differing," and
> illustrated it with a hand-aligned vaccins table. That was wrong on both
> counts. The vaccins columns share 9 of 13 leaves semantically (21% by the
> measure above, close to the 19% average, not exceptional), and the aligned
> table paired items that are not equivalent — `disponibilité du vaccin` against
> `qualité du vaccin`, `structures de vaccination` against `infrastructures
> liées au vaccin` — while omitting `Questions sur l'existence d'un vaccin pour
> la maladie`, which has no Observations counterpart. The example was also a
> single hand-picked branch presented as typical.
>
> The correction does not change the configuration. The config asks Type and
> Topic and derives the **sous-dimension**; it never asks for leaf codes and
> never depended on leaves being shared. The lockstep claim was rhetorical
> support, and it was unsound.

### The grid

| Topic | Obs | Quest | Sugg | Appr | Dem | Signal | Allég |
|---|---|---|---|---|---|---|---|
| Maladie | ● | ● | · | · | ● | ● | · |
| Épidémie | ● | ● | · | · | ● | · | · |
| Impact de l'épidémie | ● | ● | ● | · | ● | ● | · |
| Comportements | ● | ● | ● | · | ● | · | · |
| MSP | ● | ● | ● | ● | ● | ● | · |
| Traitement | ● | ● | ● | · | ● | ● | · |
| Services de santé | ● | ● | ● | ● | ● | ● | · |
| Vaccins | ● | ● | ● | ● | ● | ● | · |
| Réponse | ● | ● | ● | ● | ● | ● | ● |
| Acteurs de la réponse | ● | ● | ● | ● | ● | ● | ● |
| Sûreté et sécurité | ● | ● | ● | · | ● | ● | · |
| Autres | ● | ● | · | · | ● | · | · |

*61 of 84 cells exist; 23 gaps (·). The gaps are combinations the framework
simply never labelled — there is no "Appréciation sur la maladie".*

### Merges applied

Four topics are spelled inconsistently across types in the sheet. Every merge
is declared in the `MERGE` dict of `scripts/build_framework2_topic_grid.py`:

| Canonical topic | Merged from |
| --- | --- |
| **Traitement** | `traitement` (Sugg, Signal) · `traitements` (Quest, Dem) · `traitement de la maladie` (Obs) |
| **Vaccins** | `vaccins` (5 types) · `vaccination` (Signal) |
| **Services de santé** | `services de sante` (5 types) · `services de soins de sante` (Quest) |
| **Sûreté et sécurité** | `surete et la securite` (3) · `surete et securite` (Sugg) · `securite et la surete` (Signal — words reversed) |

`Autres` merges the three type-specific catch-alls (`Autres questions`,
`Autre demande`, `Autres observations…`).

> **One judgment call, flagged for review.** `Demandes` has two "épidémie"
> sous-dimensions. Their contents differ: `Demande de plus d'informations sur
> l'épidémie` covers symptoms, transmission, recovery and the infectious agent
> — **about the disease** — while `Demande relative à l'épidémie` covers
> trends, evolution and stopping it — **about the outbreak**. The first is
> therefore mapped to topic `Maladie`, not `Épidémie`. The label says
> "épidémie"; the contents say disease. Worth a second opinion from whoever
> maintains the framework.

---

## 4. What the Jev config does

Two questions, one per dimension, asked **once each**:

| Output column | Options |
| --- | --- |
| `framework2_type` | 7 — the speech act |
| `framework2_topic` | 12 — the subject |

Plus the five type-independent judgments carried over unchanged from
`examples/ebola-feedback-coding.jev-config.json`: `feedback_sensibilite`,
`feedback_criticalite`, `feedback_action`, `feedback_secteur`,
`feedback_referent`.

**The sous-dimension is then derived by lookup**, not asked for:
`(Type, Topic) → Sous-dimension` via `framework2_topic_grid.json`.

### Why this shape and not the sheet's

**Asking for the sous-dimension directly asks for the Type twice.** A 61-option
sous-dimension question offers labels that each restate a speech act; if it
returns `Questions sur les vaccins` on a row typed `Observations`, the two
columns openly contradict each other. Deriving it makes that contradiction
**structurally impossible**. It is also an easier question: 12 clean topics
beats 61 labels where two-thirds of each label restates something already
decided.

**The app sends one request per row, not per question.** `insights/jev.py:196`
sends one request per row carrying every question, evaluated in parallel
against one shared state. The 7-question config is **1 API call per row**.

A cascade (ask Type → narrow options → ask Sous-dimension → narrow → ask Code)
cannot use that endpoint: each question's options depend on the previous
*answer*, so the calls must run sequentially — **3 calls per row instead of 1**
— and it bypasses the batching, the 8-way row concurrency (`MAX_JEV_WORKERS`)
and the checkpointing in `analyze_batch_jev`.

> **Correction to earlier advice in this project's history:** the cascade was
> described as "~4 calls instead of the naive 9." That was wrong about this
> architecture. The flat config is 1 call per row; the cascade is 3. The
> cascade is the expensive option, not the cheap one.

**The import format has no dependency field.** `deserializeJevConfig`
(`static/js/jev-config.js:120`) keeps exactly four keys per question:
`outputColumnName`, `questionType`, `instructions`, `options`. Any `dependsOn`
field would be **silently dropped at import** — a config that looks right and
is not.

---

## 5. What this costs

**The leaf code level is not reachable in-app.** 444 options against a 255 cap,
with no dependency field to narrow them per row. The config delivers
**Type + Topic → 61 derived sous-dimensions, not 444 codes.**

**The 23 grid gaps are guarded by instructions, not structure.** The valid
topic list for each type is written into the topic question's prompt, but
nothing prevents Jev returning Type=`Appréciations` + Topic=`Maladie`, which
has no sous-dimension.

**Mitigation:** after a run, look each (type, topic) pair up in
`framework2_topic_grid.json`. A miss means the pair landed in a gap — flag
those rows for review. That check is cheap and catches the only failure mode
this design has.

### The trade, stated plainly

- If 61 sous-dimensions is enough resolution → this shape is better than the
  sheet's: fewer questions, each easier, and contradiction between Type and
  Sous-dimension is structurally impossible rather than merely discouraged.
- If the 444-code leaf is genuinely needed → it requires sequential calls
  outside the app, giving up the batch endpoint. Note the leaf level carries
  real information the sous-dimension does not (see §3: 19% mean overlap), so
  this is a genuine loss, not a cosmetic one.

### Hypotheses, not findings

Two arguments for this approach are **plausible but unmeasured on this data**,
and are recorded here as hypotheses to test rather than conclusions:

1. *That a 12-option question yields better-separated confidence scores than a
   61-option one.* Reasonable — two-thirds of each 61-label is a prefix
   restating the Type, so the scores should reflect topical ambiguity rather
   than label-lookalike confusion. Untested.
2. *That machine coding is more consistent than human coding at volume.* Also
   reasonable given a 444-item cascading dropdown over thousands of rows, but
   there is no inter-rater reliability data for the manual process here and no
   calibration run for Jev.

Jev returns a calibrated confidence per answer and `includeConfidence` is
enabled, so the raw material for testing both is produced on every run. **A
calibration pass — hand-code a sample, compare, and see where disagreements
fall on the confidence scale — is what sets a defensible review threshold.**
Until that is run, treat the confidence column as a triage aid, not as evidence
of accuracy.

**Open question:** does downstream reporting need the leaf code, or is the
sous-dimension the level actually analyzed?

---

## 6. Two source-sheet typos (these break the Excel dropdowns too)

Both were found by the parser leaving exactly two unresolved labels and exactly
two orphan columns, which paired up. Both are recorded in `HEADER_OVERRIDES` in
`scripts/build_framework2_taxonomy.py`, bound by position and verified contents — not
by name similarity.

| Label used in the parent column | Actual column header | Evidence |
| --- | --- | --- |
| `Questions sur les MSP` | `Demande relative aux MSP2` | Column 51, between `Questions sur les comportements` (50) and `Questions sur les traitements` (52); contents are all `Questions sur…` |
| `Allégations concernant &nbsp;les acteurs de la réponse` (double space) | `Allegations sur les acteurs de la réponse` | Different preposition, missing accent; contents are all `Allégations de…` |

Excel's `MATCH` is exact, so **selecting either value in the sheet returns
`#N/A`** and the dependent dropdown stays empty. Worth fixing at the source.

Also dropped during parsing: **75 `______ X ______` separator cells**, which are
visual grouping only, not codes.

---

## 7. The codes

Full listing, grouped by Type and Sous-dimension, with each sous-dimension's
derived topic. Counts in brackets.

### Observations, perceptions ou croyances

*12 sous-dimensions, 158 codes*

<details>
<summary><strong>Croyances concernant la maladie</strong> — topic <code>Maladie</code>, 27 codes</summary>

- Croyances concernant l’existence de la maladie
- Croyances concernant la non-existence de la maladie
- Croyances indiquant le déni de la maladie
- Croyances concernant l’agent infectieux (virus ou bactérie)
- Croyances concernant la résistance du virus
- Croyances concernant les symptômes
- Croyances indiquant l’ignorance des symptômes
- Croyances concernant la comparaison avec d’autres maladies
- Croyances concernant les modes de transmission
- Croyances concernant la transmission de personne à personne
- Croyances concernant la transmission à partir de corps morts
- Croyances concernant les matériaux contaminés
- Croyances concernant l’eau contaminée
- Croyances concernant les transmissions par les animaux
- Croyances selon lesquelles certains groupes de personnes sont les propagateurs de la maladie
- Croyances relatives aux comportements à risque de contamination
- Croyances concernant d’autres modes de transmission
- Croyances ou commentaires concernant le risque et la gravité
- Peurs d’être contaminé
- Croyances concernant le fait d’être contaminé
- Croyances selon lesquelles quelqu’un est contaminé
- Croyances selon lesquelles certaines communautés sont plus exposées au virus
- Croyances concernant la gravité de la maladie
- Observations ou croyances concernant le rétablissement
- Croyances concernant les personnes rétablies de la maladie
- Croyances selon lesquelles les personnes rétablies pourraient propager le virus
- Autres croyances concernant la maladie

</details>

<details>
<summary><strong>Observations sur l’épidémie</strong> — topic <code>Épidémie</code>, 11 codes</summary>

- Croyances concernant l’origine de l’épidémie
- Croyances concernant l’origine zoonotique de l’épidémie
- Croyances selon lesquelles l’épidémie est d’origine surnaturelle
- Croyances selon lesquelles l’épidémie fait partie d’un complot
- Croyances selon lesquelles l’épidémie a été apportée par des étrangers
- Croyances concernant l’évolution de l’épidémie
- Croyances selon lesquelles l’épidémie n’existe pas dans cette région ou ce pays
- Croyances concernant le nombre de cas et les tendances
- Croyances concernant la durée de l’épidémie
- Croyances selon lesquelles l’épidémie est terminée
- Autres croyances concernant l’épidémie

</details>

<details>
<summary><strong>Observations sur l’impact de l’épidémie</strong> — topic <code>Impact de l'épidémie</code>, 17 codes</summary>

- Observations sur l'impact de l'épidémie
- Observations concernant l’impact social
- Préoccupations liées à la stigmatisation due à la maladie
- Préoccupations concernant la stigmatisation ethnique ou culturelle
- Préoccupations concernant la stigmatisation des agents de santé
- Préoccupations concernant la stigmatisation de la famille ou des contacts du cas
- Préoccupations concernant la stigmatisation des personnes rétablies
- Observations concernant l’impact sur la cohésion sociale
- Observations concernant l’impact économique
- Observations concernant l’impact sur les moyens de subsistance
- Observations concernant la sécurité alimentaire
- Observations concernant l’impact sur l’emploi
- Observations concernant les conséquences sur les services de base
- Observations concernant l’impact sur les services administratifs
- Observations concernant l’impact sur les services de santé
- Observations concernant l’impact sur l’éducation
- Observations concernant les conséquences sur la santé mentale

</details>

<details>
<summary><strong>Croyances concernant les comportements</strong> — topic <code>Comportements</code>, 13 codes</summary>

- Croyances concernant les comportements préventifs
- Croyances concernant le nettoyage et la désinfection des surfaces touchées
- Croyances concernant le lavage des mains
- Croyances concernant la distanciation sociale
- Croyances concernant les pratiques sexuelles sans risque
- Croyances concernant la grossesse et l’allaitement
- Croyances concernant les cas suspects
- Croyances concernant l’isolement avec une personne malade
- Croyances concernant les comportements culturels et traditionnels
- Croyances concernant la prière comme mesure préventive
- Croyances concernant l’éradication des animaux comme mesure préventive
- Croyances concernant les enterrements et comportements funéraires traditionnels
- Croyances concernant l’utilisation de remèdes naturels pour la prévention

</details>

<details>
<summary><strong>Observations ou croyances concernant les MSP</strong> — topic <code>MSP</code>, 9 codes</summary>

- Observation concernant les mesures de santé
- Observations concernant la campagne de communication sur l’épidémie
- Observations sur l’adhésion aux mesures de santé
- Observations sur les obstacles à l’application des mesures de santé
- Observations ou croyances concernant le traçage des contacts
- Observations ou croyances concernant les enterrements sûrs et dignes
- Observations ou croyances concernant les restrictions
- Observations ou croyances concernant les restrictions de mouvement
- Observations ou croyances concernant les voyages internationaux

</details>

<details>
<summary><strong>Croyances concernant le traitement de la maladie</strong> — topic <code>Traitement</code>, 11 codes</summary>

- Croyances concernant les centres de traitement
- Observations concernant l’accès aux centres de traitement
- Observations ou croyances concernant la sécurité des centres de traitement
- Croyances concernant les traitements approuvés
- Croyances concernant la disponibilité des traitements approuvés
- Croyances concernant l’utilisation de médicaments spécifiques
- Croyances concernant l’efficacité des traitements approuvés
- Croyances concernant les effets secondaires des traitements approuvés
- Croyances concernant l’utilisation de traitements non approuvés
- Croyances concernant l’utilisation de remèdes naturels pour le traitement
- Croyances concernant l’utilisation de la médecine traditionnelle pour le traitement

</details>

<details>
<summary><strong>Croyances concernant les services de santé</strong> — topic <code>Services de santé</code>, 9 codes</summary>

- Croyances concernant l’accès aux soins de santé
- Croyances concernant la qualité des soins de santé
- Croyances selon lesquelles les structures de santé ne sont pas sûres
- Croyances concernant la gestion des décès dans les structures de santé
- Croyances concernant les agents de santé
- Croyances concernant le rôle des agents de santé
- Croyances concernant le rôle des agents de santé communautaires
- Croyances concernant les compétences des agents de santé
- Croyances concernant les équipements de protection pour les agents de santé

</details>

<details>
<summary><strong>Observations ou croyances concernant les vaccins</strong> — topic <code>Vaccins</code>, 13 codes</summary>

- Observations ou croyances concernant les campagnes de vaccination
- Observations ou croyances concernant les informations reçues sur le vaccin
- Déclarations indiquant la non-acceptation du vaccin
- Observations ou croyances concernant le développement des vaccins
- Observations ou croyances concernant la disponibilité du vaccin
- Observations ou croyances concernant l’accès au vaccin
- Observations ou croyances concernant les structures de vaccination
- Observations ou croyances concernant le prix du vaccin
- Observations ou croyances concernant la qualité du vaccin
- Observations ou croyances concernant l’efficacité du vaccin
- Observations ou croyances concernant les effets secondaires du vaccin
- Déclarations indiquant une suspicion envers le vaccin
- Autres croyances et observations concernant les vaccins

</details>

<details>
<summary><strong>Observations sur la réponse</strong> — topic <code>Réponse</code>, 21 codes</summary>

- Observations concernant l’information sur la réponse
- Déclarations concernant le manque d’information
- Observations concernant les rumeurs dans la communauté
- Croyances selon lesquelles les informations reçues sont fausses
- Observations sur l’accès à l’assistance
- Observations sur la fourniture d’une assistance sanitaire
- Observations sur la fourniture d’articles de prévention des maladies
- Observations ou croyances concernant les activités de réponse pour des groupes spécifiques
- Observations ou croyances concernant le processus de sélection de l’assistance
- Observations ou croyances concernant l’accès au soutien des moyens de subsistance
- Observations sur la qualité de la réponse
- Observations sur les compétences des acteurs de la réponse
- Observations sur le calendrier de l’action
- Observations sur l’efficacité de l’action
- Observations sur l’engagement communautaire
- Observations sur la localisation de la réponse
- Observations sur l’implication des membres de la communauté dans la réponse
- Observations sur les solutions dirigées par la communauté
- Déclarations concernant la non-acceptation de la réponse
- Déclarations concernant la fatigue face à la réponse
- Déclarations concernant le rejet de la réponse

</details>

<details>
<summary><strong>Observations sur les acteurs de la réponse</strong> — topic <code>Acteurs de la réponse</code>, 14 codes</summary>

- Déclarations indiquant une suspicion envers les acteurs et intervenants
- Déclarations indiquant une suspicion envers le gouvernement
- Déclarations indiquant une suspicion envers la science
- Déclarations indiquant une suspicion envers les médias et la presse
- Déclarations indiquant une suspicion envers les leaders locaux
- Déclarations indiquant une suspicion envers les agents de santé
- Déclarations indiquant une suspicion envers les acteurs humanitaires
- Croyances selon lesquelles certaines personnes/organisations gagnent de l’argent à cause de l’épidémie
- Croyances selon lesquelles la maladie est propagée par les acteurs de la réponse
- Observations sur le rôle des acteurs et intervenants
- Observations ou croyances concernant le gouvernement
- Observations ou croyances concernant les organisations internationales
- Observations ou croyances concernant la Croix-Rouge
- Observations ou croyances concernant les acteurs locaux

</details>

<details>
<summary><strong>Observations sur la sûreté et la sécurité</strong> — topic <code>Sûreté et sécurité</code>, 12 codes</summary>

- Observations sur la sûreté et la sécurité
- Déclarations concernant les conditions de sécurité
- Déclarations concernant le sentiment de menace
- Déclarations concernant les routes ou accès dangereux
- Déclarations concernant la violence
- Déclarations concernant le SEA ou les atteintes à la dignité
- Déclarations concernant la discrimination
- Déclarations concernant les protestations violentes
- Déclarations concernant les violences communautaires
- Déclarations concernant les groupes armés
- Déclarations concernant les forces de sécurité
- Déclarations concernant les personnes disparues

</details>

<details>
<summary><strong>Autres observations, perceptions ou croyances</strong> — topic <code>Autres</code>, 1 codes</summary>

- Autres observations, perceptions ou croyances

</details>

### Questions

*12 sous-dimensions, 140 codes*

<details>
<summary><strong>Questions sur la maladie</strong> — topic <code>Maladie</code>, 25 codes</summary>

- Questions sur l’existence de la maladie
- Questions sur la négation de la maladie
- Questions sur l’agent infectieux (virus, bactérie)
- Questions sur la résistance du virus
- Questions sur les symptômes
- Questions sur la comparaison avec d’autres maladies
- Questions sur la transmission
- Questions sur la transmission de personne à personne
- Questions sur la transmission par les corps des défunts
- Questions sur les matériaux contaminés
- Questions sur l’eau contaminée
- Questions sur les transmissions animales
- Questions sur qui propage la maladie
- Questions sur les comportements à risque de contamination
- Questions sur d’autres modes de transmission
- Questions sur le risque et la gravité
- Questions sur le risque d’être contaminé
- Questions sur le fait d’avoir des symptômes
- Questions sur une personne présentant des symptômes
- Questions sur les communautés à risque
- Questions sur la gravité de la maladie
- Questions sur la guérison
- Questions sur les personnes guéries de la maladie
- Questions sur la transmission par les personnes guéries
- Autres questions sur la maladie

</details>

<details>
<summary><strong>Questions sur l'épidémie</strong> — topic <code>Épidémie</code>, 9 codes</summary>

- Questions sur l’origine de l’épidémie
- Questions sur l’origine zoonotique de l’épidémie
- Questions sur le rôle surnaturel dans l’origine de l’épidémie
- Questions indiquant que l’épidémie fait partie d’un complot
- Questions sur l’introduction de l’épidémie par des étrangers
- Questions sur l’évolution de l’épidémie
- Autre question sur l’épidémie
- Questions sur l’existance de l’épidémie
- Question concernant la fin de l'épidémie.

</details>

<details>
<summary><strong>Questions sur l'impact de l'épidémie</strong> — topic <code>Impact de l'épidémie</code>, 20 codes</summary>

- Impact social
- Questions sur l’impact social
- Questions sur l’exclusion sociale ou la stigmatisation due à la maladie
- Questions sur la stigmatisation ethnique ou culturelle
- Questions sur la stigmatisation des agents de santé
- Questions sur la stigmatisation de la famille ou des contacts du cas
- Questions sur la stigmatisation des personnes guéries
- Questions sur l’impact sur la cohésion sociale
- Impact économique
- Questions sur l’impact économique
- Questions sur l’impact sur les moyens de subsistance
- Questions sur la sécurité alimentaire
- Questions sur l’impact sur l’emploi
- Services basiques
- Questions sur les conséquences sur les services de base
- Questions sur l’impact sur les services administratifs
- Questions sur l’impact sur les services de santé
- Questions sur l’impact sur l’éducation
- Santé mentale
- Questions sur les conséquences sur la santé mentale

</details>

<details>
<summary><strong>Questions sur les comportements</strong> — topic <code>Comportements</code>, 13 codes</summary>

- Questions sur les comportements préventifs
- Questions sur le nettoyage et la désinfection des surfaces touchées
- Questions sur le lavage des mains
- Questions sur la distanciation sociale
- Questions sur les pratiques sexuelles sûres
- Questions sur la grossesse et l’allaitement
- Questions sur ce qu’il faut faire si l’on soupçonne quelqu’un d’être malade
- Questions sur l’isolement des personnes malades
- Questions sur les comportements culturels et traditionnels
- Questions sur la prière comme mesure préventive
- Questions sur l’éradication des animaux comme mesure préventive
- Questions sur les enterrements et pratiques funéraires traditionnels
- Questions sur l’utilisation de remèdes naturels pour la prévention

</details>

<details>
<summary><strong>Questions sur les MSP</strong> — topic <code>MSP</code>, 8 codes</summary>

- Questions sur les mesures de santé
- Questions sur la campagne de communication autour de l’épidémie
- Questions sur la manière de faire appliquer les mesures de santé
- Questions sur le traçage des contacts
- Questions sur les enterrements sûrs et dignes
- Questions sur les restrictions
- Questions sur les restrictions de mouvement
- Observations ou croyances concernant les voyages internationaux

</details>

<details>
<summary><strong>Questions sur les traitements</strong> — topic <code>Traitement</code>, 12 codes</summary>

- Questions sur les centres de traitement
- Questions sur l’accès aux centres de traitement
- Questions sur la sécurité des centres de traitement
- Questions concernant les raison pour lesquelles les gens meurent a l’hopitale
- Questions sur les traitements approuvés pour la maladie
- Questions sur la disponibilité des traitements approuvés
- Questions sur l’utilisation de médicaments spécifiques
- Questions sur l’efficacité des traitements approuvés
- Questions sur les effets secondaires des traitements approuvés
- Questions sur l’utilisation de traitements non approuvés
- Questions sur l’utilisation de remèdes naturels pour le traitement
- Questions sur l’utilisation de la médecine traditionnelle pour le traitement

</details>

<details>
<summary><strong>Questions sur les services de soins de santé</strong> — topic <code>Services de santé</code>, 9 codes</summary>

- Questions sur l’accès aux soins de santé
- Questions sur la qualité des soins de santé
- Questions sur la sécurité dans les établissements de santé
- Questions sur la gestion des décès dans les établissements de santé
- Questions sur les agents de santé
- Questions sur le rôle des agents de santé
- Questions sur le rôle des agents de santé communautaires
- Questions sur les compétences des agents de santé
- Questions sur les équipements de protection pour les agents de santé

</details>

<details>
<summary><strong>Questions sur les vaccins</strong> — topic <code>Vaccins</code>, 13 codes</summary>

- Questions sur les campagnes de vaccination
- Questions sur les informations reçues sur le vaccin
- Questions sur la non-acceptation
- Questions sur le développement du vaccin
- Questions sur l’existence d’un vaccin pour la maladie
- Questions sur l’accès au vaccin
- Questions sur les infrastructures liées au vaccin
- Questions sur le prix du vaccin
- Questions sur la qualité du vaccin
- Questions sur l’efficacité du vaccin
- Questions sur les effets secondaires du vaccin
- Questions indiquant une suspicion à l’égard du vaccin
- Autres questions sur les vaccins

</details>

<details>
<summary><strong>Questions sur la réponse</strong> — topic <code>Réponse</code>, 16 codes</summary>

- Questions sur les informations concernant la réponse
- Questions liées aux fausses nouvelles suspectées ou aux rumeurs dans la communauté
- Questions sur la confiance dans les informations reçues
- Questions sur l’accès à l’assistance
- Questions sur la fourniture d’une assistance sanitaire
- Questions sur la fourniture d’autres articles de prévention des maladies
- Questions sur les activités de réponse pour des groupes spécifiques
- Questions sur le processus de sélection de l’assistance
- Questions sur l’accès au soutien des moyens de subsistance
- Questions relatives à la qualité de la réponse
- Questions sur le calendrier des actions
- Questions sur l’efficacité des actions
- Questions sur l’engagement communautaire
- Questions sur la localisation de la réponse
- Questions sur la manière de participer à la réponse
- Questions sur les solutions menées par la communauté

</details>

<details>
<summary><strong>Questions relative aux acteurs de la réponse</strong> — topic <code>Acteurs de la réponse</code>, 13 codes</summary>

- Questions révélant une méfiance envers des personnes ou des institutions
- Questions révélant une méfiance envers le gouvernement
- Questions révélant une méfiance envers la science
- Questions révélant une méfiance envers les médias
- Questions révélant une méfiance envers les dirigeants locaux
- Questions révélant une méfiance envers les agents de santé
- Questions révélant une méfiance envers les acteurs humanitaires
- Questions révélant que les acteurs humanitaires profitent financièrement de l’épidémie
- Questions sur le rôle des acteurs de la réponse
- Questions sur le rôle du gouvernement
- Questions sur le rôle des organisations internationales
- Questions sur le rôle de la Croix-Rouge
- Questions sur le rôle des acteurs locaux

</details>

<details>
<summary><strong>Questions sur la sûreté et la sécurité</strong> — topic <code>Sûreté et sécurité</code>, 1 codes</summary>

- Questions sur la sûreté et la sécurité

</details>

<details>
<summary><strong>Autres questions</strong> — topic <code>Autres</code>, 1 codes</summary>

- Autres questions

</details>

### Suggestions

*9 sous-dimensions, 32 codes*

<details>
<summary><strong>Suggestions sur l'impact de l'épidémie</strong> — topic <code>Impact de l'épidémie</code>, 3 codes</summary>

- Suggestions concernant l’impact sur les services de santé
- Suggestions concernant l’impact sur l’éducation
- Suggestions concernant la santé mentale

</details>

<details>
<summary><strong>Suggestions concernant les comportements</strong> — topic <code>Comportements</code>, 1 codes</summary>

- Suggestions concernant les comportements

</details>

<details>
<summary><strong>Suggestions sur les MSP</strong> — topic <code>MSP</code>, 8 codes</summary>

- Suggestions concernant les mesures de santé
- Suggestions concernant la campagne de communication sur l’épidémie
- Suggestions pour lever les obstacles à l’application des mesures de santé
- Suggestions concernant le traçage des contacts
- Suggestions concernant les enterrements sûrs et dignes
- Suggestions concernant les restrictions
- Suggestions concernant les restrictions de mouvement
- Suggestions concernant les voyages internationaux

</details>

<details>
<summary><strong>Suggestions sur le traitement</strong> — topic <code>Traitement</code>, 2 codes</summary>

- Suggestions concernant les traitements non approuvés
- Suggestions concernant l’utilisation de la médecine traditionnelle comme traitement

</details>

<details>
<summary><strong>Suggestions sur les services de santé</strong> — topic <code>Services de santé</code>, 2 codes</summary>

- Suggestions concernant l’accès aux soins de santé
- Suggestions concernant le rôle des agents de santé communautaires

</details>

<details>
<summary><strong>Suggestions sur les vaccins</strong> — topic <code>Vaccins</code>, 3 codes</summary>

- Suggestions concernant les campagnes de vaccination
- Suggestions concernant les informations reçues sur le vaccin
- Suggestions concernant la non-acceptation du vaccin

</details>

<details>
<summary><strong>Suggestions sur la réponse</strong> — topic <code>Réponse</code>, 10 codes</summary>

- Suggestions concernant l’information sur la réponse
- Suggestion de réduire l’information sur la maladie
- Suggestion d’adapter le contenu de la campagne d’information
- Suggestion d’utiliser un canal spécifique pour communiquer avec les communautés
- Suggestions concernant l’accès à la réponse
- Suggestions de concentrer les activités pour des groupes spécifiques
- Suggestions concernant la qualité de la réponse
- Suggestions relatives à l’engagement communautaire
- Suggestions relatives aux solutions dirigées par la communauté
- Suggestions pour modifier la réponse

</details>

<details>
<summary><strong>Suggestions sur les acteurs de la réponse</strong> — topic <code>Acteurs de la réponse</code>, 1 codes</summary>

- Suggestions sur les acteurs de la réponse

</details>

<details>
<summary><strong>Suggestions concernant la sureté et sécurité</strong> — topic <code>Sûreté et sécurité</code>, 2 codes</summary>

- Suggestions concernant les conditions de sécurité
- Suggestions concernant la violence

</details>

### Appréciations

*5 sous-dimensions, 15 codes*

<details>
<summary><strong>Appréciation concernant les MSP</strong> — topic <code>MSP</code>, 2 codes</summary>

- Appréciation concernant les MSP
- Appréciations de l'équipe SDB

</details>

<details>
<summary><strong>Appréciation des services de santé</strong> — topic <code>Services de santé</code>, 3 codes</summary>

- Appréciations ou encouragements envers les agents de santé
- Appréciations du rôle des agents de santé
- Appréciations du rôle des agents de santé communautaires

</details>

<details>
<summary><strong>Appréciation sur les vaccins</strong> — topic <code>Vaccins</code>, 1 codes</summary>

- Appréciation sur les vaccins

</details>

<details>
<summary><strong>Appréciation concernant la réponse</strong> — topic <code>Réponse</code>, 4 codes</summary>

- Appréciations des informations reçues
- Appréciations de l’accès à la réponse
- Appréciations de la qualité de la réponse
- Appréciations de l’engagement communautaire

</details>

<details>
<summary><strong>Appréciation concernant les acteurs de la réponse</strong> — topic <code>Acteurs de la réponse</code>, 5 codes</summary>

- Appréciation concernant la réponse actors
- Appréciations du rôle du gouvernement
- Appréciations du rôle des organisations internationales
- Appréciations du rôle de la Croix-Rouge
- Appréciations du rôle des acteurs locaux

</details>

### Demandes

*12 sous-dimensions, 62 codes*

<details>
<summary><strong>Demande de plus d'informations sur l'épidémie</strong> — topic <code>Maladie</code>, 8 codes</summary>

- Demande de partage de preuves de la maladie ou de personnes malades
- Demande de partage d’informations sur l’agent infectieux (virus, bactérie)
- Demande de partage d’informations sur les symptômes
- Demande de partage d’informations sur la transmission
- Demande de partage d’informations sur le risque et la gravité
- Demande de partage d’informations sur le rétablissement
- Demande de soins de santé urgents
- Demande de vérification de cas suspects

</details>

<details>
<summary><strong>Demande relative à l'épidémie</strong> — topic <code>Épidémie</code>, 3 codes</summary>

- Demande d’informations sur l’évolution de l’épidémie
- Demande de partage du nombre ou des tendances de l’épidémie
- Demande d’arrêt de l’épidémie

</details>

<details>
<summary><strong>Demande liée à l'impact de l'épidémie</strong> — topic <code>Impact de l'épidémie</code>, 1 codes</summary>

- Demande d’atténuation de l’impact économique

</details>

<details>
<summary><strong>Demande d'aide liée aux comportements</strong> — topic <code>Comportements</code>, 1 codes</summary>

- Demande d'aide liée aux comportements

</details>

<details>
<summary><strong>Demande relative aux MSP</strong> — topic <code>MSP</code>, 8 codes</summary>

- Demande de prévention sanitaire concernant la maladie
- Demande de partager plus d’informations sur les mesures préventives
- Demande de supports éducatifs, y compris des flyers et affiches
- Demande de renforcer les MHP (mesures d’hygiène et de prévention)
- Demande d’adapter les MHP
- Demande relative aux tests
- Demande relative au traçage des contacts
- Demande relative aux enterrements sûrs et dignes

</details>

<details>
<summary><strong>Demande concernant les traitements</strong> — topic <code>Traitement</code>, 6 codes</summary>

- Demande de construction de centres de traitement
- Demande d’amélioration de l’accès aux centres de traitement
- Demande de fournir un traitement approuvé
- Demande de rendre le traitement disponible
- Demande d’information sur l’utilisation des traitements
- Demande de fournir un traitement non approuvé

</details>

<details>
<summary><strong>Demande relative aux services de santé</strong> — topic <code>Services de santé</code>, 9 codes</summary>

- Demande d’amélioration de l’accès aux soins de santé
- Demande relative aux coûts des soins de santé
- Demande de soutien médical pour un groupe spécifique
- Demande d’amélioration de la qualité des soins de santé
- Demande de création de morgue
- Demande de service spécifique de la part des agents de santé
- Demande de formation pour les agents de santé communautaires
- Demande de formation pour les agents de santé
- Demande de davantage d’équipements de protection pour les agents de santé

</details>

<details>
<summary><strong>Demande relative aux vaccins</strong> — topic <code>Vaccins</code>, 3 codes</summary>

- Demande de démarrer la campagne de vaccination
- Demande de partager plus d’informations sur le vaccin
- Demande d’arrêter la vaccination

</details>

<details>
<summary><strong>Demande relative à la réponse</strong> — topic <code>Réponse</code>, 20 codes</summary>

- Demande relative aux informations reçues sur la réponse
- Demande de plus d’informations sur les activités de réponse
- Demande d’amélioration de la campagne de communication
- Demande de traitement des rumeurs
- Demande relative à l’accès à l’assistance
- Demande d’assistance sanitaire
- Demande d’autres équipements de prévention
- Demande de cibler les activités pour des groupes spécifiques
- Demande de partage d’informations sur le processus de sélection
- Demande relative à la qualité de la réponse
- Demande de former ou d’impliquer certaines personnes ou institutions
- Demande d’accélérer l’action
- Demande d’améliorer l’efficacité de l’action
- Demande relative à l’engagement communautaire
- Demande de localisation de la réponse
- Demande de participation à l’action
- Demande relative aux solutions menées par les communautés
- Demande d’arrêt de la réponse
- Demande d’être laissé tranquille
- Demande de quitter la zone

</details>

<details>
<summary><strong>Demande relative aux acteurs de la réponse</strong> — topic <code>Acteurs de la réponse</code>, 1 codes</summary>

- Demande relative aux acteurs de la réponse

</details>

<details>
<summary><strong>Demande liée à la sûreté et la sécurité</strong> — topic <code>Sûreté et sécurité</code>, 1 codes</summary>

- Demande liée à la sûreté et la sécurité

</details>

<details>
<summary><strong>Autre demande</strong> — topic <code>Autres</code>, 1 codes</summary>

- Autre demande

</details>

### Signalements et préoccupations

*9 sous-dimensions, 31 codes*

<details>
<summary><strong>Signalement sur la maladie</strong> — topic <code>Maladie</code>, 2 codes</summary>

- Signalement concernant un point d’eau contaminé
- Signalement concernant une personne qui propage la maladie

</details>

<details>
<summary><strong>Signalement lié à l'impact de l'épidémie</strong> — topic <code>Impact de l'épidémie</code>, 5 codes</summary>

- Signalement sur la stigmatisation associée à la maladie
- Signalement sur la stigmatisation ethnique ou culturelle
- Signalement sur la stigmatisation des agents de santé
- Signalement sur la stigmatisation de la famille ou des cas contacts
- Signalement sur la stigmatisation des personnes guéries

</details>

<details>
<summary><strong>Signalement lié aux MSP</strong> — topic <code>MSP</code>, 1 codes</summary>

- Signalement relatif aux restrictions

</details>

<details>
<summary><strong>Signalement lié au traitement</strong> — topic <code>Traitement</code>, 2 codes</summary>

- Signalement sur le centre de traitement
- Signalement sur les effets secondaires du traitement approuvé

</details>

<details>
<summary><strong>Signalement sur les services de santé</strong> — topic <code>Services de santé</code>, 2 codes</summary>

- Plaintes concernant le manque d’informations sur la personne décédée
- Plaintes concernant les agents de santé

</details>

<details>
<summary><strong>Signalement sur la vaccination</strong> — topic <code>Vaccins</code>, 4 codes</summary>

- Plaintes concernant l’accès au vaccin
- Plaintes concernant les infrastructures de vaccination
- Plaintes concernant le prix du vaccin
- Plaintes concernant l’indisponibilité du vaccin

</details>

<details>
<summary><strong>Signalement sur la réponse</strong> — topic <code>Réponse</code>, 2 codes</summary>

- Plaintes concernant l’information reçue
- Signalement sur l’accès à l’assistance

</details>

<details>
<summary><strong>Signalement sur les acteurs de la réponse</strong> — topic <code>Acteurs de la réponse</code>, 5 codes</summary>

- Plaintes concernant les acteurs de la réponse
- Plaintes concernant le rôle du gouvernement
- Plaintes concernant le rôle des organisations internationales
- Plaintes concernant le rôle de la Croix-Rouge
- Plaintes concernant le rôle des acteurs locaux

</details>

<details>
<summary><strong>Signalement sur la sécurité et la sûreté</strong> — topic <code>Sûreté et sécurité</code>, 8 codes</summary>

- Signalement concernant le manque de sécurité
- Signalement concernant des routes ou accès dangereux
- Signalement d’incidents violents
- Signalement de violences basées sur le genre
- Signalement de discrimination
- Signalement de manifestations violentes
- Signalement de violences communautaires
- Signalement d’incidents liés aux forces de sécurité

</details>

### Allégations et Signalement d’incident

*2 sous-dimensions, 7 codes*

<details>
<summary><strong>Allegations concernant la réponse</strong> — topic <code>Réponse</code>, 1 codes</summary>

- Allégations concernant la réponse

</details>

<details>
<summary><strong>Allégations concernant  les acteurs de la réponse</strong> — topic <code>Acteurs de la réponse</code>, 6 codes</summary>

- Allégations de mauvaise conduite en général
- Allégations de consommation illégale
- Allégations de corruption, de pot-de-vin, de vol ou de fraude
- Allégations de népotisme
- Allégations de discrimination, d’abus ou d’atteinte à la dignité
- Allégations d’exploitation et d’abus sexuels (EAS)

</details>

