# Sb Mineral Classification — implemented requirements and safeguards

These requirements record the user's instructions of 2026-11-12. The user
subsequently authorized implementation in Sb Mineral Classification.py. The
primary ranges below are preserved exactly; additional safeguards are listed
explicitly below for testing against real samples.

## Verified authoritative reference (recovered 2026-11-12)

- User re-supplied `AuAgTe-Code-Handoff.zip`; all 16 Python files matched its
  manifest and passed syntax checks. All seven app script constants and all
  eight current A2 dropdown scripts are present.
- `Updated-AllAgMinerals-41-19.py`: 419 lines, SHA-256
  `256d65756f308da710119eca203d35d542aa87ef1ff0de3f94c449f8d13527b6`.
- The archive, manifest and earlier tested summary copy are preserved under
  `reference-baselines/2026-11-12/`. The new archive is authoritative for active
  script paths initially. The user subsequently restored the tested enhanced
  Summary_Each_Sample.py as the sole current Summary authority.
- Recovery branch: `codex/auagte-authoritative-2026-11-12`. A local commit alone
  is not a persistence guarantee; confirm remote storage before claiming backup.

## Confirmed behavior to inherit/adapt

- GUI selects the Full Analysis workbook. The first input sheet supplies data.
- An Area Size Cutoff GUI prompt accepts a nonnegative value in square micrometres.
  Initial value 0 disables size filtering; cancelling also uses 0.
- In the reference, elemental eligibility is checked before size routing.
  Eligible features with `Area < cutoff` go to `After Size-CutOff`; equality
  passes. No fixed size-fraction presets or automatic diameter threshold exist
  in this particular reference. Do not substitute Tellurides' size logic.
- Reference eligibility is Ag >1.4 OR Au >1.4. This is element-specific and
  must be adapted for Sb and the requested Sb-free cinnabar, not copied literally.
- Non-empty category sheets only, a separate blank Area sheet, full Raw Data,
  Summary, Integrity Check, original feature/coordinate columns, and Excel-safe
  sheet names. The reference's output name is Classified<original filename>.
- Summary lists category names but intentionally leaves area and percentage
  cells blank for later manual microscope-image corrections. Do not replace
  these blanks with calculated totals in the new classifier.
- Integrity Check includes the actual area cutoff and compares eligible raw
  count/area with all output category sheets, including After Size-CutOff.
  The Sb version must also account for the new W/La exclusions consistently.
- The reference handles missing/NaN optional element values through v(); it
  validates required Ag and area columns. Required-element validation must be
  adapted to the Sb workflow rather than assuming Ag is always required.
- Preserve existing column highlight colors, Summary styling and standalone
  execution. Do not modify the reference itself or the app.
- W >10 and La >7 are explicit new user requirements for this classifier;
  they are not implemented in the verified Updated-AllAg reference.

## Scope

- Eventual script: `Ag-Liberation/Sb Mineral Classification.py`.
- Independently runnable. Existing A2 Refresh scripts discovers it without
  application modifications.
- Preserve the verified reference's non-classification behavior, adapting
  element-specific selection and integrity checks for Sb. The verified manual area cutoff
  applies; do not introduce automatic fraction presets.
- All ranges below are wt.%; listed range endpoints are inclusive.
- These primary ranges require additional mineral-specific exclusion conditions;
  ranges alone are not a finished classification specification.

## Free categories — user-specified ranges

| Category | Element ranges |
| --- | --- |
| Gudmundite | Sb 44–65; Fe 17–32; S 5–20 |
| Aurostibite | Au 35–51; Sb 42–65 |
| Stibnite | Sb 58–78; S 14–38 |
| Boulangerite | Pb 42–64; Sb 16–34; S 6–28 |
| Cupropolybasite | Ag 30–58; Cu 13–28; Sb 4–18; S 7–23 |
| Cinnabar | Hg 60–95; S 4–20 |
| Ag/Ag-Hg-bearing tetrahedrite | Ag 2–45; Cu 15–45; Sb 15–35; S 10–31; Hg 0–25 |
| Tetrahedrite-(Fe) | Cu 22–45; Sb 20–37; S 15–31; Fe 1–10 |
| Andorite | Ag 4–18; Pb 14–31; Sb 30–48; S 12–28 |
| Zn-bearing tetrahedrite | Cu 22–45; Sb 20–37; S 15–31; Zn 2–11; Fe 0–8 |
| Bournonite | Pb 30–46; Sb 12–30; S 7–24; Cu 4–19 |
| Pyrargyrite | Ag 30–65; Sb 14–29; S 7–22 |

## Association categories — user-specified ranges

| Category | Element ranges |
| --- | --- |
| Gudmundite-Asso | Sb 15–60; Fe 7–65; S 3–45 |
| Aurostibite-Asso | Au 4–51; Sb 10–62 |
| Stibnite-Asso | Sb 15–77; S 5–48 |
| Boulangerite-Asso | Pb 15–65; Sb 5–31; S 4–40 |
| Cupropolybasite-Asso | Ag 4–56; Cu 5–40; Sb 4–18; S 4–40 |
| Cinnabar-Asso | Hg 20–90; S 4–40 |
| Andorite-Asso | Ag 2–17; Pb 5–50; Sb 10–47; S 4–40 |
| Tetrahedrite-(Fe)-Asso | Cu 8–43; Sb 5–35; S 5–45; Fe 1–45 |
| Bournonite-Asso | Pb 12–46; Sb 5–30; S 4–24; Cu 3–19 |
| Pyrargyrite-Asso | Ag 4–65; Sb 5–29; S 4–22 |

The last association entry was written as "Pyrargyrite" in the user's message;
the working interpretation is Pyrargyrite-Asso because it appears in the
association list. No separate association categories were requested for the
broad Ag/Ag-Hg-bearing or Zn-bearing tetrahedrite categories.

## Exclusions and final routing

1. Exclude W >10 wt.% or La >7 wt.% from every mineral and leftover category.
   Equality at W=10 or La=7 is not excluded by these tests.
2. Apply the verified manual area cutoff and route eligible undersized features
   to After Size-CutOff. Do not assume a global Ag requirement or invent fraction presets.
3. Evaluate free-mineral conditions before association conditions.
4. Apply final leftovers only after named-mineral evaluation:
   - Au-Sb Leftover: Sb >2 and Au >2.
   - Ag-Sb Leftover: Sb >2 and Ag >2, if not already assigned.
   - Sb Leftover: Sb >2, if not already assigned.
5. A grain satisfying both Au-Sb and Ag-Sb leftovers goes to Au-Sb under the
   listed precedence, unless the user changes it.
6. Cinnabar must be considered without requiring Sb, because HgS has no Sb.
7. One output category per feature. Excluded features remain traceable in raw
   data and integrity counts. Input workbooks must not be modified.

## Discrimination approach

- Use mineral-specific non-member-element caps, following the verified
  reference's approach. Do not invent a single cap for every mineral.
- Aurostibite: sulphur and unrelated base-metal caps are needed to separate
  clean Au-Sb compositions from sulphide associations.
- Stibnite versus gudmundite: Fe discrimination is essential. Their user-defined
  primary ranges overlap at Sb 58–65 and S 14–20.
- Boulangerite versus bournonite: Cu discrimination is essential. Primary ranges
  overlap at Pb 42–46, Sb 16–30 and S 7–24.
- Pyrargyrite versus cupropolybasite: Cu discrimination is essential; the Sb/Ag/S
  windows overlap. Do not impose a Cu cap on a category that requires Cu.
- Cupropolybasite versus Ag-rich tetrahedrite: the supplied major-element
  windows overlap even with Cu present. Additional elemental proportions or
  an explicit ambiguity policy are needed; rule order alone is insufficient.
- Tetrahedrite subcategories: propose Ag-bearing first, then Zn-bearing, then
  ordinary Fe tetrahedrite, with mutually exclusive additional tests. This
  priority is implemented with Ag >=2 before Zn >=2; remaining ordinary
  tetrahedrite requires Ag <1.5 and Zn <2.
  Under an Ag >=2 priority, grains with Ag 3–4% enter the Ag-bearing category.
- Association envelopes overlap especially strongly. Au, Hg, Pb, Cu and Ag
  safeguards must be mineral-specific so a broad association rule does not
  capture another named mineral. For genuinely ambiguous mixtures, retain
  candidate/review information rather than claiming a unique mineral identity.
- Additional candidate elements As, Bi, Te, Se, O, Si, Al, Ca and Mg need
  appropriate treatment for substitution, oxidation, matrix contribution and
  interference. Missing optional element columns use the reference
  zero-default behavior, with an explicit console warning and a record in
  Integrity Check. Sb and area columns are required.

## Validation required before delivery

- Retain the verified reference provenance and non-classification-rule audit.
- Boundary tests for every user-specified range and strict exclusion threshold.
- Representative grains, known competitors and deliberate multi-match cases.
- No mineral or association is swallowed by a premature leftover rule.
- Equal eligible raw/output row counts and areas with separate excluded counts.
- Standalone execution, syntax checking, safe Excel sheet names and unique
  sheet-name mapping (some requested display labels exceed Excel's 31 characters).


## Implemented processing and output

- Eligibility: Sb >2 wt.% or a cinnabar/cinnabar-association match that passes
  its safeguards. There is no global minimum Ag or Au requirement.
- W >10 or La >7 never enters any category, including After Size-CutOff.
- Non-numeric, negative or non-finite reported compositions, and invalid area,
  are kept in Raw Data with Invalid data status, without mineral assignment.
- All free rules are evaluated before association rules. Exactly one match
  assigns the named category. Multiple matches within a stage go to the
  appropriate leftover, with candidate names and a review note in Raw Data.
- Categories and After Size-CutOff retain the original data columns. Raw Data
  adds classification/status/candidate/note audit columns. No review sheet is
  added that a downstream master script could mistake for a mineral category.
- Original workbooks are not saved or modified. Output is Classified<input
  stem>.xlsx, with _2, _3, etc. when an output already exists.
- Summary area/percentage cells remain blank; Integrity Check reconciles
  eligible counts/areas including After Size-CutOff, and accounts for excluded,
  invalid and outside-scope rows separately.
- The full Ag/Hg tetrahedrite label maps to the Excel-safe sheet name
  `Ag-AgHg Tetrahedrite (Fe)`; Summary and raw audit retain the full label.
- Association labels are compositional screening assignments. Images and
  actual samples are needed to validate mineral associations and liberation.

## Additional checks used in the initial implementation

All caps are strict (<), and all ratio bounds are inclusive unless stated.
Primary wt.% intervals above remain unchanged. Ratios use atomic proportions,
not direct wt.% ratios. These supplemental conditions can reject a candidate
that passes its primary ranges; the feature then reaches another named rule or
one of the requested leftovers.

- Clean matrix caps: O <15 and Si, Al, Ca, Mg each <8. These do not apply to
  association categories.
- Cupropolybasite (free and Asso): (Ag+Cu)/(Sb+As) atomic ratio >4.
- Free tetrahedrite families: (Ag+Cu)/(Sb+As) atomic ratio 2–4.
- Cupropolybasite, pyrargyrite and free tetrahedrite: Sb atomic amount >=As.
- Free boulangerite: Pb/(Sb+As) atomic ratio 0.9–1.8, to reduce assignment of
  jamesonite/geocronite-like compositions to the clean boulangerite category.
- Free pyrargyrite: Ag/(Sb+As) atomic ratio 2–4; Asso uses 1.5–4.2.
- Aurostibite-Asso: Au/Sb atomic ratio 0.1–1.0 to prevent every Au+Sb mixture
  from automatically receiving this named association.

## User update: precious-element caps and workbook order

- Au and Ag each must be <1.5 wt.% wherever that element is not one of the
  category's primary required elements. This applies to all free and association
  rules. Aurostibite still requires Au; Ag minerals still require Ag.
- Other former impurity caps below 8 are relaxed to <8, except these diagnostic
  boundaries: free boulangerite Cu <4; boulangerite-Asso Cu <3; ordinary
  tetrahedrite Zn <2; stibnite-Asso Fe <7; pyrargyrite-Asso Cu <5.
- Existing allowances above 8 for oxygen, substitution and association hosts
  remain unchanged. The initial minor-element caps are therefore not a blanket
  maximum for elements that are part of the required mineral or its host.
- Primary ranges, atomic-ratio checks, W/La exclusion, area cutoff, stages and
  rule evaluation sequence remain unchanged. Ag at 1.5–<2 fails the ordinary
  tetrahedrite cap and does not meet the Ag-bearing category's minimum; it is
  retained in the existing appropriate leftover, not forced into a named group.
- Workbook category order (and Summary order), independently of rule evaluation:
  Aurostibite; Ag/Ag-Hg-bearing tetrahedrite; Andorite; Cupropolybasite;
  Pyrargyrite; Gudmundite; Stibnite; Tetrahedrite-(Fe); Zn-bearing tetrahedrite;
  Boulangerite; Bournonite; Cinnabar. Then the existing association categories
  in the corresponding order, followed by leftovers and After Size-CutOff.
- Categories with no rows are omitted. Existing Area/Summary/Raw Data/Integrity
  Check worksheet placement remains unchanged.

The following tables are generated from the exact rules in the delivered script.

### Free caps

| Category | Additional element caps (wt.%) | Ratio checks |
| --- | --- | --- |
| Gudmundite | O <15; Si <8; Al <8; Ca <8; Mg <8; Cu <8; Pb <8; Ag <1.5; Au <1.5; Hg <8; Bi <8; Te <8; Se <8; As <8 |  |
| Aurostibite | O <15; Si <8; Al <8; Ca <8; Mg <8; S <8; Fe <8; Cu <8; Pb <8; Ag <1.5; Hg <8; Bi <8; Te <8; Se <8; As <8 |  |
| Stibnite | O <15; Si <8; Al <8; Ca <8; Mg <8; Fe <8; Pb <8; Cu <8; Ag <1.5; Au <1.5; Hg <8; Bi <8; As <8; Te <8; Se <8 |  |
| Boulangerite | O <15; Si <8; Al <8; Ca <8; Mg <8; Cu <4; Ag <1.5; Fe <8; Au <1.5; Hg <8; Bi <8; As <8; Te <8; Se <8 | Boulangerite ratio |
| Cupropolybasite | O <15; Si <8; Al <8; Ca <8; Mg <8; Pb <8; Fe <8; Hg <8; Au <1.5; Bi <8; Te <8; Se <8 | Sb dominant over As, Cupropolybasite ratio |
| Cinnabar | O <15; Si <8; Al <8; Ca <8; Mg <8; Ag <1.5; Cu <8; Fe <8; Pb <8; Au <1.5; Sb <8; Te <8; Se <8; Bi <8; As <8 |  |
| Ag/Ag-Hg bearing Tetrahedrite (Fe) | O <15; Si <8; Al <8; Ca <8; Mg <8; Pb <8; Au <1.5; Bi <8; Te <8; Se <8; Fe <12; Zn <12 | Sb dominant over As, Tetrahedrite ratio |
| Tetrahedrite-(Fe) | O <15; Si <8; Al <8; Ca <8; Mg <8; Pb <8; Au <1.5; Bi <8; Te <8; Se <8; Ag <1.5; Hg <8; Zn <2 | Sb dominant over As, Tetrahedrite ratio |
| Andorite | O <15; Si <8; Al <8; Ca <8; Mg <8; Cu <8; Fe <8; Au <1.5; Hg <8; Bi <8; As <8; Te <8; Se <8 |  |
| Zn-bearing Tetrahedrite | O <15; Si <8; Al <8; Ca <8; Mg <8; Pb <8; Au <1.5; Bi <8; Te <8; Se <8; Ag <1.5; Hg <8 | Sb dominant over As, Tetrahedrite ratio |
| Bournonite | O <15; Si <8; Al <8; Ca <8; Mg <8; Ag <1.5; Fe <8; Au <1.5; Hg <8; Bi <8; As <8; Te <8; Se <8 |  |
| Pyrargyrite | O <15; Si <8; Al <8; Ca <8; Mg <8; Cu <8; Pb <8; Fe <8; Au <1.5; Hg <8; Bi <8; Te <8; Se <8 | Sb dominant over As, Pyrargyrite ratio |

### Association caps

| Category | Additional element caps (wt.%) | Ratio checks |
| --- | --- | --- |
| Gudmundite-Asso | As <12; Bi <30; Te <8; Se <8; Cu <8; Pb <8; Ag <1.5; Au <1.5; Hg <10 |  |
| Aurostibite-Asso | As <12; Bi <30; Te <8; Se <8; Ag <1.5; Cu <25; Pb <30; Hg <20; Fe <65 | Aurostibite association ratio |
| Stibnite-Asso | As <12; Bi <30; Te <8; Se <8; Fe <7; Cu <8; Pb <8; Ag <1.5; Au <1.5; Hg <10 |  |
| Boulangerite-Asso | As <12; Bi <30; Te <8; Se <8; Cu <3; Ag <1.5; Fe <12; Au <1.5; Hg <10 |  |
| Cupropolybasite-Asso | As <12; Bi <30; Te <8; Se <8; Pb <8; Fe <25; Au <1.5; Hg <10 | Sb dominant over As, Cupropolybasite ratio |
| Cinnabar-Asso | As <12; Bi <30; Te <8; Se <8; Ag <1.5; Cu <15; Pb <20; Fe <35; Au <1.5; Sb <10 |  |
| Andorite-Asso | As <12; Bi <30; Te <8; Se <8; Cu <8; Fe <25; Au <1.5; Hg <10 |  |
| Tetrahedrite-(Fe)-Asso | As <12; Bi <30; Te <8; Se <8; Ag <1.5; Pb <8; Zn <15; Hg <10; Au <1.5 |  |
| Bournonite-Asso | As <12; Bi <30; Te <8; Se <8; Ag <1.5; Fe <25; Au <1.5; Hg <10 |  |
| Pyrargyrite-Asso | As <12; Bi <30; Te <8; Se <8; Cu <5; Pb <8; Fe <30; Hg <10; Au <1.5 | Sb dominant over As, Pyrargyrite association ratio |

## Validation performed

Sixteen unittest cases cover all 22 named examples; independently transcribed
primary ranges and endpoints; all W/La and size boundaries; leftover precedence;
Ag/Zn tetrahedrite boundaries; representative competing-mineral cases;
ambiguity handling; invalid data; all output categories and sheet names;
blank Summary totals; input-file hashes; unique output naming; empty inputs;
required columns; and execution through the existing app runner with GUI
selections mocked. Added cases verify strict Au/Ag boundaries for every
relevant rule, impurity caps and diagnostic exceptions, unchanged rule
evaluation order, exact workbook/Summary ordering, and omission of empty
categories. These are software checks, not validation on real lab data.
