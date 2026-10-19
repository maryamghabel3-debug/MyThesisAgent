* ============================================================================.
* Official scoring syntax for the three research instruments (v1.0, 2026-10-03).
* Key source: data/scoring_sources/official_scoring_keys.json
* Assumes item variables: STAI A01-A20 (1-4), YSQ S01-S75 (1-6), Grable R01-R13 (1-4).
* ============================================================================.

* ---- 1. STAI-Y2: reverse 7 items (1,6,7,10,13,16,19); formula: 5 - score ----.
RECODE A01 A06 A07 A10 A13 A16 A19 (1=4) (2=3) (3=2) (4=1) INTO A01R A06R A07R A10R A13R A16R A19R.
EXECUTE.

* ---- 2. STAI-Y2 total (range 20-80) ----.
COMPUTE STAI_Total = SUM(A02, A03, A04, A05, A08, A09, A11, A12, A14, A15, A17, A18, A20,
  A01R, A06R, A07R, A10R, A13R, A16R, A19R).
EXECUTE.
VARIABLE LABELS STAI_Total 'STAI-Y2 trait anxiety total (20-80)'.
EXECUTE.

* ---- 3. YSQ-SF: 15 subscales x 5 items (no reversed items; each range 5-30) ----.
COMPUTE S_ED = SUM(S01, S16, S31, S46, S61).
COMPUTE S_AB = SUM(S02, S17, S32, S47, S62).
COMPUTE S_MA = SUM(S03, S18, S33, S48, S63).
COMPUTE S_SI = SUM(S04, S19, S34, S49, S64).
COMPUTE S_DS = SUM(S05, S20, S35, S50, S65).
COMPUTE S_FA = SUM(S06, S21, S36, S51, S66).
COMPUTE S_DI = SUM(S07, S22, S37, S52, S67).
COMPUTE S_VH = SUM(S08, S23, S38, S53, S68).
COMPUTE S_EM = SUM(S09, S24, S39, S54, S69).
COMPUTE S_SU = SUM(S10, S25, S40, S55, S70).
COMPUTE S_SS = SUM(S11, S26, S41, S56, S71).
COMPUTE S_EI = SUM(S12, S27, S42, S57, S72).
COMPUTE S_US = SUM(S13, S28, S43, S58, S73).
COMPUTE S_ET = SUM(S14, S29, S44, S59, S74).
COMPUTE S_IS = SUM(S15, S30, S45, S60, S75).
EXECUTE.

* ---- 4. YSQ-SF total (range 75-450) ----.
COMPUTE YSQ_Total = SUM(S01 TO S75).
EXECUTE.
VARIABLE LABELS YSQ_Total 'YSQ-SF total (75-450)'.
EXECUTE.

* ---- 5. Grable & Lytton total: 13 items x 1-4 (range 13-52) ----.
COMPUTE FRT_Total = SUM(R01, R02, R03, R04, R05, R06, R07, R08, R09, R10, R11, R12, R13).
EXECUTE.
VARIABLE LABELS FRT_Total 'Grable-Lytton risk tolerance total (13-52)'.
EXECUTE.
