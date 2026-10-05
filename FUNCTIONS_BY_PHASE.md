# FUNCTIONS BY PHASE

## PHASE 1: DATA CLEANING (Pandas)
- **Status Filter**: Remove rows where Status is "BP".
- **Remark By Filter**: Remove specific excluded names (DCCAUNTE, etc.).
- **Account Formatting**: Add six leading zeros and force as text.
- **Remark Sanitization**: Remove leading "=" to prevent Excel errors.
- **Column Reordering**: Move "Next Call" to its designated index.

## TEMPLATE PHASE (Excel COM Integration)
- **Data Insertion**: Paste cleaned data starting at cell A3.
- **Formula Replication**: Replicate formulas for columns AZ:BL from Row 1.
- **Bulk Clearing**: Clear BF (Reaction) and BG/BH (Amt/When) for non-PTP rows.
- **Batch Row Deletion**: Remove rows where CH CODE is #N/A, ACTION is EXCLUDED, or AMT is 0 (while preserving blanks).
- **SYSTEM Note Swap**: Update Remark By (L) using Collector (O) if Notes 1 (BC) contains "SYSTEM".
- **Notes 2 Truncation**: Limit BD to 250 characters.
- **Character Count (BM)**: Calculate length of truncated Notes 2 and store in Column BM (with header "LENGTH").
- **Final Baking**: Convert range AZ3:BM from formulas to static values.

## BUCKET PHASE (Multi-Campaign Splitting)
- **Campaign Detection**: Split raw data into 120, 150, and 180 groups based on Column BJ (PLACEMENT).
- **Template Validation UI**: Provide visual feedback (✅/❌) for filename pattern, existence of the selected week sheet, and whether cell A2 is currently empty.
- **Target Week Selection**: Allow user to select WEEK 1 through WEEK 5 via UI.
- **Smart Recognition**: If the target sheet already contains data at A2, automatically clear the existing data before pasting to ensure a clean refresh.
- **A2 Alignment**: Paste data starting at A2 (offset from standard A3).
- **Organized Output**: Automatically create a subfolder named `[MONTH] [WEEK]` inside the `Automated` directory to store the resulting campaign reports.
- **Pasting Schema (Starting A2)**: Data is mapped and pasted in the following order:
    1. **LAN**
    2. **DATE**
    3. **NOTES 1**
    4. **NOTES 2**
    5. **ACTION**
    6. **REACTION**
    7. **AMT**
    8. **WHEN**
    9. **REF CODE**
    10. **STATUS**
    11. **ENDO DATE**
    12. **NO OF CHARACTERS**
- **Automated Naming**: Save as `UPDATED PRODUCTIVITY REPORT [Month] [Week] PL[Campaign]` in the new subfolder.

### 🗺️ SOURCE > BUCKET Mapping
| Source (Template Phase Output) | Bucket Column (Target) |
| :--- | :--- |
| **BA** (LAN) | 1. **A** (LAN) |
| **BB** (DATE) | 2. **B** (DATE) |
| **BC** (NOTES 1) | 3. **C** (NOTES 1) |
| **BD** (NOTES 2) | 4. **D** (NOTES 2) |
| **BE** (ACTION) | 5. **E** (ACTION) |
| **BF** (REACTION) | 6. **F** (REACTION) |
| **BG** (AMT) | 7. **G** (AMT) |
| **BH** (WHEN) | 8. **H** (WHEN) |
| **(BLANK)** (REF CODE) | 9. **I** (REF CODE) |
| **BL** (STATUS) | 10. **J** (STATUS) |
| **(BLANK)** (ENDO DATE) | 11. **K** (ENDO DATE) |
| **BM** (LENGTH) | 12. **L** (NO OF CHARACTERS) |
