# Production Collection Flow

## Data contract

| Template metric | Authoritative source | Value |
| --- | --- | --- |
| `건수` | Configured daily-detail endpoint | Displayed transaction count |
| `총매출` | Configured daily-detail endpoint | Displayed gross sales |

## Safety rules

- Blank values in a valid daily response are written as `-`.
- Explicit numeric zero remains numeric zero.
- Rows marked `중단` or `폐점` are not queried and receive `-`.
- Unmatched names are skipped rather than written to an inferred row.
- Formula cells are write-blocked.

## Write procedure

1. Resolve the target month sheet, store row, date column, and writable cells.
2. Preview the changes from the service response.
3. Require explicit confirmation before writing.
4. Copy the source workbook and update only the approved cells.
5. Validate values, formulas, sheet structure, merged ranges, hidden state, and freeze panes.

The source workbook is never modified in place.
