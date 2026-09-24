# 🧠 AGENTS.md — DOCX TABLE ENGINE

## 🎯 PURPOSE

This project processes DOCX documents with complex table structures, including nested tables.

The system must remain:
- stable
- deterministic
- backward-compatible
- resistant to document structure changes

---

# 🚨 CORE RULES (MANDATORY)

1. NEVER use `doc.tables` directly
2. ALWAYS use `TableManager` for table access
3. ALWAYS support nested tables via XPath
4. DO NOT rely on table indices alone
5. ALWAYS implement fallback via table signatures
6. DO NOT break existing Excel → Word mapping
7. DO NOT silently ignore errors — log them

---

# 🧱 ARCHITECTURE PRINCIPLES

## 1. Table Access Layer

All table operations must go through:

- `TableManager`
- No exceptions

---

## 2. Table Identification

Each table must be identifiable by:

- shape (rows x columns)
- preview text
- hash (md5)

Indexes are NOT reliable identifiers.

---

## 3. Fallback Strategy

If table lookup by index fails:

1. Attempt match by signature hash
2. If not found:
   - try shape + preview match
3. If still not found:
   - log warning
   - skip safely

---

## 4. Change Tracking

All modifications MUST be logged:

- table index
- row / column
- old value
- new value
- timestamp

Use `ChangeLogger`.

---

## 5. Comparison Rules

When comparing documents:

1. Match tables by:
   - index (primary)
   - signature (fallback)
2. Compare cell values after normalization:
   - strip whitespace
3. Ignore false diffs caused by structure shifts

---

## 6. Nested Tables

Nested tables are FIRST-CLASS citizens.

The system MUST:
- detect them
- process them
- compare them

---

# ⚙️ DEVELOPMENT RULES

## DO

- Write clear, deterministic code
- Keep functions pure when possible
- Handle edge cases explicitly
- Add debug modes where useful

## DO NOT

- Introduce hidden side effects
- Break existing APIs
- Change behavior without reason
- Hardcode table indices

---

# 🧪 TESTING STRATEGY

Every change MUST pass:

1. Baseline test:
   - Generated document == reference document

2. Nested table test:
   - Add nested table → system still works

3. Mapping integrity:
   - Excel mapping still applies correctly

---

# 🔍 DEBUGGING TOOLS

Always available:

- `print_table_map()` → inspect structure
- signature comparison
- visual diff

---

# 🚀 PERFORMANCE RULES

- Avoid repeated XPath scans
- Cache table signatures when possible
- Minimize nested loops over tables

---

# 📦 FILE RESPONSIBILITIES

## table_manager.py
- Single source of truth for table access

## data_filler_v2.py
- Applies Excel mapping to document

## compare_documents.py
- Compares two documents reliably

## table_mapper.py
- Matches tables across documents

## change_logger.py
- Tracks all changes

## visual_diff.py
- Outputs human-readable differences

## check_report.py
- Preflight checks before filling (files, table titles vs Word)
- ЧТО_ПРОВЕРИТЬ.xlsx and highlighted Бюллетень_ПРОВЕРКА.docx for the economist
- Must never change Бюллетень_ГОТОВЫЙ.docx

---

# ⚠️ FAILURE HANDLING

If something goes wrong:

- DO NOT crash immediately
- Log error
- Attempt fallback
- Continue processing where safe

---

# 🧠 AGENT BEHAVIOR

When modifying code:

1. Read existing implementation
2. Preserve intent
3. Apply minimal safe changes
4. Verify output
5. Only then proceed further

---

# 🏁 FINAL RULE

Stability > cleverness

Correctness > performance

Determinism > assumptions