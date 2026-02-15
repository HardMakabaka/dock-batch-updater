# PR1 (HIGH RISK): Backup Concurrency Collision Fix

## Problem
When the user selects a global `backup_dir`, multiple same-basename source files (e.g. `report.docx` from different folders) processed concurrently may create backups with colliding names (previously `_backup_1`, `_backup_2`, ...) and overwrite each other. This creates a silent data-loss path and removes the safety net needed for subsequent higher-risk replacements.

## Approach
- Use a new backup naming scheme that is unique under concurrency and still human-readable.
- Make the copy+finalize step more robust by copying to a temp file in the target directory, then finalizing via rename.

### Backup naming
`{stem}_backup_{parentDirHint}_{pathHash6}_{uuid8}{ext}`

- `parentDirHint`: sanitized parent directory name, limited to 20 chars (letters/digits/CJK/_ only). If empty, `_root_`.
- `pathHash6`: first 6 hex chars of SHA-256 of absolute source path.
- `uuid8`: first 8 hex chars of a UUID4.

Example:
`report_backup_财务部_a3f2c1_e7b4d9f0.docx`

### Atomic-ish finalize
- Copy original to `.{backup_name}.tmp` in the backup directory.
- Finalize with `os.replace(tmp, final)`.
- If an extremely unlikely collision occurs, retry generating a new name (max 50 tries).
- Best-effort cleanup of temp files.

## Tests
- Add concurrency regression test for two files with the same basename from different directories into a shared `backup_dir`.
- Existing backup tests are updated to assert uniqueness without depending on the old numeric suffix.

## Acceptance Criteria
- Concurrent backup creation to a global `backup_dir` never overwrites another backup.
- Backup files are unique and present on disk.
- Backup naming remains short enough for typical Windows MAX_PATH constraints.

## Rollback Plan
- Revert this PR.
- Note: backups already created with the new naming scheme are user data; rollback must not delete them.

## Files
- `src/core/docx_processor.py`
- `tests/test_backup_concurrency.py`
- `tests/test_docx_processor_additional.py`
