# MFD_Tools

Distributor-facing utilities. Each top-level subfolder is a self-contained tool with its own README.

| Folder | Description |
|--------|-------------|
| [**Rebalancer**](Rebalancer/) | MF portfolio rebalancer (Python; run from source). See [Rebalancer/README.md](Rebalancer/README.md). |

**Maintainers:** source for `Rebalancer/` is built from the main `MF_rebalancer` development repo with `tools/assemble_mfd_source_distribution.py` (target this folder, optional `--include-nav`).

**Do not** commit client holdings or any client-specific exports to this repository.

### Publish to GitHub (first time)

1. On GitHub: **New repository** → name **`MFD_Tools`** (empty, no README/license).
2. In this folder:

```text
git remote add origin https://github.com/YOUR_USER/MFD_Tools.git
git push -u origin main
```

Use SSH if you prefer: `git@github.com:YOUR_USER/MFD_Tools.git`.
