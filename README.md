# MFD_Tools

Distributor-facing utilities. Each top-level subfolder is a self-contained tool with its own README.

| Folder | Description |
|--------|-------------|
| [**Rebalancer**](Rebalancer/) | MF portfolio rebalancer (Python; run from source). See [Rebalancer/README.md](Rebalancer/README.md). |

**Maintainers:** source for `Rebalancer/` is built from the main `MF_rebalancer` development repo with `tools/assemble_mfd_source_distribution.py` (target this folder, optional `--include-nav`).

**Do not** commit client holdings or any client-specific exports to this repository.
