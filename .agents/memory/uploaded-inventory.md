---
name: Inventory source of truth
description: The LabVault catalog is grounded in the user's uploaded equipment and consumables workbooks.
---

The uploaded master-stock and consumables workbooks are the authoritative starting point for the LabVault catalog. New inventory records should only come from those source files or explicit lab-assistant additions.

**Why:** The user specifically supplied the real laboratory inventory and asked that component names and quantities not be invented.

**How to apply:** Preserve workbook identifiers where available, import idempotently, and label any manually added item as an explicit lab-assistant addition.