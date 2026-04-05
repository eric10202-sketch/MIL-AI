# Repository Content Structure

This repository uses a four-zone content model:

1. references/
- Methodology-only assets and reusable guidance.

2. templates/
- Blank project-agnostic templates used by generators.

3. active-projects/
- Current engagement files and generated deliverables.

4. archive/
- Closed engagement artifacts.

## Guardrail
Reference files are methodology-only. Never copy reference project facts into active engagement deliverables.

## Migration Guidance (Safe First Pass)
- Start by placing new files into the correct zone.
- Migrate old files incrementally to avoid breaking scripts.
- Validate generator paths after each move.

## Versioning Convention
- All newly created project folders and deliverable bundles start at `v1.0`.
- Minor approved changes increment only the minor digit: `v1.1`, `v1.2`, `v1.3`.
- Major approved changes or rebaselines increment the major digit and reset the minor digit: `v2.0`, `v3.0`.
- Do not overwrite the prior approved version; create a new versioned project folder instead.
- Preferred folder naming for new work: `{Project Name} v1.0`, `{Project Name} v1.1`, `{Project Name} v2.0`.
- Legacy unversioned folders may remain in place until they are explicitly migrated.

## Active Inventory Snapshot
- Active project folders currently include Eagle, Trinity-CAM (Sonnet), Trinity-CAM (GPT), Trinity-CAM (GPT) v1.1, and Zebra under active-projects/.
- The latest Trinity-CAM (GPT) change-request deliverables are stored in active-projects/Trinity-CAM (GPT) v1.1/.
- The project-specific schedule generator for this engagement is generate_trinity_cam_gpt_schedule.py at the repository root and writes the change-request outputs into the versioned folder above.
