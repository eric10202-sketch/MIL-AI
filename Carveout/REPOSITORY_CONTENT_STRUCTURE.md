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
