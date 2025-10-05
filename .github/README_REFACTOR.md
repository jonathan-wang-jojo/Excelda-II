\# Excelda-II Refactor Guide



\## Purpose

Excelda-II is a 2D VBA game implemented in Excel. The goal of this refactor is modernization: 

improve clarity, reduce redundancy, and maintain full functionality.



\## Refactor Rules

1\. \*\*Simplify:\*\* Shorten and reorganize existing code; do not expand.

2\. \*\*Preserve:\*\* Keep gameplay behavior identical.

3\. \*\*Unify:\*\* Merge redundant logic across files.

4\. \*\*Clarify:\*\* Use consistent naming, spacing, and explicit types.

5\. \*\*Stabilize:\*\* Replace global reliance with local scope where possible.



\## File Roles

| File | Purpose |

|------|----------|

| `AA\_GameLoop.bas` | Main loop and update sequence |

| `AG\_LinkActions.bas` | Player movement and actions |

| `AH\_Enemies.bas` | Enemy AI and collision |

| `AJ\_Triggers.bas` | Trigger and event management |



\## Refactoring Priorities

\- Trim unused functions, constants, or variables.

\- Collapse duplicated code paths.

\- Simplify long nested logic with clear if/else blocks.

\- Remove redundant `Call` wrappers.

\- Keep functions short (<100 lines).

\- Replace magic numbers with constants.



\## Notes

\- Always run a full test after each major refactor step.

\- Keep backup of pre-refactor code in `Archive/`.



