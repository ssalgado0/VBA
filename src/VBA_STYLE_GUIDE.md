# VBA Macro Style Guide

This document defines the coding, documentation, and language conventions
followed by all VBA macros stored in the `VBA/` folder of this repository.

The objective of this guide is to ensure consistency, readability, robustness,
and long-term maintainability of VBA automation scripts.

---

## 1. General Principles

- The primary focus of these macros is practical functionality and problem-solving.
  As a result, some stylistic conventions and structural refinements may not be
  applied consistently across all scripts.

- Functionality has been prioritized over extensive error handling, as these macros
  were developed incrementally during available time within a professional work
  environment, focusing on the automation and optimization of daily work tasks.

- All macros are self-documented through structured headers and comments.

- Language usage (English / Spanish) follows the next rules:
  - **ENGLISH** → macro headers, code comments, debugging output, and most variables
  - **SPANISH** → user interface elements (`MsgBox`, worksheet messages, etc.) and
    business-domain variable names when appropriate

---

## 2. Macro Header (Initial Description)

Each macro **must start with a structured header comment block**.

The header includes:
- Macro name
- High-level description of its purpose
- Main processing steps

### Header Example

```vb
'------------------------------------------------------------
' Macro: ExampleMacro
' Description:
'   Reads input data from the worksheet and processes each
'   record by querying an external service.
'
'   The macro validates the retrieved data using rule-based
'   checks and produces structured output for further review.
'
' Author: name@example.com
'------------------------------------------------------------
