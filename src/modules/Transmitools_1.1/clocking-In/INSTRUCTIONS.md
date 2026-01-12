# Clocking-In Macros

This folder contains VBA macros related to daily clock-in and clock-out
workflows, including teleworking support.

---

## Modules

- `FichajeEntradaTeletrabajo.bas`  
  Registers daily clock-in time and (optionally) teleworking status.

- `FichajeSalidaTeletrabajo.bas`  
  Registers daily clock-out time.

- `EntornoTeletrabajo.bas`  
  Shared teleworking state and helper procedure to display the UserForm.

---

## Dependencies

These macros depend on the following UserForm:

- **FrmTeletrabajo**
  - Form definition: [`src/forms/clocking-In/FrmTeletrabajo.frm`](../../forms/clocking-In/FrmTeletrabajo.frm)
  - Form resources:  [`src/forms/clocking-In/FrmTeletrabajo.frx`](../../forms/clocking-In/FrmTeletrabajo.frx)

---

## Notes

- Email generation is handled via Outlook and requires user review
  before manual sending.
- File paths and recipients must be adapted to the target environment.
