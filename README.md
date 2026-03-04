# autodocproduct

**GR:** Παραγωγή τελικών εγγράφων Word (DOCX) από templates με placeholders, με ένα κλικ μέσω VBA macro.  
**EN:** One-click generation of final Word documents (DOCX) from placeholder templates using a VBA macro.

---

## Ελληνικά (GR)

### Τι κάνει
- Διαβάζει τιμές από ένα **Controller Word (.docm)** που περιέχει έναν πίνακα **ΠΕΔΙΟ / ΤΙΜΗ**
- Παίρνει όλα τα templates που ταιριάζουν στο μοτίβο **`TEMPLATE_*.docx`**
- Δημιουργεί φάκελο **`OUTPUT`**
- Παράγει **ξεχωριστές τελικές εκθέσεις** στο `OUTPUT` χωρίς να πειράζει τα templates
- Αντικαθιστά placeholders τύπου `{{Key}}` με τις αντίστοιχες τιμές
- Υπολογίζει **ώρα έναρξης/περάτωσης** ανά έγγραφο:
  - 10’ διάρκεια για όλα
  - 20’ διάρκεια μόνο για “Κατάθεση Αστυνομικού” (ανίχνευση από όνομα αρχείου: περιέχει ΚΑΤΑΘΕΣΗ + ΑΣΤΥΝΟΜ)
  - διάλειμμα `BreakMinutes` (2’/5’ κτλ) μεταξύ εγγράφων

### Γρήγορο Setup (Word 2016 Windows)
1. Άνοιξε το `examples/00_Controller_example.docx` και κάνε **Save As → .docm**.
2. **Alt+F11 → File → Import File…** και κάνε import το `src/ControllerModule_TimeOutput_GR_ANSI.bas`.
3. Βάλε τα δικά σου templates στον ίδιο φάκελο με το `.docm` και ονόμασέ τα `TEMPLATE_*.docx`.
4. Βεβαιώσου ότι στα templates έχεις:
   - `{{OraEnarxis}}` στην αρχή (ώρα έναρξης)
   - `{{OraPeratosis}}` στο τέλος (ώρα περάτωσης)
5. **Developer → Macros → Generate_Reports_To_OUTPUT_From_Table → Run**.

> Αν τα macros είναι blocked: κάνε Unblock στο .docm ή βάλε τον φάκελο σε Trusted Locations.

---

## English (EN)

### Quick Setup (Word 2016 Windows)
1. Open `examples/00_Controller_example.docx` and **Save As → .docm**.
2. **Alt+F11 → File → Import File…** and import `src/ControllerModule_TimeOutput_GR_ANSI.bas`.
3. Put your templates next to the `.docm` and name them `TEMPLATE_*.docx`.
4. Ensure templates contain:
   - `{{OraEnarxis}}` for start time
   - `{{OraPeratosis}}` for end time
5. Run: **Developer → Macros → Generate_Reports_To_OUTPUT_From_Table → Run**.

---

## License
MIT (see `LICENSE`).
