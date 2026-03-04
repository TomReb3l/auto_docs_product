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

### Πεδίο χρήσης
Το project **δεν είναι αποκλειστικά για χρήση από αστυνομικές/δημόσιες αρχές**.  
Είναι ένα γενικό εργαλείο “template → output” για Word, που μπορεί να χρησιμοποιηθεί με τον ίδιο τρόπο για **οποιοδήποτε έγγραφο** (π.χ. αναφορές, πρακτικά, αιτήσεις, βεβαιώσεις, εταιρικά έντυπα, checklists), αρκεί να υπάρχουν placeholders τύπου `{{...}}` μέσα στα templates.

### Γρήγορο Setup (Word 2016 Windows)
1. Άνοιξε το `examples/00_Controller_example.docx` και κάνε **Save As → .docm**.
2. **Alt+F11 → File → Import File…** και κάνε import το `src/ControllerModule_TimeOutput_GR_ANSI.bas`.
3. Βάλε τα δικά σου templates στον ίδιο φάκελο με το `.docm` και ονόμασέ τα `TEMPLATE_*.docx`.
4. Βεβαιώσου ότι στα templates έχεις:
   - `{{OraEnarxis}}` στην αρχή (ώρα έναρξης)
   - `{{OraPeratosis}}` στο τέλος (ώρα περάτωσης)
5. **Developer → Macros → Generate_Reports_To_OUTPUT_From_Table → Run**.

> Αν τα macros είναι blocked: κάνε Unblock στο .docm ή βάλε τον φάκελο σε Trusted Locations.

## Σημείωση ασφάλειας

Το αποθετήριο περιέχει **απλό VBA κώδικα** και παραδείγματα templates.  
Δεν κάνει λήψεις ή εκτέλεση εξωτερικού κώδικα, δεν κάνει δικτυακές κλήσεις και δεν επιχειρεί “persistence”.  
Ο κώδικας είναι **πλήρως αναγνώσιμος/ελέγξιμος**. Όπως με κάθε macro-enabled λύση, να το τρέχετε μόνο σε αξιόπιστο περιβάλλον και μπορείτε να κάνετε έλεγχο με τα εργαλεία ασφαλείας που χρησιμοποιείτε.

---

## English (EN)

### What it does

**autodocproduct** is a lightweight Word automation workflow that generates finalized documents from reusable templates — and it can be used for **any document type**, not tied to a specific domain or organization.

- You maintain one **Controller** document (a `.docm`) containing a simple **KEY / VALUE** table.
- You create one or more **template `.docx` files** containing placeholders like `{{Key}}`.
- With a single macro run, the tool:
  - Reads the values from the Controller table
  - Copies each template into an `OUTPUT` folder (templates stay unchanged)
  - Replaces all placeholders with the corresponding values across the document (including headers/footers and text boxes when present)
  - Optionally applies time-based rules (e.g., calculated start/end timestamps per generated document)

Each run produces a clean set of output documents ready to share, print, or archive.
### Quick Setup (Word 2016 Windows)
1. Open `examples/00_Controller_example.docx` and **Save As → .docm**.
2. **Alt+F11 → File → Import File…** and import `src/ControllerModule_TimeOutput_GR_ANSI.bas`.
3. Put your templates next to the `.docm` and name them `TEMPLATE_*.docx`.
4. Ensure templates contain:
   - `{{OraEnarxis}}` for start time
   - `{{OraPeratosis}}` for end time
5. Run: **Developer → Macros → Generate_Reports_To_OUTPUT_From_Table → Run**.

### Scope of use
This project is **not exclusively intended for law enforcement / public authorities**.  
It is a general-purpose “template → output” Word automation tool and can be used the same way to generate **any kind of document** (e.g., reports, minutes, requests, certificates, corporate forms, checklists), as long as the templates contain `{{...}}` placeholders.

## Security note

This repository contains **plain VBA source code** and example templates.  
It does **not** download or execute external code, does **not** make network requests, and does **not** attempt persistence.  
The code is **fully readable/auditable** in this repo. As with any macro-enabled workflow, you should run it only in trusted environments and you may scan the files with your preferred security tools.

---

## License
MIT (see `LICENSE`).
