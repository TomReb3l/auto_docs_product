# auto_doc_product

**GR:** Παραγωγή τελικών εγγράφων Word (DOCX) από templates με placeholders, με ένα κλικ μέσω VBA macro.  
**EN:** One-click generation of final Word documents (DOCX) from placeholder templates using a VBA macro.

---

## Ελληνικά (GR)

### Τι κάνει
- Διαβάζει τιμές από ένα **Controller Word (.docm)** που περιέχει έναν πίνακα **ΠΕΔΙΟ / ΤΙΜΗ**
- Παίρνει templates που ταιριάζουν στο μοτίβο **`TEMPLATE_*.docx`**
- Δημιουργεί φάκελο εξαγωγής με όνομα **ίσο με την τιμή του `CaseID`** (π.χ. `3008-12-150\`)
- Παράγει **ξεχωριστά τελικά έγγραφα** στον φάκελο `CaseID` χωρίς να πειράζει τα templates
- Αντικαθιστά placeholders τύπου `{{Key}}` με τις αντίστοιχες τιμές
- Τα εξαγόμενα αρχεία παίρνουν όνομα **όπως το template**, αλλά:
  - **χωρίς** το πρόθεμα `TEMPLATE_`
  - **χωρίς** το `CaseID_` μπροστά

### Timestamp (με ή χωρίς)
Υπάρχουν δύο τρόποι εκτέλεσης:
- **Με timestamp** (ώρα έναρξης/περάτωσης): συμπληρώνει αυτόματα `{{OraEnarxis}}` και `{{OraPeratosis}}`
- **Χωρίς timestamp**: αφήνει τα `{{OraEnarxis}}` και `{{OraPeratosis}}` **κενά**

> Και στις δύο περιπτώσεις, ο φάκελος εξαγωγής είναι πάντα ο φάκελος με όνομα `CaseID`, και τα αρχεία εξόδου έχουν όνομα “σαν το template” χωρίς `TEMPLATE_`.

### Πεδίο χρήσης
Το project **δεν είναι αποκλειστικά για χρήση από αστυνομικές/δημόσιες αρχές**.  
Είναι ένα γενικό εργαλείο “template → output” για Word, που μπορεί να χρησιμοποιηθεί με τον ίδιο τρόπο για **οποιοδήποτε έγγραφο** (π.χ. αναφορές, πρακτικά, αιτήσεις, βεβαιώσεις, εταιρικά έντυπα, checklists), αρκεί να υπάρχουν placeholders τύπου `{{...}}` μέσα στα templates.

### Σειρά εκθέσεων (πολύ σημαντικό για τις ώρες)
Για να υπολογίζονται σωστά οι ώρες έναρξης/περάτωσης, τα templates πρέπει να έχουν **σαφή σειρά**.

Προτείνεται να βάζετε **αρίθμηση στην αρχή του ονόματος αρχείου**, π.χ.:
- `TEMPLATE_1. ... .docx`
- `TEMPLATE_2. ... .docx`
- `TEMPLATE_10. ... .docx`

✅ Το macro ταξινομεί τα templates **αριθμητικά** (με βάση τον **πρώτο αριθμό** στο όνομα αρχείου), ώστε η σειρά να είναι πάντα προβλέψιμη και οι υπολογισμοί ώρας να βγαίνουν σωστά.

### Γρήγορο Setup (Word 2016 Windows)
1. Ανοίξτε το `examples/00_Controller_example.docx` και κάντε **Save As → .docm**.
2. Πατήστε **Alt+F11 → File → Import File…** και κάντε import **το `.bas` που θέλετε**:
   - **Με timestamp:** `src/ControllerModule_CaseFolder_StripTemplate_Sorted_Time_GR_ANSI.bas`
   - **Χωρίς timestamp:** `src/ControllerModule_CaseFolder_StripTemplate_Sorted_NoTime_GR_ANSI.bas`
   Μετά κάντε αποθήκευση **Ctrl+S**.
3. Βάλτε τα templates στον ίδιο φάκελο με το `.docm` και ονομάστε τα `TEMPLATE_*.docx`.
4. Αν θέλετε ώρες, βάλτε στα templates:
   - `{{OraEnarxis}}` στην αρχή (ώρα έναρξης)
   - `{{OraPeratosis}}` στο τέλος (ώρα περάτωσης)

### Εκτέλεση (Run)
Μπορείτε να τρέξετε το macro είτε από το tab **Developer → Macros**, είτε πιο γρήγορα:
- Πατήστε **Alt + F8**
- Επιλέξτε το αντίστοιχο macro:
  - **Με timestamp:** `Generate_Reports_To_CaseIDFolder_From_Table`
  - **Χωρίς timestamp:** `Generate_Reports_To_CaseIDFolder_NoTime_From_Table`
- Πατήστε **Run / Εκτέλεση**

> Αν τα macros είναι blocked: κάντε Unblock στο `.docm` ή προσθέστε τον φάκελο σε **Trusted Locations**.

### Σημείωση ασφάλειας
Το αποθετήριο περιέχει **απλό VBA κώδικα** και παραδείγματα templates.  
Δεν κάνει λήψεις/εκτέλεση εξωτερικού κώδικα, δεν κάνει δικτυακές κλήσεις και δεν επιχειρεί να “παραμείνει” στο σύστημα.  
Ο κώδικας είναι **πλήρως αναγνώσιμος/ελέγξιμος**. Όπως με κάθε macro-enabled λύση, παρακαλούμε να το τρέχετε μόνο σε αξιόπιστο περιβάλλον και να κάνετε έλεγχο με τα εργαλεία ασφαλείας που χρησιμοποιείτε.

---

## English (EN)

### What it does
**auto_doc_product** is a lightweight Word automation workflow that generates finalized documents from reusable templates — and it can be used for **any document type**, not tied to a specific domain or organization.

- You maintain one **Controller** document (a `.docm`) containing a simple **KEY / VALUE** table.
- You create one or more **template `.docx` files** containing placeholders like `{{Key}}`.
- The tool creates an output folder named **exactly as `CaseID`** (e.g., `3008-12-150\`).
- With a single macro run, it:
  - Reads values from the Controller table
  - Copies each selected template into the **CaseID folder** (templates stay unchanged)
  - Replaces all placeholders across the document (including headers/footers and text boxes when present)
- Output filenames are based on the template name, but:
  - the `TEMPLATE_` prefix is removed
  - no `CaseID_` prefix is added

### Timestamp (with or without)
Two execution modes are supported:
- **With timestamp:** automatically fills `{{OraEnarxis}}` and `{{OraPeratosis}}`
- **Without timestamp:** leaves `{{OraEnarxis}}` and `{{OraPeratosis}}` blank

### Document order (important for time calculations)
To calculate start/end times correctly, templates should have a **clear and predictable order**.

We recommend adding a **leading number** to each template filename, for example:
- `TEMPLATE_1. ... .docx`
- `TEMPLATE_2. ... .docx`
- `TEMPLATE_10. ... .docx`

✅ The macro sorts templates **numerically** (based on the **first number** in the filename) to ensure a consistent, predictable sequence and correct time progression.

### Quick Setup (Word 2016 Windows)
1. Open `examples/00_Controller_example.docx` and **Save As → .docm**.
2. Press **Alt+F11 → File → Import File…** and import the `.bas` you want:
   - **With timestamp:** `src/ControllerModule_CaseFolder_StripTemplate_Sorted_Time_GR_ANSI.bas`
   - **Without timestamp:** `src/ControllerModule_CaseFolder_StripTemplate_Sorted_NoTime_GR_ANSI.bas`
   Then save with **Ctrl+S**.
3. Put your templates next to the `.docm` and name them `TEMPLATE_*.docx`.
4. If you want timestamps, ensure templates contain:
   - `{{OraEnarxis}}` for start time
   - `{{OraPeratosis}}` for end time

### Run
You can run the macro from **Developer → Macros**, or quickly:
- Press **Alt + F8**
- Select the macro you want:
  - **With timestamp:** `Generate_Reports_To_CaseIDFolder_From_Table`
  - **Without timestamp:** `Generate_Reports_To_CaseIDFolder_NoTime_From_Table`
- Click **Run**

### Scope of use
This project is **not exclusively intended for law enforcement / public authorities**.  
It is a general-purpose “template → output” Word automation tool and can be used the same way to generate **any kind of document** (e.g., reports, minutes, requests, certificates, corporate forms, checklists), as long as the templates contain `{{...}}` placeholders.

### Security note
This repository contains **plain VBA source code** and example templates.  
It does **not** download or execute external code, does **not** make network requests, and does not try to **persist** on the system.  
The code is **fully readable/auditable** in this repo. As with any macro-enabled workflow, please run it only in trusted environments and feel free to scan the files with your preferred security tools.

---

## License
MIT (see `LICENSE`).
