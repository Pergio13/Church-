function sendChurchReport() {

  // Παίρνουμε το ενεργό Google Spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Το email που θα λαμβάνει την αναφορά
  const recipient = "pergiorgos13@gmail.com";

  // Πόσες μέρες πριν από λήξη φαρμάκου να ειδοποιεί
  const medicineWarningDays = 30;

  // Πόσες μέρες πριν από συντήρηση να ειδοποιεί
  const maintenanceWarningDays = 30;



  // =====================================================
  // ΒΑΣΙΚΗ ΔΟΜΗ ΔΕΔΟΜΕΝΩΝ ΤΟΥ REPORT
  // =====================================================
  const report = {
    "Καθαριότητα": {
      buy: [],
      low: []
    },
    "Τράπεζα Αγάπης": {
      buy: [],
      low: []
    },
    "Φαρμακείο": {
      expiring: []
    },
    "Συντήρηση": {
      upcoming: []
    }
  };



  // =====================================================
  // ΣΥΝΑΡΤΗΣΗ: ΜΕΤΑΤΡΕΠΕΙ ΜΙΑ ΤΙΜΗ ΣΕ ΑΡΙΘΜΟ ΜΕ ΑΣΦΑΛΕΙΑ
  // =====================================================
  function toNumber(value) {

    // Αν το κελί είναι κενό επιστρέφουμε null
    if (value === "" || value === null || value === undefined) {
      return null;
    }

    // Προσπαθούμε να μετατρέψουμε την τιμή σε αριθμό
    const num = Number(value);

    // Αν δεν είναι αριθμός επιστρέφουμε null
    if (isNaN(num)) {
      return null;
    }

    // Αλλιώς επιστρέφουμε τον αριθμό
    return num;
  }



  // =====================================================
  // ΕΛΕΓΧΟΣ ΑΠΟΘΗΚΗΣ
  // =====================================================
  const inventorySheets = [
    { name: "Καθαριότητα ", label: "Καθαριότητα" },
    { name: "Τράπεζα Αγάπης ", label: "Τράπεζα Αγάπης" }
  ];

  // Για κάθε φύλλο αποθήκης
  inventorySheets.forEach(config => {

    // Παίρνουμε το φύλλο από το spreadsheet
    const sheet = ss.getSheetByName(config.name);

    // Αν δεν υπάρχει, το προσπερνάμε
    if (!sheet) return;

    // Παίρνουμε όλα τα δεδομένα του φύλλου
    const data = sheet.getDataRange().getValues();

    // Ξεκινάμε από τη 2η γραμμή γιατί η 1η έχει επικεφαλίδες
    for (let i = 1; i < data.length; i++) {

      // Στήλη A = Κατηγορία
      const category = data[i][0];

      // Στήλη B = Είδος
      const item = data[i][1];

      // Στήλη C = Ελάχιστη ποσότητα
      const minQty = toNumber(data[i][2]);

      // Στήλη D = Ποσότητα
      const qty = toNumber(data[i][3]);

      // Αν δεν υπάρχει είδος, προσπερνάμε
      if (!item) continue;

      // Αν λείπουν αριθμοί, προσπερνάμε
      if (minQty === null || qty === null) continue;

      // Αν η ποσότητα είναι μικρότερη από την ελάχιστη → προς αγορά
      if (qty < minQty) {
        report[config.label].buy.push({
          item: item,
          category: category || "-",
          qty: qty,
          min: minQty,
          missing: minQty - qty
        });
      }

      // Αν η ποσότητα είναι ίση με την ελάχιστη → χαμηλό απόθεμα
      else if (qty === minQty) {
        report[config.label].low.push({
          item: item,
          category: category || "-",
          qty: qty,
          min: minQty
        });
      }
    }
  });



  // =====================================================
  // ΕΛΕΓΧΟΣ ΦΑΡΜΑΚΕΙΟΥ
  // =====================================================
  const pharmacySheet = ss.getSheetByName("Φαρμακείο ");

  if (pharmacySheet) {

    // Παίρνουμε όλα τα δεδομένα του φύλλου
    const data = pharmacySheet.getDataRange().getValues();

    // Σημερινή ημερομηνία
    const today = new Date();

    // Δημιουργούμε ημερομηνία ορίου
    const limit = new Date();

    // Προσθέτουμε τις μέρες ειδοποίησης
    limit.setDate(today.getDate() + medicineWarningDays);

    // Ξεκινάμε από τη 2η γραμμή
    for (let i = 1; i < data.length; i++) {

      // Στήλη A = Όνομα φαρμάκου
      const name = data[i][0];

      // Στήλη B = Ημερομηνία λήξης
      const expiry = data[i][1];

      // Αν λείπει όνομα ή ημερομηνία, προσπερνάμε
      if (!name || !expiry) continue;

      // Αν είναι κανονική ημερομηνία
      if (expiry instanceof Date) {

        // Αν λήγει μέσα στο όριο ειδοποίησης
        if (expiry <= limit) {
          report["Φαρμακείο"].expiring.push({
            item: name,
            date: Utilities.formatDate(
              expiry,
              Session.getScriptTimeZone(),
              "dd/MM/yyyy"
            )
          });
        }
      }
    }
  }



  // =====================================================
  // ΕΛΕΓΧΟΣ ΣΥΝΤΗΡΗΣΗΣ
  // =====================================================
  const maintenanceSheet = ss.getSheetByName("Συντήρηση ");

  if (maintenanceSheet) {

    // Παίρνουμε τα δεδομένα του φύλλου
    const data = maintenanceSheet.getDataRange().getValues();

    // Σημερινή ημερομηνία
    const today = new Date();

    // Ημερομηνία ορίου
    const limit = new Date();

    // Προσθέτουμε τις μέρες ειδοποίησης
    limit.setDate(today.getDate() + maintenanceWarningDays);

    // Ξεκινάμε από τη 2η γραμμή
    for (let i = 1; i < data.length; i++) {

      // Στήλη A = Είδος
      const item = data[i][0];

      // Στήλη C = Επόμενη ενέργεια
      const next = data[i][2];

      // Αν λείπει είδος ή ημερομηνία, προσπερνάμε
      if (!item || !next) continue;

      // Αν είναι ημερομηνία
      if (next instanceof Date) {

        // Αν είναι μέσα στο όριο ειδοποίησης
        if (next <= limit) {
          report["Συντήρηση"].upcoming.push({
            item: item,
            date: Utilities.formatDate(
              next,
              Session.getScriptTimeZone(),
              "dd/MM/yyyy"
            )
          });
        }
      }
    }
  }



  // =====================================================
  // ΥΠΟΛΟΓΙΣΜΟΣ ΣΥΝΟΛΩΝ ΓΙΑ ΤΟ HEADER
  // =====================================================
  const totalBuy =
    report["Καθαριότητα"].buy.length +
    report["Τράπεζα Αγάπης"].buy.length;

  const totalLow =
    report["Καθαριότητα"].low.length +
    report["Τράπεζα Αγάπης"].low.length;

  const totalExpiring = report["Φαρμακείο"].expiring.length;

  const totalUpcoming = report["Συντήρηση"].upcoming.length;

  const reportDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy"
  );



  // =====================================================
  // ΣΥΝΑΡΤΗΣΗ ΠΟΥ ΦΤΙΑΧΝΕΙ ΚΑΡΤΕΣ ΓΙΑ ΑΠΟΘΗΚΗ
  // =====================================================
  function createInventoryTable(items, type) {

    // Αν δεν υπάρχουν στοιχεία επιστρέφουμε ήρεμο μήνυμα
    if (items.length === 0) {
      return `
        <div style="padding:14px 16px; background:#f8fafc; border:1px dashed #d6dde8; border-radius:10px; color:#6b7280; font-size:14px;">
          Δεν υπάρχουν καταχωρήσεις.
        </div>
      `;
    }

    // Ξεκινάμε HTML table
    let html = `
      <table style="width:100%; border-collapse:collapse;">
    `;

    // Για κάθε γραμμή
    items.forEach(row => {

      // Αν είναι "προς αγορά" δείχνουμε και πόσα λείπουν
      const extraInfo = type === "buy"
        ? `<div style="margin-top:4px; font-size:12px; color:#b91c1c;">Λείπουν ακόμη: ${row.missing}</div>`
        : ``;

      html += `
        <tr>
          <td style="padding:14px 12px; border-bottom:1px solid #edf1f5;">
            <div style="font-size:15px; font-weight:700; color:#111827;">${row.item}</div>
            <div style="margin-top:4px; font-size:12px; color:#6b7280;">Κατηγορία: ${row.category}</div>
            <div style="margin-top:4px; font-size:12px; color:#6b7280;">Ποσότητα: ${row.qty} | Ελάχιστη: ${row.min}</div>
            ${extraInfo}
          </td>
        </tr>
      `;
    });

    // Κλείνουμε table
    html += `</table>`;

    return html;
  }



  // =====================================================
  // ΣΥΝΑΡΤΗΣΗ ΠΟΥ ΦΤΙΑΧΝΕΙ ΚΑΡΤΕΣ ΜΕ ΗΜΕΡΟΜΗΝΙΕΣ
  // =====================================================
  function createDateTable(items, label) {

    // Αν δεν υπάρχουν στοιχεία
    if (items.length === 0) {
      return `
        <div style="padding:14px 16px; background:#f8fafc; border:1px dashed #d6dde8; border-radius:10px; color:#6b7280; font-size:14px;">
          Δεν υπάρχουν καταχωρήσεις.
        </div>
      `;
    }

    let html = `
      <table style="width:100%; border-collapse:collapse;">
    `;

    items.forEach(row => {
      html += `
        <tr>
          <td style="padding:14px 12px; border-bottom:1px solid #edf1f5;">
            <div style="font-size:15px; font-weight:700; color:#111827;">${row.item}</div>
            <div style="margin-top:4px; font-size:12px; color:#6b7280;">${label}: ${row.date}</div>
          </td>
        </tr>
      `;
    });

    html += `</table>`;

    return html;
  }



  // =====================================================
  // ΣΥΝΑΡΤΗΣΗ ΠΟΥ ΦΤΙΑΧΝΕΙ ΤΟ ΠΛΑΙΣΙΟ ΚΑΘΕ ΚΥΡΙΑΣ ΕΝΟΤΗΤΑΣ
  // =====================================================
  function createMainSection(title, emoji, content) {
    return `
      <div style="margin-bottom:24px; background:#ffffff; border:1px solid #e5eaf0; border-radius:16px; overflow:hidden;">
        <div style="padding:16px 20px; background:#f8fafc; border-bottom:1px solid #e5eaf0;">
          <div style="font-size:18px; font-weight:800; color:#0f172a;">${emoji} ${title}</div>
        </div>
        <div style="padding:20px;">
          ${content}
        </div>
      </div>
    `;
  }



  // =====================================================
  // ΣΥΝΑΡΤΗΣΗ ΠΟΥ ΦΤΙΑΧΝΕΙ ΥΠΟΕΝΟΤΗΤΕΣ
  // =====================================================
  function createSubSection(title, count, color, content) {
    return `
      <div style="margin-bottom:18px;">
        <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:10px;">
          <div style="font-size:15px; font-weight:700; color:${color};">${title}</div>
          <div style="background:${color}; color:#ffffff; font-size:12px; font-weight:700; padding:4px 10px; border-radius:999px;">
            ${count}
          </div>
        </div>
        ${content}
      </div>
    `;
  }



  // =====================================================
  // ΔΗΜΙΟΥΡΓΙΑ HTML EMAIL
  // =====================================================
  const htmlBody = `
    <div style="margin:0; padding:0; background:#eef2f7; font-family:Arial, Helvetica, sans-serif;">
      <div style="max-width:860px; margin:0 auto; padding:30px 16px;">

        <div style="background:#ffffff; border-radius:22px; overflow:hidden; box-shadow:0 8px 30px rgba(15, 23, 42, 0.08);">

          <div style="background:linear-gradient(135deg, #0d47a1, #1565c0); padding:28px 24px; color:#ffffff;">
            <div style="font-size:26px; font-weight:800; line-height:1.2;">Αναφορά Υλικών Εκκλησίας</div>
            <div style="margin-top:8px; font-size:14px; opacity:0.95;">Ημερομηνία αναφοράς: ${reportDate}</div>

            <div style="margin-top:20px; display:flex; flex-wrap:wrap; gap:10px;">
              <div style="background:rgba(255,255,255,0.16); padding:10px 14px; border-radius:12px; font-size:13px; font-weight:700;">Προς αγορά: ${totalBuy}</div>
              <div style="background:rgba(255,255,255,0.16); padding:10px 14px; border-radius:12px; font-size:13px; font-weight:700;">Χαμηλό απόθεμα: ${totalLow}</div>
              <div style="background:rgba(255,255,255,0.16); padding:10px 14px; border-radius:12px; font-size:13px; font-weight:700;">Λήξεις φαρμάκων: ${totalExpiring}</div>
              <div style="background:rgba(255,255,255,0.16); padding:10px 14px; border-radius:12px; font-size:13px; font-weight:700;">Συντηρήσεις: ${totalUpcoming}</div>
            </div>
          </div>

          <div style="padding:24px; background:#f6f8fb;">

            ${createMainSection(
              "Καθαριότητα",
              "🧼",
              createSubSection(
                "Προς αγορά",
                report["Καθαριότητα"].buy.length,
                "#b91c1c",
                createInventoryTable(report["Καθαριότητα"].buy, "buy")
              ) +
              createSubSection(
                "Χαμηλό απόθεμα",
                report["Καθαριότητα"].low.length,
                "#d97706",
                createInventoryTable(report["Καθαριότητα"].low, "low")
              )
            )}

            ${createMainSection(
              "Τράπεζα Αγάπης",
              "🛍️",
              createSubSection(
                "Προς αγορά",
                report["Τράπεζα Αγάπης"].buy.length,
                "#b91c1c",
                createInventoryTable(report["Τράπεζα Αγάπης"].buy, "buy")
              ) +
              createSubSection(
                "Χαμηλό απόθεμα",
                report["Τράπεζα Αγάπης"].low.length,
                "#d97706",
                createInventoryTable(report["Τράπεζα Αγάπης"].low, "low")
              )
            )}

            ${createMainSection(
              "Φαρμακείο",
              "💊",
              createSubSection(
                "Φάρμακα που λήγουν σύντομα",
                report["Φαρμακείο"].expiring.length,
                "#7c3aed",
                createDateTable(report["Φαρμακείο"].expiring, "Ημερομηνία λήξης")
              )
            )}

            ${createMainSection(
              "Συντήρηση",
              "🛠️",
              createSubSection(
                "Εργασίες που πλησιάζουν",
                report["Συντήρηση"].upcoming.length,
                "#2563eb",
                createDateTable(report["Συντήρηση"].upcoming, "Ημερομηνία")
              )
            )}

          </div>

          <div style="padding:18px 24px; background:#ffffff; border-top:1px solid #e5eaf0;">
            <div style="font-size:12px; color:#6b7280; line-height:1.6;">
              Το παρόν email δημιουργήθηκε αυτόματα από το αρχείο διαχείρισης υλικών της εκκλησίας.
            </div>
          </div>

        </div>
      </div>
    </div>
  `;



  // =====================================================
  // ΔΗΜΙΟΥΡΓΙΑ ΑΠΛΟΥ TEXT EMAIL ΩΣ ΕΝΑΛΛΑΚΤΙΚΟ
  // =====================================================
  let plainBody = `ΑΝΑΦΟΡΑ ΥΛΙΚΩΝ ΕΚΚΛΗΣΙΑΣ - ${reportDate}\n\n`;

  plainBody += `ΚΑΘΑΡΙΟΤΗΤΑ\n`;
  plainBody += `Προς αγορά: ${report["Καθαριότητα"].buy.length}\n`;
  report["Καθαριότητα"].buy.forEach(row => {
    plainBody += `- ${row.item} | Ποσότητα: ${row.qty} | Ελάχιστη: ${row.min} | Λείπουν: ${row.missing}\n`;
  });
  plainBody += `Χαμηλό απόθεμα: ${report["Καθαριότητα"].low.length}\n`;
  report["Καθαριότητα"].low.forEach(row => {
    plainBody += `- ${row.item} | Ποσότητα: ${row.qty} | Ελάχιστη: ${row.min}\n`;
  });

  plainBody += `\nΤΡΑΠΕΖΑ ΑΓΑΠΗΣ\n`;
  plainBody += `Προς αγορά: ${report["Τράπεζα Αγάπης"].buy.length}\n`;
  report["Τράπεζα Αγάπης"].buy.forEach(row => {
    plainBody += `- ${row.item} | Ποσότητα: ${row.qty} | Ελάχιστη: ${row.min} | Λείπουν: ${row.missing}\n`;
  });
  plainBody += `Χαμηλό απόθεμα: ${report["Τράπεζα Αγάπης"].low.length}\n`;
  report["Τράπεζα Αγάπης"].low.forEach(row => {
    plainBody += `- ${row.item} | Ποσότητα: ${row.qty} | Ελάχιστη: ${row.min}\n`;
  });

  plainBody += `\nΦΑΡΜΑΚΕΙΟ\n`;
  plainBody += `Φάρμακα που λήγουν σύντομα: ${report["Φαρμακείο"].expiring.length}\n`;
  report["Φαρμακείο"].expiring.forEach(row => {
    plainBody += `- ${row.item} | Λήξη: ${row.date}\n`;
  });

  plainBody += `\nΣΥΝΤΗΡΗΣΗ\n`;
  plainBody += `Εργασίες που πλησιάζουν: ${report["Συντήρηση"].upcoming.length}\n`;
  report["Συντήρηση"].upcoming.forEach(row => {
    plainBody += `- ${row.item} | Ημερομηνία: ${row.date}\n`;
  });



  // =====================================================
  // ΑΠΟΣΤΟΛΗ EMAIL
  // =====================================================
  MailApp.sendEmail({
    to: recipient,
    subject: "Αναφορά Υλικών Εκκλησίας",
    body: plainBody,
    htmlBody: htmlBody
  });

}
