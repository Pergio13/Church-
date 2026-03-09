function sendChurchReport() {

  // Παίρνουμε το ενεργό Google Spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Το email του παραλήπτη της αναφοράς
  const recipient = "pergiorgos13@gmail.com";

  // Πόσες μέρες πριν από λήξη φαρμάκου να θεωρείται "λήγει σύντομα"
  const medicineWarningDays = 30;

  // Πόσες μέρες πριν από συντήρηση να εμφανίζεται στην αναφορά
  const maintenanceWarningDays = 45;

  // Παίρνουμε τη ζώνη ώρας του script για σωστή μορφοποίηση ημερομηνιών
  const timeZone = Session.getScriptTimeZone();

  // Παίρνουμε τη σημερινή ημερομηνία
  const today = new Date();

  // Μηδενίζουμε ώρα/λεπτά/δευτερόλεπτα για πιο σωστές συγκρίσεις ημερομηνιών
  today.setHours(0, 0, 0, 0);

  // Δημιουργούμε την ημερομηνία αναφοράς σε μορφή dd/MM/yyyy
  const reportDate = Utilities.formatDate(today, timeZone, "dd/MM/yyyy");

  // Δημιουργούμε όριο για "λήγει σύντομα" στα φάρμακα
  const medicineLimit = new Date(today);
  medicineLimit.setDate(medicineLimit.getDate() + medicineWarningDays);

  // Δημιουργούμε όριο για "προσεχής συντήρηση"
  const maintenanceLimit = new Date(today);
  maintenanceLimit.setDate(maintenanceLimit.getDate() + maintenanceWarningDays);

  // Κεντρική δομή δεδομένων της αναφοράς
  const report = {
    cleanliness: {
      title: "Καθαριότητα",
      buy: [],
      low: []
    },
    loveBank: {
      title: "Τράπεζα Αγάπης",
      buy: [],
      low: []
    },
    pharmacy: {
      title: "Φαρμακείο",
      expiringSoon: [],
      expired: []
    },
    maintenance: {
      title: "Συντήρηση",
      upcoming: []
    }
  };

  // Μετατρέπει με ασφάλεια μια τιμή σε αριθμό
  function toNumber(value) {
    if (value === "" || value === null || value === undefined) return null;
    const num = Number(value);
    return isNaN(num) ? null : num;
  }

  // Κάνει escape ειδικούς χαρακτήρες για να μην σπάσει το HTML
  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  // Επιστρέφει σωστό ελληνικό όρο για ενικό/πληθυντικό
  function countLabel(count, singular, plural) {
    return count === 1 ? singular : plural;
  }

  // Υπολογίζει πόσες μέρες απομένουν μέχρι μια ημερομηνία
  function daysUntil(dateObj) {
    const target = new Date(dateObj);
    target.setHours(0, 0, 0, 0);
    const diffMs = target.getTime() - today.getTime();
    return Math.round(diffMs / (1000 * 60 * 60 * 24));
  }

  // Δημιουργεί badge / pill
  function createPill(text, bgColor, textColor) {
    return `
      <span style="display:inline-block; padding:4px 10px; font-size:11px; font-weight:700; border-radius:999px; background-color:${bgColor}; color:${textColor};">
        ${escapeHtml(text)}
      </span>
    `;
  }

  // Δημιουργεί το κενό μήνυμα όταν δεν υπάρχουν καταχωρήσεις
  function createEmptyBox(text) {
    return `
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
        <tr>
          <td style="padding:14px 16px; border:1px dashed #d6dde8; border-radius:12px; background-color:#f8fafc; font-size:14px; color:#6b7280;">
            ${escapeHtml(text)}
          </td>
        </tr>
      </table>
    `;
  }

  // 1. Έλεγχος αποθήκης
  const inventorySheets = [
    { sheetName: "Καθαριότητα ", targetKey: "cleanliness" },
    { sheetName: "Τράπεζα Αγάπης ", targetKey: "loveBank" }
  ];

  inventorySheets.forEach(config => {
    const sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const item = data[i][1];
      const minQty = toNumber(data[i][2]);
      const qty = toNumber(data[i][3]);

      if (!item) continue;
      if (minQty === null || qty === null) continue;

      if (qty < minQty) {
        report[config.targetKey].buy.push({
          item: item,
          qty: qty,
          minQty: minQty,
          missing: minQty - qty
        });
      } else if (qty === minQty) {
        report[config.targetKey].low.push({
          item: item,
          qty: qty,
          minQty: minQty
        });
      }
    }
  });

  // 2. Έλεγχος φαρμακείου
  const pharmacySheet = ss.getSheetByName("Φαρμακείο ");

  if (pharmacySheet) {
    const data = pharmacySheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const medicineName = data[i][0];
      const expiryDate = data[i][1];

      if (!medicineName || !expiryDate) continue;
      if (!(expiryDate instanceof Date)) continue;

      const cleanExpiry = new Date(expiryDate);
      cleanExpiry.setHours(0, 0, 0, 0);

      const formattedDate = Utilities.formatDate(cleanExpiry, timeZone, "dd/MM/yyyy");

      if (cleanExpiry < today) {
        report.pharmacy.expired.push({
          item: medicineName,
          date: formattedDate
        });
      } else if (cleanExpiry <= medicineLimit) {
        report.pharmacy.expiringSoon.push({
          item: medicineName,
          date: formattedDate
        });
      }
    }
  }

  // 3. Έλεγχος συντήρησης
  const maintenanceSheet = ss.getSheetByName("Συντήρηση ");

  if (maintenanceSheet) {
    const data = maintenanceSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const itemName = data[i][0];
      const nextActionDate = data[i][2];

      if (!itemName || !nextActionDate) continue;
      if (!(nextActionDate instanceof Date)) continue;

      const cleanNextAction = new Date(nextActionDate);
      cleanNextAction.setHours(0, 0, 0, 0);

      if (cleanNextAction <= maintenanceLimit) {
        report.maintenance.upcoming.push({
          item: itemName,
          date: Utilities.formatDate(cleanNextAction, timeZone, "dd/MM/yyyy"),
          daysLeft: daysUntil(cleanNextAction)
        });
      }
    }
  }

  // 4. Ταξινόμηση
  report.cleanliness.buy.sort((a, b) => b.missing - a.missing);
  report.loveBank.buy.sort((a, b) => b.missing - a.missing);

  report.cleanliness.low.sort((a, b) => String(a.item).localeCompare(String(b.item), "el"));
  report.loveBank.low.sort((a, b) => String(a.item).localeCompare(String(b.item), "el"));

  report.pharmacy.expired.sort((a, b) => a.date.localeCompare(b.date, "el"));
  report.pharmacy.expiringSoon.sort((a, b) => a.date.localeCompare(b.date, "el"));

  report.maintenance.upcoming.sort((a, b) => a.daysLeft - b.daysLeft);

  // 5. Υπολογισμός συνόλων
  const totalBuy = report.cleanliness.buy.length + report.loveBank.buy.length;
  const totalLow = report.cleanliness.low.length + report.loveBank.low.length;
  const totalExpired = report.pharmacy.expired.length;
  const totalUpcomingMaintenance = report.maintenance.upcoming.length;

  const urgentItems = [];

  report.pharmacy.expired.forEach(item => {
    urgentItems.push({
      type: "expiredMedicine",
      item: item.item,
      date: item.date
    });
  });

  report.cleanliness.buy.forEach(item => {
    if (item.missing >= 5) {
      urgentItems.push({
        type: "buy",
        item: item.item,
        extra: `Χρειάζονται τουλάχιστον ακόμη: ${item.missing}`
      });
    }
  });

  report.loveBank.buy.forEach(item => {
    if (item.missing >= 5) {
      urgentItems.push({
        type: "buy",
        item: item.item,
        extra: `Χρειάζονται τουλάχιστον ακόμη: ${item.missing}`
      });
    }
  });

  const totalUrgent = urgentItems.length;

  // 6. Συναρτήσεις HTML

  // Κάρτα προϊόντος προς αγορά χωρίς "Κατηγορία" και χωρίς badge
  function createBuyCard(row) {
    return `
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin-bottom:12px;">
        <tr>
          <td style="padding:16px 14px; border:1px solid #edf1f5; border-radius:12px; background-color:#fffafa;">
            <div style="font-size:15px; font-weight:700; color:#111827;">
              ${escapeHtml(row.item)}
            </div>
            <div style="margin-top:5px; font-size:13px; line-height:1.5; color:#b91c1c;">
              Πρέπει να αγοραστούν τουλάχιστον: ${escapeHtml(row.missing)}
            </div>
          </td>
        </tr>
      </table>
    `;
  }

  // Κάρτα χαμηλού αποθέματος χωρίς "Κατηγορία" και χωρίς badge
  function createLowCard(row) {
    return `
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin-bottom:12px;">
        <tr>
          <td style="padding:16px 14px; border:1px solid #edf1f5; border-radius:12px; background-color:#fffaf3;">
            <div style="font-size:15px; font-weight:700; color:#111827;">
              ${escapeHtml(row.item)}
            </div>
            <div style="margin-top:5px; font-size:13px; line-height:1.5; color:#b45309;">
              Το απόθεμα έφτασε στο όριο.
            </div>
          </td>
        </tr>
      </table>
    `;
  }

  function createExpiringMedicineCard(row) {
    return `
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin-bottom:12px;">
        <tr>
          <td style="padding:16px 14px; border:1px solid #edf1f5; border-radius:12px; background-color:#fffaf3;">
            <div style="margin-bottom:8px;">
              ${createPill("ΛΗΓΕΙ ΣΥΝΤΟΜΑ", "#fef3c7", "#b45309")}
            </div>
            <div style="font-size:15px; font-weight:700; color:#111827;">
              ${escapeHtml(row.item)}
            </div>
            <div style="margin-top:5px; font-size:13px; line-height:1.5; color:#b45309;">
              Λήγει στις: ${escapeHtml(row.date)}
            </div>
          </td>
        </tr>
      </table>
    `;
  }

  function createExpiredMedicineCard(row) {
    return `
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin-bottom:12px;">
        <tr>
          <td style="padding:16px 14px; border:1px solid #edf1f5; border-radius:12px; background-color:#fff7f7;">
            <div style="margin-bottom:8px;">
              ${createPill("ΕΛΗΞΕ", "#fee2e2", "#b91c1c")}
            </div>
            <div style="font-size:15px; font-weight:700; color:#111827;">
              ${escapeHtml(row.item)}
            </div>
            <div style="margin-top:5px; font-size:13px; line-height:1.5; color:#991b1b;">
              Έληξε στις: ${escapeHtml(row.date)}
            </div>
          </td>
        </tr>
      </table>
    `;
  }

  function createMaintenanceCard(row) {
    const daysText = row.daysLeft < 0
      ? `Έχει καθυστερήσει: ${Math.abs(row.daysLeft)} ${Math.abs(row.daysLeft) === 1 ? "μέρα" : "μέρες"}`
      : `Απομένουν: ${row.daysLeft} ${row.daysLeft === 1 ? "μέρα" : "μέρες"}`;

    return `
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin-bottom:12px;">
        <tr>
          <td style="padding:16px 14px; border:1px solid #edf1f5; border-radius:12px; background-color:#f8fafc;">
            <div style="margin-bottom:8px;">
              ${createPill("ΠΡΟΓΡΑΜΜΑΤΙΣΜΟΣ", "#dbeafe", "#1d4ed8")}
            </div>
            <div style="font-size:15px; font-weight:700; color:#111827;">
              ${escapeHtml(row.item)}
            </div>
            <div style="margin-top:5px; font-size:13px; line-height:1.5; color:#475569;">
              Επόμενη συντήρηση: ${escapeHtml(row.date)}
            </div>
            <div style="margin-top:4px; font-size:13px; line-height:1.5; color:#475569;">
              ${escapeHtml(daysText)}
            </div>
          </td>
        </tr>
      </table>
    `;
  }

  function createUrgentSection() {
    if (urgentItems.length === 0) {
      return "";
    }

    let itemsHtml = "";

    urgentItems.forEach(row => {
      if (row.type === "expiredMedicine") {
        itemsHtml += `
          <tr>
            <td style="padding:14px 0 0 0;">
              <div style="margin-bottom:8px;">
                ${createPill("ΕΛΗΞΕ", "#fee2e2", "#b91c1c")}
              </div>
              <div style="font-size:15px; font-weight:700; color:#111827;">
                ${escapeHtml(row.item)}
              </div>
              <div style="margin-top:5px; font-size:13px; line-height:1.5; color:#991b1b;">
                Έληξε στις: ${escapeHtml(row.date)}
              </div>
            </td>
          </tr>
        `;
      } else if (row.type === "buy") {
        itemsHtml += `
          <tr>
            <td style="padding:14px 0 0 0;">
              <div style="margin-bottom:8px;">
                ${createPill("ΑΜΕΣΗ ΑΓΟΡΑ", "#fee2e2", "#b91c1c")}
              </div>
              <div style="font-size:15px; font-weight:700; color:#111827;">
                ${escapeHtml(row.item)}
              </div>
              <div style="margin-top:5px; font-size:13px; line-height:1.5; color:#991b1b;">
                ${escapeHtml(row.extra)}
              </div>
            </td>
          </tr>
        `;
      }
    });

    return `
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:24px; border-collapse:separate; border-spacing:0; background-color:#fff7f7; border:1px solid #fecaca; border-radius:16px;">
        <tr>
          <td style="padding:18px 20px;">
            <div style="font-size:16px; font-weight:800; color:#991b1b; margin-bottom:12px;">
              Άμεση προσοχή
            </div>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
              ${itemsHtml}
            </table>
          </td>
        </tr>
      </table>
    `;
  }

  function createSummarySection() {
    return `
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:24px; border-collapse:separate; border-spacing:0; background-color:#ffffff; border:1px solid #e5eaf0; border-radius:16px;">
        <tr>
          <td style="padding:18px 20px;">
            <div style="font-size:15px; font-weight:800; color:#0f172a; margin-bottom:14px;">
              Σύνοψη ενεργειών
            </div>

            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
              <tr>
                <td style="padding:8px 0; font-size:13px; color:#475569; border-bottom:1px solid #eef2f7;">
                  Άμεσες ενέργειες
                </td>
                <td align="right" style="padding:8px 0; font-size:13px; font-weight:700; color:#111827; border-bottom:1px solid #eef2f7;">
                  ${totalUrgent}
                </td>
              </tr>
              <tr>
                <td style="padding:8px 0; font-size:13px; color:#475569; border-bottom:1px solid #eef2f7;">
                  Προς αγορά
                </td>
                <td align="right" style="padding:8px 0; font-size:13px; font-weight:700; color:#111827; border-bottom:1px solid #eef2f7;">
                  ${totalBuy}
                </td>
              </tr>
              <tr>
                <td style="padding:8px 0; font-size:13px; color:#475569; border-bottom:1px solid #eef2f7;">
                  Χαμηλό απόθεμα
                </td>
                <td align="right" style="padding:8px 0; font-size:13px; font-weight:700; color:#111827; border-bottom:1px solid #eef2f7;">
                  ${totalLow}
                </td>
              </tr>
              <tr>
                <td style="padding:8px 0; font-size:13px; color:#475569;">
                  Προσεχείς συντηρήσεις
                </td>
                <td align="right" style="padding:8px 0; font-size:13px; font-weight:700; color:#111827;">
                  ${totalUpcomingMaintenance}
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    `;
  }

  // Εδώ μπαίνουν οι υποκατηγορίες "Προς αγορά" και "Χαμηλό απόθεμα"
  function createInventorySection(icon, title, buyItems, lowItems) {
    const totalItems = buyItems.length + lowItems.length;
    const totalLabel = countLabel(totalItems, "εκκρεμότητα", "εκκρεμότητες");

    let innerHtml = "";

    if (buyItems.length > 0) {
      innerHtml += `
        <div style="font-size:14px; font-weight:800; color:#991b1b; margin-bottom:12px;">
          Προς αγορά
        </div>
      `;
      buyItems.forEach(row => {
        innerHtml += createBuyCard(row);
      });
    }

    if (lowItems.length > 0) {
      innerHtml += `
        <div style="font-size:14px; font-weight:800; color:#b45309; margin-bottom:12px; margin-top:${buyItems.length > 0 ? "16px" : "0"};">
          Χαμηλό απόθεμα
        </div>
      `;
      lowItems.forEach(row => {
        innerHtml += createLowCard(row);
      });
    }

    if (totalItems === 0) {
      innerHtml = createEmptyBox("Δεν υπάρχουν εκκρεμότητες σε αυτή την ενότητα.");
    }

    return `
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:24px; border-collapse:separate; border-spacing:0; background-color:#ffffff; border:1px solid #e5eaf0; border-radius:16px;">
        <tr>
          <td style="padding:16px 20px; background-color:#f8fafc; border-bottom:1px solid #e5eaf0;">
            <div style="font-size:18px; font-weight:800; color:#0f172a;">
              ${icon} ${escapeHtml(title)} — ${totalItems} ${totalLabel}
            </div>
          </td>
        </tr>
        <tr>
          <td style="padding:20px;">
            ${innerHtml}
          </td>
        </tr>
      </table>
    `;
  }

  function createPharmacySection() {
    const totalItems = report.pharmacy.expiringSoon.length + report.pharmacy.expired.length;
    const totalLabel = countLabel(totalItems, "καταχώριση", "καταχωρίσεις");

    let html = "";

    if (report.pharmacy.expiringSoon.length > 0) {
      html += `
        <div style="font-size:14px; font-weight:800; color:#92400e; margin-bottom:12px;">
          Λήγουν σύντομα
        </div>
      `;
      report.pharmacy.expiringSoon.forEach(row => {
        html += createExpiringMedicineCard(row);
      });
    }

    if (report.pharmacy.expired.length > 0) {
      html += `
        <div style="font-size:14px; font-weight:800; color:#991b1b; margin-bottom:12px; margin-top:${report.pharmacy.expiringSoon.length > 0 ? "16px" : "0"};">
          Έχουν λήξει
        </div>
      `;
      report.pharmacy.expired.forEach(row => {
        html += createExpiredMedicineCard(row);
      });
    }

    if (totalItems === 0) {
      html = createEmptyBox("Δεν υπάρχουν φάρμακα που λήγουν ή έχουν λήξει.");
    }

    return `
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:24px; border-collapse:separate; border-spacing:0; background-color:#ffffff; border:1px solid #e5eaf0; border-radius:16px;">
        <tr>
          <td style="padding:16px 20px; background-color:#f8fafc; border-bottom:1px solid #e5eaf0;">
            <div style="font-size:18px; font-weight:800; color:#0f172a;">
              💊 Φαρμακείο — ${totalItems} ${totalLabel}
            </div>
          </td>
        </tr>
        <tr>
          <td style="padding:20px;">
            ${html}
          </td>
        </tr>
      </table>
    `;
  }

  function createMaintenanceSection() {
    const totalItems = report.maintenance.upcoming.length;
    const totalLabel = countLabel(totalItems, "καταχώριση", "καταχωρίσεις");

    let html = "";

    if (totalItems === 0) {
      html = createEmptyBox("Δεν υπάρχουν προσεχείς συντηρήσεις στο επιλεγμένο χρονικό όριο.");
    } else {
      report.maintenance.upcoming.forEach(row => {
        html += createMaintenanceCard(row);
      });
    }

    return `
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:0; border-collapse:separate; border-spacing:0; background-color:#ffffff; border:1px solid #e5eaf0; border-radius:16px;">
        <tr>
          <td style="padding:16px 20px; background-color:#f8fafc; border-bottom:1px solid #e5eaf0;">
            <div style="font-size:18px; font-weight:800; color:#0f172a;">
              🛠️ Συντήρηση — ${totalItems} ${totalLabel}
            </div>
          </td>
        </tr>
        <tr>
          <td style="padding:20px;">
            ${html}
          </td>
        </tr>
      </table>
    `;
  }

  const htmlBody = `
<!DOCTYPE html>
<html lang="el">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Αναφορά Υλικών Εκκλησίας</title>
</head>
<body style="margin:0; padding:0; background-color:#eef2f7; font-family:Arial, Helvetica, sans-serif;">

  <div style="width:100%; background-color:#eef2f7; padding:30px 0;">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
      <tr>
        <td align="center" style="padding:0 16px;">

          <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width:760px; border-collapse:separate; border-spacing:0; background-color:#ffffff; border-radius:22px; overflow:hidden; box-shadow:0 8px 30px rgba(15,23,42,0.08);">

            <tr>
              <td style="background-color:#1565c0; padding:30px 24px; color:#ffffff;">
                <div style="font-size:26px; font-weight:800; line-height:1.2;">
                  Αναφορά Υλικών Εκκλησίας
                </div>
                <div style="margin-top:8px; font-size:14px; line-height:1.5; color:#dbeafe;">
                  Ημερομηνία αναφοράς: ${reportDate}
                </div>

                <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin-top:20px; border-collapse:collapse;">
                  <tr>
                    <td style="padding:0 10px 10px 0;">
                      <div style="display:inline-block; background-color:rgba(255,255,255,0.16); padding:10px 14px; border-radius:12px; font-size:13px; font-weight:700; color:#ffffff;">
                        Προς αγορά: ${totalBuy}
                      </div>
                    </td>
                    <td style="padding:0 10px 10px 0;">
                      <div style="display:inline-block; background-color:rgba(255,255,255,0.16); padding:10px 14px; border-radius:12px; font-size:13px; font-weight:700; color:#ffffff;">
                        Χαμηλό απόθεμα: ${totalLow}
                      </div>
                    </td>
                    <td style="padding:0 10px 10px 0;">
                      <div style="display:inline-block; background-color:rgba(255,255,255,0.16); padding:10px 14px; border-radius:12px; font-size:13px; font-weight:700; color:#ffffff;">
                        Ληγμένα φάρμακα: ${totalExpired}
                      </div>
                    </td>
                    <td style="padding:0 0 10px 0;">
                      <div style="display:inline-block; background-color:rgba(255,255,255,0.16); padding:10px 14px; border-radius:12px; font-size:13px; font-weight:700; color:#ffffff;">
                        Συντηρήσεις: ${totalUpcomingMaintenance}
                      </div>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <tr>
              <td style="padding:24px; background-color:#f6f8fb;">

                ${createUrgentSection()}

                ${createSummarySection()}

                ${createInventorySection(
                  "🧼",
                  "Καθαριότητα",
                  report.cleanliness.buy,
                  report.cleanliness.low
                )}

                ${createInventorySection(
                  "🛍️",
                  "Τράπεζα Αγάπης",
                  report.loveBank.buy,
                  report.loveBank.low
                )}

                ${createPharmacySection()}

                ${createMaintenanceSection()}

              </td>
            </tr>

            <tr>
              <td style="padding:18px 24px; background-color:#ffffff; border-top:1px solid #e5eaf0;">
                <div style="font-size:12px; line-height:1.7; color:#6b7280;">
                  Η αναφορά δημιουργήθηκε αυτόματα στις ${reportDate} από το σύστημα διαχείρισης υλικών της εκκλησίας.
                </div>
              </td>
            </tr>

          </table>

        </td>
      </tr>
    </table>
  </div>

</body>
</html>
  `;

  let plainBody = `ΑΝΑΦΟΡΑ ΥΛΙΚΩΝ ΕΚΚΛΗΣΙΑΣ - ${reportDate}\n\n`;

  plainBody += `ΣΥΝΟΨΗ\n`;
  plainBody += `- Άμεσες ενέργειες: ${totalUrgent}\n`;
  plainBody += `- Προς αγορά: ${totalBuy}\n`;
  plainBody += `- Χαμηλό απόθεμα: ${totalLow}\n`;
  plainBody += `- Ληγμένα φάρμακα: ${totalExpired}\n`;
  plainBody += `- Συντηρήσεις: ${totalUpcomingMaintenance}\n\n`;

  plainBody += `ΚΑΘΑΡΙΟΤΗΤΑ\n`;
  if (report.cleanliness.buy.length === 0 && report.cleanliness.low.length === 0) {
    plainBody += `- Δεν υπάρχουν εκκρεμότητες\n`;
  } else {
    if (report.cleanliness.buy.length > 0) {
      plainBody += `Προς αγορά:\n`;
      report.cleanliness.buy.forEach(row => {
        plainBody += `- ${row.item} | Λείπουν: ${row.missing}\n`;
      });
    }
    if (report.cleanliness.low.length > 0) {
      plainBody += `Χαμηλό απόθεμα:\n`;
      report.cleanliness.low.forEach(row => {
        plainBody += `- ${row.item}\n`;
      });
    }
  }

  plainBody += `\nΤΡΑΠΕΖΑ ΑΓΑΠΗΣ\n`;
  if (report.loveBank.buy.length === 0 && report.loveBank.low.length === 0) {
    plainBody += `- Δεν υπάρχουν εκκρεμότητες\n`;
  } else {
    if (report.loveBank.buy.length > 0) {
      plainBody += `Προς αγορά:\n`;
      report.loveBank.buy.forEach(row => {
        plainBody += `- ${row.item} | Λείπουν: ${row.missing}\n`;
      });
    }
    if (report.loveBank.low.length > 0) {
      plainBody += `Χαμηλό απόθεμα:\n`;
      report.loveBank.low.forEach(row => {
        plainBody += `- ${row.item}\n`;
      });
    }
  }

  plainBody += `\nΦΑΡΜΑΚΕΙΟ\n`;
  if (report.pharmacy.expiringSoon.length === 0 && report.pharmacy.expired.length === 0) {
    plainBody += `- Δεν υπάρχουν καταχωρήσεις\n`;
  } else {
    report.pharmacy.expiringSoon.forEach(row => {
      plainBody += `- Λήγει σύντομα: ${row.item} | ${row.date}\n`;
    });
    report.pharmacy.expired.forEach(row => {
      plainBody += `- Έληξε: ${row.item} | ${row.date}\n`;
    });
  }

  plainBody += `\nΣΥΝΤΗΡΗΣΗ\n`;
  if (report.maintenance.upcoming.length === 0) {
    plainBody += `- Δεν υπάρχουν προσεχείς συντηρήσεις\n`;
  } else {
    report.maintenance.upcoming.forEach(row => {
      plainBody += `- ${row.item} | ${row.date} | Απομένουν: ${row.daysLeft} μέρες\n`;
    });
  }

  MailApp.sendEmail({
    to: recipient,
    subject: "Αναφορά Υλικών Εκκλησίας",
    body: plainBody,
    htmlBody: htmlBody
  });
}
