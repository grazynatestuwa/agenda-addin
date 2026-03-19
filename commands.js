Office.initialize = function () {};

// Ta funkcja jest wywoływana przy każdej próbie wysłania spotkania
function validateAgenda(event) {
  const item = Office.context.mailbox.item;

  // Działa tylko dla spotkań (nie zwykłych maili)
  if (item.itemType !== Office.MailboxEnums.ItemType.Appointment) {
    event.completed({ allowEvent: true });
    return;
  }

  // Pobieramy opis spotkania (body)
  item.body.getAsync(Office.CoercionType.Text, function (result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      // Jeśli nie możemy odczytać — przepuszczamy
      event.completed({ allowEvent: true });
      return;
    }

    const body = result.value.trim();

    // === TWOJA LOGIKA WALIDACJI ===
    const hasAgenda = checkForAgenda(body);

    if (hasAgenda) {
      // Agenda jest — pozwól wysłać
      event.completed({ allowEvent: true });
    } else {
      // Brak agendy — BLOKUJ i pokaż komunikat
      item.notificationMessages.addAsync("agenda-error", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: "❌ Brak agendy! Dodaj cel spotkania w opisie (np. 'Cel:', 'Agenda:', 'Temat:')."
      });

      event.completed({ allowEvent: false }); // ← to blokuje wysyłkę
    }
  });
}

function checkForAgenda(bodyText) {
  if (!bodyText || bodyText.length < 20) return false;

  // Słowa kluczowe — dostosuj do potrzeb firmy
  const keywords = [
    "agenda:",
    "cel:",
    "cel spotkania",
    "topics:",
    "temat:",
    "tematyka:",
    "omówimy",
    "omawiane punkty",
    "porządek obrad",
    "plan spotkania",
  ];

  const lower = bodyText.toLowerCase();
  return keywords.some(keyword => lower.includes(keyword));
}

// Rejestracja handlera
Office.actions.associate("validateAgenda", validateAgenda);
