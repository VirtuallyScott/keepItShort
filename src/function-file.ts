Office.onReady(() => {
  // Office is ready
});

let expenseMultiplier: number = 1; // Set a default value

function calculateMeetingCost(eventArgs: Office.AddinCommands.EventArgs) {
  const item = Office.context.mailbox.item as Office.AppointmentCompose;

  item.requiredAttendees.getAsync((result: Office.AsyncResult<Office.EmailAddressDetails[]>) => {
    if (result.status === Office.MailboxEnums.ItemType.Error) {
      console.error(result.error);
      eventArgs.completed({ allowEvent: false });
      return;
    }

    const numberOfAttendees = result.value.length;
    const meetingCost = numberOfAttendees * expenseMultiplier;

    const message: string = `The estimated cost of this meeting is $${meetingCost}.`;

    Office.context.ui.displayDialogAsync(`https://localhost:3000/popup.html?message=${encodeURIComponent(message)}`, { 
height: 30, width: 20 }, (result: Office.AsyncResult<Office.Dialog>) => {
      if (result.status === Office.MailboxEnums.ItemType.Error) {
        console.error(result.error);
        eventArgs.completed({ allowEvent: false });
        return;
      }

      eventArgs.completed({ allowEvent: true });
    });
  });
}

