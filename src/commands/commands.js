/* global Office */

function openConfirmDialog(event) {

  const url =
    "https://dellasiegaexternal2.github.io/Herisson24022026/src/confirm.html?mode=dialog&v=300";

  Office.context.ui.displayDialogAsync(
    url,
    {
      height: 60,
      width: 50,
      displayInIframe: true
    }
  );

  event.completed();
}

// obligatoire pour ExecuteFunction
if (typeof window !== "undefined") {
  window.openConfirmDialog = openConfirmDialog;
}
