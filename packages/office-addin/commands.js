/* global Office */
Office.onReady(function () {
  // Commands host initialized
});

function openDashboard(event) {
  try {
    Office.context.ui.displayDialogAsync(
      'https://127.0.0.1:18443/dashboard',
      { height: 70, width: 70 }
    );
  } catch (e) {
    // ignore
  }
  if (event && typeof event.completed === 'function') {
    event.completed();
  }
}

if (typeof Office !== 'undefined' && Office.actions) {
  Office.actions.associate('openDashboard', openDashboard);
}
