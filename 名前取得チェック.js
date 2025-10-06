function debugWhoAmI_() {
  const name = getActiveUserDisplayName_();
  const email = Session.getActiveUser().getEmail();
  SpreadsheetApp.getUi().alert('表示名='+ name + '\nメール='+ email);
}