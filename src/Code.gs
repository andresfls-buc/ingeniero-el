const AUTHORIZED_EMAILS = ['inelcoingeniero@gmail.com'];

function doGet() {
  const email = Session.getActiveUser().getEmail();
  if (!AUTHORIZED_EMAILS.includes(email)) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:60px;text-align:center">' +
      '<h2 style="color:#c62828">Acceso denegado</h2>' +
      '<p style="color:#555;margin-top:12px">No tienes permiso para acceder a esta aplicación.</p>' +
      '</div>'
    );
  }
  return HtmlService.createHtmlOutputFromFile("index");
}
