/**
 * Fonction demandant une entrée à l'utilisateur en fonction du type attendu
 * @param {string} Phrase - La question posée à l'utilisateur
 * @param {string} data - Le type de données attendu : "numérique", "lettre", "lettre&numérique", "touttype"
 * @return {string|number|null} - Retourne la valeur saisie par l'utilisateur si elle est valide, sinon affiche une erreur
 */
function demanderValeur(phrase, data) {
    var ui = SpreadsheetApp.getUi(); // Interface utilisateur de Google Sheets
    var valeur; // Variable pour stocker la valeur saisie
    var regex; // Variable qui stockera l'expression régulière de validation
  
    // Définition du regex selon le type de données attendu
    switch (data.toLowerCase()) {
      case "numérique":
        regex = /^[0-9]+$/; // Accepte uniquement les chiffres
        break;
      case "lettre":
        regex = /^[a-zA-ZÀ-ÿ\s]+$/; // Accepte uniquement les lettres et les espaces
        break;
      case "lettre&numérique":
        regex = /^[a-zA-Z0-9À-ÿ\s]+$/; // Accepte lettres, chiffres et espaces
        break;
      case "touttype":
        regex = /.+/; // Accepte tout type de données
        break;
      default:
        ui.alert("❌ Erreur : Type de données inconnu. Veuillez choisir parmi : 'numérique', 'lettre', 'lettre&numérique', 'touttype'.");
        return null;
    }
  
    // Boucle pour redemander la valeur si elle est incorrecte
    do {
      var response = ui.prompt(phrase);
      
      if (response.getSelectedButton() == ui.Button.CANCEL) {
        return null; // Si l'utilisateur annule, on retourne null
      }
  
      valeur = response.getResponseText().trim(); // Récupère la réponse et enlève les espaces inutiles
  
      // Vérification si la valeur correspond au type attendu
      if (!regex.test(valeur)) {
        ui.alert("⚠️ Erreur : La saisie ne correspond pas au type attendu (" + data + "). Veuillez réessayer.");
        valeur = null; // Réinitialise la valeur pour relancer la boucle
      }
  
    } while (valeur === null);
  
    return valeur; // Retourne la valeur valide
  }