/**
 * Function that prompts the user for input based on the expected data type.
 * @param {string} question - The question displayed to the user.
 * @param {string} dataType - The expected data type: "numeric", "letter", "alphanumeric", "any".
 * @return {string|number|null} - Returns the user input if valid, otherwise displays an error.
 */
function requestValue(question, dataType) {
    var ui = SpreadsheetApp.getUi(); // Google Sheets User Interface
    var value; // Variable to store the user input
    var regex; // Variable to store the validation regex

    // Define the regex based on the expected data type
    switch (dataType.toLowerCase()) {
        case "numeric":
            regex = /^[0-9]+$/; // Accepts only numbers
            break;
        case "letter":
            regex = /^[a-zA-ZÀ-ÿ\s]+$/; // Accepts only letters and spaces
            break;
        case "alphanumeric":
            regex = /^[a-zA-Z0-9À-ÿ\s]+$/; // Accepts letters, numbers, and spaces
            break;
        case "any":
            regex = /.+/; // Accepts any type of data
            break;
        default:
            ui.alert("❌ Error: Unknown data type. Please choose from: 'numeric', 'letter', 'alphanumeric', 'any'.");
            return null;
    }

    // Loop to re-prompt the user if the input is invalid
    do {
        var response = ui.prompt(question);
        
        if (response.getSelectedButton() == ui.Button.CANCEL) {
            return null; // Return null if the user cancels
        }

        value = response.getResponseText().trim(); // Get user input and trim spaces

        // Validate input against the expected type
        if (!regex.test(value)) {
            ui.alert("⚠️ Error: The input does not match the expected type (" + dataType + "). Please try again.");
            value = null; // Reset value to restart the loop
        }

    } while (value === null);

    return value; // Return the valid input
}