/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getDataFromSelection);
            $('#btnEcho').click(echoText);
            $('#btnEchoNamedCell').click(echoTextNamedCell);
            $('#btnPopulateMatrix').click(populateMatrix);
        });
    };

    // Main function and callback function for simple text insert
    function echoText() {
        var enteredText = $('#echoText').val();
        Office.context.document.setSelectedDataAsync(enteredText, messageCallback);
    }

    function messageCallback(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            app.showNotification("Message sent successfully!", "Success");
        }
        else {
            app.showNotification("Error sending message:", result.error.message);
        }
    }

    // Main function for insert into a named cell
    function echoTextNamedCell() {
        var enteredText = $('#echoTextNamedCell').val();

        // Bind to the cell, assumes a single cell named "SingleCellB2" (doesn't have to be the B2 cell)
        Office.context.document.bindings.addFromNamedItemAsync("SingleCellB2", Office.BindingType.Text, { id: "SingleCell" },
            function (asyncResult) {
                if (asyncResult.status == 'failed') {
                    app.showNotification("Error sending message:", asyncResult.error.message);
                }
                else {
                    // Write the data, using generic messageCallback function
                    Office.select("bindings#SingleCell").setDataAsync(enteredText, { coercionType: Office.CoercionType.Text }, messageCallback);
                }
            });
    }

    // Main function and callback to populate multiple cells using a matrix
    function populateMatrix() {
        // Bind to the range, assumes four cells named ColorMatrix
        Office.context.document.bindings.addFromNamedItemAsync("ColorMatrix", Office.BindingType.Matrix, { id: "myMatrix" },
            function (asyncResult) {
                if (asyncResult.status == 'failed') {
                    app.showNotification("Error sending message:", asyncResult.error.message);
                }
                else {
                    // Write the data, using generic messageCallback function
                    Office.select("bindings#myMatrix").setDataAsync([['1', 'Blue'], ['2', 'Green'], ['3', 'Yellow'], ['4', 'Orange']], { coercionType: Office.CoercionType.Matrix }, messageCallback);
                }
            });
    }

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }
})();