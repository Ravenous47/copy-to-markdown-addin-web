// 新しいページが読み込まれるたびに初期化関数を実行する必要があります。
console.log("FunctionFile.js loaded!");

(function () {
    Office.initialize = function (reason) {
        console.log("Office.initialize called, reason:", reason);
        // 必要な初期化は、ここで実行できます。
    };
})();

// Register functions globally
if (typeof window !== 'undefined') {
    window.copyToMarkdown = copyToMarkdown;
    window.pasteFromMarkdown = pasteFromMarkdown;
    console.log("Functions registered globally");
}

// Also register with Office
if (typeof Office !== 'undefined' && Office.actions) {
    Office.actions.associate("copyToMarkdown", copyToMarkdown);
    Office.actions.associate("pasteFromMarkdown", pasteFromMarkdown);
    console.log("Functions registered with Office.actions");
}

var NewLine = "\n";

function copyToMarkdown(event) {
    console.log("copyToMarkdown called");

    Excel.run(function (ctx) {
        var cells = [];
        var range = ctx.workbook.getSelectedRange().load(["rowCount", "columnCount"]);
        return ctx.sync()
            .then(function () {
                console.log("Range loaded: " + range.rowCount + "x" + range.columnCount);
                for (var row = 0; row < range.rowCount; row++) {
                    for (var col = 0; col < range.columnCount; col++) {
                        cells.push(range.getCell(row, col).load(["text", "format"]));
                    }
                }
            })
            .then(ctx.sync)
            .then(function() {
                console.log("Cells loaded: " + cells.length);

                // Convert cells to 2D array for markdown-table library
                var tableData = [];
                for (var row = 0; row < range.rowCount; row++) {
                    var rowData = [];
                    for (var col = 0; col < range.columnCount; col++) {
                        var cell = cells[row * range.columnCount + col];
                        rowData.push(formatText(cell.text));
                    }
                    tableData.push(rowData);
                }

                // Use markdown-table library to generate markdown
                var result;
                if (typeof markdownTable !== 'undefined') {
                    result = markdownTable(tableData);
                    console.log("Generated markdown using markdownTable library");
                } else {
                    // Fallback to manual generation if library not loaded
                    console.warn("markdown-table library not loaded, using fallback");
                    result = generateMarkdownManually(tableData);
                }

                console.log("Generated markdown (" + result.length + " chars)");
                console.log("Markdown preview:", result.substring(0, 100));

                // Try clipboard APIs first, fall back to dialog
                tryAllClipboardMethods(result, event);
            });
    }).catch(function (error) {
        console.error("Error in copyToMarkdown:", error);
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug info:", JSON.stringify(error.debugInfo));
        }
        showNotification("Error", error.toString());
        event.completed({allowEvent: false});
    });
}

function tryAllClipboardMethods(result, event) {
    console.log("Attempting clipboard copy...");

    // Method 1: Try modern Clipboard API with focus
    if (navigator.clipboard && navigator.clipboard.writeText) {
        window.focus();
        navigator.clipboard.writeText(result).then(function() {
            console.log("✓ Copied using navigator.clipboard!");
            event.completed();
        }).catch(function(err) {
            console.log("navigator.clipboard failed:", err.message);
            tryTextareaMethod(result, event);
        });
    } else {
        console.log("navigator.clipboard not available");
        tryTextareaMethod(result, event);
    }
}

function tryTextareaMethod(result, event) {
    console.log("Trying textarea method...");
    // Force focus on window first
    window.focus();

    var textarea = document.createElement('textarea');
    textarea.value = result;
    textarea.style.position = 'fixed';
    textarea.style.top = '0';
    textarea.style.left = '0';
    textarea.style.width = '2em';
    textarea.style.height = '2em';
    textarea.style.padding = '0';
    textarea.style.border = 'none';
    textarea.style.outline = 'none';
    textarea.style.boxShadow = 'none';
    textarea.style.background = 'transparent';
    document.body.appendChild(textarea);

    // Give it focus explicitly
    textarea.focus();
    textarea.select();
    textarea.setSelectionRange(0, textarea.value.length);

    var success = false;
    try {
        success = document.execCommand('copy');
        console.log("execCommand result:", success);
    } catch (err) {
        console.error("execCommand error:", err);
    }

    document.body.removeChild(textarea);

    if (success) {
        console.log("✓ Copied using execCommand!");
        event.completed();
    } else {
        console.log("All clipboard methods failed, showing dialog");
        showMarkdownDialog(result, event);
    }
}

function showMarkdownDialog(markdown, event) {
    console.log("showMarkdownDialog called");

    // Use hosted dialog page with markdown as URL parameter
    var dialogUrl = 'https://ravenous47.github.io/copy-to-markdown-addin-web/dialog.html?md=' +
                    encodeURIComponent(markdown);

    console.log("Opening dialog at:", dialogUrl);

    try {
        Office.context.ui.displayDialogAsync(
            dialogUrl,
            { height: 50, width: 40 },
            function(result) {
                console.log("Dialog callback received");
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("✓ Dialog opened successfully");
                } else {
                    console.error("✗ Dialog failed:", result.error);
                    console.log("\n=== MARKDOWN OUTPUT ===\n" + markdown + "\n======================\n");
                }
                event.completed();
            }
        );
    } catch (err) {
        console.error("✗ Exception opening dialog:", err);
        console.log("\n=== MARKDOWN OUTPUT ===\n" + markdown + "\n======================\n");
        event.completed();
    }
}

function escapeHtml(text) {
    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function showNotification(title, message) {
    console.log("NOTIFICATION:", title, "-", message);
}

function formatText(range)
{
    if (range == undefined) {
        return "";
    }
    else {
        return range.join().replace(NewLine, "<br>");
    }
}

function generateMarkdownManually(tableData) {
    // Fallback manual markdown generation
    var result = '';

    if (tableData.length === 0) return result;

    // Header row
    result += '| ' + tableData[0].join(' | ') + ' |\n';

    // Separator row
    var separators = tableData[0].map(function() { return '---'; });
    result += '| ' + separators.join(' | ') + ' |\n';

    // Data rows
    for (var i = 1; i < tableData.length; i++) {
        result += '| ' + tableData[i].join(' | ') + ' |\n';
    }

    return result;
}


var StringBuilder = function (string) {
    this.buffer = [];

    this.append = function (string) {
        this.buffer.push(string);
        return this;
    };

    this.toString = function () {
        return this.buffer.join('');
    };

    if (string) {
        this.append(string);
    }
};

// Paste from Markdown function
function pasteFromMarkdown(event) {
    console.log("pasteFromMarkdown called");

    var dialogUrl = 'https://ravenous47.github.io/copy-to-markdown-addin-web/paste-dialog.html';

    try {
        Office.context.ui.displayDialogAsync(
            dialogUrl,
            { height: 60, width: 50 },
            function(result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("✓ Paste dialog opened");
                    var dialog = result.value;

                    // Handle messages from dialog
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                        console.log("Message from dialog:", arg.message);

                        try {
                            var message = JSON.parse(arg.message);
                            console.log("Parsed message:", message);
                            console.log("Message action:", message.action);
                            console.log("Message has data:", !!message.data);

                            if (message.action === 'insert' && message.data) {
                                console.log("Received data to insert:", message.data);
                                console.log("Closing dialog...");
                                dialog.close();

                                console.log("About to call insertDataIntoExcel...");
                                // Insert data into Excel
                                insertDataIntoExcel(message.data, event);
                            } else if (arg.message === 'cancel') {
                                console.log("User cancelled");
                                dialog.close();
                                event.completed();
                            } else {
                                console.log("Unknown action or no data");
                                dialog.close();
                                event.completed();
                            }
                        } catch (parseErr) {
                            console.error("Error in message handler:", parseErr);
                            console.error("Parse error details:", parseErr.message);
                            // Not JSON, might be simple message
                            if (arg.message === 'cancel') {
                                console.log("User cancelled");
                                dialog.close();
                                event.completed();
                            } else {
                                console.error("Error parsing message:", parseErr);
                                dialog.close();
                                event.completed();
                            }
                        }
                    });
                } else {
                    console.error("✗ Failed to open paste dialog:", result.error);
                    event.completed();
                }
            }
        );
    } catch (err) {
        console.error("✗ Exception opening paste dialog:", err);
        event.completed();
    }
}

function insertDataIntoExcel(data, event) {
    console.log("insertDataIntoExcel called, rows:", data.length);

    Excel.run(function(ctx) {
        var sheet = ctx.workbook.worksheets.getActiveWorksheet();
        var range = ctx.workbook.getSelectedRange();

        // Load the range to get starting position
        range.load('address');

        return ctx.sync().then(function() {
            console.log('Inserting at:', range.address);

            // Calculate dimensions
            var numRows = data.length;
            var numCols = data[0] ? data[0].length : 0;
            console.log('Dimensions:', numRows, 'rows x', numCols, 'cols');

            // Get range for data
            var targetRange = range.getResizedRange(numRows - 1, numCols - 1);

            // Set values
            targetRange.values = data;

            return ctx.sync();
        });
    }).then(function() {
        console.log("✓ Data inserted successfully!");
        event.completed();
    }).catch(function(error) {
        console.error("✗ Error inserting data:", error);
        if (error.debugInfo) {
            console.error("Debug info:", JSON.stringify(error.debugInfo));
        }
        event.completed();
    });
}
