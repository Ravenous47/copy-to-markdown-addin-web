// 新しいページが読み込まれるたびに初期化関数を実行する必要があります。
(function () {
    Office.initialize = function (reason) {
        // 必要な初期化は、ここで実行できます。
    };
})();

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
                var resultBuffer = new StringBuilder();
                var separatorBuffer = new StringBuilder();
                for (var x = 0; x < range.columnCount; x++)
                {
                    var cell = cells[x];

                    resultBuffer.append("|");
                    resultBuffer.append(formatText(cell.text));
                    switch (cell.format.horizontalAlignment)
                    {
                        case "Center":
                            separatorBuffer.append("|:-:");
                            break;
                        case "Right":
                            separatorBuffer.append("|--:");
                            break;
                        default:
                            separatorBuffer.append("|:--");
                            break;
                    }
                }
                resultBuffer.append("|");
                resultBuffer.append(NewLine);
                separatorBuffer.append("|");
                separatorBuffer.append(NewLine);
                resultBuffer.append(separatorBuffer.toString());

                for (var row = 1; row < range.rowCount; row++)
                {
                    for (var col = 0; col < range.columnCount; col++)
                    {
                        var valueCell = cells[row * range.columnCount + col];

                        resultBuffer.append("|");
                        resultBuffer.append(formatText(valueCell.text));
                    }
                    resultBuffer.append("|");
                    resultBuffer.append(NewLine);
                }

                var result = resultBuffer.toString();
                console.log("Generated markdown (" + result.length + " chars)");
                console.log("Markdown preview:", result.substring(0, 100));

                // Use Office.js built-in method to set clipboard data
                Office.context.document.setSelectedDataAsync(result, {
                    coercionType: Office.CoercionType.Text
                }, function(asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        console.log("✓ Copied to clipboard successfully");
                        // Show notification
                        showNotification("Success", "Markdown copied to clipboard!");
                        event.completed();
                    } else {
                        console.error("✗ Clipboard failed:", asyncResult.error.message);
                        // Fallback: Try navigator.clipboard
                        tryModernClipboard(result, event);
                    }
                });
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

function tryModernClipboard(result, event) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(result).then(function() {
            console.log("✓ Copied using navigator.clipboard");
            showNotification("Success", "Copied to clipboard!");
            event.completed();
        }).catch(function(err) {
            console.error("✗ navigator.clipboard failed:", err);
            tryTextareaFallback(result, event);
        });
    } else {
        tryTextareaFallback(result, event);
    }
}

function tryTextareaFallback(result, event) {
    var textarea = document.createElement('textarea');
    textarea.value = result;
    textarea.style.position = 'fixed';
    textarea.style.left = '-999999px';
    document.body.appendChild(textarea);
    textarea.select();

    var success = false;
    try {
        success = document.execCommand('copy');
        console.log("execCommand copy:", success ? "✓ success" : "✗ failed");
    } catch (err) {
        console.error("✗ execCommand failed:", err);
    }

    document.body.removeChild(textarea);

    if (success) {
        showNotification("Success", "Copied using fallback method!");
    } else {
        showNotification("Error", "Could not copy. Please use Ctrl+C manually.");
        // Select the data in Excel so user can copy manually
        console.log("Clipboard unavailable. Result:", result);
    }

    event.completed();
}

function showNotification(title, message) {
    console.log("NOTIFICATION:", title, "-", message);
    // Try to show Office notification
    if (Office.context.ui && Office.context.ui.displayDialogAsync) {
        var html = '<html><body style="font-family:Arial;padding:20px;"><h2>' +
                   title + '</h2><p>' + message + '</p></body></html>';
        Office.context.ui.displayDialogAsync(
            'data:text/html,' + encodeURIComponent(html),
            {height: 30, width: 40}
        );
    }
}
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
