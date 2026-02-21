// 新しいページが読み込まれるたびに初期化関数を実行する必要があります。
(function () {
    Office.initialize = function (reason) {
        // 必要な初期化は、ここで実行できます。
    };
})();

var NewLine = "\n";

function copyToMarkdown(event) {
    Excel.run(function (ctx) {
        var cells = [];
        var range = ctx.workbook.getSelectedRange().load(["rowCount", "columnCount"]);
        return ctx.sync()
            .then(function () {
                for (var row = 0; row < range.rowCount; row++) {
                    for (var col = 0; col < range.columnCount; col++) {
                        cells.push(range.getCell(row, col).load(["text", "format"]));
                    }
                }
            })
            .then(ctx.sync)
            .then(function() {
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

                // Copy to clipboard using modern API
                if (navigator.clipboard && navigator.clipboard.writeText) {
                    // Modern browsers
                    navigator.clipboard.writeText(result).then(function() {
                        console.log("Copied to clipboard");
                        event.completed();
                    }).catch(function(err) {
                        console.error("Clipboard error:", err);
                        // Fallback: show result in dialog
                        Office.context.ui.displayDialogAsync(
                            'data:text/html,' + encodeURIComponent('<textarea style="width:100%;height:100%">' + result + '</textarea>'),
                            {height: 50, width: 50}
                        );
                        event.completed();
                    });
                } else if (window.clipboardData) {
                    // IE fallback
                    window.clipboardData.setData("Text", result);
                    event.completed();
                } else {
                    // Last resort: create temporary textarea
                    var textarea = document.createElement('textarea');
                    textarea.value = result;
                    textarea.style.position = 'fixed';
                    textarea.style.opacity = '0';
                    document.body.appendChild(textarea);
                    textarea.select();
                    try {
                        document.execCommand('copy');
                        console.log("Copied using execCommand");
                    } catch (err) {
                        console.error("Copy failed:", err);
                    }
                    document.body.removeChild(textarea);
                    event.completed();
                }
            });
    }).catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
        event.completed({allowEvent: false});
    });
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
