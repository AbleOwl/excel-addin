(function () {
    "use strict";

    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#set-color').click(setColor);
              $('#open-dia').click(openBox);
                $('#delete').click(deleteSheet);
        });
        
     function openBox(){
              Office.context.ui.displayDialogAsync('https://google.com');
        }
    };

    function setColor() {
        Excel.run(function (context) {
            var range = context.workbook.getSelectedRange();
            range.format.fill.color = 'green';
            alert('test');
            return ctx.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    
    function deleteSheet() {
        Excel.run(function (context) {
            var sheet = context.workbook.worksheets.getItem("Sheet1");
          //  var range = context.workbook.getSelectedRange();
        //  var range = sheet.getRange("A1:B3").format.protection.locked = false;
           sheet.delete();

            return ctx.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    
})();
