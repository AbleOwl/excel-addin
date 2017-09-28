
=
(function () {
    "use strict";

    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#set-color').click(setColor);
              $('#open-dia').click(openBox);
        });
        
        
        function openBox(){

              if(Office.context.ui.displayDialogAsync!==undefined)
              {
              Office.context.ui.displayDialogAsync('https://google.com');
              }

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
    
})();
