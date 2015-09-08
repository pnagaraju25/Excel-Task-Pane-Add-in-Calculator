 /// <reference path="../App.js" />

var calComp = true;

function getNumberInput(data) {

    var textBox = document.getElementById("result");

    if (calComp) {
        textBox.innerText = '0';
    }

    if (data.value == '.') {
        if (textBox.innerText == '') {
            textBox.innerText = '0';
        }
    } else if (textBox.innerText == '0') {
        textBox.innerText = '';
    }

    if (data.value == '.' && textBox.innerText.indexOf('.') > 0) {
    } else {
        textBox.innerText += data.value;
    }
}

(function () {
    "use strict";

    var bindingId;
    var currentData = '';
    var currentPosArray = new Array;
    var operatorsArray = new Array;
    var posIndex = 0;
    var operandsIndex = 0;
    var operatorsIndex = 0;
    var textBox;

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            textBox = document.getElementById("result");

            $('#nbOne').click(function (event) {
                getOperand();
            });
            $('#nbTwo').click(function (event) {
                getOperand();
            });
            $('#nbThree').click(function (event) {
                getOperand();
            });
            $('#nbFour').click(function (event) {
                getOperand();
            });
            $('#nbFive').click(function (event) {
                getOperand();
            });
            $('#nbSix').click(function (event) {
                getOperand();
            });
            $('#nbSeven').click(function (event) {
                getOperand();
            });
            $('#nbEight').click(function (event) {
                getOperand();
            });
            $('#nbNine').click(function (event) {
                getOperand();
            });
            $('#nbZero').click(function (event) {
                getOperand();
            });
            $('#nbDeci').click(function (event) {
                getOperand();
            });

            $('#opAdd').click(function (event) {
                getOperator("+");
            });
            $('#opSub').click(function (event) {
                getOperator("-");
            });
            $('#opPro').click(function (event) {
                getOperator("*");
            });
            $('#opDiv').click(function (event) {
                getOperator("/");
            });

            $('#opEquals').click(function (event) {
                calculate();
            });

            $('#clear').click(function (event) {
                clear();
            });

            $('#back').click(function (event) {
                back();
            });
        });
    };

    function getOperand(callback) {
        calComp = false;
        if (operatorsArray.length == 0 && currentData == '') {
            currentData = Number(textBox.innerText);
            bindingId = 'firstOperand';
            bindData(bindingId);
            getIndex(callback);
        } else {
            currentData = Number(textBox.innerText);
            autoNaviAndSet(currentData);
            if (callback != undefined) {
                callback();
            }
        }
    }

    function getOperator(op) {
        if (calComp) {
            getOperand(function () {
                getOperator(op);
            });
        } else {
            currentPosArray[posIndex + 1] = getNextRow(currentPosArray[posIndex]);
            posIndex++;
            Office.context.document.goToByIdAsync(currentPosArray[posIndex], Office.GoToType.NamedItem, function (asyncResult) {
                if (asyncResult.status == 'failed') {
                    app.showNotification('Cannot navigate back to the row under last operand, please click C to restart. Detailed error: ' + asyncResult.error.message);
                }
            });

            switch (op) {
                case "+":
                    operatorsArray[operatorsIndex] = "+";
                    break;
                case "-":
                    operatorsArray[operatorsIndex] = "-";
                    break;
                case "*":
                    operatorsArray[operatorsIndex] = "*";
                    break;
                case "/":
                    operatorsArray[operatorsIndex] = "/";
                    break;
            }

            operatorsIndex++;
            currentData = '';
            textBox.innerText = '';
        }
    }

    function calculate() {
        if (textBox.innerText == '') {
            textBox.innerText = '0';
            getOperand(calculate);
        } else {
            currentPosArray[posIndex + 1] = getNextRow(currentPosArray[posIndex]);
            posIndex++;
            Office.context.document.goToByIdAsync(currentPosArray[posIndex], Office.GoToType.NamedItem, function (asyncResult) {
                if (asyncResult.status == "failed") {
                    app.showNotification('Cannot navigate back to the row under last operand, please click C to restart. Detailed error: ' + asyncResult.error.message);
                } else {
                    var resultFormular = getResultFormuar(posIndex, operatorsArray);

                    Office.context.document.setSelectedDataAsync(resultFormular, function (asyncResult) {                        
                        if (asyncResult.status == "failed") {
                            app.showNotification("Formula wasn't able to write in the cell one row under the same column of last operand. Please make sure the cell is not written protected and click C to restart. Detailed error: " + asyncResult.error.message);
                        } else {
                            currentPosArray = [];
                            posIndex = 0;
                            operatorsArray = [];
                            operandsIndex = 0;
                            operatorsIndex = 0;
                            currentData = '';
                            Office.context.document.getSelectedDataAsync('text', function (asyncResult) {
                                if (asyncResult.status == "failed") {
                                    app.showNotification("Didn't get the calculation result from current cell. The data shown in the task pane may be not accurate calculation result. Detailed error: " + asyncResult.error.message);
                                } else {
                                    // round up the value to two decimal places to show in the task pane if the result is decimal.
                                    var temp = String(asyncResult.value);
                                    if (temp.length > 10 && temp.indexOf('.') > 0) {
                                        textBox.innerText = Math.round(asyncResult.value * Math.pow(10, 2)) / Math.pow(10, 2);
                                    } else {
                                        textBox.innerText = asyncResult.value;
                                    }
                                }
                            });
                            calComp = true;
                            Office.context.document.bindings.releaseByIdAsync(bindingId);
                        }
                    });
                }
            });
        }
    }

    function clear() {
        currentPosArray = [];
        posIndex = 0;
        operatorsArray = [];
        operandsIndex = 0;
        operatorsIndex = 0;
        currentData = '';
        textBox.innerText = '0';
        Office.context.document.setSelectedDataAsync('', function (asynResult) {
            Office.context.document.bindings.releaseByIdAsync(bindingId); 
        });               
    }

    function back() {
        if (!calComp) {
            if (textBox.innerText != '' || textBox.innerText != '0') {
                textBox.innerText = textBox.innerText.substring(0, textBox.innerText.length - 1);
            }

            if (textBox.innerText == '') {
                textBox.innerText = '0';
            }

            autoNaviAndSet(textBox.innerText);
        }
    }

    // Unitities 
    function bindData(bindingId) {
        if (bindingId == undefined) {
            bindingId = 'mybinding';
        }
        Office.context.document.bindings.addFromSelectionAsync("text", { id: bindingId }, function (asyncResult) {
            if (asyncResult.status == "failed") {
                app.showNotification('Error when trying to bind to the first operand. Please click C to restart. Detialed error: ' + asyncResult.error.message);
            }
        });
    }

    function getIndex(callback) {
        // input =CELL("address") formular to current selection
        Office.select("bindings#firstOperand").setDataAsync('=CELL("address")', function (asyncResult) {
            if (asyncResult.status == "failed") {
                app.showNotification("Error when trying to input =CELL(\"address\") into the first operator cell. Please make sure the cell is not written protected and click C to restart. Detialed error: " + asyncResult.error.message);
            } else {
                //get the coordinates
                Office.select("bindings#firstOperand").getDataAsync(function (asyncResult) {
                    if (asyncResult.status == "failed") {
                        app.showNotification("Error when trying to get the cell's coordiantes using formula = CELL(\"address\"). Please click C to restart. Detailed error: " + asyncResult.error.message);
                    } else {
                        currentPosArray[posIndex] = asyncResult.value;
                        //put data back
                        Office.select("bindings#firstOperand").setDataAsync(currentData, function (asyncResult) {
                            if (asyncResult.status == "failed") {
                                app.showNotification("Error when trying to put the first operand back to current cell. Please make sure the cell is not written protected and click C to restart. Detialed error: " + asyncResult.error.message);
                            } else {
                                if (callback != undefined) {
                                    callback();
                                }
                            }
                        });
                    }
                });
            }
        });
    }

    function autoNaviAndSet(data) {
        Office.context.document.goToByIdAsync(currentPosArray[posIndex], Office.GoToType.NamedItem, function (asyncResult) {
            if (asyncResult.status == "failed") {
                app.showNotification('Error when trying auto navigate back to one row under the same column of the previous operation cell. Please click C to restart. Detailed error: ' + asyncResult.error.message);
            } else {
                Office.context.document.setSelectedDataAsync(data, function (asyncResult) {
                    if (asyncResult.status == "failed") {
                        app.showNotification('Error when trying to set the user input or current data to the current active cell. Please click C to restart. Detailed error: ' + asyncResult.error.message);
                    }
                });
            }
        });
    }

    function getNextRow(initialPos) {
        var tempArray = initialPos.split("$");
        tempArray[2] = Number(tempArray[2]) + 1;
        var newPos = '$' + tempArray[1] + '$' + tempArray[2];
        return newPos;
    }

    function getResultFormuar(i, ops) {
        var resultFormular = '=';
        var sumEnabled = true;
        var productEnabled = true;

        for (var k = 0; k <= ops.length; k++) {
            if (productEnabled && ops[k] == '*') {
                var temp = k;
                for (k = k + 1; k < ops.length; k++) {
                    if (ops[k] != '*') {
                        break;
                    }
                }
                resultFormular += 'PRODUCT(' + reformatPos(currentPosArray[temp]) + ':' + reformatPos(currentPosArray[k]) + ')';
            } else if (sumEnabled && (ops[k] == '+')) {
                var temp = k;
                for (k = k + 1; k < ops.length; k++) {
                    if (ops[k] != '+') {
                        break;
                    }
                }

                if ((ops[k] == '*') || (ops[k] == '/')) {
                    k--;
                }

                if (k > temp) {

                    resultFormular += 'SUM(' + reformatPos(currentPosArray[temp]) + ':' + reformatPos(currentPosArray[k]) + ')';
                } else {
                    resultFormular += reformatPos(currentPosArray[k]);
                }
            } else {
                resultFormular += reformatPos(currentPosArray[k]);
            }

            if (k < ops.length) {
                resultFormular += ops[k];
                sumEnabled = ops[k] == '+';
                productEnabled = ops[k] != '/';
            } else {
                break;
            }
        }
        return resultFormular;
    }

    function reformatPos(position) {
        var tempArray = position.split('$');
        return tempArray[1] + tempArray[2];
    }
})();