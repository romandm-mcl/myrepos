/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";
    document.getElementById("druk-etikety").onclick = drukEtikety;
    document.getElementById("druk-x").onclick = drukEtikety;
  }
});

// $("#druk-etikety").click(() => tryCatch(drukEtikety));
// $("#btn-test").click(() => tryCatch(main));

async function drukEtikety() {
  Excel.run(async (context) => {
    const inputSheet = context.workbook.worksheets.getActiveWorksheet();
    const printSheet = context.workbook.worksheets.getItem("etykieta sklepu");
    let selectedValue = context.workbook.getSelectedRange();
    selectedValue.load("rowIndex,columnIndex,rowCount,columnCount,address");

    console.clear();
    await context.sync();
    let y = selectedValue.rowIndex;
    let x = selectedValue.columnIndex;
    let ilosc = selectedValue.rowCount;


    if (selectedValue.columnCount > 2) {
      document.getElementById("test-p").innerText = " чтото пошло не так : selected columns > 2 ";
      return context.sync();
    }

    if (selectedValue.columnIndex % 2 === 0) {
      // если выбран четный столбец с номерами палет
      // получаем  номер последней заполненой строки в листе
      selectedValue = inputSheet.getUsedRange().getLastRow();
      selectedValue.load("rowIndex");
      await context.sync();
      ilosc = selectedValue.rowIndex;
      // получаем кол-во заполненных строк
      selectedValue = inputSheet.getRangeByIndexes(y, x + 1, ilosc, 1).getExtendedRange(Excel.KeyboardDirection.up);
      selectedValue.load("rowCount");
      await context.sync();
      ilosc = selectedValue.rowCount;
    } else {
      x--;
    }
    selectedValue = inputSheet.getRangeByIndexes(y, x, ilosc, 2);
    selectedValue.load("rowCount,rowIndex,values");

    await context.sync();
    const numerPalety = selectedValue.values[0][0];
    if (numerPalety === "" || numerPalety === undefined)
      throw new Error(" что-то пошло не так : nie wybrano numeru palety / не выбран номер палеты");
    let arrKod=selectedValue.values.filter(a=> a[1]!=='' && a[1]!==undefined );//&& a[1]!=='x');
    let indx=1;
    while(indx<arrKod.length&&arrKod[indx][0]===''){indx++ };
    ilosc = indx;
    
    //  собираем данные для печати
    selectedValue = inputSheet.getRangeByIndexes(2, x + 1, 1, 1);
    selectedValue.load("values");
    inputSheet.load("name");
    await context.sync();
    const fulNameSklep=selectedValue.values;
    const nazwaSklepa = fulNameSklep
        .toString()
        .toUpperCase()
        .replace("GALERIA ", "")
        .substring(0, 10) + ".";
    const namePage = inputSheet.name.toString().substring(0, 3);

    selectedValue = printSheet.getUsedRange();
    selectedValue.load("columnCount");
    await context.sync();
    const printColumn = selectedValue.columnCount;
    const currentday = new Date(Date.now()); //.format("[$-409]m/d/yy h:mm AM/PM;@");
    const options = { year: "numeric", month: "numeric", day: "numeric", hour: "numeric", minute: "numeric" };
    switch (inputSheet.name) {
      case "MODIVO":
        printSheet.getRange("1:14").rowHidden = true;
        printSheet.getRange("15:29").rowHidden = false;
        printSheet.getRange("C19").values = [["L1"]];
        printSheet.getRange("A26").values = [["Kartonow: " + ilosc + " szt."]];
        printSheet.getRange("A24").values = [["#" + numerPalety]];
        selectedValue = inputSheet.getRangeByIndexes(y, x + 2, 1, 1);
        selectedValue.load("values");
        await context.sync();
        if (selectedValue.values[0][0] === "" || selectedValue.values[0][0] === undefined) {
          for (let i = 0; i < ilosc; i++) {
            inputSheet.getRangeByIndexes(y + i, x + 2, 1, 1).values = [
              [currentday.toLocaleDateString("pl-PL", options)]
            ];
            inputSheet.getRangeByIndexes(y + i, x + 3, 1, 1).values = [["PRZEKAZANO"]];
          }
          // await context.sync();
        }
        break;
      case "REKL-PSPRZ (PR2)":
        printSheet.getRange("1:14").rowHidden = true;
        printSheet.getRange("15:29").rowHidden = false;
        printSheet.getRange("C19").values = [["PR2"]];
        printSheet.getRange("A26").values = [["Kartonow: " + ilosc + " szt."]];
        printSheet.getRange("A24").values = [["#" + numerPalety]];
        selectedValue = inputSheet.getRangeByIndexes(y, x + 2, 1, 1);
        selectedValue.load("values");
        await context.sync();
        if (selectedValue.values[0][0] === "" || selectedValue.values[0][0] === undefined) {
          for (let i = 0; i < ilosc; i++) {
            inputSheet.getRangeByIndexes(y + i, x + 2, 1, 1).values = [
              //[currentday.toLocaleDateString('pl-PL', options)]
              [currentday.toLocaleTimeString().replace(/([\d]+:[\d]{2})(:[\d]{2})(.*)/, "'$1")]
            ];
            inputSheet.getRangeByIndexes(y + i, x + 3, 1, 1).values = [["PR2"]];
          }
          // await context.sync();
        }
        break;
      case "CCC.EU":
        // printSheet.getRange("I:P").col
        // await context.sync();
      default:
        printSheet.getRange("1:14").rowHidden = false;
        printSheet.getRange("15:29").rowHidden = true;
        // await context.sync();
        printSheet.getRange("A1").values = [[nazwaSklepa]];
        printSheet.getRange("C14").values = [["Kartonow: " + ilosc + " szt."]];
        printSheet.getRange("A10").values = [[namePage]];
        printSheet.getRange("F10").values = [[numerPalety]];
        break;
    }
    printSheet.activate();
    return context.sync();
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
    alert(error);
    // document.getElementById("test-p").innerText = error.message ;
  }
}

// export async function run() {
//   try {
//     await Excel.run(async (context) => {
//       /**
//        * Insert your Excel code here
//        */
//       const range = context.workbook.getSelectedRange();

//       // Read the range address
//       range.load("address");

//       // Update the fill color
//       range.format.fill.color = "yellow";

//       await context.sync();
//       console.log(`The range address was ${range.address}.`);
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }
