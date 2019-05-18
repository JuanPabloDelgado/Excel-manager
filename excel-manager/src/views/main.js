const { dialog } = require("electron").remote;
const path = require("path");

if (typeof require !== "undefined") XLSX = require("xlsx");
var originalFilePath = [];
var wb;
var cantidadFilasProcesadas = 0;
var cantidadTotalFilas;
var workbooksProcessed = [];

var holder = document.getElementById("drag-file");

holder.ondragover = () => {
  return false;
};

holder.ondragleave = () => {
  return false;
};

holder.ondragend = () => {
  return false;
};

holder.ondrop = e => {
  e.preventDefault();

  document.getElementById("drag-file").innerHTML = "File(s) you dragged here: ";

  /* show the progress bar when the files was loaded */
  document.getElementById("progress-bar").style.display = "inline";

  /* This function calculates the total rows in all files that was loaded */
  /* The purpose of this is show correctly the load of the progress bar */
  totalRows(e.dataTransfer.files, "droped");

  for (let f of e.dataTransfer.files) {
    originalFilePath.push(f.path);
    var name = f.path.substring(f.path.lastIndexOf("\\") + 1, f.path.length);

    var node = document.createElement("LI"); // Create a <li> node
    var textnode = document.createTextNode(name); // Create a text node
    node.appendChild(textnode); // Append the text to <li>
    document.getElementById("drag-file").appendChild(node);

    console.log("On drag - processFile: ", f.path);
    processFile(f.path);
  }

  /* Once the file is ready, the button is enable to the user to save the new file*/
  document.getElementById("saveFileButton").disabled = false;

  return false;
};

function selectFile() {
  dialog.showOpenDialog(
    {
      properties: ["openFile", "multiSelections"]
    },
    fileNames => {
      if (fileNames === undefined) {
        console.log("No file selected");
        return;
      }

      document.getElementById("drag-file").innerHTML = "File(s) you selected: ";
      /* show the progress bar when the files was loaded */
      document.getElementById("progress-bar").style.display = "inline";

      /* This function calculates the total rows in all files that was loaded */
      /* The purpose of this is show correctly the load of the progress bar */

      totalRows(fileNames, "selected");

      fileNames.forEach((file, index) => {
        originalFilePath.push(file);

        var name = file.substring(file.lastIndexOf("\\") + 1, file.length);

        var node = document.createElement("LI"); // Create a <li> node
        var textnode = document.createTextNode(name); // Create a text node
        node.appendChild(textnode); // Append the text to <li>
        document.getElementById("drag-file").appendChild(node);
        console.log("On select - processFile: ", file);
        processFile(file);
      });
      /* Once the file is ready, the button is enable to the user to save the new file*/
      document.getElementById("saveFileButton").disabled = false;
    }
  );
}

function processFile(fileURL) {
  let porcentaje = 0;
  wb = XLSX.utils.book_new();

  workbookToProcess = XLSX.readFile(fileURL);
  workbookToProcess.SheetNames.forEach((sheet, index) => {
    var worksheet = workbookToProcess.Sheets[sheet];

    /* Create an array of json that will contains the proccesed data */
    var ws_data = [];

    /** Test sheet to json */
    var jsonSheet = XLSX.utils.sheet_to_json(worksheet);

    const rowNumber = jsonSheet.length;

    /* start the iteration */
    for (let r = 0; r < rowNumber; r++) {
      /** getting the date */
      let dateCell = "A" + (r + 2);
      let date = worksheet[dateCell];
      let day = date["w"];
      /** getting the "Col1" */
      /** getting the "Col2" */
      /** getting the "Col3" */
      /** getting the "Col4" */
      /** getting the "Col5" */
      /** getting the "Col6" */
      /** getting the "Col7" */
      /** getting the "Col8" */
      /** getting the "Col9" */
      /** getting the "Col10" */
      /** ------------------------------------------------- */

      /** for each row in the input file, we create a new json who contain the new processed information */

      let newRow = {
        Date: day,
        Col1: "NewCol1Val: " + r,
        Col2: "NewCol2Val" + r,
        Col3: "NewCol3Val" + r,
        Col4: "NewCol4Val" + r,
        Col5: "NewCol5Val" + r,
        Col6: "NewCol6Val" + r,
        Col7: "NewCol7Val" + r,
        Col8: "NewCol8Val" + r,
        Col9: "NewCol9Val" + r,
        Col10: "NewCol10Val" + r
      };

      /** update the worksheet */
      ws_data.push(newRow);

      cantidadFilasProcesadas += 1;
      /** processed percentage */
      porcentaje = (cantidadFilasProcesadas / cantidadTotalFilas) * 100 + "%";
      /** update the progress bar */
      document.getElementById("progress-bar-id").innerHTML = porcentaje;
      document.getElementById("progress-bar-id").style.width = porcentaje;
    }

    /** transform the array of json into a excel whorksheet */
    var newJsonSheet = XLSX.utils.json_to_sheet(ws_data, {
      header: [
        "Date",
        "Col1",
        "Col2",
        "Col3",
        "Col4",
        "Col5",
        "Col6",
        "Col7",
        "Col8",
        "Col9",
        "Col10"
      ],
      skipHeader: false
    });

    /** append the new worksheet to a workbook we create previusly */
    XLSX.utils.book_append_sheet(wb, newJsonSheet, sheet);
  });

  /** update the color of progress bar to green. This means that all rows of the input file was processed */
  document.getElementById("progress-bar-id").classList.add("bg-success");

  workbooksProcessed.push(wb);
}

function saveFile() {
  dialog.showOpenDialog(
    {
      title: "Select a folder",
      properties: ["openDirectory"]
    },
    folderPaths => {
      // folderPaths is an array that contains all the selected paths
      if (folderPaths === undefined) {
        console.log("No destination folder selected");
        return;
      } else {
        workbooksProcessed.forEach((workbook, index) => {
          console.log("originalFilePath: ", originalFilePath);
          console.log("originalFilePath index: ", index);

          var name = originalFilePath[index].substring(
            originalFilePath[index].lastIndexOf("\\") + 1,
            originalFilePath[index].length - 5
          );

          XLSX.writeFile(
            workbook,
            folderPaths[0] + "\\" + name + "-modificado" + ".xlsx"
          );
        });

        notification(
          "success",
          "Su(s) archivo(s) modificado(s) fue guardado(s) correctamente"
        );

        document.getElementById("drag-file").innerHTML = `
          <img src="../../assets/icons/win/test5.png" alt="" />
          <span class="text-center">
            <a href="#" onclick="selectFile()">Elija un archivo</a> o
            arrastrelo aquí.
          </span>
          `;
        document.getElementById("progress-bar").style.display = "none";
      }
    }
  );
}

function notification(type, message) {
  const notifier = require("node-notifier");

  var onError = function(err, response) {
    console.error(err, response);
  };

  if (type == "error") {
    notifier.notify(
      {
        message: message,
        title: "Error",
        sound: true,
        icon: path.join(__dirname, "../../assets/icons/win/error.png"),
        wait: true,
        timeout: 5,
        closeLabel: void 0
      },
      onError
    );
  } else {
    notifier.notify(
      {
        message: message,
        title: "Success",
        sound: true,
        icon: path.join(__dirname, "../../assets/icons/win/success.png"),
        wait: true,
        timeout: 5,
        closeLabel: void 0
      },
      onError
    );
  }
}

function totalRows(files, type) {
  var totalRows = 0;
  if (type === "droped") {
    for (let f of files) {
      workbookTotalRows = XLSX.readFile(f.path);
      workbookTotalRows.SheetNames.forEach((sheet, index) => {
        totalRows += XLSX.utils.sheet_to_json(workbookTotalRows.Sheets[sheet])
          .length;
      });
    }
  } else if (type === "selected") {
    files.forEach((file, index) => {
      workbookTotalRows = XLSX.readFile(file);
      workbookTotalRows.SheetNames.forEach((sheet, index) => {
        totalRows += XLSX.utils.sheet_to_json(workbookTotalRows.Sheets[sheet])
          .length;
      });
    });
  } else {
    return;
  }

  cantidadTotalFilas = totalRows;
}
