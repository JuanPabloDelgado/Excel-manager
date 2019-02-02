const { dialog } = require("electron").remote;
const path = require("path");

if (typeof require !== "undefined") XLSX = require("xlsx");
var workbook;
var originalFilePath;
var wb = XLSX.utils.book_new();

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

  for (let f of e.dataTransfer.files) {
    var name = f.path.substring(f.path.lastIndexOf("\\") + 1, f.path.length);

    document.getElementById("drag-file").innerHTML =
      "File(s) you dragged here: " + name;
    readFile(f.path);
  }

  return false;
};

function readFile(filepath) {
  workbook = XLSX.readFile(filepath);
  processFile(workbook);

  document.getElementById("progress-bar").style.display = "inline";
}

function selectFile() {
  dialog.showOpenDialog(fileNames => {
    // fileNames is an array that contains all the selected
    if (fileNames === undefined) {
      console.log("No file selected");
      return;
    }
    originalFilePath = fileNames[0];

    var name = fileNames[0].substring(
      fileNames[0].lastIndexOf("\\") + 1,
      fileNames[0].length
    );

    document.getElementById("drag-file").innerHTML =
      "File(s) you selected: " + name;

    readFile(fileNames[0]);
  });

  /*
  dialog.showOpenDialog({
    properties: [
      "openFile",
      "multiSelections",
      fileNames => {
        console.log(fileNames);
      }
    ]
  });*/
}

function processFile(file) {
  var first_sheet_name = file.SheetNames[0];
  /* Get the input file first worksheet */
  var worksheet = file.Sheets[first_sheet_name];

  /* Create an array of json that will contains the proccesed data */
  var ws_data = [];

  /** Test sheet to json */

  var jsonSheet = XLSX.utils.sheet_to_json(worksheet);

  const rowNumber = jsonSheet.length;

  /* start the iteration */
  for (let r = 0; r < rowNumber; r++) {
    /** getting the date */
    let celdaDia = "A" + (r + 2);
    let fecha = worksheet[celdaDia];
    let dia = fecha["w"].split("/")[1];
    /** getting the "Debe" */
    /** getting the "Haber" */
    /** getting the " Moneda" */
    /** getting the "Importe Total" */
    /** getting the "Tipo impuesto" */
    /** getting the "Iva" */
    /** getting the "Leyenda" */
    /** getting the "L" */
    /** getting the "Cotización" */
    /** getting the "Centro" */
    /** getting the "Fecha Vto." */
    /** ------------------------------------------------- */

    /** for each row in the input file, we create a new json who contain the new processed information */
    let newRow = {
      Dia: dia,
      Debe: "Debe row: " + r,
      Haber: "Haber row:" + r,
      M: "M row: " + r,
      Importe: "Importe row: " + r,
      Total: "Total row: " + r,
      I: "I row: " + r,
      "I.V.A.": "I.V.A. row: " + r,
      Leyenda: "Leyenda row: " + r,
      L: "L row: " + r,
      Cotizacion: "Cotizacion row: " + r,
      Centro: "Centro row: " + r,
      "Fecha Vto.": "Fecha Vto. row: " + r
    };

    /** update the worksheet */
    ws_data.push(newRow);

    /** processed percentage */
    let porcentaje = ((r + 1) / rowNumber) * 100 + "%";
    /** update the progress bar */
    document.getElementById("progress-bar-id").innerHTML = porcentaje;
    document.getElementById("progress-bar-id").style.width = porcentaje;
  }

  /** update the color of progress bar to green. This means that all rows of the input file was processed */
  document.getElementById("progress-bar-id").classList.add("bg-success");

  /** transform the array of json into a excel whorksheet */
  var newJsonSheet = XLSX.utils.json_to_sheet(ws_data, {
    header: [
      "Dia",
      "Debe",
      "Haber",
      "M",
      "Importe",
      "Total",
      "I",
      "I.V.A.",
      "Leyenda",
      "L",
      "Cotizacion",
      "Centro",
      "Fecha Vto."
    ],
    skipHeader: false
  });

  /** append the new worksheet to a workbook we create previusly */
  XLSX.utils.book_append_sheet(wb, newJsonSheet, first_sheet_name);

  /* Once the file is ready, the button is enable to the user to save the new file*/
  document.getElementById("saveFileButton").disabled = false;
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
        var name = originalFilePath.substring(
          originalFilePath.lastIndexOf("\\") + 1,
          originalFilePath.length - 5
        );

        XLSX.writeFile(
          wb,
          folderPaths[0] + "\\" + name + "-modificado" + ".xlsx"
        );

        notification(
          "success",
          "Su archivo modificado fue guardado correctamente"
        );
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
        icon: path.join(__dirname, "../../assets/icons/win/success12.png"),
        wait: true,
        timeout: 5,
        closeLabel: void 0
      },
      onError
    );
  }
}
