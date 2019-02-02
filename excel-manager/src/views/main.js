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
  var first_sheet_name = workbook.SheetNames[0];
  /* Get worksheet */
  var worksheet = workbook.Sheets[first_sheet_name];

  var address_of_cell = "A2";

  /* Find desired cell */
  var desired_cell = worksheet[address_of_cell];

  /* Get the value */
  var desired_value = desired_cell ? desired_cell.v : undefined;

  /* Create an array of json that will contains the proccesed data */
  var ws_data = [
    {
      A: "Día",
      B: "Debe",
      C: "Haber",
      D: "M",
      E: "Importe",
      F: "Total",
      G: "I",
      H: "I.V.A.",
      I: "Leyenda",
      J: "L",
      K: "Cotización",
      L: "Centro",
      M: "Fecha Vto."
    }
  ];

  /** Test sheet to json */

  var jsonSheet = XLSX.utils.sheet_to_json(worksheet);

  const rowNumber = jsonSheet.length;

  /* start the iteration */
  for (let r = 0; r < rowNumber; r++) {
    /** created a json to store the new data */
    let json = {};
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
    /** processed percentage */

    let porcentaje = ((r + 1) / rowNumber) * 100 + "%";
    console.log(porcentaje);
    document.getElementById("progress-bar-id").innerHTML = porcentaje;
    document.getElementById("progress-bar-id").style.width = porcentaje;
  }

  document.getElementById("progress-bar-id").classList.add("bg-success");

  /*
  for(var R = range.s.r; R <= range.e.r; ++R) {
    for(var C = range.s.c; C <= range.e.c; ++C) {
      var cell_address = {c:C, r:R};     
      var cell_ref = XLSX.utils.encode_cell(cell_address);
    }
  }

  var ws = XLSX.utils.json_to_sheet([
    { A: "S", B: "h", C: "e", D: "e", E: "t", F: "J", G: "S" }
  ], {header: ["A", "B", "C", "D", "E", "F", "G"], skipHeader: true});


  XLSX.utils.sheet_add_json(ws, [
    { A: 1, B: 2 }, { A: 2, B: 3 }, { A: 3, B: 4 }
  ], {skipHeader: true, origin: "A2"});


  XLSX.utils.sheet_add_json(ws, [
    { A: 5, B: 6, C: 7 }, { A: 6, B: 7, C: 8 }, { A: 7, B: 8, C: 9 }
  ], {skipHeader: true, origin: { r: 1, c: 4 }, header: [ "A", "B", "C" ]});


  XLSX.utils.sheet_add_json(ws, [
    { A: 4, B: 5, C: 6, D: 7, E: 8, F: 9, G: 0 }
  ], {header: ["A", "B", "C", "D", "E", "F", "G"], skipHeader: true, origin: -1});

  */

  /* end of the iteration */

  /* 
  var ws = XLSX.utils.aoa_to_sheet(ws_data);

  
  var new_ws_name = first_sheet_name;

  
  XLSX.utils.book_append_sheet(wb, ws, new_ws_name);
  */

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
