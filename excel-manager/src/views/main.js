const { dialog } = require("electron").remote;
var fs = require("fs");

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
    console.log("File(s) you dragged here: ", f.path);
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
  var address_of_cell = "A1";

  /* Get worksheet */
  var worksheet = workbook.Sheets[first_sheet_name];

  /* Find desired cell */
  var desired_cell = worksheet[address_of_cell];

  /* Get the value */
  var desired_value = desired_cell ? desired_cell.v : undefined;

  /*
  Add a new worksheet to the new bookbook
  */
  var new_ws_name = first_sheet_name;

  /* make worksheet */

  var ws_data = [["Hello world"]];
  var ws = XLSX.utils.aoa_to_sheet(ws_data);

  /* Add the worksheet to the workbook */
  XLSX.utils.book_append_sheet(wb, ws, new_ws_name);

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
        icon: "../../assets/icons/win/cancel.png",
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
        icon: "../../assets/icons/win/ok.png",
        wait: true,
        timeout: 5,
        closeLabel: void 0
      },
      onError
    );
  }
}
