const { dialog } = require("electron").remote;
var fs = require("fs");

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
  fs.readFile(filepath, "utf-8", (err, data) => {
    if (err) {
      alert("An error ocurred reading the file :" + err.message);
      return;
    }
    console.log("The file content is : " + data);
    document.getElementById("progress-bar").style.display = "inline";
  });
}

function selectFile() {
  console.log("select files was clicked");
  dialog.showOpenDialog(fileNames => {
    // fileNames is an array that contains all the selected
    if (fileNames === undefined) {
      console.log("No file selected");
      return;
    }

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

function saveFile() {
  let content = "Some text to save into the file";

  dialog.showSaveDialog(fileName => {
    if (fileName === undefined) {
      console.log("You didn't save the file");
      return;
    }

    fs.writeFile(fileName + ".txt", content, err => {
      if (err) {
        notification(
          "error",
          "An error ocurred creating the file " + err.message
        );
        return;
      }

      notification("success", "The file was saved successfully");
    });
  });

  /*
  dialog.showOpenDialog(
    {
      title: "Select a folder",
      properties: ["openDirectory"]
    },
    folderPaths => {
      // folderPaths is an array that contains all the selected paths
      if (fileNames === undefined) {
        console.log("No destination folder selected");
        return;
      } else {
        console.log(folderPaths);
        fs.writeFile(fileName, content, err => {
          if (err) {
            notification(
              error,
              "An error ocurred creating the file " + err.message
            );
            return;
          }
          notification(success, "The file was saved successfully");
        });
      }
    }
  );*/
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
