const { app, BrowserWindow, Menu } = require("electron");

const url = require("url");
const path = require("path");

let mainWindow;

// Reload in Development for Browser Windows
if (process.env.NODE_ENV !== "production") {
  require("electron-reload")(__dirname, {
    electron: path.join(__dirname, "../node_modules", ".bin", "electron")
  });
}

app.on("ready", () => {
  // The Main Window
  mainWindow = new BrowserWindow({ width: 720, height: 600 });

  mainWindow.loadURL(
    url.format({
      pathname: path.join(__dirname, "views/index.html"),
      protocol: "file",
      slashes: true
    })
  );

  // Menu
  const mainMenu = Menu.buildFromTemplate(templateMenu);
  // Set The Menu to the Main Window
  Menu.setApplicationMenu(mainMenu);

  // If we close main Window the App quit
  mainWindow.on("closed", () => {
    app.quit();
  });
});

// Menu Template
const templateMenu = [
  {
    label: "File",
    submenu: [
      {
        label: "Load file",
        accelerator: process.platform == "darwin" ? "command+L" : "Ctrl+L",
        click() {
          loadFile();
        }
      },
      {
        label: "Exit",
        accelerator: process.platform == "darwin" ? "command+Q" : "Ctrl+Q",
        click() {
          app.quit();
        }
      }
    ]
  }
];

// if you are in Mac, just add the Name of the App
if (process.platform === "darwin") {
  templateMenu.unshift({
    label: app.getName()
  });
}

// Developer Tools in Development Environment
if (process.env.NODE_ENV !== "production") {
  templateMenu.push({
    label: "DevTools",
    submenu: [
      {
        label: "Show/Hide Dev Tools",
        accelerator: process.platform == "darwin" ? "Comand+D" : "Ctrl+D",
        click(item, focusedWindow) {
          focusedWindow.toggleDevTools();
        }
      },
      {
        role: "reload"
      }
    ]
  });
}
