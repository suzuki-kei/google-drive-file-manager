//
//
// Google Drive を対象にドキュメントインデックスを生成する Google Apps Script.
//
// Google Spreadsheet にバンドルして使用する.
//
//

// ディレクトリの深さのデフォルト.
var DEFAULT_MAX_DEPTH = 5

// パスの区切り文字のデフォルト.
var DEFAULT_PATH_SEPARATOR = " > "

var Files = {
  isFile: function(file) {
    return !file.addFile
  },
  isFolder: function(file) {
    return !!file.addFile
  },
  getRootFolder: function() {
    return DriveApp.getRootFolder()
  },
  getCurrentFolder: function() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    var spreadsheetFile = DriveApp.getFileById(spreadsheet.getId())
    return spreadsheetFile.getParents().next()
  },
  join: function(paths, pathSeparator) {
    var value = ""
    var separator = ""
    for (var i = 0; i < paths.length; i++) {
      value += separator + paths[i].getName()
      separator = pathSeparator
    }
    return value
  },
}

var Cells = {
  setNumber: function(range, value) {
    range.setValue(value)
    range.setNumberFormat("0")
  },
  setText: function(range, value) {
    range.setValue(value)
    range.setNumberFormat("@")
  },
  setTextLink: function(range, url, value) {
    value = value.replace(/"/g, '""')
    range.setValue('=HYPERLINK("' + url + '", "' + value + '")')
    range.setNumberFormat("@")
    range.setShowHyperlink(true)
  },
}

function onScheduleTriggered() {
  var sheet = SpreadsheetApp.getActiveSheet()
  var rootFolder = DriveApp.getFolderById(Files.getCurrentFolder().getId())
  var maxDepth = DEFAULT_MAX_DEPTH
  var pathSeparator = DEFAULT_PATH_SEPARATOR
  var includeFiles = true
  var includeFolders = true
  generateDocumentIndex(
    sheet,
    rootFolder,
    maxDepth,
    pathSeparator,
    includeFiles,
    includeFolders)
}

function onOpen() {
  setupUi(SpreadsheetApp.getUi())
}

function setupUi(ui) {
  var menu = ui.createMenu("Document Index")
  menu.addItem("Generate Document Index...", "openDocumentIndexOptionsDialog")
  menu.addToUi()
}

function openDocumentIndexOptionsDialog() {
  var template = HtmlService.createTemplateFromFile("document-index-options.template.html")
  template.includeFiles = "checked"
  template.includeFolders = "checked"
  template.maxDepth = DEFAULT_MAX_DEPTH
  template.pathSeparator = DEFAULT_PATH_SEPARATOR
  template.driveRootFolderId = Files.getRootFolder().getId()
  template.currentFolderId = Files.getCurrentFolder().getId()

  var ui = SpreadsheetApp.getUi()
  ui.showModelessDialog(template.evaluate(), "Generate Document Index")
}

function onGenerateButtonClicked(options) {
  generateDocumentIndex(
    SpreadsheetApp.getActiveSheet(),
    DriveApp.getFolderById(options.rootFolderId),
    options.maxDepth,
    options.pathSeparator,
    options.includeFiles,
    options.includeFolders)
}

function generateDocumentIndex(sheet, rootFolder, maxDepth, pathSeparator, includeFiles, includeFolders) {
  var files = getFiles(rootFolder, maxDepth, includeFiles, includeFolders)
  updateSheet(sheet, files, pathSeparator)
}

function getFiles(rootFolder, maxDepth, includeFiles, includeFolders) {
  var files = []
  traverseDrive(rootFolder, maxDepth, function(parents, file) {
    if (Files.isFile(file) && !includeFiles) {
      return
    }
    if (Files.isFolder(file) && !includeFolders) {
      return
    }
    files.push({
      file: file,
      routes: parents.concat(file),
      parents: parents,
    })
  })

  files.sort(function(lhs, rhs) {
    var lhsSortKey = Files.join(lhs.routes)
    var rhsSortKey = Files.join(rhs.routes)
    return lhsSortKey.localeCompare(rhsSortKey)
  })
  return files
}

function traverseDrive(rootFolder, maxDepth, callback) {
  function traverse(parents, folder, depth, maxDepth, callback) {
    if (depth > maxDepth) {
      return
    }
    var query = "'" + folder.getId() + "' in parents"
    var subFolders = DriveApp.searchFolders(query)
    while (subFolders.hasNext()) {
      var subFolder = subFolders.next()
      callback(parents.concat(folder), subFolder)
      traverse(parents.concat(folder), subFolder, depth + 1, maxDepth, callback)
    }
    var files = DriveApp.searchFiles(query)
    while (files.hasNext()) {
      var file = files.next()
      callback(parents.concat(folder), file)
    }
  }

  callback([], rootFolder)
  traverse([], rootFolder, 1, maxDepth, callback)
}

function updateSheet(sheet, filePaths, pathSeparator) {
  initializeSheet(sheet)
  generateHeaderRow(sheet, filePaths)
  generateValueRows(sheet, filePaths, pathSeparator)
}

function initializeSheet(sheet) {
  sheet.clear()
}

function generateHeaderRow(sheet, filePaths) {
  var headers = ["No.", "Type", "MIME Type", "Full Path"]

  var maxDepth = 0
  for (var i = 0; i < filePaths.length; i++) {
    maxDepth = Math.max(maxDepth, filePaths[i].parents.length + 1)
  }
  for (var i = 0; i < maxDepth; i++) {
    headers.push("Path (Level " + i + ")")
  }

  var row = 1
  var range = sheet.getRange(row, 1, 1, headers.length)
  range.setValues([headers])
  range.setBackground("orange")
  range.setHorizontalAlignment("center")
}

function generateValueRows(sheet, filePaths, pathSeparator) {
  var row = 2
  for (var i = 0; i < filePaths.length; i++) {
    generateValueRow(sheet, filePaths[i], pathSeparator, i, row++)
  }
}

function generateValueRow(sheet, filePath, pathSeparator, index, row) {
  function setNoCell(column) {
    var range = sheet.getRange(row, column)
    Cells.setNumber(range, index + 1)
  }
  function setTypeCell(column) {
    var range = sheet.getRange(row, column)
    var value = Files.isFile(filePath.file) ? "File" : "Directory"
    Cells.setText(range, value)
  }
  function setMimeTypeCell(column) {
    var range = sheet.getRange(row, column)
    if (Files.isFile(filePath.file)) {
      Cells.setText(range, filePath.file.getMimeType())
    }
  }
  function setFullPathCell(column) {
    var value = Files.join(filePath.routes, pathSeparator)
    var url = filePath.file.getUrl()
    var range = sheet.getRange(row, column)
    Cells.setTextLink(range, url, value)
  }
  function setPathCells(column) {
    for (var i = 0; i < filePath.routes.length; i++) {
      var range = sheet.getRange(row, column++)
      var route = filePath.routes[i]
      Cells.setTextLink(range, route.getUrl(), route.getName())
    }
  }

  var column = 1
  setNoCell(column++)
  setTypeCell(column++)
  setMimeTypeCell(column++)
  setFullPathCell(column++)
  setPathCells(column++)
}
