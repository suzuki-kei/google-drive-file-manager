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
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
        const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId())
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
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    const sheet = spreadsheet.getSheetByName("Document Index")
    const rootFolder = DriveApp.getFolderById(Files.getCurrentFolder().getId())
    const maxDepth = DEFAULT_MAX_DEPTH
    const pathSeparator = DEFAULT_PATH_SEPARATOR
    const includeFiles = true
    const includeFolders = true
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
    const menu = ui.createMenu("Document Index")
    menu.addItem("Generate...", "openDocumentIndexOptionsDialog")
    menu.addToUi()
}

function openDocumentIndexOptionsDialog() {
    const template = HtmlService.createTemplateFromFile("document-index-options.template.html")
    template.driveRootFolderId = Files.getRootFolder().getId()
    template.currentFolderId = Files.getCurrentFolder().getId()
    template.maxDepth = DEFAULT_MAX_DEPTH
    template.outputSheetName = SpreadsheetApp.getActiveSheet().getName()
    template.pathSeparator = DEFAULT_PATH_SEPARATOR
    template.includeFiles = "checked"
    template.includeFolders = "checked"

    const htmlOutput = template.evaluate()
                               .setWidth(600)
                               .setHeight(300)

    const ui = SpreadsheetApp.getUi()
    ui.showModelessDialog(htmlOutput, "Document Index")
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
    const files = getFiles(rootFolder, maxDepth, includeFiles, includeFolders)
    updateSheet(sheet, files, pathSeparator)
}

function getFiles(rootFolder, maxDepth, includeFiles, includeFolders) {
    const files = []
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
        const lhsSortKey = Files.join(lhs.routes)
        const rhsSortKey = Files.join(rhs.routes)
        return lhsSortKey.localeCompare(rhsSortKey)
    })
    return files
}

function traverseDrive(rootFolder, maxDepth, callback) {
    function traverse(parents, folder, depth, maxDepth, callback) {
        if (depth > maxDepth) {
            return
        }
        const query = "'" + folder.getId() + "' in parents"
        const subFolders = DriveApp.searchFolders(query)
        while (subFolders.hasNext()) {
            const subFolder = subFolders.next()
            callback(parents.concat(folder), subFolder)
            traverse(parents.concat(folder), subFolder, depth + 1, maxDepth, callback)
        }
        const files = DriveApp.searchFiles(query)
        while (files.hasNext()) {
            const file = files.next()
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
    const headers = ["No.", "Type", "MIME Type", "File Path", "File Name"]
    const range = sheet.getRange(1, 1, 1, headers.length)
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
        const range = sheet.getRange(row, column)
        Cells.setNumber(range, index + 1)
    }
    function setTypeCell(column) {
        const range = sheet.getRange(row, column)
        const value = Files.isFile(filePath.file) ? "File" : "Directory"
        Cells.setText(range, value)
    }
    function setMimeTypeCell(column) {
        const range = sheet.getRange(row, column)
        if (Files.isFile(filePath.file)) {
            Cells.setText(range, filePath.file.getMimeType())
        }
    }
    function setFilePathCell(column) {
        const richText = SpreadsheetApp.newRichTextValue()
        richText.setText(Files.join(filePath.routes, pathSeparator))

        var startOffset = 0
        for (var i = 0; i < filePath.routes.length; i++) {
            const route = filePath.routes[i]
            const endOffset = startOffset + route.getName().length
            richText.setLinkUrl(startOffset, endOffset, route.getUrl())
            startOffset = endOffset + pathSeparator.length
        }
        const range = sheet.getRange(row, column)
        range.setRichTextValue(richText.build())
    }
    function setFileNameCell(column) {
        const range = sheet.getRange(row, column)
        Cells.setTextLink(range, filePath.file.getUrl(), filePath.file.getName())
    }

    var column = 1
    setNoCell(column++)
    setTypeCell(column++)
    setMimeTypeCell(column++)
    setFilePathCell(column++)
    setFileNameCell(column++)
}

