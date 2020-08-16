
/**
 *
 * ドキュメントインデックス.
 *
 */
class DocumentIndex {

    /**
     *
     * メニューを作成する.
     *
     * @return {Menu}
     *     メニュー.
     *
     */
    static createMenu() {
        const ui = SpreadsheetApp.getUi()
        const menu = ui.createMenu("File Manager")
        menu.addItem("Document Index...", "DocumentIndex.openDialog")
        return menu
    }

    /**
     *
     * ドキュメントインデックス生成ダイアログを開く.
     *
     */
    static openDialog() {
        const settings = this.getSettings()
        const templateFileName = "document-index.dialog.template.html"
        const template = HtmlService.createTemplateFromFile(templateFileName)
        template.rootFolderUrl = settings.rootFolderUrl
        template.maxDepth = settings.maxDepth
        template.outputSheetName = settings.outputSheetName
        template.pathSeparator = settings.pathSeparator
        template.includeFiles = settings.includeFiles
        template.includeFolders = settings.includeFolders

        const htmlOutput = template.evaluate().setWidth(600).setHeight(300)
        const ui = SpreadsheetApp.getUi()
        ui.showModelessDialog(htmlOutput, "Document Index")
    }

    /**
     *
     * 設定を取得する.
     *
     * @return {object}
     *     設定.
     *
     */
    static getSettings() {
        const settings = new DocumentIndexSettings()
        try {
            settings.load()
        } catch(exception) {
            // シートからの読み込みに失敗した場合でもデフォルト値を返すため例外は無視する.
        }
        return settings
    }

    /**
     *
     * ドキュメントインデックス生成ダイアログの "Generate" ボタンが押された時の処理.
     *
     * @param {object} options
     *     ドキュメントインデックス生成ダイアログで指定されたオプション.
     *
     */
    static onGenerateButtonClicked(options) {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
        const sheet = Sheets.getOrCreateSheetByName(spreadsheet, options.outputSheetName)
        const rootFolder = Paths.getFolderByUrl(options.rootFolderUrl)
        this.generate(
            sheet,
            rootFolder,
            options.maxDepth,
            options.pathSeparator,
            options.includeFiles,
            options.includeFolders)
        sheet.activate()
    }

    /**
     *
     * ドキュメントインデックスを生成する.
     *
     * @param {Sheet} sheet
     *     このシートに結果を出力する.
     *
     * @param {Folder} rootFolder
     *     このフォルダ以下を探索する.
     *
     * @param {number} maxDepth
     *     探索する最大の深さ.
     *     1 を指定すると rootFolder 直下が対象となる.
     *     2 を指定すると rootFolder 直下とサブフォルダが対象となる.
     *
     * @param {string} pathSeparator
     *     パスの区切り文字.
     *
     * @param {boolean} includeFiles
     *     結果にファイルを含める場合は true.
     *
     * @param {boolean} includeFolders
     *     結果にフォルダを含める場合は true.
     *
     */
    static generate(sheet, rootFolder, maxDepth, pathSeparator, includeFiles, includeFolders) {
        const files = this.getFiles(rootFolder, maxDepth, includeFiles, includeFolders)
        this.updateSheet(sheet, files, pathSeparator)
    }

    /**
     *
     * 指定したフォルダに含まれるファイルを取得する.
     *
     * @param {Folder} rootFolder
     *     このフォルダ以下を探索する.
     *
     * @param {number} maxDepth
     *     探索する最大の深さ.
     *     1 を指定すると rootFolder 直下が対象となる.
     *     2 を指定すると rootFolder 直下とサブフォルダが対象となる.
     *
     * @param {boolean} includeFiles
     *     結果にファイルを含める場合は true.
     *
     * @param {boolean} includeFolders
     *     結果にフォルダを含める場合は true.
     *
     * @return {Array.<FilePath>}
     *     FilePath の配列.
     *     各要素は FilePath.routes を文字列連結した昇順にソートされている.
     *
     */
    static getFiles(rootFolder, maxDepth, includeFiles, includeFolders) {
        const filePaths = []
        this.traverseFiles(rootFolder, maxDepth, (parents, file) => {
            if (Paths.isFile(file) && !includeFiles) {
                return
            }
            if (Paths.isFolder(file) && !includeFolders) {
                return
            }
            const routes = parents.concat(file)
            filePaths.push(new FilePath(file, parents, routes))
        })

        filePaths.sort((lhs, rhs) => {
            const lhsSortKey = Paths.join(lhs.routes)
            const rhsSortKey = Paths.join(rhs.routes)
            return lhsSortKey.localeCompare(rhsSortKey)
        })
        return filePaths
    }

    /**
     *
     * 指定したフォルダに含まれるファイルを探索する.
     *
     * @param {Folder} rootFolder
     *     このフォルダ以下を探索する.
     *
     * @param {number} maxDepth
     *     探索する最大の深さ.
     *     1 を指定すると rootFolder 直下が対象となる.
     *     2 を指定すると rootFolder 直下とサブフォルダが対象となる.
     *
     * @param {function} callback
     *     発見したファイル情報を受け取るコールバック関数.
     *     callback(parents, file) という形式で呼び出される.
     *     parents は rootFolder から file までのパス (file を含まない).
     *
     */
    static traverseFiles(rootFolder, maxDepth, callback) {
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

    /**
     *
     * ドキュメントインデックスのシートを更新する.
     *
     * @param {Sheet} sheet
     *     更新対象の Sheet.
     *
     * @param {Array.<FilePath>} filePaths
     *     シートに追加する FilePath の配列.
     *
     * @param {string} pathSeparator
     *     パスの区切り文字.
     *
     */
    static updateSheet(sheet, filePaths, pathSeparator) {
        this.initializeSheet(sheet)
        this.updateHeaderRow(sheet)
        this.updateValueRows(sheet, filePaths, pathSeparator)
        this.doLayout(sheet)
    }

    /**
     *
     * Sheet を初期化する.
     *
     * @param {Sheet} sheet
     *     初期化対象の Sheet.
     *
     */
    static initializeSheet(sheet) {
        sheet.clear()
    }

    /**
     *
     * ヘッダ行を更新する.
     *
     * @param {Sheet} sheet
     *     更新対象の Sheet.
     *
     */
    static updateHeaderRow(sheet) {
        const headers = ["No.", "Type", "MIME Type", "File Path", "File Name"]
        const range = sheet.getRange(1, 1, 1, headers.length)
        range.setValues([headers])
        range.setBackground("orange")
        range.setHorizontalAlignment("center")
    }

    /**
     *
     * ヘッダ以外の行を更新する.
     *
     * @param {Sheet} sheet
     *     更新対象の Sheet.
     *
     * @param {Array.<FilePath>} filePaths
     *     シートに追加する FilePath の配列.
     *
     * @param {string} pathSeparator
     *     パスの区切り文字.
     *
     */
    static updateValueRows(sheet, filePaths, pathSeparator) {
        for (var i = 0; i < filePaths.length; i++) {
            const rowIndex = i + 2
            this.updateValueRow(sheet, filePaths[i], pathSeparator, rowIndex)
        }
    }

    /**
     *
     * ヘッダ以外の行を更新する.
     *
     * @param {Sheet} sheet
     *     更新対象の Sheet.
     *
     * @param {FilePath} filePath
     *     シートに追加する FilePath.
     *
     * @param {string} pathSeparator
     *     パスの区切り文字.
     *
     * @param {number} row
     *     更新対象の Sheet の行インデックス.
     *
     */
    static updateValueRow(sheet, filePath, pathSeparator, row) {
        function setNoCell(column) {
            const range = sheet.getRange(row, column)
            Cells.setNumber(range, "=ROW() - 1")
        }
        function setTypeCell(column) {
            const range = sheet.getRange(row, column)
            const value = Paths.isFile(filePath.file) ? "File" : "Directory"
            Cells.setText(range, value)
        }
        function setMimeTypeCell(column) {
            const range = sheet.getRange(row, column)
            if (Paths.isFile(filePath.file)) {
                Cells.setText(range, filePath.file.getMimeType())
            }
        }
        function setFilePathCell(column) {
            const richText = SpreadsheetApp.newRichTextValue()
            richText.setText(Paths.join(filePath.routes, pathSeparator))

            var startOffset = 0
            filePath.routes.forEach(route => {
                const endOffset = startOffset + route.getName().length
                richText.setLinkUrl(startOffset, endOffset, route.getUrl())
                startOffset = endOffset + pathSeparator.length
            })
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

    /**
     *
     * TODO コメントを書く.
     *
     */
    static doLayout(sheet) {
        const range = sheet.getDataRange()
        range.setVerticalAlignment("top")
        sheet.autoResizeColumns(range.getColumn(), range.getNumColumns())
    }

}

/**
 *
 * ファイルパスの情報を保持する.
 *
 */
class FilePath {

    /**
     *
     * インスタンスを初期化する.
     *
     * @param {File} file
     *     File オブジェクト.
     *
     * @param {Array.<File>} parents
     *     起点となるフォルダから file までのパス (file を含まない).
     *
     * @param {Array.<File>} routes
     *     起点となるフォルダから file までのパス (file を含む).
     *
     */
    constructor(file, parents, routes) {
        this.file = file
        this.parents = parents
        this.routes = routes
    }

}

/**
 *
 * TODO
 *
 */
function DocumentIndex_saveSettings(newSettings) {
    const settings = new DocumentIndexSettings()
    for (var key in newSettings) {
        settings[key] = newSettings[key]
    }
    settings.save()
}

/**
 *
 * ドキュメントインデックス生成ダイアログで "生成" ボタンが押されたときの処理.
 *
 * コールバック関数のためトップレベルに定義する必要がある.
 * クライアントサイドから google.script.run() によって呼び出される.
 * 全ての処理を DocumentIndex.onGenerateButtonClicked() に移譲する.
 *
 * @param {object} options
 *     ドキュメントインデックス生成ダイアログで指定されたオプション.
 *
 */
function DocumentIndex_onGenerateButtonClicked(options) {
    DocumentIndex.onGenerateButtonClicked(options)
}

/**
 *
 * Google Apps Script のトリガーから呼び出される.
 * 定期的にドキュメントインデックスを自動更新するために利用する.
 *
 */
function DocumentIndex_onScheduleTriggered() {
    const settings = DocumentIndex.getSettings()
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    DocumentIndex.generate(
        Sheets.getOrCreateSheetByName(spreadsheet, settings.outputSheetName),
        Paths.getFolderByUrl(settings.rootFolderUrl),
        settings.maxDepth,
        settings.pathSeparator,
        settings.includeFiles,
        settings.includeFolders)
}

