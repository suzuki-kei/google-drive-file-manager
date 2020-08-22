
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
        menu.addItem("Rename", "DocumentIndex.rename")
        return menu
    }

    /**
     *
     * ファイル名を変更する.
     *
     */
    static rename() {
        const settings = this.getSettings()
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
        const sheet = spreadsheet.getSheetByName(settings.outputSheetName)
        const range = sheet.getDataRange()
        const dictionaries = Sheets.getTableAsRichTextDictionaries(range)

        dictionaries.forEach(dict => {
            const fileUrl = dict["File Name"].getLinkUrl()
            const fileName = dict["File Name"].getText()
            const newFileName = dict["New File Name"].getText()
            if (newFileName) {
                const file = Paths.getFileByUrl(fileUrl)
                file.setName(newFileName)
            }
        })
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
     * @return {DocumentIndexSettings}
     *     設定.
     *
     * @throws {Error}
     *     何らかのエラーが発生した場合.
     *     Settings.SheetNotFound が発生した場合は例外を無視してデフォルト値を返す.
     *
     */
    static getSettings() {
        const settings = new DocumentIndexSettings()
        try {
            settings.load()
        } catch(exception) {
            if (exception instanceof Settings.SheetNotFound) {
                // シートが存在しない場合はデフォルト値を使用して処理を継続する.
            } else {
                throw exception
            }
        }
        return settings
    }

    /**
     *
     * ダイアログの "Save settings" ボタンが押されたときの処理.
     *
     * @param {object} rawSettings
     *     ダイアログで指定された設定情報.
     *     DocumentIndexSettings ではなく通常のオブジェクト.
     *
     */
    static onSaveSettingsClicked(rawSettings) {
        const settings = new DocumentIndexSettings()
        for (let key in rawSettings) {
            settings[key] = rawSettings[key]
        }
        settings.save()
    }

    /**
     *
     * ダイアログの "Generate document index" ボタンが押されたときの処理.
     *
     * @param {object} rawSettings
     *     ダイアログで指定された設定情報.
     *     DocumentIndexSettings ではなく通常のオブジェクト.
     *
     */
    static onGenerateDocumentIndexClicked(rawSettings) {
        const settings = this.getSettings()
        settings.outputSheetName = rawSettings.outputSheetName
        settings.rootFolderUrl = rawSettings.rootFolderUrl

        const sheet = this.generate(settings)
        sheet.activate()
    }

    /**
     *
     * デフォルトの設定でドキュメントインデックスを生成する.
     *
     * Google Apps Script のトリガーを用いて
     * 定期的にドキュメントインデックスを自動更新するために利用する.
     *
     */
    static onScheduleTriggered() {
        const settings = this.getSettings()
        this.generate(settings)
    }

    /**
     *
     * ドキュメントインデックスを生成する.
     *
     * @param {DocumentIndexSettings} settings
     *     設定.
     *
     * @return {Sheet}
     *     生成したドキュメントインデックス.
     *
     */
    static generate(settings) {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
        const sheet = Sheets.getOrCreateSheetByName(spreadsheet, settings.outputSheetName)
        const rootFolder = Paths.getFolderByUrl(settings.rootFolderUrl)
        const maxDepth = settings.maxDepth
        const pathSeparator = settings.pathSeparator
        const includeFiles = settings.includeFiles
        const includeFolders = settings.includeFolders

        this.initializeSheet(sheet)
        this.updateHeaderRow(sheet)
        this.updateValueRows(sheet, rootFolder, maxDepth, pathSeparator, includeFiles, includeFolders)
        this.updateLayout(sheet)
        return sheet
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
        const headers = ["No.", "Type", "File Path", "File Name", "New File Name"]
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
    static updateValueRows(sheet, rootFolder, maxDepth, pathSeparator, includeFiles, includeFolders) {
        let updateInterval = 10
        let count = 0

        Paths.traverse(rootFolder, maxDepth, (parents, file) => {
            if (Paths.isFile(file) && !includeFiles) {
                return
            }
            if (Paths.isFolder(file) && !includeFolders) {
                return
            }
            const routes = parents.concat(file)
            const filePath = new FilePath(file, parents, routes)
            const row = count + 2
            this.updateValueRow(sheet, filePath, pathSeparator, row)

            if (count++ % updateInterval == updateInterval - 1) {
                SpreadsheetApp.flush()
            }
        })
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
        const setNoCell = column => {
            const range = sheet.getRange(row, column)
            Cells.setNumber(range, "=ROW() - 1")
        }
        const setTypeCell = column => {
            const range = sheet.getRange(row, column)
            const value = Paths.isFile(filePath.file) ? "File" : "Directory"
            Cells.setText(range, value)
        }
        const setFilePathCell = column => {
            const richText = SpreadsheetApp.newRichTextValue()
            richText.setText(Paths.join(filePath.routes, pathSeparator))

            let startOffset = 0
            filePath.routes.forEach(route => {
                const endOffset = startOffset + route.getName().length
                richText.setLinkUrl(startOffset, endOffset, route.getUrl())
                startOffset = endOffset + pathSeparator.length
            })
            const range = sheet.getRange(row, column)
            range.setRichTextValue(richText.build())
        }
        const setFileNameCell = column => {
            const range = sheet.getRange(row, column)
            Cells.setTextLink(range, filePath.file.getUrl(), filePath.file.getName())
        }

        let column = 1
        setNoCell(column++)
        setTypeCell(column++)
        setFilePathCell(column++)
        setFileNameCell(column++)
    }

    /**
     *
     * レイアウトを変更する.
     *
     * @param {Sheet} sheet
     *     更新対象の Sheet.
     *
     */
    static updateLayout(sheet) {
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
 * ダイアログの "Save settings" ボタンが押されたときの処理.
 *
 * コールバック関数のためトップレベルに定義している.
 * ダイアログ側から google.script.run() によって呼び出される.
 * 全ての処理を DocumentIndex.onSaveSettings() に移譲する.
 *
 * @param {object} rawSettings
 *     ダイアログで指定された設定情報.
 *     DocumentIndexSettings ではなく通常のオブジェクト.
 *
 */
function DocumentIndex_onSaveSettingsClicked(rawSettings) {
    DocumentIndex.onSaveSettingsClicked(rawSettings)
}

/**
 *
 * ダイアログの "Generate document index" ボタンが押されたときの処理.
 *
 * コールバック関数のためトップレベルに定義する必要がある.
 * クライアントサイドから google.script.run() によって呼び出される.
 * 全ての処理を DocumentIndex.onGenerateButtonClicked() に移譲する.
 *
 * @param {object} rawSettings
 *     ダイアログで指定された設定情報.
 *     DocumentIndexSettings ではなく通常のオブジェクト.
 *
 */
function DocumentIndex_onGenerateDocumentIndexClicked(rawSettings) {
    DocumentIndex.onGenerateDocumentIndexClicked(rawSettings)
}

/**
 *
 * Google Apps Script のトリガーから呼び出される.
 * 定期的にドキュメントインデックスを自動更新するために利用する.
 *
 */
function DocumentIndex_onScheduleTriggered() {
    DocumentIndex.onScheduleTriggered()
}

