/**
 *
 * 全体で使用するユーティリティ.
 *
 */

/**
 *
 * Google Drive のファイルパスに関するユーティリティ.
 *
 */
const Paths = {

    /**
     *
     * File がファイルであることを判定する.
     *
     * @param {File} file
     *     File オブジェクト.
     *
     * @return {boolean}
     *     file がファイルである場合は true.
     *     file がフォルダである場合は false.
     *
     */
    isFile: function(file) {
        return !file.addFile
    },

    /**
     *
     * File がフォルダであることを判定する.
     *
     * @param {File} file
     *     File オブジェクト.
     *
     * @return {boolean}
     *     file がフォルダである場合は true.
     *     file がファイルである場合は false.
     *
     */
    isFolder: function(file) {
        return !!file.addFile
    },

    /**
     *
     * Google Drive のルートフォルダを取得する.
     *
     * @preturn {File}
     *     Google Drive のルートフォルダ.
     *
     */
    getRootFolder: function() {
        return DriveApp.getRootFolder()
    },

    /**
     *
     * このスクリプトが関連付く Spreadsheet が保存されているフォルダを取得する.
     *
     * @preturn {File}
     *     このスクリプトが関連付く Spreadsheet が保存されているフォルダ.
     *
     */
    getCurrentFolder: function() {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
        const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId())
        return spreadsheetFile.getParents().next()
    },

    /**
     *
     * フォルダの URL から Folder オブジェクトを取得する.
     *
     * @param {string} url
     *     フォルダの URL.
     *
     * @return {File}
     *     Folder オブジェクト.
     *
     */
    getFolderByUrl: function(url) {
        const urlPrefix = "https://drive.google.com/drive/folders/"
        const folderId = url.replace(urlPrefix, "").split("?")[0]
        return DriveApp.getFolderById(folderId)
    },

    /**
     *
     * パスを連結する.
     *
     * @param {Array.<File>} paths
     *     連結するパス.
     *
     * @param {string} pathSeparator
     *     パスの区切り文字.
     *
     * @return {string}
     *     パスを連結した文字列.
     *
     */
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

/**
 *
 * Google Spreadsheet のシートに関するユーティリティ.
 *
 */
const Sheets = {

    /**
     *
     * 指定した名前のシートを取得もしくは作成する.
     *
     * @param {Spreadsheet} spreadsheet
     *     捜査対象の Spreadsheet.
     *
     * @param {string} sheetName
     *     シートの名前.
     *
     * @return {Sheet}
     *     取得もしくは作成したシート.
     *
     */
    getOrCreateSheetByName: function(spreadsheet, sheetName) {
        const sheet = spreadsheet.getSheetByName(sheetName)
        if (sheet) {
            return sheet
        } else {
            const newSheet = spreadsheet.insertSheet()
            newSheet.setName(sheetName)
            return newSheet
        }
    },

    /**
     *
     * Range の内容を辞書の配列として取得する.
     *
     * @param {Range} range
     *     値を取得する範囲.
     *
     * @return {Array.<object>}
     *     range の 1 行目をキーとした辞書の配列.
     *
     */
    getTableAsDictArray: function(range) {
        const values = range.getValues()
        const dictArray = []

        for (var row = 1; row < range.getNumRows(); row++) {
            const dict = {}
            for (var column = 0; column < range.getNumColumns(); column++) {
                dict[values[0][column]] = values[row][column]
            }
            dictArray.push(dict)
        }
        return dictArray
    },

}

/**
 *
 * Google Spreadsheet のセルに関するユーティリティ.
 *
 */
const Cells = {

    /**
     *
     * セルに値を設定する.
     *
     * セルのフォーマットは "自動" を設定する.
     * セルのフォーマットを明示的に指定したい場合は他の関数を使用する必要がある.
     *
     * @param {Range} range
     *     値を設定するセル.
     *
     * @param {object} value
     *     設定する値.
     *
     */
    setValue: function(range, value) {
        range.setValue(value)
        range.setNumberFormat("General")
    },

    /**
     *
     * セルに数値として値を設定する.
     *
     * @param {Range} range
     *     値を設定するセル.
     *
     * @param {number|string} value
     *     設定する値.
     *     数値を指定するか "=ROW()" のように数値に評価される数式を指定する.
     *
     */
    setNumber: function(range, value) {
        range.setValue(value)
        range.setNumberFormat("0")
    },

    /**
     *
     * セルに文字列として値を設定する.
     *
     * @param {Range} range
     *     値を設定するセル.
     *
     * @param {string} value
     *     設定する値.
     *
     */
    setText: function(range, value) {
        range.setValue(value)
        range.setNumberFormat("@")
    },

    /**
     *
     * セルにリンク文字列として値を設定する.
     *
     * @param {Range} range
     *     値を設定するセル.
     *
     * @param {string} url
     *     設定する URL.
     *
     * @param {string} value
     *     表示文字列.
     *
     */
    setTextLink: function(range, url, value) {
        value = value.replace(/"/g, '""')
        range.setValue('=HYPERLINK("' + url + '", "' + value + '")')
        range.setNumberFormat("@")
        range.setShowHyperlink(true)
    },

}
