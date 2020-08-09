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
var Paths = {

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
 * Google Spreadsheet のセルに関するユーティリティ.
 *
 */
var Cells = {

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
