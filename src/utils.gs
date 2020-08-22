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
class Paths {

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
    static isFile(file) {
        return !file.addFile
    }

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
    static isFolder(file) {
        return !!file.addFile
    }

    /**
     *
     * Google Drive のルートフォルダを取得する.
     *
     * @preturn {File}
     *     Google Drive のルートフォルダ.
     *
     */
    static getRootFolder() {
        return DriveApp.getRootFolder()
    }

    /**
     *
     * このスクリプトが関連付く Spreadsheet が保存されているフォルダを取得する.
     *
     * @preturn {File}
     *     このスクリプトが関連付く Spreadsheet が保存されているフォルダ.
     *
     */
    static getCurrentFolder() {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
        const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId())
        return spreadsheetFile.getParents().next()
    }

    /**
     *
     * ファイルやフォルダの URL から ID を取得する.
     *
     * @param {string} url
     *     ファイルやフォルダの URL.
     *
     * @return {string}
     *     ファイルやフォルダの ID.
     *
     */
    static getIdFromUrl(url) {
        return url.split("/")[5]
    }

    /**
     *
     * ファイルの URL から File オブジェクトを取得する.
     *
     * @param {string} url
     *     ファイルの URL.
     *
     * @return {string}
     *     File オブジェクト.
     *
     */
    static getFileByUrl(url) {
        return DriveApp.getFileById(this.getIdFromUrl(url))
    }

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
    static getFolderByUrl(url) {
        if (url === "https://drive.google.com/drive/my-drive") {
            return DriveApp.getRootFolder()
        }
        return DriveApp.getFolderById(this.getIdFromUrl(url))
    }

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
    static join(paths, pathSeparator) {
        let value = ""
        let separator = ""
        paths.forEach(path => {
            value += separator + path.getName()
            separator = pathSeparator
        })
        return value
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
     *     callback(index, parents, file) という形式で呼び出される.
     *     parents は rootFolder から file までのパス (file を含まない).
     *
     */
    static traverse(rootFolder, maxDepth, callback) {
        const traverse = (parents, folder, depth, maxDepth, callback) => {
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
        traverse([], rootFolder, 1, maxDepth, callback)
    }

}

/**
 *
 * Google Spreadsheet のシートに関するユーティリティ.
 *
 */
class Sheets {

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
    static getOrCreateSheetByName(spreadsheet, sheetName) {
        const sheet = spreadsheet.getSheetByName(sheetName)
        if (sheet) {
            return sheet
        } else {
            const newSheet = spreadsheet.insertSheet()
            newSheet.setName(sheetName)
            return newSheet
        }
    }

    /**
     *
     * Range の内容をテキストの辞書配列として取得する.
     *
     * @param {Range} range
     *     値を取得する範囲.
     *
     * @return {Array.<object>}
     *     range の 1 行目をキーとした辞書の配列.
     *
     */
    static getTableAsTextDictionaries(range) {
        const values = range.getValues()
        const dictionaries = []

        for (let row = 1; row < range.getNumRows(); row++) {
            const dictionary = {}
            for (let column = 0; column < range.getNumColumns(); column++) {
                dictionary[values[0][column]] = values[row][column]
            }
            dictionaries.push(dictionary)
        }
        return dictionaries
    }

    /**
     *
     * Range の内容を RichText の辞書配列として取得する.
     *
     * @param {Range}
     *     値を取得する範囲.
     *
     * @return {Array.<object>}
     *     range の一行目をキーとした辞書の配列.
     *
     */
    static getTableAsRichTextDictionaries(range) {
        const values = range.getRichTextValues()
        const dictionaries = []

        for (let row = 1; row < range.getNumRows(); row++) {
            const dictionary = {}
            for (let column = 0; column < range.getNumColumns(); column++) {
                dictionary[values[0][column].getText()] = values[row][column]
            }
            dictionaries.push(dictionary)
        }
        return dictionaries
    }

}

/**
 *
 * Google Spreadsheet のセルに関するユーティリティ.
 *
 */
class Cells {

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
    static setValue(range, value) {
        range.setValue(value)
        range.setNumberFormat("General")
    }

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
    static setNumber(range, value) {
        range.setValue(value)
        range.setNumberFormat("0")
    }

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
    static setText(range, value) {
        range.setValue(value)
        range.setNumberFormat("@")
    }

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
    static setTextLink(range, url, value) {
        value = value.replace(/"/g, '""')
        range.setValue('=HYPERLINK("' + url + '", "' + value + '")')
        range.setNumberFormat("@")
        range.setShowHyperlink(true)
    }

}

