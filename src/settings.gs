
/**
 *
 * 設定シートが存在しないことを表す例外.
 *
 */
class SettingsSheetNotFound extends Error {

    /**
     *
     * インスタンスを初期化する.
     *
     * @param {string} sheetName
     *     シート名.
     *
     */
    constructor(sheetName) {
        super("Sheet Not Found: " + sheetName)
    }

}

/**
 *
 * 設定項目の値が期待するデータ型ではないことを表す例外.
 *
 */
class SettingsInvalidDataType extends Error {

    /**
     *
     * インスタンスを初期化する.
     *
     * @param {string} key
     *     設定項目のキー.
     *
     * @param {string} type
     *     設定項目のデータ型.
     *
     * @param {object} value
     *     設定項目の値.
     *
     */
    constructor(key, type, value) {
        super("The " + key + " is must be " + type + ", but was " + typeof(value) + ".")
    }

}

/**
 *
 * シートに対して設定情報を読み書きする機能を提供する.
 *
 * シートの 1 行目をヘッダとし, 2 行目以降を設定項目として扱う.
 *
 * 設定項目は {キー, データ型, 値, 説明} で構成され,
 * インスタンスのフィールドを経由してアクセスすることができる.
 * そのためキーは JavaScript の識別子として使用可能な名前である必要がある.
 *
 * この制限を見かけ上回避するために,
 * コンストラクタでキーのプレフィックスを指定できる.
 * Settings.load(), Settings.save() はシートに読み書きするときに,
 * キーに対して自動的にプレフィックスを除去もしくは付加する.
 *
 */
class Settings {

    /**
     *
     * インスタンスを初期化する.
     *
     * @param {string} sheetName
     *     設定情報を保持するシートの名前.
     *
     * @param {{key: string, type: string, value: string, description: string}} headerNames
     *     設定情報を保持するシートのヘッダ名.
     *     key, type, value, description にはそれぞれ
     *     設定項目のキー, データ型, 値, 説明を保持するヘッダの名前を指定する.
     *
     * @param {string} keyPrefix
     *     設定項目のキーのプレフィックス.
     *
     * @param {Array.<{key: string, type: string, value: object, description: string}>} definitions
     *     設定項目の情報 (キー, データ型, 初期値, 説明).
     *
     */
    constructor(sheetName, headerNames, keyPrefix, definitions) {
        this.sheetName = sheetName
        this.headerNames = headerNames
        this.keyPrefix = keyPrefix
        this.definitions = definitions

        definitions.forEach(definition => {
            const key = definition.key
            const type = definition.type
            const value = definition.value
            this.setItem(key, type, value)
        })
    }

    /**
     *
     * 設定を更新する.
     *
     * @param {string} key
     *     設定項目のキー.
     *
     * @param {string} type
     *     設定項目のデータ型.
     *
     * @param {object} value
     *     設定項目の値.
     *
     * @throws {SettingsInvalidDataType}
     *     設定項目の値が期待するデータ型ではない場合.
     *
     */
    setItem(key, type, value) {
        if (type != typeof(value)) {
            throw new SettingsInvalidDataType(key, type, value)
        }
        this[key] = value
    }

    /**
     *
     * 設定を保持するシートを取得する.
     *
     * @return {Sheet}
     *     設定を保持するシート.
     *
     * @throws {SettingsSheetNotFound}
     *     シートが存在しない場合.
     *
     */
    getSheet() {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
        const sheet = spreadsheet.getSheetByName(this.sheetName)
        if (sheet === null) {
            throw new SettingsSheetNotFound(this.sheetName)
        }
        return sheet
    }

    /**
     *
     * 設定を保持するシートを取得もしくは作成する.
     *
     * @return {Sheet}
     *     設定を保持するシート.
     *     シートが存在しない場合は新しく作成した空のシート.
     *
     */
    getOrCreateSheet() {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
        return Sheets.getOrCreateSheetByName(spreadsheet, this.sheetName)
    }

    /**
     *
     * Sheet から設定情報を読み込む.
     *
     * @throws {SettingsSheetNotFound}
     *     シートが存在しない場合.
     *
     * @throws {SettingsInvalidDataType}
     *     設定項目の値が期待するデータ型ではない場合.
     *
     */
    load() {
        const sheet = this.getSheet()
        const range = sheet.getDataRange()
        const dictArray = Sheets.getTableAsDictArray(range)

        dictArray.forEach(dict => {
            if (!dict[this.headerNames.key].startsWith(this.keyPrefix)) {
                return
            }
            const key = dict[this.headerNames.key].substring(this.keyPrefix.length)
            const type = dict[this.headerNames.type]
            const value = dict[this.headerNames.value]
            this.setItem(key, type, value)
        })
    }

    /**
     *
     * 設定情報をシートに保存する.
     *
     */
    save() {
        const prepareSheet = () => {
            const sheet = this.getOrCreateSheet()
            sheet.clear()
            return sheet
        }
        const updateHeaderRow = sheet => {
            const values = [
                this.headerNames.key,
                this.headerNames.type,
                this.headerNames.value,
                this.headerNames.description,
            ]
            const range = sheet.getRange(1, 1, 1, values.length)
            range.setValues([values])
            range.setBackground("orange")
            range.setHorizontalAlignment("center")
        }
        const updateValueRows = sheet => {
            for (let i = 0; i < this.definitions.length; i++) {
                const key = this.keyPrefix + this.definitions[i].key
                const type = this.definitions[i].type
                const value = this[this.definitions[i].key]
                const description = this.definitions[i].description
                Cells.setValue(sheet.getRange(i + 2, 1), key)
                Cells.setValue(sheet.getRange(i + 2, 2), type)
                Cells.setValue(sheet.getRange(i + 2, 3), value)
                Cells.setValue(sheet.getRange(i + 2, 4), description)
            }
        }
        const updateLayout = sheet => {
            const range = sheet.getDataRange()
            range.setVerticalAlignment("top")
            range.setHorizontalAlignment("left")
            sheet.autoResizeColumns(range.getColumn(), range.getNumColumns())
        }

        const sheet = prepareSheet()
        updateHeaderRow(sheet)
        updateValueRows(sheet)
        updateLayout(sheet)
    }

}

/**
 *
 * ドキュメントインデックスの設定.
 *
 */
class DocumentIndexSettings extends Settings {

    /**
     *
     * インスタンスを初期化する.
     *
     */
    constructor() {
        super(
            "Settings",
            {
                key: "Key",
                type: "Type",
                value: "Value",
                description: "Description",
            },
            "FileManager.DocumentIndex.",
            [
                {
                    key: "rootFolderUrl",
                    type: "string",
                    value: Paths.getCurrentFolder().getUrl(),
                    description: "探索の起点とするフォルダの URL.",
                },
                {
                    key: "maxDepth",
                    type: "number",
                    value: 5,
                    description: "再帰的にサブフォルダを探索するときの最大の深さ.",
                },
                {
                    key: "outputSheetName",
                    type: "string",
                    value: "Document Index",
                    description: "結果を出力するシートの名前.",
                },
                {
                    key: "pathSeparator",
                    type: "string",
                    value: " > ",
                    description: "パスの区切りに使用する文字列.",
                },
                {
                    key: "includeFiles",
                    type: "boolean",
                    value: true,
                    description: "結果にファイルを含める場合に真.",
                },
                {
                    key: "includeFolders",
                    type: "boolean",
                    value: true,
                    description: "結果にフォルダを含める場合に真.",
                },
            ]
        )
    }

}

