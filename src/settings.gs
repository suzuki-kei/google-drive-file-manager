
/**
 *
 * 設定.
 *
 */
class Settings {

    /**
     *
     * インスタンスを初期化する.
     *
     * @param {Array.<{key: string, type: string, value: object, description: string}>}
     *     設定項目の情報 (キー, データ型, 初期値, 説明).
     *
     */
    constructor(definitions) {
        for (var i = 0; i < definitions.length; i++) {
            this[definitions[i].key] = definitions[i].value
        }
    }

    /**
     *
     * Sheet から設定情報を読み込む.
     *
     * Sheet の 1 行目はヘッダ, 2 列目以降を値として扱う.
     * 読み込んだ設定はインスタンスフィールドに設定される.
     *
     * @param {Sheet} sheet
     *     設定情報を保持する Sheet.
     *
     * @param {Array.<string>} scopes
     *     設定情報のスコープ.
     *     例えば ["A", "B"] を指定すると "A.B." から始まる設定のみ読み込む.
     *
     *
     * @param {string} keyColumn
     *     設定のキーを保持する列の名前.
     *
     * @param {string} typeColumn
     *     設定のデータ型を保持する列の名前.
     *
     * @param {string} valueColumn
     *     設定の値を保持する列の名前.
     *
     * @throws {string}
     *     typeColumn で指定されるデータ型と
     *     valueColumn で指定される値の型が異なる場合.
     *
     */
    load(sheet, scopes, keyColumn, typeColumn, valueColumn) {
        const range = sheet.getDataRange()
        const dictArray = Sheets.getTableAsDictArray(range)
        const scopePrefix = scopes.concat("").join(".")

        for (var i = 0; i < dictArray.length; i++) {
            if (!dictArray[i][keyColumn].startsWith(scopePrefix)) {
                continue
            }
            const key = dictArray[i][keyColumn].substring(scopePrefix.length)
            const type = dictArray[i][typeColumn]
            const value = dictArray[i][valueColumn]
            if (type != typeof(value)) {
                throw "The " + key + " is must be " + type + ", but was " + typeof(value) + "."
            }
            this[key] = value
        }
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
        super([
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
            {
                key: "outputSheetName",
                type: "string",
                value: "Document Index",
                description: "結果を出力するシートの名前.",
            },
        ])
    }

    /**
     *
     * Sheet から設定情報を読み込む.
     *
     */
    load() {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
        const sheet = spreadsheet.getSheetByName("Settings")
        const scopes = ["FileManager", "DocumentIndex"]
        super.load(sheet, scopes, "Key", "Type", "Value")
    }

}

