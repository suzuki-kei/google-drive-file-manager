
/**
 *
 * Sheet に保持された設定情報を扱う.
 *
 */
const Settings = {

    /**
     *
     * Sheet から設定情報を読み込む.
     *
     * Sheet の 1 行目はヘッダ, 2 列目以降を値として扱う.
     *
     * @param {Sheet} sheet
     *     設定情報を保持する Sheet.
     *
     * @param {string} keyColumn
     *     設定のキー名を保持する列の名前.
     *
     * @param {string} typeColumn
     *     設定のデータ型を保持する列の名前.
     *
     * @param {string} valueColumn
     *     設定の値を保持する列の名前.
     *
     * @return {{key: string, value: object}}
     *     設定情報を保持するオブジェクト.
     *     key は keyColumn で指定された列の値.
     *     value は valueColumn で指定された列の値.
     *
     * @throws {string}
     *     typeColumn で指定されるデータ型と実際の値が異なる場合.
     *
     */
    fromSheet: function(sheet, keyColumn, typeColumn, valueColumn) {
        const range = sheet.getDataRange()
        return this.fromRange(range, keyColumn, typeColumn, valueColumn)
    },

    fromRange: function(range, keyColumn, typeColumn, valueColumn) {
        const dictArray = Sheets.getTableAsDictArray(range)
        return this.fromDictArray(dictArray, keyColumn, typeColumn, valueColumn)
    },

    fromDictArray: function(dictArray, keyColumn, typeColumn, valueColumn) {
        const settings = {}
        for (var i = 0; i < dictArray.length; i++) {
            const key = dictArray[i][keyColumn]
            const type = dictArray[i][typeColumn]
            const value = dictArray[i][valueColumn]
            if (type != typeof(value)) {
                throw "The " + key + " is must be " + type + ", but was " + typeof(value) + "."
            }
            settings[key] = value
        }
        return settings
    },

    /**
     *
     * 指定されたスコープの設定だけを取り出す.
     *
     *     var settings = {
     *         "taro.name": "taro",
     *         "taro.email": "taro@example.com",
     *         "jiro.user": "jiro",
     *         "jiro.password": "jiro@example.com",
     *     }
     *     Settings.scope(settings, "taro") // => {"name": "taro", "email": "taro@example.com"}
     *
     * @param {object} settings
     *     設定.
     *
     * @param {string} name
     *     スコープ.
     *     "common" を指定すると "common." で始まる設定のみを取り出す.
     *
     * @return {object}
     *     name で指定された設定.
     *
     */
    scope: function(settings, name) {
        const prefix = name + "."
        const scoped = {}

        for (const key in settings) {
            if (!key.startsWith(prefix)) {
                continue
            }
            scoped[key.substring(prefix.length)] = settings[key]
        }
        return scoped
    },

    /**
     *
     * 設定をマージする.
     *
     * マージ対象の設定を可変長引数で受け取る.
     * 空のオブジェクトに対して指定された設定を先頭から順番に上書きする.
     *
     * @return {object}
     *     マージした設定.
     *     引数を 1 つも指定しない場合は空のオブジェクト.
     *
     */
    merge: function() {
        const merged = {}

        for (var i = 0; i < arguments.length; i++) {
            for (const key in arguments[i]) {
                merged[key] = arguments[i][key]
            }
        }
        return merged
    },

}

