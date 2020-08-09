
/**
 *
 * Sheet に保持された設定情報を扱う.
 *
 */
const Settings = {

    /**
     *
     * 指定した Sheet から設定情報を読み込む.
     *
     * @param {Sheet} sheet
     *     設定情報を保持する Sheet.
     *     シートの 1 行目をヘッダとして扱う.
     *
     * @param {string} keyColumnName
     *     プロパティのキーを保持する列のヘッダ名.
     *
     * @param {string} typeColumnName
     *     プロパティのデータ型を保持する列のヘッダ名.
     *
     * @param {string} valueColumnName
     *     プロパティの値を保持する列のヘッダ名.
     *
     * @return {key: string, value: object}
     *     設定情報を保持するオブジェクト.
     *
     */
    load: function(sheet, keyColumnName, typeColumnName, valueColumnName) {
        const range = sheet.getDataRange()
        const values = range.getValues()

        const properties = []
        for (var row = 1; row < range.getNumRows(); row++) {
            const setting = {}
            for (var column = 0; column < range.getNumColumns(); column++) {
                setting[values[0][column]] = values[row][column]
            }
            properties.push(setting)
        }

        const settings = {}
        for (var i = 0; i < properties.length; i++) {
            const key = properties[i][keyColumnName]
            const type = properties[i][typeColumnName]
            const value = properties[i][valueColumnName]
            if (type != typeof(value)) {
                throw "" + key + " is must be " + type + ", but was " + typeof(value) + "."
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

