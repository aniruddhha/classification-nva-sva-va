
class Converter {

    constructor(xlJson) {
        xlJson.map(row => this.convert(row))
    }

    convert(row) {
        const arr = getClassificationArray(row[`Category`])
        const rwDs = row[`Description`].toLowerCase()
        arr.forEach(classifier => {
            const keywords = classifier[`keywords`]
            keywords.forEach(kw => {
                if (rwDs.includes(kw.toLowerCase())) {
                    row[`type`] = classifier[`descrition`]
                    return
                }
            })
        })
    }

    getClassificationArray(category) {
        if (category == 'nva') return []
        else if (category == 'sva') return []
        return []
    }
}

module.exports = Converter