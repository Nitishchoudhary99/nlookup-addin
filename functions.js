
/**
 * NLOOKUP: A powerful lookup function that supports wildcards, multi-column return, etc.
 * @customfunction
 * @param {string} lookupValue The value to search for.
 * @param {range} lookupArray The range to search within.
 * @param {range} returnArray The range of values to return from.
 * @param {boolean} [exactMatch] TRUE for exact match, FALSE for partial/wildcard match.
 * @param {string} [notFoundText] Text to return if no match is found.
 * @returns {any[][]} The found value(s).
 */
function NLOOKUP(lookupValue, lookupArray, returnArray, exactMatch = true, notFoundText = "Not Found") {
    const result = [];
    const lowerLookup = lookupValue.toLowerCase();

    for (let i = 0; i < lookupArray.length; i++) {
        let match = exactMatch
            ? lookupArray[i][0].toLowerCase() === lowerLookup
            : lowerLookup.includes("*")
                ? new RegExp("^" + lowerLookup.replace(/\*/g, ".*") + "$", "i").test(lookupArray[i][0])
                : lookupArray[i][0].toLowerCase().includes(lowerLookup);

        if (match) {
            result.push(returnArray[i]);
        }
    }

    return result.length > 0 ? result : [[notFoundText]];
}

CustomFunctions.associate("NLOOKUP", NLOOKUP);
