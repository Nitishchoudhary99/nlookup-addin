/**
 * NLOOKUP custom function - search a value in the first column and return a value in the same row from another column.
 */
function NLOOKUP(lookupValue, lookupRange, returnRange) {
    for (let i = 0; i < lookupRange.length; i++) {
        if (lookupRange[i][0] === lookupValue) {
            return returnRange[i][0];
        }
    }
    return "Not Found";
}

// Associate the function
CustomFunctions.associate("NLOOKUP", NLOOKUP);
