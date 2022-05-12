package com.kaos.excel.dto

/*

{
    "fileName": "grocery_stock",
    "sheet": [
        {
            "name": "fruit",
            "isRowTitle": true,
            "data": {
                "fruit": [
                    "apple",
                    "orange",
                    "strawberry"
                ],
                "stock": [
                    2007,
                    101,
                    4899
                ],
                "price": [
                    1000,
                    1500,
                    500
                ]
            },
            "titleOption": {
                "fontColor": "996633",
                "fontSize": 20,
                "bold": true,
                "italic": true,
                "backgroundColor": "EEDDFF"
            },
            "arrangeOption": {
                "leftGap": 1,
                "topGap": 3
            },
            "borderOption": {
                "inside" : true,
                "insideColor" : "coral",
                "outside" : true,
                "outsideColor" : "violet"
            }
        },
        {
            "name": "drink",
            "isRowTitle": true,
            "data": {
                "drink": [
                    "water",
                    "juice",
                    "sprit"
                ],
                "stock": [
                    2007,
                    101,
                    4899
                ],
                "price": [
                    1000,
                    1500,
                    500
                ]
            }
        }
    ]
}

 */

data class ExcelInputDto (
    val fileName: String, // excel file path
    val sheet: Array<SheetDto>, // sheet information
    ) {
    override fun equals(other: Any?): Boolean {
        if (this === other) return true
        if (javaClass != other?.javaClass) return false

        other as ExcelInputDto

        if (fileName != other.fileName) return false
        if (!sheet.contentEquals(other.sheet)) return false

        return true
    }

    override fun hashCode(): Int {
        var result = fileName.hashCode()
        result = 31 * result + sheet.contentHashCode()
        return result
    }
}





