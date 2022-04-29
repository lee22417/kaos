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
                "fontColor": "993366",
                "fontSize": 20,
                "bold": true,
                "italic": true,
                "backgroundColor": "AAAAFF",
                "setBorder": true //TODO
            },
            "arrangeOption": {
                "leftGap": 1, //TODO
                "upGap": 3  //TODO
            },
            "decorationOption": {
                "borderColor": "red", //TODO
                "borderSize": 5 //TODO
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





