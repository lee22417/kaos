package com.kaos.excel.dto

/*

{
    "fileName": "grocery_stock",
    "sheet": [
        {
            "name": "fruit",
            "isRowTitle": false,
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
            "arrangeOption": {
                "leftGap": 1,
                "upGap": 3
            },
            "decorationOption": {
                "titleColor": "blue",
                "borderColor": "red"
            }
        },
        {
            "name": "drink",
            "isRowTitle": true,
            "data": {
                "fruit": [
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

// sheet information
data class SheetDto (
    val name: String, // sheet name
    val isRowTitle: Boolean, // whether title is in row or col
    val data: HashMap<String, Array<String>>, // data written in excel file
    val arrangeOption: SheetArrangeDto?, // sheet table arrange option
    val decorationOption: SheetDecorationDto? // data decoration option
)

// sheet data table arrange
data class SheetArrangeDto (
    val leftGap: Int?, // empty cell on the left table
    val upGap: Int?, // empty cell above table
)

// sheet data decoration
data class SheetDecorationDto (
    val titleColor: String?, //
    val borderSize: Int?,
    val borderColor: String?,
)