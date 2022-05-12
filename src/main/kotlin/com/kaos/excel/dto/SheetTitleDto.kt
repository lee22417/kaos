package com.kaos.excel.dto

data class SheetTitleDto (
    val fontColor: String?, // title color - RGB HEX code
    val fontSize: Int?, // title font size
    val backgroundColor: String?, // title cell background color - RGB HEX code
    val bold: Boolean?, // set title font as bold
    val italic: Boolean?, // set title font as italic
)