package com.kaos.excel.dto

/*
Color Index
https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/IndexedColors.html
*/

// sheet data decoration
data class SheetBorderDto (
    val inside: Boolean?, // table inside border
    val insideColor: String?, // inside border color
    val outside: Boolean?, // table outside border
    val outsideColor: String?, // outside border color
)