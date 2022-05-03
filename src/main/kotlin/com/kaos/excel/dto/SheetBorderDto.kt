package com.kaos.excel.dto

// sheet data decoration
data class SheetBorderDto (
    val borderColor: String?,   // table border color
    val inside: Boolean?, // table inside border
    val outside: Boolean?, // table outside border
)