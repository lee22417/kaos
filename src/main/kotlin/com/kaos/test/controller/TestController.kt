package com.kaos.test.controller

import org.springframework.http.ResponseEntity
import org.springframework.web.bind.annotation.GetMapping
import org.springframework.web.bind.annotation.RequestMapping
import org.springframework.web.bind.annotation.RestController

@RestController
@RequestMapping("/test")
class TestController {
    @GetMapping()
    fun test(): ResponseEntity<String> {
        val msg = "Hello World"
        return ResponseEntity.ok(msg)
    }
}

