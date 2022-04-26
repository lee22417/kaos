package com.kaos.controller

import org.springframework.http.ResponseEntity
import org.springframework.stereotype.Controller
import org.springframework.ui.Model
import org.springframework.web.bind.annotation.GetMapping
import org.springframework.web.bind.annotation.RequestMapping
import org.springframework.web.bind.annotation.RestController
import org.springframework.web.servlet.ModelAndView

@RestController
@RequestMapping("/test")
class TestController {
    @GetMapping()
    fun test(): ResponseEntity<String> {
        val msg = "Hello World"
        return ResponseEntity.ok(msg)
    }
}

