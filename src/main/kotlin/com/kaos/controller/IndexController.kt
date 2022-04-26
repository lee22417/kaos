package com.kaos.controller

import org.springframework.stereotype.Controller
import org.springframework.ui.Model
import org.springframework.ui.set
import org.springframework.web.bind.annotation.GetMapping

@Controller
class IndexController {
    @GetMapping("/")
    fun getIndex(model :Model):String{
        model["title"] = "index page" // index.mustache - {{title}}
        model["msg"] = "Hello World!" // index.mustache - {{msg}}
        return "index"
    }
}