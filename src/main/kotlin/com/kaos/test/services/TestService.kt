package com.kaos.test.services

import org.springframework.stereotype.Service

@Service
class TestService {
    fun getName():String{
        return "Hello World!"
    }
}