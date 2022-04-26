package com

import org.springframework.boot.autoconfigure.SpringBootApplication
import org.springframework.boot.runApplication

@SpringBootApplication
class KaosApplication

fun main(args: Array<String>) {
	runApplication<KaosApplication>(*args)
}
