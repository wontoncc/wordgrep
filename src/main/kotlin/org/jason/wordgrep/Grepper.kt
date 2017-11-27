package org.jason.wordgrep

import kotlinx.coroutines.experimental.*
import org.apache.poi.hwpf.extractor.WordExtractor
import org.apache.poi.xwpf.extractor.XWPFWordExtractor
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.fusesource.jansi.Ansi.ansi
import org.fusesource.jansi.AnsiConsole
import java.io.File
import java.io.FileInputStream
import java.io.OutputStream
import java.io.PrintStream

class NullOutputStream: OutputStream() {
    override fun write(b: Int) {
    }
}
val nullos = NullOutputStream()

fun doc2string (path: String): String {
    var result = ""
    try {
        result = WordExtractor(FileInputStream(path)).textFromPieces
    } catch (e: java.lang.IllegalArgumentException) {
        println("ERROR: document type error: ${path}, treating as .docx ...\n")
        result = docx2string(path)
    } catch (e: Exception) {
        println("* error while extracting ${path}\n")
        e.printStackTrace()
    }
    return result
}

fun docx2string (path: String): String {
    var result = ""
    try {
        // TODO: No disabling `stderr` while exception catchable
        val origErr = System.err
        System.setErr(PrintStream(nullos))

        result = XWPFWordExtractor(XWPFDocument(FileInputStream(path))).text

        System.setErr(origErr)
    } catch (e: Exception) {
        println("* error while extracting ${path}")
    }
    return result
}

fun getDocumentsTree(path: String): List<File> {
    val current = File(path)
    var documents = current.listFiles { _, name ->
        val normalized = name.toLowerCase()
        normalized.endsWith(".docx") or normalized.endsWith(".doc") and !normalized.startsWith("~$")
    }.toList()
    val directories = current.listFiles { item ->
        item.isDirectory()
    }
    directories.map { dir ->
        val inner = getDocumentsTree(dir.absolutePath)
        documents += inner
    }
    return documents
}

suspend fun singleDocumentGrep (pattern: String, path: String): List<String> {
    var digests = listOf<String>()
    val content = when {
        path.endsWith(".docx") -> docx2string(path)
        path.endsWith(".doc") -> doc2string(path)
        else -> ""
    }
    if (content.contains(pattern)) {
        var start = 0
        while (start + pattern.length < content.length) {
            val position = content.indexOf(pattern, start)
            if (position > -1) {
                var cutStart = position - 50
                var cutEnd = position + 50 + pattern.length
                if (cutStart < 0) {
                    cutStart = 0
                }
                if (cutEnd > content.length - 1) {
                    cutEnd = content.length - 1
                }
                val digest =
                        content.slice(IntRange(cutStart, cutEnd))
                                .replace("\r", "")
                                .trimEnd()
                                .trimStart()
                digests += digest
                start = position + pattern.length
            } else {
                break
            }
        }
    }
    return digests
}

fun grep (pattern: String, tree: List<File>): Map<String, List<String>> = runBlocking(CommonPool) {
    var deferred = mapOf<String, Deferred<List<String>>>()
    var result = mapOf<String, List<String>>()
    val job = launch(CommonPool) {
        for (f in tree) {
            val path = f.absolutePath
            val digests = async(CommonPool) {
                singleDocumentGrep(pattern, path)
            }
            deferred = deferred.plus(path to digests)
        }

        for ((p, d) in deferred) {
            async(CommonPool) {
                val dd = d.await()
                if (dd.isNotEmpty()) {
                    result += p to dd
                }
            }.await()
        }
    }
    job.join()
    return@runBlocking result
}

fun main (args: Array<String>) {
    if (!System.getenv().keys.contains("DEBUG")) {
        AnsiConsole.systemInstall()
    }

    val result = when (args.size) {
        0 -> mapOf<String, List<String>>()
        1 -> grep(args[0], getDocumentsTree(System.getProperty("user.dir")))
        else -> {
            var collected = mapOf<String, List<String>>()
            for (i in 1..(args.size - 1)) {
                collected += grep(args[0], getDocumentsTree(args[i]))
            }
            collected
        }
    }

    result.forEach { path, digests ->
        println((ansi().fgCyan().a(path).reset()))
        for (digest in digests) {
            println("...${digest.replace(args[0], ansi().fgRed().a(args[0]).reset().toString())}...")
        }
        println()
    }
    AnsiConsole.systemUninstall()
}