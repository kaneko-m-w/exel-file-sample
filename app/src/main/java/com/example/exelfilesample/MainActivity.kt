package com.example.exelfilesample

import android.app.Activity
import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.widget.Button
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter

class MainActivity : AppCompatActivity() {
    private val CREATE_FILE = 1002

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        var button: Button = findViewById(R.id.write)
        button.setOnClickListener() {
            createFile()
        }
    }

    private fun createFile() {
        val current = LocalDateTime.now()
        val formatter = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss")
        val formatted = current.format(formatter)
        val fileName = "${formatted}.xlsx"
        val intent = Intent(Intent.ACTION_CREATE_DOCUMENT)
        intent.addCategory(Intent.CATEGORY_OPENABLE)
        intent.type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        intent.putExtra(Intent.EXTRA_TITLE, fileName)
        startActivityForResult(intent, CREATE_FILE)
    }

    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)
        if (requestCode == CREATE_FILE && resultCode == Activity.RESULT_OK) {
            if (data != null) {
                var uri: Uri? = data.data

                var hssfWorkbook: HSSFWorkbook = HSSFWorkbook()
                var hssfSheet: HSSFSheet = hssfWorkbook.createSheet()

                var count: Int = 0
                for (y in 0..5) {
                    var hssfRow: HSSFRow = hssfSheet.createRow(y)
                    for (x in 0..5) {
                        var hssfCell: HSSFCell = hssfRow.createCell(x)
                        hssfCell.setCellValue((count++).toString())
                    }
                }

                try {
                    if(uri != null) {
                        applicationContext.contentResolver.openOutputStream(uri).use { outputStream ->
                            hssfWorkbook.write(outputStream)
                        }
                    }
                } catch (e: Exception) {
                    e.printStackTrace()
                }
            }
        }
    }
}