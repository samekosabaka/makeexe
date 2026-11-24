package com.iamhisu.timekeep

import android.content.ContentValues
import android.content.Context
import android.os.Build
import android.os.Environment
import android.provider.MediaStore
import android.util.Base64
import android.webkit.JavascriptInterface
import android.webkit.WebView
import androidx.core.view.ViewCompat
import androidx.core.view.WindowInsetsCompat
import java.io.File
import java.io.FileOutputStream

class MainActivity : TauriActivity() {
    override fun onWebViewCreate(webView: WebView) {
        super.onWebViewCreate(webView)
        webView.addJavascriptInterface(AndroidExcelBridge(this), "AndroidBridge")

        // Tự động căn lề tránh thanh trạng thái / thanh điều hướng Edge-to-Edge trên Android 15-16
        ViewCompat.setOnApplyWindowInsetsListener(webView) { v, insets ->
            val bars = insets.getInsets(
                WindowInsetsCompat.Type.systemBars() or WindowInsetsCompat.Type.displayCutout()
            )
            v.setPadding(bars.left, bars.top, bars.right, bars.bottom)
            insets
        }
    }
}

class AndroidExcelBridge(private val context: Context) {
    @JavascriptInterface
    fun saveExcel(fileName: String, base64Data: String): String {
        return try {
            val bytes = Base64.decode(base64Data, Base64.DEFAULT)
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
                val values = ContentValues().apply {
                    put(MediaStore.Downloads.DISPLAY_NAME, fileName)
                    put(
                        MediaStore.Downloads.MIME_TYPE,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    put(
                        MediaStore.Downloads.RELATIVE_PATH,
                        Environment.DIRECTORY_DOWNLOADS + "/TIMEKEEP"
                    )
                }
                val resolver = context.contentResolver
                val uri = resolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, values)
                    ?: return "ERR: Không thể tạo file trong Download/TIMEKEEP"
                resolver.openOutputStream(uri)?.use { out ->
                    out.write(bytes)
                } ?: return "ERR: Không thể mở luồng ghi file"
                "OK: Đã lưu vào Download/TIMEKEEP/$fileName"
            } else {
                @Suppress("DEPRECATION")
                val dir = File(
                    Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS),
                    "TIMEKEEP"
                )
                if (!dir.exists()) dir.mkdirs()
                val outFile = File(dir, fileName)
                FileOutputStream(outFile).use { it.write(bytes) }
                "OK: Đã lưu vào ${outFile.absolutePath}"
            }
        } catch (e: Exception) {
            "ERR: ${e.message}"
        }
    }
}
