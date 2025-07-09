# ExcelPro for Android

**ExcelPro** is a lightweight, easy-to-use Android library for reading and writing Microsoft Excel (`.xls` and `.xlsx`) files. It acts as a simplified wrapper around the powerful Apache POI library, tailored for common application tasks like data lookup and updates.

This library was created for the "Palm Information" project to handle on-device processing of palm tree data stored in Excel sheets.

## ✨ Features

-   **Read & Write:** Seamlessly read and write data to both `.xls` and `.xlsx` files.
-   **URI Support:** Works directly with Android's Storage Access Framework using `Uri` for modern, secure file access.
-   **Simple API:** An intuitive API for opening files, reading all data, finding specific rows, and updating cells.
-   **Dynamic Columns:** Automatically handles adding new columns to the sheet if they don't exist during an update.
-   **PNO-Based Logic:** Optimized for finding and updating rows based on a unique identifier (e.g., `PNO`).

## 🛠️ Setup

### 1. Add Dependencies

Add the following dependencies to your module's `build.gradle.kts` or `build.gradle` file. The library relies on Apache POI.

```groovy
// build.gradle (Groovy)
dependencies {
    // For reading and writing .xlsx files (Excel 2007+)
    implementation 'org.apache.poi:poi-ooxml:5.2.5'
    
    // For reading and writing .xls files (Excel 97-2003)
    implementation 'org.apache.poi:poi:5.2.5'

    // Required dependencies for POI on Android
    implementation 'org.apache.xmlbeans:xmlbeans:5.2.0'
    implementation 'org.apache.commons:commons-compress:1.26.1'
    implementation 'org.apache.commons:commons-collections4:4.4'
    implementation 'com.github.virtuald:curvesapi:1.08'
}
```

### 2. Add the Library Source
Copy the `ExcelPro.kt` file into your project's source directory (e.g., `app/src/main/java/com/yourpackage/excelpro/`).

## 🚀 How to Use

### 1. Initialize ExcelPro
Create an instance of `ExcelPro` in your Activity or ViewModel.

```kotlin
val excelPro = ExcelPro(context)
```

### 2. Open an Excel File
Use the `ActivityResultLauncher` to let the user pick an Excel file. In the result callback, pass the file's `Uri` to the `openFile` method.

```kotlin
// In your Activity/Fragment
private val filePickerLauncher = registerForActivityResult(ActivityResultContracts.OpenDocument()) { uri: Uri? ->
    uri?.let {
        try {
            excelPro.openFile(it)
            // File is now open and ready for operations
            // e.g., read the data or find a specific row
        } catch (e: Exception) {
            // Handle exceptions: file not found, wrong format, etc.
            Toast.makeText(this, "Error opening file: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }
}

fun selectFile() {
    filePickerLauncher.launch(arrayOf("application/vnd.ms-excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
}
```

### 3. Find Data by PNO
After opening the file, you can search for a row using its unique `PNO`. This is ideal for when you scan a QR code.

```kotlin
val pnoFromQrCode = "P12345"
val palmData: Map<String, String>? = excelPro.findRowByPno(pnoFromQrCode)

if (palmData != null) {
    // Row found, display the data
    val yield = palmData["Yield"] ?: "N/A"
    val oer = palmData["OER"] ?: "N/A"
    println("PNO: $pnoFromQrCode, Yield: $yield, OER: $oer")
} else {
    println("PNO: $pnoFromQrCode not found in the Excel file.")
}
```

### 4. Update Data and Save
Update a cell for a specific PNO. You can update existing columns or create new ones on the fly (like `PhotoUri`, `Timestamp`, etc.).

```kotlin
val pnoToUpdate = "P12345"
val photoUriString = "content://path/to/your/photo.jpg"
val timestamp = System.currentTimeMillis().toString()
val keterangan = "Kondisi baik."

// Update cells one by one
val photoUpdated = excelPro.updateCell(pnoToUpdate, "PhotoUri", photoUriString)
val timestampUpdated = excelPro.updateCell(pnoToUpdate, "Timestamp", timestamp)
val keteranganUpdated = excelPro.updateCell(pnoToUpdate, "Keterangan", keterangan)

if (photoUpdated && timestampUpdated && keteranganUpdated) {
    // All updates were successful, now save the file
    try {
        excelPro.saveFile()
        Toast.makeText(this, "File saved successfully!", Toast.LENGTH_SHORT).show()
    } catch (e: Exception) {
        Toast.makeText(this, "Failed to save file: ${e.message}", Toast.LENGTH_LONG).show()
    }
} else {
    Toast.makeText(this, "Failed to update record for PNO: $pnoToUpdate", Toast.LENGTH_LONG).show()
}
```

### 5. Read All Data
If you need to display all records from the Excel file in a list, you can use `readData()`.

```kotlin
val allRecords: List<Map<String, String>> = excelPro.readData()
// Now you can populate a RecyclerView adapter with this list.
allRecords.forEach { record ->
    println(record)
}
```

### 6. Close the Workbook
When you're done with the file (e.g., in `onDestroy`), close the workbook to release memory.

```kotlin
override fun onDestroy() {
    super.onDestroy()
    excelPro.close()
}
```

This completes the external library for Excel processing. The next step will be to integrate this `ExcelPro` library into the main "Palm Information" application, connect it with the QR code scanner, and implement the photo-taking and file management logic.
