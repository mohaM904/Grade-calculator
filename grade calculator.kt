import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream

// Data class to represent a student
data class Student(
    val name: String,
    val marks: Double
) {
    val status: String
        get() = if (marks >= 50.0) "Pass" else "Fail"
}

// FUNCTION 1: Validation function
fun validateStudentMarks(student: Student): Boolean {
    return student.marks in 0.0..100.0 && student.name.isNotBlank()
}

// FUNCTION 2: Formatting function (for display purposes)
fun formatStudentRecord(student: Student): String {
    return "${student.name.uppercase()} -> Marks: ${String.format("%.2f", student.marks)} | Status: ${student.status}"
}

// Custom higher-order function that takes a lambda
fun processStudents(students: List<Student>, operation: (Student) -> Unit) {
    students.forEach(operation)
}

fun main() {
    println("=== DEMONSTRATION: Kotlin Features ===\n")

    val sampleStudents = listOf(
        Student("Alice", 75.5),
        Student("Bob", 45.0),
        Student("Charlie", 88.0),
        Student("Diana", 52.5),
        Student("Eve", 39.0)
    )

    println("1. CUSTOM HIGHER-ORDER FUNCTION with Lambda:")
    println("-".repeat(50))
    processStudents(sampleStudents) { student ->
        if (validateStudentMarks(student)) {
            println(formatStudentRecord(student))
        }
    }

    println("\n2. COLLECTION OPERATION - Filter (Students who PASSED):")
    println("-".repeat(50))
    val passingStudents = sampleStudents.filter { it.marks >= 50.0 }
    passingStudents.forEach { println("[PASS] ${it.name}: ${it.marks}") }

    println("\n3. MAP OPERATION - Extract student names:")
    println("-".repeat(50))
    val studentNames = sampleStudents.map { it.name.uppercase() }
    println(studentNames)

    println("\n4. ANY/ALL OPERATIONS:")
    println("-".repeat(50))
    val hasFailures = sampleStudents.any { it.marks < 50.0 }
    val allValid = sampleStudents.all { validateStudentMarks(it) }
    println("Has failing students: $hasFailures")
    println("All students valid: $allValid")

    println("\n" + "=".repeat(50))
    println("PROCESSING EXCEL FILE")
    println("=".repeat(50) + "\n")

    val inputFile = File("student_marks_15.xlsx")
    val outputFile = File("student_pass_status.xlsx")

    if (!inputFile.exists()) {
        System.err.println("Input file not found: ${inputFile.absolutePath}")
        return
    }

    FileInputStream(inputFile).use { fis ->
        XSSFWorkbook(fis).use { workbook ->
            val sheet = workbook.getSheetAt(0)

            val students = mutableListOf<Student>()
            for (i in 1..sheet.lastRowNum) {
                val row = sheet.getRow(i) ?: continue
                val nameCell = row.getCell(0)
                val marksCell = row.getCell(1)

                val name = when {
                    nameCell == null -> ""
                    nameCell.cellType == CellType.STRING -> nameCell.stringCellValue
                    nameCell.cellType == CellType.NUMERIC -> nameCell.numericCellValue.toString()
                    else -> ""
                }

                val marks = when {
                    marksCell == null -> 0.0
                    marksCell.cellType == CellType.NUMERIC -> marksCell.numericCellValue
                    marksCell.cellType == CellType.STRING -> marksCell.stringCellValue.toDoubleOrNull() ?: 0.0
                    else -> 0.0
                }

                students.add(Student(name.trim(), marks))
            }

            XSSFWorkbook().use { outWorkbook ->
                val outSheet = outWorkbook.createSheet("PassStatus")

                val outHeader = outSheet.createRow(0)
                outHeader.createCell(0).setCellValue("Student")
                outHeader.createCell(1).setCellValue("Marks")
                outHeader.createCell(2).setCellValue("Status")

                var outRowNum = 1
                students
                    .filter { validateStudentMarks(it) }
                    .forEach { student ->
                        val outRow = outSheet.createRow(outRowNum++)
                        outRow.createCell(0).setCellValue(student.name)
                        outRow.createCell(1).setCellValue(student.marks)
                        outRow.createCell(2).setCellValue(student.status)
                    }

                FileOutputStream(outputFile).use { fos ->
                    outWorkbook.write(fos)
                }
            }

            println("[OK] Output written to: ${outputFile.absolutePath}")
            println("[OK] Total students processed: ${students.size}")
            println("[OK] Valid students written: ${students.count { validateStudentMarks(it) }}")
        }
    }
}
