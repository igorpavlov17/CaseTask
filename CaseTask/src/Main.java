import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

//библиотека Apache POI позволяет работать с файлами MS Office (в т.ч. с Excel)
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

public class Main {
    public static void main(String[] args) throws IOException {
        FileInputStream fileInputStream = new FileInputStream("input.xls"); //входной файл с исходными данными
        FileOutputStream fileOutputStream = new FileOutputStream("output.xls"); //выходной файл с результатами

        Workbook input = new HSSFWorkbook(fileInputStream);
        Workbook output = new HSSFWorkbook();
        Sheet sheet = input.getSheetAt(0);
        Sheet result = output.createSheet("Результат");

        result.setDefaultColumnWidth(15);
        CellStyle cellStyle = output.createCellStyle();
        setMyCellStyle(cellStyle);

        group(sheet, result, cellStyle);

        output.write(fileOutputStream); //запись в файл
        fileInputStream.close(); //закрытие потоков
        fileOutputStream.close();
    }

    private static void group(Sheet sheet, Sheet result, CellStyle cellStyle) { //функция группировки данных
        for (int i = 0; ; i++) { //цикл для перебора строк исходной таблицы
            Row row = sheet.getRow(i);
            if (row == null) break; //выход, если строка пустая

            int rowType = row.getCell(0).getCellType();
            if (rowType == Cell.CELL_TYPE_BLANK) continue; //строку с пустыми ячейками не записываем

            Row resultLastRow = result.getRow(result.getLastRowNum());
            if (rowType == Cell.CELL_TYPE_NUMERIC && resultLastRow.getCell(0).getCellType() == Cell.CELL_TYPE_NUMERIC) { //если обе строки имеют числовые ячейки
                if (checkGroup(row, resultLastRow)) { //если критерии двух соседних строк совпадают
                    setSum(row, resultLastRow); //вычисление суммы и максимального числа в пределах группы и их запись
                    setMax(row, resultLastRow);
                    continue;
                }
            }

            Row resultRow = result.createRow(result.getPhysicalNumberOfRows()); //новая строка в новой таблице

            for (int j = 0; ; j++) { //перебор ячеек текущей строки
                Cell cell = row.getCell(j);
                if (cell == null) break; //выход, если столбец пустой
                if (!sheet.getRow(0).getCell(j).getStringCellValue().equals("C")) { //столбец с критерием "С" не записывается
                    Cell resultCell = resultRow.createCell(resultRow.getPhysicalNumberOfCells());
                    setCell(cell, resultCell); //записать значение в новую ячейку
                    resultCell.setCellStyle(cellStyle); //установить стиль ячейки
                }
            }

            resultRow.setHeightInPoints(15);
        }
    }

    private static void setMyCellStyle(CellStyle cellStyle) { //установить цвет фона ячейки и стиль границ
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
        cellStyle.setBorderBottom(CellStyle.BORDER_MEDIUM);
        cellStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
    }

    private static boolean checkGroup(Row row, Row resultLastRow) { //функция проверки равенства значений критериев
        double a = row.getCell(0).getNumericCellValue();
        double b = row.getCell(1).getNumericCellValue();
        double resultA = resultLastRow.getCell(0).getNumericCellValue();
        double resultB = resultLastRow.getCell(1).getNumericCellValue();
        return (a == resultA && b == resultB);
    }

    private static void setCell(Cell cell, Cell resultCell) {
        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
            resultCell.setCellValue(cell.getStringCellValue());
        } else {
            resultCell.setCellValue(cell.getNumericCellValue());
        }
    }

    private static void setSum(Row row, Row resultLastRow) { //функция вычисления и записи суммы
        double sum = row.getCell(3).getNumericCellValue() + resultLastRow.getCell(2).getNumericCellValue();
        resultLastRow.getCell(2).setCellValue(sum);
    }

    private static void setMax(Row row, Row resultLastRow) { //функция вычисления и записи максимального числа
        double max = Math.max(row.getCell(4).getNumericCellValue(), resultLastRow.getCell(3).getNumericCellValue());
        resultLastRow.getCell(3).setCellValue(max);
    }
}