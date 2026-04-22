package com.excelconfig.extract;

/**
 * 单元格引用解析工具
 */
public class CellReference {

    private final int row;
    private final int col;

    public CellReference(String cellRef) {
        // 解析 "A1", "B2" 等引用
        cellRef = cellRef.toUpperCase();

        int i = 0;
        StringBuilder colRef = new StringBuilder();

        // 提取列字母
        while (i < cellRef.length() && Character.isLetter(cellRef.charAt(i))) {
            colRef.append(cellRef.charAt(i++));
        }

        // 提取行号
        String rowRef = cellRef.substring(i);

        // 转换列字母为数字（A=0, B=1, ...）
        this.col = columnLetterToNum(colRef.toString());
        this.row = Integer.parseInt(rowRef) - 1;  // 转换为 0-based
    }

    public int getRow() {
        return row;
    }

    public int getCol() {
        return col;
    }

    /**
     * 将列字母转换为数字（A=0, B=1, ..., Z=26, AA=27, ...）
     */
    private int columnLetterToNum(String colRef) {
        int col = 0;
        for (int i = 0; i < colRef.length(); i++) {
            col = col * 26 + (colRef.charAt(i) - 'A' + 1);
        }
        return col - 1;  // 转换为 0-based
    }
}
