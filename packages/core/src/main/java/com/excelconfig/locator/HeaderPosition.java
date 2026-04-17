package com.excelconfig.locator;

/**
 * 表头位置
 */
public class HeaderPosition {

    /**
     * 行号（0-based）
     */
    private final int row;

    /**
     * 列号（0-based）
     */
    private final int column;

    public HeaderPosition(int row, int column) {
        this.row = row;
        this.column = column;
    }

    public int getRow() {
        return row;
    }

    public int getColumn() {
        return column;
    }

    @Override
    public String toString() {
        return String.format("%s%d", columnToRef(column), row + 1);
    }

    /**
     * 将列号转换为列引用（0 -> A, 1 -> B, ...）
     */
    private String columnToRef(int column) {
        StringBuilder result = new StringBuilder();
        while (column >= 0) {
            result.insert(0, (char) ('A' + (column % 26)));
            column = (column / 26) - 1;
        }
        return result.toString();
    }
}
