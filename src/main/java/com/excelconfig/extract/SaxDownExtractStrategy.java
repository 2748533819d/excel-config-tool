package com.excelconfig.extract;

import com.excelconfig.model.ExtractConfig;
import com.excelconfig.model.HeaderConfig;
import com.excelconfig.model.RangeConfig;
import com.excelconfig.sax.SaxReader;
import com.excelconfig.spi.ExtractContext;
import com.excelconfig.spi.ExtractStrategy;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 基于 SAX 流式读取的向下提取策略
 *
 * 适用于大文件场景，内存优化
 */
public class SaxDownExtractStrategy implements ExtractStrategy {

    @Override
    public List<Object> extract(Sheet sheet, ExtractConfig config, ExtractContext context) {
        // 对于 SAX 模式，需要从上下文获取输入流
        InputStream inputStream = context.getInputStream();
        if (inputStream == null) {
            // 退回到普通模式
            return new DownExtractStrategy().extract(sheet, config, context);
        }

        try {
            return extractWithSax(inputStream, config, context);
        } catch (Exception e) {
            throw new ExtractException("SAX 提取失败：" + e.getMessage(), e);
        }
    }

    /**
     * 使用 SAX 流式读取提取数据
     *
     * 由于需要查找表头位置，我们一次性读取所有数据到内存
     * 但对于大文件，仍然比 POI 用户模型节省内存
     */
    private List<Object> extractWithSax(InputStream inputStream, ExtractConfig config, ExtractContext context) throws Exception {
        HeaderConfig headerConfig = config.getHeader();
        RangeConfig rangeConfig = config.getRange();

        // 读取所有行
        List<List<String>> allRows = new SaxReader().readAll(inputStream, context.getSheetIndex());

        // 查找表头
        int headerRowIndex = -1;
        int targetColumnIndex = -1;
        int searchStartRow = 0;
        int searchEndRow = allRows.size();

        if (headerConfig.getInRows() != null) {
            searchStartRow = headerConfig.getInRows()[0] - 1;
            searchEndRow = headerConfig.getInRows()[1];
        }

        for (int rowNum = searchStartRow; rowNum < searchEndRow && rowNum < allRows.size(); rowNum++) {
            List<String> row = allRows.get(rowNum);
            for (int i = 0; i < row.size(); i++) {
                if (row.get(i).equals(headerConfig.getMatch())) {
                    headerRowIndex = rowNum;
                    targetColumnIndex = i;
                    break;
                }
            }
            if (headerRowIndex != -1) break;
        }

        if (headerRowIndex == -1 || targetColumnIndex == -1) {
            throw new ExtractException("未找到表头：" + headerConfig.getMatch());
        }

        // 提取数据
        List<Object> result = new ArrayList<>();
        int maxRows = rangeConfig != null && rangeConfig.getMaxRows() != null ?
                rangeConfig.getMaxRows() : Integer.MAX_VALUE;
        boolean skipEmpty = rangeConfig != null && Boolean.TRUE.equals(rangeConfig.getSkipEmpty());

        for (int rowNum = headerRowIndex + 1; rowNum < allRows.size() && result.size() < maxRows; rowNum++) {
            List<String> row = allRows.get(rowNum);

            String value = "";
            if (targetColumnIndex < row.size()) {
                value = row.get(targetColumnIndex);
            }

            boolean isEmpty = value == null || value.trim().isEmpty();

            if (skipEmpty && isEmpty) {
                continue;
            }

            if (!isEmpty) {
                result.add(value);
            }
        }

        return result;
    }

    @Override
    public com.excelconfig.spi.ExtractMode getSupportedMode() {
        return com.excelconfig.spi.ExtractMode.DOWN;
    }
}
