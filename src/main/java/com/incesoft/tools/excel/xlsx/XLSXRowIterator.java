package com.incesoft.tools.excel.xlsx;

import com.incesoft.tools.excel.ExcelRowIterator;
import com.incesoft.tools.excel.xlsx.Cell;

/**
 * Created by robert on 4/5/16.
 */
public interface XLSXRowIterator extends ExcelRowIterator {
    public Cell[] getCurRow();
}
