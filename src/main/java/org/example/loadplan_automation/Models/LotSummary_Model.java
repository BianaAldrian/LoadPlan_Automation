package org.example.loadplan_automation.Models;

public class LotSummary_Model {
    private int lot;
    private int rowNum;
    private int rowCounts;

    public LotSummary_Model(int lot, int rowNum, int rowCounts) {
        this.lot = lot;
        this.rowNum = rowNum;
        this.rowCounts = rowCounts;
    }

    public int getLot() {
        return lot;
    }

    public void setLot(int lot) {
        this.lot = lot;
    }

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public int getRowCounts() {
        return rowCounts;
    }

    public void setRowCounts(int rowCounts) {
        this.rowCounts = rowCounts;
    }
}
