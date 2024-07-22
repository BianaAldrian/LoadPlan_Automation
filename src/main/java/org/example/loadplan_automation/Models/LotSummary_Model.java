package org.example.loadplan_automation.Models;

public class LotSummary_Model {
    private int lot;
    private int startRow;
    private int endRow;

    public LotSummary_Model(int lot, int startRow, int endRow) {
        this.lot = lot;
        this.startRow = startRow;
        this.endRow = endRow;
    }

    public int getLot() {
        return lot;
    }

    public void setLot(int lot) {
        this.lot = lot;
    }

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    public int getEndRow() {
        return endRow;
    }

    public void setEndRow(int endRow) {
        this.endRow = endRow;
    }
}
