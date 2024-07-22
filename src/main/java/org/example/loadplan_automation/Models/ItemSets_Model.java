package org.example.loadplan_automation.Models;

public class ItemSets_Model {
    private String itemName;
    private int qty;

    public ItemSets_Model(String itemName, int qty) {
        this.itemName = itemName;
        this.qty = qty;
    }

    public String getItemName() {
        return itemName;
    }

    public void setItemName(String itemName) {
        this.itemName = itemName;
    }

    public int getQty() {
        return qty;
    }

    public void setQty(int qty) {
        this.qty = qty;
    }
}
