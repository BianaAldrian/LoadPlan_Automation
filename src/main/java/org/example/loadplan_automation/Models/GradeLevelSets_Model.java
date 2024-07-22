package org.example.loadplan_automation.Models;

import java.util.List;

public class GradeLevelSets_Model {
    private String GradeLevel;
    private List<ItemSets_Model> itemSetsModelList;

    public GradeLevelSets_Model(String gradeLevel, List<ItemSets_Model> itemSetsModelList) {
        GradeLevel = gradeLevel;
        this.itemSetsModelList = itemSetsModelList;
    }

    public String getGradeLevel() {
        return GradeLevel;
    }

    public void setGradeLevel(String gradeLevel) {
        GradeLevel = gradeLevel;
    }

    public List<ItemSets_Model> getItemSetsModelList() {
        return itemSetsModelList;
    }

    public void setItemSetsModelList(List<ItemSets_Model> itemSetsModelList) {
        this.itemSetsModelList = itemSetsModelList;
    }
}
