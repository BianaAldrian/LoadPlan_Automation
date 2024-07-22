package org.example.loadplan_automation.Models;

import java.util.List;

public class SummaryItems_Model {
    private String LOT;
    private List<GradeLevelSets_Model> gradeLevelSetsModels;

    public SummaryItems_Model(String LOT, List<GradeLevelSets_Model> gradeLevelSetsModels) {
        this.LOT = LOT;
        this.gradeLevelSetsModels = gradeLevelSetsModels;
    }

    public String getLOT() {
        return LOT;
    }

    public void setLOT(String LOT) {
        this.LOT = LOT;
    }

    public List<GradeLevelSets_Model> getGradeLevelSetsModels() {
        return gradeLevelSetsModels;
    }

    public void setGradeLevelSetsModels(List<GradeLevelSets_Model> gradeLevelSetsModels) {
        this.gradeLevelSetsModels = gradeLevelSetsModels;
    }
}
