package org.example.loadplan_automation;

import javafx.scene.control.CheckBox;

public class Region_model {
    private String region;
    private String division;
    private int schoolID;
    private String schoolName;
    private String gradeLevel;
    private CheckBox select;

    public Region_model(String region, String division, int schoolID, String schoolName, String gradeLevel) {
        this.region = region;
        this.division = division;
        this.schoolID = schoolID;
        this.schoolName = schoolName;
        this.gradeLevel = gradeLevel;
        this.select = new CheckBox();
    }

    public String getRegion() {
        return region;
    }

    public void setRegion(String region) {
        this.region = region;
    }

    public String getDivision() {
        return division;
    }

    public void setDivision(String division) {
        this.division = division;
    }

    public int getSchoolID() {
        return schoolID;
    }

    public void setSchoolID(int schoolID) {
        this.schoolID = schoolID;
    }

    public String getSchoolName() {
        return schoolName;
    }

    public void setSchoolName(String schoolName) {
        this.schoolName = schoolName;
    }

    public String getGradeLevel() {
        return gradeLevel;
    }

    public void setGradeLevel(String gradeLevel) {
        this.gradeLevel = gradeLevel;
    }

    public CheckBox getSelect() {
        return select;
    }

    public void setSelect(CheckBox select) {
        this.select = select;
    }
}
