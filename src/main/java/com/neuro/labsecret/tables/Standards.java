package com.neuro.labsecret.tables;

import lombok.Data;

@Data
public class Standards {
    private String normType;
    private String positionName;
    private double recommendedStaffNorms;
    private String conditionParameter;
    private String condition;
    private String measurementUnit;
    private double recommendedStaffNormsQuantity;
}

