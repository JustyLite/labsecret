package com.neuro.labsecret.tables;

import lombok.Data;

@Data
public class Timetable {
    private String organizationName;
    private String branchName;
    private String branchAddress;
    private String positionName;
    private double staffUnitsCount;
    private double occupiedUnitsCount;
    private double individualCount;
    private double externalPartTimeCount;
}

