package com.neuro.labsecret.tables;

import lombok.Data;

@Data
public class MedCard {
    private String nameOfMedicalOrganization;
    private String nameOfFilial;
    private String addressOfFilial;
    private double longitude;
    private double latitude;
    private String type;
    private boolean distribution;
    private boolean dayHospital;
    private boolean medicalOutpatientClinic;
    private boolean generalPractice;
    private boolean childCenter;
    private double allergologyRoomCount;
    private double visionCabinetCount;
    private double traumatologyOrthopedicRoomCount;
    private double medicalSocialCaresRoomCount;
    private double emergencyRoomCount;
    private double childrenRoomCount;
    private double badChildrenRoomCount;
    private double pediatriciansRoomCount;
    private double goodChildrenRoomCount;
    private double freeRoomsCount;
    private double totalPopulationsCount;
    private double workersCount;
    private double oldsCount;
    private double womenCount;
    private double studentsCount;
    private double cancersCount;
    private double dayHospitalBedsCount;
    private double pediatricianBedsCount;
    private double ultrasoundMachineShiftsCount;
    private double totalVisitsCount;

}
