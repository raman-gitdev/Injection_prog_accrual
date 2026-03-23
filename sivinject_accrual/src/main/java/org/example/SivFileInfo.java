package org.example;

public class SivFileInfo {

    private String fileName;
    private String mappedTo;
    private String fileType;
    private String frequency;
    private String sheetName;
    private int uidMaster;

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getMappedTo() {
        return mappedTo;
    }

    public void setMappedTo(String mappedTo) {
        this.mappedTo = mappedTo;
    }

    public String getFileType() {
        return fileType;
    }

    public void setFileType(String fileType) {
        this.fileType = fileType;
    }

    public String getFrequency() {
        return frequency;
    }

    public void setFrequency(String frequency) {
        this.frequency = frequency;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public int getUidMaster() {
        return uidMaster;
    }

    public void setUidMaster(int uidMaster) {
        this.uidMaster = uidMaster;
    }
}

