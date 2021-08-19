package models;

import java.util.List;

public class FieldData {
    private String fieldHeaderName;
    private String fieldName;
    private List<Occurence> occurrenceList;

    public String getFieldHeaderName() {
        return fieldHeaderName;
    }

    public void setFieldHeaderName(String fieldHeaderName) {
        this.fieldHeaderName = fieldHeaderName;
    }

    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }

    public List<Occurence> getOccurrenceList() {
        return occurrenceList;
    }

    public void setOccurrenceList(List<Occurence> occurrenceList) {
        this.occurrenceList = occurrenceList;
    }
}
