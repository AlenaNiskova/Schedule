package alena;

public class DataModel_1 {

    private String lesson;
    private String cabinet;

    public DataModel_1() {
    }

    public DataModel_1(String lesson, String cabinet) {
        this.lesson = lesson;
        this.cabinet = cabinet;
    }

    public String getLesson() {
        return lesson;
    }

    public void setLesson(String lesson) {
        this.lesson = lesson;
    }

    public String getCabinet() {
        return cabinet;
    }

    public void setCabinet(String cabinet) {
        this.cabinet = cabinet;
    }

}
