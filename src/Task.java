import java.util.Date;

public class Task {
    private Date date;
    private String command;

    public Task(){

    }

    public Task(Date date,String command){
        this.date=date;
        this.command=command;
    }
    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public String getCommand() {
        return command;
    }

    public void setCommand(String command) {
        this.command = command;
    }

}
