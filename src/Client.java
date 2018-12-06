import com.sun.org.apache.bcel.internal.generic.IF_ACMPEQ;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.io.*;
import java.net.*;
import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Scanner;

public class Client {
    public static final String IP_ADDR = "169.254.151.243";//服务器地址
    public static final int PORT = 3602;//服务器端口号

    public static final String SOURCE_PATH = "D:/command.xls";//excel命令路径
    public static final String RESULT_PATH = "D:/result.xls";//excel结果输出路径

    public static ArrayList<Task> tasks = new ArrayList<>();

    public static int curIndex=0; //当前执行命令\

    public static ArrayList<String> results =new ArrayList<>();
    
//    public static Socket socket=null;

    public void sendCommand(String command) {
//        if (socket==null){
        Socket socket=null;
            try {
                socket = new Socket(IP_ADDR, PORT);
            } catch (IOException e) {
                System.out.println("TCP/IP 连接失败！");
                e.printStackTrace();
                return;
            }
            System.out.println("连接成功！");
//        }
        try {
            BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(socket.getOutputStream()));     //可用PrintWriter
            BufferedReader reader = new BufferedReader(new InputStreamReader(socket.getInputStream()));
            writer.write(command+"\r");
            writer.flush();

            String res=reader.readLine();
            results.add(res);
            System.out.println(res);
            writer.close();
            reader.close();
            socket.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void start() throws ParseException {
        if (tasks.size()==0){
            System.out.println("Excel中没有指令！！");
            return;
        }
        Runnable runnable = () -> {
            while (true) {
                Date current = new Date();
                if (!tasks.get(curIndex).getDate().equals(current))
                    continue;
                String command=tasks.get(curIndex).getCommand();
                sendCommand(command);
                try {
                    Thread.sleep(500);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                //执行完成后，定位到下一条指令
                if (curIndex==tasks.size()-1){
                    File file = new File(SOURCE_PATH);
                    //最后保存结果
                    batchSaveResult(file);
                    break;
                }
                ++curIndex;
            }
        };
        Thread thread = new Thread(runnable);
        thread.start();
    }

    public static void main(String[] args) throws UnknownHostException, IOException, ParseException {
        Client client = new Client();

        File file = new File(SOURCE_PATH);
        client.readExcel(file);
        client.start();

//        String str="\r";
//        System.out.println(client.StringToAsciiString(str));
//        client.sendCommand("hello world");
    }

    public Date stringToDate(String time) throws ParseException {
        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        //设置要读取的时间字符串格式
        Date date = format.parse(time);
        return date;
    }

    // 去读Excel的方法readExcel，该方法的入口参数为一个File对象
    public void readExcel(File file) throws ParseException {
        try {
            // 创建输入流，读取Excel
            InputStream is = new FileInputStream(file.getAbsolutePath());
            // jxl提供的Workbook类
            Workbook wb = Workbook.getWorkbook(is);
            // 每个页签创建一个Sheet对象
            Sheet sheet = wb.getSheet(0);
            // sheet.getRows()返回该页的总行数
            for (int i = 1; i < sheet.getRows(); i++) {
                Task task = new Task();
                // sheet.getColumns()返回该页的总列数
//                for (int j = 0; j < sheet.getColumns(); j++) {
                String dateStr = sheet.getCell(0, i).getContents();
                String command = sheet.getCell(1, i).getContents();
                //如果出现空行了，就跳过后面的所有内容
                if (dateStr==null||dateStr.equals("")||command==null||command.equals(""))
                    break;
                task.setDate(stringToDate(dateStr));
                task.setCommand(command);
                tasks.add(task);
                System.out.println(dateStr + command);
//                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //批量保存接收到的结果
    public void batchSaveResult(File file){
        try {
            Workbook wb=Workbook.getWorkbook(file);
            WritableWorkbook book = Workbook.createWorkbook(file,wb);
            WritableSheet sheet=book.getSheet(0);
            for (int i=0;i<results.size();i++){
                Label label=new Label(2,i+1,results.get(i));
                sheet.addCell(label);
            }
            book.write();
            book.close();
            System.out.println("结果保存成功！");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
