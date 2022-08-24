import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;

public class EmployeeSheet {
    public static String[][] data1 =  new String[10][10];
    static String[] columnNames = { "ID", "temps d'entrée", "temps de sortie", "Total des heures travaillées"};
    static JTable j = new JTable();

    private static DefaultTableModel model = new DefaultTableModel(columnNames, 0);
    public static void main(String[] args) {
        JFrame frame = new JFrame("Compteur des heures travaillées");
        JLabel selected_file = new JLabel();
        frame.setSize(820, 600);
        j = new JTable(model);
        j.setBounds(30, 40, 200, 300);
        JButton btn_SelectFile = new JButton("Choix du fichier");
        btn_SelectFile.setBounds(3, 50, 170, 30);
        selected_file.setBounds(20, 70, 100, 200);
        frame.add(btn_SelectFile);
        JPanel panel = new JPanel();
        panel.setBounds(0, 0, 900, 900);
        panel.setBackground(Color.lightGray);



        JScrollPane sp = new JScrollPane(j);
        frame.add(selected_file);
        panel.add(sp);
        frame.add(panel);
        btn_SelectFile.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                JFileChooser FileChooser = new JFileChooser();
                int i = FileChooser.showOpenDialog(null);
                if (i == JFileChooser.APPROVE_OPTION) {
                    File f = FileChooser.getSelectedFile();
                    String filepath = f.getPath();
                    String fi = f.getName();
                    //Parsing CSV Data
                    ArrayList<String> myList;
                    selected_file.setText("\t\t" + fi);
                    try {
                         myList = fetchDataFromExcel(f);
                        for (int c=0;  c<= myList.toArray().length -1; c++) {
                            System.out.println(myList.get(c));
                        }
                    } catch (IOException | ParseException ex) {
                        throw new RuntimeException(ex);
                    }
                    for (int c=0;  c<= myList.toArray().length -1; c++) {
                        String[] data=myList.get(c).split(",");
                        model.addRow(

                                new Object[]{
                                    data[0],
                                        data[1],
                                        data[2],
                                        data[3]

                                }
                        );
                        System.out.println();
                    }
                }
            }
        });
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setVisible(true);
    }
    private static ArrayList<String> fetchDataFromExcel(File myFile) throws IOException, ParseException {
        //File myFile = new File("table.xlsx");
        FileInputStream FIStream = new FileInputStream(myFile);
        XSSFWorkbook wb = new XSSFWorkbook(FIStream);
        XSSFSheet mySheet = wb.getSheetAt(0);
        Row row;

        ArrayList<String> lst = new ArrayList<>();
        String exitTimeHr = "17:00:00";
        for (int i = mySheet.getLastRowNum(); i >= 1; i--) {
            row = mySheet.getRow(i);
            String Nombredupersonnel = row.getCell(1).getStringCellValue();
            String[] EntryDateTime = row.getCell(0).getStringCellValue().split(" ");
            String getTime = EntryDateTime[1];
            String Device = row.getCell(5).getStringCellValue();
            lst.add(Nombredupersonnel + "," + Device + "," + getTime);
        }

        int[] NoDuplicate = {2, 4, 14, 8, 11};
        boolean flag = false;
        ArrayList<String> lstOtherIDs = new ArrayList<>();
        String ID5Time = null;
        String ID5EntryTime = null;
        int count = 0;
        for (int i = 0; i < (long) lst.size(); i++) {
            String[] GetListData1 = lst.get(i).split(",");
            String getTime1 = GetListData1[2];
            boolean strDeviceIP1_2 = GetListData1[1].contains("202");
            if (!strDeviceIP1_2 & (Integer.parseInt(GetListData1[0]) == NoDuplicate[0] || Integer.parseInt(GetListData1[0]) == NoDuplicate[1] || Integer.parseInt(GetListData1[0]) == NoDuplicate[2] || Integer.parseInt(GetListData1[0]) == NoDuplicate[3] || Integer.parseInt(GetListData1[0]) == NoDuplicate[4])) {
                String DiffTime = CalcTimeDifference(getTime1,exitTimeHr);
                lstOtherIDs.add(GetListData1[0] + "," + getTime1 + "," + exitTimeHr + "," + DiffTime);
            }
            for (int j = i - 1; j >= 0; j--) {
                String[] GetListData2 = lst.get(j).split(","); //splits
                String getTime2 = GetListData2[2];
                boolean strDeviceIP2_1 = GetListData2[1].contains("201");
                if (Integer.parseInt(GetListData1[0]) == Integer.parseInt(GetListData2[0])) {
                    if (strDeviceIP1_2 & strDeviceIP2_1) {
                        String CalcTotalTime = CalcTimeDifference(getTime2,getTime1);
                        if (Integer.parseInt(GetListData1[0]) == 5)
                        {
                            if(ID5EntryTime == null) {
                                ID5EntryTime = getTime2;
                                ID5Time = CalcTotalTime;
                            }
                            if (flag)
                            {
                                String ID5TimeAdd = CalcTotalSumTime1(ID5Time,CalcTotalTime);
                                //lstOtherIDs.add(GetListData2[0] + "  \t\t " + ID5EntryTime + " \t\t " + getTime1 + " \t\t " + ID5TimeAdd + "\n");
                                lstOtherIDs.add(GetListData2[0] + "," + ID5EntryTime + "," + getTime1 + "," + ID5TimeAdd);

                            }
                            flag = true;
                        }
                        else
                        {
                            //lstOtherIDs.add(GetListData2[0] + "        " + getTime2 + "        " + getTime1 + "        " + CalcTotalTime + "\n");
                            lstOtherIDs.add(GetListData2[0] + "," + getTime2 + "," + getTime1 + "," + CalcTotalTime);

                        }
                        //System.out.println(lst.get(i) + " || " + lst.get(j) + "\n");
                    }
                    break;
                }
            }
            //System.out.println(i + " => " + GetListData1[0]);
        }
        /*System.out.println("------------Result------------");
        System.out.println("ID \t    Entry Time  \t Exit Time      Total Working Hours  \n");
        for (int i=0;  i<= lstOtherIDs.toArray().length -1; i++) {
             System.out.println(lstOtherIDs.get(i));
            //lstOtherIDs.get(i).split(());
        }*/
        return  lstOtherIDs;
    }

    public static String CalcTimeDifference(String time1, String time2) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("HH:mm:ss");
        Date date1 = simpleDateFormat.parse(time1);
        Date date2 = simpleDateFormat.parse(time2);
        long differenceInMilliSeconds = Math.abs(date2.getTime() - date1.getTime());
        long differenceInHours = (differenceInMilliSeconds / (60 * 60 * 1000)) % 24;
        long differenceInMinutes = (differenceInMilliSeconds / (60 * 1000)) % 60;
        long differenceInSeconds = (differenceInMilliSeconds / 1000) % 60;
        return differenceInHours + ":"+ differenceInMinutes + ":"+ differenceInSeconds;
    }
    public static String CalcTotalSumTime1(String time1, String time2) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("HH:mm:ss");
        Date date1 = simpleDateFormat.parse(time1);
        Date date2 = simpleDateFormat.parse(time2);
        String[] getTime1 = date1.toString().split(" ");
        String[] getTime2 = date2.toString().split(" ");

        LocalTime t1 = LocalTime.parse(getTime1[3]);
        LocalTime t2 = LocalTime.parse(getTime2[3]);

        LocalTime sumT = t1.plusHours(t2.getHour())
                .plusMinutes(t2.getMinute())
                .plusSeconds(t2.getSecond());
        return sumT.toString();
    }
}