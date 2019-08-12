import constant.Constants;
import constant.OAAttendanceStatusEnum;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import utils.ImportExcel;

import javax.swing.filechooser.FileSystemView;
import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @program: excel
 * @description: 修改excel文件
 * @author: WCQian
 * @create: 2019-07-03 17:48
 **/
public class ChangeData {

    public Object changeExcel(String oaPath, String ddPath) throws Exception {
//        logger.info("开始读取excel表，并修改钉钉excel表");

        //获取当前系统的桌面路径（另一种方式）
//        File desktopDir = FileSystemView.getFileSystemView() .getHomeDirectory();
//        String desktopPath = desktopDir.getAbsolutePath();
//        //设立公共文件夹来存放生成文件，这里是取用户桌面来存放
//        String oaName = oaFile.getOriginalFilename();
//        String ddName = ddFile.getOriginalFilename();
//        File oa = new File(desktopPath, oaName);
//        File dd = new File(desktopPath, ddName);
//        FileUtils.copyInputStreamToFile(oaFile.getInputStream(), oa);
//        FileUtils.copyInputStreamToFile(ddFile.getInputStream(), dd);
//        String oaPath = oa.getPath();
//        String ddPath = dd.getPath();

        //定义输入流对象
        FileInputStream excelFileInputStream = new FileInputStream(ddPath);

        //拿到文件转化为JavaPoi可操纵类型
        Workbook workbook = WorkbookFactory.create(excelFileInputStream);
        //获取excel表格
        Sheet sheet = workbook.getSheetAt(0);
        ImportExcel poi = new ImportExcel();

        List<List<String>> OAList = poi.read(oaPath);
        List<List<String>> DDList = poi.read(ddPath);

        //创建存放表格坐标的map
        Map<String, String> map = new HashMap();

        //从钉钉表第四行开始为所需数据
        for (int i = 4; i < DDList.size(); i++) {
            List<String> dList = DDList.get(i);

            //钉钉表此行所属人姓名
            String name = dList.get(0);
            //对钉钉表中包含离职标识的姓名进行处理
            if (name.contains("离职")) {
                name = name.replace("（离职）", "");
            }

            //获取OA表中对应姓名一行数据
            List<String> list = verifyExcel(name, OAList);
            //如果未查询到对应姓名，则跳过循环
            if (null == list || list.size() == 0) continue;

            //旷工天数及坐标
            String coordinateKG = "N" + String.valueOf(i + 1);
            double kgNumOfDays = Double.valueOf(dList.get(13));

            //请假天数及坐标
            String coordinateQJ = "F" + String.valueOf(i + 1);
            double qjNumOfDays = 0f;

            //从钉钉表第15列开始为每日考勤数据
            for (int j = 15; j < dList.size(); j++) {
                String str = dList.get(j);
                //修改旷工字段
                if (Constants.STATUS_COMPLETION.equals(str)) {
                    //获取对应日期对应列数
                    int row = j - 14 + 6;

                    //取到需更改单元格坐标
                    String coordinate = transformColumn(j + 1) + String.valueOf(i + 1);

                    //获取OA表中对应状态
                    String status = list.get(row);
                    //status.replace("√", "");
                    if ("".equals(status)) {

                        map.put(coordinate, Constants.STATUS_REST);
                        //如果oa表中对应为空，那么此旷工无效，旷工天数减一
                        kgNumOfDays = kgNumOfDays - 1;
                    } else if ("√".equals(status)) {

                        map.put(coordinate, Constants.STARUS_NORMAL);
                        //如果oa表中对应为"√"，那么此旷工无效，天数减一
                        kgNumOfDays = kgNumOfDays - 1;
                    } else if (status.contains("√")) {
                        String newStatus = status.replace("√", "");

                        map.put(coordinate, "半天" + transformAttendance(newStatus));
                        //如果oa表中对应为"√N"等，那么旷工天数减0.5，请假天数加0.5
                        kgNumOfDays = kgNumOfDays - 0.5;
                        qjNumOfDays = qjNumOfDays + 0.5;
                    } else if (!status.equals("K")) {

                        map.put(coordinate, transformAttendance(status));
                        //如果oa表中对应为"N"等，那么旷工天数减1，请假天数加1
                        kgNumOfDays = kgNumOfDays - 1;
                        if(!status.equals("-")){
                            qjNumOfDays = qjNumOfDays + 1;
                        }
                    }
                //修改缺卡状态
                } else if (str.contains(Constants.STATUS_NO_COLCK_IN) && !str.contains(Constants.CELL_STATUS_ALREADY_EDITED)) {
                    //获取对应日期对应列数
                    int row = j - 14 + 6;
                    if (null == list || list.size() == 0) continue;
                    //取到需更改单元格坐标
                    String coordinate = transformColumn(j + 1) + String.valueOf(i + 1);
                    String status = list.get(row);
                    if ("√".equals(status)) {

                        continue;
                    } else if (status.contains("√")) {
                        String newStatus = status.replace("√", "");

                        map.put(coordinate, str + "\n" + transformAttendance(newStatus) + "半天(OA)");
                        qjNumOfDays = qjNumOfDays + 0.5;
                    } else {
                        map.put(coordinate, str + "\n" + transformAttendance(status) + "(OA)");
                        if (!"出差".equals(transformAttendance(status))) {
                            qjNumOfDays = qjNumOfDays + 1;
                        }
                    }
                //修改外勤状态
                } else if (str.contains(Constants.STATUS_FIELD) &&
                        !str.contains(Constants.CELL_STATUS_ALREADY_EDITED) &&
                        !str.contains(Constants.STATUS_NO_COLCK_IN) &&
                        !str.contains(Constants.STARUS_LEAVY_EARLY) &&
                        !str.contains(Constants.STATUS_WORK_BE_LATE)) {
                    //获取对应日期对应列数
                    int row = j - 14 + 6;
                    if (null == list || list.size() == 0) continue;
                    //取到需更改单元格坐标
                    String coordinate = transformColumn(j + 1) + String.valueOf(i + 1);
                    String status = list.get(row);
                    if ("√".equals(status)) {
                        continue;
                    } else if (status.contains("√")) {
                        String newStatus = status.replace("√", "");
                        map.put(coordinate, str + "\n" + transformAttendance(newStatus) + "半天(OA)");
                        if (!"出差".equals(transformAttendance(newStatus))) {
                            qjNumOfDays = qjNumOfDays + 0.5;
                        }
                    } else {
                        map.put(coordinate, str + "\n" + transformAttendance(status) + "(OA)");
                        if (!"出差".equals(transformAttendance(status))) {
                            qjNumOfDays = qjNumOfDays + 1;
                        }
                    }
                }else if (str.contains(Constants.STARUS_LEAVY_EARLY) || str.contains(Constants.STATUS_WORK_BE_LATE)){
                    //获取对应日期对应列数
                    int row = j - 14 + 6;
                    //取到需更改单元格坐标
                    String coordinate = transformColumn(j + 1) + String.valueOf(i + 1);
                    String status = list.get(row);
                    if ("√".equals(status)) {
                        continue;
                    } else if (status.contains("√")) {
                        String newStatus = status.replace("√", "");
                        map.put(coordinate, str + "\n" + transformAttendance(newStatus) + "半天(OA)");
                        qjNumOfDays = qjNumOfDays + 0.5;
                    }
                }
            }
            map.put(coordinateKG, String.valueOf(kgNumOfDays));
            map.put(coordinateQJ, String.valueOf(qjNumOfDays));
        }

        writeCells(ddPath, workbook, sheet, map);
        excelFileInputStream.close();
        return "修改结束";
    }

    /**
     * EXCEL转换列数为字母
     *
     * @param column
     * @return
     */
    private String transformColumn(int column) {
        String rs = "";
        do {
            column--;
            rs = ((char) (column % 26 + (int) 'A')) + rs;
            column = (int) ((column - column % 26) / 26);
        } while (column > 0);
        return rs;
    }

    /**
     * 修改指定单元格数值
     *
     * @param path
     * @param workbook
     * @param sheet
     * @param map
     */
    private void writeCells(String path, Workbook workbook, Sheet sheet, Map<String, String> map) {
        try {
            for (Map.Entry<String, String> entry : map.entrySet()) {
                //获取单元格的row和cell
                CellAddress address = new CellAddress(entry.getKey());
                // 获取行
                Row row = sheet.getRow(address.getRow());
                // 获取列
                Cell cell = row.getCell(address.getColumn());
                //设置单元的值
                cell.setCellValue(entry.getValue());
            }

            //写入数据
            OutputStream excelFileOutPutStream = new FileOutputStream(path);
            workbook.write(excelFileOutPutStream);
            excelFileOutPutStream.flush();
            excelFileOutPutStream.close();
            System.out.println("指定单元格设置数据写入完成");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (EncryptedDocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 根据名称查找对应表格中数值
     *
     * @param name
     * @param OAList
     * @return
     */
    private List<String> verifyExcel(String name, List<List<String>> OAList) {
        //从OA表第3行为考勤数据
        for (int i = 2; i < OAList.size(); i++) {
            List<String> oList = OAList.get(i);
            if (name.equals(oList.get(2))) {
                return oList;
            }
        }
        return null;
    }

    /**
     * 枚举转换
     *
     * @param code
     * @return
     */
    private String transformAttendance(String code) {
        OAAttendanceStatusEnum oaAttendanceStatusEnum = OAAttendanceStatusEnum.getTypeCode(code);
        if (null == oaAttendanceStatusEnum) {
            return "";
        }
        String result = "";
        switch (oaAttendanceStatusEnum) {
            case CHIDAO:
                result = "迟到";
                break;
            case HUNJIA:
                result = "婚假";
                break;
            case JIABAN:
                result = "加班";
                break;
            case SHIJIA:
                result = "事假";
                break;
            case WAIQIN:
                result = "外勤";
                break;
            case ZAOTUI:
                result = "早退";
                break;
            case BINGJIA:
                result = "病假";
                break;
            case BURUJIA:
                result = "哺乳假";
                break;
            case CHANJIA:
                result = "产假";
                break;
            case CHUCHAI:
                result = "出差";
                break;
            case GONGJIA:
                result = "公假";
                break;
            case NIANJIA:
                result = "年假";
                break;
            case QITAJIA:
                result = "其他假";
                break;
            case SANGJIA:
                result = "丧假";
                break;
            case TIAOXIU:
                result = "调休";
                break;
            case PJJJBWCQ:
                result = "排节假加班未出勤";
                break;
            case KUANGGONG:
                result = "旷工";
                break;
            case PEICHANJIA:
                result = "陪产假";
                break;
            case YUNWANQIJIA:
                result = "孕晚期假";
                break;
            case CHANGBINGJIA:
                result = "长病假";
                break;
            case GONGSHANGJIA:
                result = "工伤假";
                break;
            case CHANQIANKOUXINJIA:
                result = "产前扣薪假";
                break;
            case CHANQIANJIANCHAJIA:
                result = "产前检查假";
                break;
            case YIDIJIAOLIUFULIJIA:
                result = "异地交流福利假";
                break;
            case NIANJIATIAOXIU:
                result = "年假加调休";
                break;
            case TIAOXIUNIANJIA:
                result = "调休加年假";
                break;
            default:
                result = "";
        }
        return result;
    }
}
