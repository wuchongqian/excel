package constant;

public enum OAAttendanceStatusEnum {
    CHIDAO("C", "迟到"),
    SHIJIA("S","事假"),
    BINGJIA("B","病假"),
    JIABAN("O","加班"),
    TIAOXIU("T","调休"),
    KUANGGONG("K","旷工"),
    ZAOTUI("Z","早退"),
    GONGJIA("G","公假"),
    NIANJIA("N","年假"),
    HUNJIA("H","婚假"),
    YIDIJIAOLIUFULIJIA("Y","异地交流福利假"),
    SANGJIA("F","丧假"),
    CHANJIA("M","产假"),
    CHANQIANJIANCHAJIA("CJ","产前检查假"),
    PEICHANJIA("P","陪护-陪产假"),
    CHUCHAI("-","出差"),
    WAIQIN("W","外勤"),
    GONGSHANGJIA("I","工伤假"),
    BURUJIA("BR","哺乳假"),
    YUNWANQIJIA("YW","孕晚期假"),
    QITAJIA("Q","其他假"),
    CHANQIANKOUXINJIA("CQ","产前扣薪假"),
    CHANGBINGJIA("L","长病假"),
    TIAOXIUNIANJIA("TN","调休加年假"),
    NIANJIATIAOXIU("NT","年假加调休"),
    PJJJBWCQ("X","排节假加班未出勤");


    private String code;
    private String name;

    OAAttendanceStatusEnum(String code, String name) {
        this.code = code;
        this.name = name;
    }

    public static OAAttendanceStatusEnum getTypeCode(String code) {
        for (OAAttendanceStatusEnum type : OAAttendanceStatusEnum.values()) {
            if (type.getCode().equals(code)) {
                return type;
            }
        }
        return null;
    }
    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
