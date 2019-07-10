import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

public class ExcelTool {

    private JFrame frame;
    private JTextField textFieldOA;
    private JTextField textFieldDD;
    private JButton button;
    JFileChooser jfc = new JFileChooser();

    /**
     * Launch the application.
     */
    public static void main(String[] args) {
        EventQueue.invokeLater(new Runnable() {
            public void run() {
                try {
                    ExcelTool window = new ExcelTool();
                    window.frame.setVisible(true);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });
    }

    /**
     * Create the application.
     */
    public ExcelTool() {
        initialize();
    }

    /**
     * Initialize the contents of the frame.
     */
    private void initialize() {
        frame = new JFrame();
        frame.setBounds(100, 100, 500, 300);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.getContentPane().setLayout(null);
//		定义文本框，用于显示按钮取得的地址
        textFieldOA = new JTextField();
        textFieldOA.setBounds(197, 44, 240, 21);
        frame.getContentPane().add(textFieldOA);
        textFieldOA.setColumns(10);

//		定义按钮，获得OA表数据上传路径
        JButton buttonOA = new JButton("上传OA表格");
        buttonOA.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                JFrame f = new JFrame();
                if (jfc.showOpenDialog(f) == JFileChooser.APPROVE_OPTION) {
                    textFieldOA.setText(jfc.getSelectedFile().getAbsolutePath());
                }
            }
        });
        buttonOA.setBounds(35, 42, 150, 23);
        frame.getContentPane().add(buttonOA);
//		定义文本框，用于显示按钮取得的地址
        textFieldDD = new JTextField();
        textFieldDD.setBounds(197, 72, 240, 21);
        frame.getContentPane().add(textFieldDD);
        textFieldDD.setColumns(10);

//		定义按钮，获得钉钉表数据上传路径
        JButton buttonDD = new JButton("上传钉钉表格");

        buttonDD.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                JFrame f = new JFrame();
                if (jfc.showOpenDialog(f) == JFileChooser.APPROVE_OPTION) {
                    textFieldDD.setText(jfc.getSelectedFile().getAbsolutePath());
                }
            }
        });
        buttonDD.setBounds(35, 70, 150, 23);
        frame.getContentPane().add(buttonDD);

        button = new JButton("RUN");
        button.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                try {
                    if (textFieldOA.getText() == null || textFieldDD.getText() == null) {
                        JOptionPane.showMessageDialog(new JFrame("Warning"), "Files must be uploaded!");
                    } else {
                        JOptionPane.showMessageDialog(new JFrame("Message"), new ChangeData().changeExcel(textFieldOA.getText(), textFieldDD.getText()));
                    }
                }
                catch (Exception e1) {
                    JOptionPane.showMessageDialog(new JFrame("Error"),"Error!");
                }
            }
        });
        button.setBounds(174, 152, 144, 21);
        frame.getContentPane().add(button);

        frame.setTitle("EXCEL工具");
    }
}
