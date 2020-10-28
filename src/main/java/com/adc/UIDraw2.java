package com.adc;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * @Author wangmian
 * @Date 2020/10/26
 */
public class UIDraw2 extends JFrame {
    private JPanel contentPane;    //内容面板

    private JPanel contentPaneLeft;    //内容面板左
    private JPanel contentPaneRight;    //内容面板右

    private FileDialog fileDialog = new FileDialog(this, "文件管理");

    private final String FILE_NAME = "Data_result.xlsx";

    private List<JComboBox> boxes = new ArrayList<JComboBox>();

    private static List<String> nodeName = new ArrayList<String>(Arrays.asList(
            "长","宽","高","轴距","排量","功率","整备质量","燃料种类","额定载客"
    ));

    private int LOCATION_X=200;
    private int LOCATION_Y=50;
    private int WINDOW_WITH=1200;
    private int WINDOW_HEIGHT=800;

    public UIDraw2(){
        setTitle("车型车价库国产乘用车-字段匹配软件");    //设置窗体的标题
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);    //设置窗体退出时操作
        setBounds(LOCATION_X, LOCATION_Y, WINDOW_WITH, WINDOW_HEIGHT);    //设置窗体位置和大小

        contentPane = new JPanel();
        contentPane.setBorder(new EmptyBorder(0,0,0,0));    //设置面板的边框
        contentPane.setLayout(null);    //设置内容面板为边界布局
        setContentPane(contentPane);    //应用内容面板

        contentPaneLeft = new JPanel();
        contentPaneLeft.setBounds(0, 0, WINDOW_WITH/3, WINDOW_HEIGHT);
        contentPaneLeft.setLayout(null);

        contentPaneRight = new JPanel();
        contentPaneRight.setBounds(WINDOW_WITH/3, 0, WINDOW_WITH*2/3, WINDOW_HEIGHT);
        contentPaneRight.setLayout(null);

        contentPane.add(contentPaneLeft);
        contentPane.add(contentPaneRight);
        //-----------------------------------------------------------------------//

        //汽车之家数据ui
        JPanel panelQCZJ=new JPanel();    //新建面板用于保存文本框
        panelQCZJ.setBounds(5,WINDOW_HEIGHT/8,WINDOW_WITH,WINDOW_HEIGHT/7);
        contentPaneRight.add(panelQCZJ);    //将面板放置在边界布局的北部
        panelQCZJ.setLayout(null);

        final JButton btnQCZJ=new JButton("汽车之家数据路径");
        btnQCZJ.setFont(new Font("黑体", Font.BOLD, 16));
        btnQCZJ.setBounds(5,WINDOW_HEIGHT/7/3,200,40);
        panelQCZJ.add(btnQCZJ);

        final JTextField textQCZJ=new JTextField();    //新建文本框
        textQCZJ.setFont(new Font("黑体", Font.BOLD, 16));
        panelQCZJ.add(textQCZJ);    //将文本框增加到面板中
        textQCZJ.setBounds(215,WINDOW_HEIGHT/7/3,500,40);
        textQCZJ.setEnabled(false);
        textQCZJ.setBackground(new Color(128, 118, 105));


        //公告数据ui
        JPanel panelGG=new JPanel();    //新建面板用于保存文本框
        panelGG.setBounds(5,WINDOW_HEIGHT*3/8,WINDOW_WITH,WINDOW_HEIGHT/7);
        contentPaneRight.add(panelGG);
        panelGG.setLayout(null);

        final JButton btnGG=new JButton("公告数据路径");
        btnGG.setFont(new Font("黑体", Font.BOLD, 16));
        btnGG.setBounds(5,WINDOW_HEIGHT/7/3,200,40);
        panelGG.add(btnGG);

        final JTextField textGG=new JTextField();    //新建文本框
        textGG.setFont(new Font("黑体", Font.BOLD, 16));
        panelGG.add(textGG);    //将文本框增加到面板中
        textGG.setBounds(215,WINDOW_HEIGHT/7/3,500,40);
        textGG.setEnabled(false);
        textGG.setBackground(new Color(128, 118, 105));

        //开始执行ui
        JPanel panelOK=new JPanel();    //新建面板用于保存文本框
        panelOK.setBounds(5,WINDOW_HEIGHT/2+50,WINDOW_WITH,WINDOW_HEIGHT/7);
        contentPaneRight.add(panelOK);
        panelOK.setLayout(null);

        final JButton btnOK=new JButton("开始执行");
        btnOK.setFont(new Font("黑体", Font.BOLD, 16));
        btnOK.setBounds(400,WINDOW_HEIGHT/7/5,200,WINDOW_HEIGHT/7/5*3);
        panelOK.add(btnOK);

        //添加listener
        btnAddListener(btnQCZJ,textQCZJ);
        btnAddListener(btnGG, textGG);

        /* 确定点击 */
        btnOK.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                String filepathQCZJ=textQCZJ.getText();
                String filepathGG=textGG.getText();
                boolean flagQCZJ = validateFile(filepathQCZJ);
                boolean flagGG = validateFile(filepathGG);
                if (!flagQCZJ || !flagGG) {
                    return;
                }
                List<Integer> itemRank = new ArrayList<Integer>();
                for (JComboBox comboBox : boxes) {
                    itemRank.add(comboBox.getSelectedIndex());
                }

                /* 打开文件 */
                try {
                    openFile(filepathQCZJ, filepathGG,itemRank);
                } catch (Exception e1) {
                    e1.printStackTrace();
                }
            }
        });


        //------------------------左侧下拉列表------------------//
        generateBoxs();

        //使得程序可见
        setVisible(true);

    }


    private void jButtonActionPerformed(XSSFWorkbook workbook) {
        fileDialog.setTitle("导出");
        fileDialog.setLocation(500, 350);
        fileDialog.setMode(FileDialog.SAVE);
        fileDialog.setFile(FILE_NAME);
        fileDialog.setVisible(true);
        String stringFile = fileDialog.getDirectory() + FILE_NAME;
        if (fileDialog.getFile() != null && !fileDialog.getFile().isEmpty()) {
            stringFile = fileDialog.getDirectory()+fileDialog.getFile();
        }
        String s5 = stringFile.substring(stringFile.length() - 5);
        String s4 = stringFile.substring(stringFile.length() - 4);
        if (!s5.equals(".xlsx") && !s4.equals(".xls")) {
            stringFile = stringFile + ".xlsx";
        }
        try {
            ExcelExporter.exportTable(workbook, new File(stringFile));
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    private void generateBoxs(){
        for (int i = 0; i < nodeName.size(); i++) {
            JComboBox cmb = new JComboBox();
            boxAddItem(cmb);
            cmb.setBounds(150,20+i*80,200,40);
            cmb.setSelectedIndex(i);
            JLabel jLabel = new JLabel("第"+(i+1)+"个节点：");
            jLabel.setBounds(50,20+i*80,90,40);
            contentPaneLeft.add(cmb);
            contentPaneLeft.add(jLabel);
            boxes.add(cmb);
        }
        boxAddItemListener(boxes);
    }

    private void boxAddItemListener(List<JComboBox> boxes){
        for (int i = 0; i < boxes.size(); i++) {
            JComboBox comboBox = boxes.get(i);
            comboBox.addItemListener(new ItemListener() {
                @Override
                public void itemStateChanged(ItemEvent e) {
                    if (e.getStateChange() == ItemEvent.DESELECTED) {
                        String target = (String)comboBox.getSelectedItem();
                        if (!(target == null || target.isEmpty())) {
                            for (int j = 0; j < boxes.size(); j++) {
                                JComboBox box = boxes.get(j);
                                String all = (String)box.getSelectedItem();
                                if (!(all == null || all.isEmpty())){
                                    if (target == all) {
                                        box.setSelectedItem(e.getItem());
                                    }
                                }
                            }
                        }
                    }
                }
            });
        }
    }

    private void boxAddItem(JComboBox comboBox) {
        nodeName.forEach(node->{
            comboBox.addItem(node);
        });
    }

    private boolean validateFile(String filepath){
        if("".equals(filepath)||filepath==null){
            JOptionPane.showMessageDialog(getContentPane(), "请先选择文件~",
                    "警告", JOptionPane.WARNING_MESSAGE);
            return false;
        }

        String suffix=filepath.substring(filepath.lastIndexOf(".")+1);
        if(!(suffix.equals("xlsx")||(suffix.equals("xls")))){
            JOptionPane.showMessageDialog(getContentPane(), "请选择Excel文件~",
                    "警告", JOptionPane.WARNING_MESSAGE);
            return false;
        }
        return true;
    }

    private void btnAddListener(final JButton btn, final JTextField textField){
        btn.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                //按钮点击事件
                fileDialog.setTitle("导入");
                fileDialog.setLocation(500, 350);
                fileDialog.setMode(FileDialog.LOAD);
                fileDialog.setVisible(true);
                String path = null;
                String fileName = null;
                if (fileDialog.getDirectory() != null || !"null".equals(fileDialog.getDirectory())) {
                    path=fileDialog.getDirectory();
                }
                if (fileDialog.getFile() != null || !"null".equals(fileDialog.getFile())) {
                    fileName = fileDialog.getFile();
                }
                String filePath = path + fileName;
                if (filePath == null || filePath.equals("nullnull")) {
                    filePath = "";
                }
                textField.setText(filePath);
            }
        });
    }

    /* 打开对应的Excel文件 */
    public void openFile(String fileQCZJ,String fileGG,List<Integer> itemRank) throws IOException {
        XSSFWorkbook workbook = ExcelUtil2.validateFile(fileQCZJ, fileGG,itemRank);
        jButtonActionPerformed(workbook);
    }
}
