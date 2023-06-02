package org.example;

import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.util.ArrayList;

public class MyWindow extends JFrame {
    private JTextField textFieldMonth, textFieldDay;  // 将文本框改成两个文本框
    private JCheckBox checkBox1, checkBox2, checkBox3, checkBox4;
    private JButton button;
    private JTextField textFieldFruit, textFieldVegetable;
    public MyWindow() {
        super("加价器");

        // 设置窗口大小和位置
        setSize(800, 400);
        setLocationRelativeTo(null);


        JPanel textPanel = new JPanel(new GridLayout(2, 2, 40, 20));
        JLabel labelMonth = new JLabel("月份:");
        JLabel labelDay = new JLabel("日:");
        JLabel labelFruit = new JLabel("水果加价:");
        JLabel labelVegetable = new JLabel("蔬菜加价:");

        textFieldMonth = new JTextField(10);
        textFieldDay = new JTextField(10);
        textFieldFruit = new JTextField("3", 10);
        textFieldVegetable = new JTextField("1.5", 10);
        textFieldMonth.setFont(new Font("宋体", Font.PLAIN, 24));
        textFieldDay.setFont(new Font("宋体", Font.PLAIN, 24));
        textFieldFruit.setFont(new Font("宋体", Font.PLAIN, 24));
        textFieldVegetable.setFont(new Font("宋体", Font.PLAIN, 24));

        labelMonth.setFont(new Font("宋体", Font.PLAIN, 24));
        labelDay.setFont(new Font("宋体", Font.PLAIN, 24));
        labelFruit.setFont(new Font("宋体", Font.PLAIN, 24));
        labelVegetable.setFont(new Font("宋体", Font.PLAIN, 24));
        textPanel.add(labelMonth);
        textPanel.add(textFieldMonth);
        textPanel.add(labelDay);
        textPanel.add(textFieldDay);
        textPanel.add(labelFruit);
        textPanel.add(textFieldFruit);
        textPanel.add(labelVegetable);
        textPanel.add(textFieldVegetable);



        // 创建勾选框，并使用布局管理器将它们放置在一起
        JPanel checkBoxPanel = new JPanel(new GridLayout(2, 2));
        checkBox1 = new JCheckBox("吾悦");
        checkBox2 = new JCheckBox("兴西");
        checkBox3 = new JCheckBox("圪僚沟");
        checkBox4 = new JCheckBox("赞城");
        checkBox1.setFont(new Font("宋体", Font.PLAIN, 30));
        checkBox2.setFont(new Font("宋体", Font.PLAIN, 30));
        checkBox3.setFont(new Font("宋体", Font.PLAIN, 30));
        checkBox4.setFont(new Font("宋体", Font.PLAIN, 30));
        checkBoxPanel.add(checkBox1);
        checkBoxPanel.add(checkBox2);
        checkBoxPanel.add(checkBox3);
        checkBoxPanel.add(checkBox4);

        // 创建确定按钮, 并设置大小和字体
        button = new JButton("确定");
        button.setPreferredSize(new Dimension(150, 60));
        button.setFont(new Font("宋体", Font.PLAIN, 24));
        button.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                executeMethod();
            }
        });

        // 使用布局管理器将所有组件放置在一起
        Container contentPane = getContentPane();
        contentPane.setLayout(new BorderLayout());
        contentPane.add(textPanel, BorderLayout.NORTH);
        contentPane.add(checkBoxPanel, BorderLayout.CENTER);
        contentPane.add(button, BorderLayout.SOUTH);
    }

    private void executeMethod() {
        boolean b1 = checkBox1.isSelected();
        boolean b2 = checkBox2.isSelected();
        boolean b3 = checkBox3.isSelected();
        boolean b4 = checkBox4.isSelected();
        String month = textFieldMonth.getText();
        String day = textFieldDay.getText();
        String fruit_price = textFieldFruit.getText();
        String vege_price = textFieldVegetable.getText();


        // 校验非空
        if (month.equals("") || day.equals("")) {
            JOptionPane.showMessageDialog(this, "月份或者日期不能为空，重新输入");
            textFieldMonth.setText("");
            textFieldDay.setText("");
            return;
        }
        // 校验正确性
        Integer int_month = Integer.valueOf(month);
        Integer int_day = Integer.valueOf(day);
        if (int_month < 1 || int_month > 12 || int_day < 1 || int_day > 31) {
            JOptionPane.showMessageDialog(this, "月份或者日期不正确，重新输入");
            textFieldMonth.setText("");
            textFieldDay.setText("");
            return;
        }


        ArrayList<String> strings = new ArrayList<>();
        strings.add(month + "." + day);
        strings.add(fruit_price);
        strings.add(vege_price);

        if (b1) strings.add("吾悦");
        if (b2) strings.add("兴西");
        if (b3) strings.add("圪僚沟");
        if (b4) strings.add("赞城");
        System.out.println(strings);

        // 执行需要执行的方法，使用上面获取到的勾选框的状态作为参数
        // ...
        String[] array = strings.toArray(new String[strings.size()]);
        try {
            ExcelProcessor.main(array);
        } catch (Exception e) {
            System.out.println("执行出错");
            throw new RuntimeException(e);
        }
        JOptionPane.showMessageDialog(this, "加价成功");
        System.exit(0);
    }

    public static void main(String[] args) {
        MyWindow window = new MyWindow();
        window.setVisible(true);
    }
}
